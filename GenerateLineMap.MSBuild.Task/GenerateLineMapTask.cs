#region MIT License
/*
    MIT License

    Copyright (c) 2018-2019 Darin Higgins

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
 */
#endregion

using Microsoft.Build.Evaluation;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace GenerateLineMap.MsBuild.Task
{
	// project properties
	// http://stackoverflow.com/questions/2772426/how-to-access-the-msbuild-s-properties-list-when-coding-a-custom-task

	// dependencies
	// http://stackoverflow.com/questions/8849289/get-dependent-assemblies

	public class GenerateLineMapTask : Microsoft.Build.Utilities.Task
	{
		#region Public Properties (These are supplied by MSBuild, see the .targets files)

		[Required]
		public string SolutionDir { get; set; }

		public string SolutionPath { get; set; }

		[Required]
		public string ProjectDir { get; set; }

		public string ProjectFileName { get; set; }

		public string ProjectPath { get; set; }

		public string TargetDir { get; set; }

		public string TargetPath { get; set; }

		[Required]
		public string TargetFileName { get; set; }

		public virtual ITaskItem[] InputAssemblies { get; set; }

		public string TargetFrameworkVersion { get; set; }

		public string TargetArchitecture { get; set; }

		public string KeyFile { get; set; }

		[Required]
		public string ConfigurationFilePath { get; set; }

		#endregion


		#region Constructors

		public GenerateLineMapTask()
		{
			this.InputAssemblies = new ITaskItem[0];
		}

		#endregion


		#region Public Methods

		/// <summary>
		/// This actual Generation of the LineMap is performed here
		/// </summary>
		/// <returns></returns>
		public override bool Execute()
		{
			LogInputVariables();

			string jsonConfig;
			string exePath;

			Settings settings = null;

			// try to read configuration if file exists
			if (!ReadConfigFile(out jsonConfig)) return false;

			// replace tokens if applicable
			if (!string.IsNullOrWhiteSpace(jsonConfig))
			{
				jsonConfig = ReplaceTokens(jsonConfig);
			}

			// if json config exists, try to deserialize into settings object
			if (!string.IsNullOrWhiteSpace(jsonConfig) && !DeserializeJsonConfig(jsonConfig, out settings))
			{
				return false;
			}

			// create instance if settings still null which indicates a custom json config was not used
			settings = settings ?? new Settings();
			
			// apply defaults
			SetDefaults(settings);

			if (settings.General.AlternativeGenerateLineMapPath.IsNotNullOrWhiteSpace())
			{
				if (!File.Exists(settings.General.AlternativeGenerateLineMapPath))
				{
					Log.LogError($"An alternative path for GenerateLineMap.exe was provided but the file was not found: {settings.General.AlternativeGenerateLineMapPath}");
					return false;
				}
				else
				{
					exePath = settings.General.AlternativeGenerateLineMapPath;
					Log.LogMessage("Using alternative GenerateLineMap path: {0}", settings.General.AlternativeGenerateLineMapPath);
				}
			}
			else
			{
				exePath = this.GetGenerateLineMapPath();
			}

			return GenerateLineMapBasedOn(settings);
		}

		#endregion


		#region Private Methods

		private void LogInputVariables()
		{
			Log.LogMessage($"SolutionDir: {SolutionDir}");
			Log.LogMessage($"SolutionPath: {SolutionPath}");
			Log.LogMessage($"ProjectDir: {ProjectDir}");
			Log.LogMessage($"ProjectFileName: {ProjectFileName}");
			Log.LogMessage($"ProjectPath: {ProjectPath}");
			Log.LogMessage($"TargetDir: {TargetDir}");
			Log.LogMessage($"TargetPath: {TargetPath}");
			Log.LogMessage($"TargetFileName: {TargetFileName}");
			Log.LogMessage($"TargetFrameworkVersion: {TargetFrameworkVersion}");
			Log.LogMessage($"TargetArchitecture: {TargetArchitecture}");
			Log.LogMessage($"KeyFile: {KeyFile}");
			Log.LogMessage($"ConfigurationFilePath: {ConfigurationFilePath}");
		}


		private bool DeserializeJsonConfig(string jsonConfig, out Settings settings)
		{

			settings = null;
			bool success = true;

			if (string.IsNullOrWhiteSpace(jsonConfig))
			{
				Log.LogError("Unable to deserialize configuration. Configuration string is null or empty.");
				return false;
			}

			try
			{
				Log.LogMessage("Deserializing configuration.");
				settings = Settings.FromJson(jsonConfig);
				Log.LogMessage("Configuration file deserialized successfully.");
			}
			catch (Exception ex)
			{
				Log.LogError("Error deserializing configuration.");
				Log.LogErrorFromException(ex);
				success = false;
			}

			return success;
		}


		private string ReplaceTokens(string jsonConfig)
		{
			return jsonConfig
				.Replace("$(SolutionDir)", EscapePath(this.SolutionDir))
				.Replace("$(SolutionPath)", EscapePath(this.SolutionPath))
				.Replace("$(ProjectDir)", EscapePath(this.ProjectDir))
				.Replace("$(ProjectFileName)", EscapePath(this.ProjectFileName))
				.Replace("$(ProjectPath)", EscapePath(this.ProjectPath))
				.Replace("$(TargetDir)", EscapePath(this.TargetDir))
				.Replace("$(TargetPath)", EscapePath(this.TargetPath))
				.Replace("$(TargetFileName)", EscapePath(this.TargetFileName))
				.Replace("$(AssemblyOriginatorKeyFile)", EscapePath(this.KeyFile));
		}


		private void SetDefaults(Settings settings)
		{

			if (settings.IsNull()) throw new ArgumentNullException(nameof(settings));

			if (settings.General.IsNull())
			{
				settings.General = new GeneralSettings();
			}

			if (!settings.General.OutputFile.IsNotNullOrWhiteSpace())
			{
				settings.General.OutputFile = Path.Combine(this.TargetDir, "GenerateLineMap", this.TargetFileName);
				Log.LogMessage("Applying default value for OutputFile.");
			}

			if (settings.Advanced.IsNull())
			{
				settings.Advanced = new AdvancedSettings();
			}
		}


		private bool GenerateLineMapBasedOn(Settings settings)
		{
			bool success = true;
			//Assembly ilmerge = LoadILMerge(mergerPath);
			//Type ilmergeType = ilmerge.GetType("ILMerging.ILMerge", true, true);
			//if (ilmergeType.IsNull()) throw new InvalidOperationException("Cannot find 'ILMerging.ILMerge' in executable.");

			Log.LogMessage("Setting up GenerateLineMap.");

			//string[] tp = settings.General.TargetPlatform.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

			try
			{
				string outputDir = Path.GetDirectoryName(settings.General.OutputFile);
				if (!Directory.Exists(outputDir))
				{
					Log.LogWarning($"Output directory not found. An attempt to create the directory will be made: {outputDir}");
					Directory.CreateDirectory(outputDir);
					Log.LogMessage("Output directory created.");
				}

				var outfile = string.Format("/out:\"{0}\"", this.TargetPath);
				var cmdline = string.Format("GenerateLineMap {0} {1}", outfile, this.TargetPath);
				Log.LogMessage("Effectively executing command line: " + cmdline);

				//pass on our MSBuild logger to let GenerateLineMap use it
				GenerateLineMap.Program.MSBuildLogger = Log;

				//and actually perform the line map generation
				GenerateLineMap.Program.Main(new string[] { "{ignoredpathargument}", outfile, this.TargetPath });

				//if (merger.StrongNameLost)
				//	Log.LogMessage(MessageImportance.High, "StrongNameLost = true");
			}
			catch (Exception exception)
			{
				Log.LogErrorFromException(exception);
				success = false;
			}
		
			return success;
		}
		
		
		private bool ReadConfigFile(out string jsonConfig)
		{
			jsonConfig = string.Empty;
			bool success = true;

			if (string.IsNullOrWhiteSpace(ConfigurationFilePath))
			{
				Log.LogWarning("Path for configuration file is empty. Default values will be applied.");
			}
			else
			{
				if (!File.Exists(ConfigurationFilePath))
				{
					Log.LogMessage($"Using default configuration. A custom configuration file was not found at: {ConfigurationFilePath}");
				}
				else
				{
					try
					{

						Log.LogMessage($"Loading configuration from: {ConfigurationFilePath}");
						jsonConfig = File.ReadAllText(ConfigurationFilePath, Encoding.UTF8);
						Log.LogMessage($"Configuration file loaded successfully.");

					}
					catch (Exception ex)
					{
						Log.LogError($"Error reading configuration file.");
						Log.LogErrorFromException(ex);
						success = false;
					}
				}
			}

			return success;
		}


		public string GetGenerateLineMapPath()
		{
			string exePath = null;
			string errMsg;
			var failedPaths = new List<string>();

			// look at same directory as this assembly (task dll);
			if (ExeLocationHelper.TryValidateGenerateLineMapPath(Path.GetDirectoryName(this.GetType().Assembly.Location), out exePath))
			{
				Log.LogMessage($"GenerateLineMap.exe found at (task location): {Path.GetDirectoryName(this.GetType().Assembly.Location)}");
				return exePath;
			}

			errMsg = $"GenerateLineMap.exe not found at (task location): {Path.GetDirectoryName(this.GetType().Assembly.Location)}";
			failedPaths.Add(errMsg);
			Log.LogMessage(errMsg);

			// look at target dir;
			if (!string.IsNullOrWhiteSpace(TargetDir))
			{
				if (ExeLocationHelper.TryValidateGenerateLineMapPath(this.TargetDir, out exePath))
				{
					Log.LogMessage($"GenerateLineMap.exe found at (target dir): {this.TargetDir}");
					return exePath;
				}

				errMsg = $"GenerateLineMap.exe not found at (target dir): {this.TargetDir}";
				failedPaths.Add(errMsg);
				Log.LogMessage(errMsg);
			}

			// look for "packages" folder at the solution root and if one is found, look for DotNetLineNumbers package folder
			if (!string.IsNullOrWhiteSpace(this.SolutionDir))
			{
				if (ExeLocationHelper.TryGenerateLineMapInSolutionDir(this.SolutionDir, out exePath))
				{
					Log.LogMessage($"GenerateLineMap.exe found at (solution dir): {this.SolutionDir}");
					return exePath;
				}

				errMsg = $"GenerateLineMap.exe not found at (solution dir): {this.SolutionDir}";
				failedPaths.Add(errMsg);
				Log.LogMessage(errMsg);
			}

			// get the location of the this assembly (task dll) and assumes it is under the packages folder.
			// use this information to determine the possible location of the executable.
			if (ExeLocationHelper.TryLocatePackagesFolder(Log, out exePath))
			{
				Log.LogMessage($"GenerateLineMap.exe found at custom package location: {exePath}");
				return exePath;
			}

			foreach (var err in failedPaths)
			{
				Log.LogWarning(err);
			}

			Log.LogWarning($"Unable to determine custom package location or, location was determined but a DotNetLineNumbers package folder was not found.");

			return exePath;
		}


		private string ToAbsolutePath(string relativePath)
		{
			// if path is not rooted assume it is relative.
			// convert relative to absolute using project dir as root.

			if (string.IsNullOrWhiteSpace(relativePath)) throw new ArgumentNullException(relativePath);

			if (Path.IsPathRooted(relativePath))
			{
				return relativePath;
			}

			return Path.GetFullPath(Path.Combine(ProjectDir, relativePath));
		}


		private string EscapePath(string path)
		{
			return Regex.Replace(path ?? "", @"\\", @"\\");
		}

		#endregion

	}
}
