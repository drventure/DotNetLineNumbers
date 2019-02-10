#region MIT License
/*
    MIT License

    Copyright (c) 2018 Darin Higgins

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

using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Build.Utilities;

using Microsoft.VisualBasic.ApplicationServices;


namespace GenerateLineMap
{
	/// <summary>
	/// Main Module for the GenerateLineMap utility
	/// (c) 2008-2011 Darin Higgins All Rights Reserved
	/// This is a command line utility, so this is the main entry point.
	/// </summary>
	/// <remarks></remarks>
	public static class Program
	{
		static AssemblyInfo AsmInfo = new Microsoft.VisualBasic.ApplicationServices.AssemblyInfo(Assembly.GetExecutingAssembly());

		[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
		internal static extern void LoadLibrary(string lpFileName);


		static Program()
		{
			//default to using the console logger
			Log.Logger = new ConsoleLogger();
		}


		/// <summary>
		/// Write only prop allows the build task to tell us to use the MSBuild logger instead
		/// </summary>
		public static TaskLoggingHelper MSBuildLogger
		{
			set
			{
				Log.Logger = new MSBuildLogger(value);
			}
		}

		/// <summary>
		/// Command line application entry point
		/// </summary>
		/// <remarks></remarks>
		public static void Main(string[] args)
		{
			string fileName = args.Length > 0 ? args[0] : "";
			bool bReport = false;
			bool bFile = false;
			bool bAPIResource = true;
			bool bNETResource = false;
			string outfile = "";

			//assume success
			Environment.ExitCode = 0;

			try
			{
				//skip the first arg cause it's this apps filename
				var cmdArgs = args.Skip(1).ToList();
				foreach (string s in cmdArgs)
				{
					var bHandled = false;

					if (s.Length == 2 && (s.Contains("?") || s.ToLower().Contains("h")))
					{
						ShowHelp();

						bHandled = true;
					}
					if (string.Compare(s, "/report", true) == 0)
					{
						bReport = true;

						bHandled = true;
					}
					if (s.StartsWith("/out:", StringComparison.InvariantCultureIgnoreCase))
					{
						bHandled = true;

						outfile = s.Substring(5).Trim();
						if (outfile.StartsWith("\""))
						{
							outfile = outfile.Substring(1);

							if (outfile.EndsWith("\""))
							{
								outfile = outfile.Substring(0, outfile.Length - 1);
							}
						}
					}

					if (string.Compare(s, "/file", true) == 0)
					{
						// write the line map to a separate file
						// normally, it's written back into the EXE as a resource

						bFile = true;
						bAPIResource = false;

						bNETResource = false;
						bHandled = true;
					}

					if (string.Compare(s, "/apiresource", true) == 0)
					{
						// write the line map to a winAPI resource
						// normally, it's written back into the EXE as a.net resource

						bAPIResource = true;
						bNETResource = false;

						bHandled = true;
					}
					if (string.Compare(s, "/resource", true) == 0)
					{
						// write the line map to a .net resource

						bAPIResource = false;
						bNETResource = true;

						bHandled = true;
					}

					if (!bHandled)
					{
						if (File.Exists(s))
						{
							fileName = s;
						}
					}
				}
				
				Log.LogMessage("{0} v{1}", AsmInfo.Title, AsmInfo.Version);

				Log.LogMessage(AsmInfo.Description);
				Log.LogMessage("   {0}", AsmInfo.Copyright);

				Log.LogMessage("");

				if (fileName.Length == 0)
				{
					ShowHelp();
					return;
				}

				// extract the necessary dbghelp.dll
				if (ExtractDbgHelp())
				{
					Log.LogWarning("Unable to extract dbghelp.dll to this folder.");
					return;
				}

				var lmb = new LineMapBuilder(fileName, outfile);


				if (bReport)
				{
					// just set a flag to gen a report

					lmb.CreateMapReport = true;
				}

				if (bFile)
				{
					Log.LogMessage("Creating linemap file for file {0}...", fileName);

					lmb.CreateLineMapFile();
				}

				if (bAPIResource)
				{
					Log.LogMessage("Adding linemap WIN resource in file {0}...", fileName);

					lmb.CreateLineMapAPIResource();
				}

				if (bNETResource)
				{
					String.Format("Adding linemap .NET resource in file {0}...", fileName);

					lmb.CreateLineMapResource();
				}
			}
			catch (Exception ex)
			{
				// let em know we had a failure

				Log.LogError(ex, "Unable to complete operation.");
				Environment.ExitCode = 1;
			}

			// Return an exit code of 0 on success, 1 on failure
			// NOTE that there are several early returns in this routine
		}


		private static void ShowHelp()
		{
			Log.LogMessage("");
			Log.LogMessage("Usage:");

			Log.LogMessage("   {0} FilenameOfExeOrDllFile [options]", AsmInfo);

			Log.LogMessage("where options are:");

			Log.LogMessage("   [/report][[/file]|[/resource]|[/apiresource]]");

			Log.LogMessage("");
			Log.LogMessage("/resource      (default) Create a linemap .NET resource in the target");

			Log.LogMessage("               EXE/DLL file");

			Log.LogMessage("/report        Generate report of contents of PDB file");

			Log.LogMessage("/file          Output a linemap file with the symbol and line num buffers");

			Log.LogMessage("/apiresource   Create a linemap windows resource in the target EXE/DLL file");

			Log.LogMessage("");

			Log.LogMessage("The default is 'apiresource' which embeds the linemap into");

			Log.LogMessage("the target executable as a standard windows resource.");

			Log.LogMessage("The 'file' option is mainly for testing. The resulting *.lmp");

			Log.LogMessage("file will contain source names and line numbers but no other");

			Log.LogMessage("information commonly found in PDB files.");

			Log.LogMessage("");
			Log.LogMessage("Returns an exitcode of 0 on success, 1 on failure");
		}


		/// <summary>
		/// Extract the dbghelp.dll file from the exe
		/// </summary>
		/// <returns>true on failure</returns>
		/// <remarks></remarks>
		static private bool ExtractDbgHelp()
		{
			try
			{
				// Get our assembly
				var executing_assembly = Assembly.GetExecutingAssembly();

				// Get our namespace
				// Note that this is different from the Appname because we compile to a 
				// file call *Raw so that we can run ILMerge and result in the final filename

				var my_namespace = executing_assembly.EntryPoint.DeclaringType.Namespace;

				// write the file back out
				Stream dbghelp_stream;

				dbghelp_stream = executing_assembly.GetManifestResourceStream(my_namespace + ".dbghelp.dll");

				if (dbghelp_stream != null)
				{
					// write stream to file
					var AppPath = Path.GetDirectoryName(executing_assembly.Location);

					var outname = Path.Combine(AppPath, "dbghelp.dll");

					var bExtract = true;

					if (File.Exists(outname))
					{
						// is it the right file (just check length for now)
						if ((new FileInfo(outname)).Length == dbghelp_stream.Length)
						{
							bExtract = false;
						}
					}


					if (bExtract)
					{
						var reader = new BinaryReader(dbghelp_stream);

						var buffer = reader.ReadBytes((int)dbghelp_stream.Length);

						var output = new FileStream(outname, FileMode.Create);

						output.Write(buffer, 0, (int)dbghelp_stream.Length);

						output.Close();
						reader.Close();
					}
				}
			}
			catch (Exception ex)
			{
				Log.LogError(ex, "\r\n!!! Unable to extract dbghelp.dll to current directory.");

				return true;
			}

			return false;
		}
	}
}
