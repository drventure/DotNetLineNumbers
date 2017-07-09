using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Resources;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

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


		/// <summary>
		/// Command line application entry point
		/// </summary>
		/// <remarks></remarks>
		public static void Main(string[] args)
		{
			string fileName = "";
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
						//     normally, it's written back into the EXE as a resource

						bFile = true;
						bAPIResource = false;

						bNETResource = false;
						bHandled = true;
					}

					if (string.Compare(s, "/apiresource", true) == 0)
					{
						// write the line map to a winAPI resource
						//     normally, it's written back into the EXE as a.net resource

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
				
				Console.WriteLine(string.Format("{0} v{1}", AsmInfo.Title, AsmInfo.Version));

				Console.WriteLine(AsmInfo.Description);
				Console.WriteLine(string.Format("   {0}", AsmInfo.Copyright));

				Console.WriteLine();

				if (fileName.Length == 0)
				{
					ShowHelp();
					return;
				}

				// extract the necessary dbghelp.dll
				if (ExtractDbgHelp())
				{
					Console.WriteLine("Unable to extract dbghelp.dll to this folder.");
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
					Console.WriteLine(String.Format("Creating linemap file for file {0}...", fileName));

					lmb.CreateLineMapFile();
				}

				if (bAPIResource)
				{
					Console.WriteLine(String.Format("Adding linemap WIN resource in file {0}...", fileName));

					lmb.CreateLineMapAPIResource();
				}

				if (bNETResource)
				{
					Console.WriteLine(String.Format("Adding linemap .NET resource in file {0}...", fileName));

					lmb.CreateLineMapResource();
				}
			}
			catch (Exception ex)
			{
				// let em know we had a failure

				Console.WriteLine(string.Format("Unable to complete operation. Error: {0}\r\n\r\n{1}", ex.Message, ex.StackTrace));
				Environment.ExitCode = 1;
			}

			// Return an exit code of 0 on success, 1 on failure
			// NOTE that there are several early returns in this routine
		}


		private static void ShowHelp()
		{
			Console.WriteLine();
			Console.WriteLine("Usage:");

			Console.WriteLine(String.Format("   {0} FilenameOfExeOrDllFile [options]", AsmInfo));

			Console.WriteLine("where options are:");

			Console.WriteLine("   [/report] [[/file]|[/resource]|[/apiresource]]");

			Console.WriteLine();
			Console.WriteLine("/report        Generate report of contents of PDB file");

			Console.WriteLine("/file          Output a linemap file with the symbol and line num buffers");

			Console.WriteLine("/resource      (default) Create a linemap .NET resource in the target");

			Console.WriteLine("               EXE/DLL file");

			Console.WriteLine("/apiresource   Create a linemap windows resource in the target EXE/DLL file");

			Console.WriteLine();
			Console.WriteLine("The default is 'apiresource' which embeds the linemap into");

			Console.WriteLine("the target executable as a standard windows resource.");

			Console.WriteLine(".NET resource support is experimental at this point.");

			Console.WriteLine("The 'file' option is mainly for testing. The resulting *.linemap");

			Console.WriteLine("file will contain source names and line numbers but no other");

			Console.WriteLine("information commonly found in PDB files.");

			Console.WriteLine();
			Console.WriteLine("Returns an exitcode of 0 on success, 1 on failure");
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
				//     Note that this is different from the Appname because we compile to a 
				//     file call *Raw so that we can run ILMerge and result in the final filename

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
				Console.WriteLine(string.Format("\r\n!!! Unable to extract dbghelp.dll to current directory. Error: {0}", ex.ToString()));

				return true;
			}

			return false;
		}
	}
}
