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
# endregion
 
using System;
using System.Diagnostics;
using System.IO;

using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace GenerateLineMapUnitTests
{
	[TestClass]
	public class GenerateLineMapTests
	{
		[TestInitialize]
		public void TestInitialize()
		{
			// Set current folder to testing folder
			string assemblyCodeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;

			// Get directory name
			string dirName = Path.GetDirectoryName(assemblyCodeBase);

			// remove URL-prefix if it exists
			if (dirName.StartsWith("file:\\"))
				dirName = dirName.Substring(6);

			// set current folder
			Environment.CurrentDirectory = dirName;
		}


		[TestMethod]
		public void TestHelpScreen()
		{
			using (var consoleOutput = new ConsoleOutput())
			{
				// Check exit is normal
				StartConsoleApplication("GenerateLineMap.exe").Should().Be(0);

				// Check that help information shown correctly.
				consoleOutput.Output.Should().Contain("GenerateLineMap");
			}
		}


		[TestMethod]
		public void TestGenerateReportViaCommandLine()
		{
			using (var consoleOutput = new ConsoleOutput())
			{
				// copy the app to test on to a different filename, this is required
				// because, since TestApp1 is referenced by the unit test, VS tends to want to
				// lock it open for debugging purposes, which prevents GenerateLineMap from updating
				// its resources.
				System.IO.File.Copy("TestApp1.Exe", "TestApp1-rpt.exe", true);

				// Check exit is normal
				StartConsoleApplication("GenerateLineMap.exe", "/report TestApp1-rpt.exe").Should().Be(0);

				Debug.WriteLine(consoleOutput.Output);

				// Check that help information shown correctly.
				consoleOutput.Output.Should().Contain("Retrieved 2 strings");
			}
		}


		[TestMethod]
		public void TestGenerateReport()
		{
			// invoke the app main directly
			GenerateLineMap.Program.Main(new string[] { "path", "/report", "TestApp1.exe", "/out:TestApp1-3.exe" });

			var filename = "TestApp1-3.exe.linemapreport";
			File.Exists(filename).Should().BeTrue();

			var buf = File.ReadAllText("TestApp1-3.exe.linemapreport");
			buf.Should().Contain("SYMBOLS:");
			buf.Should().Contain("LINE NUMBERS:");
			buf.Should().Contain("NAMES:");
			buf.Should().Contain("TestApp1.Program");
		}


		[TestMethod]
		public void TestForProperLineNumInTestAppViaFileMap()
		{
			using (var consoleOutput = new ConsoleOutput())
			{
				//Apply line numbers to app
				GenerateLineMap.Program.Main(new string[] { "path", "/out:TestApp1-2.exe", "/file", "TestApp1.exe" });

				//call the console app main routine
				TestApp1.Program.Main(new string[] { });

				// Check that help information shown correctly.
				consoleOutput.Output.Should().Contain("Program.cs: line 51");
			}
		}


		[TestMethod]
		public void TestForProperLineNumInTestApp()
		{
			using (var consoleOutput = new ConsoleOutput())
			{
				//Apply line numbers to the test app
				//Generate a new version of the test app with "-1" suffix
				GenerateLineMap.Program.Main(new string[] { "path", "/out:TestApp1-1.exe", "TestApp1.exe" });

				//ideally, make sure PDB doesn't exist anymore
				//however, VS appears to lock this file open
				//if (File.Exists("TestApp1.pdb"))
				//{
				//	File.SetAttributes("TestApp1.pdb", FileAttributes.Normal);
				//	File.Delete("TestApp1.pdb");
				//}

				//now, execute test app, should file and write stack trace to console
				StartConsoleApplication("TestApp1-1.exe").Should().Be(0);

				// Check that help information shown correctly.
				consoleOutput.Output.Should().Contain("Program.cs: line 51");
			}
		}


		/// <span class="code-SummaryComment"><summary></span>
		/// Starts the console application.
		/// <span class="code-SummaryComment"></summary></span>
		/// <span class="code-SummaryComment"><param name="arguments">The arguments for console application. </span>
		/// Specify empty string to run with no arguments</param />
		/// <span class="code-SummaryComment"><returns>exit code</returns></span>
		private int StartConsoleApplication(string app, string arguments = "")
		{
			File.Exists(app).Should().BeTrue();

			// Initialize process here
			Process proc = new Process();
			proc.StartInfo.FileName = app;
			// add arguments as whole string
			proc.StartInfo.Arguments = arguments;

			// use it to start from testing environment
			proc.StartInfo.UseShellExecute = false;

			// redirect outputs to have it in testing console
			proc.StartInfo.RedirectStandardOutput = true;
			proc.StartInfo.RedirectStandardError = true;

			// set working directory
			proc.StartInfo.WorkingDirectory = Environment.CurrentDirectory;

			// start and wait for exit
			proc.Start();
			proc.WaitForExit();

			// get output to testing console.
			System.Console.WriteLine(proc.StandardOutput.ReadToEnd());
			System.Console.Write(proc.StandardError.ReadToEnd());

			// return exit code
			return proc.ExitCode;
		}
	}
}
