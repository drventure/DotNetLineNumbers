using System;
using System.IO;
using System.Text;

using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;

namespace GenerateLineMapUnitTests
{
	[TestClass]
	public class GenerateLineMapUnitTests
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
				StartConsoleApplication("").Should().Be(0);

				// Check that help information shown correctly.
				consoleOutput.Ouput.Should().Contain("GenerateLineMap");
			}
		}


		[TestMethod]
		public void TestGenerateReportViaCommandLine()
		{
			using (var consoleOutput = new ConsoleOutput())
			{
				// Check exit is normal
				StartConsoleApplication("/report TestApp1.exe").Should().Be(0);

				// Check that help information shown correctly.
				consoleOutput.Ouput.Should().Contain("Retrieved 10 lines");
			}
		}


		[TestMethod]
		public void TestGenerateReport()
		{
			// invoke the app main directly
			GenerateLineMap.Program.Main(new string[] { "path", "/report", "TestApp1.exe" });

			var filename = "TestApp1.exe.linemapreport";
			File.Exists(filename).Should().BeTrue();

			var buf = File.ReadAllText("TestApp1.exe.linemapreport");
			buf.Should().Contain("SYMBOLS:");
			buf.Should().Contain("LINE NUMBERS:");
			buf.Should().Contain("NAMES:");
			buf.Should().Contain("testapp1.program");
		}


		/// <span class="code-SummaryComment"><summary></span>
		/// Starts the console application.
		/// <span class="code-SummaryComment"></summary></span>
		/// <span class="code-SummaryComment"><param name="arguments">The arguments for console application. </span>
		/// Specify empty string to run with no arguments</param />
		/// <span class="code-SummaryComment"><returns>exit code</returns></span>
		private int StartConsoleApplication(string arguments)
		{
			// Initialize process here
			Process proc = new Process();
			proc.StartInfo.FileName = "GenerateLineMap.exe";
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
