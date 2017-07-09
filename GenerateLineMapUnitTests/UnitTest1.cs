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
		TextWriter m_normalOutput;
		StringWriter m_testingConsole;
		StringBuilder m_testingSB;

		[TestInitialize]
		public void TestFixtureInitialize()
		{
			// Set current folder to testing folder
			string assemblyCodeBase =
				System.Reflection.Assembly.GetExecutingAssembly().CodeBase;

			// Get directory name
			string dirName = Path.GetDirectoryName(assemblyCodeBase);

			// remove URL-prefix if it exists
			if (dirName.StartsWith("file:\\"))
				dirName = dirName.Substring(6);

			// set current folder
			Environment.CurrentDirectory = dirName;

			// Initialize string builder to replace console
			m_testingSB = new StringBuilder();
			m_testingConsole = new StringWriter(m_testingSB);

			// swap normal output console with testing console - to reuse 
			// it later
			m_normalOutput = System.Console.Out;
			System.Console.SetOut(m_testingConsole);
		}

		[TestCleanup]
		public void TestFixtureCleanup()
		{
			// set normal output stream to the console
			System.Console.SetOut(m_normalOutput);
		}

		public void SetUp()
		{
			// clear string builder
			m_testingSB.Remove(0, m_testingSB.Length);
		}

		public void TearDown()
		{
			// Verbose output in console
			m_normalOutput.Write(m_testingSB.ToString());
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


	public class ConsoleOutput : IDisposable
	{
		private StringWriter stringWriter;
		private TextWriter originalOutput;

		public ConsoleOutput()
		{
			stringWriter = new StringWriter();
			originalOutput = Console.Out;
			Console.SetOut(stringWriter);
		}

		public string Ouput
		{
			get
			{
				return stringWriter.ToString();
			}
		}

		public void Dispose()
		{
			Console.SetOut(originalOutput);
			stringWriter.Dispose();
		}
	}
}
