using System;
using System.IO;
using System.Text;

using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;

namespace GenerateLineMapUnitTests
{
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
