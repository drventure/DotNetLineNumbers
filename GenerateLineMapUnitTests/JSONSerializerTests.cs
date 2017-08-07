using System;
using System.Diagnostics;
using System.IO;

using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using ExceptionExtensions;


namespace GenerateLineMapUnitTests
{
	[TestClass]
	public class JSONSerializerTests
	{ 
		[TestMethod]
		public void SerializableExceptionToJSONTest()
		{
			try
			{
				throw new FileNotFoundException("3rd File was not found", "ThirdFileNotFound.txt",
					new FileNotFoundException("2nd File was not found", "SecondFileNotFound.txt",
					new FileNotFoundException("1st File was not found", "FirstFileNotFound.txt")));
			}
			catch (Exception ex)
			{
				var buf = ex.ToJSON();
				Debug.WriteLine(buf);
				buf.Should().Contain("1st File was not found");
			}
		}
	}
}
