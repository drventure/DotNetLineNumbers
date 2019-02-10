#region MIT License
/*
    MIT License

    Copyright (c) 2017 Darin Higgins

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
using System.Diagnostics;
using System.IO;

using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using ExceptionExtensions;
using ExceptionExtensions.Internal;


namespace GenerateLineMapUnitTests
{
	[TestClass]
	public class ExceptionExtensionTests
	{
		[TestInitialize]
		public void TestInitialize()
		{
		}


		[TestMethod]
		public void BaseExceptionRenderingTest()
		{
			try
			{
				throw new FileNotFoundException("File was not found", "MyFileName.txt");
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.ToString(ExceptionOptions.Default));
			}
		}


		[TestMethod]
		public void TripleNestedExceptionRenderingTest()
		{
			try
			{
				throw new FileNotFoundException("3rd File was not found", "ThirdFileNotFound.txt",
					new FileNotFoundException("2nd File was not found", "SecondFileNotFound.txt",
					new FileNotFoundException("1st File was not found", "FirstFileNotFound.txt")));
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.ToString(ExceptionOptions.Default));
			}
		}


		[TestMethod]
		public void SerializableExceptionToStringTest()
		{
			try
			{
				throw new FileNotFoundException("3rd File was not found", "ThirdFileNotFound.txt",
					new FileNotFoundException("2nd File was not found", "SecondFileNotFound.txt",
					new FileNotFoundException("1st File was not found", "FirstFileNotFound.txt")));
			}
			catch (Exception ex)
			{
				var buf = ex.ToString(ExceptionOptions.Default);
				Debug.WriteLine(buf);
				buf.Should().Contain("1st File was not found");
			}
		}
	}
}
