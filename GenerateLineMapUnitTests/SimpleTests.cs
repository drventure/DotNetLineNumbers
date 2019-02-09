using System;
using System.Diagnostics;
using System.IO;

using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace GenerateLineMapUnitTests
{
	[TestClass]
	public class SimpleTests
	{
		[TestMethod]
		public void SimpleTest()
		{
			var b = 1 + 1;
			b.Should().Be(2);
		}
	}
}
