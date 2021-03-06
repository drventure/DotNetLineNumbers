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

using DotNetLineNumbers;

using TestDLL2;


namespace TestApp1
{
	public class Program
	{
		public static void Main(string[] args)
		{
			//initialize exception extensions
			//ExceptionExtensions.UsePDB = false;

			try
			{
				var x = 1;
				var y = 2;
				var z = 3;

				y = x + z;

				throw new FormatException("Unable to format");
			}
			catch (Exception ex)
			{
				var buf = "ERROR: \r\n" + ex.ToStringEnhanced();
				Console.WriteLine(buf);
				System.Diagnostics.Debug.WriteLine("-----\r\n" + buf + "----");
				buf = "ERROR: " + ex.ToString();
				Console.WriteLine(buf);
				System.Diagnostics.Debug.WriteLine("-----\r\n" + buf + "----");
			}


			try
			{
				var lc = new TestDLL2.ListCmds();
				var l = lc.Make();
				var r = lc.AddEmUp(l);
			}
			catch (Exception ex)
			{
				var buf = "ERROR: \r\n" + ex.ToStringEnhanced();
				Console.WriteLine(buf);
				System.Diagnostics.Debug.WriteLine("-----\r\n" + buf + "----");
			}
		}
	}
}
