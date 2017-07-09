using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestApp1
{
	class Program
	{
		static void Main(string[] args)
		{
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
				var buf = "ERROR: " + ex.ToString();
				Console.WriteLine(buf);
				System.Diagnostics.Debug.WriteLine(buf);
			}
		}
	}
}
