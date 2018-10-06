using System;

namespace TestDLL1
{
	public class MathCmds
	{
		public int Divide(int x, int y)
		{
			var t = x * y;
			var s = $"Tested {x} - {y} - ";
			s = s + t;
			return x / y;
		}
	}
}
