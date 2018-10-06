using System;
using System.Collections.Generic;

using TestDLL1;

namespace TestDLL2
{
	public class ListCmds
	{
		public List<int> Make()
		{
			return new List<int> { 1, 4, 6, 3, 4, 2, 0, 6, 7, 9, 4 };
		}

		public int AddEmUp(List<int> list)
		{
			var t = 0;
			foreach (var i in list)
			{
				var c = new MathCmds();
				t = t + c.Divide(i + 7, i);
			}
			return t;
		}
	}
}
