using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelion
{
	class Util
	{
	}

	static class Extention
	{
		public static bool IsValid(this string s)
		{
			return !s.IsEmpty();
		}

		public static bool IsEmpty(this string s)
		{
			return String.IsNullOrEmpty(s);
		}
	}
}
