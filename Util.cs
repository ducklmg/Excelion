using System;
using System.Collections.Generic;
using System.Text;

namespace Excelion
{
	static class Util
	{
		public static string[] ToStringArray(object obj)
		{
			List<object> objList = (List<object>)obj;

			var result = new string[objList.Count];
			for( int i = 0; i < objList.Count; i++ )
			{
				result[i] = (string)objList[i];
			}

			return result;
		}
	}
}
