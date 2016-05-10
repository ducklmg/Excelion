using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Excelion
{
	public partial class Sheet1
	{
		#region VSTO Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Sheet1_Startup);
			this.Shutdown += new System.EventHandler(Sheet1_Shutdown);
		}

		#endregion

		private void Sheet1_Startup(object sender, System.EventArgs e)
		{
		}

		private void Sheet1_Shutdown(object sender, System.EventArgs e)
		{
		}

		public StringTable Read()
		{
			return Read(this.UsedRange);
		}

		public static StringTable Read(Excel.Range usedRange)
		{
			object[,] data = usedRange.Value;

			// header
			var languages = new List<string>();
			int columns = data.GetUpperBound(1);

			for( int i = 2; i <= columns; i++ )
			{
				string langName = data[1, i] as string;
				if( langName.IsEmpty() )
					break;

				languages.Add(langName);
			}

			// table
			var table = new Dictionary<string, string[]>();
			int rows = data.GetUpperBound(0);

			for( int i = 2; i <= rows; i++ )
			{
				string id = data[i, 1] as string;
				if( id.IsValid() )
				{
					string[] values = new string[languages.Count];

					for( int c = 0; c < languages.Count; c++ )
					{
						string v = data[i, c + 2] as string;
						if( v == null )
							v = String.Empty;

						values[c] = v;
					}

					table[id] = values;
				}
			}

			return new StringTable(languages.ToArray(), table);
		}

		public void Write(StringTable strtab)
		{
			var languages = strtab.Languages;
			var table = strtab.Table;

			// header
			for( int i = 0; i < languages.Length; i++ )
			{
				Cells[1, i + 2] = languages[i];
			}

			// table
			string[,] data = new string[table.Count, languages.Length + 1];

			int row = 0;
			foreach( var item in table )
			{
				data[row, 0] = item.Key;

				for( int col = 0; col < item.Value.Length; col++ )
				{
					string v = item.Value[col];
					if( v != null && v.StartsWith("'") )        // prepend additional one, if a value starts with quotation mark. it's excel rule :(
						v = "'" + v;

					data[row, col + 1] = v;
				}

				row++;
			}

			var tableRange = Range[Cells[2, 1], Cells[table.Count + 1, languages.Length + 1]];
			tableRange.Value = data;                // for speed, set cell values in one call
		}
	}
}
