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
			object[,] data = this.UsedRange.Value;

			var languages = new List<string>();
			int columns = data.GetUpperBound(1);

			for( int i = 2; i <= columns; i++ )
			{
				string langName = data[1, i] as string;
				if( langName.IsEmpty() )
					break;

				languages.Add(langName);
			}

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
						values[c] = data[i, c + 2] as string;
					}

					table[id] = values;
				}
			}

			return new StringTable(languages.ToArray(), table);
		}

		public void Fill(StringTable strtab)
		{
			// format
			var languages = strtab.Languages;
			for( int i = 1; i <= languages.Length; i++ )
			{
				Excel.Range colRange = this.Columns[i];
				colRange.
			}
		}

		void SetColumnFormat(int columnIndex, string name, double width, bool wrapText)
		{
			var column = this.Columns[columnIndex];

			column.Cell[1, 1] = name;
			column.WrapText = wrapText;
			column.Width = width;
		}
	}
}
