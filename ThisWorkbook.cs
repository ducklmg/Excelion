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
using System.IO;

namespace Excelion
{
    public partial class ThisWorkbook
    {
		const string JsonFile = "StringTable.json";

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
			LoadFromJson();

			this.Saved = true;
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

		private void ThisWorkbook_OnBeforeSave(bool SaveAsUI, ref bool Cancel)
		{
			SaveToJson();
			
			Cancel = true;
			this.Saved = true;
		}

		public void LoadFromJson()
		{
			string path = System.IO.Path.Combine(this.Path, JsonFile);

			if( File.Exists(path) == false )
			{
				File.WriteAllText(path, "{\"Languages\":[\"LANG1\",\"LANG2\"], \"Table\":{\"SomeId\":[\"Some Text\",\"Translated Text\"]}}", Encoding.UTF8);
			}

			var stringTable = StringTable.Load(path);

			Globals.Sheet1.Write(stringTable);
		}

		public void SaveToJson()
		{
			var stringTable = Globals.Sheet1.Read();

			string str = stringTable.ToJson();

			string path = System.IO.Path.Combine(this.Path, JsonFile);
			File.WriteAllText(path, str, Encoding.UTF8);
		}

		#region VSTO Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
        {
			this.BeforeSave += new Microsoft.Office.Interop.Excel.WorkbookEvents_BeforeSaveEventHandler(this.ThisWorkbook_OnBeforeSave);
			this.Startup += new System.EventHandler(this.ThisWorkbook_Startup);
			this.Shutdown += new System.EventHandler(this.ThisWorkbook_Shutdown);

		}

		#endregion
	}
}
