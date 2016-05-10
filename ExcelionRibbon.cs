using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;

namespace Excelion
{
	public partial class ExcelionRibbon
	{
		private void ExcelionRibbon_Load(object sender, RibbonUIEventArgs e)
		{
		}

		private void OnExportButton(object sender, RibbonControlEventArgs e)
		{
			var form = new ExportForm();
			if( form.ShowDialog() == DialogResult.OK )
			{
				var wb = Globals.ThisWorkbook.Application.Workbooks.Add();

				Globals.Sheet1.Copy(wb.Sheets[1]);

				for( int i = 2; i <= wb.Sheets.Count; i++ )
					wb.Sheets[i].Delete();

				Microsoft.Office.Interop.Excel._Worksheet ws = wb.Sheets[1];

				ws.Name = "Translation Export";

				for( int col = form.Languages.Count + 1; col >= 2; col-- )
				{
					string name = ws.Cells[1, col].Value as string;
					if( name != form.SourceLanguage && name != form.TargetLanguage )
					{
						ws.Columns[col].Delete();
					}
				}

				// save as
				string fileName = String.Format("Excelion_{0}_{1}_{2:yyyyMMdd}.xlsx", form.SourceLanguage, form.TargetLanguage, DateTime.Now);
				string path = Path.Combine(Globals.ThisWorkbook.Path, fileName);

				wb.Activate();
				wb.SaveAs(path);
			}
		}

		private void OnMergeButton(object sender, RibbonControlEventArgs e)
		{
			var openDialog = new OpenFileDialog();
			openDialog.DefaultExt = "xlsx";
			openDialog.Filter = "*.xlsx|*.xlsx|All files|*.*";
			openDialog.Title = "Select source file";
			openDialog.Multiselect = false;

			openDialog.InitialDirectory = Globals.ThisWorkbook.Path;
			openDialog.CheckFileExists = true;

			if( openDialog.ShowDialog() != DialogResult.OK )
				return;

			string filename = openDialog.FileName;

			var wb = Globals.ThisWorkbook.Application.Workbooks.Open(filename);
			var source = wb.Sheets[1];

			StringTable sourceTable = Sheet1.Read(source.UsedRange);
			StringTable currTable = Globals.Sheet1.Read();

			wb.Close();

			if( sourceTable.Languages.Length != 2 )
			{
				MessageBox.Show("Source file must have source language and target language");
				return;
			}

			if( sourceTable.Table.Count == 0 )
			{
				MessageBox.Show("Source file have no translation data");
				return;
			}

			// merge
			var log = new List<string>();

			string sourceLang = sourceTable.Languages[0];
			string targetLang = sourceTable.Languages[1];

			int srcLangIdx = currTable.GetLanguageIndex(sourceLang);
			int tgtLangIdx = currTable.GetLanguageIndex(targetLang);

			if( srcLangIdx == -1 )
			{
				MessageBox.Show("There's no source language '" + sourceLang + "'");
				return;
			}

			if( tgtLangIdx == -1 )
			{
				MessageBox.Show("There's no target language '" + targetLang + "'");
				return;
			}

			foreach( var item in sourceTable.Table )
			{
				string id = item.Key;

				if( currTable.Table.ContainsKey(id) )
				{
					string translatingSource = sourceTable.GetString(id, 0);
					string currentSource = currTable.GetString(id, srcLangIdx);

					string translatedString = sourceTable.GetString(id, 1);

					currTable.SetString(id, tgtLangIdx, translatedString);

					// translating source have been changed. (may require re-translating)
					if( translatingSource != currentSource )
					{
						log.Add(String.Format("Warning : Id '{0}' source string have been changed. '{1}' => '{2}'", id, translatingSource, currentSource));
					}
				}
				else
				{
					log.Add(String.Format("Error : Id '{0}' does not exist (maybe removed while translating).", id));
				}
			}

			Globals.Sheet1.Write(currTable);

			if( log.Count > 0 )
				AddLogSheet(log);
		}

		void AddLogSheet(List<string> logs)
		{
			var sheets = Globals.ThisWorkbook.Sheets;
			var logSheet = sheets.Add(After:sheets[1]);
			logSheet.Name = "Log";

			for( int i = 0; i < logs.Count; i++ )
			{
				logSheet.Cells[i + 1, 1].Value = String.Format("#{0}", i + 1);
				logSheet.Cells[i + 1, 2].Value = logs[i];

				if( logs[i].StartsWith("Error") )
				{
					logSheet.Cells[i+1,1].Interior.Color = 0xaaaaff;
				}
			}
		}
	}
}
