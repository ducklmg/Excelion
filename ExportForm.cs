using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Excelion
{
	public partial class ExportForm : Form
	{
		public List<string> Languages = new List<string>();
		public string SourceLanguage;
		public string TargetLanguage;

		public ExportForm()
		{
			InitializeComponent();

			int max = Globals.Sheet1.UsedRange.Columns.Count;
			for( int i = 2; i <= max; i++ )
			{
				string lang = Globals.Sheet1.Cells[1, i].Value as string;
				if( lang.IsEmpty() )
					break;

				Languages.Add(lang);
			}

			if( Languages.Count >= 2 )
			{
				var langs = Languages.ToArray();

				comboBox1.Items.AddRange(langs);
				comboBox1.SelectedIndex = 0;
				comboBox2.Items.AddRange(langs);
				comboBox2.SelectedIndex = 1;
			}
		}

		private void OnExportButton(object sender, EventArgs e)
		{
			SourceLanguage = Languages[comboBox1.SelectedIndex];
			TargetLanguage = Languages[comboBox2.SelectedIndex];

			if( SourceLanguage == TargetLanguage )
			{
				MessageBox.Show("Source language is same as Target language", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			this.DialogResult = DialogResult.OK;
			this.Close();
		}
	}
}
