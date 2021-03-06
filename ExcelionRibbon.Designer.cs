﻿namespace Excelion
{
	partial class ExcelionRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public ExcelionRibbon()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if( disposing && (components != null) )
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tab1 = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.button1 = this.Factory.CreateRibbonButton();
			this.button2 = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.group1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.ControlId.OfficeId = "TabHome";
			this.tab1.Groups.Add(this.group1);
			this.tab1.Label = "TabHome";
			this.tab1.Name = "tab1";
			// 
			// group1
			// 
			this.group1.Items.Add(this.button1);
			this.group1.Items.Add(this.button2);
			this.group1.Label = "Excelion";
			this.group1.Name = "group1";
			this.group1.Position = this.Factory.RibbonPosition.BeforeOfficeId("GroupClipboard");
			// 
			// button1
			// 
			this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.button1.Label = "Export";
			this.button1.Name = "button1";
			this.button1.OfficeImageId = "Export";
			this.button1.ShowImage = true;
			this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnExportButton);
			// 
			// button2
			// 
			this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.button2.Label = "Merge";
			this.button2.Name = "button2";
			this.button2.OfficeImageId = "Import";
			this.button2.ShowImage = true;
			this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnMergeButton);
			// 
			// ExcelionRibbon
			// 
			this.Name = "ExcelionRibbon";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExcelionRibbon_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.group1.ResumeLayout(false);
			this.group1.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
	}

	partial class ThisRibbonCollection
	{
		internal ExcelionRibbon ExcelionRibbon
		{
			get { return this.GetRibbon<ExcelionRibbon>(); }
		}
	}
}
