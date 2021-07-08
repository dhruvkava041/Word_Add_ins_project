
namespace WordAddIn1
{
    partial class WordAddInRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public WordAddInRibbon()
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
            if (disposing && (components != null))
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
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.ChNameNNum = this.Factory.CreateRibbonButton();
            this.SecNameNNum = this.Factory.CreateRibbonButton();
            this.SubSecNameNNum = this.Factory.CreateRibbonButton();
            this.GeneralText = this.Factory.CreateRibbonButton();
            this.SpecialText = this.Factory.CreateRibbonButton();
            this.SaveAsPdf = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.ChNameNNum);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.SecNameNNum);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.SubSecNameNNum);
            this.group1.Items.Add(this.separator3);
            this.group1.Items.Add(this.GeneralText);
            this.group1.Items.Add(this.separator4);
            this.group1.Items.Add(this.SpecialText);
            this.group1.Label = "SELECT THE TEXT YOU WANT TO FORMAT AND THEN PRESS THE REQUIERED BUTTON";
            this.group1.Name = "group1";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // group2
            // 
            this.group2.Items.Add(this.SaveAsPdf);
            this.group2.Label = "OTHER FUNCTIONALITIES";
            this.group2.Name = "group2";
            // 
            // ChNameNNum
            // 
            this.ChNameNNum.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ChNameNNum.Image = global::WordAddIn1.Properties.Resources.text_format;
            this.ChNameNNum.Label = "Chapter Name  And Number";
            this.ChNameNNum.Name = "ChNameNNum";
            this.ChNameNNum.ShowImage = true;
            this.ChNameNNum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChNameNNum_Click);
            // 
            // SecNameNNum
            // 
            this.SecNameNNum.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SecNameNNum.Image = global::WordAddIn1.Properties.Resources.text_format;
            this.SecNameNNum.Label = "Section Name And Number";
            this.SecNameNNum.Name = "SecNameNNum";
            this.SecNameNNum.ShowImage = true;
            this.SecNameNNum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SecNameNNum_Click);
            // 
            // SubSecNameNNum
            // 
            this.SubSecNameNNum.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SubSecNameNNum.Image = global::WordAddIn1.Properties.Resources.text_format;
            this.SubSecNameNNum.Label = "Subsection Name And Number";
            this.SubSecNameNNum.Name = "SubSecNameNNum";
            this.SubSecNameNNum.ShowImage = true;
            this.SubSecNameNNum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SubSecNameNNum_Click);
            // 
            // GeneralText
            // 
            this.GeneralText.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.GeneralText.Image = global::WordAddIn1.Properties.Resources.text_format;
            this.GeneralText.Label = "General Text";
            this.GeneralText.Name = "GeneralText";
            this.GeneralText.ShowImage = true;
            this.GeneralText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GeneralText_Click);
            // 
            // SpecialText
            // 
            this.SpecialText.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SpecialText.Image = global::WordAddIn1.Properties.Resources.text_format;
            this.SpecialText.Label = "Special Text";
            this.SpecialText.Name = "SpecialText";
            this.SpecialText.ShowImage = true;
            this.SpecialText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SpecialText_Click);
            // 
            // SaveAsPdf
            // 
            this.SaveAsPdf.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SaveAsPdf.Image = global::WordAddIn1.Properties.Resources.SaveAsPdf;
            this.SaveAsPdf.Label = "SaveAsPdf";
            this.SaveAsPdf.Name = "SaveAsPdf";
            this.SaveAsPdf.ShowImage = true;
            this.SaveAsPdf.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveAsPdf_Click);
            // 
            // WordAddInRibbon
            // 
            this.Name = "WordAddInRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.WordAddInRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChNameNNum;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SecNameNNum;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SubSecNameNNum;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GeneralText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SpecialText;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SaveAsPdf;
    }

    partial class ThisRibbonCollection
    {
        internal WordAddInRibbon WordAddInRibbon
        {
            get { return this.GetRibbon<WordAddInRibbon>(); }
        }
    }
}
