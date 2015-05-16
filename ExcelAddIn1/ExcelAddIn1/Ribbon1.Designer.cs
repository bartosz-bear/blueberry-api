namespace ExcelAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.DownloadGroup = this.Factory.CreateRibbonGroup();
            this.IDBox = this.Factory.CreateRibbonEditBox();
            this.FetchConfigurationCheckBox = this.Factory.CreateRibbonCheckBox();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.Refresh = this.Factory.CreateRibbonButton();
            this.UpdateButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.DownloadGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.DownloadGroup);
            this.tab1.Label = "Blueberry API";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.UpdateButton);
            this.group1.Label = "Publish";
            this.group1.Name = "group1";
            // 
            // DownloadGroup
            // 
            this.DownloadGroup.Items.Add(this.IDBox);
            this.DownloadGroup.Items.Add(this.FetchConfigurationCheckBox);
            this.DownloadGroup.Items.Add(this.button3);
            this.DownloadGroup.Items.Add(this.Refresh);
            this.DownloadGroup.Label = "Download";
            this.DownloadGroup.Name = "DownloadGroup";
            // 
            // IDBox
            // 
            this.IDBox.Label = "Blueberry ID";
            this.IDBox.Name = "IDBox";
            this.IDBox.Text = null;
            // 
            // FetchConfigurationCheckBox
            // 
            this.FetchConfigurationCheckBox.Checked = true;
            this.FetchConfigurationCheckBox.Label = "Repetetive";
            this.FetchConfigurationCheckBox.Name = "FetchConfigurationCheckBox";
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Publish now";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "Download now";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // Refresh
            // 
            this.Refresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Refresh.Image = ((System.Drawing.Image)(resources.GetObject("Refresh.Image")));
            this.Refresh.Label = "Refresh all";
            this.Refresh.Name = "Refresh";
            this.Refresh.ShowImage = true;
            this.Refresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Refresh_Click);
            // 
            // UpdateButton
            // 
            this.UpdateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.UpdateButton.Image = ((System.Drawing.Image)(resources.GetObject("UpdateButton.Image")));
            this.UpdateButton.Label = "Update all";
            this.UpdateButton.Name = "UpdateButton";
            this.UpdateButton.ShowImage = true;
            this.UpdateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.DownloadGroup.ResumeLayout(false);
            this.DownloadGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup DownloadGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox IDBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox FetchConfigurationCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Refresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdateButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
