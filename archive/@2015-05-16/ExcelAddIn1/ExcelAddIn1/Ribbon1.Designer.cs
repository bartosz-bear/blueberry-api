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
            this.publishButton = this.Factory.CreateRibbonButton();
            this.updateButton = this.Factory.CreateRibbonButton();
            this.DownloadGroup = this.Factory.CreateRibbonGroup();
            this.IDBox = this.Factory.CreateRibbonEditBox();
            this.FetchConfigurationCheckBox = this.Factory.CreateRibbonCheckBox();
            this.downloadButton = this.Factory.CreateRibbonButton();
            this.refreshButton = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.publishButton);
            this.group1.Items.Add(this.updateButton);
            this.group1.Label = "Publish";
            this.group1.Name = "group1";
            // 
            // publishButton
            // 
            this.publishButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.publishButton.Image = ((System.Drawing.Image)(resources.GetObject("publishButton.Image")));
            this.publishButton.Label = "Publish";
            this.publishButton.Name = "publishButton";
            this.publishButton.ShowImage = true;
            this.publishButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Publish_Click);
            // 
            // updateButton
            // 
            this.updateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.updateButton.Image = ((System.Drawing.Image)(resources.GetObject("updateButton.Image")));
            this.updateButton.Label = "Update";
            this.updateButton.Name = "updateButton";
            this.updateButton.ShowImage = true;
            this.updateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Update_Click);
            // 
            // DownloadGroup
            // 
            this.DownloadGroup.Items.Add(this.IDBox);
            this.DownloadGroup.Items.Add(this.FetchConfigurationCheckBox);
            this.DownloadGroup.Items.Add(this.downloadButton);
            this.DownloadGroup.Items.Add(this.refreshButton);
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
            this.FetchConfigurationCheckBox.Label = "Repetitive";
            this.FetchConfigurationCheckBox.Name = "FetchConfigurationCheckBox";
            // 
            // downloadButton
            // 
            this.downloadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.downloadButton.Image = ((System.Drawing.Image)(resources.GetObject("downloadButton.Image")));
            this.downloadButton.Label = "Download";
            this.downloadButton.Name = "downloadButton";
            this.downloadButton.ShowImage = true;
            this.downloadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Fetch_Click);
            // 
            // refreshButton
            // 
            this.refreshButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.refreshButton.Image = ((System.Drawing.Image)(resources.GetObject("refreshButton.Image")));
            this.refreshButton.Label = "Refresh";
            this.refreshButton.Name = "refreshButton";
            this.refreshButton.ShowImage = true;
            this.refreshButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Refresh_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton publishButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup DownloadGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton downloadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox IDBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox FetchConfigurationCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton refreshButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
