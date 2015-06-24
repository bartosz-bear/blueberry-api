namespace ExcelAddIn1
{
    partial class BlueberryRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public BlueberryRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BlueberryRibbon));
            this.BluberryTab = this.Factory.CreateRibbonTab();
            this.ArgumentsGroup = this.Factory.CreateRibbonGroup();
            this.IDBox = this.Factory.CreateRibbonEditBox();
            this.FetchConfigurationCheckBox = this.Factory.CreateRibbonCheckBox();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.DownloadGroup = this.Factory.CreateRibbonGroup();
            this.WebPlatformGroup = this.Factory.CreateRibbonGroup();
            this.Other = this.Factory.CreateRibbonGroup();
            this.LoginGroup = this.Factory.CreateRibbonGroup();
            this.publishButton = this.Factory.CreateRibbonButton();
            this.updateButton = this.Factory.CreateRibbonButton();
            this.downloadButton = this.Factory.CreateRibbonButton();
            this.refreshButton = this.Factory.CreateRibbonButton();
            this.GoToWebPlatformButton = this.Factory.CreateRibbonButton();
            this.LogInButton = this.Factory.CreateRibbonButton();
            this.LogOutButton = this.Factory.CreateRibbonButton();
            this.TestButton = this.Factory.CreateRibbonButton();
            this.usernameBox = this.Factory.CreateRibbonEditBox();
            this.passwordBox = this.Factory.CreateRibbonEditBox();
            this.BluberryTab.SuspendLayout();
            this.ArgumentsGroup.SuspendLayout();
            this.group1.SuspendLayout();
            this.DownloadGroup.SuspendLayout();
            this.WebPlatformGroup.SuspendLayout();
            this.Other.SuspendLayout();
            this.LoginGroup.SuspendLayout();
            // 
            // BluberryTab
            // 
            this.BluberryTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.BluberryTab.Groups.Add(this.ArgumentsGroup);
            this.BluberryTab.Groups.Add(this.group1);
            this.BluberryTab.Groups.Add(this.DownloadGroup);
            this.BluberryTab.Groups.Add(this.WebPlatformGroup);
            this.BluberryTab.Groups.Add(this.LoginGroup);
            this.BluberryTab.Groups.Add(this.Other);
            this.BluberryTab.Label = "Blueberry API";
            this.BluberryTab.Name = "BluberryTab";
            // 
            // ArgumentsGroup
            // 
            this.ArgumentsGroup.Items.Add(this.IDBox);
            this.ArgumentsGroup.Items.Add(this.FetchConfigurationCheckBox);
            this.ArgumentsGroup.Label = "Arguments";
            this.ArgumentsGroup.Name = "ArgumentsGroup";
            // 
            // IDBox
            // 
            this.IDBox.Label = "Blueberry ID";
            this.IDBox.Name = "IDBox";
            this.IDBox.SizeString = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
            this.IDBox.Text = null;
            // 
            // FetchConfigurationCheckBox
            // 
            this.FetchConfigurationCheckBox.Checked = true;
            this.FetchConfigurationCheckBox.Label = "Repetitive";
            this.FetchConfigurationCheckBox.Name = "FetchConfigurationCheckBox";
            // 
            // group1
            // 
            this.group1.Items.Add(this.publishButton);
            this.group1.Items.Add(this.updateButton);
            this.group1.Label = "Publish";
            this.group1.Name = "group1";
            // 
            // DownloadGroup
            // 
            this.DownloadGroup.Items.Add(this.downloadButton);
            this.DownloadGroup.Items.Add(this.refreshButton);
            this.DownloadGroup.Label = "Download";
            this.DownloadGroup.Name = "DownloadGroup";
            // 
            // WebPlatformGroup
            // 
            this.WebPlatformGroup.Items.Add(this.GoToWebPlatformButton);
            this.WebPlatformGroup.Label = "Web Platform";
            this.WebPlatformGroup.Name = "WebPlatformGroup";
            // 
            // Other
            // 
            this.Other.Items.Add(this.TestButton);
            this.Other.Label = "Other";
            this.Other.Name = "Other";
            // 
            // LoginGroup
            // 
            this.LoginGroup.Items.Add(this.usernameBox);
            this.LoginGroup.Items.Add(this.passwordBox);
            this.LoginGroup.Items.Add(this.LogInButton);
            this.LoginGroup.Items.Add(this.LogOutButton);
            this.LoginGroup.Label = "Login";
            this.LoginGroup.Name = "LoginGroup";
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
            // GoToWebPlatformButton
            // 
            this.GoToWebPlatformButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.GoToWebPlatformButton.Image = ((System.Drawing.Image)(resources.GetObject("GoToWebPlatformButton.Image")));
            this.GoToWebPlatformButton.Label = "Go to Web Platform";
            this.GoToWebPlatformButton.Name = "GoToWebPlatformButton";
            this.GoToWebPlatformButton.ShowImage = true;
            this.GoToWebPlatformButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GoToWebPlatformButton_Click);
            // 
            // LogInButton
            // 
            this.LogInButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.LogInButton.Image = ((System.Drawing.Image)(resources.GetObject("LogInButton.Image")));
            this.LogInButton.Label = "Log in";
            this.LogInButton.Name = "LogInButton";
            this.LogInButton.ShowImage = true;
            this.LogInButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LogInButton_Click);
            // 
            // LogOutButton
            // 
            this.LogOutButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.LogOutButton.Image = ((System.Drawing.Image)(resources.GetObject("LogOutButton.Image")));
            this.LogOutButton.Label = "Log out";
            this.LogOutButton.Name = "LogOutButton";
            this.LogOutButton.ShowImage = true;
            this.LogOutButton.Visible = false;
            this.LogOutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LogOutButton_Click);
            // 
            // TestButton
            // 
            this.TestButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.TestButton.Label = "Test";
            this.TestButton.Name = "TestButton";
            this.TestButton.ShowImage = true;
            this.TestButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestButton_Click);
            // 
            // usernameBox
            // 
            this.usernameBox.Label = "username";
            this.usernameBox.Name = "usernameBox";
            // 
            // passwordBox
            // 
            this.passwordBox.Label = "password";
            this.passwordBox.Name = "passwordBox";
            // 
            // BlueberryRibbon
            // 
            this.Name = "BlueberryRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.BluberryTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.BluberryTab.ResumeLayout(false);
            this.BluberryTab.PerformLayout();
            this.ArgumentsGroup.ResumeLayout(false);
            this.ArgumentsGroup.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.DownloadGroup.ResumeLayout(false);
            this.DownloadGroup.PerformLayout();
            this.WebPlatformGroup.ResumeLayout(false);
            this.WebPlatformGroup.PerformLayout();
            this.Other.ResumeLayout(false);
            this.Other.PerformLayout();
            this.LoginGroup.ResumeLayout(false);
            this.LoginGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab BluberryTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton publishButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup DownloadGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton downloadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox IDBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox FetchConfigurationCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton refreshButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Other;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ArgumentsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup WebPlatformGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GoToWebPlatformButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup LoginGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LogInButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LogOutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox usernameBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox passwordBox;
    }

    partial class ThisRibbonCollection
    {
        internal BlueberryRibbon Ribbon1
        {
            get { return this.GetRibbon<BlueberryRibbon>(); }
        }
    }
}
