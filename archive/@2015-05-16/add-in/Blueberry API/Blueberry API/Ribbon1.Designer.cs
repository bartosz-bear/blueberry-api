namespace Blueberry_API
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
            this.BlueberryRibbon = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.publish = this.Factory.CreateRibbonButton();
            this.BlueberryRibbon.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // BlueberryRibbon
            // 
            this.BlueberryRibbon.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.BlueberryRibbon.Groups.Add(this.group1);
            this.BlueberryRibbon.Label = "Blueberry API";
            this.BlueberryRibbon.Name = "BlueberryRibbon";
            // 
            // group1
            // 
            this.group1.Items.Add(this.publish);
            this.group1.Label = "Publishing";
            this.group1.Name = "group1";
            // 
            // publish
            // 
            this.publish.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.publish.Image = global::Blueberry_API.Properties.Resources.publish_icon;
            this.publish.Label = "Publish";
            this.publish.Name = "publish";
            this.publish.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.BlueberryRibbon);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.BlueberryRibbon.ResumeLayout(false);
            this.BlueberryRibbon.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab BlueberryRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton publish;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
