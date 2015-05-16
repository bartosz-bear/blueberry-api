namespace ExcelAddIn1
{
    partial class BlueberryTaskPane
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.NameTextBox = new System.Windows.Forms.TextBox();
            this.NameLabel = new System.Windows.Forms.Label();
            this.DescriptionLabel = new System.Windows.Forms.Label();
            this.DescriptionTextBox = new System.Windows.Forms.TextBox();
            this.OrganizationLabel = new System.Windows.Forms.Label();
            this.OrganizationTextBox = new System.Windows.Forms.TextBox();
            this.DataOwnerLabel = new System.Windows.Forms.Label();
            this.DataOwnerTextBox = new System.Windows.Forms.TextBox();
            this.PublishButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // NameTextBox
            // 
            this.NameTextBox.Location = new System.Drawing.Point(12, 31);
            this.NameTextBox.Name = "NameTextBox";
            this.NameTextBox.Size = new System.Drawing.Size(122, 20);
            this.NameTextBox.TabIndex = 0;
            // 
            // NameLabel
            // 
            this.NameLabel.AutoSize = true;
            this.NameLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(174)))), ((int)(((byte)(236)))), ((int)(((byte)(251)))));
            this.NameLabel.Location = new System.Drawing.Point(9, 15);
            this.NameLabel.Name = "NameLabel";
            this.NameLabel.Size = new System.Drawing.Size(35, 13);
            this.NameLabel.TabIndex = 1;
            this.NameLabel.Text = "Name";
            // 
            // DescriptionLabel
            // 
            this.DescriptionLabel.AutoSize = true;
            this.DescriptionLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(174)))), ((int)(((byte)(236)))), ((int)(((byte)(251)))));
            this.DescriptionLabel.Location = new System.Drawing.Point(9, 54);
            this.DescriptionLabel.Name = "DescriptionLabel";
            this.DescriptionLabel.Size = new System.Drawing.Size(60, 13);
            this.DescriptionLabel.TabIndex = 2;
            this.DescriptionLabel.Text = "Description";
            // 
            // DescriptionTextBox
            // 
            this.DescriptionTextBox.Location = new System.Drawing.Point(12, 70);
            this.DescriptionTextBox.Multiline = true;
            this.DescriptionTextBox.Name = "DescriptionTextBox";
            this.DescriptionTextBox.Size = new System.Drawing.Size(122, 63);
            this.DescriptionTextBox.TabIndex = 3;
            // 
            // OrganizationLabel
            // 
            this.OrganizationLabel.AutoSize = true;
            this.OrganizationLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(174)))), ((int)(((byte)(236)))), ((int)(((byte)(251)))));
            this.OrganizationLabel.Location = new System.Drawing.Point(9, 136);
            this.OrganizationLabel.Name = "OrganizationLabel";
            this.OrganizationLabel.Size = new System.Drawing.Size(66, 13);
            this.OrganizationLabel.TabIndex = 4;
            this.OrganizationLabel.Text = "Organization";
            // 
            // OrganizationTextBox
            // 
            this.OrganizationTextBox.Location = new System.Drawing.Point(12, 152);
            this.OrganizationTextBox.Name = "OrganizationTextBox";
            this.OrganizationTextBox.Size = new System.Drawing.Size(122, 20);
            this.OrganizationTextBox.TabIndex = 5;
            // 
            // DataOwnerLabel
            // 
            this.DataOwnerLabel.AutoSize = true;
            this.DataOwnerLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(174)))), ((int)(((byte)(236)))), ((int)(((byte)(251)))));
            this.DataOwnerLabel.Location = new System.Drawing.Point(9, 175);
            this.DataOwnerLabel.Name = "DataOwnerLabel";
            this.DataOwnerLabel.Size = new System.Drawing.Size(62, 13);
            this.DataOwnerLabel.TabIndex = 6;
            this.DataOwnerLabel.Text = "Data owner";
            // 
            // DataOwnerTextBox
            // 
            this.DataOwnerTextBox.Location = new System.Drawing.Point(12, 191);
            this.DataOwnerTextBox.Name = "DataOwnerTextBox";
            this.DataOwnerTextBox.Size = new System.Drawing.Size(122, 20);
            this.DataOwnerTextBox.TabIndex = 7;
            // 
            // PublishButton
            // 
            this.PublishButton.Location = new System.Drawing.Point(12, 217);
            this.PublishButton.Name = "PublishButton";
            this.PublishButton.Size = new System.Drawing.Size(75, 23);
            this.PublishButton.TabIndex = 8;
            this.PublishButton.Text = "Publish";
            this.PublishButton.UseVisualStyleBackColor = true;
            this.PublishButton.Click += new System.EventHandler(this.PublishButton_Click);
            // 
            // BlueberryTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(35)))), ((int)(((byte)(90)))));
            this.Controls.Add(this.PublishButton);
            this.Controls.Add(this.DataOwnerTextBox);
            this.Controls.Add(this.DataOwnerLabel);
            this.Controls.Add(this.OrganizationTextBox);
            this.Controls.Add(this.OrganizationLabel);
            this.Controls.Add(this.DescriptionTextBox);
            this.Controls.Add(this.DescriptionLabel);
            this.Controls.Add(this.NameLabel);
            this.Controls.Add(this.NameTextBox);
            this.Name = "BlueberryTaskPane";
            this.Size = new System.Drawing.Size(150, 282);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox NameTextBox;

        public System.Windows.Forms.TextBox PublishingNameTextBox
        {
            get { return NameTextBox; }
            set { NameTextBox = value; }
        }
        private System.Windows.Forms.Label NameLabel;

        public System.Windows.Forms.Label PublishingNameLabel
        {
            get { return NameLabel; }
            set { NameLabel = value; }
        }
        private System.Windows.Forms.Label DescriptionLabel;
        private System.Windows.Forms.TextBox DescriptionTextBox;

        public System.Windows.Forms.TextBox PublishingDescriptionTextBox
        {
            get { return DescriptionTextBox; }
            set { DescriptionTextBox = value; }
        }
        private System.Windows.Forms.Label OrganizationLabel;
        private System.Windows.Forms.TextBox OrganizationTextBox;

        public System.Windows.Forms.TextBox PublishingOrganizationTextBox
        {
            get { return OrganizationTextBox; }
            set { OrganizationTextBox = value; }
        }
        private System.Windows.Forms.Label DataOwnerLabel;
        private System.Windows.Forms.TextBox DataOwnerTextBox;

        public System.Windows.Forms.TextBox PublishingDataOwnerTextBox
        {
            get { return DataOwnerTextBox; }
            set { DataOwnerTextBox = value; }
        }
        private System.Windows.Forms.Button PublishButton;
    }
}
