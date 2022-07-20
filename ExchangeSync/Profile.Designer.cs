namespace ExchangeSync
{
    partial class Profile
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Profile));
            this.label1 = new System.Windows.Forms.Label();
            this.cmbProfileName = new System.Windows.Forms.ComboBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOptions = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkSetDefault = new System.Windows.Forms.CheckBox();
            this.txtAlternateProfile = new System.Windows.Forms.TextBox();
            this.chkAlternateProfile = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Profile Name:";
            // 
            // cmbProfileName
            // 
            this.cmbProfileName.DisplayMember = "Value";
            this.cmbProfileName.FormattingEnabled = true;
            this.cmbProfileName.Location = new System.Drawing.Point(89, 13);
            this.cmbProfileName.Name = "cmbProfileName";
            this.cmbProfileName.Size = new System.Drawing.Size(218, 21);
            this.cmbProfileName.TabIndex = 1;
            this.cmbProfileName.ValueMember = "Key";
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(89, 77);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(71, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(166, 77);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(72, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOptions
            // 
            this.btnOptions.Location = new System.Drawing.Point(244, 77);
            this.btnOptions.Name = "btnOptions";
            this.btnOptions.Size = new System.Drawing.Size(71, 23);
            this.btnOptions.TabIndex = 4;
            this.btnOptions.Text = "Options > >";
            this.btnOptions.UseVisualStyleBackColor = true;
            this.btnOptions.Click += new System.EventHandler(this.btnOptions_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkAlternateProfile);
            this.groupBox1.Controls.Add(this.chkSetDefault);
            this.groupBox1.Location = new System.Drawing.Point(13, 117);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(302, 47);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Options";
            // 
            // chkSetDefault
            // 
            this.chkSetDefault.AutoSize = true;
            this.chkSetDefault.Location = new System.Drawing.Point(7, 20);
            this.chkSetDefault.Name = "chkSetDefault";
            this.chkSetDefault.Size = new System.Drawing.Size(122, 17);
            this.chkSetDefault.TabIndex = 0;
            this.chkSetDefault.Text = "Set as default profile";
            this.chkSetDefault.UseVisualStyleBackColor = true;
            // 
            // txtAlternateProfile
            // 
            this.txtAlternateProfile.Location = new System.Drawing.Point(89, 14);
            this.txtAlternateProfile.Name = "txtAlternateProfile";
            this.txtAlternateProfile.Size = new System.Drawing.Size(218, 20);
            this.txtAlternateProfile.TabIndex = 6;
            // 
            // chkAlternateProfile
            // 
            this.chkAlternateProfile.AutoSize = true;
            this.chkAlternateProfile.Location = new System.Drawing.Point(168, 19);
            this.chkAlternateProfile.Name = "chkAlternateProfile";
            this.chkAlternateProfile.Size = new System.Drawing.Size(126, 17);
            this.chkAlternateProfile.TabIndex = 1;
            this.chkAlternateProfile.Text = "Enter alternate profile";
            this.chkAlternateProfile.UseVisualStyleBackColor = true;
            this.chkAlternateProfile.CheckedChanged += new System.EventHandler(this.chkAlternateProfile_CheckedChanged);
            // 
            // Profile
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(319, 110);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnOptions);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.cmbProfileName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtAlternateProfile);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Profile";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Choose Profile";
            this.Load += new System.EventHandler(this.Profile_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbProfileName;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOptions;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chkSetDefault;
        private System.Windows.Forms.CheckBox chkAlternateProfile;
        private System.Windows.Forms.TextBox txtAlternateProfile;
    }
}

