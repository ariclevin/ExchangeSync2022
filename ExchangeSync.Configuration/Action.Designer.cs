namespace ExchangeSync
{
    partial class StartAction
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
            this.rbInitialConfiguration = new System.Windows.Forms.RadioButton();
            this.rbNewProfile = new System.Windows.Forms.RadioButton();
            this.rbEditProfile = new System.Windows.Forms.RadioButton();
            this.cmbProfileName = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // rbInitialConfiguration
            // 
            this.rbInitialConfiguration.AutoSize = true;
            this.rbInitialConfiguration.Location = new System.Drawing.Point(24, 23);
            this.rbInitialConfiguration.Margin = new System.Windows.Forms.Padding(6);
            this.rbInitialConfiguration.Name = "rbInitialConfiguration";
            this.rbInitialConfiguration.Size = new System.Drawing.Size(390, 29);
            this.rbInitialConfiguration.TabIndex = 0;
            this.rbInitialConfiguration.TabStop = true;
            this.rbInitialConfiguration.Text = "Create/Update Default Configuration";
            this.rbInitialConfiguration.UseVisualStyleBackColor = true;
            this.rbInitialConfiguration.CheckedChanged += new System.EventHandler(this.rbInitialConfiguration_CheckedChanged);
            // 
            // rbNewProfile
            // 
            this.rbNewProfile.AutoSize = true;
            this.rbNewProfile.Location = new System.Drawing.Point(24, 67);
            this.rbNewProfile.Margin = new System.Windows.Forms.Padding(6);
            this.rbNewProfile.Name = "rbNewProfile";
            this.rbNewProfile.Size = new System.Drawing.Size(222, 29);
            this.rbNewProfile.TabIndex = 1;
            this.rbNewProfile.TabStop = true;
            this.rbNewProfile.Text = "Create New Profile";
            this.rbNewProfile.UseVisualStyleBackColor = true;
            this.rbNewProfile.CheckedChanged += new System.EventHandler(this.rbNewProfile_CheckedChanged);
            // 
            // rbEditProfile
            // 
            this.rbEditProfile.AutoSize = true;
            this.rbEditProfile.Location = new System.Drawing.Point(24, 112);
            this.rbEditProfile.Margin = new System.Windows.Forms.Padding(6);
            this.rbEditProfile.Name = "rbEditProfile";
            this.rbEditProfile.Size = new System.Drawing.Size(229, 29);
            this.rbEditProfile.TabIndex = 2;
            this.rbEditProfile.TabStop = true;
            this.rbEditProfile.Text = "Edit Existing Profile";
            this.rbEditProfile.UseVisualStyleBackColor = true;
            this.rbEditProfile.CheckedChanged += new System.EventHandler(this.rbEditProfile_CheckedChanged);
            // 
            // cmbProfileName
            // 
            this.cmbProfileName.DisplayMember = "Value";
            this.cmbProfileName.Enabled = false;
            this.cmbProfileName.FormattingEnabled = true;
            this.cmbProfileName.Location = new System.Drawing.Point(246, 156);
            this.cmbProfileName.Margin = new System.Windows.Forms.Padding(6);
            this.cmbProfileName.Name = "cmbProfileName";
            this.cmbProfileName.Size = new System.Drawing.Size(294, 33);
            this.cmbProfileName.TabIndex = 4;
            this.cmbProfileName.ValueMember = "Key";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(94, 156);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(141, 25);
            this.label1.TabIndex = 3;
            this.label1.Text = "Profile Name:";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(400, 233);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(6);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(144, 44);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(246, 233);
            this.btnOK.Margin = new System.Windows.Forms.Padding(6);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(142, 44);
            this.btnOK.TabIndex = 5;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // StartAction
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(568, 300);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.cmbProfileName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.rbEditProfile);
            this.Controls.Add(this.rbNewProfile);
            this.Controls.Add(this.rbInitialConfiguration);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "StartAction";
            this.Text = "Select Action";
            this.Load += new System.EventHandler(this.StartAction_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton rbInitialConfiguration;
        private System.Windows.Forms.RadioButton rbNewProfile;
        private System.Windows.Forms.RadioButton rbEditProfile;
        private System.Windows.Forms.ComboBox cmbProfileName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
    }
}

