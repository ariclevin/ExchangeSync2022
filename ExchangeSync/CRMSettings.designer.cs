namespace ExchangeSync
{
    partial class CRMSettings
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
            this.txtConnectionUrl = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtCustomField = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtRevisionField = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtExchangeAliasField = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtRequiredAttributes = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label26 = new System.Windows.Forms.Label();
            this.cmbAuthenticationType = new System.Windows.Forms.ComboBox();
            this.label27 = new System.Windows.Forms.Label();
            this.cmbServiceEndpoint = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkIntegrated = new System.Windows.Forms.CheckBox();
            this.label23 = new System.Windows.Forms.Label();
            this.txtCRMDomain = new System.Windows.Forms.TextBox();
            this.txtCRMUsername = new System.Windows.Forms.TextBox();
            this.label25 = new System.Windows.Forms.Label();
            this.btnUpdateSolution = new System.Windows.Forms.Button();
            this.btnSaveSettings = new System.Windows.Forms.Button();
            this.lblConfirmation = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtConnectionUrl
            // 
            this.txtConnectionUrl.Location = new System.Drawing.Point(3, 68);
            this.txtConnectionUrl.Name = "txtConnectionUrl";
            this.txtConnectionUrl.Size = new System.Drawing.Size(462, 20);
            this.txtConnectionUrl.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(138, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "CRM Server Connection Url";
            // 
            // txtCustomField
            // 
            this.txtCustomField.Location = new System.Drawing.Point(233, 369);
            this.txtCustomField.Name = "txtCustomField";
            this.txtCustomField.Size = new System.Drawing.Size(226, 20);
            this.txtCustomField.TabIndex = 26;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(230, 353);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(146, 13);
            this.label9.TabIndex = 25;
            this.label9.Text = "Custom Field Name (Optional)";
            // 
            // txtRevisionField
            // 
            this.txtRevisionField.Location = new System.Drawing.Point(17, 369);
            this.txtRevisionField.Name = "txtRevisionField";
            this.txtRevisionField.Size = new System.Drawing.Size(206, 20);
            this.txtRevisionField.TabIndex = 34;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 353);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(169, 13);
            this.label2.TabIndex = 27;
            this.label2.Text = "Revision No Field Name (Optional)";
            // 
            // txtExchangeAliasField
            // 
            this.txtExchangeAliasField.Location = new System.Drawing.Point(16, 330);
            this.txtExchangeAliasField.Name = "txtExchangeAliasField";
            this.txtExchangeAliasField.Size = new System.Drawing.Size(443, 20);
            this.txtExchangeAliasField.TabIndex = 30;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 314);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(136, 13);
            this.label3.TabIndex = 29;
            this.label3.Text = "Exchange Alias Field Name";
            // 
            // txtRequiredAttributes
            // 
            this.txtRequiredAttributes.Location = new System.Drawing.Point(16, 291);
            this.txtRequiredAttributes.Name = "txtRequiredAttributes";
            this.txtRequiredAttributes.Size = new System.Drawing.Size(443, 20);
            this.txtRequiredAttributes.TabIndex = 20;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 275);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(231, 13);
            this.label4.TabIndex = 31;
            this.label4.Text = "Required Attributes - Separated by semicolon (;)";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(3, 9);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(102, 13);
            this.label26.TabIndex = 46;
            this.label26.Text = "Authentication Type";
            // 
            // cmbAuthenticationType
            // 
            this.cmbAuthenticationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAuthenticationType.FormattingEnabled = true;
            this.cmbAuthenticationType.Items.AddRange(new object[] {
            "Active Directory (On-Premise)",
            "Federated (CRM On-Premise or Partner-Hosted CRM)",
            "Live Id (Legacy CRM Online)",
            "Online Federated (Office 365 CRM Online)"});
            this.cmbAuthenticationType.Location = new System.Drawing.Point(3, 25);
            this.cmbAuthenticationType.Name = "cmbAuthenticationType";
            this.cmbAuthenticationType.Size = new System.Drawing.Size(462, 21);
            this.cmbAuthenticationType.TabIndex = 45;
            this.cmbAuthenticationType.SelectedIndexChanged += new System.EventHandler(this.cmbAuthenticationType_SelectedIndexChanged);
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(3, 91);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(88, 13);
            this.label27.TabIndex = 48;
            this.label27.Text = "Service Endpoint";
            // 
            // cmbServiceEndpoint
            // 
            this.cmbServiceEndpoint.Enabled = false;
            this.cmbServiceEndpoint.FormattingEnabled = true;
            this.cmbServiceEndpoint.Items.AddRange(new object[] {
            "SOAP 2011 (CRM 2011, CRM 2013, CRM 2015, CRM 2016)"});
            this.cmbServiceEndpoint.Location = new System.Drawing.Point(6, 107);
            this.cmbServiceEndpoint.Name = "cmbServiceEndpoint";
            this.cmbServiceEndpoint.Size = new System.Drawing.Size(459, 21);
            this.cmbServiceEndpoint.TabIndex = 47;
            this.cmbServiceEndpoint.Text = "SOAP 2011 (CRM 2011, CRM 2013, CRM 2015, CRM 2016)";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.chkIntegrated);
            this.groupBox2.Controls.Add(this.label23);
            this.groupBox2.Controls.Add(this.txtCRMDomain);
            this.groupBox2.Controls.Add(this.txtCRMUsername);
            this.groupBox2.Controls.Add(this.label25);
            this.groupBox2.Location = new System.Drawing.Point(7, 138);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(458, 122);
            this.groupBox2.TabIndex = 49;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Authentication";
            // 
            // chkIntegrated
            // 
            this.chkIntegrated.AutoSize = true;
            this.chkIntegrated.Location = new System.Drawing.Point(9, 20);
            this.chkIntegrated.Name = "chkIntegrated";
            this.chkIntegrated.Size = new System.Drawing.Size(167, 17);
            this.chkIntegrated.TabIndex = 43;
            this.chkIntegrated.Text = "Use Integrated Authentication";
            this.chkIntegrated.UseVisualStyleBackColor = true;
            this.chkIntegrated.CheckedChanged += new System.EventHandler(this.chkIntegrated_CheckedChanged);
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label23.Location = new System.Drawing.Point(6, 78);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(43, 13);
            this.label23.TabIndex = 41;
            this.label23.Text = "Domain";
            // 
            // txtCRMDomain
            // 
            this.txtCRMDomain.Enabled = false;
            this.txtCRMDomain.Location = new System.Drawing.Point(132, 71);
            this.txtCRMDomain.Name = "txtCRMDomain";
            this.txtCRMDomain.PasswordChar = '*';
            this.txtCRMDomain.Size = new System.Drawing.Size(320, 20);
            this.txtCRMDomain.TabIndex = 42;
            // 
            // txtCRMUsername
            // 
            this.txtCRMUsername.Enabled = false;
            this.txtCRMUsername.Location = new System.Drawing.Point(132, 45);
            this.txtCRMUsername.Name = "txtCRMUsername";
            this.txtCRMUsername.Size = new System.Drawing.Size(320, 20);
            this.txtCRMUsername.TabIndex = 38;
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label25.Location = new System.Drawing.Point(6, 48);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(55, 13);
            this.label25.TabIndex = 37;
            this.label25.Text = "Username";
            // 
            // btnUpdateSolution
            // 
            this.btnUpdateSolution.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnUpdateSolution.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpdateSolution.Image = global::ExchangeSync.Properties.Resources.solutions32;
            this.btnUpdateSolution.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnUpdateSolution.Location = new System.Drawing.Point(17, 403);
            this.btnUpdateSolution.Name = "btnUpdateSolution";
            this.btnUpdateSolution.Size = new System.Drawing.Size(198, 39);
            this.btnUpdateSolution.TabIndex = 50;
            this.btnUpdateSolution.Text = "Solution Update";
            this.btnUpdateSolution.UseVisualStyleBackColor = true;
            this.btnUpdateSolution.Click += new System.EventHandler(this.btnUpdateSolution_Click);
            // 
            // btnSaveSettings
            // 
            this.btnSaveSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSaveSettings.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveSettings.Image = global::ExchangeSync.Properties.Resources.save32;
            this.btnSaveSettings.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSaveSettings.Location = new System.Drawing.Point(261, 403);
            this.btnSaveSettings.Name = "btnSaveSettings";
            this.btnSaveSettings.Size = new System.Drawing.Size(198, 39);
            this.btnSaveSettings.TabIndex = 44;
            this.btnSaveSettings.Text = "Save Settings";
            this.btnSaveSettings.UseVisualStyleBackColor = true;
            this.btnSaveSettings.Click += new System.EventHandler(this.btnSaveSettings_Click);
            // 
            // lblConfirmation
            // 
            this.lblConfirmation.Location = new System.Drawing.Point(16, 449);
            this.lblConfirmation.Name = "lblConfirmation";
            this.lblConfirmation.Size = new System.Drawing.Size(443, 23);
            this.lblConfirmation.TabIndex = 51;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(9, 101);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(409, 13);
            this.label5.TabIndex = 44;
            this.label5.Text = "Please run the Configuration/Profile Wizard to change your Credentials";
            // 
            // CRMSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblConfirmation);
            this.Controls.Add(this.btnUpdateSolution);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.label27);
            this.Controls.Add(this.cmbServiceEndpoint);
            this.Controls.Add(this.label26);
            this.Controls.Add(this.cmbAuthenticationType);
            this.Controls.Add(this.txtRequiredAttributes);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtExchangeAliasField);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtRevisionField);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtCustomField);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.btnSaveSettings);
            this.Controls.Add(this.txtConnectionUrl);
            this.Controls.Add(this.label1);
            this.Name = "CRMSettings";
            this.Size = new System.Drawing.Size(480, 476);
            this.Load += new System.EventHandler(this.CRMSettings_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtConnectionUrl;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSaveSettings;
        private System.Windows.Forms.TextBox txtCustomField;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtRevisionField;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtExchangeAliasField;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtRequiredAttributes;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.ComboBox cmbAuthenticationType;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.ComboBox cmbServiceEndpoint;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox chkIntegrated;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.TextBox txtCRMDomain;
        private System.Windows.Forms.TextBox txtCRMUsername;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Button btnUpdateSolution;
        private System.Windows.Forms.Label lblConfirmation;
        private System.Windows.Forms.Label label5;
    }
}
