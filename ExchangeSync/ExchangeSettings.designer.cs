namespace ExchangeSync
{
    partial class ExchangeSettings
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
            this.components = new System.ComponentModel.Container();
            Infragistics.Win.UltraWinToolTip.UltraToolTipInfo ultraToolTipInfo3 = new Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Enter the # sign in the Exchange Contact Alias or Exchange Contact Name fields fo" +
        "r the Custom Field Value", Infragistics.Win.ToolTipImage.Info, null, Infragistics.Win.DefaultableBoolean.Default);
            Infragistics.Win.UltraWinToolTip.UltraToolTipInfo ultraToolTipInfo2 = new Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Enter the # sign in the Exchange Contact Alias or Exchange Contact Name fields fo" +
        "r the Custom Field Value", Infragistics.Win.ToolTipImage.Info, null, Infragistics.Win.DefaultableBoolean.Default);
            Infragistics.Win.UltraWinToolTip.UltraToolTipInfo ultraToolTipInfo1 = new Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Internal Recipient Type for use with Office 365 and ADSync", Infragistics.Win.ToolTipImage.Info, "Internal Recipient Type", Infragistics.Win.DefaultableBoolean.Default);
            this.label1 = new System.Windows.Forms.Label();
            this.txtServer = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbVersion = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbAliasType = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtAliasValue = new System.Windows.Forms.TextBox();
            this.txtNameValue = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.cmbNameType = new System.Windows.Forms.ComboBox();
            this.uttmExchangeSettings = new Infragistics.Win.UltraWinToolTip.UltraToolTipManager(this.components);
            this.label13 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.cmbAutoUpdatePolicy = new System.Windows.Forms.ComboBox();
            this.btnSaveSettings = new System.Windows.Forms.Button();
            this.txtDomain = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.uceDistributionGroups = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
            this.uceContacts = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
            this.panel3 = new System.Windows.Forms.Panel();
            this.rbMailEnabled = new System.Windows.Forms.RadioButton();
            this.rbMailboxEnabled = new System.Windows.Forms.RadioButton();
            this.lblConfirmation = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.uceDistributionGroups)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.uceContacts)).BeginInit();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Exchange Server Name";
            // 
            // txtServer
            // 
            this.txtServer.Location = new System.Drawing.Point(7, 30);
            this.txtServer.Name = "txtServer";
            this.txtServer.Size = new System.Drawing.Size(324, 20);
            this.txtServer.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(336, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(42, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Version";
            // 
            // cmbVersion
            // 
            this.cmbVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbVersion.FormattingEnabled = true;
            this.cmbVersion.Items.AddRange(new object[] {
            "2010",
            "2013",
            "Online"});
            this.cmbVersion.Location = new System.Drawing.Point(339, 29);
            this.cmbVersion.Name = "cmbVersion";
            this.cmbVersion.Size = new System.Drawing.Size(126, 21);
            this.cmbVersion.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 53);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Username";
            // 
            // txtUsername
            // 
            this.txtUsername.Location = new System.Drawing.Point(7, 69);
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.Size = new System.Drawing.Size(324, 20);
            this.txtUsername.TabIndex = 6;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(4, 102);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(183, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Distribution Group Organizational Unit";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(4, 144);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(178, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Default Contacts Organizational Unit";
            // 
            // cmbAliasType
            // 
            this.cmbAliasType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAliasType.FormattingEnabled = true;
            this.cmbAliasType.Items.AddRange(new object[] {
            "Prefix",
            "Suffix",
            "Separator",
            "Append Domain"});
            this.cmbAliasType.Location = new System.Drawing.Point(239, 204);
            this.cmbAliasType.Name = "cmbAliasType";
            this.cmbAliasType.Size = new System.Drawing.Size(94, 21);
            this.cmbAliasType.TabIndex = 17;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(236, 188);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(120, 13);
            this.label7.TabIndex = 18;
            this.label7.Text = "Exchange Contact Alias";
            // 
            // txtAliasValue
            // 
            this.txtAliasValue.Location = new System.Drawing.Point(339, 205);
            this.txtAliasValue.Name = "txtAliasValue";
            this.txtAliasValue.Size = new System.Drawing.Size(126, 20);
            this.txtAliasValue.TabIndex = 19;
            ultraToolTipInfo3.ToolTipImage = Infragistics.Win.ToolTipImage.Info;
            ultraToolTipInfo3.ToolTipText = "Enter the # sign in the Exchange Contact Alias or Exchange Contact Name fields fo" +
    "r the Custom Field Value";
            this.uttmExchangeSettings.SetUltraToolTip(this.txtAliasValue, ultraToolTipInfo3);
            // 
            // txtNameValue
            // 
            this.txtNameValue.Location = new System.Drawing.Point(107, 205);
            this.txtNameValue.Name = "txtNameValue";
            this.txtNameValue.Size = new System.Drawing.Size(126, 20);
            this.txtNameValue.TabIndex = 22;
            ultraToolTipInfo2.ToolTipImage = Infragistics.Win.ToolTipImage.Info;
            ultraToolTipInfo2.ToolTipText = "Enter the # sign in the Exchange Contact Alias or Exchange Contact Name fields fo" +
    "r the Custom Field Value";
            this.uttmExchangeSettings.SetUltraToolTip(this.txtNameValue, ultraToolTipInfo2);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(4, 188);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(126, 13);
            this.label8.TabIndex = 21;
            this.label8.Text = "Exchange Contact Name";
            // 
            // cmbNameType
            // 
            this.cmbNameType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbNameType.FormattingEnabled = true;
            this.cmbNameType.Items.AddRange(new object[] {
            "Prefix",
            "Suffix",
            "Separator",
            "Append Domain"});
            this.cmbNameType.Location = new System.Drawing.Point(7, 204);
            this.cmbNameType.Name = "cmbNameType";
            this.cmbNameType.Size = new System.Drawing.Size(94, 21);
            this.cmbNameType.TabIndex = 20;
            this.cmbNameType.SelectedIndexChanged += new System.EventHandler(this.cmbNameType_SelectedIndexChanged);
            // 
            // uttmExchangeSettings
            // 
            this.uttmExchangeSettings.ContainingControl = this;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(7, 282);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(117, 13);
            this.label13.TabIndex = 46;
            this.label13.Text = "Internal Recipient Type";
            ultraToolTipInfo1.ToolTipImage = Infragistics.Win.ToolTipImage.Info;
            ultraToolTipInfo1.ToolTipText = "Internal Recipient Type for use with Office 365 and ADSync";
            ultraToolTipInfo1.ToolTipTextFormatted = "Internal Recipient Type for use with Office 365 and ADSync";
            ultraToolTipInfo1.ToolTipTitle = "Internal Recipient Type";
            this.uttmExchangeSettings.SetUltraToolTip(this.label13, ultraToolTipInfo1);
            this.label13.Visible = false;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(7, 234);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(293, 13);
            this.label11.TabIndex = 39;
            this.label11.Text = "Automatically update e-mail addresses based on e-mail policy";
            // 
            // cmbAutoUpdatePolicy
            // 
            this.cmbAutoUpdatePolicy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAutoUpdatePolicy.FormattingEnabled = true;
            this.cmbAutoUpdatePolicy.Items.AddRange(new object[] {
            "Yes",
            "No",
            "Ignore"});
            this.cmbAutoUpdatePolicy.Location = new System.Drawing.Point(10, 251);
            this.cmbAutoUpdatePolicy.Name = "cmbAutoUpdatePolicy";
            this.cmbAutoUpdatePolicy.Size = new System.Drawing.Size(121, 21);
            this.cmbAutoUpdatePolicy.TabIndex = 40;
            // 
            // btnSaveSettings
            // 
            this.btnSaveSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSaveSettings.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveSettings.Image = global::ExchangeSync.Properties.Resources.save32;
            this.btnSaveSettings.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSaveSettings.Location = new System.Drawing.Point(267, 370);
            this.btnSaveSettings.Name = "btnSaveSettings";
            this.btnSaveSettings.Size = new System.Drawing.Size(198, 39);
            this.btnSaveSettings.TabIndex = 24;
            this.btnSaveSettings.Text = "Save Settings";
            this.btnSaveSettings.UseVisualStyleBackColor = true;
            this.btnSaveSettings.Click += new System.EventHandler(this.btnSaveSettings_Click);
            // 
            // txtDomain
            // 
            this.txtDomain.Location = new System.Drawing.Point(339, 69);
            this.txtDomain.Name = "txtDomain";
            this.txtDomain.Size = new System.Drawing.Size(126, 20);
            this.txtDomain.TabIndex = 43;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(336, 53);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(43, 13);
            this.label12.TabIndex = 42;
            this.label12.Text = "Domain";
            // 
            // uceDistributionGroups
            // 
            this.uceDistributionGroups.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button;
            this.uceDistributionGroups.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
            this.uceDistributionGroups.Location = new System.Drawing.Point(7, 117);
            this.uceDistributionGroups.Name = "uceDistributionGroups";
            this.uceDistributionGroups.Size = new System.Drawing.Size(458, 21);
            this.uceDistributionGroups.TabIndex = 44;
            // 
            // uceContacts
            // 
            this.uceContacts.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button;
            this.uceContacts.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
            this.uceContacts.Location = new System.Drawing.Point(7, 159);
            this.uceContacts.Name = "uceContacts";
            this.uceContacts.Size = new System.Drawing.Size(458, 21);
            this.uceContacts.TabIndex = 45;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.rbMailEnabled);
            this.panel3.Controls.Add(this.rbMailboxEnabled);
            this.panel3.Location = new System.Drawing.Point(10, 296);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(204, 50);
            this.panel3.TabIndex = 47;
            this.panel3.Visible = false;
            // 
            // rbMailEnabled
            // 
            this.rbMailEnabled.AutoSize = true;
            this.rbMailEnabled.Enabled = false;
            this.rbMailEnabled.Location = new System.Drawing.Point(3, 26);
            this.rbMailEnabled.Name = "rbMailEnabled";
            this.rbMailEnabled.Size = new System.Drawing.Size(86, 17);
            this.rbMailEnabled.TabIndex = 32;
            this.rbMailEnabled.TabStop = true;
            this.rbMailEnabled.Tag = "INTERNALRECIPIENTTYPEMAILENABLED";
            this.rbMailEnabled.Text = "Mail Enabled";
            this.rbMailEnabled.UseVisualStyleBackColor = true;
            // 
            // rbMailboxEnabled
            // 
            this.rbMailboxEnabled.AutoSize = true;
            this.rbMailboxEnabled.Enabled = false;
            this.rbMailboxEnabled.Location = new System.Drawing.Point(3, 3);
            this.rbMailboxEnabled.Name = "rbMailboxEnabled";
            this.rbMailboxEnabled.Size = new System.Drawing.Size(103, 17);
            this.rbMailboxEnabled.TabIndex = 31;
            this.rbMailboxEnabled.TabStop = true;
            this.rbMailboxEnabled.Tag = "INTERNALRECIPIENTTYPEMAILBOX";
            this.rbMailboxEnabled.Text = "Mailbox Enabled";
            this.rbMailboxEnabled.UseVisualStyleBackColor = true;
            // 
            // lblConfirmation
            // 
            this.lblConfirmation.Location = new System.Drawing.Point(22, 412);
            this.lblConfirmation.Name = "lblConfirmation";
            this.lblConfirmation.Size = new System.Drawing.Size(443, 23);
            this.lblConfirmation.TabIndex = 52;
            // 
            // ExchangeSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblConfirmation);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.uceContacts);
            this.Controls.Add(this.uceDistributionGroups);
            this.Controls.Add(this.txtDomain);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.cmbAutoUpdatePolicy);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.txtNameValue);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.cmbNameType);
            this.Controls.Add(this.txtAliasValue);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.cmbAliasType);
            this.Controls.Add(this.btnSaveSettings);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtUsername);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmbVersion);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtServer);
            this.Controls.Add(this.label1);
            this.Name = "ExchangeSettings";
            this.Size = new System.Drawing.Size(480, 450);
            this.Load += new System.EventHandler(this.ExchangeSettings_Load);
            ((System.ComponentModel.ISupportInitialize)(this.uceDistributionGroups)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.uceContacts)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtServer;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbVersion;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtUsername;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnSaveSettings;
        private System.Windows.Forms.ComboBox cmbAliasType;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtAliasValue;
        private System.Windows.Forms.TextBox txtNameValue;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cmbNameType;
        private Infragistics.Win.UltraWinToolTip.UltraToolTipManager uttmExchangeSettings;
        private System.Windows.Forms.ComboBox cmbAutoUpdatePolicy;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtDomain;
        private System.Windows.Forms.Label label12;
        private Infragistics.Win.UltraWinEditors.UltraComboEditor uceDistributionGroups;
        private Infragistics.Win.UltraWinEditors.UltraComboEditor uceContacts;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.RadioButton rbMailEnabled;
        private System.Windows.Forms.RadioButton rbMailboxEnabled;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label lblConfirmation;
    }
}
