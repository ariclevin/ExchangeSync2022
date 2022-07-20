namespace ExchangeSync
{
    partial class GeneralSettings
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
            Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
            Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
            Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
            Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
            this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
            this.cmbMailUserFormat = new System.Windows.Forms.ComboBox();
            this.cmbMailContactFormat = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.rbIgnore = new System.Windows.Forms.RadioButton();
            this.rbShow = new System.Windows.Forms.RadioButton();
            this.rbHide = new System.Windows.Forms.RadioButton();
            this.label4 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.rbUpdateDelta = new System.Windows.Forms.RadioButton();
            this.rbUpdateAll = new System.Windows.Forms.RadioButton();
            this.label3 = new System.Windows.Forms.Label();
            this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
            this.chkDeleteOrphanedContacts = new System.Windows.Forms.CheckBox();
            this.chkDeleteExchangeList = new System.Windows.Forms.CheckBox();
            this.chkRemoveConnection = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
            this.panel3 = new System.Windows.Forms.Panel();
            this.rbMapCRM = new System.Windows.Forms.RadioButton();
            this.rbMapMinimal = new System.Windows.Forms.RadioButton();
            this.rbMapNone = new System.Windows.Forms.RadioButton();
            this.label5 = new System.Windows.Forms.Label();
            this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
            this.panel4 = new System.Windows.Forms.Panel();
            this.rbDuplicationSyncIgnore = new System.Windows.Forms.RadioButton();
            this.rbDuplicationSyncModify = new System.Windows.Forms.RadioButton();
            this.chkDuplicateDetectionAction = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.uttmExchangeSettings = new Infragistics.Win.UltraWinToolTip.UltraToolTipManager(this.components);
            this.btnSaveSettings = new System.Windows.Forms.Button();
            this.utGeneral = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
            this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
            this.lblConfirmation = new System.Windows.Forms.Label();
            this.ultraTabPageControl1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.ultraTabPageControl2.SuspendLayout();
            this.ultraTabPageControl3.SuspendLayout();
            this.panel3.SuspendLayout();
            this.ultraTabPageControl4.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.utGeneral)).BeginInit();
            this.utGeneral.SuspendLayout();
            this.SuspendLayout();
            // 
            // ultraTabPageControl1
            // 
            this.ultraTabPageControl1.Controls.Add(this.cmbMailUserFormat);
            this.ultraTabPageControl1.Controls.Add(this.cmbMailContactFormat);
            this.ultraTabPageControl1.Controls.Add(this.label8);
            this.ultraTabPageControl1.Controls.Add(this.label7);
            this.ultraTabPageControl1.Controls.Add(this.label6);
            this.ultraTabPageControl1.Controls.Add(this.panel2);
            this.ultraTabPageControl1.Controls.Add(this.label4);
            this.ultraTabPageControl1.Controls.Add(this.panel1);
            this.ultraTabPageControl1.Controls.Add(this.label3);
            this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 20);
            this.ultraTabPageControl1.Name = "ultraTabPageControl1";
            this.ultraTabPageControl1.Size = new System.Drawing.Size(444, 331);
            // 
            // cmbMailUserFormat
            // 
            this.cmbMailUserFormat.FormattingEnabled = true;
            this.cmbMailUserFormat.Items.AddRange(new object[] {
            "No Update",
            "CRM Setting",
            "First Last",
            "Last, First"});
            this.cmbMailUserFormat.Location = new System.Drawing.Point(202, 242);
            this.cmbMailUserFormat.Name = "cmbMailUserFormat";
            this.cmbMailUserFormat.Size = new System.Drawing.Size(230, 21);
            this.cmbMailUserFormat.TabIndex = 49;
            // 
            // cmbMailContactFormat
            // 
            this.cmbMailContactFormat.FormattingEnabled = true;
            this.cmbMailContactFormat.Items.AddRange(new object[] {
            "No Update",
            "CRM Setting",
            "First Last",
            "Last, First"});
            this.cmbMailContactFormat.Location = new System.Drawing.Point(202, 215);
            this.cmbMailContactFormat.Name = "cmbMailContactFormat";
            this.cmbMailContactFormat.Size = new System.Drawing.Size(230, 21);
            this.cmbMailContactFormat.TabIndex = 48;
            this.cmbMailContactFormat.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbMailContactFormat_KeyDown);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(19, 239);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(169, 13);
            this.label8.TabIndex = 39;
            this.label8.Text = "Mail User (Mailbox) Update Format";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(19, 215);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(139, 13);
            this.label7.TabIndex = 38;
            this.label7.Text = "Mail Contact Update Format";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(16, 198);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(151, 13);
            this.label6.TabIndex = 37;
            this.label6.Text = "Exchange Naming Format";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.rbIgnore);
            this.panel2.Controls.Add(this.rbShow);
            this.panel2.Controls.Add(this.rbHide);
            this.panel2.Location = new System.Drawing.Point(16, 110);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(416, 71);
            this.panel2.TabIndex = 36;
            // 
            // rbIgnore
            // 
            this.rbIgnore.AutoSize = true;
            this.rbIgnore.Location = new System.Drawing.Point(3, 49);
            this.rbIgnore.Name = "rbIgnore";
            this.rbIgnore.Size = new System.Drawing.Size(227, 17);
            this.rbIgnore.TabIndex = 33;
            this.rbIgnore.TabStop = true;
            this.rbIgnore.Tag = "IGNORE";
            this.rbIgnore.Text = "Ignore Hide/Show Action for Mail Contacts";
            this.rbIgnore.UseVisualStyleBackColor = true;
            // 
            // rbShow
            // 
            this.rbShow.AutoSize = true;
            this.rbShow.Location = new System.Drawing.Point(3, 26);
            this.rbShow.Name = "rbShow";
            this.rbShow.Size = new System.Drawing.Size(205, 17);
            this.rbShow.TabIndex = 32;
            this.rbShow.TabStop = true;
            this.rbShow.Tag = "SHOW";
            this.rbShow.Text = "Show Mail Contacts in Exchange Lists";
            this.rbShow.UseVisualStyleBackColor = true;
            // 
            // rbHide
            // 
            this.rbHide.AutoSize = true;
            this.rbHide.Location = new System.Drawing.Point(3, 3);
            this.rbHide.Name = "rbHide";
            this.rbHide.Size = new System.Drawing.Size(212, 17);
            this.rbHide.TabIndex = 31;
            this.rbHide.TabStop = true;
            this.rbHide.Tag = "HIDE";
            this.rbHide.Text = "Hide Mail Contacts from Exchange Lists";
            this.rbHide.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(16, 94);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(118, 13);
            this.label4.TabIndex = 35;
            this.label4.Text = "Default Hide Action";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.rbUpdateDelta);
            this.panel1.Controls.Add(this.rbUpdateAll);
            this.panel1.Location = new System.Drawing.Point(16, 29);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(416, 50);
            this.panel1.TabIndex = 34;
            // 
            // rbUpdateDelta
            // 
            this.rbUpdateDelta.AutoSize = true;
            this.rbUpdateDelta.Location = new System.Drawing.Point(3, 26);
            this.rbUpdateDelta.Name = "rbUpdateDelta";
            this.rbUpdateDelta.Size = new System.Drawing.Size(332, 17);
            this.rbUpdateDelta.TabIndex = 32;
            this.rbUpdateDelta.TabStop = true;
            this.rbUpdateDelta.Tag = "DELTA";
            this.rbUpdateDelta.Text = "Update only Contact Records modified since last Synchronization";
            this.rbUpdateDelta.UseVisualStyleBackColor = true;
            // 
            // rbUpdateAll
            // 
            this.rbUpdateAll.AutoSize = true;
            this.rbUpdateAll.Location = new System.Drawing.Point(3, 3);
            this.rbUpdateAll.Name = "rbUpdateAll";
            this.rbUpdateAll.Size = new System.Drawing.Size(157, 17);
            this.rbUpdateAll.TabIndex = 31;
            this.rbUpdateAll.TabStop = true;
            this.rbUpdateAll.Tag = "ALL";
            this.rbUpdateAll.Text = "Update All Contact Records";
            this.rbUpdateAll.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(13, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(210, 13);
            this.label3.TabIndex = 33;
            this.label3.Text = "Exchange Update Action (All/Delta)";
            // 
            // ultraTabPageControl2
            // 
            this.ultraTabPageControl2.Controls.Add(this.chkDeleteOrphanedContacts);
            this.ultraTabPageControl2.Controls.Add(this.chkDeleteExchangeList);
            this.ultraTabPageControl2.Controls.Add(this.chkRemoveConnection);
            this.ultraTabPageControl2.Controls.Add(this.label1);
            this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
            this.ultraTabPageControl2.Name = "ultraTabPageControl2";
            this.ultraTabPageControl2.Size = new System.Drawing.Size(444, 331);
            // 
            // chkDeleteOrphanedContacts
            // 
            this.chkDeleteOrphanedContacts.AutoSize = true;
            this.chkDeleteOrphanedContacts.Enabled = false;
            this.chkDeleteOrphanedContacts.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.chkDeleteOrphanedContacts.Location = new System.Drawing.Point(16, 32);
            this.chkDeleteOrphanedContacts.Name = "chkDeleteOrphanedContacts";
            this.chkDeleteOrphanedContacts.Size = new System.Drawing.Size(393, 17);
            this.chkDeleteOrphanedContacts.TabIndex = 31;
            this.chkDeleteOrphanedContacts.Tag = "DELETECONTACTS";
            this.chkDeleteOrphanedContacts.Text = "Delete Exchange Mail Contacts that do not belong to other Distribution Groups";
            this.chkDeleteOrphanedContacts.UseVisualStyleBackColor = true;
            // 
            // chkDeleteExchangeList
            // 
            this.chkDeleteExchangeList.AutoSize = true;
            this.chkDeleteExchangeList.Enabled = false;
            this.chkDeleteExchangeList.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.chkDeleteExchangeList.Location = new System.Drawing.Point(16, 51);
            this.chkDeleteExchangeList.Name = "chkDeleteExchangeList";
            this.chkDeleteExchangeList.Size = new System.Drawing.Size(216, 17);
            this.chkDeleteExchangeList.TabIndex = 30;
            this.chkDeleteExchangeList.Tag = "DELETEGROUP";
            this.chkDeleteExchangeList.Text = "Delete Distribution Group from Exchange";
            this.chkDeleteExchangeList.UseVisualStyleBackColor = true;
            // 
            // chkRemoveConnection
            // 
            this.chkRemoveConnection.AutoSize = true;
            this.chkRemoveConnection.Checked = true;
            this.chkRemoveConnection.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkRemoveConnection.Location = new System.Drawing.Point(16, 70);
            this.chkRemoveConnection.Name = "chkRemoveConnection";
            this.chkRemoveConnection.Size = new System.Drawing.Size(266, 17);
            this.chkRemoveConnection.TabIndex = 29;
            this.chkRemoveConnection.Tag = "REMOVECONNECTION";
            this.chkRemoveConnection.Text = "Remove Connection between CRM and Exchange";
            this.chkRemoveConnection.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 13);
            this.label1.TabIndex = 28;
            this.label1.Text = "List Delete Action";
            // 
            // ultraTabPageControl3
            // 
            this.ultraTabPageControl3.Controls.Add(this.panel3);
            this.ultraTabPageControl3.Controls.Add(this.label5);
            this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
            this.ultraTabPageControl3.Name = "ultraTabPageControl3";
            this.ultraTabPageControl3.Size = new System.Drawing.Size(444, 331);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.rbMapCRM);
            this.panel3.Controls.Add(this.rbMapMinimal);
            this.panel3.Controls.Add(this.rbMapNone);
            this.panel3.Location = new System.Drawing.Point(13, 36);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(412, 71);
            this.panel3.TabIndex = 38;
            // 
            // rbMapCRM
            // 
            this.rbMapCRM.AutoSize = true;
            this.rbMapCRM.Location = new System.Drawing.Point(3, 49);
            this.rbMapCRM.Name = "rbMapCRM";
            this.rbMapCRM.Size = new System.Drawing.Size(248, 17);
            this.rbMapCRM.TabIndex = 33;
            this.rbMapCRM.TabStop = true;
            this.rbMapCRM.Tag = "CRM";
            this.rbMapCRM.Text = "Map Fields based on CRM Field Mapping Entity";
            this.rbMapCRM.UseVisualStyleBackColor = true;
            // 
            // rbMapMinimal
            // 
            this.rbMapMinimal.AutoSize = true;
            this.rbMapMinimal.Location = new System.Drawing.Point(3, 26);
            this.rbMapMinimal.Name = "rbMapMinimal";
            this.rbMapMinimal.Size = new System.Drawing.Size(326, 17);
            this.rbMapMinimal.TabIndex = 32;
            this.rbMapMinimal.TabStop = true;
            this.rbMapMinimal.Tag = "MIN";
            this.rbMapMinimal.Text = "Minimal Field Mapping (First/Last Name, Title, Phone, Company)";
            this.rbMapMinimal.UseVisualStyleBackColor = true;
            // 
            // rbMapNone
            // 
            this.rbMapNone.AutoSize = true;
            this.rbMapNone.Location = new System.Drawing.Point(3, 3);
            this.rbMapNone.Name = "rbMapNone";
            this.rbMapNone.Size = new System.Drawing.Size(266, 17);
            this.rbMapNone.TabIndex = 31;
            this.rbMapNone.TabStop = true;
            this.rbMapNone.Tag = "NONE";
            this.rbMapNone.Text = "No Field Mapping (Only First Name and Last Name)";
            this.rbMapNone.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(10, 20);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(171, 13);
            this.label5.TabIndex = 37;
            this.label5.Text = "Default Field Mapping Action";
            // 
            // ultraTabPageControl4
            // 
            this.ultraTabPageControl4.Controls.Add(this.panel4);
            this.ultraTabPageControl4.Controls.Add(this.chkDuplicateDetectionAction);
            this.ultraTabPageControl4.Controls.Add(this.label2);
            this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
            this.ultraTabPageControl4.Name = "ultraTabPageControl4";
            this.ultraTabPageControl4.Size = new System.Drawing.Size(444, 331);
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.rbDuplicationSyncIgnore);
            this.panel4.Controls.Add(this.rbDuplicationSyncModify);
            this.panel4.Location = new System.Drawing.Point(32, 43);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(393, 51);
            this.panel4.TabIndex = 42;
            // 
            // rbDuplicationSyncIgnore
            // 
            this.rbDuplicationSyncIgnore.AutoSize = true;
            this.rbDuplicationSyncIgnore.Location = new System.Drawing.Point(3, 26);
            this.rbDuplicationSyncIgnore.Name = "rbDuplicationSyncIgnore";
            this.rbDuplicationSyncIgnore.Size = new System.Drawing.Size(291, 17);
            this.rbDuplicationSyncIgnore.TabIndex = 32;
            this.rbDuplicationSyncIgnore.TabStop = true;
            this.rbDuplicationSyncIgnore.Tag = "IGNORE";
            this.rbDuplicationSyncIgnore.Text = "Ignore (Abort) updates to Exchange of Duplicate records";
            this.rbDuplicationSyncIgnore.UseVisualStyleBackColor = true;
            // 
            // rbDuplicationSyncModify
            // 
            this.rbDuplicationSyncModify.AutoSize = true;
            this.rbDuplicationSyncModify.Location = new System.Drawing.Point(3, 3);
            this.rbDuplicationSyncModify.Name = "rbDuplicationSyncModify";
            this.rbDuplicationSyncModify.Size = new System.Drawing.Size(286, 17);
            this.rbDuplicationSyncModify.TabIndex = 31;
            this.rbDuplicationSyncModify.TabStop = true;
            this.rbDuplicationSyncModify.Tag = "ALL";
            this.rbDuplicationSyncModify.Text = "Modify duplicate records in Exchange with Minimal data";
            this.rbDuplicationSyncModify.UseVisualStyleBackColor = true;
            // 
            // chkDuplicateDetectionAction
            // 
            this.chkDuplicateDetectionAction.AutoSize = true;
            this.chkDuplicateDetectionAction.Location = new System.Drawing.Point(13, 27);
            this.chkDuplicateDetectionAction.Name = "chkDuplicateDetectionAction";
            this.chkDuplicateDetectionAction.Size = new System.Drawing.Size(211, 17);
            this.chkDuplicateDetectionAction.TabIndex = 41;
            this.chkDuplicateDetectionAction.Tag = "NEW_EXISTING";
            this.chkDuplicateDetectionAction.Text = "Use Existing Duplicate Detection Rules";
            this.chkDuplicateDetectionAction.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(10, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(227, 13);
            this.label2.TabIndex = 40;
            this.label2.Text = "Duplication Detection and Sync Action";
            // 
            // uttmExchangeSettings
            // 
            this.uttmExchangeSettings.ContainingControl = this;
            // 
            // btnSaveSettings
            // 
            this.btnSaveSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSaveSettings.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveSettings.Image = global::ExchangeSync.Properties.Resources.save32;
            this.btnSaveSettings.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSaveSettings.Location = new System.Drawing.Point(247, 373);
            this.btnSaveSettings.Name = "btnSaveSettings";
            this.btnSaveSettings.Size = new System.Drawing.Size(211, 40);
            this.btnSaveSettings.TabIndex = 24;
            this.btnSaveSettings.Text = "Save Settings";
            this.btnSaveSettings.UseVisualStyleBackColor = true;
            this.btnSaveSettings.Click += new System.EventHandler(this.btnSaveSettings_Click);
            // 
            // utGeneral
            // 
            this.utGeneral.Controls.Add(this.ultraTabSharedControlsPage1);
            this.utGeneral.Controls.Add(this.ultraTabPageControl1);
            this.utGeneral.Controls.Add(this.ultraTabPageControl2);
            this.utGeneral.Controls.Add(this.ultraTabPageControl3);
            this.utGeneral.Controls.Add(this.ultraTabPageControl4);
            this.utGeneral.Location = new System.Drawing.Point(13, 15);
            this.utGeneral.Name = "utGeneral";
            this.utGeneral.SharedControlsPage = this.ultraTabSharedControlsPage1;
            this.utGeneral.Size = new System.Drawing.Size(446, 352);
            this.utGeneral.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Flat;
            this.utGeneral.TabIndex = 40;
            ultraTab1.TabPage = this.ultraTabPageControl1;
            ultraTab1.Text = "Update Filter";
            ultraTab2.TabPage = this.ultraTabPageControl2;
            ultraTab2.Text = "List Delete";
            ultraTab3.TabPage = this.ultraTabPageControl3;
            ultraTab3.Text = "Field Mapping";
            ultraTab4.TabPage = this.ultraTabPageControl4;
            ultraTab4.Text = "Duplicate Sync Actions";
            this.utGeneral.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
            ultraTab1,
            ultraTab2,
            ultraTab3,
            ultraTab4});
            this.utGeneral.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Standard;
            // 
            // ultraTabSharedControlsPage1
            // 
            this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
            this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
            this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(444, 331);
            // 
            // lblConfirmation
            // 
            this.lblConfirmation.Location = new System.Drawing.Point(16, 416);
            this.lblConfirmation.Name = "lblConfirmation";
            this.lblConfirmation.Size = new System.Drawing.Size(443, 23);
            this.lblConfirmation.TabIndex = 52;
            // 
            // GeneralSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblConfirmation);
            this.Controls.Add(this.utGeneral);
            this.Controls.Add(this.btnSaveSettings);
            this.Name = "GeneralSettings";
            this.Size = new System.Drawing.Size(480, 450);
            this.Load += new System.EventHandler(this.ExchangeSettings_Load);
            this.ultraTabPageControl1.ResumeLayout(false);
            this.ultraTabPageControl1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ultraTabPageControl2.ResumeLayout(false);
            this.ultraTabPageControl2.PerformLayout();
            this.ultraTabPageControl3.ResumeLayout(false);
            this.ultraTabPageControl3.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ultraTabPageControl4.ResumeLayout(false);
            this.ultraTabPageControl4.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.utGeneral)).EndInit();
            this.utGeneral.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnSaveSettings;
        private Infragistics.Win.UltraWinToolTip.UltraToolTipManager uttmExchangeSettings;
        private Infragistics.Win.UltraWinTabControl.UltraTabControl utGeneral;
        private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
        private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RadioButton rbIgnore;
        private System.Windows.Forms.RadioButton rbShow;
        private System.Windows.Forms.RadioButton rbHide;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton rbUpdateDelta;
        private System.Windows.Forms.RadioButton rbUpdateAll;
        private System.Windows.Forms.Label label3;
        private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
        private System.Windows.Forms.CheckBox chkDeleteOrphanedContacts;
        private System.Windows.Forms.CheckBox chkDeleteExchangeList;
        private System.Windows.Forms.CheckBox chkRemoveConnection;
        private System.Windows.Forms.Label label1;
        private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.RadioButton rbMapCRM;
        private System.Windows.Forms.RadioButton rbMapMinimal;
        private System.Windows.Forms.RadioButton rbMapNone;
        private System.Windows.Forms.Label label5;
        private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.RadioButton rbDuplicationSyncIgnore;
        private System.Windows.Forms.RadioButton rbDuplicationSyncModify;
        private System.Windows.Forms.CheckBox chkDuplicateDetectionAction;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbMailUserFormat;
        private System.Windows.Forms.ComboBox cmbMailContactFormat;
        private System.Windows.Forms.Label lblConfirmation;
    }
}
