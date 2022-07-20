namespace ExchangeSync
{
    partial class FormatSettings
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
            this.panel4 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.chkUpdateMailboxFormat = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lstMailUserFields = new System.Windows.Forms.ListBox();
            this.panel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.lstMailUserFields);
            this.panel4.Controls.Add(this.label1);
            this.panel4.Location = new System.Drawing.Point(30, 42);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(341, 110);
            this.panel4.TabIndex = 42;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(119, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Update Mail User Fields";
            // 
            // chkUpdateMailboxFormat
            // 
            this.chkUpdateMailboxFormat.AutoSize = true;
            this.chkUpdateMailboxFormat.Location = new System.Drawing.Point(11, 26);
            this.chkUpdateMailboxFormat.Name = "chkUpdateMailboxFormat";
            this.chkUpdateMailboxFormat.Size = new System.Drawing.Size(207, 17);
            this.chkUpdateMailboxFormat.TabIndex = 41;
            this.chkUpdateMailboxFormat.Tag = "UPDATEMAILBOX";
            this.chkUpdateMailboxFormat.Text = "Update Exchange Mailbox (Mail User) ";
            this.chkUpdateMailboxFormat.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(245, 13);
            this.label2.TabIndex = 40;
            this.label2.Text = "Mailbox and Mail Contact Update Settings";
            // 
            // lstMailUserFields
            // 
            this.lstMailUserFields.FormattingEnabled = true;
            this.lstMailUserFields.Location = new System.Drawing.Point(130, 8);
            this.lstMailUserFields.Name = "lstMailUserFields";
            this.lstMailUserFields.Size = new System.Drawing.Size(208, 95);
            this.lstMailUserFields.TabIndex = 1;
            // 
            // FormatSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.chkUpdateMailboxFormat);
            this.Controls.Add(this.label2);
            this.Name = "FormatSettings";
            this.Size = new System.Drawing.Size(480, 450);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkUpdateMailboxFormat;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox lstMailUserFields;

    }
}
