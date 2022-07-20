namespace ExchangeSync
{
    partial class History
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
            this.label3 = new System.Windows.Forms.Label();
            this.btnExchangeSync = new System.Windows.Forms.Button();
            this.btnExchangeAutoSync = new System.Windows.Forms.Button();
            this.btnExchangeConsoleSync = new System.Windows.Forms.Button();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnRecent = new System.Windows.Forms.Button();
            this.ulvFiles = new Infragistics.Win.UltraWinListView.UltraListView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.ulvFiles)).BeginInit();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI Light", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.SystemColors.GrayText;
            this.label3.Location = new System.Drawing.Point(6, 16);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(125, 47);
            this.label3.TabIndex = 65;
            this.label3.Text = "History";
            // 
            // btnExchangeSync
            // 
            this.btnExchangeSync.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.btnExchangeSync.FlatAppearance.BorderSize = 0;
            this.btnExchangeSync.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExchangeSync.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExchangeSync.ForeColor = System.Drawing.Color.DimGray;
            this.btnExchangeSync.Image = global::ExchangeSync.Properties.Resources.exchangesync32;
            this.btnExchangeSync.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExchangeSync.Location = new System.Drawing.Point(14, 157);
            this.btnExchangeSync.Name = "btnExchangeSync";
            this.btnExchangeSync.Size = new System.Drawing.Size(260, 46);
            this.btnExchangeSync.TabIndex = 70;
            this.btnExchangeSync.Text = "Exchange Sync";
            this.btnExchangeSync.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExchangeSync.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnExchangeSync.UseVisualStyleBackColor = true;
            this.btnExchangeSync.Click += new System.EventHandler(this.btnExchangeSync_Click);
            this.btnExchangeSync.MouseEnter += new System.EventHandler(this.Button_MouseHover);
            this.btnExchangeSync.MouseLeave += new System.EventHandler(this.Button_MouseLeave);
            this.btnExchangeSync.MouseHover += new System.EventHandler(this.Button_MouseHover);
            // 
            // btnExchangeAutoSync
            // 
            this.btnExchangeAutoSync.FlatAppearance.BorderSize = 0;
            this.btnExchangeAutoSync.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExchangeAutoSync.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExchangeAutoSync.ForeColor = System.Drawing.Color.DimGray;
            this.btnExchangeAutoSync.Image = global::ExchangeSync.Properties.Resources.exchangeautosync32;
            this.btnExchangeAutoSync.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExchangeAutoSync.Location = new System.Drawing.Point(14, 222);
            this.btnExchangeAutoSync.Name = "btnExchangeAutoSync";
            this.btnExchangeAutoSync.Size = new System.Drawing.Size(260, 46);
            this.btnExchangeAutoSync.TabIndex = 71;
            this.btnExchangeAutoSync.Text = "Exchange AutoSync";
            this.btnExchangeAutoSync.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExchangeAutoSync.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnExchangeAutoSync.UseVisualStyleBackColor = true;
            this.btnExchangeAutoSync.MouseEnter += new System.EventHandler(this.Button_MouseHover);
            this.btnExchangeAutoSync.MouseLeave += new System.EventHandler(this.Button_MouseLeave);
            this.btnExchangeAutoSync.MouseHover += new System.EventHandler(this.Button_MouseHover);
            // 
            // btnExchangeConsoleSync
            // 
            this.btnExchangeConsoleSync.FlatAppearance.BorderSize = 0;
            this.btnExchangeConsoleSync.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExchangeConsoleSync.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExchangeConsoleSync.ForeColor = System.Drawing.Color.DimGray;
            this.btnExchangeConsoleSync.Image = global::ExchangeSync.Properties.Resources.exchangeconsolesync32;
            this.btnExchangeConsoleSync.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExchangeConsoleSync.Location = new System.Drawing.Point(14, 289);
            this.btnExchangeConsoleSync.Name = "btnExchangeConsoleSync";
            this.btnExchangeConsoleSync.Size = new System.Drawing.Size(260, 46);
            this.btnExchangeConsoleSync.TabIndex = 72;
            this.btnExchangeConsoleSync.Text = "Exchange AutoSync";
            this.btnExchangeConsoleSync.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExchangeConsoleSync.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnExchangeConsoleSync.UseVisualStyleBackColor = true;
            this.btnExchangeConsoleSync.MouseEnter += new System.EventHandler(this.Button_MouseHover);
            this.btnExchangeConsoleSync.MouseLeave += new System.EventHandler(this.Button_MouseLeave);
            this.btnExchangeConsoleSync.MouseHover += new System.EventHandler(this.Button_MouseHover);
            // 
            // btnBrowse
            // 
            this.btnBrowse.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.btnBrowse.FlatAppearance.BorderSize = 0;
            this.btnBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowse.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowse.ForeColor = System.Drawing.Color.DimGray;
            this.btnBrowse.Image = global::ExchangeSync.Properties.Resources.exportlog32;
            this.btnBrowse.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnBrowse.Location = new System.Drawing.Point(14, 402);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(260, 46);
            this.btnBrowse.TabIndex = 73;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnBrowse.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.MouseEnter += new System.EventHandler(this.Button_MouseHover);
            this.btnBrowse.MouseLeave += new System.EventHandler(this.Button_MouseLeave);
            this.btnBrowse.MouseHover += new System.EventHandler(this.Button_MouseHover);
            // 
            // btnRecent
            // 
            this.btnRecent.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.btnRecent.FlatAppearance.BorderSize = 0;
            this.btnRecent.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRecent.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRecent.ForeColor = System.Drawing.Color.DimGray;
            this.btnRecent.Image = global::ExchangeSync.Properties.Resources.extendtrial32;
            this.btnRecent.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRecent.Location = new System.Drawing.Point(14, 87);
            this.btnRecent.Name = "btnRecent";
            this.btnRecent.Size = new System.Drawing.Size(260, 46);
            this.btnRecent.TabIndex = 74;
            this.btnRecent.Text = "Recent Logs";
            this.btnRecent.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRecent.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnRecent.UseVisualStyleBackColor = true;
            this.btnRecent.Click += new System.EventHandler(this.btnRecent_Click);
            this.btnRecent.MouseEnter += new System.EventHandler(this.Button_MouseHover);
            this.btnRecent.MouseLeave += new System.EventHandler(this.Button_MouseLeave);
            this.btnRecent.MouseHover += new System.EventHandler(this.Button_MouseHover);
            // 
            // ulvFiles
            // 
            this.ulvFiles.ItemSettings.AllowEdit = Infragistics.Win.DefaultableBoolean.False;
            this.ulvFiles.Location = new System.Drawing.Point(345, 87);
            this.ulvFiles.Name = "ulvFiles";
            this.ulvFiles.Size = new System.Drawing.Size(498, 431);
            this.ulvFiles.TabIndex = 75;
            this.ulvFiles.View = Infragistics.Win.UltraWinListView.UltraListViewStyle.List;
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(32, 32);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // History
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.ulvFiles);
            this.Controls.Add(this.btnRecent);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.btnExchangeConsoleSync);
            this.Controls.Add(this.btnExchangeAutoSync);
            this.Controls.Add(this.btnExchangeSync);
            this.Controls.Add(this.label3);
            this.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Name = "History";
            this.Size = new System.Drawing.Size(862, 581);
            ((System.ComponentModel.ISupportInitialize)(this.ulvFiles)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnExchangeSync;
        private System.Windows.Forms.Button btnExchangeAutoSync;
        private System.Windows.Forms.Button btnExchangeConsoleSync;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnRecent;
        private Infragistics.Win.UltraWinListView.UltraListView ulvFiles;
        private System.Windows.Forms.ImageList imageList1;
    }
}
