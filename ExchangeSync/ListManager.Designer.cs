namespace ExchangeSync
{
    partial class ListManager
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            Infragistics.Win.UltraWinToolbars.UltraToolbar ultraToolbar1 = new Infragistics.Win.UltraWinToolbars.UltraToolbar("utMain");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool5 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolImportNewLists");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool1 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolPopulateExchangeGroupNames");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool3 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolImportExistingLists");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool2 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolPopulateExchangeGroupNames");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool4 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolImportExistingLists");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool6 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolImportNewLists");
            this.dgvDisconnected = new System.Windows.Forms.DataGridView();
            this.ObjectId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ObjectTypeName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ObjectType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.GroupName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ExchangeListName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Active = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.utmList = new Infragistics.Win.UltraWinToolbars.UltraToolbarsManager(this.components);
            this.ListManager_Fill_Panel = new System.Windows.Forms.Panel();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this._ListManager_Toolbars_Dock_Area_Left = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            this._ListManager_Toolbars_Dock_Area_Right = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            this._ListManager_Toolbars_Dock_Area_Top = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            this._ListManager_Toolbars_Dock_Area_Bottom = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDisconnected)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.utmList)).BeginInit();
            this.ListManager_Fill_Panel.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvDisconnected
            // 
            this.dgvDisconnected.AllowUserToAddRows = false;
            this.dgvDisconnected.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvDisconnected.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvDisconnected.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvDisconnected.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDisconnected.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ObjectId,
            this.ObjectTypeName,
            this.ObjectType,
            this.GroupName,
            this.ExchangeListName,
            this.Active});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvDisconnected.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvDisconnected.Dock = System.Windows.Forms.DockStyle.Top;
            this.dgvDisconnected.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnF2;
            this.dgvDisconnected.Location = new System.Drawing.Point(0, 0);
            this.dgvDisconnected.Name = "dgvDisconnected";
            this.dgvDisconnected.Size = new System.Drawing.Size(709, 423);
            this.dgvDisconnected.TabIndex = 11;
            // 
            // ObjectId
            // 
            this.ObjectId.HeaderText = "Object Id";
            this.ObjectId.Name = "ObjectId";
            this.ObjectId.ReadOnly = true;
            this.ObjectId.Visible = false;
            // 
            // ObjectTypeName
            // 
            this.ObjectTypeName.HeaderText = "Object Type Name";
            this.ObjectTypeName.Name = "ObjectTypeName";
            this.ObjectTypeName.ReadOnly = true;
            this.ObjectTypeName.Visible = false;
            // 
            // ObjectType
            // 
            this.ObjectType.HeaderText = "Object Type";
            this.ObjectType.Name = "ObjectType";
            this.ObjectType.ReadOnly = true;
            // 
            // GroupName
            // 
            this.GroupName.HeaderText = "Group Name";
            this.GroupName.Name = "GroupName";
            this.GroupName.ReadOnly = true;
            // 
            // ExchangeListName
            // 
            this.ExchangeListName.HeaderText = "Exchange List Name";
            this.ExchangeListName.Name = "ExchangeListName";
            // 
            // Active
            // 
            this.Active.HeaderText = "Active";
            this.Active.Name = "Active";
            this.Active.ReadOnly = true;
            // 
            // utmList
            // 
            this.utmList.DesignerFlags = 1;
            this.utmList.DockWithinContainer = this;
            this.utmList.DockWithinContainerBaseType = typeof(System.Windows.Forms.Form);
            this.utmList.ShowFullMenusDelay = 500;
            ultraToolbar1.DockedColumn = 0;
            ultraToolbar1.DockedRow = 0;
            buttonTool3.InstanceProps.IsFirstInGroup = true;
            ultraToolbar1.NonInheritedTools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            buttonTool5,
            buttonTool1,
            buttonTool3});
            ultraToolbar1.Text = "utMain";
            this.utmList.Toolbars.AddRange(new Infragistics.Win.UltraWinToolbars.UltraToolbar[] {
            ultraToolbar1});
            buttonTool2.SharedPropsInternal.Caption = "Populate Exchange Group Names";
            buttonTool2.SharedPropsInternal.CustomizerCaption = "Populate Exchange Group Names";
            buttonTool2.SharedPropsInternal.DisplayStyle = Infragistics.Win.UltraWinToolbars.ToolDisplayStyle.TextOnlyAlways;
            buttonTool2.SharedPropsInternal.ToolTipText = "Automatically populate the Exchange Group Names for any lists that do not have a " +
    "Group Name";
            buttonTool2.SharedPropsInternal.ToolTipTitle = "Populate Exchange Group Names";
            buttonTool4.SharedPropsInternal.Caption = "Import Existing Lists";
            buttonTool4.SharedPropsInternal.CustomizerCaption = "Import Existing Lists";
            buttonTool4.SharedPropsInternal.DisplayStyle = Infragistics.Win.UltraWinToolbars.ToolDisplayStyle.TextOnlyAlways;
            buttonTool4.SharedPropsInternal.ToolTipText = "Import Lists from CRM that were synched in the previous versions of Exchange Sync" +
    "";
            buttonTool4.SharedPropsInternal.ToolTipTitle = "Import Existing Lists";
            buttonTool6.SharedPropsInternal.Caption = "Import New Lists";
            buttonTool6.SharedPropsInternal.CustomizerCaption = "Import New Lists";
            buttonTool6.SharedPropsInternal.DisplayStyle = Infragistics.Win.UltraWinToolbars.ToolDisplayStyle.TextOnlyAlways;
            buttonTool6.SharedPropsInternal.ToolTipText = "Import Lists from CRM that were not synched in the previous versions of Exchange " +
    "Sync";
            buttonTool6.SharedPropsInternal.ToolTipTitle = "Import New Lists";
            this.utmList.Tools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            buttonTool2,
            buttonTool4,
            buttonTool6});
            this.utmList.ToolClick += new Infragistics.Win.UltraWinToolbars.ToolClickEventHandler(this.utmList_ToolClick);
            // 
            // ListManager_Fill_Panel
            // 
            this.ListManager_Fill_Panel.Controls.Add(this.btnCancel);
            this.ListManager_Fill_Panel.Controls.Add(this.btnOK);
            this.ListManager_Fill_Panel.Controls.Add(this.dgvDisconnected);
            this.ListManager_Fill_Panel.Cursor = System.Windows.Forms.Cursors.Default;
            this.ListManager_Fill_Panel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ListManager_Fill_Panel.Location = new System.Drawing.Point(0, 23);
            this.ListManager_Fill_Panel.Name = "ListManager_Fill_Panel";
            this.ListManager_Fill_Panel.Size = new System.Drawing.Size(709, 472);
            this.ListManager_Fill_Panel.TabIndex = 0;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(622, 435);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 13;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(535, 435);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 12;
            this.btnOK.Text = "Accept";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // _ListManager_Toolbars_Dock_Area_Left
            // 
            this._ListManager_Toolbars_Dock_Area_Left.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._ListManager_Toolbars_Dock_Area_Left.BackColor = System.Drawing.SystemColors.Control;
            this._ListManager_Toolbars_Dock_Area_Left.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Left;
            this._ListManager_Toolbars_Dock_Area_Left.ForeColor = System.Drawing.SystemColors.ControlText;
            this._ListManager_Toolbars_Dock_Area_Left.Location = new System.Drawing.Point(0, 23);
            this._ListManager_Toolbars_Dock_Area_Left.Name = "_ListManager_Toolbars_Dock_Area_Left";
            this._ListManager_Toolbars_Dock_Area_Left.Size = new System.Drawing.Size(0, 472);
            this._ListManager_Toolbars_Dock_Area_Left.ToolbarsManager = this.utmList;
            // 
            // _ListManager_Toolbars_Dock_Area_Right
            // 
            this._ListManager_Toolbars_Dock_Area_Right.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._ListManager_Toolbars_Dock_Area_Right.BackColor = System.Drawing.SystemColors.Control;
            this._ListManager_Toolbars_Dock_Area_Right.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Right;
            this._ListManager_Toolbars_Dock_Area_Right.ForeColor = System.Drawing.SystemColors.ControlText;
            this._ListManager_Toolbars_Dock_Area_Right.Location = new System.Drawing.Point(709, 23);
            this._ListManager_Toolbars_Dock_Area_Right.Name = "_ListManager_Toolbars_Dock_Area_Right";
            this._ListManager_Toolbars_Dock_Area_Right.Size = new System.Drawing.Size(0, 472);
            this._ListManager_Toolbars_Dock_Area_Right.ToolbarsManager = this.utmList;
            // 
            // _ListManager_Toolbars_Dock_Area_Top
            // 
            this._ListManager_Toolbars_Dock_Area_Top.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._ListManager_Toolbars_Dock_Area_Top.BackColor = System.Drawing.SystemColors.Control;
            this._ListManager_Toolbars_Dock_Area_Top.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Top;
            this._ListManager_Toolbars_Dock_Area_Top.ForeColor = System.Drawing.SystemColors.ControlText;
            this._ListManager_Toolbars_Dock_Area_Top.Location = new System.Drawing.Point(0, 0);
            this._ListManager_Toolbars_Dock_Area_Top.Name = "_ListManager_Toolbars_Dock_Area_Top";
            this._ListManager_Toolbars_Dock_Area_Top.Size = new System.Drawing.Size(709, 23);
            this._ListManager_Toolbars_Dock_Area_Top.ToolbarsManager = this.utmList;
            // 
            // _ListManager_Toolbars_Dock_Area_Bottom
            // 
            this._ListManager_Toolbars_Dock_Area_Bottom.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._ListManager_Toolbars_Dock_Area_Bottom.BackColor = System.Drawing.SystemColors.Control;
            this._ListManager_Toolbars_Dock_Area_Bottom.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Bottom;
            this._ListManager_Toolbars_Dock_Area_Bottom.ForeColor = System.Drawing.SystemColors.ControlText;
            this._ListManager_Toolbars_Dock_Area_Bottom.Location = new System.Drawing.Point(0, 495);
            this._ListManager_Toolbars_Dock_Area_Bottom.Name = "_ListManager_Toolbars_Dock_Area_Bottom";
            this._ListManager_Toolbars_Dock_Area_Bottom.Size = new System.Drawing.Size(709, 0);
            this._ListManager_Toolbars_Dock_Area_Bottom.ToolbarsManager = this.utmList;
            // 
            // ListManager
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(709, 495);
            this.ControlBox = false;
            this.Controls.Add(this.ListManager_Fill_Panel);
            this.Controls.Add(this._ListManager_Toolbars_Dock_Area_Left);
            this.Controls.Add(this._ListManager_Toolbars_Dock_Area_Right);
            this.Controls.Add(this._ListManager_Toolbars_Dock_Area_Bottom);
            this.Controls.Add(this._ListManager_Toolbars_Dock_Area_Top);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "ListManager";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "ListManager";
            this.Load += new System.EventHandler(this.ListManager_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDisconnected)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.utmList)).EndInit();
            this.ListManager_Fill_Panel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvDisconnected;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsManager utmList;
        private System.Windows.Forms.Panel ListManager_Fill_Panel;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _ListManager_Toolbars_Dock_Area_Left;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _ListManager_Toolbars_Dock_Area_Right;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _ListManager_Toolbars_Dock_Area_Bottom;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _ListManager_Toolbars_Dock_Area_Top;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.DataGridViewTextBoxColumn ObjectId;
        private System.Windows.Forms.DataGridViewTextBoxColumn ObjectTypeName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ObjectType;
        private System.Windows.Forms.DataGridViewTextBoxColumn GroupName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ExchangeListName;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Active;
    }
}