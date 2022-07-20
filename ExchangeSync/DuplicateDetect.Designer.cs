namespace ExchangeSync
{
    partial class DuplicateDetect
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            Infragistics.Win.UltraWinToolbars.UltraToolbar ultraToolbar1 = new Infragistics.Win.UltraWinToolbars.UltraToolbar("utDuplicates");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool1 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolDetectDuplicates");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool3 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolUpdateExchangeAliases");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool4 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolClearExchangeAliases");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool7 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolChangeStatus");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool2 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolDetectDuplicates");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool5 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolUpdateExchangeAliases");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool6 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolClearExchangeAliases");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool8 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolChangeStatus");
            this.dgvDuplicates = new System.Windows.Forms.DataGridView();
            this.entityId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ContactNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fullname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.emailaddress1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.xrm_exchangealias = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chkSelected = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Revision = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ultraToolbarsManager1 = new Infragistics.Win.UltraWinToolbars.UltraToolbarsManager(this.components);
            this.DuplicateDetect_Fill_Panel = new System.Windows.Forms.Panel();
            this._DuplicateDetect_Toolbars_Dock_Area_Left = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            this._DuplicateDetect_Toolbars_Dock_Area_Right = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            this._DuplicateDetect_Toolbars_Dock_Area_Top = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            this._DuplicateDetect_Toolbars_Dock_Area_Bottom = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDuplicates)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ultraToolbarsManager1)).BeginInit();
            this.DuplicateDetect_Fill_Panel.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvDuplicates
            // 
            this.dgvDuplicates.AllowUserToAddRows = false;
            this.dgvDuplicates.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvDuplicates.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvDuplicates.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvDuplicates.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDuplicates.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.entityId,
            this.ContactNo,
            this.fullname,
            this.emailaddress1,
            this.xrm_exchangealias,
            this.chkSelected,
            this.Revision});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvDuplicates.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvDuplicates.Dock = System.Windows.Forms.DockStyle.Left;
            this.dgvDuplicates.Location = new System.Drawing.Point(0, 0);
            this.dgvDuplicates.MultiSelect = false;
            this.dgvDuplicates.Name = "dgvDuplicates";
            this.dgvDuplicates.ReadOnly = true;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvDuplicates.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dgvDuplicates.Size = new System.Drawing.Size(914, 471);
            this.dgvDuplicates.TabIndex = 13;
            this.dgvDuplicates.Visible = false;
            // 
            // entityId
            // 
            this.entityId.HeaderText = "Object Id";
            this.entityId.Name = "entityId";
            this.entityId.ReadOnly = true;
            this.entityId.Visible = false;
            // 
            // ContactNo
            // 
            this.ContactNo.HeaderText = "Contact Id";
            this.ContactNo.Name = "ContactNo";
            this.ContactNo.ReadOnly = true;
            this.ContactNo.Visible = false;
            // 
            // fullname
            // 
            this.fullname.HeaderText = "Full Name";
            this.fullname.Name = "fullname";
            this.fullname.ReadOnly = true;
            // 
            // emailaddress1
            // 
            this.emailaddress1.HeaderText = "Email Address";
            this.emailaddress1.Name = "emailaddress1";
            this.emailaddress1.ReadOnly = true;
            // 
            // xrm_exchangealias
            // 
            this.xrm_exchangealias.HeaderText = "Exchange Alias";
            this.xrm_exchangealias.Name = "xrm_exchangealias";
            this.xrm_exchangealias.ReadOnly = true;
            // 
            // chkSelected
            // 
            this.chkSelected.HeaderText = "Select";
            this.chkSelected.Name = "chkSelected";
            this.chkSelected.ReadOnly = true;
            // 
            // Revision
            // 
            this.Revision.HeaderText = "Revision";
            this.Revision.Name = "Revision";
            this.Revision.ReadOnly = true;
            this.Revision.Visible = false;
            // 
            // ultraToolbarsManager1
            // 
            this.ultraToolbarsManager1.DesignerFlags = 1;
            this.ultraToolbarsManager1.DockWithinContainer = this;
            this.ultraToolbarsManager1.DockWithinContainerBaseType = typeof(System.Windows.Forms.Form);
            ultraToolbar1.DockedColumn = 0;
            ultraToolbar1.DockedRow = 0;
            buttonTool3.InstanceProps.IsFirstInGroup = true;
            buttonTool7.InstanceProps.IsFirstInGroup = true;
            ultraToolbar1.NonInheritedTools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            buttonTool1,
            buttonTool3,
            buttonTool4,
            buttonTool7});
            ultraToolbar1.Text = "utDuplicates";
            this.ultraToolbarsManager1.Toolbars.AddRange(new Infragistics.Win.UltraWinToolbars.UltraToolbar[] {
            ultraToolbar1});
            buttonTool2.SharedPropsInternal.Caption = "Detect Duplicates";
            buttonTool2.SharedPropsInternal.CustomizerCaption = "Detect Duplicates";
            buttonTool2.SharedPropsInternal.DisplayStyle = Infragistics.Win.UltraWinToolbars.ToolDisplayStyle.ImageAndText;
            buttonTool5.SharedPropsInternal.Caption = "Update Exchange Aliases";
            buttonTool5.SharedPropsInternal.CustomizerCaption = "Update Exchange Aliases";
            buttonTool5.SharedPropsInternal.DisplayStyle = Infragistics.Win.UltraWinToolbars.ToolDisplayStyle.ImageAndText;
            buttonTool6.SharedPropsInternal.Caption = "Clear Exchange Aliases";
            buttonTool6.SharedPropsInternal.CustomizerCaption = "Clear Exchange Aliases";
            buttonTool6.SharedPropsInternal.DisplayStyle = Infragistics.Win.UltraWinToolbars.ToolDisplayStyle.ImageAndText;
            buttonTool8.SharedPropsInternal.Caption = "Change Status";
            buttonTool8.SharedPropsInternal.CustomizerCaption = "Change Status";
            buttonTool8.SharedPropsInternal.DisplayStyle = Infragistics.Win.UltraWinToolbars.ToolDisplayStyle.ImageAndText;
            this.ultraToolbarsManager1.Tools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            buttonTool2,
            buttonTool5,
            buttonTool6,
            buttonTool8});
            // 
            // DuplicateDetect_Fill_Panel
            // 
            this.DuplicateDetect_Fill_Panel.Controls.Add(this.dgvDuplicates);
            this.DuplicateDetect_Fill_Panel.Cursor = System.Windows.Forms.Cursors.Default;
            this.DuplicateDetect_Fill_Panel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DuplicateDetect_Fill_Panel.Location = new System.Drawing.Point(0, 25);
            this.DuplicateDetect_Fill_Panel.Name = "DuplicateDetect_Fill_Panel";
            this.DuplicateDetect_Fill_Panel.Size = new System.Drawing.Size(916, 471);
            this.DuplicateDetect_Fill_Panel.TabIndex = 0;
            // 
            // _DuplicateDetect_Toolbars_Dock_Area_Left
            // 
            this._DuplicateDetect_Toolbars_Dock_Area_Left.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._DuplicateDetect_Toolbars_Dock_Area_Left.BackColor = System.Drawing.SystemColors.Control;
            this._DuplicateDetect_Toolbars_Dock_Area_Left.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Left;
            this._DuplicateDetect_Toolbars_Dock_Area_Left.ForeColor = System.Drawing.SystemColors.ControlText;
            this._DuplicateDetect_Toolbars_Dock_Area_Left.Location = new System.Drawing.Point(0, 25);
            this._DuplicateDetect_Toolbars_Dock_Area_Left.Name = "_DuplicateDetect_Toolbars_Dock_Area_Left";
            this._DuplicateDetect_Toolbars_Dock_Area_Left.Size = new System.Drawing.Size(0, 471);
            this._DuplicateDetect_Toolbars_Dock_Area_Left.ToolbarsManager = this.ultraToolbarsManager1;
            // 
            // _DuplicateDetect_Toolbars_Dock_Area_Right
            // 
            this._DuplicateDetect_Toolbars_Dock_Area_Right.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._DuplicateDetect_Toolbars_Dock_Area_Right.BackColor = System.Drawing.SystemColors.Control;
            this._DuplicateDetect_Toolbars_Dock_Area_Right.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Right;
            this._DuplicateDetect_Toolbars_Dock_Area_Right.ForeColor = System.Drawing.SystemColors.ControlText;
            this._DuplicateDetect_Toolbars_Dock_Area_Right.Location = new System.Drawing.Point(916, 25);
            this._DuplicateDetect_Toolbars_Dock_Area_Right.Name = "_DuplicateDetect_Toolbars_Dock_Area_Right";
            this._DuplicateDetect_Toolbars_Dock_Area_Right.Size = new System.Drawing.Size(0, 471);
            this._DuplicateDetect_Toolbars_Dock_Area_Right.ToolbarsManager = this.ultraToolbarsManager1;
            // 
            // _DuplicateDetect_Toolbars_Dock_Area_Top
            // 
            this._DuplicateDetect_Toolbars_Dock_Area_Top.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._DuplicateDetect_Toolbars_Dock_Area_Top.BackColor = System.Drawing.SystemColors.Control;
            this._DuplicateDetect_Toolbars_Dock_Area_Top.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Top;
            this._DuplicateDetect_Toolbars_Dock_Area_Top.ForeColor = System.Drawing.SystemColors.ControlText;
            this._DuplicateDetect_Toolbars_Dock_Area_Top.Location = new System.Drawing.Point(0, 0);
            this._DuplicateDetect_Toolbars_Dock_Area_Top.Name = "_DuplicateDetect_Toolbars_Dock_Area_Top";
            this._DuplicateDetect_Toolbars_Dock_Area_Top.Size = new System.Drawing.Size(916, 25);
            this._DuplicateDetect_Toolbars_Dock_Area_Top.ToolbarsManager = this.ultraToolbarsManager1;
            // 
            // _DuplicateDetect_Toolbars_Dock_Area_Bottom
            // 
            this._DuplicateDetect_Toolbars_Dock_Area_Bottom.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._DuplicateDetect_Toolbars_Dock_Area_Bottom.BackColor = System.Drawing.SystemColors.Control;
            this._DuplicateDetect_Toolbars_Dock_Area_Bottom.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Bottom;
            this._DuplicateDetect_Toolbars_Dock_Area_Bottom.ForeColor = System.Drawing.SystemColors.ControlText;
            this._DuplicateDetect_Toolbars_Dock_Area_Bottom.Location = new System.Drawing.Point(0, 496);
            this._DuplicateDetect_Toolbars_Dock_Area_Bottom.Name = "_DuplicateDetect_Toolbars_Dock_Area_Bottom";
            this._DuplicateDetect_Toolbars_Dock_Area_Bottom.Size = new System.Drawing.Size(916, 0);
            this._DuplicateDetect_Toolbars_Dock_Area_Bottom.ToolbarsManager = this.ultraToolbarsManager1;
            // 
            // DuplicateDetect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(916, 496);
            this.Controls.Add(this.DuplicateDetect_Fill_Panel);
            this.Controls.Add(this._DuplicateDetect_Toolbars_Dock_Area_Left);
            this.Controls.Add(this._DuplicateDetect_Toolbars_Dock_Area_Right);
            this.Controls.Add(this._DuplicateDetect_Toolbars_Dock_Area_Bottom);
            this.Controls.Add(this._DuplicateDetect_Toolbars_Dock_Area_Top);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "DuplicateDetect";
            this.Text = "Duplicate Detection";
            ((System.ComponentModel.ISupportInitialize)(this.dgvDuplicates)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ultraToolbarsManager1)).EndInit();
            this.DuplicateDetect_Fill_Panel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvDuplicates;
        private System.Windows.Forms.DataGridViewTextBoxColumn entityId;
        private System.Windows.Forms.DataGridViewTextBoxColumn ContactNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn fullname;
        private System.Windows.Forms.DataGridViewTextBoxColumn emailaddress1;
        private System.Windows.Forms.DataGridViewTextBoxColumn xrm_exchangealias;
        private System.Windows.Forms.DataGridViewCheckBoxColumn chkSelected;
        private System.Windows.Forms.DataGridViewTextBoxColumn Revision;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsManager ultraToolbarsManager1;
        private System.Windows.Forms.Panel DuplicateDetect_Fill_Panel;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _DuplicateDetect_Toolbars_Dock_Area_Left;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _DuplicateDetect_Toolbars_Dock_Area_Right;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _DuplicateDetect_Toolbars_Dock_Area_Bottom;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _DuplicateDetect_Toolbars_Dock_Area_Top;
    }
}