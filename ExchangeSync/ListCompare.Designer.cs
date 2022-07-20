namespace ExchangeSync
{
    partial class ListCompare
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
            Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
            Infragistics.Win.UltraWinToolbars.UltraToolbar ultraToolbar1 = new Infragistics.Win.UltraWinToolbars.UltraToolbar("ListCompareToolbar");
            Infragistics.Win.UltraWinToolbars.PopupMenuTool popupMenuTool1 = new Infragistics.Win.UltraWinToolbars.PopupMenuTool("toolSetDisplayStatus");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool1 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolRefresh");
            Infragistics.Win.UltraWinToolbars.PopupMenuTool popupMenuTool2 = new Infragistics.Win.UltraWinToolbars.PopupMenuTool("toolSetDisplayStatus");
            Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool1 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("toolShowAll", "");
            Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool3 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("toolShowActive", "");
            Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool5 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("toolShowInactive", "");
            Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool2 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("toolShowAll", "");
            Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool4 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("toolShowActive", "");
            Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool6 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("toolShowInactive", "");
            Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool2 = new Infragistics.Win.UltraWinToolbars.ButtonTool("toolRefresh");
            this.ugCompareList = new Infragistics.Win.UltraWinGrid.UltraGrid();
            this.ultraPanel1 = new Infragistics.Win.Misc.UltraPanel();
            this.btnResetCrmAlias = new System.Windows.Forms.Button();
            this.btnResetExchangeAlias = new System.Windows.Forms.Button();
            this.utmListCompare = new Infragistics.Win.UltraWinToolbars.UltraToolbarsManager(this.components);
            this.ListCompare_Fill_Panel = new System.Windows.Forms.Panel();
            this._ListCompare_Toolbars_Dock_Area_Left = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            this._ListCompare_Toolbars_Dock_Area_Right = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            this._ListCompare_Toolbars_Dock_Area_Top = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            this._ListCompare_Toolbars_Dock_Area_Bottom = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
            ((System.ComponentModel.ISupportInitialize)(this.ugCompareList)).BeginInit();
            this.ultraPanel1.ClientArea.SuspendLayout();
            this.ultraPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.utmListCompare)).BeginInit();
            this.ListCompare_Fill_Panel.SuspendLayout();
            this.SuspendLayout();
            // 
            // ugCompareList
            // 
            appearance1.BackColor = System.Drawing.Color.White;
            this.ugCompareList.DisplayLayout.Appearance = appearance1;
            this.ugCompareList.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
            this.ugCompareList.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
            this.ugCompareList.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.True;
            this.ugCompareList.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
            this.ugCompareList.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.None;
            appearance2.BackColor = System.Drawing.Color.Transparent;
            this.ugCompareList.DisplayLayout.Override.CardAreaAppearance = appearance2;
            this.ugCompareList.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
            this.ugCompareList.DisplayLayout.Override.CellPadding = 3;
            appearance3.TextHAlignAsString = "Left";
            this.ugCompareList.DisplayLayout.Override.HeaderAppearance = appearance3;
            this.ugCompareList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
            appearance4.BorderColor = System.Drawing.Color.LightGray;
            appearance4.TextVAlignAsString = "Middle";
            this.ugCompareList.DisplayLayout.Override.RowAppearance = appearance4;
            this.ugCompareList.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False;
            appearance5.BackColor = System.Drawing.Color.LightSteelBlue;
            appearance5.BorderColor = System.Drawing.Color.Black;
            appearance5.ForeColor = System.Drawing.Color.Black;
            this.ugCompareList.DisplayLayout.Override.SelectedRowAppearance = appearance5;
            this.ugCompareList.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
            this.ugCompareList.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
            this.ugCompareList.DisplayLayout.Override.SelectTypeGroupByRow = Infragistics.Win.UltraWinGrid.SelectType.None;
            this.ugCompareList.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
            this.ugCompareList.DisplayLayout.RowConnectorStyle = Infragistics.Win.UltraWinGrid.RowConnectorStyle.None;
            this.ugCompareList.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
            this.ugCompareList.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControlOnLastCell;
            this.ugCompareList.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
            this.ugCompareList.Dock = System.Windows.Forms.DockStyle.Top;
            this.ugCompareList.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ugCompareList.Location = new System.Drawing.Point(0, 0);
            this.ugCompareList.Name = "ugCompareList";
            this.ugCompareList.Size = new System.Drawing.Size(752, 507);
            this.ugCompareList.TabIndex = 17;
            this.ugCompareList.Text = "List Comparison";
            this.ugCompareList.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugCompareList_InitializeLayout);
            this.ugCompareList.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ugCompareList_InitializeRow);
            // 
            // ultraPanel1
            // 
            // 
            // ultraPanel1.ClientArea
            // 
            this.ultraPanel1.ClientArea.Controls.Add(this.btnResetCrmAlias);
            this.ultraPanel1.ClientArea.Controls.Add(this.btnResetExchangeAlias);
            this.ultraPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.ultraPanel1.Location = new System.Drawing.Point(0, 511);
            this.ultraPanel1.Name = "ultraPanel1";
            this.ultraPanel1.Size = new System.Drawing.Size(752, 41);
            this.ultraPanel1.TabIndex = 18;
            // 
            // btnResetCrmAlias
            // 
            this.btnResetCrmAlias.Location = new System.Drawing.Point(192, 11);
            this.btnResetCrmAlias.Name = "btnResetCrmAlias";
            this.btnResetCrmAlias.Size = new System.Drawing.Size(154, 23);
            this.btnResetCrmAlias.TabIndex = 6;
            this.btnResetCrmAlias.Text = "Reset Crm Alias";
            this.btnResetCrmAlias.UseVisualStyleBackColor = true;
            this.btnResetCrmAlias.Click += new System.EventHandler(this.btnResetCrmAlias_Click);
            // 
            // btnResetExchangeAlias
            // 
            this.btnResetExchangeAlias.Location = new System.Drawing.Point(12, 11);
            this.btnResetExchangeAlias.Name = "btnResetExchangeAlias";
            this.btnResetExchangeAlias.Size = new System.Drawing.Size(154, 23);
            this.btnResetExchangeAlias.TabIndex = 5;
            this.btnResetExchangeAlias.Text = "Reset Exchange Alias";
            this.btnResetExchangeAlias.UseVisualStyleBackColor = true;
            this.btnResetExchangeAlias.Click += new System.EventHandler(this.btnResetExchangeAlias_Click);
            // 
            // utmListCompare
            // 
            this.utmListCompare.DesignerFlags = 1;
            this.utmListCompare.DockWithinContainer = this;
            this.utmListCompare.DockWithinContainerBaseType = typeof(System.Windows.Forms.Form);
            ultraToolbar1.DockedColumn = 0;
            ultraToolbar1.DockedRow = 0;
            ultraToolbar1.NonInheritedTools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            popupMenuTool1,
            buttonTool1});
            ultraToolbar1.Text = "ListCompareToolbar";
            this.utmListCompare.Toolbars.AddRange(new Infragistics.Win.UltraWinToolbars.UltraToolbar[] {
            ultraToolbar1});
            popupMenuTool2.SharedPropsInternal.Caption = "Set Display Status";
            popupMenuTool2.SharedPropsInternal.CustomizerCaption = "Set Display Status";
            popupMenuTool2.Tools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            stateButtonTool1,
            stateButtonTool3,
            stateButtonTool5});
            stateButtonTool2.SharedPropsInternal.Caption = "Show All";
            stateButtonTool2.SharedPropsInternal.CustomizerCaption = "Show All";
            stateButtonTool4.SharedPropsInternal.Caption = "Show Active";
            stateButtonTool4.SharedPropsInternal.CustomizerCaption = "Show Active";
            stateButtonTool6.SharedPropsInternal.Caption = "Show Inactive";
            stateButtonTool6.SharedPropsInternal.CustomizerCaption = "Show Inactive";
            buttonTool2.SharedPropsInternal.Caption = "Refresh";
            buttonTool2.SharedPropsInternal.CustomizerCaption = "Refresh";
            buttonTool2.SharedPropsInternal.DisplayStyle = Infragistics.Win.UltraWinToolbars.ToolDisplayStyle.TextOnlyAlways;
            this.utmListCompare.Tools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            popupMenuTool2,
            stateButtonTool2,
            stateButtonTool4,
            stateButtonTool6,
            buttonTool2});
            this.utmListCompare.ToolClick += new Infragistics.Win.UltraWinToolbars.ToolClickEventHandler(this.utmListCompare_ToolClick);
            // 
            // ListCompare_Fill_Panel
            // 
            this.ListCompare_Fill_Panel.Controls.Add(this.ultraPanel1);
            this.ListCompare_Fill_Panel.Controls.Add(this.ugCompareList);
            this.ListCompare_Fill_Panel.Cursor = System.Windows.Forms.Cursors.Default;
            this.ListCompare_Fill_Panel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ListCompare_Fill_Panel.Location = new System.Drawing.Point(0, 25);
            this.ListCompare_Fill_Panel.Name = "ListCompare_Fill_Panel";
            this.ListCompare_Fill_Panel.Size = new System.Drawing.Size(752, 552);
            this.ListCompare_Fill_Panel.TabIndex = 0;
            // 
            // _ListCompare_Toolbars_Dock_Area_Left
            // 
            this._ListCompare_Toolbars_Dock_Area_Left.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._ListCompare_Toolbars_Dock_Area_Left.BackColor = System.Drawing.SystemColors.Control;
            this._ListCompare_Toolbars_Dock_Area_Left.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Left;
            this._ListCompare_Toolbars_Dock_Area_Left.ForeColor = System.Drawing.SystemColors.ControlText;
            this._ListCompare_Toolbars_Dock_Area_Left.Location = new System.Drawing.Point(0, 25);
            this._ListCompare_Toolbars_Dock_Area_Left.Name = "_ListCompare_Toolbars_Dock_Area_Left";
            this._ListCompare_Toolbars_Dock_Area_Left.Size = new System.Drawing.Size(0, 552);
            this._ListCompare_Toolbars_Dock_Area_Left.ToolbarsManager = this.utmListCompare;
            // 
            // _ListCompare_Toolbars_Dock_Area_Right
            // 
            this._ListCompare_Toolbars_Dock_Area_Right.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._ListCompare_Toolbars_Dock_Area_Right.BackColor = System.Drawing.SystemColors.Control;
            this._ListCompare_Toolbars_Dock_Area_Right.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Right;
            this._ListCompare_Toolbars_Dock_Area_Right.ForeColor = System.Drawing.SystemColors.ControlText;
            this._ListCompare_Toolbars_Dock_Area_Right.Location = new System.Drawing.Point(752, 25);
            this._ListCompare_Toolbars_Dock_Area_Right.Name = "_ListCompare_Toolbars_Dock_Area_Right";
            this._ListCompare_Toolbars_Dock_Area_Right.Size = new System.Drawing.Size(0, 552);
            this._ListCompare_Toolbars_Dock_Area_Right.ToolbarsManager = this.utmListCompare;
            // 
            // _ListCompare_Toolbars_Dock_Area_Top
            // 
            this._ListCompare_Toolbars_Dock_Area_Top.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._ListCompare_Toolbars_Dock_Area_Top.BackColor = System.Drawing.SystemColors.Control;
            this._ListCompare_Toolbars_Dock_Area_Top.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Top;
            this._ListCompare_Toolbars_Dock_Area_Top.ForeColor = System.Drawing.SystemColors.ControlText;
            this._ListCompare_Toolbars_Dock_Area_Top.Location = new System.Drawing.Point(0, 0);
            this._ListCompare_Toolbars_Dock_Area_Top.Name = "_ListCompare_Toolbars_Dock_Area_Top";
            this._ListCompare_Toolbars_Dock_Area_Top.Size = new System.Drawing.Size(752, 25);
            this._ListCompare_Toolbars_Dock_Area_Top.ToolbarsManager = this.utmListCompare;
            // 
            // _ListCompare_Toolbars_Dock_Area_Bottom
            // 
            this._ListCompare_Toolbars_Dock_Area_Bottom.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this._ListCompare_Toolbars_Dock_Area_Bottom.BackColor = System.Drawing.SystemColors.Control;
            this._ListCompare_Toolbars_Dock_Area_Bottom.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Bottom;
            this._ListCompare_Toolbars_Dock_Area_Bottom.ForeColor = System.Drawing.SystemColors.ControlText;
            this._ListCompare_Toolbars_Dock_Area_Bottom.Location = new System.Drawing.Point(0, 577);
            this._ListCompare_Toolbars_Dock_Area_Bottom.Name = "_ListCompare_Toolbars_Dock_Area_Bottom";
            this._ListCompare_Toolbars_Dock_Area_Bottom.Size = new System.Drawing.Size(752, 0);
            this._ListCompare_Toolbars_Dock_Area_Bottom.ToolbarsManager = this.utmListCompare;
            // 
            // ListCompare
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(752, 577);
            this.Controls.Add(this.ListCompare_Fill_Panel);
            this.Controls.Add(this._ListCompare_Toolbars_Dock_Area_Left);
            this.Controls.Add(this._ListCompare_Toolbars_Dock_Area_Right);
            this.Controls.Add(this._ListCompare_Toolbars_Dock_Area_Bottom);
            this.Controls.Add(this._ListCompare_Toolbars_Dock_Area_Top);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "ListCompare";
            this.Text = "ListCompare";
            this.Load += new System.EventHandler(this.ListCompare_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ugCompareList)).EndInit();
            this.ultraPanel1.ClientArea.ResumeLayout(false);
            this.ultraPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.utmListCompare)).EndInit();
            this.ListCompare_Fill_Panel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private Infragistics.Win.UltraWinGrid.UltraGrid ugCompareList;
        private Infragistics.Win.Misc.UltraPanel ultraPanel1;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsManager utmListCompare;
        private System.Windows.Forms.Panel ListCompare_Fill_Panel;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _ListCompare_Toolbars_Dock_Area_Left;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _ListCompare_Toolbars_Dock_Area_Right;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _ListCompare_Toolbars_Dock_Area_Bottom;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _ListCompare_Toolbars_Dock_Area_Top;
        private System.Windows.Forms.Button btnResetExchangeAlias;
        private System.Windows.Forms.Button btnResetCrmAlias;
    }
}