using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExchangeSync
{
    public partial class TraceViewer : Form
    {
        public TraceViewer()
        {
            InitializeComponent();
        }

        private void TraceViewer_Load(object sender, EventArgs e)
        {
            if (!Trace.Null)
            {
                if (!Trace.Empty)
                {
                    ugSyncLog.DataSource = Trace.RetrieveLog();
                }
            }
        }

        private void ugSyncLog_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
        {
            Color color = Color.Black;

            switch (e.Row.Cells[0].Value.ToString())
            {
                case "Warning":
                    color = Color.Orange;
                    break;
                case "Error":
                    color = Color.Red;
                    break;
                default:
                    color = Color.Black;
                    break;
            }

            for (int i = 0; i < e.Row.Cells.Count; i++)
            {
                e.Row.Cells[i].Appearance.ForeColor = color;
            }

        }
    }
}
