using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;

using System.Text;
using System.Windows.Forms;

namespace MSC.WinFormControlLib
{
    [ToolboxBitmap(typeof(DataGridView))]
    public partial class dgvHasRowNum : DataGridView
    {
        public dgvHasRowNum()
        {
            InitializeComponent();
            this.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgvHasRowNum_RowPostPaint);
        }

        void dgvHasRowNum_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            Rectangle rect = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y,
                dgv.RowHeadersWidth - 4, e.RowBounds.Height);

            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dgv.RowHeadersDefaultCellStyle.Font,
                rect,
                dgv.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

    }
}
