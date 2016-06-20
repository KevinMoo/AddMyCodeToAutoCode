using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MSC.WinFormControlLib
{
    public partial class FrmColumnSet : Form
    {
        public FrmColumnSet(DataGridView _dgv)
        {
            InitializeComponent();
            dgv = _dgv;
        }
        private DataGridView dgv;

        private void frmColumnSet_Load(object sender, EventArgs e)
        {
            if (dgv == null) return;
            TreeNode node;
            DataGridViewColumn column;
            DataGridViewColumn tmpcol;

            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    tmpcol=dgv.Columns[j];
                    if (tmpcol.DisplayIndex == i)
                    {
                        column = dgv.Columns[j];
                        if (column.Tag == null || (int)(column.Tag) != 0) //不显示的隐藏列初始化为tag＝0
                        {
                            node = new TreeNode(column.HeaderText);
                            node.Tag = column;
                            node.Checked = column.Visible;
                            tv.Nodes.Add(node);
                        }
                    }
                }
            }

        }


        private void btnAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < tv.Nodes.Count; i++)
            {
                tv.Nodes[i].Checked = true;
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < tv.Nodes.Count; i++)
            {
                tv.Nodes[i].Checked = false;
            }
        }

        private void btnReverse_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < tv.Nodes.Count; i++)
            {
                tv.Nodes[i].Checked = !tv.Nodes[i].Checked;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            DataGridViewColumn column;
            for (int i = 0; i < tv.Nodes.Count; i++)
            {
                column = (DataGridViewColumn)tv.Nodes[i].Tag;
                column.Visible = tv.Nodes[i].Checked;
            }
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}