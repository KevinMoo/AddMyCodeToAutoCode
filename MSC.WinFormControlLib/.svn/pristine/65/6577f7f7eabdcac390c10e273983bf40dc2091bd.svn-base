using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MSC.WinFormControlLib
{
    /// <summary>
    /// DataGridView自动加合计行类
    /// </summary>
    public class DataGridViewAddSumRow
    {
        private DataGridView dgv = null;
        private DataTable dt = null;
        //private MSC.CommonLib.MyDataView _myDataView;
        public DataGridViewAddSumRow(DataGridView pDgv, DataTable pDt)
        {
            this.dgv = pDgv;
            this.dt = pDt;
            this.begin();
            //_myDataView = (MSC.CommonLib.MyDataView)pDt.DefaultView;
           // _myDataView.OnMyRowFilterChanged += new MSC.CommonLib.MyDataView.MyRowFilterChanged(_myDataView_OnMyRowFilterChanged);
        }

        //void _myDataView_OnMyRowFilterChanged(object sender, EventArgs e)
        //{
        //    this.dataGridView_DataSourceChanged(this.dgv, new EventArgs());
        //}
        /**/
        /// <summary>
        /// 设置表格的数据源
        /// </summary>
        public DataTable dataTableName
        {
            set
            {
                this.dt = value;
            }
        }
        /**/
        /// <summary>
        ///传递表格的名称
        /// </summary>
        public DataGridView DgvName
        {
            set
            {
                dgv = value;
            }
        }
        /**/
        /// <summary>
        /// 开始添加合计行
        /// </summary>
        public void begin()
        {
            initDgv();
        }
        private void initDgv()
        {
            if (dgv != null)
            {

                this.dgv.DataSourceChanged += new EventHandler(dataGridView_DataSourceChanged);
                this.dgv.ColumnHeaderMouseClick += new DataGridViewCellMouseEventHandler(dataGridView_ColumnHeaderMouseClick);
                this.dgv.CellValueChanged += new DataGridViewCellEventHandler(dataGridView_CellValueChanged);
                this.dgv.AllowUserToAddRows = false;
                dgv.DataSource = dt;
            }
        }
        /**/
        /// <summary>
        /// 计算合计算
        /// </summary>
        /// <param name="dgv">要计算的DataGridView</param>
        private void SumDataGridView(DataGridView dgv)
        {
            if (dgv.DataSource == null) return;
            //DataTable dt = (DataTable)dgv.DataSource;
            if (dt.Rows.Count < 1) return;
            decimal[] tal = new decimal[dt.Columns.Count];

            DataRow ndr = dt.NewRow();

            string talc = "";

            int number = 1;
            foreach (DataRow dr in dt.Rows)
            {
                dr["@xu.Hao"] = number++;
                int n = 0;
                foreach (DataColumn dc in dt.Columns)
                {


                    if (talc == "" && dc.DataType.Name.ToUpper().IndexOf("STRING") >= 0)
                    { talc = dc.ColumnName; }


                    if (dc.DataType.IsValueType)
                    {
                        string hej = dr[dc.ColumnName].ToString();
                        try
                        {
                            if (hej != string.Empty) tal[n] += decimal.Parse(hej);
                        }
                        catch (Exception) { }
                        //if (hej != string.Empty) tal[n] += decimal.Parse(hej);
                    }


                    n++;
                }
            }

            ndr.BeginEdit();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (tal[i] != 0)
                    ndr[i] = tal[i];
            }
            ndr["@xu.Hao"] = ((int)(dt.Rows.Count + 1)).ToString();
            if (talc != "") ndr[talc] = "合计";
            ndr.EndEdit();
            dt.Rows.Add(ndr);

            dgv.Rows[dgv.Rows.Count - 1].DefaultCellStyle.BackColor = Color.FromArgb(255, 222, 210);
            dgv.Rows[dgv.Rows.Count - 1].ReadOnly = true;

            if (dgv.Tag == null)
            {
                foreach (DataGridViewColumn dgvc in dgv.Columns)
                {
                    dgvc.SortMode = DataGridViewColumnSortMode.Programmatic;
                }
            }
            dgv.Tag = ndr;
        }
        private void dataGridView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            
            DataGridView sortDgv = (DataGridView)sender;
            if (e.Button != MouseButtons.Left)
            {
                //不是左键不排序了
                return;
            }
            if (sortDgv.Rows.Count == 0)
            {
                return;
            }
            int fx = 0;
            if (sortDgv.AccessibleDescription == null)
            {
                fx = 1;
            }
            else
            {
                fx = int.Parse(sortDgv.AccessibleDescription);
                fx = (fx == 0 ? 1 : 0);
            }
            sortDgv.AccessibleDescription = fx.ToString();
            if (sortDgv.Columns[e.ColumnIndex].Name == "@xu.Hao") return;
            DataGridViewColumn nCol = sortDgv.Columns[e.ColumnIndex];

            if (nCol.DataPropertyName == string.Empty) return;

            if (nCol != null)
            {
                sortDgv.Sort(nCol, fx == 0 ? ListSortDirection.Ascending : ListSortDirection.Descending);

            }
            //--
            DataRow dr = (DataRow)sortDgv.Tag;
            DataTable dt = (DataTable)sortDgv.DataSource;
            DataRow ndr = dt.NewRow();
            ndr.BeginEdit();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ndr[i] = dr[i];
            }
            dt.Rows.Remove(dr);


            //if (e.ColumnIndex != 0)
            {
                int n = 1;
                for (int i = 0; i < sortDgv.Rows.Count; i++)
                {
                    DataGridViewRow dgRow = sortDgv.Rows[i];
                    DataRowView drv = (DataRowView)dgRow.DataBoundItem;
                    DataRow tdr = drv.Row;
                    tdr.BeginEdit();
                    tdr["@xu.Hao"] = n;
                    n++;
                    tdr.EndEdit();

                }
                sortDgv.Refresh();
                sortDgv.RefreshEdit();

            }
            ndr["@xu.Hao"] = ((int)(dt.Rows.Count + 1)).ToString();
            ndr.EndEdit();
            dt.Rows.Add(ndr);
            sortDgv.Tag = ndr;

            //--
            sortDgv.Sort(sortDgv.Columns["@xu.Hao"], ListSortDirection.Ascending);
            sortDgv.Columns["@xu.Hao"].HeaderCell.SortGlyphDirection = SortOrder.None;
            nCol.HeaderCell.SortGlyphDirection = fx == 0 ? SortOrder.Ascending : SortOrder.Descending;
            sortDgv.Rows[sortDgv.Rows.Count - 1].DefaultCellStyle.BackColor = Color.FromArgb(255, 222, 210);

        }

        private void dataGridView_DataSourceChanged(object sender, EventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            //DataTable dt = (DataTable)dgv.DataSource;
            if (dt == null) return;
            decimal[] tal = new decimal[dt.Columns.Count];
            if (dt.Columns.IndexOf("@xu.Hao") < 0)
            {
                DataColumn dc = new DataColumn("@xu.Hao", System.Type.GetType("System.Int32"));
                dt.Columns.Add(dc);

                //显示最后
                dgv.Columns["@xu.Hao"].DisplayIndex = dt.Columns.Count - 1;

                dgv.Columns["@xu.Hao"].HeaderText = "序号";

                dgv.Columns["@xu.Hao"].SortMode = DataGridViewColumnSortMode.Programmatic;
                dgv.AutoResizeColumn(dgv.Columns["@xu.Hao"].Index);

                dgv.Columns["@xu.Hao"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv.Columns["@xu.Hao"].Visible = true;
            }
            SumDataGridView(dgv);
        }
        private void dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            DataGridView dgv = (DataGridView)sender;
            if (dgv.Tag == null || e.RowIndex < 0 || e.RowIndex == dgv.Rows.Count - 1) return;

            string col = dgv.Columns[e.ColumnIndex].DataPropertyName;
            if (col == string.Empty) return;
            if (((DataRowView)dgv.Rows[e.RowIndex].DataBoundItem).Row.Table.Columns[col].DataType.IsValueType)
            {
                decimal tal = 0;
                foreach (DataGridViewRow dgvr in dgv.Rows)
                {
                    if (dgvr.Index != dgv.Rows.Count - 1)
                    {
                        string hej = dgvr.Cells[e.ColumnIndex].Value.ToString();
                        if (hej != string.Empty) tal += decimal.Parse(hej);
                    }
                }
                if (tal == 0)
                    dgv[e.ColumnIndex, dgv.Rows.Count - 1].Value = DBNull.Value;
                else
                    dgv[e.ColumnIndex, dgv.Rows.Count - 1].Value = tal;
            }
        }

    }
}
