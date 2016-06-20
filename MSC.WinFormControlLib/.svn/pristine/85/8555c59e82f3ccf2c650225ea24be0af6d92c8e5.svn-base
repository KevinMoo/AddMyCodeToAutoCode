using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;
using System.Reflection;
using MSC.CommonLib;


namespace MSC.WinFormControlLib
{
    public partial class frmManagerBill : Form
    {
       
        string[] whereString = new string[1];
        private string mastBllClassName = "";
        private string mastModelClassName = "";
        private string detlBllClassName = "";
        private string detlModelClassName = "";
        private const string BLL_ASSEMBLE = "MSC.MSC.WinFormControlLib.BLL";
        private const string MODEL_ASSEMBLE = "MSC.MSC.WinFormControlLib.Model";

        /// <summary>
        /// 编辑窗口
        /// </summary>
        private frmCommonInputBill editForm;




        public frmManagerBill(frmCommonInputBill pEditForm, string pBllClassName, string pModelClassName,
            string pDetlBllClassName,string pDetlModelClassName)
        {
            InitializeComponent();
            mastBllClassName = pBllClassName;
            mastModelClassName = pModelClassName;
            detlBllClassName = pDetlBllClassName;
            detlModelClassName = pDetlModelClassName;
            editForm = pEditForm;

            InitGridMaster();
            InitGridDetl();
        }

        private void InitGridDetl()
        {
            ModelColumnInfo[] infos = GetModelColumnInfos(detlModelClassName);
            InitGrid(dgvDetl, infos);
        }

        private void InitGridMaster()
        {
            ModelColumnInfo [] infos = GetModelColumnInfos(mastModelClassName);
            InitGrid(dgvMast, infos);

        }

        /// <summary>
        /// 初始化表格
        /// </summary>
        /// <param name="pDgv"></param>
        /// <param name="pColumnNames"></param>
        private void InitGrid(dgvHasRowNum pDgv, ModelColumnInfo[] pInfos)
        {
            DataTable dt = new DataTable();
            DataColumn dc = null;

            foreach(ModelColumnInfo info in pInfos)
            {
                dc = new DataColumn(info.ColumnDescription);
                dt.Columns.Add(dc);
            }

            pDgv.DataSource = dt;
        }

        private ModelColumnInfo[] GetModelColumnInfos(string pModelClassName)
        {
            Assembly modelAssembly = Assembly.Load(MODEL_ASSEMBLE);
            Type modelType = modelAssembly.GetType(MODEL_ASSEMBLE + "." + pModelClassName);
            FieldInfo fi = modelType.GetField("ColumnInfos");
            object obj = new object();
            return (ModelColumnInfo[])fi.GetValue(null);
        }



        /// <summary>
        /// 绑定数据
        /// </summary>
        /// <param name="pDestination">目的表格</param>
        /// <param name="pDataSource">数据源</param>
        /// <param name="pColnumNum">列数</param>
        private void BindData(dgvHasRowNum pDestination, DataTable pDataSource, int pColnumNum)
        {

            DataTable destinationTable = (DataTable)pDestination.DataSource;

            //先清空
            destinationTable.Rows.Clear();

            for (int i = 0; i < pDataSource.Rows.Count; i++)
            {
                DataRow newDR = destinationTable.NewRow();
                for (int j = 0; j < pColnumNum; j++)
                {
                    newDR[j] = pDataSource.Rows[i][j];
                }
                destinationTable.Rows.Add(newDR);
            }
        }

        private void toolStripQuery_Click(object sender, EventArgs e)
        {
            ModelColumnInfo []  infos = GetModelColumnInfos(mastModelClassName);
            whereString[0] = null;
            dbQuery frmQuery = new dbQuery(infos, whereString);
            frmQuery.ShowDialog();
            if (whereString[0] == null)
            {
                return;
            }
            DataSet ds = GetDataSet(mastBllClassName,whereString[0]);
            BindData(dgvMast, ds.Tables[0], ds.Tables[0].Columns.Count);

       }

        private DataSet GetDataSet(string pBllClassName,string pWhere)
        {
            Assembly bllAssembly = Assembly.Load(BLL_ASSEMBLE);
            Type bllType = bllAssembly.GetType(BLL_ASSEMBLE + "." + pBllClassName);
            object bllObj = Activator.CreateInstance(bllType);
            //利用MethodInfo类来获得从指定类中符合条件的成员函数
            MethodInfo mi = bllType.GetMethod("GetList",
                BindingFlags.Public | BindingFlags.Instance, null, new Type[] { typeof(string) }, null);
            try
            {
                DataSet ds = (DataSet)(mi.Invoke(bllObj, new object[] { pWhere }));
                return ds;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.InnerException.Message);
                return null;
            }
 
        }

        private void toolStripExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dgvMast_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            editBill();
        }

        private void editBill()
        {
            try
            {
                if (dgvMast.CurrentRow.Cells[0].Value == null ||
                    dgvMast.CurrentRow.Cells[0].Value.ToString() == string.Empty)
                {
                    return;
                }
            }
            catch
            {
                return;
            }

            string billNo = dgvMast.CurrentRow.Cells[0].Value.ToString();
             //绑定数据
            editForm.BindData(billNo);
            try
            {
                editForm.Show();
                editForm.Activate();
            }
            catch
            {

            }


        }

        private void toolStripEdit_Click(object sender, EventArgs e)
        {
            editBill();
        }

        private void frmManagerBill_FormClosing(object sender, FormClosingEventArgs e)
        {
            editForm.Close();
            editForm.Dispose();
        }

        private void dgvMast_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            //先清空明细表的内容
            DataTable detailTable = (DataTable)dgvDetl.DataSource;
            if (detailTable != null)
            {
                detailTable.Rows.Clear();
            }

            if (dgvMast.Rows[e.RowIndex].Cells[0].Value == null
                || dgvMast.Rows[e.RowIndex].Cells[0].Value.ToString() == string.Empty)
            {
                return;
            }
            string billNo = dgvMast.Rows[e.RowIndex].Cells[0].Value.ToString();

            DataSet ds = GetDataSet(detlBllClassName,"BillNum='"+billNo+"'");



            BindData(dgvDetl, ds.Tables[0],ds.Tables[0].Columns.Count );
        }
    }
}
