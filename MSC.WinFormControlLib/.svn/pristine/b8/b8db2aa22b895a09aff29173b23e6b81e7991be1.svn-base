using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using MSC.CommonLib;



namespace MSC.WinFormControlLib
{
    /// <summary>
    /// 返回条件字段串
    /// </summary>
    public partial class dbQuery : Form
    {
        private ModelColumnInfo[] fieldInfos = null;
        private DateTimePicker dtpDate=new DateTimePicker();

        

        private string [] groups = {"AND","OR"};
        private string [] expressions={"=","LIKE","<>",">",">=","<","<="};



        private string[] result;


        protected const string CR = "\r\n";
        protected const string SHOW_ALL = "*";


        //dgvQuery各列的含义

        /// <summary>
        /// 非列
        /// </summary>
        private const int COLUMN_NOT = 0;

        /// <summary>
        /// 字段列
        /// </summary>
        private const int COLUMN_FIELD = 1;

        /// <summary>
        /// 表达式列
        /// </summary>
        private const int COLUMN_EXP = 2;

        /// <summary>
        /// 值列
        /// </summary>
        private const int COLUMN_VALUE = 3;

        /// <summary>
        /// 组合方式列
        /// </summary>
        private const int COLUMN_GROUP = 5;

        /// <summary>
        /// 只取日期
        /// </summary>
        private const int COLUMN_IS_DATE = 4;






        /// <summary>
        /// 通用单表查询类,以*返回所有字段
        /// </summary>
        /// <param name="pConString">连接数据库的字符串</param>
        /// <param name="pTableName">要查询的数据库表名</param>
        /// <param name="pResultTable">返回的结果表</param>
        public dbQuery(ModelColumnInfo [] pFields,string [] pResult)
        {
            
            //系统初始化
            InitializeComponent();

            this.fieldInfos = pFields;
            result = pResult;

            //初始化表达式
            InitExpression();

            //初始化组合方式
            InitGroup();


            //初始化字段名
            InitField();

            dtpDate.ValueChanged += new EventHandler(dtpDate_ValueChanged);



            dgvQuery.Controls.Add(dtpDate);
            dtpDate.Visible = false;

        }




        /// <summary>
        /// 当日期选择的值改变了，将值赋给DataGridVew(dgvQuery).UserValue.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void dtpDate_ValueChanged(object sender, EventArgs e)
        {
            //如果当前单元格不空，且是Value单元格，就把值赋给value单元格
            if (dgvQuery.CurrentCell != null && dgvQuery.CurrentCell.ColumnIndex == COLUMN_VALUE)
            {
                if (dtpDate.Format == DateTimePickerFormat.Custom)
                {
                    dgvQuery.CurrentCell.Value = dtpDate.Value.ToString();
                }
                else
                {
                    dgvQuery.CurrentCell.Value = dtpDate.Value.ToShortDateString();
                }
                
            }
        }

        /// <summary>
        /// 初始化表的字段名
        /// </summary>
        private void InitField()
        {
            foreach (ModelColumnInfo model in fieldInfos)
            {
                DbField.Items.Add(model);
            }
            DbField.ValueMember = "Self";
            DbField.DisplayMember = "ColumnDescription";
            
        }

        /// <summary>
        /// 初始化组合方式
        /// </summary>
        private void InitGroup()
        {
            Group.DataSource = groups;
        }

        /// <summary>
        /// 初始化表达式组合框
        /// </summary>
        private void InitExpression()
        {
            Expression.DataSource = expressions;
        }


        private void toolStripQuery_Click(object sender, EventArgs e)
        {
            string whereText = "";

            if (!CheckExpression(ref whereText, false))
            {
                DialogBox.ShowError(whereText);
                return;
            }

            this.result[0] = whereText;
            this.Close();
        }


       
        
        /// <summary>
        /// 检验表达式的合法性，如果成功并返回where 表达式
        /// </summary>
        /// <param name="pWhereText">where表达式</param>
        /// <param name="pIsSave">返回的表达式是否用于保存！</param>
        /// <returns></returns>
        private bool CheckExpression(ref string pWhereText,bool pIsSave)
        {
            for (int i = 0; i < (this.dgvQuery.Rows.Count - 1); i++)
            {
                if ((this.dgvQuery.Rows[i].Cells[3].Value == null) || (this.dgvQuery.Rows[i].Cells[3].Value.ToString() == string.Empty))
                {
                    this.dgvQuery.Rows.RemoveAt(i--);
                }
            }
            if (this.dgvQuery.Rows.Count == 0)
            {
                pWhereText = "请先设定查询的条件！";
                return false;
            }
            string columnName = "";
            string str2 = "";
            string str3 = "";
            for (int j = 0; j < (this.dgvQuery.Rows.Count - 1); j++)
            {
                if ((this.dgvQuery.Rows[j].Cells[0].Value != null) && ((bool)this.dgvQuery.Rows[j].Cells[0].Value))
                {
                    pWhereText = pWhereText + " NOT ";
                }
                if ((this.dgvQuery.Rows[j].Cells[1].Value == null) || (this.dgvQuery.Rows[j].Cells[1].Value.ToString() == string.Empty))
                {
                    pWhereText = "第" + ((j + 1)).ToString() + "行,字段为空";
                    return false;
                }
                ModelColumnInfo info = this.dgvQuery.Rows[j].Cells[1].Value as ModelColumnInfo;
                columnName = info.ColumnName;
                if ((info.ColumnType == typeof(DateTime)) && Convert.ToBoolean(this.dgvQuery.Rows[j].Cells[4].Value))
                {
                    columnName = "CONVERT(datetime,CONVERT(varchar(10)," + columnName + ",101))";
                }
                if ((this.dgvQuery.Rows[j].Cells[2].Value == null) || (this.dgvQuery.Rows[j].Cells[2].Value.ToString() == string.Empty))
                {
                    pWhereText = "第" + ((j + 1)).ToString() + "行,表达式空！";
                    return false;
                }
                str2 = this.dgvQuery.Rows[j].Cells[2].Value.ToString();
                if ((info.ColumnType == typeof(DateTime)) && Convert.ToBoolean(this.dgvQuery.Rows[j].Cells[4].Value))
                {
                    str3 = Convert.ToDateTime(this.dgvQuery.Rows[j].Cells[3].Value).ToShortDateString();
                }
                else
                {
                    str3 = this.dgvQuery.Rows[j].Cells[3].Value.ToString();
                }
                if ((str2.ToUpper() == "LIKE") && !pIsSave)
                {
                    string str4 = pWhereText;
                    pWhereText = str4 + " " + columnName + " " + str2 + " '%" + str3 + "%'";
                }
                else
                {
                    string str5 = pWhereText;
                    pWhereText = str5 + " " + columnName + " " + str2 + " '" + str3 + "'";
                }
                if (j != (this.dgvQuery.Rows.Count - 2))
                {
                    if ((this.dgvQuery.Rows[j].Cells[5].Value == null) || (this.dgvQuery.Rows[j].Cells[5].Value.ToString() == string.Empty))
                    {
                        pWhereText = "第" + ((j + 1)).ToString() + "行,组合条件为空";
                        return false;
                    }
                    pWhereText = pWhereText + " " + this.dgvQuery.Rows[j].Cells[5].Value.ToString();
                }
                if (pIsSave)
                {
                    pWhereText = pWhereText + CR;
                }
            }
            return true;

        }


        private void toolStripExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void dgvQuery_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            dgv.Rows[e.RowIndex].Cells[COLUMN_NOT].Value = false;
            dgv.Rows[e.RowIndex].Cells[COLUMN_VALUE].Value = "";

            //设定默认值
            dgv.Rows[e.RowIndex].Cells[COLUMN_GROUP].Value = "AND";
            dgv.Rows[e.RowIndex].Cells[COLUMN_EXP].Value = "=";
            

        }

        private void dgvQuery_CurrentCellChanged(object sender, EventArgs e)
        {
            DataGridView view = (DataGridView)sender;
            bool flag = false;
            //日期
            if (((view.CurrentCell != null) && (view.CurrentCell.ColumnIndex == 3)) 
                && ((view.CurrentRow.Cells[1].Value != null) 
                && (view.CurrentRow.Cells[1].Value.ToString() != string.Empty)))
            {
                string columnName = (view.CurrentRow.Cells[1].Value as ModelColumnInfo).ColumnName;
                if ((view.CurrentRow.Cells[1].Value as ModelColumnInfo).ColumnType == typeof(DateTime))
                {
                    Rectangle rectangle = view.GetCellDisplayRectangle(view.CurrentCell.ColumnIndex, view.CurrentCell.RowIndex, false);
                    this.dtpDate.Top = rectangle.Top;
                    this.dtpDate.Left = rectangle.Left;
                    this.dtpDate.Height = rectangle.Height;
                    this.dtpDate.Width = rectangle.Width;
                    try
                    {
                        DateTime time = Convert.ToDateTime(view.CurrentCell.Value);
                        this.dtpDate.Value = time;
                    }
                    catch
                    {
                        this.dtpDate.Value = DateTime.Now;
                    }
                    this.dtpDate.Format = DateTimePickerFormat.Custom;
                    this.dtpDate.CustomFormat = "yyyy年MM月dd日 hh:mm:ss";
                    flag = true;
                }
            }
            this.dtpDate.Visible = flag;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.result[0] = "1=1";
            this.Close();
        }
    }
}