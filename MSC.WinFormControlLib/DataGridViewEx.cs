/****************************************************************************************************
 * Copyright (C) 2010 ����½���Ƽ��ɷ����޹�˾ ��Ȩû�У����⿽����ʹ�ã�����ʹ����ɵ��κκ�������κ����Σ����࿪ԴӰ�죬��ͬ����
 * �� �� ����DataGridViewEx.cs
 * �� �� �ˣ������ʿ
 * Email:nzj.163@163.com   qq:342155124
 * ����ʱ�䣺2010-06-01
 * ����޸�ʱ�䣺2012-1-19  ���ӵ�10����ʾ�Ĺ��ܣ��޸�����ͷ����26�еĴ��󣬵���excelΪ���鷽ʽ���ٶȸ��죬�����ĵ�Ԫ������Ϊ�ı���ʽ��
 * ��    �⣺�û��Զ����DataGridView�ؼ�
 * ������������չDataGridView�ؼ�����
 * ��չ���ܣ�
 * 1������Search(); ������ͬ��������������ͬ F3Ϊ��ݼ�������������
 * 2����TreeView HeadSource �����ø��ӵı�����ʽ�����ĳ���ڵ��Ӧ����ʾ�����أ��뽫�ýڵ�Tag����Ϊhide�������е�����λ���������Ԫ��λ�ö�Ӧ����Ҷ�ڵ��˳����Ҫ����������˳��һ��
 * 3��ͨ�����䵼��Excel����������com���������ExportExcel() ���������������ƣ���ͷͬ�����Ե���,AutoFit�������õ���excel���Ƿ��Զ�������Ԫ����
 *    ��������֧���Զ���ģ�Title List<string> Header   List<string> Footer,֧�������ʱֵ���趨�����ڹر�ʱExcel��Դ�Զ������ͷ�
 * 4�������Լ������趨��Щ����ʾ������ʾ��ͨ�����÷���SetColumnVisible()ʵ�֡�
 * 5�������б���SetHeader(),��������Զ�ɼ�AlwaysShowCols(),��������ʱ���ɼ�HideCols()
 *    ע�⣬��ʹ����TreeView��Ϊ����Headerʱ����Ҫʹ�ñ�������Header��ʾ�����ݸ���treeview���ݶ���ʾ
 * 6���п�ȼ�˳��ı���SaveGridView()������LoadGridView()
 * 7��֧�����������õĴ�ӡ���ܣ���������
 *     private void button5_Click(object sender, EventArgs e)
        {
            DGVPrinter printer = new DGVPrinter();
            printer.PrintPreviewDataGridView(DataGridViewEx1);
        }
 * 8���Զ���ϲ������У��кϲ��� MergeRowColumn ���ԣ��кϲ���MergeColumnNames���ԣ������Զ�������
 * 9���б�ŵ����� bool ShowRowNumber;
 * 10���������һ�еĻ����У�֧���еľۺϺ������μ�http://msdn.microsoft.com/zh-cn/library/system.data.datacolumn.expression(v=VS.100).aspx
 *     �����id����ʾ���ϼơ��ַ���avgPrice����ƽ��ֵ��total����ʾ�ϼƣ����ComputeColumns�����������ݣ�id,�ϼƣ���avgPrice,Avg(avgPrice)��total,Sum(total)
 *     �����Ҫ��ֵ���и�ʽ���ƣ���ʵ��beforeShow�¼�
 *     �����˵����ʹ�ӡ��Ӧ��֧�֣����������õĶ��뷽Ӧ����ʽ��������ӡ��
 ****************************************************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Drawing.Design;
using System.Linq;
using System.Data;
using System.Collections;

namespace MSC.WinFormControlLib
{
    public partial class DataGridViewEx : DataGridView
    {
        public DataGridViewEx()
        {
            InitializeComponent();
            footer = new List<string>();
            header = new List<string>();
        }

        /// <summary>
        /// �����¼�����������ʾ�����еĸ�ʽ����
        /// </summary>
        /// <param name="sender">����</param>
        /// <param name="value">�е�ֵ</param>
        /// <returns>���ص��ӹ����ֵ</returns>
        public delegate string BeforeShow(string sender,string value);
        public event BeforeShow beforeShow;

        #region private paras ---------------------------------------
        //---------------------------------------search paras------------------------------------------------------
        private string searchValue = "";  // Ҫ�������ַ���
        private int currentRow=0;       // ��ǰ�������кţ�������������ʱ��
        //----------------------------------------���ӱ�ͷ���ϲ���-------------------------------------------------
        private List<string> _mergecolumnname = new List<string>();
        private bool showRowNumber=false;   //�Ƿ���ʾ�к�
        //---------------------------------------excell export------------------------------------------------------
        object objApp_Late;
        object objBook_Late;
        object objBooks_Late;
        object objSheets_Late;
        object objSheet_Late;
        object objRange_Late;
        object[] Parameters;
        int visibleCols;

        private int treeDepth;
        private int iNodeLevels;                                    // ����������
        private int iCellHeight=22;                                 // һά�б����ĸ߶�
        private IList<TreeNode> ColLists = new List<TreeNode>();    // ������ʾ��ҳ�ڵ�
        private IList<TreeNode> allColLists = new List<TreeNode>();    // ���е�ҳ�ڵ�

        private string title;
        private List<string> header;
        private List<string> footer;
        private bool autoFit=true;
        /// <summary>
        /// �ӳ־û������п�
        /// </summary>
        private bool loadColWidthFromDb = false;


        private List<string> _mergeRowColumn = new List<string>();
        private List<List<string>> _lstMergeRowColumn = new List<List<string>>();
        private List<string> computeColumns = new List<string>();
        private void resetMergeRowColumns()
        {
            _lstMergeRowColumn.Clear();
            foreach (string mergeRowTemp in _mergeRowColumn)
            {
                char[] cs = { ',' };
                string[] rowColumns = mergeRowTemp.Split(cs, StringSplitOptions.RemoveEmptyEntries);
                if (rowColumns.Length >= 2)
                {
                    List<string> tempStringLst = new List<string>();

                    for (int r = 0; r < rowColumns.Length; r++)
                    {
                        if (rowColumns[r].Trim().Length > 0)
                        {
                            if (!Columns.Contains(rowColumns[r]) || Columns[rowColumns[r]].Displayed)
                                tempStringLst.Add(rowColumns[r].Trim());
                        }
                    }
                    if (tempStringLst.Count < 2) continue;

                    bool added = false;
                    for (int r = 0; r < _lstMergeRowColumn.Count; r++)
                    {
                        if (_lstMergeRowColumn[r][0] == tempStringLst[0] && _lstMergeRowColumn[r].Count < tempStringLst.Count)
                        {
                            _lstMergeRowColumn.Insert(r, tempStringLst);
                            added = true;
                            break;
                        }
                    }
                    if (!added)
                    {
                        _lstMergeRowColumn.Add(tempStringLst);
                    }
                }
            }
        }

        internal bool ReCalculateRowMerge(int rowNo, string columnsName, out string value, out int mergeCount)
        {
            mergeCount = checkHasMergeTheColumn(columnsName, rowNo, out value) - 1;
            return mergeCount >= 1;
        }



        private bool loadedFinish = false;

        /// <summary>
        /// �Ƿ����������ݣ�����������    
        /// ��ʹ��merge��ʱ������û��޸����ݣ�Ӧ������ˢ�����ô�ˢ����ʾЧ����
        /// Ϊ�˼���ˢ�´����������˴�����
        /// </summary>
        private bool LoadedFinish
        {
            get { return loadedFinish; }
            set
            {
                loadedFinish = value;
                if (loadedFinish)
                {
                    stockOfAllRows.Clear();
                    resetColumnDisplayArray();
                }
            }
        }
        #endregion

        #region properties-------------------------------------------

        /// <summary>
        /// ���û��ȡ�ϲ��еļ���
        /// </summary>
        [MergableProperty(false)]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        //[DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
        [Localizable(true)]
        [Description("���û��ȡ�ϲ��еļ���"), Browsable(true), Category("�Զ�������")]
        public List<string> MergeColumnNames
        {
            get
            {
                return _mergecolumnname;
            }
            set
            {
                _mergecolumnname = value;
            }
        }

        /// <summary>
        /// ����ͷ
        /// </summary>
        [Bindable(true), Category("�Զ�������"), Description("�����title")]
        public string Title
        {
            get { return title; }
            set { title = value; }
        }

        /// <summary>
        /// ��ͷ��Ҫд�������
        /// </summary>
        [Bindable(true), Category("�Զ�������"), Description("�����header��Ҫд�������")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> Header
        {
            get { return header; }
            set { header = value; }
        }

        /// <summary>
        /// ҳ������
        /// </summary>
        [Bindable(true), Category("�Զ�������"), Description("�����footer��Ҫд�������")]
        //[DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]  ֻ��ȥ����仰�ſ��������ʱ���丳ֵ
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> Footer
        {
            get { return footer; }
            set { footer = value; }
        }

        /// <summary>
        /// ���������Ƿ��Զ��������еĿ��
        /// </summary>
        [Description("���������Ƿ��Զ��������еĿ��"), Category("�Զ�������")]
        public bool AutoFit
        {
            get { return autoFit; }
            set { autoFit = value; }
        }


        /// <summary>
        /// �Ƿ���ʾ�к�
        /// </summary>
        [Description("�趨�Ƿ���ʾ�к�"), Category("�Զ�������")]
        public bool ShowRowNumber
        {
            get { return showRowNumber; }
            set { showRowNumber = value; }
        }


        /// <summary>
        /// ���û��ȡ�ϲ��еļ���
        /// </summary>
        [MergableProperty(false)]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
        [Localizable(true)]
        [Description("���û��ȡ�кϲ���������,�����ö��ŷָ����������ö����ͬ�ĺϲ��м���"), Browsable(true), Category("�Զ�������")]
        public List<string> MergeRowColumn
        {
            get
            {
                return _mergeRowColumn;
            }
            set
            {
                _mergeRowColumn = value;
                resetMergeRowColumns();
            }
        }

        /// <summary>
        /// ���û��ȡ�����еļ���
        /// </summary>
        [MergableProperty(false)]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
        [Localizable(true)]
        [Description("���û��ȡ��������,ÿ��������һ��,���id����sum����,������Ϊ:id,sum(id)"), Browsable(true), Category("�Զ�������")]
        public List<string> ComputeColumns
        {
            get{return computeColumns;}
            set { computeColumns = value; }
        }
        #endregion

        #region -------------------public functions------------------
        /// <summary>
        /// ��ǰ�������Ϣ���������Դ�Сд
        /// Ĭ��Ϊ�������пɼ��У�ÿ����ʾ������������
        /// </summary>
        public void Search()
        {
            string searchCol = "";  //Ҫ�������ֶ������ö������ӵ��ַ�����

            //��������ʾ�������ö��ŷָ����ַ���
            foreach (DataGridViewColumn dgvColumn in this.Columns)
            {
                if (dgvColumn.Visible == true)
                {
                    searchCol += dgvColumn.Name + ",";
                }
            }

            if (searchCol == "")
            {
                MessageBox.Show("û�пɼ��������У�", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                InputBox frm = new InputBox();
                DialogResult dlrResult = frm.ShowDialog();
                this.searchValue = frm.SearchValue;

                if (dlrResult == DialogResult.OK)
                {
                    //ȥ�����һ������
                    searchCol = searchCol.Substring(0, searchCol.Length - 1);
                    this.Search(searchCol, searchValue, currentRow);
                }
            }
        }

        /// <summary>
        /// ��ǰ�������Ϣ���������Դ�Сд��ָ���������У����ݼ���ʼ��
        /// </summary>
        /// <param name="searchCol">Ҫ�������У��ö��ŷָ����ַ�����</param>
        /// <param name="searchValue">������ֵ</param>
        /// <param name="startRow">��������ʼ��</param>
        public void Search(string searchCol, string searchValue,  int startRow)
        {
            int find = 0;
            this.searchValue = searchValue;
            if (this.Rows.Count < startRow) return;

            string[] searchCols = searchCol.Split(',');
            int foundCol = 0;
            try
            {
                for (int i = startRow; i < this.Rows.Count; i++)
                {

                    string str = "";
                    foreach (string curSearchName in searchCols)
                    {
                        if (this[curSearchName.Trim(), i].Value != null)
                        {
                            str += this[curSearchName.Trim(), i].Value.ToString() + ";";
                        }
                    }

                    startRow = i;

                    if (str.IndexOf(searchValue, 0, StringComparison.OrdinalIgnoreCase) > -1)
                    {
                        find = 1;
                        int fountPos = str.IndexOf(searchValue, 0, StringComparison.OrdinalIgnoreCase);
                        string subStr = str.Substring(0, fountPos + searchValue.Length);
                        string[] weidu = subStr.Split(';');
                        foundCol = weidu.Length - 1;
                        startRow++;
                        currentRow = startRow;
                        break;
                    }
                    else
                        find = 0;
                }

                if (find == 0)
                {
                    currentRow = 0;
                    MessageBox.Show("�Ҳ����������������ݣ�", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    this.CurrentCell = this[searchCols[foundCol], currentRow - 1];
                }

                this.Focus();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "��������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>
        /// �����������ݵ���Ϊexcel
        /// </summary>
        public void ExportExcel()
        {
            Cursor.Current = Cursors.WaitCursor;
            Export2Excel(this, true,true); //�Ż���excel�������ٶȼӿ�
            Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// �趨��ǰ������Щ��ʾ���������ز�����Ҫ�趨���У��趨����tag��0
        /// TreeView�ı���ͷ���ܽ��д��趨
        /// </summary>
        public void SetColumnVisible()
        {
            if (_ColHeaderTreeView.Nodes.Count>0) return; //���ڸ��ӽṹ���Ͳ����Լ�������Щ�п�����ʾ
            FrmColumnSet frm = new FrmColumnSet(this);
            frm.ShowDialog();
        }

        /// <summary>
        /// ���ÿɼ��м��б��⣬��ʾ��˳���봫���˳����ͬ,��ʹ����TreeView��ΪHeaderʱ����Ҫʹ�ñ�����
        /// �����˷������ú����õ��в��ǿɼ��ģ�������SetColumnVisible()����ͬ��
        /// �����Ҫ����ĳЩ����Զ�ɼ����ڵ��ô˷����󣬵��÷���AlwaysShowCols��������
        /// ���ҪʹĳЩ����ʱ���ɼ���ͨ��SetColumnVisible()��������Ϊ�ɼ��������HideCols()����
        /// </summary>
        /// <param name="columns">�ö��ŷָ��������</param>
        /// <param name="headers">�ö��ŷָ����ͷ����Ҫ��columns��Գ���</param>
        public void SetHeader(string columns, string headers)
        {
            if (_ColHeaderTreeView != null && _ColHeaderTreeView.Nodes.Count > 0) return;
            string[] cols = columns.Split(',');
            string[] heads = headers.Split(',');
            if (cols.Length != heads.Length)
            {
                throw new Exception("�趨����ʾ���������б�����������ԣ��������趨��");
            }

            //����������
            foreach (DataGridViewColumn curColumn in this.Columns)
            {
                curColumn.Tag = 0;
                curColumn.Visible = false;
            }
            try
            {
                for (int i = 0; i < cols.Length; i++)
                {
                    if (this.Columns[cols[i].ToString()] != null)
                    {
                        this.Columns[cols[i].ToString()].Tag = null;
                        this.Columns[cols[i].ToString()].DisplayIndex = i;
                        this.Columns[cols[i].ToString()].HeaderText = heads[i].ToString();
                        this.Columns[cols[i].ToString()].Visible = true;
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// �������������Ϊ��Զ�ɼ�
        /// </summary>
        /// <param name="columns">�ö��ŷָ��������</param>
        public void AlwaysShowCols(string columns)
        {
            string[] cols = columns.Split(',');
            try
            {
                foreach (string col in cols)
                {
                    this.Columns[col].Tag = 0;
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// ����ĳЩ����ʱ���ɼ�,���ŷָ������
        /// </summary>
        /// <param name="columns">�ö��ŷָ������</param>
        public void HideCols(string columns)
        { 
            string[] cols=columns.Split(',');
            try
            {
                foreach (string col in cols)
                {
                    this.Columns[col].Visible = false;
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// �����е����ã���˳��Ϊ������,��ʾ˳��,�п�,�Ƿ�ɼ�(True,false�ַ���)
        /// �Լ���Ҫ�����ص����ݳ־û�,ͨ���־û���Ҫ�ٷ�װһ�Σ������û���ʹ��λ�õı�ʶ��������ͬ�û��Ϳ����ж���������
        /// </summary>
        /// <returns>���ص�����Ϊdictionary[0]����ʾ˳��(dictionary[0])[0],�п�(dictionary[0])[1]���Ƿ�ɼ�(dictionary[0])[2]</returns>
        public Dictionary<string, string[]> SaveGridView()
        {
            Dictionary<string, string[]> dict = new Dictionary<string, string[]>();

            for (int i = 0; i < this.Columns.Count; i++)
            {
                foreach (DataGridViewColumn col in this.Columns)
                {
                    if (col.DisplayIndex == i)
                    {
                        string[] value = { col.DisplayIndex.ToString(), col.Width.ToString(), col.Visible.ToString() };
                        dict.Add(col.Name, value);
                    }
                }
            }
            return dict;
        }

        /// <summary>
        /// �п�ļ��أ��Ӵ�����ֵ��ڼ����û����������Ϣ
        /// ���������Ӧ�ð�����ʾ˳����С�������������
        /// ���־û�������Ϣ���أ�һ��Ӧ������ǰ�û����ؼ�ʹ��λ����Ϣ
        /// </summary>
        /// <param name="dict">Dictionary<����,{��ʾ˳��,�п�,�Ƿ�ɼ�(True,false�ַ���)}></param>
        public void LoadGridView(Dictionary<string, string[]> dict)
        {
            if (dict.Count <= 0) return;
            loadColWidthFromDb = true;
            try
            {
                foreach (string dic in dict.Keys)
                {
                    if (this.Columns.Contains(dic))
                    {
                        DataGridViewColumn column = this.Columns[dic];
                        column.DisplayIndex = Convert.ToInt32(dict[dic][0]);
                        column.Visible = Convert.ToBoolean(dict[dic][2]);
                        if (column.Visible)
                        {
                            column.Width = Convert.ToInt32(dict[dic][1]) > 20 ? Convert.ToInt32(dict[dic][1]) : 120; //С��20���Զ�����Ϊ120.
                        }
                    }
                }
            }
            catch (Exception e) 
            {
                loadColWidthFromDb = false;
                throw e; 
            }
            loadColWidthFromDb = false;
        }
        #endregion

        #region --------------private functions----------------------
       
        /// <summary>
        /// Exports a passed datagridview to an Excel worksheet.
        /// If captions is true, grid headers will appear in row 1.
        /// Data will start in row 2.
        /// </summary>
        /// <param name="datagridview"></param>
        /// <param name="captions"></param>
        private void Export2Excel(DataGridView datagridview, bool captions)
        {
            int kk = 0;
            foreach (DataGridViewColumn col in datagridview.Columns)
            {
                if (col.GetType().Name == "DataGridViewTextBoxColumn" && col.Visible == true)// 
                {
                    kk++;
                }
            }
            visibleCols = kk;
            string[] headers = new string[kk];
            string[] columns = new string[kk];
            string[] colName = new string[kk];

            int i = 0;
            int c = 0;
            int m = 0;

            for (c = 0; c < datagridview.Columns.Count; c++)
            {
                for (int j = 0; j < datagridview.Columns.Count; j++)
                {
                    DataGridViewColumn tmpcol = datagridview.Columns[j];
                    if (tmpcol.DisplayIndex == c)
                    {
                        if (tmpcol.GetType().Name == "DataGridViewTextBoxColumn" && tmpcol.Visible) //����ʾ�������г�ʼ��Ϊtag��0 
                        {
                            headers[c - m] = tmpcol.HeaderText;
                            i = c - m+1;
                            columns[c - m] = ConvertColumnNum2String(i);
                            colName[c - m] = tmpcol.Name;
                        }
                        else
                        {
                            m++;
                        }
                        break;
                    }
                }
            }

            try
            {
                // Get the class type and instantiate Excel.
                Type objClassType;
                objClassType = Type.GetTypeFromProgID("Excel.Application");
                objApp_Late = Activator.CreateInstance(objClassType);
                //Get the workbooks collection.
                objBooks_Late = objApp_Late.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, objApp_Late, null);
                //Add a new workbook.
                objBook_Late = objBooks_Late.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, objBooks_Late, null);
                //Get the worksheets collection.
                objSheets_Late = objBook_Late.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, objBook_Late, null);
                //Get the first worksheet.
                Parameters = new Object[1];
                Parameters[0] = 1;
                objSheet_Late = objSheets_Late.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, objSheets_Late, Parameters);

                if (captions)
                {
                    //��������д��excel��ͷ
                    //��ά��ͷ
                    //_ColHeaderTreeView  ��ͷTreeView
                    //iNodeLevels ��ͷ�Ĳ���
                    //_ColHeaderTreeView.Nodes.a
                    if (this._ColHeaderTreeView.Nodes.Count > 0)
                    {
                        TreeView tr = new TreeView();
                        CopyTree(_ColHeaderTreeView.Nodes, tr.Nodes);
                        WriteCell(tr.Nodes, 1, 1);
                    }
                    else
                    {
                        // Create the headers in the first row of the sheet
                        for (c = 0; c < kk; c++)
                        {
                            //Get a range object that contains cell.
                            Parameters = new Object[2];
                            Parameters[0] = columns[c] + "1";
                            Parameters[1] = Missing.Value;
                            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, Parameters);
                            //Write Headers in cell.
                            Parameters = new Object[1];
                            Parameters[0] = headers[c];
                            objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);
                        }
                    }
                }

                // Now add the data from the grid to the sheet starting in row 2

                int startRow = 2;
                if (iNodeLevels >1)
                {
                    startRow = iNodeLevels + 1;
                }

                for (i = 0; i < datagridview.RowCount; i++)
                {
                    c = 0;
                    foreach (string txtCol in colName)
                    {
                        DataGridViewColumn col = datagridview.Columns[txtCol];
                        if (col.Visible)
                        {
                            //Get a range object that contains cell.
                            Parameters = new Object[2];

                            Parameters[0] = columns[c] + Convert.ToString(i + startRow);
                            Parameters[1] = Missing.Value;
                            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, Parameters);
                            //Write Headers in cell.
                            Parameters = new Object[1];
                            if (datagridview.Rows[i].Cells[col.Name].Value != null)
                            {
                                Parameters[0] = datagridview.Rows[i].Cells[col.Name].Value.ToString().Replace(" 0:00:00", ""); //string.Format("{0:d}",dt);//2005-11-5
                            }
                            else
                            {
                                Parameters[0] = "";
                            }
                            objRange_Late.GetType().InvokeMember("NumberFormatLocal", BindingFlags.SetProperty, null, objRange_Late, colTypeToxlsType(col));
                            objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);
                            c++;
                        }

                    }
                }
                //���Һϲ���Ԫ��
                mergeRowsCell(startRow);

                //���ºϲ���Ԫ��
                //start row = startRow, end row= startRow+ startRow+datagridview.RowCount
                //calculate column index tobe bombined
                int endRow = startRow + datagridview.RowCount - 1;
                if (_mergecolumnname != null && _mergecolumnname.Count > 0)
                {
                    foreach (string cn in _mergecolumnname)
                    {
                        int col = Array.IndexOf<string>(colName, cn) + 1;
                        if (col > 0) mergeColumnsCell(startRow, endRow, col);//���õĺϲ���û����ʾ�Ĳ����кϲ�����
                    }
                }

                drawLine();

                // д��footer
                int maxRows = iNodeLevels + 1 + datagridview.RowCount;
                if (iNodeLevels == 0) maxRows++;
                if (footer != null) writeFooter(maxRows);

                //д��header
                if (header != null) writeHeader();

                if (title != null) writeTitle();

                //Return control of Excel to the user.
                Parameters = new Object[1];
                Parameters[0] = true;
                objApp_Late.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, objApp_Late, Parameters);
                objApp_Late.GetType().InvokeMember("UserControl", BindingFlags.SetProperty, null, objApp_Late, Parameters);
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        /// <summary>
        /// Exports a passed datagridview to an Excel worksheet.
        /// If captions is true, grid headers will appear in row 1.
        /// Data will start in row 2.
        /// </summary>
        /// <param name="datagridview"></param>
        /// <param name="captions"></param>
        private void Export2Excel(DataGridView datagridview, bool captions,bool fast)
        {
            int kk = 0;
            foreach (DataGridViewColumn col in datagridview.Columns)
            {
                if (col.GetType().Name == "DataGridViewTextBoxColumn" && col.Visible == true)// 
                {
                    kk++;
                }
            }
            visibleCols = kk;
            string[] headers = new string[kk];
            string[] columns = new string[kk];
            string[] colName = new string[kk];

            int i = 0;
            int c = 0;
            int m = 0;

            for (c = 0; c < datagridview.Columns.Count; c++)
            {
                for (int j = 0; j < datagridview.Columns.Count; j++)
                {
                    DataGridViewColumn tmpcol = datagridview.Columns[j];
                    if (tmpcol.DisplayIndex == c)
                    {
                        if (tmpcol.GetType().Name == "DataGridViewTextBoxColumn" && tmpcol.Visible) //����ʾ�������г�ʼ��Ϊtag��0 
                        {
                            headers[c - m] = tmpcol.HeaderText;
                            i = c - m + 1;
                            columns[c - m] = ConvertColumnNum2String(i);
                            colName[c - m] = tmpcol.Name;
                        }
                        else
                        {
                            m++;
                        }
                        break;
                    }
                }
            }

            try
            {
                // Get the class type and instantiate Excel.
                Type objClassType;
                objClassType = Type.GetTypeFromProgID("Excel.Application");
                objApp_Late = Activator.CreateInstance(objClassType);
                //Get the workbooks collection.
                objBooks_Late = objApp_Late.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, objApp_Late, null);
                //Add a new workbook.
                objBook_Late = objBooks_Late.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, objBooks_Late, null);
                //Get the worksheets collection.
                objSheets_Late = objBook_Late.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, objBook_Late, null);
                //Get the first worksheet.
                Parameters = new Object[1];
                Parameters[0] = 1;
                objSheet_Late = objSheets_Late.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, objSheets_Late, Parameters);

                if (captions)
                {
                    //��������д��excel��ͷ
                    //��ά��ͷ
                    //_ColHeaderTreeView  ��ͷTreeView
                    //iNodeLevels ��ͷ�Ĳ���
                    //_ColHeaderTreeView.Nodes.a
                    if (this._ColHeaderTreeView.Nodes.Count > 0)
                    {
                        TreeView tr = new TreeView();
                        CopyTree(_ColHeaderTreeView.Nodes, tr.Nodes);
                        WriteCell(tr.Nodes, 1, 1);
                    }
                    else
                    {
                        // Create the headers in the first row of the sheet
                        for (c = 0; c < kk; c++)
                        {
                            //Get a range object that contains cell.
                            Parameters = new Object[2];
                            Parameters[0] = columns[c] + "1";
                            Parameters[1] = Missing.Value;
                            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, Parameters);
                            //Write Headers in cell.
                            Parameters = new Object[1];
                            Parameters[0] = headers[c];
                            objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);
                        }
                    }
                }

                // Now add the data from the grid to the sheet starting in row 2

                int startRow = 2;
                if (iNodeLevels != 0)
                {
                    startRow = iNodeLevels + 1;
                }

                //ʹ��2ά�������ճ�����ٶȸ���
                string[,] vals=new string[this.Rows.Count,headers.Length];
                string startCell = columns[0] + Convert.ToString(startRow);
                string endCell = columns[columns.Length-1] + Convert.ToString(startRow+ this.Rows.Count-1);

                //TODO: writh to excel
                for (i = 0; i < datagridview.RowCount; i++)
                {
                    int j=0;
                    
                    if (i < datagridview.NewRowIndex)
                    #region data
                    {
                        foreach (string txtCol in colName)
                        {
                            DataGridViewColumn col = datagridview.Columns[txtCol];
                            if (col.Visible)
                            {
                                if (datagridview.Rows[i].Cells[col.Name].Value != null)
                                {
                                    vals[i, j] = datagridview.Rows[i].Cells[col.Name].Value.ToString().Replace(" 0:00:00", ""); //string.Format("{0:d}",dt);//2005-11-5
                                }
                                else
                                {
                                    vals[i, j] = "";
                                }
                                j++;
                            }
                        }
                    }
                    #endregion data
                    else if (i == datagridview.NewRowIndex && computeColumns.Count > 0)
                    #region sum footer
                    {
                        Dictionary<string, string> dict = dictComputeColumns();
                        foreach (string txtCol in colName)
                        {
                            string v;
                            if (dict.TryGetValue(txtCol, out v))
                            {
                                vals[i, j] = v;
                            }
                            j++;
                        }
                    }
                    #endregion sum footer
                }

                //��д��
                Parameters = new Object[2];
                Parameters[0] = startCell;
                Parameters[1] = endCell;
                objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, Parameters);
                //��Ϊ�ı���ʽ
                object[] o = new object[1];
                o[0] = "@";
                objRange_Late.GetType().InvokeMember("NumberFormatLocal", BindingFlags.SetProperty, null, objRange_Late,o);
                //д��ֵ
                Parameters = new Object[1];
                Parameters[0] = vals;
                objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);

                //���Һϲ���Ԫ��
                mergeRowsCell(startRow);

                //���ºϲ���Ԫ��
                //start row = startRow, end row= startRow+ startRow+datagridview.RowCount
                //calculate column index tobe bombined
                int endRow = startRow + datagridview.RowCount - 1;
                if (_mergecolumnname != null && _mergecolumnname.Count > 0)
                {
                    foreach (string cn in _mergecolumnname)
                    {
                        int col = Array.IndexOf<string>(colName, cn) + 1;
                        if (col > 0) mergeColumnsCell(startRow, endRow, col);//���õĺϲ���û����ʾ�Ĳ����кϲ�����
                    }
                }

                drawLine();
                

                
                int maxRows = iNodeLevels + 1 + datagridview.RowCount;
                if (iNodeLevels == 0) maxRows++;

                setAlignment(colName, columns,maxRows); //���õ������ݵĶ��뷽ʽ

                // д��footer
                if (footer != null) writeFooter(maxRows);

                //д��header
                if (header != null) writeHeader();

                if (title != null) writeTitle();

                //Return control of Excel to the user.
                Parameters = new Object[1];
                Parameters[0] = true;
                objApp_Late.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, objApp_Late, Parameters);
                objApp_Late.GetType().InvokeMember("UserControl", BindingFlags.SetProperty, null, objApp_Late, Parameters);
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        
        private object[] colTypeToxlsType(DataGridViewColumn col)
        {
            object[] o=new object[1];
            switch (col.ValueType.ToString())
            {
                case "System.DateTime":
                    o[0]="yyyy-m-d";
                    break;
                case "System.Int16":
                case "System.Int32":
                case "System.Int64":
                case "System.long":
                case "System.float":
                case "System.decimal":
                    o[0] = "0.00_ ";
                    break;
                default:
                    o[0]="@";
                    break;
            }
            return o;
        }


        //����tree�������صĽڵ�ȥ��
        private void CopyTree(TreeNodeCollection nodes, TreeNodeCollection tgNodes)
        {
            foreach (TreeNode nd in nodes)
            {
                if (nd.Tag == null || nd.Tag.ToString() != "hide")
                {
                    TreeNode newNode = new TreeNode(nd.Text);
                    tgNodes.Add(newNode);
                    if (nd.Nodes.Count > 0) CopyTree(nd.Nodes, newNode.Nodes);
                }
            }
        }

        private void removeHideNode(TreeNode node)
        {
            if (node.Tag != null && node.Tag.ToString() == "hide")
            {
                node.Remove();
            }
            else
            {
                for (int i = node.Nodes.Count - 1; i >= 0; i--)
                {
                    TreeNode nd = node.Nodes[i];
                    removeHideNode(nd);
                }
            }
        }

        /// <summary>
        /// ����treeview��д��
        /// </summary>
        /// <param name="rootNodes">treeview �ڵ㼯��</param>
        /// <param name="excelRow">д�����:��1��ʼ</param>
        /// <param name="excelColumn">д�����:��1��ʼ</param>
        /// <returns>ռ�õ�����</returns>
        private int WriteCell(TreeNodeCollection rootNodes, int excelRow, int excelColumn)
        {
            int cellCount = 1;
            if (rootNodes.Count < 1)
            {
                WriteExcelROW(excelRow - 1, excelColumn, cellCount, treeDepth);//sdΪtreeview�����
                return cellCount;
            }
            int cellCountSum = 0;
            foreach (TreeNode TNode in rootNodes)
            {
                cellCount = WriteCell(TNode.Nodes, excelRow + 1, excelColumn);
                WriteExcel(excelRow, excelColumn, cellCount, TNode.Text);
                cellCountSum += cellCount;
                excelColumn += cellCount;
            }
            return cellCountSum;
        }

        private void WriteExcelROW(int excelRow, int excelColumn, int cellCountSum, int sd)
        {
            //����û������
            string point1 = ConvertColumnNum2String(excelColumn) + excelRow.ToString();//��Ԫ����ʼ�� ��:A1 
            string point2 = ConvertColumnNum2String(excelColumn) + sd.ToString();//��Ԫ����ʼ�� ��:A1 
            if (point1 != point2)
            {
                objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
                objRange_Late.GetType().InvokeMember("MergeCells", BindingFlags.SetProperty, null, objRange_Late, new object[] { true });
            }

        }

        /// <summary>
        /// ����header�����
        /// </summary>
        /// <param name="excelRow"></param>
        /// <param name="excelColumn"></param>
        /// <param name="cellCountSum"></param>
        /// <param name="excelValue"></param>
        private void WriteExcel(int excelRow, int excelColumn, int cellCountSum, string excelValue)
        {
            //����û������
            string point1 = ConvertColumnNum2String(excelColumn) + excelRow.ToString();//��Ԫ����ʼ�� ��:A1
            string point2 = ConvertColumnNum2String(excelColumn + cellCountSum- 1) + excelRow.ToString();//��Ԫ������� ��:B4

            //_Range = _Worksheet.get_Range(point1, point2);//��ȡ��Ԫ��
            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
            if (cellCountSum > 0)
            {
                //_Range.MergeCells = true; //�ϲ���Ԫ��
                objRange_Late.GetType().InvokeMember("MergeCells", BindingFlags.SetProperty, null, objRange_Late, new object[] { true });
            }

            objRange_Late.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });//����ˮƽ���� Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter 
            objRange_Late.GetType().InvokeMember("VerticalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });  //���ݴ�ֱ����Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter 

            //_Worksheet.Cells[_Range.Row, _Range.Column] = excelValue;//������д�뵥Ԫ��
            Parameters = new Object[1];
            Parameters[0] = excelValue;
            objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);
        }
        /// <summary>
        /// ����footer�����
        /// </summary>
        /// <param name="excelRow"></param>
        /// <param name="excelColumn"></param>
        /// <param name="cellCountSum"></param>
        /// <param name="excelValue"></param>
        private void WriteExcels(int excelRow, int excelColumn, int cellCountSum, string excelValue)
        {
            //����û������..
            string point1 = ConvertColumnNum2String(excelColumn) + excelRow.ToString();//��Ԫ����ʼ�� ��:A1
            string point2 = ConvertColumnNum2String(excelColumn + cellCountSum- 1).ToString() + excelRow.ToString();//��Ԫ������� ��:B4


            //_Range = _Worksheet.get_Range(point1, point2);//��ȡ��Ԫ��
            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
            if (cellCountSum > 0)
            {
                //_Range.MergeCells = true; //�ϲ���Ԫ��
                objRange_Late.GetType().InvokeMember("MergeCells", BindingFlags.SetProperty, null, objRange_Late, new object[] { true });
            }
            //_Worksheet.Cells[_Range.Row, _Range.Column] = excelValue;//������д�뵥Ԫ��
            Parameters = new Object[1];
            Parameters[0] = excelValue;
            objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);
            objRange_Late.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4131 });//����� Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft 
        }

        /// <summary>
        /// ����title�����
        /// </summary>
        /// <param name="excelRow"></param>
        /// <param name="excelColumn"></param>
        /// <param name="cellCountSum"></param>
        /// <param name="excelValue"></param>
        private void WriteExcelTitle(int excelRow, int excelColumn, int cellCountSum, string excelValue)
        {
            string point1 = ConvertColumnNum2String(excelColumn) + excelRow.ToString();//��Ԫ����ʼ�� ��:A1
            string point2 = ConvertColumnNum2String(excelColumn + cellCountSum - 1).ToString() + excelRow.ToString();//��Ԫ������� ��:B4


            //_Range = _Worksheet.get_Range(point1, point2);//��ȡ��Ԫ��
            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
            if (cellCountSum > 0)
            {
                //_Range.MergeCells = true; //�ϲ���Ԫ��
                objRange_Late.GetType().InvokeMember("MergeCells", BindingFlags.SetProperty, null, objRange_Late, new object[] { true });
            }

            objRange_Late.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });//����ˮƽ���� Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter 
            objRange_Late.GetType().InvokeMember("VerticalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });  //���ݴ�ֱ����Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter 

            object Font = objRange_Late.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, objRange_Late, null);
            objRange_Late.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, Font, new object[] { 20 });

            //_Worksheet.Cells[_Range.Row, _Range.Column] = excelValue;//������д�뵥Ԫ��
            Parameters = new Object[1];
            Parameters[0] = excelValue;
            objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);
        }

        /// <summary>
        /// ������
        /// </summary>
        /// <param name="Index">�������</param>
        /// <param name="value">�����ֵ</param>
        /// <param name="isTitle">�Ƿ�Ϊtitle</param>
        private void Insert(int Index, string value, bool isTitle)
        {
            object EntireRow_Late;

            Parameters = new Object[2];
            Parameters[0] = "A1";
            Parameters[1] = Missing.Value;

            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, Parameters);
            EntireRow_Late = objRange_Late.GetType().InvokeMember("EntireRow", BindingFlags.GetProperty, null, objRange_Late, null);
            EntireRow_Late.GetType().InvokeMember("Insert", BindingFlags.InvokeMethod, null, EntireRow_Late, null);
            if (!isTitle)
            {
                WriteExcels(Index, 1, visibleCols, value);
            }
            else
            {
                WriteExcelTitle(Index, 1, visibleCols, value);
            }

        }

        /// <summary>
        /// д���ݵ����ݻ�ʵ��
        /// </summary>
        private void drawLine()
        {
            object Range = objSheet_Late.GetType().InvokeMember("UsedRange", BindingFlags.GetProperty, null, objSheet_Late, null);
            object[] args = new object[] { 1 };
            object Borders = Range.GetType().InvokeMember("Borders", BindingFlags.GetProperty, null, Range, null);
            Borders = Range.GetType().InvokeMember("LineStyle", BindingFlags.SetProperty, null, Borders, args);
            //�Զ���Ӧ���еĿ��
            if (autoFit)
            {
                Object dataColumns = Range.GetType().InvokeMember("Columns", BindingFlags.GetProperty, null, Range, null);
                dataColumns.GetType().InvokeMember("AutoFit", BindingFlags.InvokeMethod, null, dataColumns, null);
            }
        }

        //TODO:excel alignment
        /// <summary>
        /// �������е����ݶ��뷽ʽ����ʾ��һ��
        /// </summary>
        /// <param name="colName">��������</param>
        /// <param name="colName">Excel��������</param>
        /// <param name="maxRows">������ֵ</param>
        private void setAlignment(string[] colName,string[] columns, int maxRows)
        {
            //Ĭ�������
            string point1; 
            string point2;

            int startRow = 2;
            if (iNodeLevels >1)
            {
                startRow = iNodeLevels + 1;
            }
            int i = 0;
            
            foreach (string col in colName)
            {
                point1 = columns[i] + startRow.ToString();//��Ԫ����ʼ�� ��:A3
                point2 = columns[i] + maxRows.ToString();//��Ԫ������� ��:A15
                objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
                objRange_Late.GetType().InvokeMember("VerticalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });  //���ݴ�ֱ����Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter 
                if (this.Columns[col].DefaultCellStyle.Alignment.ToString().Contains("Right"))
                {
                    objRange_Late.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4152 });//Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight 
                }
                else if (this.Columns[col].DefaultCellStyle.Alignment.ToString().Contains("Center"))
                {
                    objRange_Late.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });//����ˮƽ���� Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter 
                }
                i++;
            }
        }

        /// <summary>
        /// �����кϲ�
        /// </summary>
        /// <param name="startRow"></param>
        private void mergeRowsCell(int startRow)
        {
            reCalculateAllStock();
            objApp_Late.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, objApp_Late, new object[] { false });
            foreach (int i in stockOfAllRows.Keys)
            {
                foreach (string tempKeys in stockOfAllRows[i].Keys)
                {
                    {
                        string theValue = getValue(startRow + i, displayedColumns.IndexOf(tempKeys) + 1);
                        string[] points = new string[2];

                        int smallvalue = 10000, bigvalue=0;
                        if(stockOfAllRows[i][tempKeys].Count<2) continue ;
                        for (int r = 0; r < stockOfAllRows[i][tempKeys].Count; r++)
                        {
                            int tempvalue = displayedColumns.IndexOf(stockOfAllRows[i][tempKeys][r]) + 1;
                            if (smallvalue > tempvalue)
                                smallvalue = displayedColumns.IndexOf(stockOfAllRows[i][tempKeys][r]) + 1;
                            if (bigvalue < tempvalue)
                                bigvalue = tempvalue;                           
                        }
                        points[0] = ConvertColumnNum2String(smallvalue) + (startRow + i).ToString();
                        points[1] = ConvertColumnNum2String(bigvalue) + (startRow + i).ToString();
                        objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, points);
                        objRange_Late.GetType().InvokeMember("Merge", BindingFlags.InvokeMethod, null, objRange_Late, null);

                    }

                }
            }

        }
        /// <summary>
        /// ���ºϲ���Ԫ��
        /// </summary>
        /// <param name="startRow">��ʼ��</param>
        /// <param name="endRow">��ֹ��</param>
        /// <param name="col">�ڼ���</param>
        private void mergeColumnsCell(int startRow, int endRow, int excelColumn)
        {
            objApp_Late.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, objApp_Late, new object[] { false });
            for (int i = endRow; i > startRow; i--)
            {
                if (displayedColumns.Count <= excelColumn || excelColumn < 1 ||
                    (checkHasTheColumn(displayedColumns[excelColumn - 1], i - startRow).Length > 0
                    || checkHasTheColumn(displayedColumns[excelColumn - 1], i - 1 - startRow).Length > 0))
                    continue;

                string s = getValue(i, excelColumn);

                if (!string.IsNullOrEmpty(s) && s == getValue(i - 1, excelColumn))
                {
                    string point1 = ConvertColumnNum2String(excelColumn) + i.ToString();
                    string point2 = ConvertColumnNum2String(excelColumn) + (i - 1).ToString();
                    objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
                    objRange_Late.GetType().InvokeMember("Merge", BindingFlags.InvokeMethod, null, objRange_Late, null);
                }
            }
        } 
       
        private string getValue(int row, int col)
        {
            string point1 = ConvertColumnNum2String(col) + row.ToString();
            string point2 = ConvertColumnNum2String(col) + row.ToString();
            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
            Parameters = new Object[1];
            Parameters[0] = Missing.Value;
            object rt = objRange_Late.GetType().InvokeMember("Value", BindingFlags.GetProperty, null, objRange_Late, Parameters);
            if (rt==null)
                return string.Empty;
            else
                return rt.ToString();
        }

        /// <summary>
        /// д��title
        /// </summary>
        private void writeTitle()
        {
            if (!string.IsNullOrEmpty(title))
            {
                Insert(1, title, true);
            }
        }

        /// <summary>
        /// д��header
        /// </summary>
        private void writeHeader()
        {
            for (int i = header.Count - 1; i >= 0; i--)
            {
                Insert(1, header[i], false);
            }
        }

        /// <summary>
        /// д��footer
        /// </summary>
        /// <param name="maxRows">��ʼд�����</param>
        private void writeFooter(int maxRows)
        {
            for (int i = 0; i < footer.Count; i++)
            {
                WriteExcels(maxRows + i, 1, visibleCols, footer[i]);
            }
        }

        /// <summary>
        /// ���õݹ���������,��ͷ�ܸ�
        /// </summary>
        private void myNodeLevels()
        {

            iNodeLevels = 1;//��ʼֵΪ1

            ColLists.Clear();
            allColLists.Clear();

            int iNodeDeep = myGetNodeLevels(_ColHeaderTreeView.Nodes);
            treeDepth = iNodeDeep;

            this.ColumnHeadersHeight = iCellHeight * iNodeDeep;//��ͷ�ܸ�=һά�и�*����
            this.ColumnDeep = iNodeDeep;

        }
        /// <summary>
        /// �ݹ��������������,�������е�Ҷ�ڵ�
        /// </summary>
        /// <param name="tnc"></param>
        /// <returns></returns>
        private int myGetNodeLevels(TreeNodeCollection tnc)
        {
            if (tnc == null) return 0;

            foreach (TreeNode tn in tnc)
            {
                if ((tn.Level + 1) > iNodeLevels)//tn.Level�Ǵ�0��ʼ��
                {
                    iNodeLevels = tn.Level + 1;
                }

                if (tn.Nodes.Count > 0)
                {
                    myGetNodeLevels(tn.Nodes);
                }
                else
                {
                    allColLists.Add(tn);
                    if (tn.Tag == null || tn.Tag.ToString() != "hide")
                    {
                        ColLists.Add(tn);//ҳ�ڵ�
                    }
                }
            }

            return iNodeLevels;
        }

        /// <summary>
        /// ����ת��ΪExcel��ĸ���ֵ���
        /// </summary>
        /// <param name="columnNum">����</param>
        /// <returns>�����ַ���</returns>
        private string ConvertColumnNum2String(int columnNum)
        {
            if (columnNum > 26)
            {
                return string.Format("{0}{1}", (char)(((columnNum - 1) / 26) + 64), (char)(((columnNum - 1) % 26) + 65));
            }
            else
            {
                return ((char)(columnNum + 64)).ToString();
            }
        }
        /// <summary>
        /// Excel��ĸ��ʽ����ת��Ϊ����
        /// </summary>
        /// <param name="letters">�ַ�</param>
        /// <returns>���ص�intֵ</returns>
        private int ConvertLetters2ColumnName(string letters)
        {
            int num = 0;
            if (letters.Length == 1)
            {
                num = Convert.ToInt32(letters[0]) - 64;
            }
            else if (letters.Length == 2)
            {
                num = (Convert.ToInt32(letters[0]) - 64) * 26 + Convert.ToInt32(letters[1]) - 64;
            }
            return num;
        }

        /// <summary>
        /// ����Ԫ��
        /// </summary>
        /// <param name="e"></param>
        private void DrawCell(DataGridViewCellPaintingEventArgs e)
        {
            if (e.CellStyle.Alignment == DataGridViewContentAlignment.NotSet)
            {
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            Brush gridBrush = new SolidBrush(this.GridColor);
            SolidBrush backBrush = new SolidBrush(e.CellStyle.BackColor);
            SolidBrush fontBrush = new SolidBrush(e.CellStyle.ForeColor);
                        
            //������ͬ������
            int upRows = 0;
            //������ͬ������
            int downRows = 0;
            //���Ҫ�ϲ�����
            int leftColumns = 0;
            //�ұ�Ҫ�ϲ�����
            int rightColumns = 0;

            Pen gridLinePen = new Pen(gridBrush);
            bool mergeRow = false;
            resetMergeRowColumns();
            #region ����ϲ�,���������ϲ��ܲ���
            if (e.RowIndex != -1 && this._lstMergeRowColumn.Count > 0 && displayedColumns.Count >0)
            {
                foreach (List<string> lstRowColumn in _lstMergeRowColumn)
                {
                    //����ϲ���,��������ؼ��в���ʾ,�Ͳ�������;
                    if (mergeRow || !Columns[lstRowColumn[0]].Visible ) break;

                    if (!lstRowColumn.Contains(this.Columns[e.ColumnIndex].Name)) continue;
                    #region ��֤������һ��
                    
                    leftColumns = 0;
                    rightColumns = 0;
                    for (int i = displayedColumns.IndexOf(Columns[e.ColumnIndex].Name) - 1; i >= 0; i--)
                    {
                        if (lstRowColumn.Contains(displayedColumns[i]))                            
                        { 
                            leftColumns++;
                            continue;
                        } 
                        break ;
                    }
                    for (int i = displayedColumns.IndexOf(Columns[e.ColumnIndex].Name) + 1; i < displayedColumns.Count; i++)
                    {
                        if (lstRowColumn.Contains(displayedColumns[i]))
                        {
                            rightColumns++;
                            continue;
                        }
                        break;
                    }
                    if (leftColumns + rightColumns + 1 != lstRowColumn.Count) continue;
                    #endregion

                    string haveTheItem = checkHasTheColumn(Columns[e.ColumnIndex].Name, e.RowIndex);

                    if (haveTheItem.Length > 0)
                    {
                        stockOfAllRows[e.RowIndex].Remove(haveTheItem);
                    }

                    #region ��֤ͷΪ�ǿ�,����Ϊ��
                
                    if (null != this.Rows[e.RowIndex].Cells[lstRowColumn[0]].Value )
                    {
                        string mainValue = this.Rows[e.RowIndex].Cells[lstRowColumn[0]].Value.ToString().Trim();
                        if (mainValue.Length == 0) continue;

                        bool canMergeRow = true;

                        for (int i = 1; i < lstRowColumn.Count ; i++)
                        {
                            //Ϊ��,���߲��ɼ�,����Ϊ���ַ���
                            if (this.Rows[e.RowIndex].Cells[lstRowColumn[i]].Value == null ||
                                !this.Columns[lstRowColumn[i]].Visible ||                                
                                this.Rows[e.RowIndex].Cells[lstRowColumn[i]].Value.ToString() == mainValue )
                            {
                                continue;
                            }
                            canMergeRow = false;
                            break;
                        }
                        if (canMergeRow)
                        {
                            //�Ա���ɫ���
                            e.Graphics.FillRectangle(backBrush, e.CellBounds);
                            PaintingFont(e,this.Rows[e.RowIndex].Cells[lstRowColumn[0]].Value.ToString().Trim(), 0, 0, leftColumns , rightColumns );
                            if(!stockOfAllRows.ContainsKey (e.RowIndex))
                            {
                                stockOfAllRows.Add(e.RowIndex, new Dictionary<string, List<string>>());
                            }
                            if (stockOfAllRows[e.RowIndex].ContainsKey(lstRowColumn[0]))
                            {
                                stockOfAllRows[e.RowIndex].Remove(lstRowColumn[0]);
                            }
                            stockOfAllRows[e.RowIndex].Add(lstRowColumn[0], lstRowColumn);
                            // ���ұ���
                            if (rightColumns  == 0)
                            {
                                e.Graphics.DrawLine(gridLinePen, e.CellBounds.Right - 1, e.CellBounds.Top, e.CellBounds.Right - 1, e.CellBounds.Bottom);
                            }

                            e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);
                            e.Handled = true;
                            mergeRow = true;
                        }

                    }
                    #endregion
                }
            }
            #endregion

            #region ����ϲ�
            if (!mergeRow && this._mergecolumnname.Contains(this.Columns[e.ColumnIndex].Name) 
                && e.RowIndex != -1 && e.Value != null && !string.IsNullOrEmpty (e.Value.ToString()))
            {
                string curValue = e.Value.ToString();
                
                if (!string.IsNullOrEmpty(curValue))
                {
                    #region ��ȡ���������
                    for (int i = e.RowIndex + 1; i < this.Rows.Count; i++)
                    {
                        if (checkHasTheColumn(Columns[e.ColumnIndex].Name, i).Length > 0) break;
                        if (this.Rows[i].Cells[e.ColumnIndex].Value != null && this.Rows[i].Cells[e.ColumnIndex].Value.ToString().Equals(curValue))
                        {
                            downRows++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    #endregion
                    #region ��ȡ���������
                    for (int i = e.RowIndex - 1; i >= 0; i--)
                    {
                        if (checkHasTheColumn(Columns[e.ColumnIndex].Name, i).Length >0) break;
                        if (this.Rows[i].Cells[e.ColumnIndex].Value != null && this.Rows[i].Cells[e.ColumnIndex].Value.ToString().Equals(curValue))
                        { 
                            upRows++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    #endregion
                    if (downRows == 0 && upRows == 0) return;
                   
                }
                if (this.Rows[e.RowIndex].Selected)
                {
                    backBrush.Color = e.CellStyle.SelectionBackColor;
                    fontBrush.Color = e.CellStyle.SelectionForeColor;
                }
                //�Ա���ɫ���
                e.Graphics.FillRectangle(backBrush, e.CellBounds);
                //���ַ���
                PaintingFont(e, curValue,upRows, downRows, 0, 0);
                if (downRows == 0)
                {
                    e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);
                }
                // ���ұ���
                e.Graphics.DrawLine(gridLinePen, e.CellBounds.Right - 1, e.CellBounds.Top, e.CellBounds.Right - 1, e.CellBounds.Bottom);

                e.Handled = true;
            }
            #endregion
        }

        internal string checkHasTheColumn(string columnName, int rowIndex)
        {
            if (stockOfAllRows.ContainsKey(rowIndex))
            {
                foreach (string sTemp in stockOfAllRows[rowIndex].Keys)
                {
                    if (stockOfAllRows[rowIndex][sTemp].Contains(columnName))
                    {
                        return sTemp;
                    }
                }
            }
            return "";
        }
        internal int checkHasMergeTheColumn(string columnName, int rowIndex,out string value)
        {
            value = "";
            if (stockOfAllRows.ContainsKey(rowIndex))
            {
                foreach (string sTemp in stockOfAllRows[rowIndex].Keys)
                {
                    if (stockOfAllRows[rowIndex][sTemp].Contains(columnName))
                    {                        
                        value = Rows[rowIndex].Cells[sTemp].Value.ToString();
                        return stockOfAllRows[rowIndex][sTemp].Count;
                    }
                }
            }
            return 0;
        }

        private void PaintingFont(System.Windows.Forms.DataGridViewCellPaintingEventArgs e, string showValue, int UpRowsIn, int DownRowsIn, int leftColumnsIn, int rightColumnsIn)
        {
            int UpRows = UpRowsIn;
            int DownRows = DownRowsIn;
            int leftColumns = leftColumnsIn;
            int rightColumns = rightColumnsIn;
            int cellwidth = e.CellBounds.Width;
            int cellheight = e.CellBounds.Height;
            int fondHalf = (int)(e.CellStyle.Font.Size);
            int sumHeight = 0;
            for (int i = e.RowIndex - UpRows; i <= e.RowIndex + DownRows; i++)
            {
                //sumHeight += this.Rows[i].Displayed ? this.Rows[i].Height : 0;
                sumHeight += this.Rows[i].Height;
            }
            int sumWidth = 0;
            if (leftColumns > 0 || rightColumns > 0)
            {
                for (int i = displayedColumns.IndexOf(Columns[e.ColumnIndex].Name) - leftColumns;
                    i <= displayedColumns.IndexOf(Columns[e.ColumnIndex].Name) + rightColumns; i++)
                {
                    sumWidth += this.Columns[displayedColumns[i]].Visible ? this.Columns[displayedColumns[i]].Width : 0;
                }
            }
            else sumWidth = Columns[e.ColumnIndex].Width;

            SolidBrush fontBrush = new SolidBrush(e.CellStyle.ForeColor);
            int fontheight = (int)e.Graphics.MeasureString(showValue, e.CellStyle.Font).Height + 1;
            int fontwidth = (int)e.Graphics.MeasureString(showValue, e.CellStyle.Font).Width + 1;

            int wordWidth = fontwidth > (sumWidth) ? (sumWidth) : fontwidth;

            int wordHeight = ((fontwidth - 1) / (sumWidth - fondHalf) + 1) * fontheight;
            //�������һ�е������,�ֵĸ߶�����Ϊ����ĸ߶�
            if (fontwidth <= sumWidth) wordHeight = fontheight;
            //����ֵĸ߶ȴ����ܸ߶�,ֻ�ð��ܸ߶ȸ�ֵ���ֵĸ߶�
            else if (wordHeight > sumHeight) wordHeight = sumHeight;

            int x0 = e.CellBounds.X;
            for (int i = displayedColumns.IndexOf(Columns[e.ColumnIndex].Name) - 1;
                i >= displayedColumns.IndexOf(Columns[e.ColumnIndex].Name) - leftColumns; i--)
            {
                if (this.Columns[displayedColumns[i]].Visible)
                    x0 -= this.Columns[displayedColumns[i]].Width;
            }
            int y0 = e.CellBounds.Y;
            for (int i = e.RowIndex - 1; i >= e.RowIndex - UpRows; i--)
            {
                // if (this.Rows[i].Displayed)
                y0 -= this.Rows[i].Height;
            }
            Rectangle drawRec;
            int recX, recY;

            StringFormat theStringFormat = StringFormat.GenericTypographic;

            if (e.CellStyle.Alignment == DataGridViewContentAlignment.BottomCenter ||
                e.CellStyle.Alignment == DataGridViewContentAlignment.BottomLeft ||
                e.CellStyle.Alignment == DataGridViewContentAlignment.BottomRight)
            {
                recY = y0 + sumHeight - wordHeight;
            }
            else if (e.CellStyle.Alignment == DataGridViewContentAlignment.TopCenter ||
            e.CellStyle.Alignment == DataGridViewContentAlignment.TopLeft ||
            e.CellStyle.Alignment == DataGridViewContentAlignment.TopRight)
            {
                recY = y0;
            }
            else
            {
                recY = y0 + (sumHeight - wordHeight) / 2;
            }
            if (e.CellStyle.Alignment == DataGridViewContentAlignment.BottomLeft ||
              e.CellStyle.Alignment == DataGridViewContentAlignment.TopLeft ||
              e.CellStyle.Alignment == DataGridViewContentAlignment.MiddleLeft)
            {
                recX = x0;
                theStringFormat.Alignment = StringAlignment.Near;
            }
            else if (e.CellStyle.Alignment == DataGridViewContentAlignment.BottomRight ||
              e.CellStyle.Alignment == DataGridViewContentAlignment.TopRight ||
              e.CellStyle.Alignment == DataGridViewContentAlignment.MiddleRight)
            {
                recX = x0 + sumWidth - wordWidth;
                theStringFormat.Alignment = StringAlignment.Far;
            }
            else
            {
                recX = x0 + (sumWidth - wordWidth) / 2;
                theStringFormat.Alignment = StringAlignment.Center;
            }
            Rectangle r = this.GetCellDisplayRectangle(FirstDisplayedCell.ColumnIndex, FirstDisplayedCell.RowIndex, true);
            if (recY >= r.Y)
            {
                drawRec = new Rectangle(recX, recY, wordWidth, wordHeight);
                e.Graphics.DrawString(showValue, e.CellStyle.Font, fontBrush, drawRec, theStringFormat);
            }
        }
        internal void reCalculateAllStock()
        {
            stockOfAllRows.Clear();
            foreach (List<string> lstRowColumn in _lstMergeRowColumn)
            {
                for (int rowNo = 0; rowNo < Rows.Count; rowNo++)
                {
                    if (null != this.Rows[rowNo].Cells[lstRowColumn[0]] && rowNo!=this.NewRowIndex)
                    {
                        string mainValue = this.Rows[rowNo].Cells[lstRowColumn[0]].Value.ToString().Trim();
                        if (mainValue.Length == 0) continue;

                        //���Ժϲ�
                        bool canMergeRow = true;
                        //���в��ü����,��Ϊ�����ϵ�ĳ�����Ѿ�����ĺϲ���
                        bool nowNeedToCheckThisGroup = false;

                        for (int i = 1; i < lstRowColumn.Count; i++)
                        {
                            //�ںϲ��������ҵ��˵�ǰ�к���
                            if (checkHasTheColumn(lstRowColumn[0], rowNo).Length > 0)
                            {
                                nowNeedToCheckThisGroup = true;
                                break;
                            }
                            //Ϊ��,���߲��ɼ�,����Ϊ���ַ���
                            if (this.Rows[rowNo].Cells[lstRowColumn[i]].Value == null ||
                                !this.Columns[lstRowColumn[i]].Visible ||
                                this.Rows[rowNo].Cells[lstRowColumn[i]].Value.ToString() == mainValue)
                            {
                                continue;
                            }
                            canMergeRow = false;
                            break;
                        }
                        if (nowNeedToCheckThisGroup) continue;
                        if (canMergeRow)
                        {
                            if (!stockOfAllRows.ContainsKey(rowNo))
                            {
                                stockOfAllRows.Add(rowNo, new Dictionary<string, List<string>>());
                            }
                            if (stockOfAllRows[rowNo].ContainsKey(lstRowColumn[0]))
                            {
                                stockOfAllRows[rowNo].Remove(lstRowColumn[0]);
                            }
                            stockOfAllRows[rowNo].Add(lstRowColumn[0], lstRowColumn);
                        }
                    }
                }
            }
        }

        #region ��ǿ����ϲ�����

        //��¼���еĺ���ϲ������;
        private Dictionary<int, Dictionary<string, List<string>>> stockOfAllRows = new Dictionary<int, Dictionary<string, List<string>>>();

        private List<string> displayedColumns = new List<string>();
        /// <summary>
        /// ����������ʾ��
        /// </summary>
        private void resetColumnDisplayArray()
        {
            displayedColumns.Clear();
            for (int i = 0; i < Columns.Count; i++)
            {
                if (Columns[i].Visible)
                {
                    bool find = false;
                    for (int r = displayedColumns.Count - 1; r >= 0; r--)
                    {
                        if (Columns[displayedColumns[r]].DisplayIndex < Columns[i].DisplayIndex)
                        {
                            find = true;
                            displayedColumns.Insert(r + 1, Columns[i].Name);
                            break;
                        }
                    }
                    if (!find)
                    {
                        displayedColumns.Insert(0, Columns[i].Name);
                    }
                }
            }
        }
        #endregion 
        /// <summary>
        /// �ͷ���Դ
        /// </summary>
        /// <param name="o"></param>
        private void NAR(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch { }
            finally
            {
                o = null;
            }
        }
        #endregion

        #region override functions
        protected override void OnKeyUp(KeyEventArgs e)
        {
            base.OnKeyUp(e);
            string searchCol = "";  //Ҫ�������ֶ������ö������ӵ��ַ�����
            if (e.KeyData == Keys.F3)
            {
                //��������ʾ�������ö��ŷָ����ַ���
                foreach (DataGridViewColumn dgvColumn in this.Columns)
                {
                    if (dgvColumn.Visible == true)
                    {
                        searchCol += dgvColumn.Name + ",";
                    }
                }

                //ȥ�����һ������
                searchCol = searchCol.Substring(0, searchCol.Length - 1);

                if (searchCol == "")
                {
                    MessageBox.Show("�Ҳ����������������ݣ�", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    this.Search(searchCol, searchValue, currentRow);
                }
            }
        }


        protected override void OnColumnDisplayIndexChanged(DataGridViewColumnEventArgs e)
        {
            base.OnColumnDisplayIndexChanged(e);
            resetColumnDisplayArray();
        }

        /// <summary>
        /// ������ϳ�ˢ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected override void OnSorted(EventArgs e)
        {
            base.OnSorted(e);
            refreshHeader();
        }

        /// <summary>
        /// ̧�����ʱ��ˢ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            this.Invalidate();
            refreshHeader();
        }
        protected override void OnDataSourceChanged(EventArgs e)
        {
            base.OnDataSourceChanged(e);
            loadedFinish = false;
        }

        //���α�ͷ�����п�ʱ�Զ�ˢ�¿ؼ�
        protected override void OnColumnWidthChanged(DataGridViewColumnEventArgs e)
        {
            base.OnColumnWidthChanged(e);
            if (!loadColWidthFromDb && this._ColHeaderTreeView != null) refreshHeader();
        }

        //ʹ��ˮƽ������ʱˢ�¿ؼ�
        protected override void OnScroll(ScrollEventArgs e)
        {
            base.OnScroll(e);
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
            {
                refreshHeader();
                refresh(e);
            }
            
        }

        //��ʾ�к�
        protected override void OnRowPostPaint(DataGridViewRowPostPaintEventArgs e)
        {
            base.OnRowPostPaint(e);
            if (showRowNumber)
            {
                int rownum = (e.RowIndex + 1);
                Rectangle rct = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, this.RowHeadersWidth - 4, e.RowBounds.Height);
                TextRenderer.DrawText(e.Graphics, rownum.ToString(), this.RowHeadersDefaultCellStyle.Font, rct, this.RowHeadersDefaultCellStyle.ForeColor,
                    this.RowHeadersDefaultCellStyle.BackColor, TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
            }

        }
        /// <summary>
        /// ��Ԫ�����(��д)
        /// </summary>
        /// <param name="e"></param>
        /// <remarks></remarks>
        protected override void OnCellPainting(DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex > -1)
            {
                DrawCell(e);
            }
            //�б��ⲻ��д
            if (e.ColumnIndex < 0)
            {
                base.OnCellPainting(e);
                return;
            }

            if (_columnDeep == 1 && _ColHeaderTreeView == null && _ColHeaderTreeView.Nodes.Count==0)
            {
                base.OnCellPainting(e);
                return;
            }

            //���Ʊ�ͷ
            if (e.RowIndex == -1 && _ColHeaderTreeView != null && _ColHeaderTreeView.Nodes.Count > 0)
            {
                PaintUnitHeader((TreeNode)NadirColumnList[e.ColumnIndex], e, _columnDeep);
                e.Handled = true;
                return;
            }
            //��ҳ�Ż�����Ϣ
            this.AllowUserToAddRows = true;
            if (e.RowIndex == this.NewRowIndex && e.ColumnIndex > -1 && computeColumns.Count>0)
            {
                DrawFooter(e);
            }

        }

        ///TODO:����ҳ����Ϣ
        //����footer�Ļ�����Ϣ
        private void DrawFooter(DataGridViewCellPaintingEventArgs e)
        {
            // �˴�ֻ��ҳ����Ϣ
            this.AllowUserToAddRows = true;
            if (e.RowIndex == -1) return;
            
            ///����������s����0���Ӳ���������xȴ��ֵ����֪��lambda���ʽ����ʲô�����ˡ�ֻ��ѭ����
            //var drs = from DataGridViewRow a in this.Rows
            //          where e.RowIndex < this.NewRowIndex
            //          select a;
            //int ss = drs.Sum(x => int.Parse(x.Cells["WEEKDAY"].Value.ToString()));
            
            e.PaintBackground(e.CellBounds, false);
            Dictionary<string, string> dict= dictComputeColumns();
            for (int i = 0; i < this.NewRowIndex; i++)
            {
                foreach (string col in dict.Keys )
                {
                    if (this.Columns[e.ColumnIndex].Name == col)
                    {
                        StringFormat sf = new StringFormat();
                        if (this.Columns[col].DefaultCellStyle.Alignment.ToString().Contains("Right"))
                            sf.Alignment = StringAlignment.Far;
                        else if (this.Columns[col].DefaultCellStyle.Alignment.ToString().Contains("Center"))
                            sf.Alignment = StringAlignment.Center;
                        else
                            sf.Alignment = StringAlignment.Near;
                        RectangleF rf = new RectangleF(e.CellBounds.Left + 2, e.CellBounds.Top + 4, e.CellBounds.Width-4, e.CellBounds.Height-4);
                        e.Graphics.DrawString(dict[col], this.Font, Brushes.Black,rf,sf);
                    }
                }
            }
            e.Handled = true;
               
        }

        /// <summary>
        /// ��������м����е���ֵ�ԣ������ء�
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, string> dictComputeColumns()
        {
            DataTable dt = GridToTable();
            Dictionary<string, string> dict = new Dictionary<string, string>();
            int i = 0; //����˳��ֵд���ֵ�
            foreach (DataColumn col in dt.Columns)
            {
                string val=computeColumns[i].Substring(computeColumns[i].IndexOf(",") + 1);
                if (computeColumns[i].Contains('(')) //ʹ���˾ۺϺ���
                {
                    string v;
                    if (val.Split(',').Length > 1) //�й�������
                    {
                        v = (dt.Compute(val[0].ToString(), val[1].ToString())).ToString();
                    }
                    else
                    {
                        v = (dt.Compute(val, "")).ToString();
                    }
                    if (beforeShow != null)
                        v = beforeShow(col.ColumnName,v);
                    dict.Add(col.ColumnName, v);
                    
                }
                else //����Ϊ�ϼƵ��ַ��������ݣ�û��()���������ж�
                {
                    dict.Add(col.ColumnName, val);
                }
                i++;
            }
            return dict;
        }



        /// <summary>
        /// ת��Ϊһ��DataTable
        /// </summary>
        /// <typeparam name="TResult"></typeparam>
        ///// <param name="value"></param>
        /// <returns></returns>
        private DataTable GridToTable()
        {
            DataTable dt;
            List<string> cols = oColumns();
            if (this.DataSource.GetType().Equals(typeof(System.Data.DataTable)))
            {
                dt = ((DataTable)this.DataSource).Copy();
                for (int i = dt.Columns.Count - 1; i >= 0; i--)
                {
                    if (cols.IndexOf(dt.Columns[i].ColumnName) == -1)
                    {
                        dt.Columns.RemoveAt(i);
                    }
                }
            }
            else
            {
                
                dt = new DataTable();
                foreach (string s in cols)
                {
                    DataColumn col = new DataColumn(s, this.Columns[s].ValueType);
                }
                for(int i=0;i<this.NewRowIndex;i++)
                {
                    foreach (string c in cols)
                    {
                        DataRow dr = dt.NewRow();
                        dr[c] = this.Rows[i].Cells[c].Value;
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// �������õ�computeColumns���м���
        /// </summary>
        /// <returns>������ʾ���м���</returns>
        private List<string> oColumns()
        {
            List<string> o=new List<string>();
            foreach (string v in computeColumns)
            {
                o.Add(v.Split(',')[0]);
            }
            return o;
        }


        protected override void OnDataBindingComplete(DataGridViewBindingCompleteEventArgs e)
        {
            base.OnDataBindingComplete(e);
            myNodeLevels();
            LoadedFinish = true;
        }
       

        private void refresh(ScrollEventArgs e)
        {
            Rectangle rect;

            Point pt = PointToScreen(this.Location);

            if (pt.X < 0)
            {
                int left = -pt.X;
                int top = this.ColumnHeadersHeight;
                int width = e.OldValue - e.NewValue;
                int height = this.ClientSize.Height;
                this.Invalidate(new Rectangle(new Point(left, top), new Size(width, height)));
            }

            pt.X += this.Width;
            rect = Screen.GetBounds(pt);

            if (pt.X > rect.Right)
            {
                int left = this.ClientSize.Width - (pt.X - rect.Right) - (e.NewValue - e.OldValue);
                int top = this.ColumnHeadersHeight;
                int width = e.NewValue - e.OldValue;
                int height = this.ClientSize.Height;
                this.Invalidate(new Rectangle(new Point(left, top), new Size(width, height)));
            }

            pt.Y += this.Height;
            if (pt.Y > rect.Bottom)
            {
                int left = 0;
                int top = this.ColumnHeadersHeight;
                int width = this.ClientSize.Width;
                int height = this.ClientSize.Height - (pt.Y - rect.Bottom) - (e.NewValue - e.OldValue);
                this.Invalidate(new Rectangle(new Point(left, top), new Size(width, height)));
            }
        }

        private void refreshHeader()
        {
            this.Invalidate(new Rectangle(this.Location, new Size(this.Width, this.ColumnHeadersHeight)));
        }

        #endregion --------------end override functions---------------

        #region ---------����ͷ����-----------------------------
        public DataGridViewEx(IContainer container)
        {
            container.Add(this);
            InitializeComponent();
        }

        #region ����ͷ���趨
        private TreeView _ColHeaderTreeView = new TreeView();


        /// <summary>
        /// ��ά�б�������ṹ
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        [Description("��ά�б�������ṹ,ʹ�ñ�����ʱ,����Ҫ����ʾ�������οؼ�Ҷ�ڵ�����һ��,������ʾ���쳣,���ĳ���ڵ㲻��ʾ,����Tag��Ϊhide"), Category("�Զ�������")]
        public TreeNodeCollection HeadSource
        {
            get 
            {
                iNodeLevels = 0;
                ColLists.Clear();
                allColLists.Clear();
                this.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                myNodeLevels();
                setColumns();
                return this._ColHeaderTreeView.Nodes; 
            }
        }

        private void setColumns()
        {
            if (DesignMode && _ColHeaderTreeView.Nodes.Count>0)
            {
                this.Columns.Clear();
                foreach (TreeNode tn in allColLists)
                {
                    DataGridViewColumn col = new DataGridViewTextBoxColumn();
                    col.Name = tn.Name;
                    col.HeaderText = tn.Text;
                    if (tn.Tag != null && tn.Tag.ToString() == "hide")
                    {
                        col.Visible = false;
                        col.Width = 0;
                    }
                     this.Columns.Add(col);
                }
            }
        }


        private int _cellHeight = 22;
        private int _columnDeep = 1;
        [Description("���û��úϲ���ͷ�������")]
        public int ColumnDeep
        {
            get
            {
                if (this.Columns.Count == 0)
                    _columnDeep = 1;
                this.ColumnHeadersHeight = _cellHeight * _columnDeep;
                return _columnDeep;
            }

            set
            {
                if (value < 1)
                    _columnDeep = 1;
                else
                    _columnDeep = value;
                this.ColumnHeadersHeight = _cellHeight * _columnDeep;
            }
        }
        
        ///<summary>
        ///���ƺϲ���ͷ
        ///</summary>
        ///<param name="node">�ϲ���ͷ�ڵ�</param>
        ///<param name="e">��ͼ������</param>
        ///<param name="level">������</param>
        ///<remarks></remarks>
        public void PaintUnitHeader(TreeNode node, DataGridViewCellPaintingEventArgs e, int level)
        {
            //���ڵ�ʱ�˳��ݹ����
            if (level == 0 ) //|| ( node.Tag!=null && node.Tag.ToString()=="hide"))
                return;

            RectangleF uhRectangle;
            int uhWidth;
            SolidBrush gridBrush = new SolidBrush(this.GridColor);

            Pen gridLinePen = new Pen(gridBrush);
            StringFormat textFormat = new StringFormat();


            textFormat.Alignment = StringAlignment.Center;

            
            uhWidth = GetUnitHeaderWidth(node);

            //��ԭ���㷨�����������⡣
            if (node.Nodes.Count == 0)
            {
                uhRectangle = new Rectangle(e.CellBounds.Left,
                            e.CellBounds.Top + node.Level * _cellHeight,
                            uhWidth - 1,
                            _cellHeight * (_columnDeep - node.Level) - 1);
            }
            else
            {
                uhRectangle = new Rectangle(
                            e.CellBounds.Left,
                            e.CellBounds.Top + node.Level * _cellHeight,
                            uhWidth - 1,
                            _cellHeight - 1);
            }

            Color backColor = e.CellStyle.BackColor;
            if (node.BackColor != Color.Empty)
            {
                backColor = node.BackColor;
            }
            SolidBrush backColorBrush = new SolidBrush(backColor);
            //������
            e.Graphics.FillRectangle(backColorBrush, uhRectangle);

            //������
            e.Graphics.DrawLine(gridLinePen
                                , uhRectangle.Left
                                , uhRectangle.Bottom
                                , uhRectangle.Right
                                , uhRectangle.Bottom);
            //���Ҷ���
            e.Graphics.DrawLine(gridLinePen
                                , uhRectangle.Right
                                , uhRectangle.Top
                                , uhRectangle.Right
                                , uhRectangle.Bottom);

            ////д�ֶ��ı�
            Color foreColor = Color.Black;
            if (node.ForeColor != Color.Empty)
            {
                foreColor = node.ForeColor;
            }
            float x = uhRectangle.Left + uhRectangle.Width / 2 - e.Graphics.MeasureString(node.Text, this.Font).Width / 2 - 1;
            x = x > uhRectangle.Left ? x : uhRectangle.Left;
            e.Graphics.DrawString(node.Text, this.Font
                                    , new SolidBrush(foreColor)
                                    , x
                                    , uhRectangle.Top + uhRectangle.Height / 2 - e.Graphics.MeasureString(node.Text, this.Font).Height / 2);

            ////�ݹ����()
            if (node.PrevNode == null)
                if (node.Parent != null)
                    PaintUnitHeader(node.Parent, e, level - 1);
        }

        /// <summary>
        /// ��úϲ������ֶεĿ��
        /// </summary>
        /// <param name="node">�ֶνڵ�</param>
        /// <returns>�ֶο��</returns>
        private int GetUnitHeaderWidth(TreeNode node)
        {
            int uhWidth = 0;
            //�����ײ��ֶεĿ��
            if (node.Nodes == null)
                return this.Columns[GetColumnListNodeIndex(node)].Width;

            if (node.Nodes.Count == 0)
                return this.Columns[GetColumnListNodeIndex(node)].Width;

            //��÷���ײ��ֶεĿ��
            for (int i = 0; i <= node.Nodes.Count - 1; i++)
            {
                if (node.Nodes[i].Tag == null || node.Nodes[i].Tag.ToString() != "hide")
                {
                    uhWidth = uhWidth + GetUnitHeaderWidth(node.Nodes[i]);
                }
            }
            return uhWidth;
        }


        /// <summary>
        /// ��õײ��ֶ�����
        /// </summary>
        /// <param name="node">�ײ��ֶνڵ�</param>
        /// <returns>����</returns>
        /// <remarks></remarks>
        private int GetColumnListNodeIndex(TreeNode node)
        {
            for (int i = 0; i <= _columnList.Count - 1; i++)
            {
                if (((TreeNode)_columnList[i]).Equals(node))
                    return i;
            }
            return -1;
        }

        private List<TreeNode> _columnList = new List<TreeNode>();
        [Description("��ײ�ڵ㼯��")]
        public List<TreeNode> NadirColumnList
        {
            get
            {
                if (this._ColHeaderTreeView == null)
                    return null;

                if (this._ColHeaderTreeView.Nodes == null)
                    return null;

                if (this._ColHeaderTreeView.Nodes.Count == 0)
                    return null;

                _columnList.Clear();
                foreach (TreeNode node in this._ColHeaderTreeView.Nodes)
                {
                    GetNadirColumnNodes(_columnList, node);
                }
                return _columnList;
            }
        }

        private void GetNadirColumnNodes(List<TreeNode> alList, TreeNode node)
        {
            if (node.FirstNode == null) //&& (node.Tag == null || node.Tag.ToString() != "hide"))
            {
                alList.Add(node);
            }
            foreach (TreeNode n in node.Nodes)
            {
                GetNadirColumnNodes(alList, n);
            }
        }

        /// <summary>
        /// ��õײ��ֶμ���
        /// </summary>
        /// <param name="alList">�ײ��ֶμ���</param>
        /// <param name="node">�ֶνڵ�</param>
        /// <param name="checked">�����������</param>
        /// <remarks></remarks>
        private void GetNadirColumnNodes(List<TreeNode> alList, TreeNode node, Boolean isChecked)
        {
            if (isChecked == false)
            {
                if (node.FirstNode == null)
                {
                    alList.Add(node);
                    if (node.NextNode != null)
                    {
                        GetNadirColumnNodes(alList, node.NextNode, false);
                        return;
                    }
                    if (node.Parent != null)
                    {
                        GetNadirColumnNodes(alList, node.Parent, true);
                        return;
                    }
                }
                else
                {
                    if (node.FirstNode != null)
                    {
                        GetNadirColumnNodes(alList, node.FirstNode, false);
                        return;
                    }
                }
            }
            else
            {
                if (node.FirstNode == null)
                {
                    return;
                }
                else
                {
                    if (node.NextNode != null)
                    {
                        GetNadirColumnNodes(alList, node.NextNode, false);
                        return;
                    }

                    if (node.Parent != null)
                    {
                        GetNadirColumnNodes(alList, node.Parent, true);
                        return;
                    }
                }
            }
        }



        
        #endregion

        #endregion

    }
}