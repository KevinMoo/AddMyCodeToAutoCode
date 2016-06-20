/****************************************************************************************************
 * Copyright (C) 2010 大连陆海科技股份有限公司 版权没有，任意拷贝及使用，但对使用造成的任何后果不负任何责任，互相开源影响，共同进步
 * 文 件 名：DataGridViewEx.cs
 * 创 建 人：明振居士
 * Email:nzj.163@163.com   qq:342155124
 * 创建时间：2010-06-01
 * 最后修改时间：2012-1-19  增加第10条所示的功能；修改了列头超过26列的错误，导出excel为数组方式，速度更快，导出的单元格设置为文本格式。
 * 标    题：用户自定义的DataGridView控件
 * 功能描述：扩展DataGridView控件功能
 * 扩展功能：
 * 1、搜索Search(); 有两个同明方法，参数不同 F3为快捷键继续向下搜索
 * 2、用TreeView HeadSource 来设置复杂的标题样式，如果某个节点对应的显示列隐藏，请将该节点Tag设置为hide，隐藏列的排列位置与绑定数据元列位置对应，树叶节点的顺序需要与结果集的列顺序一致
 * 3、通过反射导出Excel，无需引用com组件，方法ExportExcel() ，不受列数的限制，表头同样可以导出,AutoFit属性设置导出excel后是否自动调整单元格宽度
 *    导出内容支持自定义的：Title List<string> Header   List<string> Footer,支持在设计时值的设定，窗口关闭时Excel资源自动彻底释放
 * 4、可以自己任意设定那些列显示及不显示，通过调用方法SetColumnVisible()实现。
 * 5、设置列标题SetHeader(),设置列永远可见AlwaysShowCols(),设置列暂时不可见HideCols()
 *    注意，当使用了TreeView作为复杂Header时，不要使用本方法，Header显示的内容根据treeview内容而显示
 * 6、列宽度及顺序的保存SaveGridView()，加载LoadGridView()
 * 7、支持所见即所得的打印功能，举例如下
 *     private void button5_Click(object sender, EventArgs e)
        {
            DGVPrinter printer = new DGVPrinter();
            printer.PrintPreviewDataGridView(DataGridViewEx1);
        }
 * 8、自定义合并行与列，行合并用 MergeRowColumn 属性，列合并用MergeColumnNames属性，都可以定义多个列
 * 9、行标号的设置 bool ShowRowNumber;
 * 10、增加最后一行的汇总行，支持列的聚合函数，参见http://msdn.microsoft.com/zh-cn/library/system.data.datacolumn.expression(v=VS.100).aspx
 *     假设对id列显示“合计”字符，avgPrice进行平均值，total列显示合计，则对ComputeColumns增加三行内容：id,合计：；avgPrice,Avg(avgPrice)；total,Sum(total)
 *     如果需要对值进行格式控制，请实现beforeShow事件
 *     增加了导出和打印对应的支持，所见即所得的对齐方应用于式导出及打印。
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
        /// 定义事件，用于在显示计算列的格式控制
        /// </summary>
        /// <param name="sender">列名</param>
        /// <param name="value">列的值</param>
        /// <returns>返回到加工后的值</returns>
        public delegate string BeforeShow(string sender,string value);
        public event BeforeShow beforeShow;

        #region private paras ---------------------------------------
        //---------------------------------------search paras------------------------------------------------------
        private string searchValue = "";  // 要搜索的字符串
        private int currentRow=0;       // 当前搜索的行号（用于连续搜索时）
        //----------------------------------------复杂表头及合并列-------------------------------------------------
        private List<string> _mergecolumnname = new List<string>();
        private bool showRowNumber=false;   //是否显示行号
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
        private int iNodeLevels;                                    // 树的最大层数
        private int iCellHeight=22;                                 // 一维列表标题的高度
        private IList<TreeNode> ColLists = new List<TreeNode>();    // 所有显示的页节点
        private IList<TreeNode> allColLists = new List<TreeNode>();    // 所有的页节点

        private string title;
        private List<string> header;
        private List<string> footer;
        private bool autoFit=true;
        /// <summary>
        /// 从持久化加载列宽
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
        /// 是否加载完毕数据，如果加载完毕    
        /// 当使用merge的时候，如果用户修改数据，应该重新刷对象，让此刷新显示效果。
        /// 为了减少刷新次数，加载了此属性
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
        /// 设置或获取合并列的集合
        /// </summary>
        [MergableProperty(false)]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        //[DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
        [Localizable(true)]
        [Description("设置或获取合并列的集合"), Browsable(true), Category("自定义属性")]
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
        /// 报表头
        /// </summary>
        [Bindable(true), Category("自定义属性"), Description("报表的title")]
        public string Title
        {
            get { return title; }
            set { title = value; }
        }

        /// <summary>
        /// 表头需要写入的内容
        /// </summary>
        [Bindable(true), Category("自定义属性"), Description("报表的header需要写入的内容")]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> Header
        {
            get { return header; }
            set { header = value; }
        }

        /// <summary>
        /// 页脚内容
        /// </summary>
        [Bindable(true), Category("自定义属性"), Description("报表的footer需要写入的内容")]
        //[DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]  只有去除这句话才可以在设计时对其赋值
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public List<string> Footer
        {
            get { return footer; }
            set { footer = value; }
        }

        /// <summary>
        /// 导出数据是否自动设置行列的宽度
        /// </summary>
        [Description("导出数据是否自动设置行列的宽度"), Category("自定义属性")]
        public bool AutoFit
        {
            get { return autoFit; }
            set { autoFit = value; }
        }


        /// <summary>
        /// 是否显示行号
        /// </summary>
        [Description("设定是否显示行号"), Category("自定义属性")]
        public bool ShowRowNumber
        {
            get { return showRowNumber; }
            set { showRowNumber = value; }
        }


        /// <summary>
        /// 设置或获取合并列的集合
        /// </summary>
        [MergableProperty(false)]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
        [Localizable(true)]
        [Description("设置或获取行合并的列名称,列名用逗号分隔，可以设置多个不同的合并列集合"), Browsable(true), Category("自定义属性")]
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
        /// 设置或获取计算列的集合
        /// </summary>
        [MergableProperty(false)]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Visible)]
        [Localizable(true)]
        [Description("设置或获取计算列组,每个计算列一行,如对id进行sum计算,就设置为:id,sum(id)"), Browsable(true), Category("自定义属性")]
        public List<string> ComputeColumns
        {
            get{return computeColumns;}
            set { computeColumns = value; }
        }
        #endregion

        #region -------------------public functions------------------
        /// <summary>
        /// 当前网格的信息搜索，忽略大小写
        /// 默认为搜索所有可见列，每次提示输入搜索内容
        /// </summary>
        public void Search()
        {
            string searchCol = "";  //要搜索的字段名（用逗号连接的字符串）

            //把所有显示的列名用逗号分隔成字符串
            foreach (DataGridViewColumn dgvColumn in this.Columns)
            {
                if (dgvColumn.Visible == true)
                {
                    searchCol += dgvColumn.Name + ",";
                }
            }

            if (searchCol == "")
            {
                MessageBox.Show("没有可见的搜索列！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                InputBox frm = new InputBox();
                DialogResult dlrResult = frm.ShowDialog();
                this.searchValue = frm.SearchValue;

                if (dlrResult == DialogResult.OK)
                {
                    //去掉最后一个逗号
                    searchCol = searchCol.Substring(0, searchCol.Length - 1);
                    this.Search(searchCol, searchValue, currentRow);
                }
            }
        }

        /// <summary>
        /// 当前网格的信息搜索，忽略大小写，指定搜索的列，内容及起始行
        /// </summary>
        /// <param name="searchCol">要搜索的列（用逗号分隔的字符串）</param>
        /// <param name="searchValue">搜索的值</param>
        /// <param name="startRow">搜索的起始行</param>
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
                    MessageBox.Show("找不到正在搜索的数据！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    this.CurrentCell = this[searchCols[foundCol], currentRow - 1];
                }

                this.Focus();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "搜索出错", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>
        /// 将所见的内容导出为excel
        /// </summary>
        public void ExportExcel()
        {
            Cursor.Current = Cursors.WaitCursor;
            Export2Excel(this, true,true); //优化的excel导出，速度加快
            Cursor.Current = Cursors.Default;
        }

        /// <summary>
        /// 设定当前的列哪些显示，对于隐藏并不需要设定的列，设定它的tag＝0
        /// TreeView的标题头不能进行此设定
        /// </summary>
        public void SetColumnVisible()
        {
            if (_ColHeaderTreeView.Nodes.Count>0) return; //对于复杂结构类型不能自己定义哪些列可以显示
            FrmColumnSet frm = new FrmColumnSet(this);
            frm.ShowDialog();
        }

        /// <summary>
        /// 设置可见列及列标题，显示的顺序与传入的顺序相同,当使用了TreeView作为Header时，不要使用本方法
        /// 经过此方法设置后，设置的列才是可见的，并且与SetColumnVisible()方法同步
        /// 如果想要设置某些列永远可见，在调用此方法后，调用方法AlwaysShowCols进行设置
        /// 如果要使某些列暂时不可见，通过SetColumnVisible()可以设置为可见，则调用HideCols()方法
        /// </summary>
        /// <param name="columns">用逗号分割的列名称</param>
        /// <param name="headers">用逗号分割的列头，需要与columns配对出现</param>
        public void SetHeader(string columns, string headers)
        {
            if (_ColHeaderTreeView != null && _ColHeaderTreeView.Nodes.Count > 0) return;
            string[] cols = columns.Split(',');
            string[] heads = headers.Split(',');
            if (cols.Length != heads.Length)
            {
                throw new Exception("设定的显示列数量与列标题数量不配对，请重新设定！");
            }

            //隐藏所有列
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
        /// 将传入的列设置为永远可见
        /// </summary>
        /// <param name="columns">用逗号分割的列名称</param>
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
        /// 设置某些列暂时不可见,逗号分割的列名
        /// </summary>
        /// <param name="columns">用逗号分割的列名</param>
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
        /// 保存列的设置，按顺序为：列名,显示顺序,列宽,是否可见(True,false字符串)
        /// 自己需要将返回的内容持久化,通畅持久化需要再封装一次，增加用户及使用位置的标识，这样不同用户就可以有独立的配置
        /// </summary>
        /// <returns>返回的列名为dictionary[0]，显示顺序(dictionary[0])[0],列宽(dictionary[0])[1]，是否可见(dictionary[0])[2]</returns>
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
        /// 列宽的加载，从传入的字典内加载用户传入的列信息
        /// 传入的内容应该按照显示顺序由小到大进行排序了
        /// 将持久化的列信息加载，一般应包含当前用户及控件使用位置信息
        /// </summary>
        /// <param name="dict">Dictionary<列名,{显示顺序,列宽,是否可见(True,false字符串)}></param>
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
                            column.Width = Convert.ToInt32(dict[dic][1]) > 20 ? Convert.ToInt32(dict[dic][1]) : 120; //小于20则自动调整为120.
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
                        if (tmpcol.GetType().Name == "DataGridViewTextBoxColumn" && tmpcol.Visible) //不显示的隐藏列初始化为tag＝0 
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
                    //这里重新写入excel的头
                    //多维表头
                    //_ColHeaderTreeView  表头TreeView
                    //iNodeLevels 表头的层数
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
                //左右合并单元格
                mergeRowsCell(startRow);

                //上下合并单元格
                //start row = startRow, end row= startRow+ startRow+datagridview.RowCount
                //calculate column index tobe bombined
                int endRow = startRow + datagridview.RowCount - 1;
                if (_mergecolumnname != null && _mergecolumnname.Count > 0)
                {
                    foreach (string cn in _mergecolumnname)
                    {
                        int col = Array.IndexOf<string>(colName, cn) + 1;
                        if (col > 0) mergeColumnsCell(startRow, endRow, col);//设置的合并列没有显示的不进行合并操作
                    }
                }

                drawLine();

                // 写入footer
                int maxRows = iNodeLevels + 1 + datagridview.RowCount;
                if (iNodeLevels == 0) maxRows++;
                if (footer != null) writeFooter(maxRows);

                //写入header
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
                        if (tmpcol.GetType().Name == "DataGridViewTextBoxColumn" && tmpcol.Visible) //不显示的隐藏列初始化为tag＝0 
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
                    //这里重新写入excel的头
                    //多维表头
                    //_ColHeaderTreeView  表头TreeView
                    //iNodeLevels 表头的层数
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

                //使用2维数组进行粘贴，速度更快
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

                //块写入
                Parameters = new Object[2];
                Parameters[0] = startCell;
                Parameters[1] = endCell;
                objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, Parameters);
                //设为文本格式
                object[] o = new object[1];
                o[0] = "@";
                objRange_Late.GetType().InvokeMember("NumberFormatLocal", BindingFlags.SetProperty, null, objRange_Late,o);
                //写入值
                Parameters = new Object[1];
                Parameters[0] = vals;
                objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);

                //左右合并单元格
                mergeRowsCell(startRow);

                //上下合并单元格
                //start row = startRow, end row= startRow+ startRow+datagridview.RowCount
                //calculate column index tobe bombined
                int endRow = startRow + datagridview.RowCount - 1;
                if (_mergecolumnname != null && _mergecolumnname.Count > 0)
                {
                    foreach (string cn in _mergecolumnname)
                    {
                        int col = Array.IndexOf<string>(colName, cn) + 1;
                        if (col > 0) mergeColumnsCell(startRow, endRow, col);//设置的合并列没有显示的不进行合并操作
                    }
                }

                drawLine();
                

                
                int maxRows = iNodeLevels + 1 + datagridview.RowCount;
                if (iNodeLevels == 0) maxRows++;

                setAlignment(colName, columns,maxRows); //设置导出数据的对齐方式

                // 写入footer
                if (footer != null) writeFooter(maxRows);

                //写入header
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


        //复制tree，将隐藏的节点去除
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
        /// 遍历treeview并写入
        /// </summary>
        /// <param name="rootNodes">treeview 节点集合</param>
        /// <param name="excelRow">写入的行:从1开始</param>
        /// <param name="excelColumn">写入的列:从1开始</param>
        /// <returns>占用的列数</returns>
        private int WriteCell(TreeNodeCollection rootNodes, int excelRow, int excelColumn)
        {
            int cellCount = 1;
            if (rootNodes.Count < 1)
            {
                WriteExcelROW(excelRow - 1, excelColumn, cellCount, treeDepth);//sd为treeview的深度
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
            //列数没有限制
            string point1 = ConvertColumnNum2String(excelColumn) + excelRow.ToString();//单元格起始点 如:A1 
            string point2 = ConvertColumnNum2String(excelColumn) + sd.ToString();//单元格起始点 如:A1 
            if (point1 != point2)
            {
                objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
                objRange_Late.GetType().InvokeMember("MergeCells", BindingFlags.SetProperty, null, objRange_Late, new object[] { true });
            }

        }

        /// <summary>
        /// 报表header的输出
        /// </summary>
        /// <param name="excelRow"></param>
        /// <param name="excelColumn"></param>
        /// <param name="cellCountSum"></param>
        /// <param name="excelValue"></param>
        private void WriteExcel(int excelRow, int excelColumn, int cellCountSum, string excelValue)
        {
            //列数没有限制
            string point1 = ConvertColumnNum2String(excelColumn) + excelRow.ToString();//单元格起始点 如:A1
            string point2 = ConvertColumnNum2String(excelColumn + cellCountSum- 1) + excelRow.ToString();//单元格结束点 如:B4

            //_Range = _Worksheet.get_Range(point1, point2);//获取单元格
            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
            if (cellCountSum > 0)
            {
                //_Range.MergeCells = true; //合并单元格
                objRange_Late.GetType().InvokeMember("MergeCells", BindingFlags.SetProperty, null, objRange_Late, new object[] { true });
            }

            objRange_Late.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });//内容水平居中 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter 
            objRange_Late.GetType().InvokeMember("VerticalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });  //内容垂直居中Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter 

            //_Worksheet.Cells[_Range.Row, _Range.Column] = excelValue;//把内容写入单元格
            Parameters = new Object[1];
            Parameters[0] = excelValue;
            objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);
        }
        /// <summary>
        /// 报表footer的输出
        /// </summary>
        /// <param name="excelRow"></param>
        /// <param name="excelColumn"></param>
        /// <param name="cellCountSum"></param>
        /// <param name="excelValue"></param>
        private void WriteExcels(int excelRow, int excelColumn, int cellCountSum, string excelValue)
        {
            //列数没有限制..
            string point1 = ConvertColumnNum2String(excelColumn) + excelRow.ToString();//单元格起始点 如:A1
            string point2 = ConvertColumnNum2String(excelColumn + cellCountSum- 1).ToString() + excelRow.ToString();//单元格结束点 如:B4


            //_Range = _Worksheet.get_Range(point1, point2);//获取单元格
            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
            if (cellCountSum > 0)
            {
                //_Range.MergeCells = true; //合并单元格
                objRange_Late.GetType().InvokeMember("MergeCells", BindingFlags.SetProperty, null, objRange_Late, new object[] { true });
            }
            //_Worksheet.Cells[_Range.Row, _Range.Column] = excelValue;//把内容写入单元格
            Parameters = new Object[1];
            Parameters[0] = excelValue;
            objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);
            objRange_Late.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4131 });//左对齐 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft 
        }

        /// <summary>
        /// 报表title的输出
        /// </summary>
        /// <param name="excelRow"></param>
        /// <param name="excelColumn"></param>
        /// <param name="cellCountSum"></param>
        /// <param name="excelValue"></param>
        private void WriteExcelTitle(int excelRow, int excelColumn, int cellCountSum, string excelValue)
        {
            string point1 = ConvertColumnNum2String(excelColumn) + excelRow.ToString();//单元格起始点 如:A1
            string point2 = ConvertColumnNum2String(excelColumn + cellCountSum - 1).ToString() + excelRow.ToString();//单元格结束点 如:B4


            //_Range = _Worksheet.get_Range(point1, point2);//获取单元格
            objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
            if (cellCountSum > 0)
            {
                //_Range.MergeCells = true; //合并单元格
                objRange_Late.GetType().InvokeMember("MergeCells", BindingFlags.SetProperty, null, objRange_Late, new object[] { true });
            }

            objRange_Late.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });//内容水平居中 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter 
            objRange_Late.GetType().InvokeMember("VerticalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });  //内容垂直居中Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter 

            object Font = objRange_Late.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, objRange_Late, null);
            objRange_Late.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, Font, new object[] { 20 });

            //_Worksheet.Cells[_Range.Row, _Range.Column] = excelValue;//把内容写入单元格
            Parameters = new Object[1];
            Parameters[0] = excelValue;
            objRange_Late.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, objRange_Late, Parameters);
        }

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="Index">插入的行</param>
        /// <param name="value">插入的值</param>
        /// <param name="isTitle">是否为title</param>
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
        /// 写数据的内容画实线
        /// </summary>
        private void drawLine()
        {
            object Range = objSheet_Late.GetType().InvokeMember("UsedRange", BindingFlags.GetProperty, null, objSheet_Late, null);
            object[] args = new object[] { 1 };
            object Borders = Range.GetType().InvokeMember("Borders", BindingFlags.GetProperty, null, Range, null);
            Borders = Range.GetType().InvokeMember("LineStyle", BindingFlags.SetProperty, null, Borders, args);
            //自动适应行列的宽高
            if (autoFit)
            {
                Object dataColumns = Range.GetType().InvokeMember("Columns", BindingFlags.GetProperty, null, Range, null);
                dataColumns.GetType().InvokeMember("AutoFit", BindingFlags.InvokeMethod, null, dataColumns, null);
            }
        }

        //TODO:excel alignment
        /// <summary>
        /// 设置所有的数据对齐方式与显示的一致
        /// </summary>
        /// <param name="colName">列名数组</param>
        /// <param name="colName">Excel列名数组</param>
        /// <param name="maxRows">最大的行值</param>
        private void setAlignment(string[] colName,string[] columns, int maxRows)
        {
            //默认左对齐
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
                point1 = columns[i] + startRow.ToString();//单元格起始点 如:A3
                point2 = columns[i] + maxRows.ToString();//单元格结束点 如:A15
                objRange_Late = objSheet_Late.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, objSheet_Late, new object[] { point1, point2 });
                objRange_Late.GetType().InvokeMember("VerticalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });  //内容垂直居中Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter 
                if (this.Columns[col].DefaultCellStyle.Alignment.ToString().Contains("Right"))
                {
                    objRange_Late.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4152 });//Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight 
                }
                else if (this.Columns[col].DefaultCellStyle.Alignment.ToString().Contains("Center"))
                {
                    objRange_Late.GetType().InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, objRange_Late, new object[] { -4108 });//内容水平居中 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter 
                }
                i++;
            }
        }

        /// <summary>
        /// 左右行合并
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
        /// 上下合并单元格
        /// </summary>
        /// <param name="startRow">起始行</param>
        /// <param name="endRow">终止行</param>
        /// <param name="col">第几列</param>
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
        /// 写入title
        /// </summary>
        private void writeTitle()
        {
            if (!string.IsNullOrEmpty(title))
            {
                Insert(1, title, true);
            }
        }

        /// <summary>
        /// 写入header
        /// </summary>
        private void writeHeader()
        {
            for (int i = header.Count - 1; i >= 0; i--)
            {
                Insert(1, header[i], false);
            }
        }

        /// <summary>
        /// 写入footer
        /// </summary>
        /// <param name="maxRows">开始写入的行</param>
        private void writeFooter(int maxRows)
        {
            for (int i = 0; i < footer.Count; i++)
            {
                WriteExcels(maxRows + i, 1, visibleCols, footer[i]);
            }
        }

        /// <summary>
        /// 调用递归求最大层数,列头总高
        /// </summary>
        private void myNodeLevels()
        {

            iNodeLevels = 1;//初始值为1

            ColLists.Clear();
            allColLists.Clear();

            int iNodeDeep = myGetNodeLevels(_ColHeaderTreeView.Nodes);
            treeDepth = iNodeDeep;

            this.ColumnHeadersHeight = iCellHeight * iNodeDeep;//列头总高=一维列高*层数
            this.ColumnDeep = iNodeDeep;

        }
        /// <summary>
        /// 递归计算树的最大层数,保存所有的叶节点
        /// </summary>
        /// <param name="tnc"></param>
        /// <returns></returns>
        private int myGetNodeLevels(TreeNodeCollection tnc)
        {
            if (tnc == null) return 0;

            foreach (TreeNode tn in tnc)
            {
                if ((tn.Level + 1) > iNodeLevels)//tn.Level是从0开始的
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
                        ColLists.Add(tn);//页节点
                    }
                }
            }

            return iNodeLevels;
        }

        /// <summary>
        /// 数字转换为Excel字母数字的列
        /// </summary>
        /// <param name="columnNum">数字</param>
        /// <returns>返回字符串</returns>
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
        /// Excel字母形式的列转换为数字
        /// </summary>
        /// <param name="letters">字符</param>
        /// <returns>返回的int值</returns>
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
        /// 画单元格
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
                        
            //上面相同的行数
            int upRows = 0;
            //下面相同的行数
            int downRows = 0;
            //左边要合并的列
            int leftColumns = 0;
            //右边要合并的列
            int rightColumns = 0;

            Pen gridLinePen = new Pen(gridBrush);
            bool mergeRow = false;
            resetMergeRowColumns();
            #region 横向合并,必须加载完毕才能操作
            if (e.RowIndex != -1 && this._lstMergeRowColumn.Count > 0 && displayedColumns.Count >0)
            {
                foreach (List<string> lstRowColumn in _lstMergeRowColumn)
                {
                    //如果合并了,或者如果关键列不显示,就不再找了;
                    if (mergeRow || !Columns[lstRowColumn[0]].Visible ) break;

                    if (!lstRowColumn.Contains(this.Columns[e.ColumnIndex].Name)) continue;
                    #region 保证都连在一起
                    
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

                    #region 保证头为非空,其他为空
                
                    if (null != this.Rows[e.RowIndex].Cells[lstRowColumn[0]].Value )
                    {
                        string mainValue = this.Rows[e.RowIndex].Cells[lstRowColumn[0]].Value.ToString().Trim();
                        if (mainValue.Length == 0) continue;

                        bool canMergeRow = true;

                        for (int i = 1; i < lstRowColumn.Count ; i++)
                        {
                            //为空,或者不可见,或者为空字符串
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
                            //以背景色填充
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
                            // 画右边线
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

            #region 竖向合并
            if (!mergeRow && this._mergecolumnname.Contains(this.Columns[e.ColumnIndex].Name) 
                && e.RowIndex != -1 && e.Value != null && !string.IsNullOrEmpty (e.Value.ToString()))
            {
                string curValue = e.Value.ToString();
                
                if (!string.IsNullOrEmpty(curValue))
                {
                    #region 获取下面的行数
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
                    #region 获取上面的行数
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
                //以背景色填充
                e.Graphics.FillRectangle(backBrush, e.CellBounds);
                //画字符串
                PaintingFont(e, curValue,upRows, downRows, 0, 0);
                if (downRows == 0)
                {
                    e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);
                }
                // 画右边线
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
            //如果不到一行的情况下,字的高度设置为字体的高度
            if (fontwidth <= sumWidth) wordHeight = fontheight;
            //如果字的高度大于总高度,只好把总高度赋值给字的高度
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

                        //可以合并
                        bool canMergeRow = true;
                        //本行不用检查了,因为这个组合的某个列已经跟别的合并了
                        bool nowNeedToCheckThisGroup = false;

                        for (int i = 1; i < lstRowColumn.Count; i++)
                        {
                            //在合并集合中找到了当前行和列
                            if (checkHasTheColumn(lstRowColumn[0], rowNo).Length > 0)
                            {
                                nowNeedToCheckThisGroup = true;
                                break;
                            }
                            //为空,或者不可见,或者为空字符串
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

        #region 增强横向合并部分

        //记录所有的横向合并的情况;
        private Dictionary<int, Dictionary<string, List<string>>> stockOfAllRows = new Dictionary<int, Dictionary<string, List<string>>>();

        private List<string> displayedColumns = new List<string>();
        /// <summary>
        /// 重新设置显示列
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
        /// 释放资源
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
            string searchCol = "";  //要搜索的字段名（用逗号连接的字符串）
            if (e.KeyData == Keys.F3)
            {
                //把所有显示的列名用逗号分隔成字符串
                foreach (DataGridViewColumn dgvColumn in this.Columns)
                {
                    if (dgvColumn.Visible == true)
                    {
                        searchCol += dgvColumn.Name + ",";
                    }
                }

                //去掉最后一个逗号
                searchCol = searchCol.Substring(0, searchCol.Length - 1);

                if (searchCol == "")
                {
                    MessageBox.Show("找不到正在搜索的数据！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        /// 排序完毕冲刷
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected override void OnSorted(EventArgs e)
        {
            base.OnSorted(e);
            refreshHeader();
        }

        /// <summary>
        /// 抬起鼠标时冲刷
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

        //树形表头调整列宽时自动刷新控件
        protected override void OnColumnWidthChanged(DataGridViewColumnEventArgs e)
        {
            base.OnColumnWidthChanged(e);
            if (!loadColWidthFromDb && this._ColHeaderTreeView != null) refreshHeader();
        }

        //使用水平滚动条时刷新控件
        protected override void OnScroll(ScrollEventArgs e)
        {
            base.OnScroll(e);
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
            {
                refreshHeader();
                refresh(e);
            }
            
        }

        //显示行号
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
        /// 单元格绘制(重写)
        /// </summary>
        /// <param name="e"></param>
        /// <remarks></remarks>
        protected override void OnCellPainting(DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex > -1)
            {
                DrawCell(e);
            }
            //行标题不重写
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

            //绘制表头
            if (e.RowIndex == -1 && _ColHeaderTreeView != null && _ColHeaderTreeView.Nodes.Count > 0)
            {
                PaintUnitHeader((TreeNode)NadirColumnList[e.ColumnIndex], e, _columnDeep);
                e.Handled = true;
                return;
            }
            //画页脚汇总信息
            this.AllowUserToAddRows = true;
            if (e.RowIndex == this.NewRowIndex && e.ColumnIndex > -1 && computeColumns.Count>0)
            {
                DrawFooter(e);
            }

        }

        ///TODO:增加页脚信息
        //绘制footer的汇总信息
        private void DrawFooter(DataGridViewCellPaintingEventArgs e)
        {
            // 此处只画页脚信息
            this.AllowUserToAddRows = true;
            if (e.RowIndex == -1) return;
            
            ///很奇怪这里的s总是0，加不出来，但x却有值，不知道lambda表达式出了什么问题了。只能循环了
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
        /// 计算出所有计算列的名值对，并返回。
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, string> dictComputeColumns()
        {
            DataTable dt = GridToTable();
            Dictionary<string, string> dict = new Dictionary<string, string>();
            int i = 0; //按照顺序将值写入字典
            foreach (DataColumn col in dt.Columns)
            {
                string val=computeColumns[i].Substring(computeColumns[i].IndexOf(",") + 1);
                if (computeColumns[i].Contains('(')) //使用了聚合函数
                {
                    string v;
                    if (val.Split(',').Length > 1) //有过滤条件
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
                else //设置为合计等字符串的内容，没有()的内容来判断
                {
                    dict.Add(col.ColumnName, val);
                }
                i++;
            }
            return dict;
        }



        /// <summary>
        /// 转换为一个DataTable
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
        /// 返回设置的computeColumns的列集合
        /// </summary>
        /// <returns>特殊显示的列集合</returns>
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

        #region ---------多层表头部分-----------------------------
        public DataGridViewEx(IContainer container)
        {
            container.Add(this);
            InitializeComponent();
        }

        #region 多层表头的设定
        private TreeView _ColHeaderTreeView = new TreeView();


        /// <summary>
        /// 多维列标题的树结构
        /// </summary>
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        [Description("多维列表标题树结构,使用本属性时,必须要求显示列与树形控件叶节点数量一致,否则显示将异常,如果某个节点不显示,将其Tag置为hide"), Category("自定义属性")]
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
        [Description("设置或获得合并表头树的深度")]
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
        ///绘制合并表头
        ///</summary>
        ///<param name="node">合并表头节点</param>
        ///<param name="e">绘图参数集</param>
        ///<param name="level">结点深度</param>
        ///<remarks></remarks>
        public void PaintUnitHeader(TreeNode node, DataGridViewCellPaintingEventArgs e, int level)
        {
            //根节点时退出递归调用
            if (level == 0 ) //|| ( node.Tag!=null && node.Tag.ToString()=="hide"))
                return;

            RectangleF uhRectangle;
            int uhWidth;
            SolidBrush gridBrush = new SolidBrush(this.GridColor);

            Pen gridLinePen = new Pen(gridBrush);
            StringFormat textFormat = new StringFormat();


            textFormat.Alignment = StringAlignment.Center;

            
            uhWidth = GetUnitHeaderWidth(node);

            //与原贴算法有所区别在这。
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
            //画矩形
            e.Graphics.FillRectangle(backColorBrush, uhRectangle);

            //划底线
            e.Graphics.DrawLine(gridLinePen
                                , uhRectangle.Left
                                , uhRectangle.Bottom
                                , uhRectangle.Right
                                , uhRectangle.Bottom);
            //划右端线
            e.Graphics.DrawLine(gridLinePen
                                , uhRectangle.Right
                                , uhRectangle.Top
                                , uhRectangle.Right
                                , uhRectangle.Bottom);

            ////写字段文本
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

            ////递归调用()
            if (node.PrevNode == null)
                if (node.Parent != null)
                    PaintUnitHeader(node.Parent, e, level - 1);
        }

        /// <summary>
        /// 获得合并标题字段的宽度
        /// </summary>
        /// <param name="node">字段节点</param>
        /// <returns>字段宽度</returns>
        private int GetUnitHeaderWidth(TreeNode node)
        {
            int uhWidth = 0;
            //获得最底层字段的宽度
            if (node.Nodes == null)
                return this.Columns[GetColumnListNodeIndex(node)].Width;

            if (node.Nodes.Count == 0)
                return this.Columns[GetColumnListNodeIndex(node)].Width;

            //获得非最底层字段的宽度
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
        /// 获得底层字段索引
        /// </summary>
        /// <param name="node">底层字段节点</param>
        /// <returns>索引</returns>
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
        [Description("最底层节点集合")]
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
        /// 获得底层字段集合
        /// </summary>
        /// <param name="alList">底层字段集合</param>
        /// <param name="node">字段节点</param>
        /// <param name="checked">向上搜索与否</param>
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