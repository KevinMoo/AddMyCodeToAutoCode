namespace MSC.WinFormControlLib
{

    using MSC.CommonLib;
    using System;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Windows.Forms;
    using System.IO;
    using System.Diagnostics;
    using System.Collections.Generic;

    public class frmManagerReward : Form
    {
        private IColumnInfos _columnInfos;
        private IManagerForm _form;
        private IDelete _iDelete;
        private IGetList _iGetList;
        private string _mustWhereString = "";
        private IContainer components;
        private dgvHasRowNum dgvHasRowNumDetl;
        //private DataGridViewEx dgvHasRowNumDetl;
        private ToolStrip toolStrip1;
        private ToolStripButton toolStripButtonDelete;
        private ToolStripButton toolStripButtonEdit;
        private ToolStripButton toolStripButtonExit;
        private ToolStripButton toolStripButtonNew;
        private ToolStripButton toolStripButtonQuery;
        private ToolStripButton toolStripButtonRefresh;
        private ToolStripButton toolStripButtonReport;
        private ToolStripButton toolStripButtonNormal;
        private ToolStripButton toolStripButton1;
        private ToolStripSplitButton toolStripSplitButtonToExcel;
        private ToolStripMenuItem 快速ToolStripMenuItem;
        private ToolStripMenuItem 格式ToolStripMenuItem;
        private string[] whereString = new string[1];


        /// <summary>
        /// 管理单表记录 
        /// </summary>
        /// <param name="pIGetList">实现GetList接口的类，一般是BLL类</param>
        /// <param name="pIDelete">实现IDelete类，因为有的是视图，所以未实理删除，可以通过这个关连删除，为null说明不能删除</param>
        /// <param name="frm">实现IManagerForm 一般是缉窗口，为null 说明不能修改</param>
        /// <param name="infos">表的Model对像,与GetList的对应</param>
        /// <param name="pMustWhereString">必须添加的条件，为"" 表示不需要加</param>
        /// <param name="formTitle">窗口标题</param>
        public frmManagerReward(IGetList pIGetList, IDelete pIDelete, IManagerForm frm, 
            IColumnInfos infos, string pMustWhereString, string formTitle):this(pIGetList,
            pIDelete,frm,infos,pMustWhereString,formTitle,true)
        {
        }

        private bool _isAddSumRowToGrid = true ;

        /// <summary>
        /// 管理单表记录 
        /// </summary>
        /// <param name="pIGetList">实现GetList接口的类，一般是BLL类</param>
        /// <param name="pIDelete">实现IDelete类，因为有的是视图，所以未实理删除，可以通过这个关连删除，为null说明不能删除</param>
        /// <param name="frm">实现IManagerForm 一般是缉窗口，为null 说明不能修改</param>
        /// <param name="infos">表的Model对像,与GetList的对应</param>
        /// <param name="pMustWhereString">必须添加的条件，为"" 表示不需要加</param>
        /// <param name="formTitle">窗口标题</param>
        /// /// <param name="isAddSumRowToGrid">是否在表格增加合计行</param>
        public frmManagerReward(IGetList pIGetList, IDelete pIDelete, IManagerForm frm, 
            IColumnInfos infos, string pMustWhereString, string formTitle,bool pIsAddSumRowToGrid)
        {
            this.InitializeComponent();
            this._iDelete = pIDelete;
            if (this._iDelete == null)
            {
                this.toolStripButtonDelete.Enabled = false;
            }
            this._iGetList = pIGetList;
            this._form = frm;
            this._columnInfos = infos;
            if (this._form != null)
            {

                this._form._IsCanClose = false;
            }
            else
            {
                //未定交编辑
                this.toolStripButtonEdit.Enabled = false;
                this.toolStripButtonNew.Enabled = false;
            }

            this._mustWhereString = pMustWhereString;
            this.Text = formTitle;
            this._isAddSumRowToGrid = pIsAddSumRowToGrid;
        }

        private void dgvHasRowNumDetl_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            dgvHasRowNum num = (dgvHasRowNum) sender;
            if (e.RowIndex < 0)
            {
                return;
            }
            if (num.Rows[e.RowIndex].Cells[e.ColumnIndex].ValueType == typeof(byte[]))
            {
                string tag = ImageOperation.CrateImageFileByBytes((byte[]) num.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                if (tag != null)
                {
                    new frmPreviewImage(tag) { MdiParent = base.ParentForm.MdiParent }.Show();
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void frmManagerReward_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (this._form != null)
            {
                (this._form as Form).Close();
            }
        }

        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButtonNew = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonQuery = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonDelete = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonEdit = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonRefresh = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonReport = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonNormal = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonExit = new System.Windows.Forms.ToolStripButton();
            this.toolStripSplitButtonToExcel = new System.Windows.Forms.ToolStripSplitButton();
            this.快速ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.格式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgvHasRowNumDetl = new MSC.WinFormControlLib.dgvHasRowNum();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvHasRowNumDetl)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButtonNew,
            this.toolStripButtonQuery,
            this.toolStripButtonDelete,
            this.toolStripButtonEdit,
            this.toolStripButtonRefresh,
            this.toolStripSplitButtonToExcel,
            this.toolStripButton1,
            this.toolStripButtonReport,
            this.toolStripButtonNormal,
            this.toolStripButtonExit});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(679, 35);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButtonNew
            // 
            this.toolStripButtonNew.Image = global::MSC.WinFormControlLib.Resource1.New;
            this.toolStripButtonNew.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonNew.Name = "toolStripButtonNew";
            this.toolStripButtonNew.Size = new System.Drawing.Size(33, 32);
            this.toolStripButtonNew.Text = "新增";
            this.toolStripButtonNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButtonNew.Click += new System.EventHandler(this.toolStripButtonNew_Click);
            // 
            // toolStripButtonQuery
            // 
            this.toolStripButtonQuery.Image = global::MSC.WinFormControlLib.Resource1.Search;
            this.toolStripButtonQuery.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonQuery.Name = "toolStripButtonQuery";
            this.toolStripButtonQuery.Size = new System.Drawing.Size(57, 32);
            this.toolStripButtonQuery.Text = "查询明细";
            this.toolStripButtonQuery.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButtonQuery.Click += new System.EventHandler(this.toolStripButtonQuery_Click);
            // 
            // toolStripButtonDelete
            // 
            this.toolStripButtonDelete.Image = global::MSC.WinFormControlLib.Resource1.delete1;
            this.toolStripButtonDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonDelete.Name = "toolStripButtonDelete";
            this.toolStripButtonDelete.Size = new System.Drawing.Size(69, 32);
            this.toolStripButtonDelete.Text = "删除当前行";
            this.toolStripButtonDelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButtonDelete.Click += new System.EventHandler(this.toolStripButtonDelete_Click);
            // 
            // toolStripButtonEdit
            // 
            this.toolStripButtonEdit.Image = global::MSC.WinFormControlLib.Resource1.EDIT;
            this.toolStripButtonEdit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonEdit.Name = "toolStripButtonEdit";
            this.toolStripButtonEdit.Size = new System.Drawing.Size(69, 32);
            this.toolStripButtonEdit.Text = "编辑当前行";
            this.toolStripButtonEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButtonEdit.Click += new System.EventHandler(this.toolStripButtonEdit_Click);
            // 
            // toolStripButtonRefresh
            // 
            this.toolStripButtonRefresh.Image = global::MSC.WinFormControlLib.Resource1.refresh;
            this.toolStripButtonRefresh.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonRefresh.Name = "toolStripButtonRefresh";
            this.toolStripButtonRefresh.Size = new System.Drawing.Size(33, 32);
            this.toolStripButtonRefresh.Text = "刷新";
            this.toolStripButtonRefresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButtonRefresh.Click += new System.EventHandler(this.toolStripButtonRefresh_Click);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.Image = global::MSC.WinFormControlLib.Resource1.打印;
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(57, 32);
            this.toolStripButton1.Text = "报表套打";
            this.toolStripButton1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // toolStripButtonReport
            // 
            this.toolStripButtonReport.Image = global::MSC.WinFormControlLib.Resource1.打印;
            this.toolStripButtonReport.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonReport.Name = "toolStripButtonReport";
            this.toolStripButtonReport.Size = new System.Drawing.Size(81, 32);
            this.toolStripButtonReport.Text = "表格样式打印";
            this.toolStripButtonReport.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButtonReport.Click += new System.EventHandler(this.toolStripButtonReport_Click);
            // 
            // toolStripButtonNormal
            // 
            this.toolStripButtonNormal.Image = global::MSC.WinFormControlLib.Resource1.Print;
            this.toolStripButtonNormal.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonNormal.Name = "toolStripButtonNormal";
            this.toolStripButtonNormal.Size = new System.Drawing.Size(69, 32);
            this.toolStripButtonNormal.Text = "无格式打印";
            this.toolStripButtonNormal.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButtonNormal.Click += new System.EventHandler(this.toolStripButtonNormal_Click);
            // 
            // toolStripButtonExit
            // 
            this.toolStripButtonExit.Image = global::MSC.WinFormControlLib.Resource1.exit;
            this.toolStripButtonExit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonExit.Name = "toolStripButtonExit";
            this.toolStripButtonExit.Size = new System.Drawing.Size(33, 32);
            this.toolStripButtonExit.Text = "退出";
            this.toolStripButtonExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButtonExit.Click += new System.EventHandler(this.toolStripButtonExit_Click);
            // 
            // toolStripSplitButtonToExcel
            // 
            this.toolStripSplitButtonToExcel.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.快速ToolStripMenuItem,
            this.格式ToolStripMenuItem});
            this.toolStripSplitButtonToExcel.Image = global::MSC.WinFormControlLib.Resource1.excel;
            this.toolStripSplitButtonToExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripSplitButtonToExcel.Name = "toolStripSplitButtonToExcel";
            this.toolStripSplitButtonToExcel.Size = new System.Drawing.Size(75, 32);
            this.toolStripSplitButtonToExcel.Text = "导出Excel";
            this.toolStripSplitButtonToExcel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripSplitButtonToExcel.ButtonClick += new System.EventHandler(this.快速ToolStripMenuItem_Click);
            // 
            // 快速ToolStripMenuItem
            // 
            this.快速ToolStripMenuItem.Name = "快速ToolStripMenuItem";
            this.快速ToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.快速ToolStripMenuItem.Text = "快速";
            this.快速ToolStripMenuItem.Click += new System.EventHandler(this.快速ToolStripMenuItem_Click);
            // 
            // 格式ToolStripMenuItem
            // 
            this.格式ToolStripMenuItem.Name = "格式ToolStripMenuItem";
            this.格式ToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.格式ToolStripMenuItem.Text = "格式";
            this.格式ToolStripMenuItem.Click += new System.EventHandler(this.格式ToolStripMenuItem_Click);
            // 
            // dgvHasRowNumDetl
            // 
            this.dgvHasRowNumDetl.AllowUserToAddRows = false;
            this.dgvHasRowNumDetl.AllowUserToDeleteRows = false;
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvHasRowNumDetl.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7;
            this.dgvHasRowNumDetl.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvHasRowNumDetl.DefaultCellStyle = dataGridViewCellStyle8;
            this.dgvHasRowNumDetl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvHasRowNumDetl.Location = new System.Drawing.Point(0, 35);
            this.dgvHasRowNumDetl.Name = "dgvHasRowNumDetl";
            this.dgvHasRowNumDetl.ReadOnly = true;
            this.dgvHasRowNumDetl.RowTemplate.Height = 23;
            this.dgvHasRowNumDetl.Size = new System.Drawing.Size(679, 464);
            this.dgvHasRowNumDetl.TabIndex = 1;
            this.dgvHasRowNumDetl.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvHasRowNumDetl_CellContentDoubleClick);
            // 
            // frmManagerReward
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(679, 499);
            this.Controls.Add(this.dgvHasRowNumDetl);
            this.Controls.Add(this.toolStrip1);
            this.Name = "frmManagerReward";
            this.Text = "管理公共";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.frmManagerReward_Load);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmManagerReward_FormClosed);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvHasRowNumDetl)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void RefreshGrid()
        {
            if (this.whereString[0] != null)
            {
                string sWhereString = "";
                if (string.IsNullOrEmpty(this._mustWhereString))
                {
                    sWhereString = this.whereString[0];
                }
                else if (string.IsNullOrEmpty(this.whereString[0]))
                {
                    sWhereString = this._mustWhereString;
                }
                else
                {
                    sWhereString = this._mustWhereString + " And (" + this.whereString[0] + ")";
                }
                DataSet listAsAlias = this._iGetList.GetListAsAlias(sWhereString);
                //this.dgvHasRowNumDetl.DataSource = listAsAlias.Tables[0];

                if (this._isAddSumRowToGrid)
                {
                    new DataGridViewAddSumRow(this.dgvHasRowNumDetl, listAsAlias.Tables[0]);
                }
                else
                {
                      this.dgvHasRowNumDetl.DataSource = listAsAlias.Tables[0];

                }
            }
        }

        private void toolStripButtonDelete_Click(object sender, EventArgs e)
        {
            DataGridViewRow currentRow = this.dgvHasRowNumDetl.CurrentRow;
            if (currentRow == null)
            {
                DialogBox.ShowError("请选择要删除的行！");
            }
            else
            {
                string str = "";
                str = currentRow.Cells[0].Value.ToString();
                if (DialogBox.ShowQuestion("你真的要删除当前记录！") == DialogResult.Yes)
                {
                    if (this._iDelete.Delete(Convert.ToInt32(str)))
                    {
                        DialogBox.ShowInfo("删除成功！");
                        this.dgvHasRowNumDetl.Rows.Remove(currentRow);
                    }
                    else
                    {
                        DialogBox.ShowError("删除失败!");
                    }
                }
            }
        }

        private void toolStripButtonEdit_Click(object sender, EventArgs e)
        {
            DataGridViewRow currentRow = this.dgvHasRowNumDetl.CurrentRow;
            if (currentRow == null)
            {
                DialogBox.ShowError("请选择要编辑的行！");
            }
            else
            {
                int pAutoID = Convert.ToInt32(currentRow.Cells[0].Value.ToString());
                this._form.BindData(pAutoID, true);
                Form form = this._form as Form;
                form.MdiParent = base.MdiParent;
                form.Show();
            }
        }

        private void toolStripButtonExit_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void toolStripButtonNew_Click(object sender, EventArgs e)
        {
            this._form.BindData(-1, false);
            (this._form as Form).Show();
        }

        private void toolStripButtonQuery_Click(object sender, EventArgs e)
        {
            ModelColumnInfo[] getColumnInfos = this._columnInfos.GetColumnInfos;
            this.whereString[0] = null;
            new dbQuery(getColumnInfos, this.whereString).ShowDialog();
            this.RefreshGrid();

        }

        private void toolStripButtonRefresh_Click(object sender, EventArgs e)
        {
            this.RefreshGrid();
        }

        private void toolStripButtonToExcel_Click(object sender, EventArgs e)
        {
            if (this.dgvHasRowNumDetl.DataSource == null)
            {
                DialogBox.ShowError("请先查询");
                return;
            }
            //old
            //ExportToExcel();

            VBprinter.VB2008Print printControl = new VBprinter.VB2008Print();
            printControl.ExportDGVToExcel(dgvHasRowNumDetl, this.Text);
        }


        private void ExportToExcel()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Execl files (*.xls)|*.xls";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = true;
            saveFileDialog.Title = "保存为Excel文件";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName.IndexOf(":") < 0) return; //被点了"取消"

            Stream myStream;
            myStream = saveFileDialog.OpenFile();
            StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));
            string columnTitle = "";
            try
            {
                //写入列标题
                for (int i = 0; i < this.dgvHasRowNumDetl.ColumnCount; i++)
                {
                    if (i > 0)
                    {
                        columnTitle += "\t";
                    }
                    columnTitle += this.dgvHasRowNumDetl.Columns[i].HeaderText;
                }
                sw.WriteLine(columnTitle);

                //写入列内容
                for (int j = 0; j < this.dgvHasRowNumDetl.Rows.Count; j++)
                {
                    string columnValue = "";
                    for (int k = 0; k < this.dgvHasRowNumDetl.Columns.Count; k++)
                    {
                        if (k > 0)
                        {
                            columnValue += "\t";
                        }
                        if (this.dgvHasRowNumDetl.Rows[j].Cells[k].Value == null)
                            columnValue += "";
                        else
                            columnValue += this.dgvHasRowNumDetl.Rows[j].Cells[k].Value.ToString().Trim();
                    }
                    sw.WriteLine(columnValue);
                }
                sw.Close();
                myStream.Close();
                if (DialogBox.ShowQuestion("导出成功！\n\r是否立即打开此文件") == DialogResult.Yes)
                {
                    //OPEN;
                    Process proc = Process.Start(saveFileDialog.FileName);
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                sw.Close();
                myStream.Close();
            }
        }

        private void toolStripButtonReport_Click(object sender, EventArgs e)
        {
            VBprinter.DGVprint printerDgv = new VBprinter.DGVprint();
            printerDgv.MainTitle = this.Text;
            printerDgv.Print(dgvHasRowNumDetl, false);
        }

        private void toolStripButtonNormal_Click(object sender, EventArgs e)
        {
            if (this.dgvHasRowNumDetl.DataSource == null)
            {
                DialogBox.ShowError("请先查询");
                return;
            }
            PrintDGV.Print_DataGridView(this.dgvHasRowNumDetl);
        }

        private void frmManagerReward_Load(object sender, EventArgs e)
        {
            new DgvFilterPopup.DgvFilterManager(this.dgvHasRowNumDetl);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void 快速ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VBprinter.DGVprint printControl = new VBprinter.DGVprint();
            printControl.ExportDGVToExcel2(dgvHasRowNumDetl, this.Text, "", true);
        }

        private void 格式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.dgvHasRowNumDetl.DataSource == null)
            {
                DialogBox.ShowError("请先查询");
                return;
            }
            //old
            //ExportToExcel();

            VBprinter.VB2008Print printControl = new VBprinter.VB2008Print();
            printControl.ExportDGVToExcel(dgvHasRowNumDetl, this.Text);
        }
    }
}

