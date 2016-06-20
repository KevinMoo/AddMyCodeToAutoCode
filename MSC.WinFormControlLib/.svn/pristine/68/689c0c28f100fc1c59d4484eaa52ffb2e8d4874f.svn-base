namespace MSC.WinFormControlLib
{
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Windows.Forms;
    using MSC.CommonLib;

    public class frmEditCommon : frmNoCloseForm, IManagerForm
    {
        private int _autoNo;
        private State _billState;
        protected IDelete _dal;
        private bool _isCanClose = true;
        protected readonly Color COLOR_READONLY_BACKCOLOR = Color.PeachPuff;
        protected readonly Color COLOR_READWRITE_BACKCOLOR = Color.White;
        private IContainer components;
        public ToolStripButton toolStripCancel;
        public ToolStripButton toolStripDelete;
        public ToolStripButton toolStripEdit;
        public ToolStripButton toolStripExit;
        public ToolStripButton toolStripNew;
        public ToolStripButton toolStripSave;
        public ToolStripSeparator toolStripSeparator2;
        private ToolStripSeparator toolStripSeparator4;
        public ToolStrip toolStripTool;

        public frmEditCommon()
        {
            this.InitializeComponent();
        }

        protected virtual bool BillDelete()
        {
            try
            {
                if (!this._dal.Delete(this._AutoNo))
                {
                    throw new Exception("操作失败，无出错信息，可能是因为数据库已无此记录！");
                }
            }
            catch (Exception exception)
            {
                DialogBox.ShowError("操作数据库失败！原因是:" + exception.Message);
                return false;
            }
            return true;
        }

        protected virtual void BindData()
        {
            if (this._AutoNo == 0)
            {
                DialogBox.ShowError("初始数据出错！");
            }
            this._BillState = State.Query;
            this.RefreshTool();
            this.ClearControls();
        }

        public void BindData(int pAutoNo)
        {
            this._AutoNo = pAutoNo;
            this.BindData();
            this.toolStripEdit_Click(new object(), new EventArgs());
        }

        public void BindData(int pAutoNo, bool pStartEdit)
        {
            if (pAutoNo == -1)
            {
                this.toolStripNew_Click(new object(), new EventArgs());
            }
            else
            {
                this.BindData(pAutoNo);
                if (pStartEdit)
                {
                    this._BillState = State.Edit;
                    this.RefreshTool();
                }
            }
        }

        private bool CanDelete()
        {
            string errorMessage = "";
            return this.CanDelete(ref errorMessage);
        }

        protected virtual bool CanDelete(ref string errorMessage)
        {
            return true;
        }

        private bool CanEdit()
        {
            string errorMessage = "";
            return this.CanEdit(ref errorMessage);
        }

        protected virtual bool CanEdit(ref string errorMessage)
        {
            return true;
        }

        protected virtual void ClearControls()
        {
            DialogBox.ShowError("未实现方法：ClearControls()");
        }

        public void ComboBoxChangeSelectByValue(ComboBox cmb, int value)
        {
            int num = -1;
            for (int i = 0; i < cmb.Items.Count; i++)
            {
                ComboBoxItem item = (ComboBoxItem) cmb.Items[i];
                if (Convert.ToInt32(item.Value) == value)
                {
                    num = i;
                    break;
                }
            }
            cmb.SelectedIndex = num;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            ComponentResourceManager resources = new ComponentResourceManager(typeof(frmEditCommon));
            this.toolStripTool = new ToolStrip();
            this.toolStripNew = new ToolStripButton();
            this.toolStripDelete = new ToolStripButton();
            this.toolStripSeparator2 = new ToolStripSeparator();
            this.toolStripEdit = new ToolStripButton();
            this.toolStripCancel = new ToolStripButton();
            this.toolStripSave = new ToolStripButton();
            this.toolStripSeparator4 = new ToolStripSeparator();
            this.toolStripExit = new ToolStripButton();
            this.toolStripTool.SuspendLayout();
            base.SuspendLayout();
            this.toolStripTool.Items.AddRange(new ToolStripItem[] { this.toolStripNew, this.toolStripDelete, this.toolStripSeparator2, this.toolStripEdit, this.toolStripCancel, this.toolStripSave, this.toolStripSeparator4, this.toolStripExit });
            this.toolStripTool.Location = new Point(0, 0);
            this.toolStripTool.Name = "toolStripTool";
            this.toolStripTool.Size = new Size(0x2ab, 0x23);
            this.toolStripTool.TabIndex = 1;
            this.toolStripTool.Text = "toolStrip1";
            this.toolStripNew.Image = Resource1.New;
            this.toolStripNew.ImageTransparentColor = Color.Magenta;
            this.toolStripNew.Name = "toolStripNew";
            this.toolStripNew.Size = new Size(0x21, 0x20);
            this.toolStripNew.Text = "新增";
            this.toolStripNew.TextImageRelation = TextImageRelation.ImageAboveText;
            this.toolStripNew.Click += new EventHandler(this.toolStripNew_Click);
            this.toolStripDelete.Enabled = false;
            this.toolStripDelete.Image = Resource1.delete1;
            this.toolStripDelete.ImageTransparentColor = Color.Magenta;
            this.toolStripDelete.Name = "toolStripDelete";
            this.toolStripDelete.Size = new Size(0x21, 0x20);
            this.toolStripDelete.Text = "删除";
            this.toolStripDelete.TextImageRelation = TextImageRelation.ImageAboveText;
            this.toolStripDelete.Click += new EventHandler(this.toolStripDelete_Click);
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new Size(6, 0x23);
            this.toolStripEdit.Enabled = false;
            this.toolStripEdit.Image = Resource1.EDIT;
            this.toolStripEdit.ImageTransparentColor = Color.Magenta;
            this.toolStripEdit.Name = "toolStripEdit";
            this.toolStripEdit.Size = new Size(0x21, 0x20);
            this.toolStripEdit.Text = "修改";
            this.toolStripEdit.TextImageRelation = TextImageRelation.ImageAboveText;
            this.toolStripEdit.Click += new EventHandler(this.toolStripEdit_Click);
            this.toolStripCancel.Enabled = false;
            this.toolStripCancel.Image = Resource1.撤销;
            this.toolStripCancel.ImageTransparentColor = Color.Magenta;
            this.toolStripCancel.Name = "toolStripCancel";
            this.toolStripCancel.Size = new Size(0x21, 0x20);
            this.toolStripCancel.Text = "取消";
            this.toolStripCancel.TextImageRelation = TextImageRelation.ImageAboveText;
            this.toolStripCancel.Click += new EventHandler(this.toolStripCancel_Click);
            this.toolStripSave.Enabled = false;
            this.toolStripSave.Image = Resource1.Save;
            this.toolStripSave.ImageTransparentColor = Color.Magenta;
            this.toolStripSave.Name = "toolStripSave";
            this.toolStripSave.Size = new Size(0x21, 0x20);
            this.toolStripSave.Text = "保存";
            this.toolStripSave.TextImageRelation = TextImageRelation.ImageAboveText;
            this.toolStripSave.Click += new EventHandler(this.toolStripSave_Click);
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new Size(6, 0x23);
            this.toolStripExit.Image = (Image) resources.GetObject("toolStripExit.Image");
            this.toolStripExit.ImageTransparentColor = Color.Magenta;
            this.toolStripExit.Name = "toolStripExit";
            this.toolStripExit.Size = new Size(0x21, 0x20);
            this.toolStripExit.Text = "退出";
            this.toolStripExit.TextImageRelation = TextImageRelation.ImageAboveText;
            this.toolStripExit.Click += new EventHandler(this.toolStripExit_Click);
            base.AutoScaleDimensions = new SizeF(6f, 12f);
            base.ClientSize = new Size(0x2ab, 0x221);
            base.Controls.Add(this.toolStripTool);
            base.Name = "frmEditCommon";
            this.toolStripTool.ResumeLayout(false);
            this.toolStripTool.PerformLayout();
            base.ResumeLayout(false);
            base.PerformLayout();
        }

        protected virtual void RefreshTool()
        {
            switch (this._BillState)
            {
                case State.None:
                    this.toolStripNew.Enabled = true;
                    this.toolStripDelete.Enabled = false;
                    this.toolStripEdit.Enabled = false;
                    this.toolStripCancel.Enabled = false;
                    this.toolStripSave.Enabled = false;
                    return;

                case State.New:
                case State.Edit:
                    this.toolStripNew.Enabled = false;
                    this.toolStripDelete.Enabled = false;
                    this.toolStripEdit.Enabled = false;
                    this.toolStripCancel.Enabled = true;
                    this.toolStripSave.Enabled = true;
                    return;

                case State.Query:
                {
                    bool flag = this.CanEdit();
                    this.toolStripNew.Enabled = true;
                    this.toolStripDelete.Enabled = this.CanDelete();
                    this.toolStripEdit.Enabled = flag;
                    this.toolStripCancel.Enabled = false;
                    this.toolStripSave.Enabled = false;
                    return;
                }
            }
        }

        protected virtual bool SaveBill(ref int pAutoNo)
        {
            DialogBox.ShowError("未实现方法：SaveBill()");
            return false;
        }

        private bool SaveBill(ref int pAutoNo, bool pIsWriteLog)
        {
            bool flag = this.SaveBill(ref pAutoNo);
            if (flag && pIsWriteLog)
            {
                State state1 = this._BillState;
            }
            return flag;
        }

        protected bool SaveCurrentOk()
        {
            if (MessageBox.Show("是否保存？", "系统询问", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                return this.SaveBill(ref this._autoNo, true);
            }
            return true;
        }

        private void toolStripCancel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("是否取消编辑", "系统询问", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.No)
            {
                if (this._BillState == State.New)
                {
                    this.ClearControls();
                    this._BillState = State.None;
                }
                else if (this._BillState == State.Edit)
                {
                    this.BindData();
                    this._BillState = State.Query;
                }
                this.RefreshTool();
            }
        }

        private void toolStripDelete_Click(object sender, EventArgs e)
        {
            if (this._autoNo == 0)
            {
                DialogBox.ShowInfo("无此单号");
            }
            else if (MessageBox.Show("是否删除?", "系统询问", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.No)
            {
                string errorMessage = "";
                if (!this.CanDelete(ref errorMessage))
                {
                    DialogBox.ShowInfo(errorMessage);
                }
                else
                {
                    if (this.BillDelete())
                    {
                        DialogBox.ShowInfo("删除成功!");
                        this._BillState = State.None;
                        this._AutoNo = 0;
                        this.ClearControls();
                    }
                    else
                    {
                        DialogBox.ShowInfo("删除失败");
                        this._BillState = State.Query;
                    }
                    this.RefreshTool();
                }
            }
        }

        private void toolStripEdit_Click(object sender, EventArgs e)
        {
            if (((this._BillState != State.New) && (this._BillState != State.Edit)) || this.SaveCurrentOk())
            {
                if (this._AutoNo == 0)
                {
                    DialogBox.ShowInfo("无单号");
                }
                else
                {
                    string errorMessage = "";
                    if (!this.CanEdit(ref errorMessage))
                    {
                        DialogBox.ShowError(errorMessage);
                    }
                    else
                    {
                        this._BillState = State.Edit;
                        this.RefreshTool();
                    }
                }
            }
        }

        private void toolStripExit_Click(object sender, EventArgs e)
        {
            if (((this._BillState != State.New) && (this._BillState != State.Edit)) || this.SaveCurrentOk())
            {
                if (this._IsCanClose)
                {
                    base.Close();
                }
                else
                {
                    this._BillState = State.Query;
                    base.Hide();
                }
            }
        }

        protected void toolStripNew_Click(object sender, EventArgs e)
        {
            if (((this._BillState != State.New) && (this._BillState != State.Edit)) || this.SaveCurrentOk())
            {
                this.ClearControls();
                this._autoNo = 0;
                this._BillState = State.New;
                this.RefreshTool();
            }
        }

        private void toolStripSave_Click(object sender, EventArgs e)
        {
            if (this.SaveBill(ref this._autoNo, true))
            {
                this.BindData();
                this._BillState = State.Query;
            }
            this.RefreshTool();
        }

        public int _AutoNo
        {
            get
            {
                return this._autoNo;
            }
            set
            {
                this._autoNo = value;
            }
        }

        public State _BillState
        {
            get
            {
                return this._billState;
            }
            set
            {
                this._billState = value;
            }
        }

        public bool _IsCanClose
        {
            get
            {
                return this._isCanClose;
            }
            set
            {
                this._isCanClose = value;
            }
        }

        public enum State
        {
            None,
            New,
            Edit,
            Query
        }
    }
}

