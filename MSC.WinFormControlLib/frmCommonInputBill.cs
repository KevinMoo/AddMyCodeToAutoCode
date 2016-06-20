using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
//using MSC.WinFormControlLib.DB.Table;
//sing MSC.WinFormControlLib.Print;

namespace MSC.WinFormControlLib
{
    public partial class frmCommonInputBill : frmNoCloseForm
    {


        /// <summary>
        /// 单据状态
        /// </summary>
        public enum State
        {
            /// <summary>
            /// 初始壮态，空
            /// </summary>
            None,

            /// <summary>
            /// 新增状态
            /// </summary>
            New,
            /// <summary>
            /// 修改状态
            /// </summary>
            Edit,
            /// <summary>
            /// 查询状态
            /// </summary>
            Query
        }



        /// <summary>
        /// 单据审核状态
        /// </summary>
        public enum ApproveState
        {
            /// <summary>
            /// 未
            /// </summary>
            No,
            /// <summary>
            /// 已
            /// </summary>
            Yes,
            /// <summary>
            /// 读取数据库出错
            /// </summary>
            Error
        }


        /// <summary>
        /// 作废壮态
        /// </summary>
        public enum CancellationState
        {
            /// <summary>
            /// 未
            /// </summary>
            No,
            /// <summary>
            /// 已
            /// </summary>
            Yes,
            /// <summary>
            /// 读取数据库出错
            /// </summary>
            Error
        }





        /// <summary>
        /// 保存当前用户，当前表单的权限
        /// </summary>
        //protected Authority MyAuthority=new Authority();

        //protected LogOperation log = null;


        //protected Print.Print print = null;

        ///// <summary>
        ///// 表，[0]主表，……
        ///// </summary>
        //protected string[] TableName = null;

        ///// <summary>
        ///// 表对应的
        ///// </summary>
        //protected string[] TableKeyField = null;

        /// <summary>
        /// 审核字段字
        /// </summary>
        protected string ApproveFieldName = "IsApproved";

        /// <summary>
        /// 作废字段名
        /// </summary>
        protected string CancellationFieldName = "IsCanceled";


        private State billState= State.None;

        /// <summary>
        /// 当前单据状态
        /// </summary>
        public  State BillState
        {
            get
            {
                return billState;
            }
            set
            {
                billState = value;
            }
        }

        private string billNo = "";

        /// <summary>
        /// 当前单据号
        /// </summary>
        public string BillNo
        {
            get
            {
                return billNo;
            }
            set
            {
                billNo = value;
            }
        }

        /// <summary>
        /// 是否能关闭窗口,如果不关闭就Hide
        /// </summary>
        private  bool isCanClose = true;

        /// <summary>
        /// 是否能关闭窗口,如果不关闭就Hide,默认值true
        /// </summary>
        public bool IsCanClose
        {
            get
            {
                return isCanClose;
            }
            set
            {
                isCanClose = value;
            }
        }

        /// <summary>
        /// 只读单元格的背景色
        /// </summary>
        protected readonly Color COLOR_READONLY_BACKCOLOR = Color.PeachPuff;

        /// <summary>
        /// 可写单元格的背景色
        /// </summary>
        protected readonly Color COLOR_READWRITE_BACKCOLOR = Color.White;

        protected MSC.CommonLib.INavigation navigation;


        public frmCommonInputBill()
        {
            InitializeComponent();
            RefreshTool();
        }

        protected void toolStripNew_Click(object sender, EventArgs e)
        {
            if (BillState == State.New || BillState == State.Edit)
            {
                if (!SaveCurrentOk())
                {
                    return;
                }
            }

            //清空各控件的值
            ClearControls();
            billNo = "";
            BillState = State.New;
            RefreshTool();
        }

        /// <summary>
        /// 退出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripExit_Click(object sender, EventArgs e)
        {
            if (BillState == State.New || BillState == State.Edit)
            {
                if (!SaveCurrentOk())
                {
                    return;
                }
            }
            if (IsCanClose)
            {
                this.Close();
            }
            else
            {
                this.Hide();
            }
        }


        

        private void toolStripEdit_Click(object sender, EventArgs e)
        {
            if (BillState == State.New || BillState == State.Edit)
            {
                if (!SaveCurrentOk())
                {
                    return;
                }
            }

            if (billNo == string.Empty)
            {
                DialogBox.ShowError("无单号");
                return;
            }

            string errorMessage = "";
            if (!CanEdit(ref errorMessage))
            {
                DialogBox.ShowError(errorMessage);
                return;
            }

            BillState = State.Edit;
            RefreshTool();

        }

        /// <summary>
        /// 询问是否保存当前据据，
        /// </summary>
        /// <returns>如果保存不成功返回false,其它情况都返回True(保存成功或不保存）</returns>
        protected bool SaveCurrentOk()
        {
            bool result = false;
            if (MessageBox.Show("是否保存？",
                "系统询问",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (SaveBill(ref billNo,true))
                {
                    result = true;
                }
                else
                {
                    result = false;
                }
            }
            else
            {
                result = true;
            }
            return result;

        }

        /// <summary>
        /// 清空控件值
        /// </summary>
        protected virtual void ClearControls()
        {
            DialogBox.ShowError("未实现方法：" + "ClearControls()");
        }

        /// <summary>
        /// 保存单据过程,包括各项合法性验证,
        /// 注意：如果保存是否成功都在重载时弹出提示对话窗口，主调用不对信息显示。
        /// </summary>
        /// <returns></returns>
        protected virtual bool SaveBill(ref string pbillNo)
        {
            DialogBox.ShowError("未实现方法：" + "SaveBill()");
            return false;
        }

        /// <summary>
        /// 保存单据
        /// </summary>
        /// <param name="pBillNo">单号</param>
        /// <param name="IsWriteLog">是否在此方法里记录日志</param>
        /// <returns></returns>
        private bool SaveBill(ref string pBillNo,bool pIsWriteLog)
        {
            bool result = SaveBill(ref pBillNo);
            if (result && pIsWriteLog)
            {
                if (BillState == State.New)
                {
                }
                else
                {
                }
            }
            return result;
        }

        private void toolStripDelete_Click(object sender, EventArgs e)
        {
            if (billNo == string.Empty)
            {
                DialogBox.ShowInfo("无此单号");
                return;
            }

            if (MessageBox.Show("是否删除?",
                "系统询问", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                == DialogResult.No )
            {
                return;
            }


            string errorMessage="";

            if (!CanDelete(ref errorMessage))
            {
                DialogBox.ShowInfo(errorMessage);
                return;
            }

            if (BillDelete())
            {
                //保存成功
                DialogBox.ShowInfo("删除成功!");
                BillState = State.None;
                
                //写日志
                this.BillNo = null;
                ClearControls();
            }
            else
            {
                //保存失败
                DialogBox.ShowError("删除失败");
                BillState = State.Query;
            }

            RefreshTool();
            
        }

        /// <summary>
        /// 执行保存过程
        /// </summary>
        /// <returns></returns>
        protected virtual bool BillDelete()
        {
            //删除过程 未完成
            return true;
        }

        /// <summary>
        /// 当前单据能否删除或修改
        /// </summary>
        /// <param name="errorMessage">不能删除的原因</param>
        /// <returns>能true,不能false</returns>
        protected virtual bool CanDelete(ref string errorMessage)
        {
            errorMessage = "单据已扣帐或审核！";
            return false;
        }

        private bool CanDelete()
        {
            string errorMessage="";
            return CanDelete(ref errorMessage);
        }


                /// <summary>
        /// 当前单据能否修改
        /// </summary>
        /// <param name="errorMessage">不能修改的原因</param>
        /// <returns>能True,不能就false</returns>
        protected virtual bool CanEdit(ref string errorMessage)
        {
            errorMessage = "单据已扣帐或审核！";
            return false;
        }

        private bool CanEdit()
        {
            string errorMessage = "";
            return CanEdit(ref errorMessage);
        }

        /// <summary>
        /// 刷新工具栏
        /// </summary>
        protected virtual void RefreshTool()
        {
            switch (BillState)
            {
                case State.None:
                    this.toolStripNew.Enabled = true;
                    this.toolStripDelete.Enabled = false;
                    this.toolStripEdit.Enabled = false;
                    this.toolStripCancel.Enabled = false;
                    this.toolStripSave.Enabled = false;
                    this.toolStripApprove.Enabled = false;
                    this.toolStripReApprove.Enabled = false;
                    this.toolStripPrint.Enabled = false;
                    this.toolStripExport.Enabled = false;
                    this.toolStripCancellation.Enabled = false;
                    this.picFlag.Image = null;
                    break;
                case State.New:
                    goto case State.Edit;
                case State.Edit:
                    this.toolStripNew.Enabled = false ;
                    this.toolStripDelete.Enabled = false;
                    this.toolStripEdit.Enabled = false;
                    this.toolStripCancel.Enabled = true;
                    this.toolStripSave.Enabled = true ;
                    this.toolStripApprove.Enabled = false ;
                    this.toolStripReApprove.Enabled = false;
                    this.toolStripPrint.Enabled = false;
                    this.toolStripExport.Enabled = false;
                    this.toolStripCancellation.Enabled = false;
                    if (billState == State.New)
                    {
                        this.picFlag.Image = Resource1.Adding;
                    }
                    else
                    {
                        this.picFlag.Image = Resource1.Editing;
                    }
                    break;
                case State.Query:
                    bool canEdit = CanEdit();
                    this.toolStripNew.Enabled = true;
                    this.toolStripDelete.Enabled = CanDelete();
                    this.toolStripEdit.Enabled = canEdit;
                    this.toolStripCancel.Enabled = false ;
                    this.toolStripSave.Enabled = false ;
                    ApproveState approveState = IsApprove();
                    CancellationState cancellationState = IsCancelled();
                    this.picFlag.Image = Resource1.Saved;
                    switch (approveState)
                    {
                        case ApproveState.No :
                            this.toolStripApprove.Enabled=true;
                            this.toolStripReApprove.Enabled=false;
                            break;
                        case ApproveState.Yes :
                            this.toolStripApprove.Enabled=false;
                            this.toolStripReApprove.Enabled=true;
                            this.picFlag.Image = Resource1.Approved;
                            break;
                        case ApproveState.Error:
                            this.toolStripApprove.Enabled = false;
                            this.toolStripReApprove.Enabled = false;
                            break;
                    }

                    this.toolStripPrint.Enabled = true;
                    this.toolStripExport.Enabled = true;
                    switch (cancellationState)
                    {
                        case CancellationState.No:
                            //如果单据未审核并且可以修改 才能被取消，否则不能再取消
                            if (approveState != ApproveState.Yes && canEdit)
                            {
                                this.toolStripCancellation.Enabled = true;
                            }
                            else
                            {
                                this.toolStripCancellation.Enabled = false;
                            }
                            break;
                        case CancellationState.Yes:
                            //取消了不能再修改，也不能再做其它操作
                            this.toolStripNew.Enabled = true;
                            this.toolStripDelete.Enabled = false;
                            this.toolStripEdit.Enabled = false;
                            this.toolStripCancel.Enabled = false;
                            this.toolStripSave.Enabled = false;
                            this.toolStripApprove.Enabled = false;
                            this.toolStripReApprove.Enabled = false;
                            this.toolStripPrint.Enabled = false;
                            this.toolStripExport.Enabled = false;
                            this.toolStripCancellation.Enabled = false;
                            this.picFlag.Image = Resource1.Cancelled;
                            break;
                        case CancellationState.Error:
                            this.toolStripCancellation.Enabled = false;
                            break;
                    }




                    break;
            }
        }

        protected virtual CancellationState IsCancelled()
        {

                return CancellationState.Error;
           
        }
         
        /// <summary>
        /// 单据是审核壮态
        /// </summary>
        /// <returns>0未审核，1已审核，2错误</returns>
        protected virtual ApproveState IsApprove()
        {
            return ApproveState.Error;
        }

        /// <summary>
        /// 取消编辑
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripCancel_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("是否取消编辑",
                "系统询问", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                == DialogResult.No)
            {
                return;
            }

            if (BillState == State.New)
            {
                ClearControls();
                BillState = State.None;
            }
            else if (BillState == State.Edit)
            {
                BindData();
                BillState = State.Query;
            }
            RefreshTool();

        }

        /// <summary>
        /// 按单号，初始化各控件的数据绑定
        /// </summary>
        protected  virtual void BindData()
        {
            //MCT.CommonTool.Error_m(Common.IniFile.Segments[ERROR_MESSAGE].Items[SYSTEM].Value + "BindData(string BillNo)");
            if (BillNo == null || BillNo == string.Empty)
            {
                DialogBox.ShowError("初始数据出错！");
            }
            this.BillState = State.Query;
            RefreshTool();

            //绑定之前先清空
            ClearControls();

        }

        public void BindData(string pBillNo)
        {
            this.BillNo = pBillNo;
            BindData();
        }

        private void toolStripSave_Click(object sender, EventArgs e)
        {
            if (SaveBill(ref billNo,true))
            {
                //保存成功重新加载单据
                BillState = State.Query;
                BindData();
            }
            RefreshTool();
        }

        private void toolStripApprove_Click(object sender, EventArgs e)
        {
            if (ApproveBill())
            {
            }
            //RefreshTool();
            BindData();
        }

        /// <summary>
        /// 审核单据过程,包括各项合法性验证,
        /// 注意：如果审核是否成功都在重载时弹出提示对话窗口，主调用不对信息显示。   
        /// </summary>
        /// <returns></returns>
        protected virtual  bool ApproveBill()
        {
            //MCT.CommonTool.Error_m(Common.IniFile.Segments[ERROR_MESSAGE].Items[SYSTEM].Value + "ApproveBill()");
            return false;
        }

        private void toolStripReApprove_Click(object sender, EventArgs e)
        {
            if (ReApproveBill())
            {
                //
            }
            //RefreshTool();
            BindData();
        }

        /// <summary>
        /// 反审核单据过程,包括各项合法性验证,
        /// 注意：如果反审核是否成功都在重载时弹出提示对话窗口，主调用不对信息显示。   
        /// </summary>
        /// <returns></returns>
        protected virtual  bool ReApproveBill()
        {
            DialogBox.ShowError("未重载方法" + " ReApproveBill()");
            return false;
        }

        private void toolStripPrint_Click(object sender, EventArgs e)
        {
            Print();
        }


        private void Print()
        {
            //打印功能未实现
            //if (print == null)
            //{
            //    MCT.CommonTool.Error_m(Common.IniFile.Segments[ERROR_MESSAGE].Items[SYSTEM].Value+"print");
            //    return;
            //}
            //print.BillNo = this.BillNo;

            //log.Write(Common.IniFile.Segments[LOG].Items[PRINT].Value,
            //    Common.IniFile.Segments[LOG].Items[PRINT_BILL].Value+this.BillNo);

            //print.Preview();
        }

        private void toolStripCancellation_Click(object sender, EventArgs e)
        {
            if (CancellationBill())
            {

            }
            RefreshTool();
        }


        /// <summary>
        /// 作废单据过程,包括各项合法性验证,
        /// 注意：如果作废是否成功都在重载时弹出提示对话窗口，主调用不对信息显示。   
        /// </summary>
        /// <returns></returns>  
        protected virtual bool CancellationBill()
        {
            //MCT.CommonTool.Error_m(Common.IniFile.Segments[ERROR_MESSAGE].Items[SYSTEM].Value + " CancellationBill()");
            //return false;
            try
            {
                //string sqlText = "UPDATE "+TableName[0] +" SET "
                //    + CancellationFieldName + "=1"
                //    + " WHERE " + TableKeyField[0] + "='" + BillNo + "'";

                ////string errorMessage = "";
                //int rowNum = 0;
                ////if (!MCT.db.ExecuteNonQuery(sqlText,Common.SapConString,ref rowNum,ref errorMessage))
                ////{
                ////    //MCT.CommonTool.Error_m(Common.IniFile.Segments[ERROR_MESSAGE].Items[OPERATION_DB].Value + errorMessage);
                ////    return false;
                ////}

                //if (rowNum < 1)
                //{
                //    return false;
                //}
                //return true;
                //未完成
                return true;


            }
            catch
            {
                return false;
            }
        }

        private void toolStripButtonTop_Click(object sender, EventArgs e)
        {
            string billNo = this.navigation.GetMinBill();
            if (billNo == "")
            {
                DialogBox.ShowInfo("没有找到首记录！");
                return;
            }
            if (BillState == State.New || BillState == State.Edit)
            {
                if (!SaveCurrentOk())
                {
                    return;
                }
            }
            this.BindData(billNo);

        }

        //protected virtual bool IsHasBill(string pBillNo)
        //{
        //    MCT.CommonTool.Error_m("未重载方法" + " IsHasBill()");
        //    return false;
        //}

        //protected virtual string GetNextBill(string pBillNo)
        //{
        //    MCT.CommonTool.Error_m("未重载方法" + " GetNextBill()");
        //    return "";
        //}

        //protected virtual string GetPrevBill(string pBillNo)
        //{
        //    MCT.CommonTool.Error_m("未重载方法" + " GetPrevBill()");
        //    return "";
        //}
        //protected virtual string GetMaxBill()
        //{
        //    MCT.CommonTool.Error_m("未重载方法" + " GetMaxBill()");
        //    return "";
        //}

        //protected virtual string GetMinBill()
        //{
        //    MCT.CommonTool.Error_m("未重载方法" + " GetMinBill()");
        //    return "";
        //}

        private void toolStripButtonPrew_Click(object sender, EventArgs e)
        {
            //string billNo = GetPrevBill(this.BillNo);
            string billNo = this.navigation.GetPrevBill(this.BillNo);
            if (billNo == "")
            {
               DialogBox.ShowInfo("没有找到上一记录！");
                return;
            }
            if (BillState == State.New || BillState == State.Edit)
            {
                if (!SaveCurrentOk())
                {
                    return;
                }
            }
            this.BindData(billNo);
        }

        private void toolStripButtonNext_Click(object sender, EventArgs e)
        {
            string billNo = this.navigation.GetNextBill(this.BillNo);
            if (billNo == "")
            {
                DialogBox.ShowInfo("没有找到下一记录！");
                return;
            }

            if (BillState == State.New || BillState == State.Edit)
            {
                if (!SaveCurrentOk())
                {
                    return;
                }
            }
            this.BindData(billNo);
        }

        private void toolStripButtonButtom_Click(object sender, EventArgs e)
        {
            
            string billNo = this.navigation.GetMaxBill();
            if (billNo == "")
            {
                DialogBox.ShowInfo("没有找到尾记录！");
                return;
            }

            if (BillState == State.New || BillState == State.Edit)
            {
                if (!SaveCurrentOk())
                {
                    return;
                }
            }
            this.BindData(billNo);
        }



        
    }
}