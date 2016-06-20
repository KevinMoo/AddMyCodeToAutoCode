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
    public partial class frmBase : Form
    {
        public frmBase()
        {
            InitializeComponent();
        }

        private bool _isEnterFocusNextControl;

        [Category("外观"), Description("是否按回车跳到下一控件"), Browsable(true)]
        public bool IsEnterFocusNextControl
        {
            get
            {
                return  this._isEnterFocusNextControl;
            }
            set 
            {
                this._isEnterFocusNextControl = value;
            }
        }
        private int _functionId;

        
        [Category("其它"), Description("功能ID号"), Browsable(true)]
        public int FunctionId
        {
            get
            {
                return this._functionId;
            }
            set
            {
                this._functionId=value;
            }
        }


        protected virtual void DoFocusNextControl(KeyPressEventArgs e)
        {
            Control activeControl = base.ActiveControl;
            if (!this.IsMultiLineInputControl(activeControl))
            {
                base.SelectNextControl(activeControl, true, false, true, true);
                e.KeyChar = '\0';
                e.Handled = true;
            }
        }

        private void frmBase_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar == '\r') && this._isEnterFocusNextControl)
            {
                this.DoFocusNextControl(e);
            }
        }

        protected virtual bool IsMultiLineInputControl(Control c)
        {
            return ((c is TextBox) && (c as TextBox).Multiline);
        }

        


    }
}
