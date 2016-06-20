using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MSC.WinFormControlLib
{
    public partial class frmNoCloseForm : MSC.WinFormControlLib.frmBase
    {
        public frmNoCloseForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 获取已设置无法关闭窗口创建参数。
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                int CS_NOCLOSE = 0x200;
                CreateParams parameters = base.CreateParams;
                parameters.ClassStyle |= CS_NOCLOSE;

                return parameters;
            }
        }
    }
}
