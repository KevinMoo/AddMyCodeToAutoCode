using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MSC.WinFormControlLib
{
    /// <summary>
    /// 快速搜索对话框窗体
    /// </summary>
    public partial class InputBox : Form
    {
        /// <summary>
        /// 搜索内容
        /// </summary>
        private string searchValue = "";

        private string msg = "请输入要查找的内容！";

        public string Msg
        {
            get { return msg; }
            set 
            {
                 msg = value;
                 this.label1.Text = msg; 
            }
        }

        /// <summary>
        /// 搜索内容
        /// </summary>
        public string SearchValue
        {
            get { return searchValue; }
            set { searchValue = value; }
        }

        /// <summary>
        /// 构造函数

        /// </summary>
        public InputBox()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 设置搜索内容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOk_Click(object sender, EventArgs e)
        {
            if (txtInputBox.Text.Trim().Length == 0)
            {
                MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtInputBox.Focus();

                //设置对话框继续运行
                this.DialogResult = DialogResult.None;
            }
            else
            {
                this.SearchValue = txtInputBox.Text;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        /// <summary>
        /// 关闭窗体
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        /// <summary>
        /// 当在快速搜索文本框中按回车时执行btnOk功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtInputBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.btnOk_Click(sender, e);
            }
        }
    }
}