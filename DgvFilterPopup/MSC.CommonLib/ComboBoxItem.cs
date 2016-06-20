using System;
using System.Collections.Generic;
using System.Text;

namespace MSC.CommonLib
{
    /// <summary>
    /// ComboBox 项类，让ComboBox 支持Text和Value方式。
    /// 取值时记得类型转换，如：
    /// ComboBoxItem myItem = (ComboBoxItem)ComboBox1.Items[0];
    /// string strValue = (string)myItem.Value;
    /// </summary>
    public class ComboBoxItem
    {

        /// <summary>
        /// ComboBox 项类，让ComboBox 支持Text和Value方式。
        /// 取值时记得类型转换，如：
        /// ComboBoxItem myItem = (ComboBoxItem)ComboBox1.Items[0];
        /// string strValue = (string)myItem.Value;
        /// </summary>
        public ComboBoxItem()
        {
        }


        private string _Text = null;
        private object _Value = null;

        public string Text
        {
            get
            {
                return this._Text;
            }
            set
            {
                this._Text = value;
            }
        }

        public object Value
        {
            get
            {
                return this._Value;
            }

            set
            {
                this._Value = value;
            }
        }

        public override string ToString()
        {
            return this._Text;
        }
    }
}
