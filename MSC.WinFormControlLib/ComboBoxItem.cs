namespace MSC.WinFormControlLib
{
    using System;

    public class ComboBoxItem
    {
        private string _Text;
        private object _Value;

        public override string ToString()
        {
            return this._Text;
        }

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
    }
}

