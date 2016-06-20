namespace MSC.WinFormControlLib
{
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Windows.Forms;

    public class ctrMyComboBox : ComboBox
    {
        private ComboBox comboBox1;
        private IContainer components;

        public ctrMyComboBox()
        {
            this.InitializeComponent();
        }

        public void BindData(ComboBoxItem[] items)
        {
            base.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            base.AutoCompleteSource = AutoCompleteSource.CustomSource;
            foreach (ComboBoxItem item in items)
            {
                base.AutoCompleteCustomSource.Add(item.Text);
            }
            base.Items.AddRange(items);
        }

        public void ChangeSelectByValue(object value)
        {
            int num = -1;
            for (int i = 0; i < base.Items.Count; i++)
            {
                ComboBoxItem item = (ComboBoxItem) base.Items[i];
                if (item.Value.Equals(value))
                {
                    num = i;
                    break;
                }
            }
            this.SelectedIndex = num;
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
            this.comboBox1 = new ComboBox();
            base.SuspendLayout();
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new Point(0, 0);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new Size(0x79, 20);
            this.comboBox1.TabIndex = 0;
            base.ResumeLayout(false);
        }
    }
}

