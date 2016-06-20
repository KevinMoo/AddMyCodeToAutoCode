using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DgvFilterPopup
{
    public class MyDataView:DataView
    {
        private string _myRowFilter;
        public string MyRowFilter
        {
            get { return _myRowFilter; }
            set
            {
                this._myRowFilter = value;
                this.RowFilter = value;
                OnMyRowFilterChanged(this, new EventArgs());
            }
        }

        public delegate void MyRowFilterChanged(object sender, EventArgs e);
        public event MyRowFilterChanged OnMyRowFilterChanged; 

    }
}
