using System;
using System.Collections.Generic;
using System.Text;

namespace MSC.CommonLib
{
    public class ModelColumnInfo
    {
        public ModelColumnInfo()
        {
        }
        public ModelColumnInfo(string pColumnName, string pColumnDescription, Type pColumnType)
        {
            this._columnName = pColumnName;
            this._columnDescription = pColumnDescription;
            this._columnType = pColumnType;
        }

        private string _columnName;
        private string _columnDescription;
        private Type _columnType;

        public string ColumnName
        {
            get
            {
                return this._columnName;
            }
        }
        public string ColumnDescription
        {
            get
            {
                return this._columnDescription;
            }
        }
        public Type ColumnType
        {
            get
            {
                return this._columnType;
            }
        }

        public override string ToString()
        {
            return this._columnDescription;
        }

        public ModelColumnInfo Self
        {
            get
            {
                return this;
            }
        }

    }
}
