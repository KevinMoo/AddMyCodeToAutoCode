using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace MSC.CommonLib
{
    public interface IGetList
    {
        DataSet GetList(string sWhereString);
        DataSet GetListAsAlias(string sWhereString);
    }
}
