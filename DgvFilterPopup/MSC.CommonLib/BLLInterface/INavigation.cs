using System;
using System.Collections.Generic;
using System.Text;

namespace MSC.CommonLib
{
    public interface INavigation
    {
         string GetMinBill();
         string GetMaxBill();
         string GetNextBill(string pBillNo);

         string GetPrevBill(string pBillNo);

    }
}
