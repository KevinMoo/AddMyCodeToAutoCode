using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MSC.CommonLib
{
    public class CommonCode
    {

        public static int ToInt(string pString)
        {
            int i = 0;
            try
            {
                i = Convert.ToInt32(pString);
            }
            catch
            {
            }
            return i;
        }
        public static decimal ToDecimal(string pString)
        {
            decimal d = 0;
            try
            {
                d = Convert.ToDecimal(pString);
            }
            catch
            {
            }
            return d;
        }

        public static decimal StringToDecimal(string pS)
        {
            decimal num = 0M;
            try
            {
                num = Convert.ToDecimal(pS);
            }
            catch
            {
            }
            return num;
        }
    }
}
