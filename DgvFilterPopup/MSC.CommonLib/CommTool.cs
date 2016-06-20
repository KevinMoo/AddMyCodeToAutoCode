using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MSC.CommonLib
{
 




        public class ASPTool
        {

            public static void Alert(string s, System.Web.UI.Page page)
            {
                page.ClientScript.RegisterClientScriptBlock(page.GetType(), "", "alert(\"" + s + "\");", true);
            }
        }


    



        /// <summary>
        /// 字符串工具
        /// </summary>
        public class StringTool
        {
            /// <summary>
            /// 用正则表达式样式判断是否match
            /// </summary>
            /// <param name="input"></param>
            /// <param name="pat"></param>
            /// <returns></returns>
            public static bool IsMatch(string input, string pat)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(input, pat))
                    return true;
                else
                    return false;
            }




            /// <summary>
            /// 检查一个字符串是否是数字.
            /// </summary>
            /// <param name="val"></param>
            /// <returns></returns>
            public static bool CheckIfInt(string val)
            {
                val = val.Trim();

                if (val.StartsWith("-"))
                    val = val.Remove(0, 1);

                if (val == "")
                    return false;

                if (System.Text.RegularExpressions.Regex.IsMatch(val, "[^0-9]"))
                    return false;
                else
                    return true;
            }




            /// <summary>
            /// 根据分割符将字符串割成数组
            /// </summary>
            /// <param name="value"></param>
            /// <param name="sep"></param>
            /// <returns></returns>
            public static string[] MSplit(string value, string sep)
            {
                string s = value;
                int len = 0;

                while (s.IndexOf(sep) >= 0)
                {
                    s = s.Remove(0, s.IndexOf(sep) + sep.Length);
                    len++;
                }

                string[] arr;
                s = value;
                arr = new string[len + 1];
                int i = 0;

                while (s.IndexOf(sep) >= 0)
                {
                    arr[i] = s.Substring(0, s.IndexOf(sep));
                    i++;
                    s = s.Remove(0, s.IndexOf(sep) + sep.Length);
                }

                arr[i] = s;

                return arr;
            }




            /// <summary>
            /// 根据长度将字符串割成数组
            /// </summary>
            /// <param name="value"></param>
            /// <param name="len"></param>
            /// <returns></returns>
            public static string[] MSplit(string value, int len)
            {
                string[] arr;
                int length;

                length = (int)(Math.Ceiling((double)(value.Length / len)));
                arr = new string[length];

                string s = value;
                for (int i = 0; i < arr.Length; i++)
                {
                    arr[i] = s.Substring(0, Math.Min(len, s.Length));
                    s = s.Remove(0, Math.Min(len, s.Length));

                    if (s == "")
                        break;
                }

                return arr;
            }




            /// <summary>
            /// 从字符串的右侧移除指定长度的子串, 返回移除后的串
            /// </summary>
            /// <param name="val"></param>
            /// <param name="len"></param>
            /// <returns></returns>
            public static string RemoveTail(string val, int len)
            {
                if (len <= 0)
                    return val;
                else if (len >= val.Length)
                    return "";
                else
                    return val.Substring(0, val.Length - len);
            }




            /// <summary>
            /// 返回串左侧指定长的子串,
            /// </summary>
            /// <param name="val"></param>
            /// <param name="len">长度小于0时,返回空串,大于母串长时,返回母串</param>
            /// <returns></returns>
            public static string Left(string val, int len)
            {
                if (len <= 0)
                    return "";
                else if (len >= val.Length)
                    return val;
                else
                    return val.Substring(0, len);
            }



            /// <summary>
            /// 返回串右侧指定长的子串
            /// </summary>
            /// <param name="val"></param>
            /// <param name="len"></param>
            /// <returns></returns>
            public static string Right(string val, int len)
            {
                if (len <= 0)
                    return "";
                else if (len >= val.Length)
                    return val;
                else
                    return val.Remove(0, val.Length - len);
            }









            /// <summary>
            /// 将日期变量转换成 'YYYY/MM/DD HH:MM:SS' 的形式
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public static string FormatDateTime(DateTime dt)
            {
                string ret = "";

                ret += dt.Year.ToString("0000") + "/" +
                    dt.Month.ToString("00") + "/" +
                    dt.Day.ToString("00") + " " +
                    dt.Hour.ToString("00") + ":" +
                    dt.Minute.ToString("00") + ":" +
                    dt.Second.ToString("00");

                return ret;
            }





            /// <summary>
            /// 将日期型变量转换成'YYYY/MM/DD' 的形式
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public static string FormatDate(DateTime dt)
            {
                return dt.Year.ToString("0000") + "/" +
                    dt.Month.ToString("00") + "/" +
                    dt.Day.ToString("00");
            }




            /// <summary>
            /// 从 'YYYY/MM/DD HH:MM:SS' 的形式提取出日期型变量 
            /// </summary>
            /// <param name="val"></param>
            /// <returns></returns>
            public static DateTime GetDateTimeFromString(string val)
            {
                string pat = "^\\d{4,4}/\\d{2,2}/\\d{2,2} \\d{2,2}:\\d{2,2}:\\d{2,2}$";
                if (!StringTool.IsMatch(val, pat))
                    return new DateTime();
                string sYear, sMonth, sDay, sHour, sMinute, sSecond;
                sYear = StringTool.MTable.getItem(val, " ", "/", 0, 0);
                sMonth = StringTool.MTable.getItem(val, " ", "/", 0, 1);
                sDay = StringTool.MTable.getItem(val, " ", "/", 0, 2);

                sHour = StringTool.MTable.getItem(val, " ", ":", 1, 0);
                sMinute = StringTool.MTable.getItem(val, " ", ":", 1, 1);
                sSecond = StringTool.MTable.getItem(val, " ", ":", 1, 2);

                if (sYear == "" || sMonth == "" || sDay == "" || sHour == "" || sMinute == "" || sSecond == "")
                    return new DateTime();
                else
                {
                    DateTime dt = new DateTime(Convert.ToInt32(sYear), Convert.ToInt32(sMonth), Convert.ToInt32(sDay),
                                    Convert.ToInt32(sHour), Convert.ToInt32(sMinute), Convert.ToInt32(sSecond));
                    return dt;
                }
            }




            /// <summary>
            /// 从'YYYY/MM/DD' 的形式提取出日期型变量 
            /// </summary>
            /// <param name="val"></param>
            /// <returns></returns>
            public static DateTime GetDateFromString(string val)
            {
                string pat = "^\\d{4,4}/\\d{2,2}/\\d{2,2}$";

                if (!StringTool.IsMatch(val, pat))
                    return new DateTime();
                string sYear, sMonth, sDay;
                sYear = StringTool.MTable.getItem(val, " ", "/", 0, 0);
                sMonth = StringTool.MTable.getItem(val, " ", "/", 0, 1);
                sDay = StringTool.MTable.getItem(val, " ", "/", 0, 2);

                if (sYear == "" || sMonth == "" || sDay == "")
                    return new DateTime();
                else
                {
                    DateTime dt = new DateTime(Convert.ToInt32(sYear), Convert.ToInt32(sMonth), Convert.ToInt32(sDay),
                                    0, 0, 0);
                    return dt;
                }

            }






            /// <summary>
            /// 生成工单号的年月日三位,参数应该是6位的字符串,月日不足两位的,前面补0.
            /// 使用时,应检查返回值. 如果参数不合法, 则返回以"Err:" 开头的错误信息.
            /// </summary>
            /// <param name="sVal"></param>
            /// <returns></returns>
            public static string Gen_XDate(string sVal)
            {
                string sPat = "^\\d{6,6}$";
                if (!MSC.CommonLib.StringTool.IsMatch(sVal, sPat))
                    return ("Err:Argument illegal.An example:071203.");

                string[] arr = MSC.CommonLib.StringTool.MSplit(sVal, 2);

                int iYear, iMonth, iDay;
                iYear = Convert.ToInt32(arr[0]);
                iMonth = Convert.ToInt32(arr[1]);
                iDay = Convert.ToInt32(arr[2]);

                if (iMonth < 1 || iMonth > 12)
                {
                    return ("Err:Month must between 1 and 12.");
                }
                if (iDay < 1 || iDay > 31)
                {
                    return ("Err:Day must between 1 and 31.");
                }

                string sYear, sMonth, sDay;
                char cYear, cMonth, cDay, c;
                c = 'A';

                if (iYear > 9)
                {
                    cYear = (char)((int)c + iYear - 10);
                    sYear = cYear.ToString();
                }
                else
                    sYear = iYear.ToString();

                if (iMonth > 9)
                {
                    cMonth = (char)((int)c + iMonth - 10);
                    sMonth = cMonth.ToString();
                }
                else
                    sMonth = iMonth.ToString();

                if (iDay > 9)
                {
                    cDay = (char)((int)c + iDay - 10);
                    sDay = cDay.ToString();
                }
                else
                    sDay = iDay.ToString();

                return sYear + sMonth + sDay;

            }






            /// <summary>
            /// 将日期变量返回成形如 YYMMDD 形式的六位字符串
            /// </summary>
            /// <param name="dt"></param>
            /// <returns></returns>
            public static string Gen_ShortDateString(DateTime dt)
            {
                int iYear, iMonth, iDay;
                iYear = dt.Year;
                iMonth = dt.Month;
                iDay = dt.Day;

                string sYear, sMonth, sDay;
                sYear = iYear.ToString("0000").Substring(2, 2);
                sMonth = iMonth.ToString("00");
                sDay = iDay.ToString("00");

                return sYear + sMonth + sDay;

            }



            /// <summary>
            /// 取得一个由分隔符分隔的字符串的分段信息. 如果参数不合法, 将取最接近的值返回,而不返回错误信息
            /// </summary>
            /// <param name="input"></param>
            /// <param name="sep"></param>
            /// <param name="section">从0开始编号的段号</param>
            /// <returns></returns>
            public static string getSectionValue(string input, string sep, int section)
            {
                string[] arr = MSC.CommonLib.StringTool.MSplit(input, sep);
                if (section >= arr.Length)
                    return arr[arr.Length - 1];
                else
                    return arr[section];
            }


            public static string GetStartAtSplitChar(string pOriString, string pSplitString)
            {
                int at = pOriString.IndexOf(pSplitString);
                return pOriString.Substring(0, at);
            }





            /// <summary>
            /// 基于字符串的二维表类, 以';' 作为行分隔符, 以',' 作为列分隔符
            /// </summary>
            public class MTable
            {
                private int iRow = -1, iClm = -1;
                private MTableRow[] oRows;

                /// <summary>
                /// 构造一张空表
                /// </summary>
                /// <param name="iRowCnt"></param>
                /// <param name="iColumnCnt"></param>
                public MTable(int iRowCnt, int iColumnCnt)
                {
                    if (iRowCnt < 1 || iColumnCnt < 1)
                        throw new Exception("Row or column count must more than 0.");
                    iRow = iRowCnt;
                    iClm = iColumnCnt;

                    oRows = new MTableRow[iRowCnt];

                    for (int i = 0; i < iRow; i++)
                        oRows[i] = new MTableRow(iClm);
                }



                /// <summary>
                /// 从字符串构造表
                /// </summary>
                /// <param name="val"></param>
                public MTable(string val)
                {
                    FromString(val);
                }



                /// <summary>
                /// 只读属性, 返回当前表的行数
                /// </summary>
                public int RowLength
                {
                    get
                    {
                        return iRow;
                    }
                }



                /// <summary>
                /// 只读属性,返回当前表的列数
                /// </summary>
                public int ColumnLength
                {
                    get
                    {
                        return iClm;
                    }
                }



                /// <summary>
                /// 获取或设置某一行
                /// </summary>
                public MTableRow[] Rows
                {
                    get
                    {
                        return oRows;
                    }
                    set
                    {
                        oRows = value;
                    }
                }



                /// <summary>
                /// 根据某一列的值, 查询此值所在的行
                /// </summary>
                /// <param name="itemIndex"></param>
                /// <param name="itemValue"></param>
                /// <returns></returns>
                public MTableRow GetRow(int itemIndex, string itemValue)
                {
                    for (int i = 0; i < iRow; i++)
                    {
                        if (oRows[i].Items[itemIndex] == itemValue)
                            return oRows[i];
                    }

                    return null;
                }




                /// <summary>
                /// 删除指定行
                /// </summary>
                /// <param name="index"></param>
                public void RemoveRowAt(int index)
                {
                    if (index < 0 || index >= iRow || iRow < 1)
                        throw new Exception("index is not valid");
                    MTableRow[] tRows = new MTableRow[iRow - 1];

                    for (int i = 0; i < iRow; i++)
                    {
                        if (i < index)
                            tRows[i] = oRows[i];
                        else if (i > index)
                            tRows[i - 1] = oRows[i];
                    }

                    iRow -= 1;
                    oRows = new MTableRow[iRow];

                    for (int i = 0; i < iRow; i++)
                        oRows[i] = tRows[i];
                }





                /// <summary>
                /// 删除最后一行
                /// </summary>
                public void RemoveRow()
                {
                    RemoveRowAt(iRow - 1);
                }






                /// <summary>
                /// 在表的末尾添加一行
                /// </summary>
                /// <param name="val"></param>
                public void AddRow(string val)
                {
                    Array.Resize<MTableRow>(ref oRows, iRow + 1);
                    oRows[iRow] = new MTableRow(val);
                    iRow++;
                }






                /// <summary>
                /// 在表的末尾添加一行
                /// </summary>
                /// <param name="oRow"></param>
                public void AddRow(MTableRow oRow)
                {
                    Array.Resize<MTableRow>(ref oRows, iRow + 1);
                    oRows[iRow] = oRow;
                    iRow++;
                }





                /// <summary>
                /// 在指定位置插入行
                /// </summary>
                /// <param name="index"></param>
                /// <param name="row"></param>
                public void InsertRowAt(int index, MTableRow row)
                {
                    if (row.Length != iClm)
                        throw new Exception("row length is not equal to current table.");
                    if (index >= iRow)
                    {
                        AddRow(row);
                        return;
                    }
                    else if (index < 0)
                        index = 0;

                    Array.Resize<MTableRow>(ref oRows, iRow + 1);
                    for (int i = iRow; i > index; i--)
                        oRows[i] = oRows[i - 1];

                    oRows[index] = row;
                    iRow++;
                }





                /// <summary>
                /// 从指定字符串构造表
                /// </summary>
                /// <param name="val"></param>
                public void FromString(string val)
                {
                    if (val.Trim() == "")
                    {
                        iRow = 0;
                        iClm = 0;
                        oRows = new MTableRow[0];
                        return;
                    }


                    string[] arrRow = StringTool.MSplit(val, ";");

                    int iTemp = 0;
                    for (int i = 0; i < arrRow.Length; i++)
                    {
                        if (i == 0)
                            iTemp = StringTool.MSplit(arrRow[i], ",").Length;
                        else if (StringTool.MSplit(arrRow[i], ",").Length != iTemp)
                            throw new Exception("Argument is not valid");
                    }

                    iClm = iTemp;
                    iRow = arrRow.Length;

                    oRows = new MTableRow[iRow];
                    for (int i = 0; i < iRow; i++)
                        oRows[i] = new MTableRow(arrRow[i]);

                }




                /// <summary>
                /// 将当前表导出为字符串
                /// </summary>
                /// <returns></returns>
                public override string ToString()
                {
                    if (iRow == -1 || iClm == -1 || oRows == null)
                        throw new Exception("null reference. object is not initialized properly.");

                    string str = "";

                    for (int i = 0; i < iRow; i++)
                        str += oRows[i].ToString() + ";";
                    str = StringTool.RemoveTail(str, 1);
                    return str;
                }




                //**********************************************************








                /// <summary>
                /// 二维表的行类
                /// </summary>
                public class MTableRow
                {

                    private string[] arrItem;


                    /// <summary>
                    /// 构造空行
                    /// </summary>
                    /// <param name="itemCnt"></param>
                    public MTableRow(int itemCnt)
                    {
                        arrItem = new string[itemCnt - 1];
                        for (int i = 0; i < arrItem.Length; i++)
                            arrItem[i] = "";
                    }



                    /// <summary>
                    /// 从字符串构造一行
                    /// </summary>
                    /// <param name="str"></param>
                    public MTableRow(string str)
                    {
                        arrItem = StringTool.MSplit(str, ",");
                    }




                    /// <summary>
                    /// 只读属性, 返回此行的列数
                    /// </summary>
                    public int Length
                    {
                        get
                        {
                            return arrItem.Length;
                        }
                    }




                    /// <summary>
                    /// 获取或设置本行的某一列的值
                    /// </summary>
                    public string[] Items
                    {
                        get
                        {
                            return arrItem;
                        }
                        set
                        {
                            arrItem = value;
                        }
                    }




                    /// <summary>
                    /// 根据值查询某一列
                    /// </summary>
                    /// <param name="itemValue"></param>
                    /// <returns></returns>
                    public string GetItem(string itemValue)
                    {
                        for (int i = 0; i < arrItem.Length; i++)
                            if (arrItem[i] == itemValue)
                                return arrItem[i];
                        return "";
                    }




                    /// <summary>
                    /// 从字串构造行
                    /// </summary>
                    /// <param name="str"></param>
                    public void FromString(string str)
                    {
                        arrItem = StringTool.MSplit(str, ",");
                    }





                    /// <summary>
                    /// 将行导出为字串
                    /// </summary>
                    /// <returns></returns>
                    public override string ToString()
                    {
                        if (arrItem == null)
                            throw new Exception("object is not initialized properly.");

                        string str = "";
                        for (int i = 0; i < arrItem.Length; i++)
                            str += arrItem[i] + ",";
                        str = StringTool.RemoveTail(str, 1);
                        return str;
                    }




                }//end class mTableRow







                //************************************************************




                /// <summary>
                /// 静态方法, 获取一个带分隔符的字串的指定段的值.第三个参数从0开始计数.
                /// 如:getItem("1,2,3",",",1)=="2"
                /// </summary>
                /// <param name="input"></param>
                /// <param name="sep">"分隔符"</param>
                /// <param name="pos">从0开始计数的位置</param>
                /// <returns></returns>
                public static string getItem(string input, string seperator, int pos)
                {
                    string[] arr = StringTool.MSplit(input, seperator);

                    if (pos >= arr.Length || pos < 0)
                        return "";
                    else
                        return arr[pos];
                }





                /// <summary>
                /// 
                /// </summary>
                /// <param name="input"></param>
                /// <param name="sep1"></param>
                /// <param name="sep2"></param>
                /// <param name="pos1"></param>
                /// <param name="flag"></param>
                /// <param name="pos2"></param>
                /// <returns></returns>
                public static string getItem(string input, string sep1, string sep2, int pos1, string flag, int pos2)
                {
                    //例如: 串 1,aa;2,bb 由一级分隔符";" 分成两段, 每段又由二级分隔符"," 分开, 第一个参数sIn 是整个字串,
                    //sep1 指一级分隔符,sep2是二级分隔符.    如果对于这个串, 想查找子串中, 第1 个位置是2 时, 第二个位置的值,
                    //即, pos1 为0, flag为"2", pos2 为1, 函数返回"bb"

                    string[] arr;
                    string[] arrb;
                    arr = StringTool.MSplit(input, sep1);

                    foreach (string st in arr)
                    {
                        arrb = StringTool.MSplit(st, sep2);
                        if (pos1 >= arrb.Length || pos2 >= arrb.Length)
                            return "";
                        for (int i = 0; i < arrb.Length; i++)
                            if (arrb[pos1] == flag)
                                return arrb[pos2];
                    }

                    return "";
                }








                public static string getItem(string input, string sep1, string sep2, int row, int pos)
                {
                    if (input == "")
                        return "";

                    string[] arr = StringTool.MSplit(input, sep1);

                    if (row >= arr.Length)
                        return "";

                    string[] arrb = StringTool.MSplit(arr[row], sep2);

                    if (pos >= arrb.Length)
                        return "";

                    return arrb[pos];
                }










            }// end class MTable



            /// <summary>
            /// 获得字符串中开始和结束字符串中间得值
            /// </summary>
            /// <param name="str"></param>
            /// <param name="s">开始</param>
            /// <param name="e">结束</param>
            /// <returns></returns>
            public static string GetValue(string str, string s, string e)
            {
                Regex rg = new Regex("(?<=(" + s + "))[.\\s\\S]*?(?=(" + e + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                return rg.Match(str).Value;
            }




        }// end class stringTool








        /// <summary>
        ///  提供加解密的工具
        /// </summary>
        public class SecurityTool
        {

            public static void Encrypt(string input, ref string output, ref string key)
            {
                System.Security.Cryptography.RSACryptoServiceProvider rsa = new System.Security.Cryptography.RSACryptoServiceProvider();

                key = rsa.ToXmlString(true);
                output = ByteArrayToString(rsa.Encrypt(System.Text.Encoding.UTF8.GetBytes(input), false));
            }


            public static string Decrypt(string input, string key)
            {
                System.Security.Cryptography.RSACryptoServiceProvider rsa = new System.Security.Cryptography.RSACryptoServiceProvider();

                rsa.FromXmlString(key);
                return System.Text.Encoding.UTF8.GetString(rsa.Decrypt(ByteArrayFromString(input), false));
            }



            public static string ByteArrayToString(byte[] arr)
            {
                string ret = "";

                for (int i = 0; i < arr.Length; i++)
                    ret += arr[i].ToString("000");
                return ret;
            }


            public static byte[] ByteArrayFromString(string str)
            {
                if (str.Length % 3 != 0)
                {
                    throw new Exception("argument invalid");
                    return null;
                }

                byte[] ret = new byte[str.Length / 3];
                for (int i = 0; i < ret.Length; i++)
                {
                    ret[i] = Convert.ToByte(str.Substring(0, 3));
                    str = str.Remove(0, 3);
                }

                return ret;


            }

        }










        /// <summary>
        /// 提供完整的cookie 操作支持
        /// </summary>
        public class Cookie
        {

            /// <summary>
            /// 检查指定cookie 是否存在
            /// </summary>
            /// <param name="cookieName">cookie 名</param>
            /// <param name="request"></param>
            /// <returns>如果存在,则返回真, 否则返回假</returns>
            public static bool CheckCookie(string cookieName, System.Web.HttpRequest request)
            {
                if (request.Cookies[cookieName] == null)
                    return false;
                else
                    return true;
            }


            /// <summary>
            /// 删除一个cookie.
            /// </summary>
            /// <param name="cookieName">cookie名</param>
            /// <param name="request"></param>
            /// <param name="response"></param>
            /// <returns>如果成功删除, 返回真, 否则返回假</returns>
            public static bool RemoveCookie(string cookieName, System.Web.HttpRequest request, System.Web.HttpResponse response)
            {
                if (CheckCookie(cookieName, request) == false)
                    return true;

                System.Web.HttpCookie objCookie = request.Cookies[cookieName];
                objCookie.Expires = DateTime.Now.AddMinutes(-1);
                try
                {
                    response.AppendCookie(objCookie);
                }
                catch (Exception ex)
                {
                    return false;
                }
                return true;
            }



            /// <summary>
            /// 生成一个单键 cookie. 如果指定的名称已存在, 则删除原有cookie.
            /// </summary>
            /// <param name="cookieName">要生成的cookie名 </param>
            /// <param name="value">键值 </param>
            /// <param name="request"></param>
            /// <param name="response"></param>
            /// <param name="server"></param>
            /// <returns>生成成功则返回真, 否则返回假.</returns>
            public static bool GenCookie(string cookieName, string value, System.Web.HttpRequest request,
                System.Web.HttpResponse response, System.Web.HttpServerUtility server)
            {
                if (CheckCookie(cookieName, request) == true)
                    if (!RemoveCookie(cookieName, request, response))
                        return false;

                System.Web.HttpCookie objCookie = new System.Web.HttpCookie(cookieName);
                objCookie.Value = server.UrlEncode(value);

                try
                {
                    response.AppendCookie(objCookie);
                }
                catch (Exception ex)
                {
                    return false;
                }

                return true;
            }



            /// <summary>
            /// 生成一个多键cookie.
            /// </summary>
            /// <param name="cookieName">cookie 文件名</param>
            /// <param name="arrKey">cookie的键数组</param>
            /// <param name="arrValue">键值数组</param>
            /// <param name="request"></param>
            /// <param name="response"></param>
            /// <param name="server"></param>
            /// <returns>生成成功则返回真, 否则返回假</returns>
            public static bool GenCookie(string cookieName, string[] arrKey, string[] arrValue,
                System.Web.HttpRequest request, System.Web.HttpResponse response, System.Web.HttpServerUtility server)
            {
                if (CheckCookie(cookieName, request) == true)
                    if (!RemoveCookie(cookieName, request, response))
                        return false;
                if (arrKey.Length != arrValue.Length)
                    return false;

                System.Web.HttpCookie objCookie = new System.Web.HttpCookie(cookieName);
                for (int i = 0; i < arrKey.Length; i++)
                    objCookie.Values[arrKey[i]] = server.UrlEncode(arrValue[i]);

                try
                {
                    response.AppendCookie(objCookie);
                }
                catch (Exception ex)
                {
                    return false;
                }
                return true;
            }




            /// <summary>
            /// 读取一个单键cookie
            /// </summary>
            /// <param name="cookieName"></param>
            /// <param name="request"></param>
            /// <param name="response"></param>
            /// <param name="server"></param>
            /// <returns>返回读取到的键值</returns>
            public static string ReadCookie(string cookieName, System.Web.HttpRequest request,
                System.Web.HttpResponse response, System.Web.HttpServerUtility server)
            {
                if (CheckCookie(cookieName, request) == false)
                    return "";
                System.Web.HttpCookie objCookie = request.Cookies[cookieName];
                return server.UrlDecode(objCookie.Value);
            }




            /// <summary>
            /// 读取一个多键cookie. 如果读取失败, 返回空
            /// </summary>
            /// <param name="cookieName"></param>
            /// <param name="key">要读取的键名</param>
            /// <param name="request"></param>
            /// <param name="response"></param>
            /// <param name="server"></param>
            /// <returns>返回读到的值或空.</returns>
            public static string ReadCookie(string cookieName, string key, System.Web.HttpRequest request,
                System.Web.HttpResponse response, System.Web.HttpServerUtility server)
            {
                if (CheckCookie(cookieName, request) == false)
                    return "";
                System.Web.HttpCookie objCookie = request.Cookies[cookieName];

                try
                {
                    return server.UrlDecode(objCookie.Values[key]);
                }
                catch (Exception ex)
                {
                    return "";
                }

            }





            /// <summary>
            /// 修改一个单键cookie的值. 如果指定的cookie不存在, 返回false
            /// </summary>
            /// <param name="cookieName"></param>
            /// <param name="value"></param>
            /// <param name="request"></param>
            /// <param name="response"></param>
            /// <param name="server"></param>
            /// <returns>修改成功,返回真, 否则返回假</returns>
            public static bool ChangeKeyValue(string cookieName, string value, System.Web.HttpRequest request,
          System.Web.HttpResponse response, System.Web.HttpServerUtility server)
            {
                if (CheckCookie(cookieName, request) == false)
                    return false;
                try
                {
                    System.Web.HttpCookie objCookie = request.Cookies[cookieName];
                    objCookie.Value = server.UrlEncode(value);
                    response.AppendCookie(objCookie);
                }
                catch (Exception ex)
                {
                    return false;
                }
                return true;
            }







            /// <summary>
            /// 修改一个多键cookie的值. 
            /// </summary>
            /// <param name="cookieName"></param>
            /// <param name="key"></param>
            /// <param name="value"></param>
            /// <param name="request"></param>
            /// <param name="response"></param>
            /// <param name="server"></param>
            /// <returns></returns>
            public static bool ChangeKeyValue(string cookieName, string key, string value, System.Web.HttpRequest request,
                System.Web.HttpResponse response, System.Web.HttpServerUtility server)
            {
                if (CheckCookie(cookieName, request) == false)
                    return false;
                try
                {
                    System.Web.HttpCookie objCookie = request.Cookies[cookieName];
                    objCookie.Values[key] = server.UrlEncode(value);
                    response.AppendCookie(objCookie);
                }
                catch (Exception ex)
                {
                    return false;
                }
                return true;
            }









            /// <summary>
            /// 设置一个cookie的过期时间
            /// </summary>
            /// <param name="cookieName"></param>
            /// <param name="dtObsolete"></param>
            /// <param name="request"></param>
            /// <param name="response"></param>
            /// <returns>设置成功返回真, 否则返回假</returns>
            public static bool SetCookieExpire(string cookieName, DateTime dtObsolete,
                System.Web.HttpRequest request, System.Web.HttpResponse response)
            {
                if (CheckCookie(cookieName, request) == false)
                    return false;
                System.Web.HttpCookie objCookie = request.Cookies[cookieName];
                objCookie.Expires = dtObsolete;

                try
                {
                    response.AppendCookie(objCookie);
                }
                catch (Exception ex)
                {
                    return false;
                }
                return true;
            }






            /// <summary>
            /// 读取一个cookie的过期时间
            /// </summary>
            /// <param name="cookieName"></param>
            /// <param name="request"></param>
            /// <returns></returns>
            public static DateTime GetCookieExpire(string cookieName, System.Web.HttpRequest request)
            {
                if (CheckCookie(cookieName, request))
                    return DateTime.Now.AddMinutes(-1);

                return request.Cookies[cookieName].Expires;
            }


        }// end class cookie









        /// <summary>
        /// 操作文本文件的类
        /// </summary>
        public class TextFileTool
        {
            public static string NewLine = System.Environment.NewLine;


            /// <summary>
            /// 检查文件是否存在.
            /// </summary>
            /// <param name="fileName"></param>
            /// <returns></returns>
            public static bool CheckFile(string fileName)
            {
                if (System.IO.File.Exists(fileName))
                    return true;
                else
                    return false;
            }



            public static bool CreateTextFile(string fileName, string data, bool replace)
            {
                string sErr = "";

                if (CheckFile(fileName) == true)
                {
                    if (replace == true)
                    {
                    }
                    else
                        return false;
                }
                return true;
            }


            // Public Shared Function CreateTextFile(ByVal sName As String, ByVal sData As String, ByVal ec As System.Text.Encoding) As Boolean


            //    Dim sErrorMessage As String

            //    If CheckFile(sName) = True Then

            //        sErrorMessage = String.Format("The file ""{0}"" to create is already exists, do you want to replace it? {1} ( Suggestion: You best move the old file now, and then select yes here.)", sName, NEWLINE)

            //        If MessageBox.Show(sErrorMessage, "File already exists:", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) = DialogResult.Yes Then

            //            If DeleteFile(sName) = False Then
            //                Return False
            //                '用户要求删除原有文件, 但删除失败, 创建停止.
            //            End If
            //        Else
            //            Return True
            //            '同名文件已存在, 用户不允许覆盖, 取消函数动作.返回真.
            //        End If
            //    End If


            //    Try
            //        Dim objWriter As New IO.StreamWriter(New IO.FileStream(sName, IO.FileMode.OpenOrCreate), ec)
            //        objWriter.Write(sData)
            //        objWriter.Close()
            //    Catch ex As Exception
            //        sErrorMessage = String.Format("INIs.create: create {0} failed!", sName)
            //        TextFileTool.MakeErrorLog(sErrorMessage)
            //        Return False
            //    End Try



            //    Return True

            //End Function

        }












    
}
