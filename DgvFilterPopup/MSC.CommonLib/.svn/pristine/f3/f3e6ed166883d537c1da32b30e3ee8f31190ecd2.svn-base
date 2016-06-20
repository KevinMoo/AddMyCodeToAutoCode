using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;
using System.Runtime.InteropServices;

namespace MSC.CommonLib
{
    public class NetCardOpeartor
    {

        public NetCardOpeartor(string pNetCardName)
        {
            this._netCardName = pNetCardName;
        }

        private string _netCardName;



        /// <summary>
        /// 网卡列表
        /// </summary>
        public static List<string> NetWorkList()
        {
            string manage = "SELECT * From Win32_NetworkAdapter";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(manage);
            ManagementObjectCollection collection = searcher.Get();
            List<string> netWorkList = new List<string>();

            foreach (ManagementObject obj in collection)
            {
                netWorkList.Add(obj["Name"].ToString());
            }

            return netWorkList;
            //this.cmbNetWork.DataSource = netWorkList;

        }

        /// <summary>
        /// 禁用网卡
        /// </summary>5
        /// <param name="netWorkName">网卡名</param>
        /// <returns></returns>
        public bool DisableNetWork(ManagementObject network)
        {
            try
            {
                network.InvokeMethod("Disable", null);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 启用网卡
        /// </summary>
        /// <param name="netWorkName">网卡名</param>
        /// <returns></returns>
        public bool EnableNetWork(ManagementObject network)
        {
            try
            {
                network.InvokeMethod("Enable", null);
                return true;
            }
            catch
            {
                return false;
            }

        }

        /// <summary>
        /// 网卡状态
        /// </summary>
        /// <param name="netWorkName">网卡名</param>
        /// <returns></returns>
        public bool NetWorkState()
        {
            string netState = "SELECT * From Win32_NetworkAdapter";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(netState);
            ManagementObjectCollection collection = searcher.Get();
            foreach (ManagementObject manage in collection)
            {
                if (manage["Name"].ToString() == this._netCardName)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 得到指定网卡
        /// </summary>
        /// <param name="networkname">网卡名字</param>
        /// <returns></returns>
        public ManagementObject NetWork()
        {
            string netState = "SELECT * From Win32_NetworkAdapter";

            ManagementObjectSearcher searcher = new ManagementObjectSearcher(netState);
            ManagementObjectCollection collection = searcher.Get();

            foreach (ManagementObject manage in collection)
            {
                if (manage["Name"].ToString() == this._netCardName)
                {
                    return manage;
                }
            }


            return null;
        }



        [DllImport("wininet.dll")]
        private extern static bool InternetCheckConnection(String url, int flag, int ReservedValue);
        /// <summary>
        /// 第一步.检测外网的一个网站，如www.baidu.com
        /// </summary>
        /// <returns></returns>
        public static bool GetExtranet(string pURL)
        {
            bool extranet = false;
            try
            {
                if (InternetCheckConnection(pURL, 1, 0).Equals(false))
                {
                    extranet = false;
                }
                else
                {
                    extranet = true;
                }
            }
            catch (Exception e)
            {
                e.ToString();
            }
            return extranet;
        }

        //禁用 SetNetworkAdapter(False)
        //启用 SetNetworkAdapter(True)
        //添加引用system32\shell32.dll

/// <summary>
/// 通过网络连接名，修改网络启用，停用状态
/// </summary>
/// <param name="status">启用true,停用false</param>
/// <param name="pNetworkConnection">网络连接名</param>
/// <returns></returns>
        public  static bool SetNetworkAdapterByShell(bool status,string pNetworkConnection)
        {
            const string discVerb = "停用(&B)"; // "停用(&B)";
            const string connVerb = "启用(&A)"; // "启用(&A)";
            const string network = "网络连接"; //"网络连接";
            //const string networkConnection = this._netCardName; // "本地连接"

            string sVerb = null;

            if (status)
            {
                sVerb = connVerb;
            }
            else
            {
                sVerb = discVerb;
            }

            Shell32.Shell sh = new Shell32.Shell();
            Shell32.Folder folder = sh.NameSpace(Shell32.ShellSpecialFolderConstants.ssfCONTROLS);

            try
            {
                //进入控制面板的所有选项
                foreach (Shell32.FolderItem myItem in folder.Items())
                {
                    //进入网络连接
                    if (myItem.Name == network)
                    {
                        Shell32.Folder fd = (Shell32.Folder)myItem.GetFolder;
                        foreach (Shell32.FolderItem fi in fd.Items())
                        {
                            //找到本地连接
                            if ((fi.Name == pNetworkConnection))
                            {
                                //找本地连接的所有右键功能菜单
                                foreach (Shell32.FolderItemVerb Fib in fi.Verbs())
                                {
                                    if (Fib.Name == sVerb)
                                    {
                                        Fib.DoIt();
                                        return true;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            return true;
        }

    }
}
