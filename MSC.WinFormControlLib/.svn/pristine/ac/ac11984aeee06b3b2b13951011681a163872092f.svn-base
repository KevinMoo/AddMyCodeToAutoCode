using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace MSC.WinFormControlLib
{
    public class DialogBox
    {
        public static void ShowError(string pMsg)
        {
            MessageBox.Show(pMsg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Hand);
        }

        public static void ShowInfo(string pMsg)
        {
            MessageBox.Show(pMsg, "信息", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        public static DialogResult ShowQuestion(string pMsg)
        {
            return MessageBox.Show(pMsg, "系统询问", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        public static void ShowWarn(string pMsg)
        {
            MessageBox.Show(pMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }

    public class CommonCode
    {
        public static string ReadConfig(string pKey)
        {
            return ConfigurationManager.AppSettings[pKey];
        }

        public static void WriteConfig(string pKey, string pNewValue)
        {
            Configuration configuration =ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            configuration.AppSettings.Settings.Remove(pKey);
            configuration.AppSettings.Settings.Add(pKey, pNewValue);
            configuration.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }
    }



    
}
