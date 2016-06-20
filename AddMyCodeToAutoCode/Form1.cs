using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using MSC.WinFormControlLib;
using MSC.CommonLib;
using System.Text.RegularExpressions;
using LTP.Utility;

namespace AddMyCodeToAutoCode
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonSelectAutoDir_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxAutoDir.Text = folderDialog.SelectedPath;
            }

        }

        private void addAutoFileToDataTable(string pFolder)
        {
            _dt.Rows.Clear();
            DirectoryInfo dirInfo = new DirectoryInfo(pFolder);
            foreach (FileSystemInfo fsi in dirInfo.GetFileSystemInfos())
            {
                if (fsi is FileInfo)
                {
                    FileInfo fi = (FileInfo)fsi;
                    if (Path.GetExtension(fi.FullName) == ".cs")
                    {
                        //检查是否有自动生成代码，有才加了

                        using (StreamReader sr = new StreamReader(fi.FullName, Encoding.Default))
                        {
                            string s = sr.ReadToEnd();
                            if (s.IndexOf(this._start) < 0)
                            {
                                continue;
                            }
                        }


                        DataRow dr = _dt.NewRow();
                        dr["IsCheck"] = true;
                        dr["FileName"] = fi.Name;
                        _dt.Rows.Add(dr);
                    }
                }
            }

            this.dataGridViewAuto.DataSource = _dt;
            
        }

        private void buttonReadAutoFile_Click(object sender, EventArgs e)
        {
            string folder = this.textBoxAutoDir.Text;

            //如果是解决方案就自动添加BLL项目的文件
            if (this.checkBoxIsProjectDir.Checked)
            {
                folder += "\\BLL";
            }
            if (!Directory.Exists(folder))
            {
                DialogBox.ShowError("文件夹不存在！");
                return;
            }
            this.addAutoFileToDataTable(folder);


        }

        private DataTable _dt;
        private DataTable _dtModify;
        private void Form1_Load(object sender, EventArgs e)
        {
            _dt = new DataTable();
            DataColumn dc = new DataColumn("IsCheck", typeof(bool));
            _dt.Columns.Add(dc);
            dc = new DataColumn("FileName", typeof(string));
            _dt.Columns.Add(dc);


            _dtModify = new DataTable();
            dc = new DataColumn("FileName", typeof(string));
            _dtModify.Columns.Add(dc);

            _dtNewAdd = new DataTable();
            dc = new DataColumn("Project", typeof(string));
            _dtNewAdd.Columns.Add(dc);
            dc = new DataColumn("FileName", typeof(string));
            _dtNewAdd.Columns.Add(dc);
            this.textBoxAutoDir.Text = MSC.WinFormControlLib.CommonCode.ReadConfig("AutoDir");
            this.textBoxModifyDir.Text = MSC.WinFormControlLib.CommonCode.ReadConfig("ModifyDir");
        }

        private void 全全ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataRow dr in _dt.Rows)
            {
                dr["IsCheck"] = true;
            }
        }

        private void 选ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataRow dr in _dt.Rows)
            {
                dr["IsCheck"] = false;
            }
        }

        private void 反选ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataRow dr in _dt.Rows)
            {
                dr["IsCheck"] = !(Boolean)dr["IsCheck"];
            }
        }

        private void buttonSelectModifyDir_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxModifyDir.Text = folderDialog.SelectedPath;
            }
        }

        private void AddModifyFile(string pFolder)
        {
            this._dtModify.Rows.Clear();
            DirectoryInfo dirInfo = new DirectoryInfo(pFolder);
            foreach (FileSystemInfo fsi in dirInfo.GetFileSystemInfos())
            {
                if (fsi is FileInfo)
                {
                    FileInfo fi = (FileInfo)fsi;
                    if (Path.GetExtension(fi.FullName) == ".cs")
                    {

                        using (StreamReader sr = new StreamReader(fi.FullName, Encoding.Default))
                        {
                            string s = sr.ReadToEnd();
                            if (s.IndexOf(this._start) < 0)
                            {
                                continue;
                            }
                        }

                        DataRow dr = this._dtModify.NewRow();
                        //dr["IsCheck"] = false;
                        dr["FileName"] = fi.Name;
                        _dtModify.Rows.Add(dr);
                    }
                }
            }

            this.dataGridViewModify.DataSource = _dtModify;
        }

        private void buttonReadModifyFile_Click(object sender, EventArgs e)
        {
            string folder = this.textBoxModifyDir.Text;
            if (!Directory.Exists(folder))
            {
                DialogBox.ShowError("文件夹不存在！");
                return;
            }

            //如果是解决方案就自动添加BLL项目的文件
            if (this.checkBoxIsProjectDir.Checked)
            {
                folder += "\\BLL";
            }

            this.AddModifyFile(folder);


        }

        private string _newstart = "//autoReplaceStart";
        private string _newend = "//autoReplaceEnd";

        private string _start = "#region  自动生成代码";
        private string _end = "#endregion  自动生成代码";

        private string _start2 = "#region 自动生成代码";
        private string _end2 = "#endregion 自动生成代码";

        private DataTable _dtNewAdd;

        private void StartModify(string pAutoDir, string pModifyDir, string pProjectName)
        {
            //add file
            this.addAutoFileToDataTable(pAutoDir);


            foreach (DataRow dr in this._dt.Rows)
            {
                if ((Boolean)dr["IsCheck"])
                {

                    //选择
                    string fileName = dr["FileName"].ToString();
                    string modiFile = pModifyDir +"\\"+ fileName;
                    string autoFile = pAutoDir +"\\"+ fileName;


                    if (!File.Exists(autoFile))
                    {
                        continue;
                    }



                    //自动加入项目
                    string projectFile = pModifyDir + "\\" + pProjectName + ".csproj";
                    if (File.Exists(projectFile))
                    {
                        if (File.ReadAllText(projectFile, Encoding.Default).IndexOf(fileName) < 1)
                        {
                            //加入
                            this.AddClassFile(projectFile, fileName, "");

                        }
                    }

                    if (!File.Exists(modiFile))
                    {
                        //DialogBox.ShowError("已修改的文件，不存在了！文件：" + modiFile);
                        //直接复制
                        if (this.checkBoxIsCopy.Checked)
                        {
                            File.Copy(autoFile, modiFile);
                            //加入表
                            DataRow drNewRow = _dtNewAdd.NewRow();
                            drNewRow["Project"] = pProjectName;
                            drNewRow["FileName"] = autoFile;
                            _dtNewAdd.Rows.Add(drNewRow);
                        }
                        continue;
                    }


                    string modifyString = "";

                    //直接用自动生成的文件内容替换现有项目的内容
                    //先备份

                    string sSource = "";

                    //读取文件
                    using (StreamReader sr = new StreamReader(modiFile, Encoding.Default))
                    {
                        sSource = sr.ReadToEnd();
                        sr.Close();
                    }

                    WriteFile(modiFile + ".bak", sSource);

                    string s = "";
                    //读取生成的文件内容
                    using (StreamReader sr = new StreamReader(autoFile, Encoding.Default))
                    {
                        modifyString = StringTool.GetValue(sr.ReadToEnd(), _start, _end);
                        sr.Close();
                    }
                    Regex renew  = new Regex("(?<=(" + this._newstart + "))[.\\s\\S]*?(?=(" + this._newend + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                    Regex re = new Regex("(?<=(" + this._start + "))[.\\s\\S]*?(?=(" + this._end + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                    Regex re2 = new Regex("(?<=(" + this._start2 + "))[.\\s\\S]*?(?=(" + this._end2 + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                    if (renew.IsMatch(sSource))
                    {
                        using (StreamReader sr = new StreamReader(autoFile, Encoding.Default))
                        {
                            modifyString = StringTool.GetValue(sr.ReadToEnd(), _newstart, _newend);
                            sr.Close();
                        }
                        sSource = renew.Replace(sSource, modifyString);
                    }

                    else if (re.IsMatch(sSource))
                    {
                        sSource = re.Replace(sSource, modifyString);
                    }
                    else
                    {
                        //一个空格
                        sSource = re2.Replace(sSource, modifyString);
                    }
                    //modifyString 是修改后的新内容
                    //modifyString = this._start + "\n\r" + modifyString+"\n\r"+this._end;




                    //sSource =  
                    WriteFile(modiFile, sSource);
                }

            }

            
        }


        private void buttonStart_Click(object sender, EventArgs e)
        {

            //SAVE 
            MSC.WinFormControlLib.CommonCode.WriteConfig("AutoDir", this.textBoxAutoDir.Text);
            MSC.WinFormControlLib.CommonCode.WriteConfig("ModifyDir", this.textBoxModifyDir.Text);

            this._dtNewAdd.Rows.Clear();

            string bll = "BLL",  dal = "DAL", model = "Model";
            if (checkBoxIsProjectDir.Checked)
            {
                string autoProjectDir = this.textBoxAutoDir.Text ;
                string modifyProjectDir = this.textBoxModifyDir.Text ;
                if (checkBoxIsCopyBll.Checked)
                {
                    string autoDir = autoProjectDir +"\\"+ bll ;
                    string modifyDir = modifyProjectDir + "\\" + bll;

                    this.addAutoFileToDataTable(autoDir);
                    this.AddModifyFile(modifyDir);


                    this.StartModify(autoDir, modifyDir, bll);
                }
                if (checkBoxIsDal. Checked)
                {
                    string autoDir = autoProjectDir + "\\" + dal;
                    string modifyDir = modifyProjectDir + "\\" + dal;

                    this.addAutoFileToDataTable(autoDir);
                    this.AddModifyFile(modifyDir);

                    this.StartModify(autoDir, modifyDir, dal);
                }
                if (checkBoxIsModel. Checked)
                {
                    string autoDir = autoProjectDir + "\\" + model;
                    string modifyDir = modifyProjectDir + "\\" + model;

                    this.addAutoFileToDataTable(autoDir);
                    this.AddModifyFile(modifyDir);

                    this.StartModify(autoDir, modifyDir, model);
                }
            }
            else
            {
                this.StartModify(this.textBoxAutoDir.Text , this.textBoxModifyDir.Text, "当前项目");
            }

            //SAVE 
            MSC.WinFormControlLib.CommonCode.WriteConfig("AutoDir", this.textBoxAutoDir.Text);
            MSC.WinFormControlLib.CommonCode.WriteConfig("ModifyDir", this.textBoxModifyDir.Text);

            this.dataGridViewAuto.DataSource = this._dtNewAdd;

            DialogBox.ShowInfo("完成， 左表是新文件,请另行加入项目！");





            

        }


        private VSProject vsp = new VSProject();

        private void AddClassFile(string ProjectFile, string classFileName, string ProType)
        {
            if (File.Exists(ProjectFile))
            {
                switch (ProType)
                {
                    case "2003":
                        this.vsp.AddClass2003(ProjectFile, classFileName);
                        return;

                    case "2005":
                        this.vsp.AddClass2005(ProjectFile, classFileName);
                        return;
                }
                this.vsp.AddClass(ProjectFile, classFileName);
            }
        }


        private void WriteFile(string Filename, string strCode)
        {
            StreamWriter writer = new StreamWriter(Filename, false, Encoding.Default);
            writer.Write(strCode);
            writer.Flush();
            writer.Close();
        }
    }
}
