namespace MSC.WinFormControlLib
{
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Windows.Forms;

    public class frmPreviewImage : frmBase
    {
        private ContextMenuStrip cms;
        private IContainer components;
        private Image img;
        //private Log log;
        private PictureBox picMain;
        private PictureBoxSizeMode[] picMode;
        private ToolTip ttInfo;
        private int x;
        private int y;
        private ToolStripMenuItem 拉伸ToolStripMenuItem;
        private ToolStripMenuItem 铺满ToolStripMenuItem;
        private ToolStripMenuItem 原始大小ToolStripMenuItem;

        public frmPreviewImage()
        {
            PictureBoxSizeMode[] modeArray = new PictureBoxSizeMode[5];
            modeArray[1] = PictureBoxSizeMode.CenterImage;
            modeArray[2] = PictureBoxSizeMode.StretchImage;
            modeArray[3] = PictureBoxSizeMode.Zoom;
            modeArray[4] = PictureBoxSizeMode.AutoSize;
            this.picMode = modeArray;
            this.InitializeComponent();
        }

        public frmPreviewImage(string tag)
        {
            PictureBoxSizeMode[] modeArray = new PictureBoxSizeMode[5];
            modeArray[1] = PictureBoxSizeMode.CenterImage;
            modeArray[2] = PictureBoxSizeMode.StretchImage;
            modeArray[3] = PictureBoxSizeMode.Zoom;
            modeArray[4] = PictureBoxSizeMode.AutoSize;
            this.picMode = modeArray;
            this.InitializeComponent();
            try
            {
                if ((tag != null) && (tag.Length > 0))
                {
                    this.showPic(tag);
                }
            }
            catch (Exception exception)
            {
                //this.log = new Log("MainForm(string[] tag)", exception.Message);
            }
        }

        private void centerScreen(PictureBox pic)
        {
            int num = base.Width - 8;
            int num2 = base.Height - 0x19;
            int width = pic.Width;
            int height = pic.Height;
            pic.Location = new Point((num / 2) - (width / 2), (num2 / 2) - (height / 2));
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
            this.components = new Container();
            ComponentResourceManager resources = new ComponentResourceManager(typeof(frmPreviewImage));
            this.picMain = new PictureBox();
            this.cms = new ContextMenuStrip(this.components);
            this.拉伸ToolStripMenuItem = new ToolStripMenuItem();
            this.铺满ToolStripMenuItem = new ToolStripMenuItem();
            this.原始大小ToolStripMenuItem = new ToolStripMenuItem();
            this.ttInfo = new ToolTip(this.components);
            ((ISupportInitialize) this.picMain).BeginInit();
            this.cms.SuspendLayout();
            base.SuspendLayout();
            this.picMain.Anchor = AnchorStyles.None;
            this.picMain.BackColor = Color.White;
            this.picMain.Location = new Point(0, 0);
            this.picMain.Margin = new Padding(0);
            this.picMain.Name = "picMain";
            this.picMain.Size = new Size(680, 0x1a3);
            this.picMain.SizeMode = PictureBoxSizeMode.Zoom;
            this.picMain.TabIndex = 0;
            this.picMain.TabStop = false;
            this.picMain.MouseMove += new MouseEventHandler(this.PicMainMouseMove);
            this.picMain.MouseDown += new MouseEventHandler(this.PicMainMouseDown);
            this.picMain.MouseUp += new MouseEventHandler(this.PicMainMouseUp);
            this.picMain.DragEnter += new DragEventHandler(this.PicMainDragEnter);
            this.cms.Items.AddRange(new ToolStripItem[] { this.拉伸ToolStripMenuItem, this.铺满ToolStripMenuItem, this.原始大小ToolStripMenuItem });
            this.cms.Name = "cms";
            this.cms.Size = new Size(0x77, 70);
            this.拉伸ToolStripMenuItem.Name = "拉伸ToolStripMenuItem";
            this.拉伸ToolStripMenuItem.Size = new Size(0x76, 0x16);
            this.拉伸ToolStripMenuItem.Tag = "2";
            this.拉伸ToolStripMenuItem.Text = "拉伸";
            this.拉伸ToolStripMenuItem.Click += new EventHandler(this.TsmiSizeModeClick);
            this.铺满ToolStripMenuItem.Name = "铺满ToolStripMenuItem";
            this.铺满ToolStripMenuItem.Size = new Size(0x76, 0x16);
            this.铺满ToolStripMenuItem.Tag = "3";
            this.铺满ToolStripMenuItem.Text = "铺满";
            this.铺满ToolStripMenuItem.Click += new EventHandler(this.TsmiSizeModeClick);
            this.原始大小ToolStripMenuItem.Name = "原始大小ToolStripMenuItem";
            this.原始大小ToolStripMenuItem.Size = new Size(0x76, 0x16);
            this.原始大小ToolStripMenuItem.Tag = "4";
            this.原始大小ToolStripMenuItem.Text = "原始大小";
            this.原始大小ToolStripMenuItem.Click += new EventHandler(this.TsmiSizeModeClick);
            this.AllowDrop = true;
            base.AutoScaleDimensions = new SizeF(6f, 12f);
            base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = Color.DimGray;
            base.ClientSize = new Size(680, 0x1a3);
            this.ContextMenuStrip = this.cms;
            base.Controls.Add(this.picMain);
            base.Icon = (Icon) resources.GetObject("$this.Icon");
            base.Name = "frmPreviewImage";
            base.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "图片浏览器";
            base.StyleChanged += new EventHandler(this.MainFormSizeChanged);
            base.Load += new EventHandler(this.MainFormLoad);
            base.SizeChanged += new EventHandler(this.MainFormSizeChanged);
            base.DragEnter += new DragEventHandler(this.PicMainDragEnter);
            base.MouseMove += new MouseEventHandler(this.MainFormMouseMove);
            ((ISupportInitialize) this.picMain).EndInit();
            this.cms.ResumeLayout(false);
            base.ResumeLayout(false);
        }

        private void MainFormLoad(object sender, EventArgs e)
        {
            //new Log(Application.ExecutablePath);
        }

        private void MainFormMouseMove(object sender, MouseEventArgs e)
        {
            this.Text = string.Concat(new object[] { "木木图片浏览器 - size(w: ", this.picMain.Size.Width, " h: ", this.picMain.Size.Height, ") - point(x: ", e.X, ", y: ", e.Y, ")" });
        }

        private void MainFormSizeChanged(object sender, EventArgs e)
        {
            this.centerScreen(this.picMain);
        }

        private void PicMainDragEnter(object sender, DragEventArgs e)
        {
            this.showpicture(e);
            this.centerScreen(this.picMain);
        }

        private void PicMainMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.x = e.X;
                this.y = e.Y;
                this.picMain.Cursor = Cursors.SizeAll;
            }
        }

        private void PicMainMouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.picMain.Location = new Point(this.picMain.Location.X + (e.X - this.x), this.picMain.Location.Y + (e.Y - this.y));
            }
            this.Text = string.Concat(new object[] { "木木图片浏览器 - size(w: ", this.picMain.Size.Width, " h: ", this.picMain.Size.Height, ") - point(x: ", e.X, ", y: ", e.Y, ")" });
        }

        private void PicMainMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.picMain.Cursor = Cursors.Arrow;
            }
        }

        private void picSizeMode(object sender, PictureBox pic)
        {
            try
            {
                if (sender != null)
                {
                    ToolStripMenuItem item = (ToolStripMenuItem) sender;
                    if (item != null)
                    {
                        this.拉伸ToolStripMenuItem.Checked = false;
                        this.铺满ToolStripMenuItem.Checked = false;
                        this.原始大小ToolStripMenuItem.Checked = false;
                        item.Checked = true;
                        object tag = item.Tag;
                        int index = 0;
                        if (tag != null)
                        {
                            index = Convert.ToInt32(tag);
                        }
                        if ((index == 2) || (index == 3))
                        {
                            pic.Size = new Size(base.Width - 8, base.Height - 0x19);
                        }
                        pic.SizeMode = this.picMode[index];
                    }
                }
            }
            catch (Exception exception)
            {
                //this.log = new Log("picSizeMode", exception.Message);
            }
        }

        private void showPic(string path)
        {
            Bitmap bitmap = new Bitmap(path);
            if ((bitmap.Width > (base.Width - 8)) || (bitmap.Height > (base.Height - 0x19)))
            {
                this.picSizeMode(this.铺满ToolStripMenuItem, this.picMain);
            }
            else
            {
                this.picSizeMode(this.原始大小ToolStripMenuItem, this.picMain);
            }
            this.img = bitmap;
            this.picMain.Image = this.img;
        }

        private void showpicture(DragEventArgs e)
        {
            try
            {
                string[] strArray;
                e.Effect = DragDropEffects.Copy;
                object data = e.Data.GetData(DataFormats.FileDrop, true);
                if (data != null)
                {
                    strArray = (string[]) data;
                }
                else
                {
                    strArray = new string[0];
                }
                if ((strArray != null) && (strArray.Length > 0))
                {
                    this.showPic(strArray[0]);
                }
            }
            catch (Exception exception)
            {
                //this.log = new Log("showpicture", exception.Message);
            }
        }

        private void TsmiSizeModeClick(object sender, EventArgs e)
        {
            this.picSizeMode(sender, this.picMain);
            this.centerScreen(this.picMain);
        }
    }
}

