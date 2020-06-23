using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace Manage_Document
{
    public partial class frmDetails : Form
    {
        sqlexpressEntities2 DE = new sqlexpressEntities2();
        private DrawItem drawNote;
        private string filename;
        private string link;
        private string ID;
        private int rowIndex;
        private string note;
        private string LinkImage;
        private int isRead;
        frmHome home;
        private Bitmap bitmapNote;
        private Bitmap bmOld;
        private Bitmap bmNew;
        //public DateTime(int year, int month, int day, int hour, int minute, int second);


        public string Note
        {
            get { return note; }
            set { note = value; }
        }
        public frmDetails(string fname, string l, int rIndex, string id, string lImage, int isread)
        {
            filename = fname;
            link = l;
            rowIndex = rIndex;
            ID = id;
            LinkImage = lImage;
            isRead = isread;
            InitializeComponent();
            bitmapNote = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());
            bmNew = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());
            bmOld = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());

        }


        private void picNote_MouseUp(object sender, MouseEventArgs e)
        {
            drawNote.isDraw = false;
        }

        private void picNote_MouseDown(object sender, MouseEventArgs e)
        {
            drawNote.isDraw = true;
            drawNote.X = e.X;
            drawNote.Y = e.Y;

        }

        private void picNote_MouseMove(object sender, MouseEventArgs e)
        {
            if (drawNote.isDraw)
            {
                Graphics G = this.picNote.CreateGraphics();
                G.DrawLine(drawNote.pen, drawNote.X, drawNote.Y, e.X, e.Y);
                using (Graphics gr = Graphics.FromImage(bitmapNote))
                {
                    gr.DrawLine(drawNote.pen, drawNote.X, drawNote.Y, e.X, e.Y);
                }
                drawNote.X = e.X;
                drawNote.Y = e.Y;

            }
        }

        private void frmDetails_Load(object sender, EventArgs e)
        {

            drawNote = new DrawItem();
            txtboxName.Text = filename;
            txtboxLink.Text = link;
            txtBoxID.Text = ID;



        }

        public List<Document> loadFile(int i)
        {
            var result = from c in DE.Documents where c.ID == i select c;
            return result.ToList();
        }
        private void btnSave_Click(object sender, EventArgs e)
        {   
            Controller cl = new Controller();
            if (!LinkImage.Contains("0.jpg"))
            {
                string s = string.Format("D:/Window/Manage Document/Manage Document/bin/Debug/{0}.jpg", ID);
                string s2 = string.Format("D:/Window/Manage Document/Manage Document/bin/Debug/{0}(1).jpg", ID);
                if (File.Exists(s))
                {
                    if (File.Exists(s2))
                    {
                        File.Delete(s2);
                        bitmapNote.Save(string.Format("{0}(1).jpg", ID));

                        bitmapNote.Dispose();
                        bitmapNote = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());
                    }
                    else
                    {
                        bitmapNote.Save(string.Format("{0}(1).jpg", ID));

                        bitmapNote.Dispose();
                        bitmapNote = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());
                    }
                }
                else
                {
                    bitmapNote.Save(string.Format("{0}.jpg", ID));
                    bitmapNote.Dispose();
                    bitmapNote = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());
                }
                using (FileStream fst = new FileStream(s2, FileMode.Open))
                {
                    bmNew = (Bitmap)System.Drawing.Image.FromStream(fst);
                    fst.Close();

                }
                using (FileStream fst2 = new FileStream(s, FileMode.Open))
                {
                    bmOld = (Bitmap)System.Drawing.Image.FromStream(fst2);
                    fst2.Close();

                }
                Graphics g = Graphics.FromImage(bmOld);
                g.DrawImage(bmNew, 0, 0, bmOld.Size.Width, bmOld.Size.Height);
                g.Dispose();
                if (File.Exists(s))
                {
                    File.Delete(s);
                    bmOld.Save(string.Format("{0}.jpg", ID));
                    bitmapNote = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());
                }
            }
            else
            {
                string s = string.Format("D:/Window/Manage Document/Manage Document/bin/Debug/0.jpg", ID);
                string s2 = string.Format("D:/Window/Manage Document/Manage Document/bin/Debug/{0}.jpg", ID);
                if (File.Exists(s))
                {
                    if (File.Exists(s2))
                    {
                        File.Delete(s2);
                        bitmapNote.Save(string.Format("{0}.jpg", ID));
                        bitmapNote = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());
                        bitmapNote.Dispose();
                    }
                    else
                    {
                        bitmapNote.Save(string.Format("{0}.jpg", ID));
                        bitmapNote = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());
                        bitmapNote.Dispose();
                    }
                }
                else
                {
                    bitmapNote.Save(string.Format("{0}.jpg", ID));
                    bitmapNote.Dispose();
                }
                using (FileStream fst = new FileStream(s2, FileMode.Open))
                {
                    bmNew = (Bitmap)System.Drawing.Image.FromStream(fst);
                    fst.Close();

                }
                using (FileStream fst2 = new FileStream(s, FileMode.Open))
                {
                    bmOld = (Bitmap)System.Drawing.Image.FromStream(fst2);
                    fst2.Close();

                }
                Graphics g = Graphics.FromImage(bmOld);
                g.DrawImage(bmNew, 0, 0, bmOld.Size.Width, bmOld.Size.Height);
                g.Dispose();
                if (File.Exists(s2))
                {
                    File.Delete(s2);
                    bmNew.Save(string.Format("{0}.jpg", ID));
                    bitmapNote = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());
                }
            }
            string s3 = string.Format("D:/Window/Manage Document/Manage Document/bin/Debug/{0}.jpg", ID);
            cl.Edit(Convert.ToInt32(ID), txtboxName.Text, s3, isRead);
            MessageBox.Show("Đã Update Thành công!");
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            string path = txtboxLink.Text;
            object readOnly = false;
            object visible = true;
            object save = false;
            object fileName = path;
            object newTemplate = false;
            object docType = 0;
            object missing = Type.Missing;
            Controller cl = new Controller();

            Microsoft.Office.Interop.Word._Application application = new Microsoft.Office.Interop.Word.Application();
            {
                Visible = true;
            }
            application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing,
                       ref missing, ref missing, ref missing, ref missing,
                       ref visible, ref missing, ref missing, ref missing, ref missing);
            home = new frmHome();
            int id1 = Convert.ToInt32(ID);
            cl.Edit(id1, txtboxName.Text, LinkImage, 1);
            //int s = Convert.ToInt32(home.dataLib.SelectedCells[0].OwningRow.Cells["ID"].Value);
            //var result = from c in DE.Documents where c.ID == s select c;
            //home.dataRead.DataSource = result.ToList();

        }         
        private void btnDelete_Click(object sender, EventArgs e)
        {
            Controller cl = new Controller();
            DialogResult result = MessageBox.Show("Do you want to delete this file?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            switch (result)
            { 
                case DialogResult.Yes:
                    int ID1 = Convert.ToInt32(ID);
                    cl.delete(ID1);
                    MessageBox.Show("Đã xóa Thành công!");
                    this.Close();
                    break;
                case DialogResult.No:
                    break;
            }
        }

        public class DrawItem
        {
            public int X { set; get; }
            public int Y { set; get; }
            public Color color { set; get; }
            public Pen pen { set; get; }
            public bool isDraw { set; get; }
            public DrawItem()
            {
                color = Color.Black;
                pen = new Pen(this.color, 2);
                isDraw = false;

            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            this.picNote.CreateGraphics().Clear(Color.White);

            Controller cl = new Controller();
            string s3 = string.Format("D:/Window/Manage Document/Manage Document/bin/Debug/{0}.jpg", ID);
            if (File.Exists(s3))
            {
                File.Delete(s3);
                bitmapNote.Save(string.Format("{0}.jpg", ID));
                bitmapNote.Dispose();
                cl.Edit(Convert.ToInt32(ID), txtboxName.Text, s3, isRead);
                MessageBox.Show("Clear note thanh cong ");
                bitmapNote = new Bitmap(this.picNote.ClientSize.Width, this.picNote.ClientSize.Height, this.picNote.CreateGraphics());
            }
        }
    }
}

