using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Manage_Document
{
    public partial class frmHome : Form
    { 
        frmDetails detail;
        sqlexpressEntities2 DE = new sqlexpressEntities2();
        private string filename;
        private string link;
        private int index=-1;

        public frmHome()
        {
            InitializeComponent();
   

        }
        private void btSearch_Click(object sender, EventArgs e)
        {
            LoadGridByWord();
        }

        private void textSearch_TextChanged(object sender, EventArgs e)
        {
            //find the name of document

        }
        public void LoadGridByWord()
        {
            Controller kn = new Controller();

            dataRead.DataSource = kn.laybang("select * from Document where Ten like '%" + textSearch.Text + "%'");

        }


        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { ValidateNames = true, Multiselect = false, Filter = "Word 97 - 2003 | *doc |Word Document | *.docx" })
            {
                Controller cl = new Controller();
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    dataLib.Rows.Add();
                    index++;
                    string l = ofd.FileName;
                    string[] a = ofd.FileName.Split('\\');
                    string fname = a[a.Length - 1];
                    //int index = dataLib.Rows.Add();
                    filename = fname;
                    link = l;
                    cl.addDoc(fname,l);
                    dataLib.Rows[index].Cells[0].Value = fname;
                    dataLib.Rows[index].Cells[1].Value = l;
                    
                    dataLib.Rows[index].Cells[3].Value = "D:/Window/Manage Document/Manage Document/bin/Debug/0.jpg";
                    dataLib.Rows[index].Cells[4].Value = 0;

                    //dataLib.DataSource = filename.ToList();
                    object readOnly = false;
                    object visible = true;
                    object save = false;
                    object fileName = ofd.FileName;
                    object newTemplate = false;
                    object docType = 0;
                    object missing = Type.Missing;

                    if (dataLib.Rows.Count > 1)
                    {
                        dataLib.Rows.Clear();
                        index = -1;
                    }
                    var result = from c in DE.Documents select new { Name = c.Ten, link = c.Link, ID = c.ID, lImage = c.LinkImage, isread = c.IsRead };
                    var data = result.ToList();
                    for (int i = 0; i < data.Count; i++)
                    {
                        dataLib.Rows.Add();
                        index++;
                        dataLib.Rows[i].Cells[0].Value = data[i].Name;
                        dataLib.Rows[i].Cells[1].Value = data[i].link;
                        dataLib.Rows[i].Cells[2].Value = data[i].ID;
                        dataLib.Rows[index].Cells[3].Value = data[i].lImage;
                        dataLib.Rows[index].Cells[4].Value = data[i].isread;
                    }
                }
            }
        }

        private void dataLib_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex > index)
            {
                MessageBox.Show("Thêm dữ liệu vào trước");
                return;
            }

            detail = new frmDetails(dataLib.Rows[e.RowIndex].Cells[0].Value.ToString(), dataLib.Rows[e.RowIndex].Cells[1].Value.ToString(), e.RowIndex, dataLib.Rows[e.RowIndex].Cells[2].Value.ToString(), dataLib.Rows[e.RowIndex].Cells[3].Value.ToString(), 0);
            int id = Convert.ToInt32(dataLib.Rows[e.RowIndex].Cells[2].Value);
            string s = string.Format(dataLib.Rows[e.RowIndex].Cells[3].Value.ToString());
            using (FileStream file = new FileStream(s, FileMode.Open))
            {
                detail.picNote.Image = System.Drawing.Image.FromStream(file);
                file.Close();
            }
            detail.Show();

        }
        
        private void btnLoad_Click(object sender, EventArgs e)
        {
            var rs = from c in DE.Documents where c.IsRead == 1 select c;
            dataRead.DataSource = rs.ToList();

        }

        private void frmHome_Load(object sender, EventArgs e)
        {
            var result = from c in DE.Documents select new { Name = c.Ten, link = c.Link, ID = c.ID , lImage =c.LinkImage , isread=c.IsRead};
            var data = result.ToList();
            for(int i = 0; i < data.Count; i++)
            {
                dataLib.Rows.Add();
                index++;
                dataLib.Rows[index].Cells[0].Value = data[i].Name;
                dataLib.Rows[index].Cells[1].Value = data[i].link;
                dataLib.Rows[index].Cells[2].Value = data[i].ID;
                dataLib.Rows[index].Cells[3].Value = data[i].lImage;
                dataLib.Rows[index].Cells[4].Value = data[i].isread;

            }
        }
       
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            if(dataLib.Rows.Count > 1)
            {
                dataLib.Rows.Clear();
                index = -1;
            }
            var result = from c in DE.Documents select new { Name = c.Ten, link = c.Link, ID = c.ID , lImage = c.LinkImage, isread = c.IsRead };
            var data = result.ToList();
            for (int i = 0; i < data.Count; i++)
            {
                    dataLib.Rows.Add();
                    index++;
                    dataLib.Rows[i].Cells[0].Value = data[i].Name;
                    dataLib.Rows[i].Cells[1].Value = data[i].link;
                    dataLib.Rows[i].Cells[2].Value = data[i].ID;
                    dataLib.Rows[index].Cells[3].Value = data[i].lImage;
                    dataLib.Rows[index].Cells[4].Value = data[i].isread;
            }

        }

        private void dataRead_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > index)
            {
                MessageBox.Show("Thêm dữ liệu vào trước");
                return;
            }

            detail = new frmDetails(dataRead.Rows[e.RowIndex].Cells[1].Value.ToString(), dataRead.Rows[e.RowIndex].Cells[3].Value.ToString(), e.RowIndex, dataRead.Rows[e.RowIndex].Cells[0].Value.ToString(), dataRead.Rows[e.RowIndex].Cells[2].Value.ToString(), 0);
            int id = Convert.ToInt32(dataRead.Rows[e.RowIndex].Cells[0].Value);
            string s = string.Format(dataRead.Rows[e.RowIndex].Cells[2].Value.ToString());
            using (FileStream file = new FileStream(s, FileMode.Open))
            {
                detail.picNote.Image = System.Drawing.Image.FromStream(file);
                file.Close();
            }
            detail.Show();
        }
    }
}
