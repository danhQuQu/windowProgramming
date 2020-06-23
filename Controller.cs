using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;

//using Manage_Document.Views;

namespace Manage_Document
{
    class Controller
    {
        sqlexpressEntities2 DE = new sqlexpressEntities2();
        public void addDoc(string filename, string link)
        {

            Document dcm = new Document() { Ten = filename, Link = link, LinkImage = "D:/Window/Manage Document/Manage Document/bin/Debug/0.jpg", IsRead = 0 };
            DE.Documents.Add(dcm);
            DE.SaveChanges();

        }
        public void Edit(int i, string s1, string s2, int s3)
        {

            Document dcm = DE.Documents.Find(i);
            dcm.Ten = s1;
            dcm.LinkImage = s2;
            dcm.IsRead = s3;
            DE.SaveChanges();
        }
        public void delete(int ID1)
        {
            Document dcm = DE.Documents.Where(b => b.ID == ID1).FirstOrDefault();
            DE.Documents.Remove(dcm);
            DE.SaveChanges();
        }
        public List<Document> LoadFile(int i)
        {

            var result = from c in DE.Documents where c.ID == i select c;
            return result.ToList();

        }
        public SqlConnection kn = new SqlConnection();
        public void kn_csdl()
        {
            string chuoikn = @"Data Source=localhost;Initial Catalog=sqlexpress;Integrated Security=True";

            kn.ConnectionString = chuoikn;
            kn.Open();
        }
        public void dongketnoi()
        {
            if (kn.State == ConnectionState.Open)
                kn.Close();
        }
        public DataTable bangdulieu = new DataTable();
        public DataTable laybang(string caulenh)
        {
            try
            {
                kn_csdl();
                SqlDataAdapter Adapter = new SqlDataAdapter(caulenh, kn);
                DataSet ds = new DataSet();

                Adapter.Fill(bangdulieu);
            }
            catch (System.Exception)
            {
                bangdulieu = null;
            }
            finally
            {
                dongketnoi();
            }
            return bangdulieu;
        }


    }
}
