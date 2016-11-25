using MoreLinq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADOExcel
{   
    //Test Mod Elisardo
    class FileReader<T>
    {
        private static List<T> list = new List<T>();

        public FileReader() { }

        public static OleDbConnection OpenExcel(string file = null)
        {
            string fileExtension = Path.GetExtension(file);
            string connectionString = "";

            try
            {
                if (fileExtension == ".xls")
                {
                    connectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};" + "Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'", file);
                }
                else if (fileExtension == ".xlsx")
                {
                    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1;\"", file);
                }

                OleDbConnection conn = new OleDbConnection(connectionString);
                conn.Open();
                return conn;
            }
            catch (Exception ex)
            {
                throw new ArgumentException("There was a problem trying to open the file");
            }
        }

        public static DataSet SelectWorkingSheet(OleDbConnection conn)
        {
            try
            {
                OleDbCommand command = new OleDbCommand("Select * from [Sheet1$]", conn);
                OleDbDataAdapter adpt = new OleDbDataAdapter(command);
                DataSet ds1 = new DataSet();
                adpt.Fill(ds1);
                OleDbDataReader reader = command.ExecuteReader();
                conn.Close();
                return ds1;
            }
            catch (Exception ex)
            {

                throw new ArgumentException("Verify if the Sheet name in the file is [Sheet1]");
            }
        }

        public static List<T> ReadExcel(string file)
        {
            list.Clear();
            var conn = OpenExcel(file);
            var ds1 = SelectWorkingSheet(conn);

            var F1 = ds1.Tables["Table"].Columns["F1"];
            var F2 = ds1.Tables["Table"].Columns["F2"];
            var F3 = ds1.Tables["Table"].Columns["F3"];
            var F4 = ds1.Tables["Table"].Columns["F4"];

            if (F1 == null || F2 == null || F3 == null || F4 == null)
            {
                throw new ArgumentException("There are columns missing in your document please verify");
            }

            string excelHeader = "<table cellspacing='0' rules='all' border='1'  style='border-collapse:collapse;'><tbody><tr><th scope='col'>Order #</th><th scope='col'>Part #</th><th scope='col'>IMEI</th><th scope='col'>Model</th></tbody></table>";

            if (ds1.Tables["Table"].Rows[0][0].ToString() != "Order #" && ds1.Tables["Table"].Rows[0][1].ToString() != "Part #" && ds1.Tables["Table"].Rows[0][2].ToString() != "IMEI" && ds1.Tables["Table"].Rows[0][3].ToString() != "Model")
                throw new ArgumentException(String.Format("Sorry, I could not find the specific format. Your document header must be {0} ", excelHeader));

            var finalList = new List<T>();
            for (int i = 0; i <= ds1.Tables["Table"].Rows.Count - 1; i++)
            {
                if (i == 0)
                    continue;

                if (String.IsNullOrEmpty(ds1.Tables["Table"].Rows[i]["F1"].ToString()) && String.IsNullOrEmpty(ds1.Tables["Table"].Rows[i]["F2"].ToString()) && String.IsNullOrEmpty(ds1.Tables["Table"].Rows[i]["F3"].ToString()) && String.IsNullOrEmpty(ds1.Tables["Table"].Rows[i]["F4"].ToString()))
                    continue;
                //throw new ArgumentException(String.Format("Sorry, I could not load all the data. I found one empty starting at row #{0}", i + 1));

                if (String.IsNullOrEmpty(ds1.Tables["Table"].Rows[i]["F1"].ToString()) || String.IsNullOrEmpty(ds1.Tables["Table"].Rows[i]["F3"].ToString()))
                    throw new ArgumentException(String.Format("Sorry, I could not load all the data. I found essensial values missing starting at row #{0}", i + 1));

                if (ds1.Tables["Table"].Rows[i]["F3"].ToString().Length < 15)
                    throw new ArgumentException(String.Format("Sorry, I could not load all the data. I found one IMEI with the wrong lenght at row #{0}", i + 1));


                var lot = ds1.Tables["Table"].Rows[i]["F1"].ToString();

                object q = new object();
                //q.ORDERNUMBER = ds1.Tables["Table"].Rows[i]["F1"].ToString();
                //q.SKU = ds1.Tables["Table"].Rows[i]["F2"].ToString();
                //q.IMEI = ds1.Tables["Table"].Rows[i]["F3"].ToString();
                //q.MODEL = ds1.Tables["Table"].Rows[i]["F4"].ToString();
                //q.CREATIONDATE = DateTime.Now;
                //q.LOTTYPE = lotType;
                //q.USERID = 12;
                //q.STATUSID = 1;

                //list.Add();
            }
            //finalList = list.DistinctBy(x => new { x.ORDERNUMBER, x.IMEI }).ToList();

            try
            {
                return finalList;
            }
            catch (Exception ex)
            {
                throw new ArgumentException("There was a problem while reading the file");
            }
        }
    }
}
