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

        
        
    }
}
