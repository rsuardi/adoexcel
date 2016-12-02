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

    class where
    {
        public string name { get; set; }
        public string columname { get; set; }
    }

    class FileReader
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

        public static DataTable SelectWorkingSheet(OleDbConnection conn, string sheet, string rangeExcel = null)
        {
            try
            {
                OleDbCommand command = new OleDbCommand("Select * from [" + sheet + "$" + rangeExcel + "]", conn);
                OleDbDataAdapter adpt = new OleDbDataAdapter(command);
                DataSet ds1 = new DataSet();
                adpt.Fill(ds1);
                command.ExecuteReader();
                conn.Close();
                DataTable table = ds1.Tables[0];
                return table;
            }
            catch (Exception ex)
            {
                return null;
            }   
        }

        public static List<T> findList<T>(DataTable data) where T : new ()
        {
            List<T> Temp = new List<T>();
            T Tempo = new T();



            PropertyInfo[] h = Tempo.GetType().GetProperties();
            List<string> find = new List<string>();
            

            foreach (PropertyInfo property in h)
            {
                    find.Add(property.Name); 
            }

            int des = -1;
            List<where> found = new List<where>();
            bool finded = false;

            
            for (int i = 0; i < data.Columns.Count; i++)
            {   
                
                for (int x = 0; x < data.Rows.Count; x++)
                {
                    if (finded == false)
                    {
                        foreach (string on in find)
                        {
                            if (data.Rows[x][i].ToString() == on)
                            {
                                finded = true;
                                des = x + 1;
                                found = scanl(data.Rows[x], find);
                                break;
                            }
                        }
                    }
                    else
                        break;
                }
                if (finded)
                    break;
            }
            
            for(int i = des; i < data.Rows.Count; i++)
            {
                Tempo = new T();
                foreach (where s in found)
                {
                    if (s.columname != null)
                    {

                        PropertyInfo propertyInfo = typeof(T).GetProperty(s.name);
                        propertyInfo.SetValue(Tempo, data.Rows[i][s.columname].ToString(), null);
                    }
                    else
                    {
                        PropertyInfo propertyInfo = typeof(T).GetProperty(s.name);
                        propertyInfo.SetValue(Tempo, string.Empty);
                    }
                }
                Temp.Add(Tempo);
            }   

            return Temp;
        }

        private static List<where> scanl(DataRow here, List<string> values)
        {
            List<where> temp = new List<where>();
            where tempo = new where();
            

            foreach (string val in values)
            {
                bool founded = false;
                foreach (DataColumn col in here.Table.Columns)
                {   
                    if (here[col].ToString() == val)
                    {
                        tempo = new where();
                        tempo.columname = col.ColumnName.ToString();
                        tempo.name = val;
                        temp.Add(tempo);
                        founded = true;
                    }
                }

                if(!founded)
                {
                    tempo = new where();
                    tempo.columname = null;
                    tempo.name = val;
                    temp.Add(tempo);
                }

            }
            

            return temp;
        }
        

    }


}
