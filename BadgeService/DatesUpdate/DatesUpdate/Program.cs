using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace DatesUpdate
{
    class Program
    {

        //private static string connString = "Data Source=DESKTOP-HR2P137\\SQLEXPRESS;Initial Catalog=HRdb;Integrated Security=True";
         private static string connString = "Data Source=KNSQL2014;Initial Catalog=WbmOlimpiasHR;User=sa;Password=onlyouolimpias";
        //private static string connString = "Data Source=KNSQL2014;Initial Catalog=WbmBackup;User=sa;Password=onlyouolimpias";
        //private static string connString = "Data Source=DESKTOP-VRT03OS\\SQLEXPRESS;Initial Catalog=WbmOlimpiasHR;Integrated Security=True";
       // private static string connString = "Data Source=DESKTOP-VRT03OS;Initial Catalog=WbmOlimpiasHR;User=sa;Password=sergiu123";
        
       static string pathNameGiusti = "D:/HR_Files/residuo";
           
        public static bool IsDirectoryEmpty(string path)
        {
            IEnumerable<string> items = Directory.EnumerateFileSystemEntries(path);
            using (IEnumerator<string> en = items.GetEnumerator())
            {
                return !en.MoveNext();
            }
        }
        static void Main(string[] args)
        {
            //bool giustiFileNotExist = IsDirectoryEmpty(pathNameGiusti);
            SqlConnection conn = new SqlConnection(connString);

           
                //SqlConnection conn = new SqlConnection(connString);
                try
                {
                    var csvFiles = Directory.EnumerateFiles(pathNameGiusti, "*.csv");
                    foreach (string currentFile in csvFiles)
                    {
                    int i = 0;
                        conn.Open();
                        DataTable res = new DataTable();
                        res = ConvertCSVtoDataTable(currentFile);
                        foreach (DataRow row in res.Rows)
                        { i++;
                            using (SqlCommand cmd = new SqlCommand($"insert into FerieTotYY (Code, FullName, Hour, Year) values ('" + row[0].ToString() + "','" + row[1].ToString() + "'," + row[2].ToString() + "," + row[3].ToString() + ")", conn))
                            {
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                            }
                        Console.WriteLine("Insertd:" + i.ToString());
                    }
                    //conn.Open();
                    //DateTime date = new DateTime();
                    //DateTime begin = new DateTime(2021, 1, 1);
                    //DateTime end = new DateTime(2021, 12, 31);
                    //for ( date = begin; date <= end; date = date.AddDays(1))
                    //{
                    //    Console.WriteLine("Insertd:" + date.ToString("yyyy-MM-dd"));

                    //    using (SqlCommand cmd = new SqlCommand($"insert into Dates(d) values('" + date.ToString("yyyy-MM-dd") + "')", conn))
                    //    {
                    //        cmd.CommandType = CommandType.Text;
                    //        cmd.ExecuteNonQuery();

                    //    }

                    //}
                    conn.Close();
                    }
                
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }

            
            }
            public static DataTable ConvertCSVtoDataTable(string strFilePath)
            {
                using (StreamReader sr = new StreamReader(strFilePath))
                {
                    char[] separator = new char[] { ',' };
                    string[] result1 = sr.ReadLine().Split(separator, StringSplitOptions.None);

                    string[] headers = result1;
                    DataTable dt = new DataTable();
                    foreach (string header in headers)
                    {
                        dt.Columns.Add(header);
                    }
                    while (!sr.EndOfStream)
                    {
                        string[] rows = sr.ReadLine().Split(separator, StringSplitOptions.None);
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i];
                        }
                        dt.Rows.Add(dr);
                    }
                    return dt;
                }
            }
        }
}
