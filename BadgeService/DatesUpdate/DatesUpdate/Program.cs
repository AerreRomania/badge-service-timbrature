using System;
using System.Data;
using System.Data.SqlClient;

namespace DatesUpdate
{
    class Program
    {

        //private static string connString = "Data Source=DESKTOP-HR2P137\\SQLEXPRESS;Initial Catalog=HRdb;Integrated Security=True";
        private static string connString = "Data Source=KNSQL2014;Initial Catalog=WbmOlimpiasHR;User=sa;Password=onlyouolimpias";
        //private static string connString = "Data Source=KNSQL2014;Initial Catalog=WbmBackup;User=sa;Password=onlyouolimpias";
        //private static string connString = "Data Source=DESKTOP-VRT03OS\\SQLEXPRESS;Initial Catalog=WbmOlimpiasHR;Integrated Security=True";
        
        static void Main(string[] args)
        {
            SqlConnection conn = new SqlConnection(connString);
           
            try
            {
                
                conn.Open();
                DateTime date = new DateTime();
                DateTime begin = new DateTime(2021, 1, 1);
                DateTime end = new DateTime(2021, 12, 31);
                for ( date = begin; date <= end; date = date.AddDays(1))
                {
                    Console.WriteLine("Insertd:" + date.ToString("yyyy-MM-dd"));

                    using (SqlCommand cmd = new SqlCommand($"insert into Dates(d) values('" + date.ToString("yyyy-MM-dd") + "')", conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.ExecuteNonQuery();

                    }

                }
                conn.Close();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
