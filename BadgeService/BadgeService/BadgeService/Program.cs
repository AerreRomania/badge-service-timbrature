using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Text;

namespace HRUploadingApp
{
    class Program
    {
        static System.Timers.Timer MainUploadTimer = new System.Timers.Timer();
        static System.Timers.Timer EmailTimer = new System.Timers.Timer();

        static void Main(string[] args)
        {
            Console.WriteLine("The application started at {0:HH:mm:ss.fff}", DateTime.Now);
            LoadMain();
            MainUploadTimer.Interval = 60000;
            MainUploadTimer.Elapsed += new System.Timers.ElapsedEventHandler(MainUploadTimer_Tick);
            MainUploadTimer.Start();

            EmailTimer.Interval = 60000;
            EmailTimer.Elapsed += new System.Timers.ElapsedEventHandler(EmailTimer_Tick);
            EmailTimer.Start();

            Console.WriteLine("Press \'q\' to quit the timers and app.");
            while (Console.Read() != 'q') ;

        }

        static private void MainUploadTimer_Tick(object sender, System.Timers.ElapsedEventArgs e)
        {
            LoadMain();
        }

        static private void EmailTimer_Tick(object sender, System.Timers.ElapsedEventArgs e)
        {
            int CurrentHour = DateTime.Now.Hour;
            int CurrentMinut = DateTime.Now.Minute;

            if (CurrentHour == 9 && CurrentMinut == 10)
            {
                SendEmails(connString,1);
            } else if ( CurrentHour == 16 && CurrentMinut == 10 )
            {
                SendEmails(connString, 2);
            } else if (CurrentHour == 1 && CurrentMinut == 10)
            {
                SendEmails(connString, 3);
            }
        }

        static void LoadMain()
        {
            string pathNameAnagrafiche = "D:/HR_Files/Anagrafiche";
            bool anagraficheFileNotExist = IsDirectoryEmpty(pathNameAnagrafiche);

            string pathNameGiusti = "D:/HR_Files/Giusti";
            bool giustiFileNotExist = IsDirectoryEmpty(pathNameGiusti);

            string pathNameTimbrature = "D:/HR_Files/Timbrature";
            bool timbratureFileNotExist = IsDirectoryEmpty(pathNameTimbrature);

            if (anagraficheFileNotExist == false)
            {
                MainUploadTimer.Stop();
                AnagraficheUpload(pathNameAnagrafiche);
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("There is no files in Anagrafiche folder, wait for next refresh in 1 min");
                MainUploadTimer.Start();
            }

            if (giustiFileNotExist == false)
            {
                MainUploadTimer.Stop();
                GiustiUpload(pathNameGiusti);
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("There is no files in Giusti folder, wait for next refresh in 1 min");
                MainUploadTimer.Start();
            }

            if (timbratureFileNotExist == false)
            {
                MainUploadTimer.Stop();
                TimbratureUpload(pathNameTimbrature);
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("There is no files in Timbrature folder, wait for next refresh in 1 min");
                MainUploadTimer.Start();
            }
        }

        public static void SendEmails(string connString, int shift)
        {
            //GET DELAYED PERSONS
            DataTable delayedPersons = new DataTable();
            DataTable delayedPersonsConf = new DataTable();
            DataTable delayedPersonsStiro = new DataTable();
            DataTable delayedPersonsTess = new DataTable();

            delayedPersonsConf.Columns.Add(new DataColumn("RCode", typeof(string)));
            delayedPersonsConf.Columns.Add(new DataColumn("FullName", typeof(string)));
            delayedPersonsConf.Columns.Add(new DataColumn("Departament", typeof(string)));
            delayedPersonsConf.Columns.Add(new DataColumn("Mansione", typeof(string)));
            delayedPersonsConf.Columns.Add(new DataColumn("Linie", typeof(string)));
            delayedPersonsConf.Columns.Add(new DataColumn("Time", typeof(string)));

            delayedPersonsStiro.Columns.Add(new DataColumn("RCode", typeof(string)));
            delayedPersonsStiro.Columns.Add(new DataColumn("FullName", typeof(string)));
            delayedPersonsStiro.Columns.Add(new DataColumn("Departament", typeof(string)));
            delayedPersonsStiro.Columns.Add(new DataColumn("Mansione", typeof(string)));
            delayedPersonsStiro.Columns.Add(new DataColumn("Linie", typeof(string)));
            delayedPersonsStiro.Columns.Add(new DataColumn("Time", typeof(string)));

            delayedPersonsTess.Columns.Add(new DataColumn("RCode", typeof(string)));
            delayedPersonsTess.Columns.Add(new DataColumn("FullName", typeof(string)));
            delayedPersonsTess.Columns.Add(new DataColumn("Departament", typeof(string)));
            delayedPersonsTess.Columns.Add(new DataColumn("Mansione", typeof(string)));
            delayedPersonsTess.Columns.Add(new DataColumn("Linie", typeof(string)));
            delayedPersonsTess.Columns.Add(new DataColumn("Time", typeof(string)));

            if (shift == 1) {
                using (var da = new SqlDataAdapter("SELECT dbo.BadgeViewMatrix.IdRnum,CONCAT(dbo.BadgeViewMatrix.Nume, ' ', dbo.BadgeViewMatrix.Prenume) As[FullName]," +
                    "dbo.BadgeViewMatrix.Departament, dbo.PosturiDeLucru.PostDeLucru AS Mansione,dbo.BadgeViewMatrix.Linie,dbo.BadgeViewMatrix.Time " +
                    "FROM dbo.PosturiDeLucru INNER JOIN dbo.AngajatiViewLastMonth ON dbo.PosturiDeLucru.Id = dbo.AngajatiViewLastMonth.IdPostDeLucru " +
                    "INNER JOIN dbo.BadgeViewMatrix ON dbo.AngajatiViewLastMonth.Marca = dbo.BadgeViewMatrix.IdRnum " +
                    "WHERE CONCAT('20', dbo.BadgeViewMatrix.Year) = (YEAR(GETDATE())) AND dbo.BadgeViewMatrix.Month >= (MONTH(GETDATE())) AND dbo.BadgeViewMatrix.Day >= (DAY(GETDATE())) " +
                    "AND CONCAT('20', dbo.BadgeViewMatrix.Year) = (YEAR(GETDATE())) AND dbo.BadgeViewMatrix.Month <= (MONTH(GETDATE())) AND dbo.BadgeViewMatrix.Day <= (DAY(GETDATE())) " +
                    "AND dbo.BadgeViewMatrix.IsLate = 1 GROUP BY dbo.BadgeViewMatrix.IdRnum, dbo.BadgeViewMatrix.Year, dbo.BadgeViewMatrix.Month, dbo.BadgeViewMatrix.Day, dbo.BadgeViewMatrix.DYear, " +
                    "dbo.BadgeViewMatrix.DMonth, dbo.BadgeViewMatrix.Nume, dbo.BadgeViewMatrix.Prenume, dbo.BadgeViewMatrix.Departament, dbo.PosturiDeLucru.PostDeLucru, dbo.BadgeViewMatrix.Linie," +
                    " dbo.BadgeViewMatrix.Time, dbo.BadgeViewMatrix.IO, dbo.BadgeViewMatrix.IsLate ORDER BY CONCAT(dbo.BadgeViewMatrix.Nume, ' ', dbo.BadgeViewMatrix.Prenume), " +
                    "CONCAT(dbo.BadgeViewMatrix.DAY, '/', dbo.BadgeViewMatrix.MONTH, '/', '20', dbo.BadgeViewMatrix.YEAR), dbo.BadgeViewMatrix.Time ASC", connString))
                {
                    da.Fill(delayedPersons);
                }
            } else if (shift == 2)
            {
                using (var da = new SqlDataAdapter("SELECT dbo.BadgeViewMatrix.IdRnum,CONCAT(dbo.BadgeViewMatrix.Nume, ' ', dbo.BadgeViewMatrix.Prenume) As[FullName]," +
               "dbo.BadgeViewMatrix.Departament, dbo.PosturiDeLucru.PostDeLucru AS Mansione,dbo.BadgeViewMatrix.Linie,dbo.BadgeViewMatrix.Time " +
               "FROM dbo.PosturiDeLucru INNER JOIN dbo.AngajatiViewLastMonth ON dbo.PosturiDeLucru.Id = dbo.AngajatiViewLastMonth.IdPostDeLucru " +
               "INNER JOIN dbo.BadgeViewMatrix ON dbo.AngajatiViewLastMonth.Marca = dbo.BadgeViewMatrix.IdRnum " +
               "WHERE CONCAT('20', dbo.BadgeViewMatrix.Year) = (YEAR(GETDATE())) AND dbo.BadgeViewMatrix.Month >= (MONTH(GETDATE())) AND dbo.BadgeViewMatrix.Day >= (DAY(GETDATE())) " +
               "AND CONCAT('20', dbo.BadgeViewMatrix.Year) = (YEAR(GETDATE())) AND dbo.BadgeViewMatrix.Month <= (MONTH(GETDATE())) AND dbo.BadgeViewMatrix.Day <= (DAY(GETDATE())) " +
               "AND dbo.BadgeViewMatrix.IsLate = 1 and dbo.BadgeViewMatrix.Time > '14:00' " +
               "GROUP BY dbo.BadgeViewMatrix.IdRnum, dbo.BadgeViewMatrix.Year, dbo.BadgeViewMatrix.Month, dbo.BadgeViewMatrix.Day, dbo.BadgeViewMatrix.DYear, " +
               "dbo.BadgeViewMatrix.DMonth, dbo.BadgeViewMatrix.Nume, dbo.BadgeViewMatrix.Prenume, dbo.BadgeViewMatrix.Departament, dbo.PosturiDeLucru.PostDeLucru, dbo.BadgeViewMatrix.Linie," +
               " dbo.BadgeViewMatrix.Time, dbo.BadgeViewMatrix.IO, dbo.BadgeViewMatrix.IsLate ORDER BY CONCAT(dbo.BadgeViewMatrix.Nume, ' ', dbo.BadgeViewMatrix.Prenume), " +
               "CONCAT(dbo.BadgeViewMatrix.DAY, '/', dbo.BadgeViewMatrix.MONTH, '/', '20', dbo.BadgeViewMatrix.YEAR), dbo.BadgeViewMatrix.Time ASC", connString))
                {
                    da.Fill(delayedPersons);
                }
            } else if (shift == 3)
            {
                using (var da = new SqlDataAdapter("SELECT dbo.BadgeViewMatrix.IdRnum,CONCAT(dbo.BadgeViewMatrix.Nume, ' ', dbo.BadgeViewMatrix.Prenume) As[FullName]," +
               "dbo.BadgeViewMatrix.Departament, dbo.PosturiDeLucru.PostDeLucru AS Mansione,dbo.BadgeViewMatrix.Linie,dbo.BadgeViewMatrix.Time " +
               "FROM dbo.PosturiDeLucru INNER JOIN dbo.AngajatiViewLastMonth ON dbo.PosturiDeLucru.Id = dbo.AngajatiViewLastMonth.IdPostDeLucru " +
               "INNER JOIN dbo.BadgeViewMatrix ON dbo.AngajatiViewLastMonth.Marca = dbo.BadgeViewMatrix.IdRnum " +
               "WHERE CONCAT('20', dbo.BadgeViewMatrix.Year) = (YEAR(GETDATE())) AND dbo.BadgeViewMatrix.Month >= (MONTH(GETDATE())) AND dbo.BadgeViewMatrix.Day >= (DAY(GETDATE())) " +
               "AND CONCAT('20', dbo.BadgeViewMatrix.Year) = (YEAR(GETDATE())) AND dbo.BadgeViewMatrix.Month <= (MONTH(GETDATE())) AND dbo.BadgeViewMatrix.Day <= (DAY(GETDATE())) " +
               "AND dbo.BadgeViewMatrix.IsLate = 1 and dbo.BadgeViewMatrix.Time > '20:00' " +
               "GROUP BY dbo.BadgeViewMatrix.IdRnum, dbo.BadgeViewMatrix.Year, dbo.BadgeViewMatrix.Month, dbo.BadgeViewMatrix.Day, dbo.BadgeViewMatrix.DYear, " +
               "dbo.BadgeViewMatrix.DMonth, dbo.BadgeViewMatrix.Nume, dbo.BadgeViewMatrix.Prenume, dbo.BadgeViewMatrix.Departament, dbo.PosturiDeLucru.PostDeLucru, dbo.BadgeViewMatrix.Linie," +
               " dbo.BadgeViewMatrix.Time, dbo.BadgeViewMatrix.IO, dbo.BadgeViewMatrix.IsLate ORDER BY CONCAT(dbo.BadgeViewMatrix.Nume, ' ', dbo.BadgeViewMatrix.Prenume), " +
               "CONCAT(dbo.BadgeViewMatrix.DAY, '/', dbo.BadgeViewMatrix.MONTH, '/', '20', dbo.BadgeViewMatrix.YEAR), dbo.BadgeViewMatrix.Time ASC", connString))
                {
                    da.Fill(delayedPersons);
                }
            }

            foreach (DataRow o in delayedPersons.Select("Departament = 'CONFEZIONE A' OR Departament = 'CONFEZIONE B'"))
            {
                delayedPersonsConf.Rows.Add(o[0], o[1], o[2], o[3], o[4], o[5]);
            }

            foreach (DataRow o in delayedPersons.Select("Departament = 'STIRO'"))
            {
                delayedPersonsStiro.Rows.Add(o[0], o[1], o[2], o[3], o[4], o[5]);
            }

            foreach (DataRow o in delayedPersons.Select("Departament = 'TESSITURA'"))
            {
                delayedPersonsTess.Rows.Add(o[0], o[1], o[2], o[3], o[4], o[5]);
            }

            SendAll(delayedPersons);
            SendConf(delayedPersonsConf);
            SendStiro(delayedPersonsStiro);
            SendTess(delayedPersonsTess);
        }

        private static void SendAll(DataTable dt)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("noreply@olimpias.rs", "Timbrature Report");
            mail.To.Add("giovanni@antonioli.it");
            mail.To.Add("jjevtic@olimpias.rs");
            mail.To.Add("bdraskovic@olimpias.rs");

            mail.Subject = "Late persons - '" + DateTime.Now.ToShortDateString() + "' - "+DateTime.Now.TimeOfDay+"";
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("<table>");

            sb.AppendLine("<thead>");
            sb.AppendLine("<th>RCode:</th>");
            sb.AppendLine("<th>Full Name:</th>");
            sb.AppendLine("<th>Departament:</th>");
            sb.AppendLine("<th>Mansione:</th>");
            sb.AppendLine("<th>Linea:</th>");
            sb.AppendLine("<th>Time:</th>");
            sb.AppendLine("</thead>");

            sb.AppendLine("<tbody>");
            if (dt.Rows.Count > 1) { 
                foreach (DataRow o in dt.Rows)
                {
                    sb.AppendLine("<tr>");
                    sb.AppendLine("<td>"+o[0]+ "</td> <td>" + o[1] + "</td> <td>" + o[2] + "</td> <td>" + o[3] + "</td> <td>" + o[4] + "</td> <td>" + o[5] + "</td>");
                    sb.AppendLine("</tr>");
                }
            } else
            {
                sb.AppendLine("<tr>");
                sb.AppendLine("<td>All persons are checked on the time</td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody>");

            sb.AppendLine("</table>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            mail.Body = sb.ToString();
            mail.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient("mail.olimpias.it");
            smtp.Port = 25;
            smtp.Send(mail);
        }
        private static void SendConf(DataTable dt)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("noreply@olimpias.rs", "Timbrature Report");
            mail.To.Add("imiljkovic@olimpias.rs");
            mail.To.Add("azivkovic@olimpias.rs");

            mail.Subject = "Late persons - '" + DateTime.Now.ToShortDateString() + "' - " + DateTime.Now.TimeOfDay + "";
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("<table>");

            sb.AppendLine("<thead>");
            sb.AppendLine("<th>RCode:</th>");
            sb.AppendLine("<th>Full Name:</th>");
            sb.AppendLine("<th>Departament:</th>");
            sb.AppendLine("<th>Mansione:</th>");
            sb.AppendLine("<th>Linea:</th>");
            sb.AppendLine("<th>Time:</th>");
            sb.AppendLine("</thead>");

            sb.AppendLine("<tbody>");
            if (dt.Rows.Count > 1)
            {
                foreach (DataRow o in dt.Rows)
                {
                    sb.AppendLine("<tr>");
                    sb.AppendLine("<td>" + o[0] + "</td> <td>" + o[1] + "</td> <td>" + o[2] + "</td> <td>" + o[3] + "</td> <td>" + o[4] + "</td> <td>" + o[5] + "</td>");
                    sb.AppendLine("</tr>");
                }
            }
            else
            {
                sb.AppendLine("<tr>");
                sb.AppendLine("<td>All persons are checked on the time</td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody>");

            sb.AppendLine("</table>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            mail.Body = sb.ToString();
            mail.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient("mail.olimpias.it");
            smtp.Port = 25;
            smtp.Send(mail);
        }
        private static void SendStiro(DataTable dt)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("noreply@olimpias.rs", "Timbrature Report");
            mail.To.Add("mcvetkovic@olimpias.rs");

            mail.Subject = "Late persons - '" + DateTime.Now.ToShortDateString() + "' - " + DateTime.Now.TimeOfDay + "";
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("<table>");

            sb.AppendLine("<thead>");
            sb.AppendLine("<th>RCode:</th>");
            sb.AppendLine("<th>Full Name:</th>");
            sb.AppendLine("<th>Departament:</th>");
            sb.AppendLine("<th>Mansione:</th>");
            sb.AppendLine("<th>Linea:</th>");
            sb.AppendLine("<th>Time:</th>");
            sb.AppendLine("</thead>");

            sb.AppendLine("<tbody>");
            if (dt.Rows.Count > 1)
            {
                foreach (DataRow o in dt.Rows)
                {
                    sb.AppendLine("<tr>");
                    sb.AppendLine("<td>" + o[0] + "</td> <td>" + o[1] + "</td> <td>" + o[2] + "</td> <td>" + o[3] + "</td> <td>" + o[4] + "</td> <td>" + o[5] + "</td>");
                    sb.AppendLine("</tr>");
                }
            }
            else
            {
                sb.AppendLine("<tr>");
                sb.AppendLine("<td>All persons are checked on the time</td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody>");

            sb.AppendLine("</table>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            mail.Body = sb.ToString();
            mail.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient("mail.olimpias.it");
            smtp.Port = 25;
            smtp.Send(mail);
        }
        private static void SendTess(DataTable dt)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("noreply@olimpias.rs", "Timbrature Report");
            mail.To.Add("dbojic@olimpias.rs");

            mail.Subject = "Late persons - '" + DateTime.Now.ToShortDateString() + "' - " + DateTime.Now.TimeOfDay + "";
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("<table>");

            sb.AppendLine("<thead>");
            sb.AppendLine("<th>RCode:</th>");
            sb.AppendLine("<th>Full Name:</th>");
            sb.AppendLine("<th>Departament:</th>");
            sb.AppendLine("<th>Mansione:</th>");
            sb.AppendLine("<th>Linea:</th>");
            sb.AppendLine("<th>Time:</th>");
            sb.AppendLine("</thead>");

            sb.AppendLine("<tbody>");
            if (dt.Rows.Count > 1)
            {
                foreach (DataRow o in dt.Rows)
                {
                    sb.AppendLine("<tr>");
                    sb.AppendLine("<td>" + o[0] + "</td> <td>" + o[1] + "</td> <td>" + o[2] + "</td> <td>" + o[3] + "</td> <td>" + o[4] + "</td> <td>" + o[5] + "</td>");
                    sb.AppendLine("</tr>");
                }
            }
            else
            {
                sb.AppendLine("<tr>");
                sb.AppendLine("<td>All persons are checked on the time</td>");
                sb.AppendLine("</tr>");
            }
            sb.AppendLine("</tbody>");

            sb.AppendLine("</table>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            mail.Body = sb.ToString();
            mail.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient("mail.olimpias.it");
            smtp.Port = 25;
            smtp.Send(mail);
        }

        //private static string connString = "Data Source=DESKTOP-HR2P137\\SQLEXPRESS;Initial Catalog=HRdb;Integrated Security=True";
          //private static string connString = "Data Source=KNSQL2014;Initial Catalog=WbmOlimpiasHR;User=sa;Password=onlyouolimpias";
        //private static string connString = "Data Source=KNSQL2014;Initial Catalog=WbmBackup;User=sa;Password=onlyouolimpias";
         private static string connString = "Data Source=DESKTOP-VRT03OS\\SQLEXPRESS;Initial Catalog=WbmOlimpiasHR;Integrated Security=True";
        public static bool IsDirectoryEmpty(string path)
        {
            IEnumerable<string> items = Directory.EnumerateFileSystemEntries(path);
            using (IEnumerator<string> en = items.GetEnumerator())
            {
                return !en.MoveNext();
            }
        }
        private static void AnagraficheUpload(string path)
        {
            SqlConnection conn = new SqlConnection(connString);
            try
            {
                var csvFiles = Directory.EnumerateFiles(path, "*.csv");
                foreach (string currentFile in csvFiles)
                {
                    DataTable res = new DataTable();
                    res = ConvertCSVtoDataTable(currentFile);
                    string filenameWithoutPath = Path.GetFileName(currentFile);
                    string mm = filenameWithoutPath.Split('.')[1];
                    string yy = filenameWithoutPath.Split('.')[2];

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine();
                    Console.WriteLine("----------------------------------------------------------------------");
                    Console.WriteLine("STARTED UPLOADING - '" + filenameWithoutPath + "'");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("YEAR - '" + yy + "'");
                    Console.WriteLine("MONTH - '" + mm + "'");

                    int fileYear = Convert.ToInt32(yy);
                    int fileMonth = Convert.ToInt32(mm);
                    int counter = -1;
                    int totalRows = res.Rows.Count - 1;
                    conn.Open();

                    Console.WriteLine();
                    foreach (DataRow row in res.Rows)
                    {
                        try
                        {

                            counter++;
                            Console.SetCursorPosition(0, Console.CursorTop - 1);
                            Console.WriteLine("INSERTED: " + counter + " / " + totalRows);

                            string mrc = Convert.ToString(row["R3MAT"]);
                            string mrcSubstringed = mrc.TrimStart('R');
                            int codAngajat = Convert.ToInt32(mrcSubstringed);

                            string marca = Convert.ToString(row["R3MAT"]);
                            string nume = Convert.ToString(row["R3NOM"]);
                            string prenume = Convert.ToString(row["R3COG"]);
                            string numaridenticare = Convert.ToString(row["R3FSC"]);
                            string sex = Convert.ToString(row["R3SEX"]);
                            string localitate = Convert.ToString(row["R3CMR"]);
                            string strada = Convert.ToString(row["R3VIR"]);
                            string codpostdelucru = Convert.ToString(row["R3DEM"]);
                            string postdelucru = Convert.ToString(row["R3DEM"]);
                            string coddepartment = Convert.ToString(row["R3DEA"]);
                            string departament = Convert.ToString(row["R3DEA"]);
                            string telefon = Convert.ToString(row["R3PTE"]);
                            string cc = Convert.ToString(row["R3IBA"]);
                            string codincadrare = Convert.ToString(row["R3CCD"]);
                            string incadrare = Convert.ToString(row["R3DCD"]);
                            string tippostdelucru = Convert.ToString(row["R3D07"]);
                            int deteindete = Convert.ToInt32(row["R3U02"]);
                            string linie = Convert.ToString(row["R3DLV"]);
                            int LevelOfStudy = Convert.ToInt32(row["R3TIS"]);
                            string Email = Convert.ToString(row["R3MAL"]);
                            int IdBadge = Convert.ToInt32(row["OPTID"]);
                            DateTime datumzaposljenja;
                            DateTime datumraskida;
                            DateTime datumistekaugovora;
                            DateTime prelazaknaneodredjeno;
                            //ok
                            int? dan = Convert.ToInt32(row["R3GAS"]);
                            int? mesec = Convert.ToInt32(row["R3MAS"]);
                            int? godina = Convert.ToInt32(row["R3AAS"]);
                            if (godina == 0 || godina == null)
                            {
                                datumzaposljenja = Convert.ToDateTime(new DateTime(0001, 1, 1).ToShortDateString());
                            }
                            else
                            {
                                datumzaposljenja = Convert.ToDateTime(new DateTime(Convert.ToInt32(godina), Convert.ToInt32(mesec), Convert.ToInt32(dan)).ToShortDateString());
                            }

                            //ok sss expire of the contract

                            int? dan1 = Convert.ToInt32(row["R3GCE"]);
                            int? mesec1 = Convert.ToInt32(row["R3MCE"]);
                            int? godina1 = Convert.ToInt32(row["R3ACE"]);
                            if (godina1 == 0 || godina == null)
                            {
                                datumraskida = Convert.ToDateTime(new DateTime(0001, 1, 1).ToShortDateString());
                            }
                            else
                            {
                                datumraskida = Convert.ToDateTime(new DateTime(Convert.ToInt32(godina1), Convert.ToInt32(mesec1), Convert.ToInt32(dan1)).ToShortDateString());
                            }


                            //CHANGE CONTRACT
                            int? dan3 = Convert.ToInt32(row["R3GS1"]);
                            int? mesec3 = Convert.ToInt32(row["R3MS1"]);
                            int? godina3 = Convert.ToInt32(row["R3AS1"]);
                            if (godina3 == 0)
                            {
                                prelazaknaneodredjeno = Convert.ToDateTime(new DateTime(0001, 1, 1).ToShortDateString());
                            }
                            else
                            {
                                prelazaknaneodredjeno = Convert.ToDateTime(new DateTime(Convert.ToInt32(godina3), Convert.ToInt32(mesec3), Convert.ToInt32(dan3)).ToShortDateString());
                            }


                            using (SqlCommand cmd = new SqlCommand("Import_Angajati", conn))
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.Add("@CodAngajat", SqlDbType.Int).Value = codAngajat;
                                cmd.Parameters.Add("@Nume", SqlDbType.VarChar).Value = nume;
                                cmd.Parameters.Add("@Prenume", SqlDbType.VarChar).Value = prenume;
                                cmd.Parameters.Add("@CodSistem", SqlDbType.Int).Value = codAngajat;
                                cmd.Parameters.Add("@Marca", SqlDbType.VarChar).Value = marca;
                                cmd.Parameters.Add("@NumarIdentificarePersonala", SqlDbType.VarChar).Value = numaridenticare;
                                cmd.Parameters.Add("@Sex", SqlDbType.VarChar).Value = sex;
                                cmd.Parameters.Add("@Localitate", SqlDbType.VarChar).Value = localitate;
                                cmd.Parameters.Add("@Strada", SqlDbType.VarChar).Value = strada;
                                cmd.Parameters.Add("@CodPostDeLucru", SqlDbType.VarChar).Value = codpostdelucru;
                                cmd.Parameters.Add("@PostDeLucru", SqlDbType.VarChar).Value = postdelucru;
                                cmd.Parameters.Add("@CodDepartament", SqlDbType.VarChar).Value = coddepartment;
                                cmd.Parameters.Add("@Departament", SqlDbType.VarChar).Value = departament;
                                cmd.Parameters.Add("@CodIncadrare", SqlDbType.VarChar).Value = codincadrare;
                                cmd.Parameters.Add("@Incadrare", SqlDbType.VarChar).Value = incadrare;
                                cmd.Parameters.Add("@CC", SqlDbType.VarChar).Value = cc;
                                cmd.Parameters.Add("@TipPostDeLucru", SqlDbType.VarChar).Value = tippostdelucru;
                                cmd.Parameters.Add("@LoculNasterii", SqlDbType.VarChar).Value = localitate;
                                cmd.Parameters.Add("@Email", SqlDbType.VarChar).Value = Email;
                                cmd.Parameters.Add("@IdLiniee", SqlDbType.VarChar).Value = linie;
                                cmd.Parameters.Add("@IdUtilizatorAdaugare", SqlDbType.Int).Value = 1;
                                cmd.Parameters.Add("@Telefon", SqlDbType.VarChar).Value = telefon;
                                cmd.Parameters.Add("@IdNivelStudi", SqlDbType.Int).Value = 1;

                                cmd.Parameters.Add("@LevelOfStudy", SqlDbType.Int).Value = LevelOfStudy;

                                cmd.Parameters.Add("@DataAngajarii", SqlDbType.Date).Value = datumzaposljenja;
                                cmd.Parameters.Add("@DataLichidare", SqlDbType.Date).Value = datumraskida;
                                cmd.Parameters.Add("@DataExpirareContract", SqlDbType.Date).Value = prelazaknaneodredjeno; // datumistekaugovora
                                cmd.Parameters.Add("@DataNedeterminat", SqlDbType.Date).Value = prelazaknaneodredjeno;

                                cmd.Parameters.Add("@DataNasterii", SqlDbType.Date).Value = datumzaposljenja;
                                cmd.Parameters.Add("@DeteIndete", SqlDbType.Int).Value = deteindete;

                                cmd.Parameters.Add("@DYear", SqlDbType.Int).Value = fileYear;
                                cmd.Parameters.Add("@DMonth", SqlDbType.Int).Value = fileMonth;
                                cmd.ExecuteNonQuery();
                            }
                            using(SqlCommand cmd =new SqlCommand("Import_Badge",conn))
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.Add("@IdBadge", SqlDbType.Int).Value = IdBadge;
                                cmd.Parameters.Add("@RNum", SqlDbType.VarChar).Value = marca;
                                cmd.ExecuteNonQuery();
                            }
                        }
                        catch (Exception ex)
                        {
                            String query = "INSERT INTO dbo.Exception (Application,Exception,Time) VALUES (@application, @exception,@time)";
                            SqlCommand cmd = new SqlCommand(query, conn);
                            cmd.Parameters.Add("@application", SqlDbType.NVarChar).Value = "Anagrafiche Upload";
                            cmd.Parameters.Add("@exception", SqlDbType.NVarChar).Value = ex.ToString();
                            cmd.Parameters.Add("@time", SqlDbType.DateTime).Value = DateTime.Now;
                            cmd.ExecuteNonQuery();

                            Console.WriteLine(ex.Message.ToString());
                            continue;
                        }

                    }

                    conn.Close();

                    Console.ForegroundColor = ConsoleColor.DarkMagenta;
                    Console.WriteLine("INSERTING DONE!");
                    //TO ENABLE IN PROD : MOVE FILE TO BACKUP
                    var fullPath = Path.Combine("D:/HR_Files/Anagrafiche", filenameWithoutPath);
                    var fullPath1 = Path.Combine("D:/HR_Files/Backup", filenameWithoutPath);
                    File.Move(fullPath, fullPath1);
                }
            }
            catch (Exception e)
            {

                Console.BackgroundColor = ConsoleColor.Red;
                Console.WriteLine(e.Message);
            }
            MainUploadTimer.Start();
        }
        private static void GiustiUpload(string path)
        {
            SqlConnection conn = new SqlConnection(connString);
            try
            {
                var csvFiles = Directory.EnumerateFiles(path, "*.csv");
                foreach (string currentFile in csvFiles)
                {
                    DataTable res = new DataTable();
                    res = ConvertCSVtoDataTable(currentFile);
                    string filenameWithoutPath = Path.GetFileName(currentFile);
                    string mm = filenameWithoutPath.Split('.')[1];
                    string yy = filenameWithoutPath.Split('.')[2];

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine();
                    Console.WriteLine("----------------------------------------------------------------------");
                    Console.WriteLine("STARTED UPLOADING - '" + filenameWithoutPath + "'");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("YEAR - '" + yy + "'");
                    Console.WriteLine("MONTH - '" + mm + "'");

                    int fileYear = Convert.ToInt32(yy);
                    int fileMonth = Convert.ToInt32(mm);
                    int counter = -1;
                    int totalRows = res.Rows.Count - 1;
                    conn.Open();


                    int codAngajat = 0;
                    string codTipOra = "";
                    Console.WriteLine();
                    foreach (DataRow row in res.Rows)
                    {
                        try
                        {
                            counter++;
                            Console.SetCursorPosition(0, Console.CursorTop - 1);
                            Console.WriteLine("INSERTED:" + counter + " / " + totalRows);

                            string mrc = Convert.ToString(row[1]);
                            string mrcSubstringed = mrc.TrimStart('R');
                            codAngajat = Convert.ToInt32(mrcSubstringed);

                            codTipOra = Convert.ToString(row[5]);

                           

                            int dan = Convert.ToInt32(row[2]);
                            int mesec = Convert.ToInt32(row[3]);
                            int godina = Convert.ToInt32(row[4]);
                            string sve = new DateTime(godina, mesec, dan).ToShortDateString();


                            //decimal dal = Convert.ToDecimal(row["R1DAL"]);
                            //decimal all = Convert.ToDecimal(row["R1ALL"]);
                            decimal tot = Convert.ToDecimal(row[6]);
                            //int glupiid = Convert.ToInt32(row["R1KEY"]);


                            using (SqlCommand cmd = new SqlCommand("Import_Giusti", conn))
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.Add("@CodAngajat", SqlDbType.Int).Value = codAngajat;
                                cmd.Parameters.Add("@Data", SqlDbType.Date).Value = sve;
                                cmd.Parameters.Add("@CodTipOra", SqlDbType.NVarChar).Value = codTipOra;
                                //cmd.Parameters.Add("@Dep", SqlDbType.Int).Value = iddepartament;
                                cmd.Parameters.Add("@R1DAL", SqlDbType.Decimal).Value = 0.0;
                                cmd.Parameters.Add("@R1ALL", SqlDbType.Decimal).Value = 0.0;
                                cmd.Parameters.Add("@R1TOT", SqlDbType.Decimal).Value = tot;
                                cmd.Parameters.Add("@IdUtilizatorAdaugare", SqlDbType.Int).Value = 1;

                                cmd.ExecuteNonQuery();
                            }
                        }
                        catch (Exception ex)
                        {
                            String query = "INSERT INTO dbo.Exception (Application,Exception,Time) VALUES (@application, @exception,@time)";
                            SqlCommand cmd = new SqlCommand(query, conn);
                            cmd.Parameters.Add("@application", SqlDbType.NVarChar).Value = "Giusti Upload";
                            cmd.Parameters.Add("@exception", SqlDbType.NVarChar).Value = codAngajat + " " + codTipOra + ex.ToString();
                            cmd.Parameters.Add("@time", SqlDbType.DateTime).Value = DateTime.Now;
                            cmd.ExecuteNonQuery();

                            Console.WriteLine(ex.Message.ToString());
                            continue;
                        }
                    }

                    conn.Close();

                    Console.ForegroundColor = ConsoleColor.DarkMagenta;
                    Console.WriteLine("INSERTING DONE!");

                    //TO ENABLE IN PROD : MOVE FILE TO BACKUP
                    var fullPath = Path.Combine("D:/HR_Files/Giusti", filenameWithoutPath);
                    var fullPath1 = Path.Combine("D:/HR_Files/Backup", filenameWithoutPath);
                    File.Move(fullPath, fullPath1);
                }
            }
            catch (Exception e)
            {
                Console.BackgroundColor = ConsoleColor.Red;
                Console.WriteLine(e.Message);
            }
            MainUploadTimer.Start();

        }
        private static void TimbratureUpload(string path)
        {
            SqlConnection conn = new SqlConnection(connString);
            try
            {
                var csvFiles = Directory.EnumerateFiles(path, "*.csv");
                foreach (string currentFile in csvFiles)
                {
                    DataTable res = new DataTable();
                    res = ConvertCSVtoDataTable(currentFile);
                    string filenameWithoutPath = Path.GetFileName(currentFile);
                    string mm = filenameWithoutPath.Split('.')[1];
                    string yy = filenameWithoutPath.Split('.')[2];

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine();
                    Console.WriteLine("----------------------------------------------------------------------");
                    Console.WriteLine("STARTED UPLOADING - '" + filenameWithoutPath + "'");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("YEAR - '" + yy + "'");
                    Console.WriteLine("MONTH - '" + mm + "'");

                    int fileYear = Convert.ToInt32(yy);
                    int fileMonth = Convert.ToInt32(mm);
                    int counter = -1;
                    int totalRows = res.Rows.Count - 1;
                    conn.Open();

                    Console.WriteLine();
                    foreach (DataRow row in res.Rows)
                    {
                        try
                        {
                            counter++;
                            Console.SetCursorPosition(0, Console.CursorTop - 1);
                            Console.WriteLine("INSERTED:" + counter + " / " + totalRows);

                            string RCode = Convert.ToString(row["R3MAT"]);
                            int RYear = fileYear;
                            int RMonth = fileMonth;
                            int daysForMonth = DateTime.DaysInMonth(RYear, RMonth);

                            for (int i = 1; i <= daysForMonth; i++)
                            {
                                if (row["" + i + ""].ToString() != "")
                                {
                                    string str = row["" + i + ""].ToString();
                                    string value = After(str, "(");

                                        String query = "INSERT INTO [dbo].[Matrix_Tibrature] ([RCode],[RYear],[RMonth],[RDay],[RTime]) VALUES (@RCode, @RYear, @RMonth, @RDay, @RTime)";

                                        using (SqlCommand command = new SqlCommand(query, conn))
                                        {
                                            command.Parameters.AddWithValue("@RCode", RCode);
                                            command.Parameters.AddWithValue("@RYear", RYear);
                                            command.Parameters.AddWithValue("@RMonth", RMonth);
                                            command.Parameters.AddWithValue("@RDay", i); 
                                            command.Parameters.AddWithValue("@RTime", value.Substring(0, 5));

                                        int result = command.ExecuteNonQuery();

                                            // Check Error
                                            if (result < 0)
                                                Console.WriteLine("Error inserting data into Database!");
                                        }
                                   
                                }
                                else
                                {
                                    String query = "INSERT INTO [dbo].[Matrix_Tibrature] ([RCode],[RYear],[RMonth],[RDay],[RTime]) VALUES (@RCode, @RYear, @RMonth, @RDay, @RTime)";

                                    using (SqlCommand command = new SqlCommand(query, conn))
                                    {
                                        command.Parameters.AddWithValue("@RCode", RCode);
                                        command.Parameters.AddWithValue("@RYear", RYear);
                                        command.Parameters.AddWithValue("@RMonth", RMonth);
                                        command.Parameters.AddWithValue("@RDay", i);
                                        command.Parameters.AddWithValue("@RTime", -1);

                                        int result = command.ExecuteNonQuery();

                                        // Check Error
                                        if (result < 0)
                                            Console.WriteLine("Error inserting data into Database!");
                                    }
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            String query = "INSERT INTO dbo.Exception (Application,Exception,Time) VALUES (@application, @exception,@time)";
                            SqlCommand cmd = new SqlCommand(query, conn);
                            cmd.Parameters.Add("@application", SqlDbType.NVarChar).Value = "Timbrature Upload";
                            cmd.Parameters.Add("@exception", SqlDbType.NVarChar).Value = ex.ToString();
                            cmd.Parameters.Add("@time", SqlDbType.DateTime).Value = DateTime.Now;
                            cmd.ExecuteNonQuery();

                            Console.WriteLine(ex.Message.ToString());
                            continue;
                        }
                    }

                    conn.Close();

                    Console.ForegroundColor = ConsoleColor.DarkMagenta;
                    Console.WriteLine("INSERTING DONE!");

                    //TO ENABLE IN PROD : MOVE FILE TO BACKUP
                    var fullPath = Path.Combine("D:/HR_Files/Timbrature", filenameWithoutPath);
                    var fullPath1 = Path.Combine("D:/HR_Files/Backup", filenameWithoutPath);
                    File.Move(fullPath, fullPath1);
                }
            }
            catch (Exception e)
            {
                Console.BackgroundColor = ConsoleColor.Red;
                Console.WriteLine(e.Message);
            }
            MainUploadTimer.Start();
        }
        public static string Between(string value, string a, string b)
        {
            int posA = value.IndexOf(a);
            int posB = value.LastIndexOf(b);
            if (posA == -1)
            {
                return "";
            }
            if (posB == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= posB)
            {
                return "";
            }
            return value.Substring(adjustedPosA, posB - adjustedPosA);
        }
        public static string After(string value, string a)
        {
            int posA = value.LastIndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= value.Length)
            {
                return "";
            }
            return value.Substring(adjustedPosA);
        }
        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                char[] separator =  new char[] { '\t' };
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

