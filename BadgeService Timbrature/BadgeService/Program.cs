using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Threading;

namespace BadgeService
{
    class Program
    {
        //Start Giusti Timer and Checking Shifts
        static void Main(string[] args)
        {
            Timer t = new Timer(GiustiServiceTimer, null, 0, 600000);
            LoadMain();
            Console.ReadLine();
        }

        //Checking Shifts -> Start Mail Timers
        private static void LoadMain()
        {
            
            String hourss = DateTime.Now.ToString("HH", System.Globalization.DateTimeFormatInfo.InvariantInfo);
            String minutess = DateTime.Now.ToString("mm", System.Globalization.DateTimeFormatInfo.InvariantInfo);

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Current Hour: " +hourss + ":" + minutess);

            if (Convert.ToInt32(hourss) >= 6 && Convert.ToInt32(hourss) <= 14)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("First Shift Selected:");
                Timer t = new Timer(FirstShift, null, 0, 60000);
            }

            if (Convert.ToInt32(hourss) >= 15 && Convert.ToInt32(hourss) <= 22)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Second Shift Selected:");
                Timer t = new Timer(SecondShift, null, 0, 60000);
            }
        }
        private static void SendEmail()
        {
            String hourss = DateTime.Now.ToString("HH", System.Globalization.DateTimeFormatInfo.InvariantInfo);
            String minutess = DateTime.Now.ToString("mm", System.Globalization.DateTimeFormatInfo.InvariantInfo);

            if (Convert.ToInt32(hourss) == 10 && Convert.ToInt32(minutess) == 40)
            {
                LoadDataFirst();
                Console.WriteLine("SALJEM MAIL PRVA SMENA");
            }

            if (Convert.ToInt32(hourss) == 15 && Convert.ToInt32(minutess) == 40)
            {
                //LoadDataSecond();
                LoadDataSecond();
                Console.WriteLine("SALJEM MAIL DRUGA SMENA");
            }

        }

        //Timers
        private static void GiustiServiceTimer(Object o)
        {
            SqlConnection conn = new SqlConnection("Data Source=KNSQL2014;Initial Catalog=WbmOlimpiasHR;User=sa;Password=onlyouolimpias");
            int counter = 0;
            try
            {
                var txtFiles = Directory.EnumerateFiles("\\\\192.168.96.35\\Users\\Time\\", "*.txt");
                foreach (string currentFile in txtFiles)
                {
                    using (StreamReader sr = new StreamReader(currentFile))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {

                            Console.WriteLine(line);

                            conn.Open();

                            using (SqlCommand cmd = new SqlCommand("TimbraturaDataImport", conn))
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.Add("@BadgeTimbrature", SqlDbType.VarChar).Value = line;
                                cmd.ExecuteNonQuery();
                            }
                            conn.Close();
                            counter++;
                        }

                        sr.Close();
                        string filenameWithoutPath = Path.GetFileName(currentFile);

                        var fullPath = Path.Combine("\\\\192.168.96.35\\Users\\Time\\", filenameWithoutPath);
                        var fullPath1 = Path.Combine("\\\\192.168.96.35\\Users\\Time\\Backup\\", filenameWithoutPath);

                        File.Move(fullPath, fullPath1);
                    }
                }
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Uploading Done - '" + counter + "' rows inserted.");
                Console.WriteLine("Last Updated Time: " + DateTime.Now);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Next Update in 10 mins '" + DateTime.Now.AddMinutes(10) + "'");

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        private static void FirstShift(object o)
        {
            SendEmail();
        }
        private static void SecondShift(object o)
        {

            SendEmail();
        }

        //SQL Loading Data First SHIFT
        private static void LoadDataFirst()
        {
            //DECLARE BadgeValues
            int _bdgAmmini = 0;
            int _bdgTess = 0;
            int _bdgStiro = 0;
            int _bdgConfA = 0;
            int _bdgConfB = 0;
            //DECLARE AngajatiViewLastMonthValues
            int _anaAmmini = 0;
            int _anaTess = 0;
            int _anaStiro = 0;
            int _anaConfA = 0;
            int _anaConfB = 0;
         

            _bdgAmmini = LoadBadgeFirst(_bdgAmmini, "AMMINISTRAZIONE");
            _bdgTess = LoadBadgeFirst(_bdgTess, "TESSITURA");
            _bdgStiro = LoadBadgeFirst(_bdgStiro, "STIRO");
            _bdgConfA = LoadBadgeFirst(_bdgConfA, "CONFEZIONE A");
            _bdgConfB = LoadBadgeFirst(_bdgConfB, "CONFEZIONE B");

            _anaAmmini = LoadAnagraficheFirst(_anaAmmini, 5);
            _anaTess = LoadAnagraficheFirst(_anaTess, 3);
            _anaStiro = LoadAnagraficheFirst(_anaStiro, 1);
            _anaConfA = LoadAnagraficheFirst(_anaConfA, 2);
            _anaConfB = LoadAnagraficheFirst(_anaConfB, 14);

          
            //DECLARE AssentValues
            int _assAmmini = _anaAmmini - _bdgAmmini;
            int _assTess = _anaTess - _bdgAmmini;
            int _assStiro = _anaStiro - _bdgStiro;
            int _assConfA = _anaConfA - _bdgConfA;
            int _assConfB = _anaConfB - _bdgConfB;
            //DECLARE PercentValues
            string _perAmmini = (decimal.Divide(_assAmmini,_anaAmmini)*100).ToString("n2");

            //_perAmmini = _assAmmini/_anaAmmini;
            string _perStiro = (decimal.Divide(_assStiro,_anaStiro)*100).ToString("n2");
            string _perConfA = (decimal.Divide(_assConfA,_anaConfA)*100).ToString("n2");
            string _perConfB = (decimal.Divide(_assConfB,_anaConfB)*100).ToString("n2");
            int _perTess = 0;

            int _totale = _assAmmini + _assStiro + _assConfA + _assConfB;

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Amministrazione: " + _bdgAmmini + " " + "Tessitura: " + _bdgTess + " " + "Stiro: " + _bdgStiro + " " + "Confezione A: " + _bdgConfA + " " + "Confezione B: " + _bdgConfB);
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("Amministrazione: " + _anaAmmini + " " + "Tessitura: " + _anaTess + " " + " " + "Stiro: " + _anaStiro + " " + "Confezione A: " + _anaConfA + " " + "Confezione B: " + _anaConfB);


            //MailMessage mail = new MailMessage();
            //mail.From = new MailAddress("noreply@olimpias.rs", "ASSENTEISMO GIORNALIERO - '"+DateTime.Now.ToShortDateString()+"'");
            //mail.To.Add("pnikolic@olimpias.rs");
            //mail.Subject = "ASSENTEISMO GIORNALIERO - Report First Shift ";
            //StringBuilder sb = new StringBuilder();

            //sb.AppendLine("<html>");
            //sb.AppendLine("<head>");
            //sb.AppendLine("</head>");
            //sb.AppendLine("<body>");
            //sb.AppendLine("<table style='font-family: arial;'>");
            //sb.AppendLine("<tr>");
            //sb.AppendLine("<td colspan='2' style='font-weight:600;color: #293d61;font-size: 14pt;'> ASSENTEISMO GIORNALIERO </td>");
            //sb.AppendLine("<td colspan='1' style='color: #334c79;'> -Oliknit </td>'");
            //sb.AppendLine("<td colspan='2' style='float:right;text-align:right;font-size: 8pt;font-weight: 600;'> Anno 2019 </td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='background: #cecece;line-height: 22px;'>");
            //sb.AppendLine("<td colspan='2'></td>");
            //sb.AppendLine("<td colspan='1'></td>");
            //sb.AppendLine("<td colspan='2' style='font-size:10pt;color:red;text-align:right;padding-right:5px;font-weight:600;'>"+DateTime.Now.ToShortDateString()+"</td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='line-height:30px;background: #f0fafd;'>");
            //sb.AppendLine("<td colspan='2' style='font-weight: 600;color:red;padding-left: 5px;vertical-align: middle;'> REPARTO </td>");
            //sb.AppendLine("<td colspan='1' style='color:red;font-size: 11pt;font-weight: 600;padding-left: 5px;padding-right: 5px;'>% assenteismo</td>");
            //sb.AppendLine("<td colspan='2' style='color:red;font-size: 11pt;font-weight: 600;padding-left: 5px;padding-right: 5px;'>nr persone assenti</td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='background:#d6ebfb;line-height: 25px;'>");
            //sb.AppendLine("<td colspan='2' style='padding-left: 5px;font-weight: 600;'>Confezione A </td>");
            //sb.AppendLine("<td colspan='1' style='text-align:center; font-weight: 600;'>" + _perConfA + "%</td>");
            //sb.AppendLine("<td colspan='2' style='text-align:center; font-weight: 600;'>" + _assConfA + "</td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='line-height: 25px;background: #f0fafd;'>");
            //sb.AppendLine("<td colspan='2' style='padding-left: 5px;font-weight: 600;'>Confezione B</td>");
            //sb.AppendLine("<td colspan='1' style='text-align:center; font-weight: 600;'>" + _perConfB + "%</td>");
            //sb.AppendLine("<td colspan='2' style='text-align:center; font-weight: 600;'>" + _assConfB + "</td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='background: #d6ebfb;line-height: 25px;'>");
            //sb.AppendLine("<td colspan='2' style='padding-left: 5px;font-weight: 600;'> Stiro </td>");
            //sb.AppendLine("<td colspan='1' style='text-align:center;font-weight: 600;'>" + _perStiro + "%</td>");
            //sb.AppendLine("<td colspan='2' style='text-align:center;font-weight: 600;'>" + _assStiro + "</td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='line-height: 25px;background: #f0fafd;'>");
            //sb.AppendLine("<td colspan='2' style='padding-left: 5px;font-weight: 600;'>Amministrazione</td>");
            //sb.AppendLine("<td colspan='1' style='text-align:center; font-weight: 600;'>" + _perAmmini + "%</td>");
            //sb.AppendLine("<td colspan='2' style='text-align:center; font-weight: 600;'>" + _assAmmini + "</td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='line-height:30px;background:#acd7f7;'>");
            //sb.AppendLine("<td colspan='3' style='padding-left:5px;font-weight: 600;'>Totale</td>");
            //sb.AppendLine("<td colspan='2' style='text-align:center;font-weight: 600;'>" +_totale+ "</td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("</table>");
            //sb.AppendLine("</body>");
            //sb.AppendLine("</html>");

            //mail.Body = sb.ToString();
            //mail.IsBodyHtml = true;
            //SmtpClient smtp = new SmtpClient("mail.olimpias.it");
            //smtp.Port = 25;
            //smtp.Send(mail);
        }
        public static int LoadBadgeFirst(int count, string Departament)
        {
            try
            {
                SqlConnection conn = new SqlConnection("Data Source=KNSQL2014;Initial Catalog=WbmOlimpiasHR;User=sa;Password=onlyouolimpias");
                conn.Open();
                string query = "SELECT COUNT(DISTINCT dbo.BadgeView.IdRnum) From dbo.BadgeView INNER JOIN dbo.AngajatiViewLastMonth ON dbo.BadgeView.IdRnum = dbo.AngajatiViewLastMonth.Marca INNER JOIN dbo.Departamente ON dbo.AngajatiViewLastMonth.IdDepartament = dbo.Departamente.Id INNER JOIN dbo.Linii ON dbo.AngajatiViewLastMonth.IdLinie = dbo.Linii.Id WHERE dbo.BadgeView.Time >= '06:00:00' AND dbo.BadgeView.Time <= '14:00:00' AND dbo.BadgeView.Departament = '"+Departament+"' and CONCAT('20',dbo.BadgeView.Year) = YEAR(GETDATE()) and dbo.BadgeView.Month = MONTH(GETDATE()) AND dbo.BadgeView.Day = DAY(GETDATE()) AND[IO] = 0";
                SqlCommand cmd = new SqlCommand(query, conn);
                count = Convert.ToInt32(cmd.ExecuteScalar());
                conn.Close();
                return count;
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static int LoadAnagraficheFirst(int count, int IdDepartament)
        {
            try
            {

                SqlConnection conn = new SqlConnection("Data Source=KNSQL2014;Initial Catalog=WbmOlimpiasHR;User=sa;Password=onlyouolimpias");
                conn.Open();
                string query = string.Empty;
                if (IdDepartament==1)
                {
                    query = "SELECT COUNT(DISTINCT dbo.AngajatiViewLastMonth.Marca) FROM dbo.AngajatiViewLastMonth " +
                        "INNER JOIN dbo.Departamente ON dbo.AngajatiViewLastMonth.IdDepartament = dbo.Departamente.Id " +
                        "INNER JOIN dbo.Linii ON dbo.AngajatiViewLastMonth.IdLinie = dbo.Linii.Id " +
                        "INNER JOIN dbo.PosturiDeLucru ON dbo.AngajatiViewLastMonth.IdPostDeLucru = dbo.PosturiDeLucru.Id " +
                        "WHERE dbo.AngajatiViewLastMonth.IdDepartament = '" + IdDepartament + "' AND dbo.AngajatiViewLastMonth.DataLichidarii = '0001-01-01' AND dbo.AngajatiViewLastMonth.IdLinie=20 OR AngajatiViewLastMonth.IdLinie=21 OR AngajatiViewLastMonth.IdLinie=22 AND dbo.AngajatiViewLastMonth.IdLinie <> 30 AND AngajatiViewLastMonth.IdLinie <> 33 AND AngajatiViewLastMonth.IdLinie <> 37 AND AngajatiViewLastMonth.IdLinie <> 39";
                }
                else
                {
                    query = "SELECT COUNT(DISTINCT dbo.AngajatiViewLastMonth.Marca) FROM dbo.AngajatiViewLastMonth " +
                        "INNER JOIN dbo.Departamente ON dbo.AngajatiViewLastMonth.IdDepartament = dbo.Departamente.Id " +
                        "INNER JOIN dbo.Linii ON dbo.AngajatiViewLastMonth.IdLinie = dbo.Linii.Id " +
                        "INNER JOIN dbo.PosturiDeLucru ON dbo.AngajatiViewLastMonth.IdPostDeLucru = dbo.PosturiDeLucru.Id " +
                        "WHERE dbo.AngajatiViewLastMonth.IdDepartament = '" + IdDepartament + "' AND dbo.AngajatiViewLastMonth.DataLichidarii = '0001-01-01' AND dbo.AngajatiViewLastMonth.IdLinie <> 30 AND AngajatiViewLastMonth.IdLinie <> 33 AND AngajatiViewLastMonth.IdLinie <> 37 AND AngajatiViewLastMonth.IdLinie <> 39";

                }
                SqlCommand cmd = new SqlCommand(query, conn);
                count = Convert.ToInt32(cmd.ExecuteScalar());
                conn.Close();

                return count;

            }
            catch (Exception)
            {

                throw;
            }
        }



        //SQL Loading Data Second SHIFT
        private static void LoadDataSecond()
        {
            //DECLARE BadgeValues
            int _bdgAmmini = 0;
            int _bdgTess = 0;
            int _bdgStiro = 0;
            int _bdgConfA = 0;
            int _bdgConfB = 0;
            //DECLARE AngajatiViewLastMonthValues
            int _anaAmmini = 0;
            int _anaTess = 0;
            int _anaStiro = 0;
            int _anaConfA = 0;
            int _anaConfB = 0;


            //_bdgAmmini = LoadBadgeFirst(_bdgAmmini, "AMMINISTRAZIONE");
            //_bdgTess = LoadBadgeFirst(_bdgTess, "TESSITURA");
            _bdgStiro = LoadBadgeFirst(_bdgStiro, "STIRO");
            //_bdgConfA = LoadBadgeFirst(_bdgConfA, "CONFEZIONE A");
            //_bdgConfB = LoadBadgeFirst(_bdgConfB, "CONFEZIONE B");

            //_anaAmmini = LoadAnagraficheFirst(_anaAmmini, 5);
            //_anaTess = LoadAnagraficheFirst(_anaTess, 3);
            _anaStiro = LoadAnagraficheFirst(_anaStiro, 1);
            //_anaConfA = LoadAnagraficheFirst(_anaConfA, 2);
            //_anaConfB = LoadAnagraficheFirst(_anaConfB, 14);

            //DECLARE PercentValues
            //int _perAmmini = _anaAmmini / _bdgAmmini;
            //int _perTess =   _anaTess / _bdgAmmini;
            int _perStiro = _anaStiro / _bdgStiro;
            //int _perConfA = _anaConfA / _bdgConfA;
            //int _perConfB = _anaConfB / _bdgConfB;

            //DECLARE AssentValues
            //int _assAmmini = _anaAmmini - _bdgAmmini;
            //int _assTess = _anaTess - _bdgAmmini;
            int _assStiro = _anaStiro - _bdgStiro;
            //int _assConfA = _anaConfA - _bdgConfA;
            //int _assConfB = _anaConfB - _bdgConfB;
            int _totale = _assStiro;

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Amministrazione: " + _bdgAmmini + " " + "Tessitura: " + _bdgTess + " " + "Stiro: " + _bdgStiro + " " + "Confezione A: " + _bdgConfA + " " + "Confezione B: " + _bdgConfB);
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("Amministrazione: " + _anaAmmini + " " + "Tessitura: " + _anaTess + " " + " " + "Stiro: " + _anaStiro + " " + "Confezione A: " + _anaConfA + " " + "Confezione B: " + _anaConfB);


            //MailMessage mail = new MailMessage();
            //mail.From = new MailAddress("noreply@olimpias.rs", "ASSENTEISMO GIORNALIERO - '" + DateTime.Now.ToShortDateString() + "'");
            //mail.To.Add("pnikolic@olimpias.rs");
            //mail.Subject = "ASSENTEISMO GIORNALIERO - Report Second Shift ";
            //StringBuilder sb = new StringBuilder();

            //sb.AppendLine("<html>");
            //sb.AppendLine("<head>");
            //sb.AppendLine("</head>");
            //sb.AppendLine("<body>");
            //sb.AppendLine("<table style='font-family: arial;'>");
            //sb.AppendLine("<tr>");
            //sb.AppendLine("<td colspan='2' style='font-weight:600;color: #293d61;font-size: 14pt;'> ASSENTEISMO GIORNALIERO </td>");
            //sb.AppendLine("<td colspan='1' style='color: #334c79;'> -Oliknit </td>'");
            //sb.AppendLine("<td colspan='2' style='float:right;text-align:right;font-size: 8pt;font-weight: 600;'> Anno 2019 </td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='background: #cecece;line-height: 22px;'>");
            //sb.AppendLine("<td colspan='2'></td>");
            //sb.AppendLine("<td colspan='1'></td>");
            //sb.AppendLine("<td colspan='2' style='font-size:10pt;color:red;text-align:right;padding-right:5px;font-weight:600;'>" + DateTime.Now.ToShortDateString() + "</td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='line-height:30px;background: #f0fafd;'>");
            //sb.AppendLine("<td colspan='2' style='font-weight: 600;color:red;padding-left: 5px;vertical-align: middle;'> REPARTO </td>");
            //sb.AppendLine("<td colspan='1' style='color:red;font-size: 11pt;font-weight: 600;padding-left: 5px;padding-right: 5px;'>% assenteismo</td>");
            //sb.AppendLine("<td colspan='2' style='color:red;font-size: 11pt;font-weight: 600;padding-left: 5px;padding-right: 5px;'>nr persone assenti</td>");
            //sb.AppendLine("</tr>");
            ////sb.AppendLine("<tr style='background:#d6ebfb;line-height: 25px;'>");
            ////sb.AppendLine("<td colspan='2' style='padding-left: 5px;font-weight: 600;'>Confezione A </td>");
            ////sb.AppendLine("<td colspan='1' style='text-align:center; font-weight: 600;'>"+_perConfA+"%</td>");
            ////sb.AppendLine("<td colspan='2' style='text-align:center; font-weight: 600;'>" + _assConfA + "</td>");
            ////sb.AppendLine("</tr>");
            ////sb.AppendLine("<tr style='line-height: 25px;background: #f0fafd;'>");
            ////sb.AppendLine("<td colspan='2' style='padding-left: 5px;font-weight: 600;'>Confezione B</td>");
            ////sb.AppendLine("<td colspan='1' style='text-align:center; font-weight: 600;'>" + _perConfB + "%</td>");
            ////sb.AppendLine("<td colspan='2' style='text-align:center; font-weight: 600;'>" + _assConfB + "</td>");
            ////sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='background: #d6ebfb;line-height: 25px;'>");
            //sb.AppendLine("<td colspan='2' style='padding-left: 5px;font-weight: 600;'> Stiro </td>");
            //sb.AppendLine("<td colspan='1' style='text-align:center;font-weight: 600;'>" + _perStiro + "%</td>");
            //sb.AppendLine("<td colspan='2' style='text-align:center;font-weight: 600;'>" + _assStiro + "</td>");
            //sb.AppendLine("</tr>");
            ////sb.AppendLine("<tr style='line-height: 25px;background: #f0fafd;'>");
            ////sb.AppendLine("<td colspan='2' style='padding-left: 5px;font-weight: 600;'>Amministrazione</td>");
            ////sb.AppendLine("<td colspan='1' style='text-align:center; font-weight: 600;'>" + _perAmmini + "%</td>");
            ////sb.AppendLine("<td colspan='2' style='text-align:center; font-weight: 600;'>" + _assAmmini + "</td>");
            ////sb.AppendLine("</tr>");
            //sb.AppendLine("<tr style='line-height:30px;background:#acd7f7;'>");
            //sb.AppendLine("<td colspan='3' style='padding-left:5px;font-weight: 600;'>Totale</td>");
            //sb.AppendLine("<td colspan='2' style='text-align:center;font-weight: 600;'>" + _totale + "</td>");
            //sb.AppendLine("</tr>");
            //sb.AppendLine("</table>");
            //sb.AppendLine("</body>");
            //sb.AppendLine("</html>");

            //mail.Body = sb.ToString();
            //mail.IsBodyHtml = true;
            //SmtpClient smtp = new SmtpClient("mail.olimpias.it");
            //smtp.Port = 25;
            //smtp.Send(mail);
        }
        public static int LoadBadgeSecond(int count, string Departament)
        {
            try
            {

                SqlConnection conn = new SqlConnection("Data Source=KNSQL2014;Initial Catalog=WbmOlimpiasHR;User=sa;Password=onlyouolimpias");
                conn.Open();
                string query = "SELECT COUNT(DISTINCT dbo.BadgeView.IdRnum) From dbo.BadgeView INNER JOIN dbo.AngajatiViewLastMonth ON dbo.BadgeView.IdRnum = dbo.AngajatiViewLastMonth.Marca INNER JOIN dbo.Departamente ON dbo.AngajatiViewLastMonth.IdDepartament = dbo.Departamente.Id INNER JOIN dbo.Linii ON dbo.AngajatiViewLastMonth.IdLinie = dbo.Linii.Id WHERE dbo.BadgeView.Time >= '14:00:00' AND dbo.BadgeView.Time <= '22:00:00' AND dbo.BadgeView.Departament = '" + Departament + "' and CONCAT('20',dbo.BadgeView.Year) = YEAR(GETDATE()) and dbo.BadgeView.Month = MONTH(GETDATE()) AND dbo.BadgeView.Day = DAY(GETDATE()) AND[IO] = 0";
                SqlCommand cmd = new SqlCommand(query, conn);
                count = Convert.ToInt32(cmd.ExecuteScalar());
                conn.Close();
                return count;

            }
            catch (Exception)
            {
                throw;
            }
        }
        public static int LoadAnagraficheSecond(int count, int IdDepartament)
        {
            try
            {

                SqlConnection conn = new SqlConnection("Data Source=KNSQL2014;Initial Catalog=WbmOlimpiasHR;User=sa;Password=onlyouolimpias");
                conn.Open();
                string query = string.Empty;
                if (IdDepartament == 1)
                {
                    query = "SELECT COUNT(DISTINCT dbo.AngajatiViewLastMonth.Marca) FROM dbo.AngajatiViewLastMonth INNER JOIN dbo.Departamente ON dbo.AngajatiViewLastMonth.IdDepartament = dbo.Departamente.Id INNER JOIN dbo.Linii ON dbo.AngajatiViewLastMonth.IdLinie = dbo.Linii.Id INNER JOIN dbo.PosturiDeLucru ON dbo.AngajatiViewLastMonth.IdPostDeLucru = dbo.PosturiDeLucru.Id WHERE dbo.AngajatiViewLastMonth.IdDepartament = '" + IdDepartament + "' AND dbo.AngajatiViewLastMonth.DataLichidarii = '0001-01-01' AND dbo.AngajatiViewLastMonth.IdLinie=26 OR AngajatiViewLastMonth.IdLinie=27 OR AngajatiViewLastMonth.IdLinie=28 AND dbo.AngajatiViewLastMonth.IdLinie <> 30 AND AngajatiViewLastMonth.IdLinie <> 33 AND AngajatiViewLastMonth.IdLinie <> 37 AND AngajatiViewLastMonth.IdLinie <> 39";
                }
                else
                {
                    query = "SELECT COUNT(DISTINCT dbo.AngajatiViewLastMonth.Marca) FROM dbo.AngajatiViewLastMonth INNER JOIN dbo.Departamente ON dbo.AngajatiViewLastMonth.IdDepartament = dbo.Departamente.Id INNER JOIN dbo.Linii ON dbo.AngajatiViewLastMonth.IdLinie = dbo.Linii.Id INNER JOIN dbo.PosturiDeLucru ON dbo.AngajatiViewLastMonth.IdPostDeLucru = dbo.PosturiDeLucru.Id WHERE dbo.AngajatiViewLastMonth.IdDepartament = '" + IdDepartament + "' AND dbo.AngajatiViewLastMonth.DataLichidarii = '0001-01-01' AND dbo.AngajatiViewLastMonth.IdLinie <> 30 AND AngajatiViewLastMonth.IdLinie <> 33 AND AngajatiViewLastMonth.IdLinie <> 37 AND AngajatiViewLastMonth.IdLinie <> 39";

                }
                SqlCommand cmd = new SqlCommand(query, conn);
                count = Convert.ToInt32(cmd.ExecuteScalar());
                conn.Close();

                return count;

            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}