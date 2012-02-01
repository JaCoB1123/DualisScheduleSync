using System;
using System.Windows.Forms;
using System.Net;
using System.Text.RegularExpressions;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Configuration;

namespace DHBWloginTest
{
    public partial class DualisConnector : Form
    {
        string username = "", password = "";
        string icalpath = "", xmlpath = "";
        string format = "both";
        int monthspast = 2, monthsfuture = 6;
        bool ftpenabled = false;
        string ftpuser = "", ftppassword = "", ftpserver = "", ftpfilename = "";
        bool gmailenabled = false;
        string gmailuser = "", gmailpassword = "", gmailcalendarid = "";
        string config = "./DHBWloginTest.config";
        private CookieCollection cookies = new CookieCollection();
        private CookieContainer cookieContainer = new CookieContainer();
        string startUrl = "https://dualis.dhbw.de/scripts/mgrqcgi";
        List<String> urls = new List<string>();
        List<String> links = new List<String>();
        static string cookie = "cnsc=" + new Random().Next(1000000000).ToString();
        string DUALIS_KALENDER_URL = "https://dualis.dhbw.de/scripts/mgrqcgi?APPNAME=CampusNet&PRGNAME=MONTH&ARGUMENTS=";

        public DualisConnector()
        {
            InitializeComponent();
            progressBar1.MarqueeAnimationSpeed = 200;
            progressBar1.Style = ProgressBarStyle.Marquee;
            this.Shown += new EventHandler(DualisConnector_Shown);
        }

        void DualisConnector_Shown(object sender, EventArgs e)
        {
            if (!InitializeConfig())
            {
                return;
            }
            string args = null;
            int errors = 0;
            while (string.IsNullOrEmpty(args) && errors++ < 10)
            {
                this.Text = (errors).ToString() +  ". try to login as: " + username;
                login(username, password);
            }
            if (string.IsNullOrEmpty(args))
            {
                MessageBox.Show("Login failed!");
                Close();
            }
            this.Text = "Logged in:" + args;

            DateTime thisMonth = DateTime.Today.AddDays(1 - DateTime.Today.Day);
            thisMonth = new DateTime(thisMonth.Year, thisMonth.Month, 1);

            String docs = string.Empty;
            for (int i = -monthspast; i < monthsfuture; i++)
            {
                this.Text = "Loading:" + thisMonth.AddMonths(i).ToString("dd.MM.yyyy");
                docs += Load(DUALIS_KALENDER_URL + args + ",-N000031,-A" + thisMonth.AddMonths(i).ToString("dd.MM.yyyy"));
            }
            //@"<a title=""(?<start>\d\d:\d\d) - (?<end>\d\d:\d\d) / (?:(?<room>[^/]+) /)? (?<name>[^""]+)""[^>]*>"
            foreach (Match m in Regex.Matches(docs, @"<a title=""([^"":.]{2}[:.][^""]+)""[^>]*>"))
            {
                this.Text = "Matching";
                links.Add(m.Groups[1].Value);
            }

            Calendar c = new Calendar(links);
            c.SaveToFile(icalpath, xmlpath, format);
            this.Text = "Saving to " + icalpath + " / " + xmlpath;
            if (ftpenabled)
            {
                this.Text = "Saving to FTP";
                c.SaveToFTP(ftpfilename, format, ftpserver, ftpuser, ftppassword);
            }
            if (gmailenabled)
            {
                this.Text = "Saving to GMail";
                c.SaveToGmail(gmailuser, gmailpassword, gmailcalendarid);
            }
            this.Text = "Finished!";
            Close();
        }

        private bool InitializeConfig()
        {
            if (File.Exists(config))
            {
                XDocument xd = XDocument.Load(config);
                username = (from c in xd.Descendants("username") select c.Value).First();
                password = (from c in xd.Descendants("password") select c.Value).First();
                icalpath = (from c in xd.Descendants("icalpath") select c.Value).First();
                xmlpath = (from c in xd.Descendants("xmlpath") select c.Value).First();
                format = (from c in xd.Descendants("format") select c.Value).First();
                monthsfuture = (from c in xd.Descendants("monthsfuture") select Convert.ToInt32(c.Value)).First();
                monthspast = (from c in xd.Descendants("monthspast") select Convert.ToInt32(c.Value)).First();
                ftppassword = (from c in xd.Descendants("ftppassword") select c.Value).First();
                ftpserver = (from c in xd.Descendants("ftpserver") select c.Value).First();
                ftpuser = (from c in xd.Descendants("ftpuser") select c.Value).First();
                ftpenabled = (from c in xd.Descendants("ftpenabled") select Convert.ToBoolean(c.Value)).First();
                ftpfilename = (from c in xd.Descendants("ftpfilename") select c.Value).First();
                gmailenabled = (from c in xd.Descendants("gmailenabled") select Convert.ToBoolean(c.Value)).First();
                gmailpassword = (from c in xd.Descendants("gmailpassword") select c.Value).First();
                gmailuser = (from c in xd.Descendants("gmailuser") select c.Value).First();
                gmailcalendarid = (from c in xd.Descendants("gmailcalendarid") select c.Value).First();
                return true;
            }
            XDocument doc = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                new XElement("config",
                    new XElement("username", "a00000@hb.dhbw-stuttgart.de"),
                    new XElement("password", "PASSWORD"),
                    new XElement("icalpath", "./dualis.ics"),
                    new XElement("xmlpath", "./dualis.xml"),
                    new XComment("Format to export (xml|ical|both)"),
                    new XElement("format", "both"),
                    new XElement("monthspast", 1),
                    new XElement("monthsfuture", 2),
                    new XElement("ftpenabled",false),
                    new XElement("ftpuser","USER"),
                    new XElement("ftppassword","PASSWORD"),
                    new XElement("ftpserver","example.com"),
                    new XElement("ftpfilename","dualis"),
                    new XElement("gmailenabled","false"),
                    new XElement("gmailuser","user@googlemail.com"),
                    new XElement("gmailpassword","PASSWORD"),
                    new XComment("Not used at the moment. Do NOT enter your main Calendar ID."),
                    new XElement("gmailcalendarid","calendarid")
                )
            );
            doc.Save(config);
            MessageBox.Show(string.Format("Please edit your Settings at:\n{0}", config));
            return false;
        }
     
        private String readResponse(HttpWebRequest conn)
        {
            String ret = "";
            using (HttpWebResponse resp = (HttpWebResponse)conn.GetResponse())
            {
				using (StreamReader rd = new StreamReader(resp.GetResponseStream()))
				{
				    ret = rd.ReadToEnd();
				}
            }
            return ret;
        }

        private String login(String username, String passwort)
        {
            String cookie = null;
            String data = "usrname=" + username + "&pass=" + passwort + "&APPNAME=CampusNet&PRGNAME=LOGINCHECK&ARGUMENTS=clino%2Cusrname%2Cpass%2Cmenuno%2Cpersno%2Cbrowser%2Cplatform&clino=000000000000001&menuno=000000&persno=00000000&browser=&platform=";

            HttpWebRequest conn = connect(startUrl);
            conn.Method = "POST";
            try
            {
                using (StreamWriter wr = new StreamWriter(conn.GetRequestStream()))
                {
                    wr.Write(data);
                    wr.Close();
                }
            }
            catch (WebException)
            {
                return null;
            }

            Regex r = new Regex("ARGUMENTS=([^,]+),", RegexOptions.Compiled);
            Match m = r.Match(getHeader(conn));

            if (m == null)
            {
                return null;
            }
            cookie = m.Groups[1].Value;
            return cookie;
        }
        private String Load(String uri)
        {
            return readResponse(connect(uri));
        }
        private HttpWebRequest connect(String surl)
        {
            Uri url = new Uri(surl);
            HttpWebRequest conn = (HttpWebRequest)WebRequest.Create(url);
            conn.Headers.Add("Accept-Charset", "UTF-8");
            conn.ContentType = "application/x-www-form-urlencoded;charset=UTF-8";
            conn.UserAgent = "Bot";
            conn.Headers.Add("Cookie", cookie);
            return conn;
        }
        private static String getHeader(HttpWebRequest con)
        {
            String ret = "";
            using (HttpWebResponse resp = (HttpWebResponse)con.GetResponse())
            {
                foreach (String field in resp.Headers.AllKeys)
                {
                    ret += field + " : " + resp.Headers[field] + "\n";
                }
            }
            return ret;
        }
    }
}
