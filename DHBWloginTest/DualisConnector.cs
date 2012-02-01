using System;
using System.Windows.Forms;
using System.Net;
using System.Text.RegularExpressions;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace DHBWloginTest
{
	public partial class DualisConnector : Form
    {
        static string cookie = "cnsc=" + new Random().Next(1000000000).ToString();
        string startUrl { 
            get { return "https://dualis.dhbw.de/scripts/mgrqcgi"; } 
        }
        string calendarUrl
        {
            get { return startUrl + "?APPNAME=CampusNet&PRGNAME=MONTH&ARGUMENTS="; }
        }

		public DualisConnector()
		{
			InitializeComponent();
			progressBar1.MarqueeAnimationSpeed = 200;
			progressBar1.Style = ProgressBarStyle.Marquee;
			this.Shown += new EventHandler(DualisConnector_Shown);
		}

        internal static string getSetting(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }

		void DualisConnector_Shown(object sender, EventArgs e)
		{
			string args = null;
			int errors = 0;
			while (string.IsNullOrEmpty(args) && errors++ < 10)
			{
                this.Text = (errors).ToString() + ". try to login as: " + getSetting("username");
				args = login(); 
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
			for (int i = -int.Parse(getSetting("monthspast")); 
                i < int.Parse(getSetting("monthsfuture")); i++)
			{
                this.Text = "Loading:" + thisMonth.AddMonths(i).ToShortDateString();
                docs += loadCalendarMonth(args, thisMonth.AddMonths(i));
			}
			//@"<a title=""(?<start>\d\d:\d\d) - (?<end>\d\d:\d\d) / (?:(?<room>[^/]+) /)? (?<name>[^""]+)""[^>]*>"

            List<String> links = new List<String>();
			foreach (Match m in Regex.Matches(docs, @"<a title=""([^"":.]{2}[:.][^""]+)""[^>]*>"))
			{
				this.Text = "Matching";
				links.Add(m.Groups[1].Value);
			}

			Calendar c = new Calendar(links);
			c.SaveToFile(getSetting("format"));
            this.Text = "Saving to File (" + getSetting("icalpath") + " / " + getSetting("xmlpath") + ")";
			if (bool.Parse(getSetting("ftpenabled")))
			{
				this.Text = "Saving to FTP";
                c.SaveToFTP(getSetting("format"));
			}
			if (bool.Parse(getSetting("gmailenabled")))
			{
				this.Text = "Saving to GMail";
				c.SaveToGmail( );
			}
			this.Text = "Finished!";
			Close();
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

		private String login( )
		{
			String cookie = null;
			String data = "usrname=" + getSetting("username") + "&pass=" + getSetting("password") 
                + "&APPNAME=CampusNet&PRGNAME=LOGINCHECK&ARGUMENTS=clino%2Cusrname%2Cpass%2Cmenuno%2Cpersno%2Cbrowser%2Cplatform&clino=000000000000001&menuno=000000&persno=00000000&browser=&platform=";

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
			String header = "";
			using (HttpWebResponse resp = (HttpWebResponse)conn.GetResponse())
			{
				foreach (String field in resp.Headers.AllKeys)
				{
					header += field + " : " + resp.Headers[field] + "\n";
				}
			}
			Match m = r.Match(header);

			if (m == null)
			{
				return null;
			}
			cookie = m.Groups[1].Value;
			return cookie;
		}
        private String loadCalendarMonth(string args, DateTime thisMonth)
		{
            String uri = calendarUrl + args + ",-N000031,-A" + thisMonth.ToString("dd.MM.yyyy");
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
	}
}
