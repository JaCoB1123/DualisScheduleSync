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

namespace DualisScheduleSync
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
                args = login(); 
			}
			if (string.IsNullOrEmpty(args))
			{
				MessageBox.Show("Login failed!");
				Close();
			}

			DateTime thisMonth = DateTime.Today.AddDays(1 - DateTime.Today.Day);
			thisMonth = new DateTime(thisMonth.Year, thisMonth.Month, 1);

            List<String> links = new List<String>();
            String docs = loadCalendarMonths(args, thisMonth);
			foreach (Match m in Regex.Matches(docs, @"<a title=""([^"":.]{2}[:.][^""]+)""[^>]*>"))
			{
				links.Add(m.Groups[1].Value);
			}

            Calendar c = new Calendar(links);
            c.Save(getSetting("format"), bool.Parse(getSetting("ftpenabled")), bool.Parse(getSetting("gmailenabled")));
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
			String data = "usrname=" + getSetting("username") + "&pass=" + getSetting("password") 
                + "&APPNAME=CampusNet&PRGNAME=LOGINCHECK&ARGUMENTS=clino%2Cusrname%2Cpass%2Cmenuno%2Cpersno%2Cbrowser%2Cplatform"
                + "&clino=000000000000001&menuno=000000&persno=00000000&browser=&platform=";

			HttpWebRequest conn = getWebRequest(startUrl);
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

            return getHeader(conn);
		}

        private static string getHeader(HttpWebRequest conn)
        {
            using (HttpWebResponse resp = (HttpWebResponse)conn.GetResponse())
            {
                Match m = Regex.Match(resp.Headers.ToString(), "ARGUMENTS=([^,]+),", RegexOptions.Compiled);

                if (m == null)
                {
                    return null;
                }

                return m.Groups[1].Value;
            }
        }

        private String loadCalendarMonths(string args, DateTime thisMonth)
		{
			String docs = string.Empty;
			for (int i = -int.Parse(getSetting("monthspast")); i < int.Parse(getSetting("monthsfuture")); i++)
			{
                String uri = calendarUrl + args + ",-N000031,-A" + thisMonth.AddMonths(i).ToString("dd.MM.yyyy");
			    docs += readResponse(getWebRequest(uri));             
			}
            return docs;
		}

		private HttpWebRequest getWebRequest(String surl)
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
