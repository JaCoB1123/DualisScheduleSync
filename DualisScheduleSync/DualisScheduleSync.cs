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
using System.Threading.Tasks;

namespace DualisScheduleSync
{
	public class DualisScheduleSync
    {
        string startUrl { 
            get { return "https://dualis.dhbw.de/scripts/mgrqcgi"; } 
        }
        string calendarUrl
        {
            get { return startUrl + "?APPNAME=CampusNet&PRGNAME=MONTH&ARGUMENTS="; }
        }

        public string Error { get; private set; }

        internal static string getSetting(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }

		public bool Sync() {
			string args = null;
			int errors = 0;
			while (string.IsNullOrEmpty(args) && errors++ < 10) {
                args = login(); 
			}
			if (string.IsNullOrEmpty(args)) {
                Error = "Login failed!";
                return false;
			}

			DateTime thisMonth = DateTime.Today.AddDays(1 - DateTime.Today.Day);
			thisMonth = new DateTime(thisMonth.Year, thisMonth.Month, 1);

            List<String> links = new List<String>();
            String docs = loadCalendarMonths(args, thisMonth);
			foreach (Match m in Regex.Matches(docs, @"<a title=""([^"":.]{2}[:.][^""]+)""[^>]*>")) {
				links.Add(m.Groups[1].Value);
			}

            Cal c = new Cal(links);
            c.Save(getSetting("format"), bool.Parse(getSetting("ftpenabled")), bool.Parse(getSetting("gmailenabled")));
            return true;
		}

		private String login( ) {
			String data = "usrname=" + getSetting("username") + "&pass=" + getSetting("password") 
                + "&APPNAME=CampusNet&PRGNAME=LOGINCHECK&ARGUMENTS=clino%2Cusrname%2Cpass%2Cmenuno%2Cpersno%2Cbrowser%2Cplatform"
                + "&clino=000000000000001&menuno=000000&persno=00000000&browser=&platform=";
            String headers;

            Web.LoadHttpText(startUrl, data, out headers);

            return getHeader(headers);
		}

        private static string getHeader(string headers) {
            Match m = Regex.Match(headers, "ARGUMENTS=([^,]+),", RegexOptions.Compiled);

            if (m == null) {
                return null;
            }

            return m.Groups[1].Value;
        }

        private String loadCalendarMonths(string args, DateTime thisMonth)
		{
			String docs = string.Empty;
			for (int i = -int.Parse(getSetting("monthspast")); i < int.Parse(getSetting("monthsfuture")); i++) {
                String uri = calendarUrl + args + ",-N000031,-A" + thisMonth.AddMonths(i).ToString("dd.MM.yyyy");
                docs += Web.LoadHttpText(uri);
			}
            return docs;
		}
	}
}
