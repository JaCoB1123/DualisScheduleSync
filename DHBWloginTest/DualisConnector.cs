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

namespace DHBWloginTest
{
    delegate void LoadFinished(object sender, WebBrowserDocumentCompletedEventArgs e);
    public partial class DualisConnector : Form
    {
        string username = "";
        string password = "";
        string icalpath = "";
        string xmlpath = "";
        string format = "";
        int monthspast = 1;
        int monthsfuture = 2;
        string config = "./DHBWloginTest.config";
        WebBrowser wb = new WebBrowser();
        LoadFinished loaded;
        LoadFinished loadFinished;
        private CookieCollection cookies = new CookieCollection();
        private CookieContainer cookieContainer = new CookieContainer();
        string startUrl = "https://dualis.dhbw.de/scripts/mgrqcgi?APPNAME=CampusNet&PRGNAME=EXTERNALPAGES&ARGUMENTS=-N000000000000001,-N,-Awelcome";
        List<String> urls = new List<string>();
        List<String> links = new List<String>();
        

        public DualisConnector()
        {
            InitializeComponent();
            if (InitializeConfig())
            {
                loaded = new LoadFinished(indexLoadFinished);
                loadFinished = (s, ea) => { loaded(s, ea); };
                wb.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(loadFinished);
                progressBar1.MarqueeAnimationSpeed = 200;
                progressBar1.Style = ProgressBarStyle.Marquee;
                wb.Navigate(startUrl);
            }
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
                    new XElement("monthsfuture", 2)
                )
            );
            doc.Save(config);
            MessageBox.Show(string.Format("Please edit your Settings at:\n{0}", config));
            return false;
        }

        void indexLoadFinished(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            loaded = new LoadFinished(loginLoadFinished);
            try
            {
                HtmlElementCollection es = wb.Document.Forms[0].GetElementsByTagName("input");
                es.GetElementsByName("usrname")[0].InnerText = username;
                es.GetElementsByName("pass")[0].InnerText = password;
                wb.Document.GetElementById("logIn_btn").InvokeMember("Click");
            }
            catch
            {
                MessageBox.Show("Failed to get Dualis Website", "Network Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        void loginLoadFinished(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (wb.DocumentText.Contains("Benutzername oder Passwort falsch"))
            {
                MessageBox.Show("Benutzername oder Passwort falsch", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            if (wb.Url.OriginalString.Contains("MLSSTART"))
            {
                String url = wb.Url.OriginalString;
                url = url.Replace("MLSSTART", "MONTH") + ",-A" + "###DATE###" + ",-A,-N000000000000000";
                for (int i = -monthspast; i < monthsfuture; i++) 
                    urls.Add(url.Replace("###DATE###",DateTime.Now.AddDays(1 - DateTime.Now.Day).AddMonths(i).ToString("dd.MM.yyyy")));
                loaded = new LoadFinished(calendarLoadFinished);
                wb.Navigate(urls[0]);
            }
        }

        void calendarLoadFinished(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (urls.Contains(wb.Url.OriginalString))
            {
                foreach (HtmlElement element in wb.Document.GetElementsByTagName("a"))
                    links.Add(element.GetAttribute("title"));
                urls.Remove(wb.Url.OriginalString);
                if (urls.Count > 0)     wb.Navigate(urls[0]);
                else
                {
                    progressBar1.Style = ProgressBarStyle.Blocks;
                    HtmlToIcal(links);
                    this.Close();
                    return;
                }
            }
        }

        void HtmlToIcal(List<String> elements)
        {
            List<String> datesntimes = new List<String>();
            foreach (String e in elements ) {
                if (Regex.IsMatch(e, "^[^\":.]{2}[:.].+$") )    datesntimes.Add(e);
            }

            Calendar c = new Calendar(datesntimes);
            c.SaveToFile(icalpath, xmlpath, format);
        }
    }
}
