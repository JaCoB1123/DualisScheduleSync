using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Calendar.v3.Data;

namespace DualisScheduleSync
{
    static class Web
    {
        private static string cookie = "cnsc=" + new Random().Next(1000000000).ToString();

        public static string LoadHttpText(string uri)
        {
            string headers;
            return LoadHttpText(uri, null, out headers);
        }
        public static string LoadHttpText(string uri, string postData)
        {
            string headers;
            return LoadHttpText(uri, postData, out headers);
        }
        public static string LoadHttpText(string uri, string postData, out string headers) {
            HttpWebRequest conn = getWebRequest(uri);

            if (!string.IsNullOrEmpty(postData))
            {
                conn.Method = "POST";
                try
                {
                    using (StreamWriter wr = new StreamWriter(conn.GetRequestStream()))
                    {
                        wr.Write(postData);
                        wr.Close();
                    }
                }
                catch (WebException)
                {
                    headers = null;
                    return null;
                }
            } 
            
            using (HttpWebResponse resp = (HttpWebResponse)conn.GetResponse())
            {
                headers = resp.Headers.ToString();
                using (StreamReader rd = new StreamReader(resp.GetResponseStream()))
                {
                    return rd.ReadToEnd();
                }
            }
        }

        public static HttpWebRequest getWebRequest(String surl)
        {
            Uri url = new Uri(surl);
            HttpWebRequest conn = (HttpWebRequest)WebRequest.Create(url);
            conn.Headers.Add("Accept-Charset", "UTF-8");
            conn.ContentType = "application/x-www-form-urlencoded;charset=UTF-8";
            conn.UserAgent = "Bot";
            conn.Headers.Add("Cookie", cookie);
            return conn;
        }

        public static void AppendTo(this Event e, ref StringBuilder ical)
        {
            ical.AppendLine("BEGIN:VEVENT");
            ical.AppendLine("UID:" + e.Id);
            ical.AppendLine("SUMMARY:" + e.Summary);
            ical.AppendLine("DTSTART:" + e.Start.DateTime);
            ical.AppendLine("DTEND:" + e.End.DateTime);
            ical.AppendLine("DTSTAMP:" + DateTime.Now.ToISO());
            ical.AppendLine("END:VEVENT");
        }
    }
}
