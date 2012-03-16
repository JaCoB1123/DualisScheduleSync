using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Net;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Util;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Requests;
using System.Configuration;
using System.Diagnostics;
using System.Threading;

namespace DualisScheduleSync
{
    class Cal
    {

        public static Event NewEvent(DateTime start, DateTime end, String title)
        {
            return NewEvent(start, end, title, null);
        }
        public static Event NewEvent(DateTime start, DateTime end, String title, String place)
        {
            Event e = new Event();
            e.Start = new EventDateTime();
            e.Start.DateTime = start.ToString("yyyy-MM-ddTHH:mm:sszzzz");
            e.End = new EventDateTime();
            e.End.DateTime = end.ToString("yyyy-MM-ddTHH:mm:sszzzz");

            if (title.Contains("/") && string.IsNullOrEmpty(place))
            {
                String[] s = title.Split('/');
                e.Location = s[0].Trim();
                e.Summary = s[1].Trim();
            }
            else
            {
                e.Summary = title.Trim();
                e.Location = place ?? string.Empty;
            }

            e.Creator = new Event.CreatorData();
            e.Creator.DisplayName = "Jan Bader";
            e.Creator.Email = "jan@javil.eu";

            return e;
        }


        public const String version = "2.0";
        List<Event> events = new List<Event>();

        public Cal(List<String> elements)
        {
            AddEvent(NewEvent(DateTime.Now, DateTime.Now.AddMinutes(1), "Updated DHBW Data", "Dualis"));

            DateTime lastDate = DateTime.MinValue;
            foreach (String title in elements)
            {
                if (title.ToCharArray()[2] == '.') {
                    DateTime.TryParse(title, out lastDate);
                } else {
                    DateTime start = lastDate.Add(TimeSpan.Parse(title.Substring(0, 5)));
                    DateTime end = lastDate.Add(TimeSpan.Parse(title.Substring(8, 5)));
                    Event e = NewEvent(start,end,title.Substring(16));
                    AddEvent(e);
                    Console.WriteLine(e);
                }
            }
        }

        public void AddEvent(Event e)
        {
            events.Add(e);
        }

        public void Save(String format, bool ftp, bool gmail)
        {
            SaveToFile(format);
            if (ftp)
            {
                SaveToFTP(format);
            }
            if (gmail)
            {
                SaveToGmail();
            }
        }

        public void SaveToFile(String format)
        {
            if (format == "ics" || format == "both")
            {
                File.WriteAllText(DualisScheduleSync.getSetting("icalpath"), ToIcal());
            }
            if (format == "xml" || format == "both")
            {
                File.WriteAllText(DualisScheduleSync.getSetting("xmlpath"), ToXml());
            }
        }
        public void SaveToFTP(String format)
        {
            String filename = DualisScheduleSync.getSetting("ftpfilename"),
                server = DualisScheduleSync.getSetting("ftpserver"),
                user = DualisScheduleSync.getSetting("ftpuser"),
                password = DualisScheduleSync.getSetting("ftppassword");
            if (format == "both")
            {
                SaveToFTP("xml");
                SaveToFTP("ics");
                return;
            }
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://"+server+"/"+filename+"."+format);
            request.Method = WebRequestMethods.Ftp.UploadFile;
            request.Credentials = new NetworkCredential(user, password);
            byte[] fileContents = null;
            if (format == "xml")
            {
                fileContents = Encoding.UTF8.GetBytes(ToXml());
            }
            else
            {
                fileContents = Encoding.UTF8.GetBytes(ToIcal());
            }
            request.ContentLength = fileContents.Length;

            Stream requestStream = request.GetRequestStream();
            requestStream.Write(fileContents, 0, fileContents.Length);
            requestStream.Close();

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            response.Close();
        }

        public override String ToString() {
            return ToIcal();
        }
        public String ToIcal()
        {
            StringBuilder ical = new StringBuilder();
            ical.AppendLine("BEGIN:VCALENDAR");
            ical.AppendLine("VERSION:" + version);
            ical.AppendLine("PRODID:http://dualis.dhbw.de");

            events.All(a => { a.AppendTo(ref ical); return true; });

            ical.AppendLine("END:VCALENDAR");

            return ical.ToString();
        }

        public String ToXml()
        {
            XDocument doc = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                new XElement("calendar",
                    from e in events
                    orderby e.Start.DateTime
                    select new XElement("event",
                        new XAttribute("ID", e.Id ?? ""),
                        new XAttribute("SUMMARY", e.Summary),
                        new XAttribute("DTSTART", e.Start.DateTime),
                        new XAttribute("DTEND", e.End.DateTime),
                        new XAttribute("DTSTAMP", DateTime.Now.ToISO())
                    )
                )
            );
            return doc.ToString();
        }
        
        public bool SaveToGmail()
        {
            var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            provider.ClientIdentifier = "953955382732.apps.googleusercontent.com";
            provider.ClientSecret = "mYUQelU9LUiozb-0Qw6l0rBK";
            var auth = new OAuth2Authenticator<NativeApplicationClient>(provider, GetAuthorization);

            var srv = new Google.Apis.Calendar.v3.CalendarService(auth);

            Calendar calendar = ClearGMail(srv);
            if (calendar == null)
            {
                return false;
            }
            
            foreach (Event e in events)
            {
                srv.Events.Insert(e, calendar.Id).Fetch();
            }

            return true;
        }

        private static Calendar ClearGMail(CalendarService srv)
        {
            var a = srv.CalendarList.List().Fetch().Items;
            string calID = null;
            foreach (CalendarListEntry c in a)
            {
                if (c.Summary == DualisScheduleSync.getSetting("gmailcalendarname"))
                {
                    calID = c.Id;
                    break;
                }
            }

            Calendar calendar = null;
            if (!string.IsNullOrEmpty(calID))
            {
                calendar = srv.Calendars.Get(calID).Fetch();
                if (calendar != null)
                {
                    var b = srv.Events.List(calendar.Id);
                    b.MaxResults = 10000;
                    b.ShowDeleted = false;
                    b.TimeMin = DualisScheduleSync.thisMonth.AddMonths(-int.Parse(DualisScheduleSync.getSetting("monthspast"))).ToString("yyyy-MM-ddTHH:mm:sszzzz");
                    Events events = b.Fetch();
                    if (events != null && events.Items != null)
                    {
                        foreach (var evententry in events.Items)
                        {
                            srv.Events.Delete(calendar.Id, evententry.Id).Fetch();
                        }
                    }
                }
            }
            
            if (calendar == null)
            {
                calendar = new Calendar();
                calendar.Summary = DualisScheduleSync.getSetting("gmailcalendarname");
                calendar.Description = "Dieser Kalender enthält den Stundenplan laut Dualis";
                calendar.TimeZone = "Europe/Berlin";
                calendar.Location = "Horb am Neckar";
                calendar = srv.Calendars.Insert(calendar).Fetch();
            }

            return calendar;
       }

        private static IAuthorizationState GetAuthorization(NativeApplicationClient arg)
        {
            IAuthorizationState state = new AuthorizationState(new[] { CalendarService.Scopes.Calendar.GetStringValue() });
            state.Callback = new Uri(NativeApplicationClient.OutOfBandCallbackUrl);

            string path = "Q:\\auth.txt";
            if (File.Exists(path))
            {
                
                state.RefreshToken = File.ReadAllText(path);
                arg.RefreshToken(state);
                //state = arg.ProcessUserAuthorization(state.AccessToken, state);
            }
            else
            {

                Uri authUri = arg.RequestUserAuthorization(state);

                // Request authorization from the user (by opening a browser window):
                Process.Start(authUri.ToString());
                Console.Write("  Authorization Code: ");
                string authCode = string.Empty;
                int count = 0;
                while (string.IsNullOrEmpty(authCode) && count < 500)
                {
                    foreach (Process proc in Process.GetProcesses())
                    {
                        if (proc.MainWindowTitle.StartsWith("Success code="))
                        {
                            authCode = proc.MainWindowTitle.Split('=').Last().Split(' ').First();
                        }
                    }
                    Thread.Sleep(50);
                    Application.DoEvents();
                    count++;
                }
                Console.WriteLine();
                // Retrieve the access token by using the authorization code:
                state = arg.ProcessUserAuthorization(authCode, state);
            }

            File.WriteAllText(path, state.RefreshToken);

            return state;
        }
    }
}