﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Net;
using Google.GData.Calendar;
using Google.GData.Client;
using Google.GData.Extensions;

namespace DHBWloginTest
{
    class Calendar
    {
        public const String version = "2.0";
        List<Event> events = new List<Event>();

        public Calendar(List<String> elements)
        {
            AddEvent(new Event(DateTime.Today, "P1D", "Created " + DateTime.Now.TimeOfDay.ToString()));

            DateTime lastDate = DateTime.MinValue;
            foreach (String title in elements)
            {
                if (title.ToCharArray()[2] == '.') {
                    DateTime.TryParse(title, out lastDate);
                } else {
                    DateTime start = lastDate.Add(TimeSpan.Parse(title.Substring(0, 5)));
                    DateTime end = lastDate.Add(TimeSpan.Parse(title.Substring(8, 5)));
                    Event e = new Event(start,end,title.Substring(16));
                    AddEvent(e);
                    Console.WriteLine(e);
                }
            }
        }

        public void AddEvent(Event e) {
            events.Add(e);
        }

        public void SaveToFile(String icalpath, String xmlpath, String format)
        {
            if (format == "ics" || format == "both")
            {
                File.WriteAllText(icalpath, ToIcal());
            }
            if (format == "xml" || format == "both")
            {
                File.WriteAllText(xmlpath, ToXml());
            }
        }
        public void SaveToFTP(String filename, String format, String server, String user, String password)
        {
            if (format == "both")
            {
                SaveToFTP(filename, "xml", server, user, password);
                SaveToFTP(filename, "ics", server, user, password);
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
                    orderby e.id
                    select new XElement("event",
                        new XAttribute("UID", e.uid),
                        new XAttribute("SUMMARY", e.title),
                        new XAttribute("DTSTART", e.begin.ToUniversalTime().ToISO()),
                        new XAttribute("DTEND", (e.end != DateTime.MinValue) ? e.end.ToUniversalTime().ToISO() : ""),
                        new XAttribute("DURATION", e.duration??"" ),
                        new XAttribute("DTSTAMP", DateTime.Now.ToISO())
                    )
                )
            );
            return doc.ToString();
        }
        


        public void SaveToGmail(string user, string password, string calendarID)
        {
            var srv = new Google.GData.Calendar.CalendarService("GoogleCalendarTest");
            srv.setUserCredentials(user, password);

            String calendarUri = ClearGMail(srv);
            EventFeed eF = srv.Query(new EventQuery(calendarUri));
            AtomFeed batchFeed = new AtomFeed(eF);
            foreach (Event e in events)
            {
                EventEntry evt = new EventEntry(e.title, e.title, e.place);
                evt.Times.Add(new Google.GData.Extensions.When(e.begin, e.end));
                evt.Authors.Add(new AtomPerson(AtomPersonType.Author, "Jan Bader"));
                evt.BatchData = new GDataBatchEntryData(GDataBatchOperationType.insert);
                batchFeed.Entries.Add(evt);
            }

            EventFeed batchResultFeed = (EventFeed)srv.Batch(batchFeed, new Uri(eF.Batch));
            bool success =  batchResultFeed.Entries.All(a => a.BatchData.Status.Code == 200 || a.BatchData.Status.Code == 201);

            if (!success) {
                //failed
            }
        }

        private static String ClearGMail(CalendarService srv)
        {
            CalendarQuery cQ = new CalendarQuery("https://www.google.com/calendar/feeds/default/owncalendars/full");
            CalendarFeed cR = srv.Query(cQ);
            foreach (CalendarEntry c in cR.Entries) {
                if (c.Title.Text == "DHBW") c.Delete();
            }

            CalendarEntry calendar = new CalendarEntry();
            calendar.Title.Text = "DHBW";
            calendar.Summary.Text = "Dieser Kalender enthält den Stundenplan laut Dualis";
            calendar.TimeZone = "Europe/Berlin";
            calendar.Hidden = false;
            calendar.Selected = true;
            calendar.Color = "#2952A3";
            calendar.Location = new Where("", "", "Horb am Neckar");

            Uri postUri = new Uri("https://www.google.com/calendar/feeds/default/owncalendars/full");
            CalendarEntry createdCalendar = (CalendarEntry)srv.Insert(postUri, calendar);
            return "http://www.google.com/calendar/feeds/" + createdCalendar.EditUri.Content.Split('/').Last() + "/private/full";
        }
    }
}