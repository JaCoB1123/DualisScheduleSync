using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

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

        public void SaveToFile(String icalpath, String xmlpath, String format) {
            if (format == "ics" || format == "both")
            {
                File.WriteAllText(icalpath, ToIcal());
            }
            if (format == "xml" || format == "both")
            {
                File.WriteAllText(xmlpath, ToXml());
            }
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
    }
}
