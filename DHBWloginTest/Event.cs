using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace DHBWloginTest
{
    class Event
    {
        private static int count = 100;
        public DateTime begin
        {
            get;
            private set;
        }
        public DateTime end
        {
            get;
            private set;
        }
        public String title
        {
            get;
            private set;
        }
        public String place
        {
            get;
            private set;
        }
        public String duration
        {
            get;
            private set;
        }
        public int id
        {
            get;
            private set;
        }
        public String uid
        {
            get
            {
                return id.ToString() + "@dhbw.de";
            }
        }

        public Event(DateTime begin, String title)
        {
            this.id = ++count;
            this.begin = begin;
            if (title.Contains("/"))
            {
                String[] s = title.Split('/');
                this.place = s[0].Trim();
                this.title = s[1].Trim();
            }
            else
            {
                this.title = title.Trim();
            }
        }

        public Event(DateTime begin, DateTime end, String title)
            : this(begin, title)
        {
            this.end = end;
        }
        public Event(DateTime begin, DateTime end, String title, String place)
            : this(begin, end, title)
        {
            this.place = place;
        }

        public Event(DateTime begin, String duration, String title)
            : this(begin, title)
        {
            this.duration = duration;
        }
        public Event(DateTime begin, String duration, String title, String place)
            : this(begin, duration, title)
        {
            this.place = place;
        }

        public override string ToString()
        {
            return begin.ToShortDateString() + " " + begin.ToShortTimeString() + " - " + end.ToShortTimeString() + " - " + title;
        }

        public void AppendTo(ref StringBuilder ical)
        {
            ical.AppendLine("BEGIN:VEVENT");
            ical.AppendLine("UID:" + uid);
            ical.AppendLine("SUMMARY:" + title);
            ical.AppendLine("DTSTART:" + begin.ToUniversalTime().ToISO());
            if (end != DateTime.MinValue)
            {
                ical.AppendLine("DTEND:" + end.ToUniversalTime().ToISO());
            }
            else if (!string.IsNullOrEmpty(duration))
            {
                ical.AppendLine("DURATION:" + duration);
            }
            ical.AppendLine("DTSTAMP:" + DateTime.Now.ToISO());
            ical.AppendLine("END:VEVENT");
        }
        
    }
}
