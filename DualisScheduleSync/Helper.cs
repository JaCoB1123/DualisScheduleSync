using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DualisScheduleSync
{
    static class Helper
    {
        public static Random r = new Random();

        public static String ToISO(this DateTime dt)
        {
            return dt.ToString(@"yyyyMMdd\THHmmssZ");
        }
    }
}
