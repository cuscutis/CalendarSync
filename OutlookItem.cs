using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarSync
{
    [Serializable]
    public class OutlookItem
    {
        public string EntryID { get; set; }
        public string Subject { get; set; }
        public string Location { get; set; }
        public DateTime Start { get; set; }
        public int Duration { get; set; }

    }
}
