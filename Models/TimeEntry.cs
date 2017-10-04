using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace timesheet.Models
{
    public class TimeEntry
    {
        public DateTime EntryDateTime{ get; set; }

        public DateTime LunchExitDateTime { get; set; }

        public DateTime LunchEntryDateTime { get; set; }

        public DateTime ExitDateTime { get; set; }

        public string Comment { get; set; }

        public double Hours { get; set; }
    }
}
