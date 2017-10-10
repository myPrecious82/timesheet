using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using timesheet.Models;

namespace timesheet
{
    public class TimeRecords
    {
        public List<TimeEntry> TimeEntries { get; set; }

        public TimeRecords()
        {
            var appSettings = ConfigurationManager.AppSettings;

            // is this running from exe outside of bin folder?
            var path = Path.GetFullPath(Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName)) + @"\";

#if DEBUG // if we're debugging, send email to consultant email
            path = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName), @"..\..\"));
            emailSendTo = emailSendFrom;
#endif

            var templatePath = $@"{path}Templates\TimeRecords.txt";

            TimeEntries = new List<TimeEntry>();

            var entries = File.ReadAllLines(templatePath);

            foreach (var entry in entries)
            {
                var newEntry = new TimeEntry();
                var e = entry.Split("\t".ToCharArray());

                if (e[1] != "")
                {
                    newEntry = new TimeEntry()
                    {
                        EntryDateTime = DateTime.ParseExact(e[0] + " " + e[1], "MM/dd/yyyy hh:mm tt", null),
                        LunchExitDateTime = DateTime.ParseExact(e[0] + " " + e[2], "MM/dd/yyyy hh:mm tt", null),
                        LunchEntryDateTime = DateTime.ParseExact(e[0] + " " + e[3], "MM/dd/yyyy hh:mm tt", null),
                        ExitDateTime = DateTime.ParseExact(e[0] + " " + e[4], "MM/dd/yyyy hh:mm tt", null)
                    };

                    if (e.Length == 6)
                    {
                        newEntry.Comment = e[5];
                    }
                }
                else
                {
                    newEntry = new TimeEntry()
                    {
                        EntryDateTime = DateTime.ParseExact(e[0] + " 07:00 am", "MM/dd/yyyy hh:mm tt", null),
                        LunchExitDateTime = DateTime.ParseExact(e[0] + " 07:00 am", "MM/dd/yyyy hh:mm tt", null),
                        LunchEntryDateTime = DateTime.ParseExact(e[0] + " 07:00 am", "MM/dd/yyyy hh:mm tt", null),
                        ExitDateTime = DateTime.ParseExact(e[0] + " 07:00 am", "MM/dd/yyyy hh:mm tt", null)
                    };
                }

                var lunch = (newEntry.LunchEntryDateTime - newEntry.LunchExitDateTime).TotalHours;
                var day = (newEntry.ExitDateTime - newEntry.EntryDateTime).TotalHours;

                newEntry.Hours = Math.Round((day - lunch) * 4, MidpointRounding.ToEven) / 4;

                TimeEntries.Add(newEntry);
            }

            // add another two weeks of dummy date rows
            using (var sw = File.AppendText(templatePath))
            {
                var lastDate = TimeEntries[TimeEntries.Count - 1].EntryDateTime.Date;
                for (var i = 1; i < 15; i++)
                {
                    lastDate = lastDate.AddDays(1);
                    var str = (lastDate.DayOfWeek == DayOfWeek.Saturday || lastDate.DayOfWeek == DayOfWeek.Sunday) ? $"{lastDate:MM/dd/yyyy}\t\t\t\t\t" : $"{lastDate:MM/dd/yyyy}\t08:00 am\t11:30 am\t12:30 pm\t05:00 pm\t";
                    sw.WriteLine(str);
                }
            }
        }
    }
}
