using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;

namespace CalendarSync
{
    public class OulookGet
    {
        private const int ResultStartOffsetDays = -2;
        private const int ResultEndOffsetDays = 30;
        private const int RecurStartOffsetDays = -15;
        private const int RecurEndOffsetDays = 30;
        private const int RecurIntervalCheckMinutes = 30;

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public async Task<List<OutlookItem>> GetCalendarItems()
        {
            var result = new List<OutlookItem>();

            // Set start value
            DateTime start = DateTime.Today.AddMonths(-6);
            // Set end value
            DateTime end = DateTime.Today.AddDays(30);
            // Initial restriction is Jet query for date range
            string filter = "[Start] >= '" + start.ToString("g") + "' AND [End] <= '" + end.ToString("g") + "'";

            var oApp = new Microsoft.Office.Interop.Outlook.Application();
            var mapiNamespace = oApp.GetNamespace("MAPI"); 
            var calendarFolder = mapiNamespace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar); 
            var outlookCalendarItems = calendarFolder.Items.Restrict(filter);
            outlookCalendarItems.IncludeRecurrences = true;

            foreach (Microsoft.Office.Interop.Outlook.AppointmentItem item in outlookCalendarItems)
            {
                if (item.IsRecurring)
                {
                    try
                    {
                        Microsoft.Office.Interop.Outlook.RecurrencePattern rp = item.GetRecurrencePattern();
                        DateTime first =
                            new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, item.Start.Hour,
                                item.Start.Minute, item.Start.Second).AddDays(RecurStartOffsetDays);
                        DateTime last = DateTime.Now.AddDays(RecurEndOffsetDays);

                        for (DateTime cur = first; cur <= last; cur = cur.AddMinutes(RecurIntervalCheckMinutes))
                        {
                            if (cur.DayOfWeek == DayOfWeek.Saturday || cur.DayOfWeek == DayOfWeek.Sunday)
                            {
                                continue;
                            }

                            if (cur.Hour > 0 && cur.Hour < 8)
                            {
                                continue;
                            }
                            if (cur.Hour > 18 && cur.Hour < 24)
                            {
                                continue;
                            }
                            try
                            {
                                Microsoft.Office.Interop.Outlook.AppointmentItem recur = rp.GetOccurrence(cur);
                                result.Add(new OutlookItem
                                {
                                    EntryID = recur.EntryID + recur.Start.ToString(":yyyy-MM-dd"),
                                    Subject = recur.Subject,
                                    Location = recur.Location,
                                    Start = recur.Start,
                                    Duration = recur.Duration
                                });
                            }
                            catch (Exception)
                            {
                                // this looks bad, but the expected way to find recurrances is to try to create one and if it fails, it throws an exception
                                //Console.Write("");
                            }
                        }
                    }
                    catch
                    {
                        Console.Write("{0} {1} {2} {3}", item.Start.Day, item.Start.Hour, item.Start.Minute, item.Start.Second);
                    }
                }
                else
                {
                    result.Add(new OutlookItem
                    {
                        EntryID = item.EntryID,
                        Subject = item.Subject,
                        Location = item.Location,
                        Start = item.Start,
                        Duration = item.Duration
                    });
                }
            }

            return result.Where(i => i.Start > DateTime.Now.AddDays(ResultStartOffsetDays) && i.Start < DateTime.Now.AddDays(ResultEndOffsetDays)).OrderBy(i => i.Start).ToList();
        }

    }
}
