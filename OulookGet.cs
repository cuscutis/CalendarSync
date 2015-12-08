using System;
using System.Collections.Generic;
using System.Linq;

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
        public List<OutlookItem> GetCalendarItems()
        {
            var result = new List<OutlookItem>();
            Microsoft.Office.Interop.Outlook.Application oApp = null;
            Microsoft.Office.Interop.Outlook.NameSpace mapiNamespace = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder calendarFolder = null;
            Microsoft.Office.Interop.Outlook.Items outlookCalendarItems = null;

            oApp = new Microsoft.Office.Interop.Outlook.Application();
            mapiNamespace = oApp.GetNamespace("MAPI"); 
            calendarFolder = mapiNamespace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar); 
            outlookCalendarItems = calendarFolder.Items;
            outlookCalendarItems.IncludeRecurrences = true;

            foreach (Microsoft.Office.Interop.Outlook.AppointmentItem item in outlookCalendarItems)
            {
                if (item.IsRecurring)
                {
                    Microsoft.Office.Interop.Outlook.RecurrencePattern rp = item.GetRecurrencePattern();
                    DateTime first = new DateTime(DateTime.Now.Year, DateTime.Now.Month, item.Start.Day, item.Start.Hour, item.Start.Minute, item.Start.Second).AddDays(RecurStartOffsetDays);
                    DateTime last = DateTime.Now.AddDays(RecurEndOffsetDays);

                    for (DateTime cur = first; cur <= last; cur = cur.AddMinutes(RecurIntervalCheckMinutes))
                    {
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
                            Console.Write("");
                        }
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
