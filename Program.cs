using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Google.Apis.Calendar.v3.Data;

namespace CalendarSync
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var outlookCalendar = new OulookGet();

                var gCalendar = new GoogleCalendarGet();
                gCalendar.Init();

                var tasks = new List<Task>
                {
                    outlookCalendar.GetCalendarItems(),
                    gCalendar.GetCalendarItems()
                };
                Task.WhenAll(tasks);

                var oTask = tasks[0] as Task<List<OutlookItem>>;
                var gTask = tasks[1] as Task<List<Event>>;
                if (oTask == null || gTask == null) return;

                var outlookItems = oTask.Result;
                var items = gTask.Result;

                MergeItems.Merge(gCalendar, outlookItems, items);
            }
            catch (Exception ex)
            {
                Console.WriteLine("========================================");
                Console.WriteLine(ex.GetType().ToString());
                Console.WriteLine(ex.Message);
                Console.WriteLine("========================================");
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine("========================================");
                Console.ReadLine();
            }
        }

    }
}
