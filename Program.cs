using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
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
                var outlookItems = outlookCalendar.GetCalendarItems();

                var gCalendar = new GoogleCalendarGet();
                gCalendar.Init();
                var items = gCalendar.GetCalendarItems();

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
