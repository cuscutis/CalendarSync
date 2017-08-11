using System;
using System.Collections.Generic;
using Google.Apis.Calendar.v3.Data;

namespace CalendarSync
{
    public static class MergeItems
    {
        /// <summary>
        /// Merge Outlook Calendar Items with Google Calendar Items
        /// </summary>
        /// <param name="gCalendar">Google Calendar Object that can insert, update or delete items</param>
        /// <param name="fromItems">List of Outlook Calendar Items</param>
        /// <param name="toEvents">List of existing Google Calendar Items</param>
        public static void Merge(GoogleCalendarGet gCalendar, List<OutlookItem> fromItems, List<Event> toEvents)
        {
            var previouslyMergedEvents = new Dictionary<string, Event>();
            var itemsToUpdate = new List<OutlookItem>();
            var itemsToAdd = new List<OutlookItem>();
            foreach (Event eventItem in toEvents)
            {
                if (eventItem.ExtendedProperties == null) continue;
                if (eventItem.ExtendedProperties.Private__.ContainsKey(gCalendar.OutlookEntryId))
                {
                    if (!previouslyMergedEvents.ContainsKey(eventItem.ExtendedProperties.Private__[gCalendar.OutlookEntryId])) {
                        previouslyMergedEvents.Add(eventItem.ExtendedProperties.Private__[gCalendar.OutlookEntryId], eventItem);
                    }
                }
            }
            foreach (OutlookItem fromItem in fromItems)
            {
                if (previouslyMergedEvents.ContainsKey(fromItem.EntryID))
                {
                    itemsToUpdate.Add(fromItem);
                }
                else
                {
                    itemsToAdd.Add(fromItem);
                }
            }

            Console.WriteLine("Merging {0} items from Outlook",fromItems.Count);

            Console.WriteLine("  Adding {0} items to Google Calendar", itemsToAdd.Count);
            foreach (OutlookItem item in itemsToAdd)
            {
                Console.WriteLine("    {0}", item.Subject);
                gCalendar.AddEvent(item);
                previouslyMergedEvents.Remove(item.EntryID);
            }
            Console.WriteLine("  Updating {0} items to Google Calendar", itemsToUpdate.Count);
            foreach (OutlookItem item in itemsToUpdate)
            {
                bool doUpdate = false;
                var gItem = previouslyMergedEvents[item.EntryID];
                if (item.Subject != gItem.Summary)
                {
                    gItem.Summary = item.Subject;
                    doUpdate = true;
                }
                if (item.Location != gItem.Location)
                {
                    gItem.Location = item.Location;
                    doUpdate = true;
                }
                if (item.Start != gItem.Start.DateTime)
                {
                    gItem.Start.DateTime = item.Start;
                    doUpdate = true;
                }
                if (item.Start.AddMinutes(item.Duration) != gItem.End.DateTime)
                {
                    gItem.End.DateTime = item.Start.AddMinutes(item.Duration);
                    doUpdate = true;
                }
                if (doUpdate)
                {
                    Console.WriteLine("    {0}", item.Subject);
                    gCalendar.UpdateEvent(gItem);
                }
                previouslyMergedEvents.Remove(item.EntryID);
            }
            Console.WriteLine("  Deleting {0} items in Google Calendar", previouslyMergedEvents.Count);
            foreach (KeyValuePair<string, Event> valuePair in previouslyMergedEvents)
            {
                if (valuePair.Value.Start.DateTime > DateTime.Now)
                {
                    Console.WriteLine("    {0}", valuePair.Value.Summary);
                    gCalendar.DeleteEvent(valuePair.Value);
                }
            }
        }

    }
}
