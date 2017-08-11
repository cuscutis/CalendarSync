using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace CalendarSync
{
    public class GoogleCalendarGet
    {
        private readonly string[] Scopes = { CalendarService.Scope.Calendar };
        private const string ApplicationName = "Calendar Sync";
        private const string DefaultCalendarId = "primary";
        public string OutlookEntryId = "OutlookEntryId";

        private UserCredential _credential;
        private CalendarService _service;

        /// <summary>
        /// Connect to Google Calendar API
        /// </summary>
        /// <returns></returns>
        public bool Init()
        {
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(
                    Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/calendar-dotnet-quickstart");

                _credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            _service = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = _credential,
                ApplicationName = ApplicationName,
            });

            return true;
        }

        /// <summary>
        /// Get list of Calendar Items from yesterday to 30 days from now
        /// </summary>
        /// <returns>A List of Events</returns>
        public async Task<List<Event>> GetCalendarItems()
        {
            var result = new List<Event>();

            // Define parameters of request.
            EventsResource.ListRequest request = _service.Events.List(DefaultCalendarId);
            request.TimeMin = DateTime.Now.AddDays(-2.0);
            request.TimeMax = DateTime.Now.AddDays(30.0);
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 250;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events.
            var events = request.Execute();
            if (events.Items != null && events.Items.Count > 0)
            {
                result.AddRange(events.Items);
            }
            return result;
        }

        /// <summary>
        /// Add new Google Calendar Item from an OutlookItem
        /// </summary>
        /// <param name="item"></param>
        public void AddEvent(OutlookItem item)
        {
            var newEvent = new Event
            {
                Summary = item.Subject,
                Start = new EventDateTime {DateTime = item.Start},
                End = new EventDateTime {DateTime = item.Start.AddMinutes(item.Duration)},
                ExtendedProperties = new Event.ExtendedPropertiesData
                {
                    Private__ = new Dictionary<string, string> {{OutlookEntryId, item.EntryID}}
                }
            };

            var request = _service.Events.Insert(newEvent, DefaultCalendarId);
            request.Execute();

        }

        /// <summary>
        /// Update an event with an updated event
        /// </summary>
        /// <param name="updatedEvent"></param>
        public void UpdateEvent(Event updatedEvent)
        {
            var request = _service.Events.Update(updatedEvent, DefaultCalendarId, updatedEvent.Id);
            request.Execute();
        }

        /// <summary>
        /// Delete the given event
        /// </summary>
        /// <param name="deletedEvent"></param>
        public void DeleteEvent(Event deletedEvent)
        {
            var request = _service.Events.Delete(DefaultCalendarId, deletedEvent.Id);
            request.Execute();            
        }
    }
}