using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Requests;

namespace GoogleCalendarToOutlookSyncer
{
    class Program
    {
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "Google-Outlook-syncer";

        /* This property will be added to the appointments in Outlook to keep track of them */
        static string UserPropertyName = "sync-id";

        static CalendarService service;
        static DateTime min = DateTime.Today;
        static DateTime max = new DateTime(2018, 1, 31);

        static void Main(string[] args)
        {
            Console.WriteLine("Started! Press CTRL+C to exit.");

            while (true)
            {
                try
                {
                    sync();
                }
                catch (System.Exception e)
                {
                    Console.WriteLine($"{DateTime.Now} something went wrong, maybe the internet is down?");
                }

                /* Sleep for 10 minutes */
                Thread.Sleep(1000 * 60 * 10);
            }

        }

        static void sync()
        {
            Events googleEvents = getGoogleEvents();
            var outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            var calendarFolder = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            var outlookEvents = GetOutlookAppointmentsInRange(calendarFolder, min, max);
            var googleRequests = new BatchRequest(service);

            List<string> googleEventIDs = new List<string>();
            Dictionary<string, Event> googleOutlookEventIDs = new Dictionary<string, Event>();

            /* Create and update Google events in Outlook */
            foreach (var googleEvent in googleEvents.Items)
            {
                if (googleEvent?.Description?.Contains("no-sync") ?? false)
                    continue;

                if (googleEvent?.PrivateCopy ?? false)
                    continue;

                if (googleEvent?.Description?.Contains("outlook-id:") ?? false)
                {
                    var id = googleEvent.Description.Split(':')[1];
                    if(!googleOutlookEventIDs.Any(i => i.Key == id))
                        googleOutlookEventIDs.Add(id, googleEvent);
                    continue;
                }

                AppointmentItem outlookEvent = findOutlookEventById(outlookEvents, googleEvent.Id);
                googleEventIDs.Add(googleEvent.Id);

                /* Find outlook event or create a new one */
                if (outlookEvent == null)
                {
                    outlookEvent = outlookApp.CreateItem(OlItemType.olAppointmentItem);
                    outlookEvent.UserProperties.Add(UserPropertyName, OlUserPropertyType.olText);
                    var property = outlookEvent.UserProperties.Find(UserPropertyName);
                    property.Value = googleEvent.Id;
                    outlookEvent.ReminderSet = false;
                    Console.WriteLine(DateTime.Now + " Added in Outlook: " + googleEvent.Summary);
                }

                /* Update properties in Outlook */
                outlookEvent.Start = googleEvent.Start?.DateTime ?? DateTime.Parse(googleEvent.Start.Date);
                outlookEvent.End = googleEvent.End?.DateTime ?? DateTime.Parse(googleEvent.End.Date);
                outlookEvent.Subject = googleEvent.Summary;
                outlookEvent.Sensitivity = googleEvent.Visibility == "private" || (googleEvent.Description?.Contains("private") ?? false) ? OlSensitivity.olPrivate : OlSensitivity.olNormal;
                outlookEvent.Location = googleEvent.Location;
                outlookEvent.Body = "Source: Google Calendar";
                outlookEvent.Save();
            }

            /* Detect deleted events from Google, loop backwards because we delete from the list */
            for (int i = countOutlookEvents(outlookEvents); i > 0; i--)
            {
                AppointmentItem outlookEvent = outlookEvents[i];
                var property = outlookEvent.UserProperties.Find(UserPropertyName);

                /* Created in Outlook, sync back to Google Calendar */
                if (property == null)
                {
                    var entryId = outlookEvent.EntryID;
                    if (outlookEvent.RecurrenceState != OlRecurrenceState.olApptNotRecurring)
                        entryId += outlookEvent.Start.ToShortDateString();

                    Event gevent = googleOutlookEventIDs.Where(kvp => kvp.Key == entryId).Select(kvp => kvp.Value).FirstOrDefault();

                    bool is_new = false;

                    if (gevent == null)
                    {
                        gevent = new Event()
                        {
                            Description = "outlook-id:" + entryId
                        };
                        is_new = true;
                        Console.WriteLine(DateTime.Now + " Added in Google Calendar: " + outlookEvent.Subject);
                    };

                    if (outlookEvent.AllDayEvent)
                    {
                        gevent.Start = new EventDateTime() { Date = outlookEvent.Start.ToString("yyyy-MM-dd") };
                        gevent.End = new EventDateTime() { Date = outlookEvent.End.ToString("yyyy-MM-dd") };
                    }
                    else
                    {
                        gevent.Start = new EventDateTime() { DateTime = outlookEvent.Start };
                        gevent.End = new EventDateTime() { DateTime = outlookEvent.End };
                    }

                    gevent.Summary = outlookEvent.Subject;
                    gevent.Location = outlookEvent.Location;
                    gevent.Visibility = outlookEvent.Sensitivity == OlSensitivity.olPrivate ? "private" : "default";

                    if (is_new)
                        googleRequests.Queue<Event>(service.Events.Insert(gevent, "primary"), (content, error, ii, message) => { });
                    else
                        googleRequests.Queue<Event>(service.Events.Update(gevent, "primary", gevent.Id), (content, error, ii, message) => { });

                    /* Remove from the list so we can track deleted events. */
                    googleOutlookEventIDs.Remove(entryId);

                    continue;
                }


                if (!googleEventIDs.Contains(property.Value))
                {
                    Console.WriteLine(DateTime.Now + " Deleted from Outlook: " + outlookEvent.Subject);
                    outlookEvent.Delete();
                }
            }

            googleRequests.ExecuteAsync();

            /* Remove deleted events Outlook events from Google Calendar */
            foreach (var kvp in googleOutlookEventIDs)
            {
                Console.WriteLine(DateTime.Now + " Deleted from Google Calendar: " + kvp.Value.Summary);
                service.Events.Delete("primary", kvp.Value.Id).Execute();
            }

        }

        static int countOutlookEvents(Items items)
        {
            int i = 0;
            foreach(var item in items)
                i++;
            return i;
        }

        static AppointmentItem findOutlookEventById(Items items, string id)
        {
            foreach(AppointmentItem item in items)
            {
                var property = item.UserProperties.Find(UserPropertyName);
                if (property == null || property.Value != id)
                    continue;
                else
                    return item;
            }

            return null;
        }

        /* https://developers.google.com/google-apps/calendar/quickstart/dotnet */
        static Events getGoogleEvents()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/google-outlook-syncer");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //Console.WriteLine("Credential file read from/saved to: " + credPath);
            }

            /* Create Google Calendar API service. */
            service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            /* Define parameters of request. */
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = min;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.TimeMax = max;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
            request.MaxResults = 2500;
            Events events = request.Execute();

            return events;
        }

        /* https://msdn.microsoft.com/en-us/library/office/gg619398.aspx*/
        static Items GetOutlookAppointmentsInRange(MAPIFolder folder, DateTime startTime, DateTime endTime)
        {
            string filter = "[Start] >= '"
                + startTime.ToString("g")
                + "' AND [End] <= '"
                + endTime.ToString("g") + "'";

            Items calItems = folder.Items;
            calItems.IncludeRecurrences = true;
            calItems.Sort("[Start]", Type.Missing);
            Items restrictItems = calItems.Restrict(filter);

            return restrictItems;

        }
    }
}
