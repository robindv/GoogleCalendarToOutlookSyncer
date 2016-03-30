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


namespace GoogleCalendarToOutlookSyncer
{
    class Program
    {
        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Google-Outlook-syncer";

        /* This property will be added to the appointments in Outlook to keep track of them */
        static string UserPropertyName = "sync-id";

        static DateTime min = DateTime.Today;
        static DateTime max = new DateTime(2016, 12, 31);

        static void Main(string[] args)
        {
            Console.WriteLine("Started! Press CTRL+C to exit.");

            while(true)
            {
                sync();

                /* Sleep for 30 minutes */
                Thread.Sleep(1000 * 60 * 30);
            }
        }

        static void sync()
        { 
            Events googleEvents = getGoogleEvents();
            var outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            var calendarFolder = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            var outlookEvents = GetOutlookAppointmentsInRange(calendarFolder, min, max);

            List<string> googleEventIDs = new List<string>();

            /* Create and update Google events in Outlook */
            foreach (var googleEvent in googleEvents.Items)
            {
                if (googleEvent?.Description?.Contains("no-sync") ?? false)
                    continue;

                AppointmentItem outlookEvent = findOutlookEventById(outlookEvents, googleEvent.Id);
                googleEventIDs.Add(googleEvent.Id);

                /* Find outlook event or create a new one */
                if(outlookEvent == null)
                {
                    outlookEvent = outlookApp.CreateItem(OlItemType.olAppointmentItem);
                    outlookEvent.UserProperties.Add(UserPropertyName, OlUserPropertyType.olText);
                    var property = outlookEvent.UserProperties.Find(UserPropertyName);
                    property.Value = googleEvent.Id;
                    outlookEvent.ReminderSet = false;
                    Console.WriteLine(DateTime.Now + " Added: " + googleEvent.Summary);
                }

                /* Update properties in Outlook */
                outlookEvent.Start = googleEvent.Start?.DateTime ?? DateTime.Parse(googleEvent.Start.Date);
                outlookEvent.End = googleEvent.End?.DateTime ?? DateTime.Parse(googleEvent.End.Date);
                outlookEvent.Subject = googleEvent.Summary;
                outlookEvent.Sensitivity = googleEvent.Visibility == "private" ? OlSensitivity.olPrivate : OlSensitivity.olNormal;
                outlookEvent.Location = googleEvent.Location;
                outlookEvent.Save();
            }

            /* Detect deleted events from Google, loop backwards because we delete from the list */
            for (int i = countOutlookEvents(outlookEvents); i > 0; i--)
            {
                AppointmentItem outlookEvent = outlookEvents[i];
                var property = outlookEvent.UserProperties.Find(UserPropertyName);

                /* Created in Outlook, do not delete. */
                if (property == null)
                    continue;

                if(! googleEventIDs.Contains(property.Value))
                {
                    Console.WriteLine(DateTime.Now + " Deleted: " + outlookEvent.Subject);
                    outlookEvent.Delete();
                }
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
            var service = new CalendarService(new BaseClientService.Initializer()
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
