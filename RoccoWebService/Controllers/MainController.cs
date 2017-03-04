using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using System.Web.Hosting;

namespace RoccoWebService.Controllers {
    // See WebApiConfig.cs for routing (other than attribute routing)

    // Restrict MVC action / method to users who are both authenticated and authorized
    // - User will receive 401 HTTP if not authorized
    // - Will need this later
    // - [Authorize]
    public class MainController : ApiController {

        // Add item to Google Calendar using service account
        // Be sure to share the calendar with the service account just like a regular user
        [System.Web.Http.HttpGet]
        public IHttpActionResult AddToGoogleCalendar() {
            String serviceAccountEmail = File.ReadAllText(HostingEnvironment.MapPath(@"~/App_Data/serviceAccountEmail"));

            // Read certificate
            // @ is for verbatim string literal as opposed to regular (prevent taking meaning from characters in key)
            var certificate = new X509Certificate2(HostingEnvironment.MapPath(@"~/App_Data/key.p12"), "notasecret", X509KeyStorageFlags.Exportable);

            // Create credential from certificate
            ServiceAccountCredential credential = new ServiceAccountCredential(
                new ServiceAccountCredential.Initializer(serviceAccountEmail) {
                    Scopes = new[] { CalendarService.Scope.Calendar }
                }.FromCertificate(certificate));

            // Create the service using credential
            var service = new CalendarService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential,
                ApplicationName = "RoccoWebService",
            });

            // Create the calendar event object
            Event calendarEvent = new Event();

            calendarEvent.Summary = "Google I/O 2015";
            calendarEvent.Location = "800 Howard St., San Francisco, CA 94103";
            calendarEvent.Description = "A chance to hear more about Google's developer products.";

            DateTime startDateTime = new DateTime(2017, 3, 4, 14, 0, 0); //("2015-05-28T09:00:00-07:00");
            EventDateTime start = new EventDateTime();
            start.DateTime = startDateTime;
            start.TimeZone = "America/Los_Angeles";
            calendarEvent.Start = start;

            DateTime endDateTime = new DateTime(2017, 3, 4, 14, 0, 0);
            EventDateTime end = new EventDateTime();
            end.DateTime = endDateTime;
            end.TimeZone = "America/Los_Angeles";
            calendarEvent.End = end;

            String[] recurrence = new String[] { "RRULE:FREQ=DAILY;COUNT=2" };
            calendarEvent.Recurrence = recurrence;

            var attendees = new List<EventAttendee>() {
                new EventAttendee() { Email = "nobodyhome1@lennysmadeupserver.com" },
                new EventAttendee() { Email = "nobodyhome2@lennysmadeupserver.com" }
            };
            calendarEvent.Attendees = attendees;

            var remindersList = new List<EventReminder>() {
                new EventReminder() { Method = "email", Minutes = (24 * 60) },
                new EventReminder() { Method = "popup", Minutes = 10 }
            };
            var reminders = new Event.RemindersData { UseDefault = false, Overrides = remindersList };
            calendarEvent.Reminders = reminders;

            var calendarId = File.ReadAllText(HostingEnvironment.MapPath(@"~/App_Data/calendarId"));
            var eventInsert = service.Events.Insert(calendarEvent, calendarId).Execute();

            return Ok();
        }
    }
}