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
        [Route("api/AddToGoogleCalendar")]
        public IHttpActionResult AddToGoogleCalendar() {
            AddToGoogleCalendar("Summary stuff", "This is the description.", new DateTime(2017, 3, 5, 14, 0, 0), new DateTime(2017, 3, 5, 14, 0, 0), 1, "11");
            return Ok();
        }



        private void AddToGoogleCalendar(string summary, string description, DateTime startDateTime, DateTime endDateTime, int reminderMinutes, string eventColorId) {
            String serviceAccountEmail = File.ReadAllText(HostingEnvironment.MapPath(@"~/App_Data/serviceAccountEmail"));

            // Read certificate
            // @ is for verbatim string literal as opposed to regular string literal (prevent taking meaning from characters in key)
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

            // Create the calendar event
            Event calendarEvent = new Event();
            calendarEvent.Summary = summary;
            calendarEvent.Description = description;

            // Start date and time
            EventDateTime start = new EventDateTime();
            start.DateTime = startDateTime;
            start.TimeZone = "America/Chicago";
            calendarEvent.Start = start;

            // End date and time
            EventDateTime end = new EventDateTime();
            end.DateTime = endDateTime;
            end.TimeZone = "America/Chicago";
            calendarEvent.End = end;

            // Reminders and notifications
            var remindersList = new List<EventReminder>() {
                new EventReminder() { Method = "popup", Minutes = reminderMinutes }
            };
            var reminders = new Event.RemindersData { UseDefault = false, Overrides = remindersList };
            calendarEvent.Reminders = reminders;

            // Event color list
            // 1 = light blue
            // 2 = light green
            // 3 = levender
            // 4 = salmon
            // 5 = yellow
            // 6 = tan
            // 7 = aqua
            // 8 = gray
            // 9 = blue
            // 10 = green
            // 11 = red
            
            // Set event color
            calendarEvent.ColorId = eventColorId;

            var calendarId = File.ReadAllText(HostingEnvironment.MapPath(@"~/App_Data/calendarId"));
            var eventInsert = service.Events.Insert(calendarEvent, calendarId).Execute();

            //return Ok();
        }
    }
}