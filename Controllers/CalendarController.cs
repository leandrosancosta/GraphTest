// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using GraphTutorial.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Web;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TimeZoneConverter;

namespace GraphTutorial.Controllers
{
    public class CalendarController : Controller
    {
        private readonly GraphServiceClient _graphClient;
        private readonly ILogger<HomeController> _logger;

        public CalendarController(
            GraphServiceClient graphClient,
            ILogger<HomeController> logger)
        {
            _graphClient = graphClient;
            _logger = logger;
        }

        [AuthorizeForScopes(Scopes = new[] { "Calendars.Read" })]
        public async Task<IActionResult> Index()
        {
            try
            {
                var userTimeZone = TZConvert.GetTimeZoneInfo(
                    User.GetUserGraphTimeZone());
                var startOfWeekUtc = CalendarController.GetUtcStartOfWeekInTimeZone(
                    DateTime.Today, userTimeZone);

                var events = await GetUserWeekCalendar(startOfWeekUtc);

                var startOfWeekInTz = TimeZoneInfo.ConvertTimeFromUtc(startOfWeekUtc, userTimeZone);
                var model = new CalendarViewModel(startOfWeekInTz, events);

                return View(model);
            }
            catch (ServiceException ex)
            {
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException)
                {
                    throw;
                }

                return View(new CalendarViewModel())
                    .WithError("Error getting calendar view", ex.Message);
            }
        }

        [AuthorizeForScopes(Scopes = new[] { "Calendars.ReadWrite" })]
        public IActionResult New()
        {
            return View();
        }
        // </CalendarNewGetSnippet>

        // <CalendarNewPostSnippet>
        [HttpPost]
        [ValidateAntiForgeryToken]
        [AuthorizeForScopes(Scopes = new[] { "Calendars.ReadWrite" })]
        public async Task<IActionResult> New([Bind("Subject,Attendees,Start,End,Body")] NewEvent newEvent)
        {
            var timeZone = User.GetUserGraphTimeZone();

           
            var graphEvent = new Event
            {
                Subject = newEvent.Subject,
                Start = new DateTimeTimeZone
                {
                    DateTime = newEvent.Start.ToString("o"),
                    
                    TimeZone = timeZone
                },
                End = new DateTimeTimeZone
                {
                    DateTime = newEvent.End.ToString("o"),
                    
                    TimeZone = timeZone
                }
            };

            
            if (!string.IsNullOrEmpty(newEvent.Body))
            {
                graphEvent.Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = newEvent.Body
                };
            }

      
            if (!string.IsNullOrEmpty(newEvent.Attendees))
            {
                var attendees =
                    newEvent.Attendees.Split(';', StringSplitOptions.RemoveEmptyEntries);

                if (attendees.Length > 0)
                {
                    var attendeeList = new List<Attendee>();
                    foreach (var attendee in attendees)
                    {
                        attendeeList.Add(new Attendee{
                            EmailAddress = new EmailAddress
                            {
                                Address = attendee
                            },
                            Type = AttendeeType.Required
                        });
                    }

                    graphEvent.Attendees = attendeeList;
                }
            }

            try
            {
                await _graphClient.Me.Events
                    .Request()
                    .AddAsync(graphEvent);

                return RedirectToAction("Index").WithSuccess("Event created");
            }
            catch (ServiceException ex)
            { 
                return RedirectToAction("Index")
                    .WithError("Error creating event", ex.Error.Message);
            }
        }
        // </CalendarNewPostSnippet>

        // <GetCalendarViewSnippet>
        private async Task<IList<Event>> GetUserWeekCalendar(DateTime startOfWeekUtc)
        {
            var endOfWeekUtc = startOfWeekUtc.AddDays(7);

            var viewOptions = new List<QueryOption>
            {
                new QueryOption("startDateTime", startOfWeekUtc.ToString("o")),
                new QueryOption("endDateTime", endOfWeekUtc.ToString("o"))
            };

            var events = await _graphClient.Me
                .CalendarView
                .Request(viewOptions)
                .Header("Prefer", $"outlook.timezone=\"{User.GetUserGraphTimeZone()}\"")
                .Top(50)
                .Select(e => new
                {
                    e.Subject,
                    e.Organizer,
                    e.Start,
                    e.End
                })
                .OrderBy("start/dateTime")
                .GetAsync();

            IList<Event> allEvents;
            if (events.NextPageRequest != null)
            {
                allEvents = new List<Event>();
                var pageIterator = PageIterator<Event>.CreatePageIterator(
                    _graphClient, events,
                    (e) => {
                        allEvents.Add(e);
                        return true;
                    }
                );
                await pageIterator.IterateAsync();
            }
            else
            {
                allEvents = events.CurrentPage;
            }

            return allEvents;
        }

        private static DateTime GetUtcStartOfWeekInTimeZone(DateTime today, TimeZoneInfo timeZone)
        {
            int diff = System.DayOfWeek.Sunday - today.DayOfWeek;

            var unspecifiedStart = DateTime.SpecifyKind(today.AddDays(diff), DateTimeKind.Unspecified);

            return TimeZoneInfo.ConvertTimeToUtc(unspecifiedStart, timeZone);
        }

    }
}
