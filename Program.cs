using System.Globalization;
using System.Net;
using CsvHelper;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using GoogleCalendarToCsv;

// Parse CLI arguments
string calendarId = "primary";
string outputFile = "calendar_events.csv";
int maxResults = 2500;

for (int i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--calendar-id" when i + 1 < args.Length:
            calendarId = args[++i];
            break;
        case "--output" when i + 1 < args.Length:
            outputFile = args[++i];
            break;
        case "--max-results" when i + 1 < args.Length:
            if (!int.TryParse(args[++i], out maxResults))
            {
                Console.Error.WriteLine($"Invalid value for --max-results: '{args[i]}'. Expected an integer.");
                return 1;
            }
            break;
    }
}

// Authenticate with OAuth 2.0
UserCredential credential;
try
{
    await using var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read);
    string credPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        ".credentials",
        "google-calendar-to-csv.json");

    credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
        GoogleClientSecrets.FromStream(stream).Secrets,
        new[] { CalendarService.Scope.CalendarReadonly },
        "user",
        CancellationToken.None,
        new FileDataStore(credPath, true));
}
catch (FileNotFoundException)
{
    Console.Error.WriteLine("credentials.json not found in the working directory. " +
        "Download OAuth 2.0 client credentials from the Google Cloud Console and place the file here.");
    return 1;
}

// Create the Calendar service
var service = new CalendarService(new BaseClientService.Initializer
{
    HttpClientInitializer = credential,
    ApplicationName = "GoogleCalendarToCsv"
});

// Fetch events with pagination
var allEvents = new List<Event>();
string? pageToken = null;

try
{
    do
    {
        var request = service.Events.List(calendarId);
        request.TimeMinDateTimeOffset = DateTimeOffset.UtcNow;
        request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
        request.ShowDeleted = false;
        request.SingleEvents = true;
        request.MaxResults = Math.Min(maxResults - allEvents.Count, 2500);
        request.PageToken = pageToken;

        Events result = await request.ExecuteAsync();
        if (result.Items != null)
            allEvents.AddRange(result.Items);

        pageToken = result.NextPageToken;
    }
    while (pageToken != null && allEvents.Count < maxResults);
}
catch (Google.GoogleApiException ex) when (ex.HttpStatusCode == HttpStatusCode.NotFound)
{
    Console.Error.WriteLine($"Calendar '{calendarId}' was not found. " +
        "Please check the calendar ID and ensure the calendar is shared with your account.");
    return 1;
}

// Map events to records
var records = allEvents.Select(e =>
{
    bool isAllDay = e.Start?.Date != null;
    string FormatDateTime(EventDateTime? dt) =>
        isAllDay
            ? dt?.Date ?? string.Empty
            : dt?.DateTimeDateTimeOffset?.ToString("yyyy-MM-dd HH:mm:ss") ?? string.Empty;

    return new CalendarEventRecord
    {
        EventId = e.Id,
        Summary = e.Summary,
        Description = e.Description,
        Location = e.Location,
        StartDateTime = FormatDateTime(e.Start),
        EndDateTime = FormatDateTime(e.End),
        IsAllDay = isAllDay,
        Status = e.Status,
        Organizer = e.Organizer?.Email,
        Creator = e.Creator?.Email,
        Attendees = e.Attendees != null
            ? string.Join(";", e.Attendees.Select(a => a.Email))
            : null,
        HtmlLink = e.HtmlLink,
        Created = e.CreatedDateTimeOffset?.ToString("yyyy-MM-dd HH:mm:ss"),
        Updated = e.UpdatedDateTimeOffset?.ToString("yyyy-MM-dd HH:mm:ss"),
        Recurrence = e.RecurringEventId
    };
});

// Write to CSV
await using (var writer = new StreamWriter(outputFile))
await using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
{
    await csv.WriteRecordsAsync(records);
}

Console.WriteLine($"Wrote {allEvents.Count} events to '{outputFile}'.");
return 0;

