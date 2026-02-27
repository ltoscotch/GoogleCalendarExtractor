namespace GoogleCalendarToCsv;

public class CalendarEventRecord
{
    public string? EventId { get; set; }
    public string? Summary { get; set; }
    public string? Description { get; set; }
    public string? Location { get; set; }
    public string? StartDateTime { get; set; }
    public string? EndDateTime { get; set; }
    public bool IsAllDay { get; set; }
    public string? Status { get; set; }
    public string? Organizer { get; set; }
    public string? Creator { get; set; }
    public string? Attendees { get; set; }
    public string? HtmlLink { get; set; }
    public string? Created { get; set; }
    public string? Updated { get; set; }
    public string? Recurrence { get; set; }
}
