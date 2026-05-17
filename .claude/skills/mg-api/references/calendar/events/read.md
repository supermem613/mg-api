# Read Events

List events with `$filter`, `$orderby`, and `$top`, and fetch a single event by id.

## Implemented commands

| Task | Command |
|------|---------|
| List events ordered by start | `mg-api calendar list --top 20 --orderby "start/dateTime asc" --select subject,start,end,attendees` |
| View a date-range window | `mg-api calendar view --start 2026-01-15T00:00:00Z --end 2026-01-16T00:00:00Z` |
| Get a single event | `mg-api calendar get --event-id AAMkAGI... --select subject,start,end,attendees,location,body` |

All read verbs route to `graph`. `calendar view` wraps `/me/calendarView?startDateTime=...&endDateTime=...` — note those parameters do NOT take a `$` prefix (`api-patterns/odata.md` covers the exception).

Inspect the live contract first:

```bash
mg-api schema calendar list
mg-api schema calendar view
mg-api schema calendar get
```

---

## List Events

### All Events (ordered by creation)

```
GET /me/events?$orderby=start/dateTime&$top=25
```

### Calendar View (time range — recommended)

```
GET /me/calendarView?startDateTime=2024-07-01T00:00:00Z&endDateTime=2024-07-31T23:59:59Z
```

> **calendarView vs events:** Use `calendarView` to get events within a time window
> (includes recurring event occurrences expanded). Use `/me/events` only when you need
> all events regardless of time, or a specific event by ID.

### Query Parameters

| Parameter | Example | Notes |
|-----------|---------|-------|
| `startDateTime` | `2024-07-15T00:00:00Z` | Required for calendarView |
| `endDateTime` | `2024-07-15T23:59:59Z` | Required for calendarView |
| `$top` | `$top=50` | Page size |
| `$select` | `$select=subject,start,end,location,organizer` | Reduce payload |
| `$filter` | `$filter=isOnlineMeeting eq true` | OData filter |
| `$orderby` | `$orderby=start/dateTime` | Sort by start time |

### Common $filter Patterns

```
# Online meetings only
$filter=isOnlineMeeting eq true

# Events I organized
$filter=organizer/emailAddress/address eq 'me@contoso.com'

# Cancelled events excluded (default excludes them)
$filter=isCancelled eq false
```

### Response Shape

```json
{
  "value": [
    {
      "id": "AAMkAGI2...",
      "subject": "Sprint Planning",
      "bodyPreview": "Review backlog and plan sprint 42...",
      "start": {
        "dateTime": "2024-07-15T09:00:00.0000000",
        "timeZone": "Pacific Standard Time"
      },
      "end": {
        "dateTime": "2024-07-15T10:00:00.0000000",
        "timeZone": "Pacific Standard Time"
      },
      "location": {
        "displayName": "Conference Room A",
        "locationType": "conferenceRoom"
      },
      "organizer": {
        "emailAddress": {
          "name": "Jane Doe",
          "address": "jane@contoso.com"
        }
      },
      "attendees": [
        {
          "emailAddress": {
            "name": "Marcus",
            "address": "marcus@contoso.com"
          },
          "type": "required",
          "status": {
            "response": "accepted",
            "time": "2024-07-10T12:00:00Z"
          }
        }
      ],
      "isOnlineMeeting": true,
      "onlineMeetingUrl": "https://teams.cloud.microsoft/l/meetup-join/...",
      "isAllDay": false,
      "isCancelled": false,
      "responseStatus": {
        "response": "accepted",
        "time": "2024-07-10T12:00:00Z"
      }
    }
  ]
}
```

---

---

---

## Get a Single Event

```
GET /me/events/{event-id}
```

Optional: `$select=subject,body,start,end,attendees,location,onlineMeetingUrl`

---

---
