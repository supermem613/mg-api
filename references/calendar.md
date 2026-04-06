# Microsoft Graph — Calendar Reference

Lazy-loaded by the agent when handling calendar/event/meeting operations.

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
      "onlineMeetingUrl": "https://teams.microsoft.com/l/meetup-join/...",
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

## Get a Single Event

```
GET /me/events/{event-id}
```

Optional: `$select=subject,body,start,end,attendees,location,onlineMeetingUrl`

---

## Create an Event

```
POST /me/events
Content-Type: application/json
```

### Minimal Example

```json
{
  "subject": "1:1 with Alice",
  "start": {
    "dateTime": "2024-07-20T14:00:00",
    "timeZone": "Pacific Standard Time"
  },
  "end": {
    "dateTime": "2024-07-20T14:30:00",
    "timeZone": "Pacific Standard Time"
  }
}
```

### Full Example (with attendees, Teams meeting, location)

```json
{
  "subject": "Project Kickoff",
  "body": {
    "contentType": "HTML",
    "content": "<p>Let's kick off the new project. Agenda attached.</p>"
  },
  "start": {
    "dateTime": "2024-07-20T10:00:00",
    "timeZone": "Pacific Standard Time"
  },
  "end": {
    "dateTime": "2024-07-20T11:00:00",
    "timeZone": "Pacific Standard Time"
  },
  "location": {
    "displayName": "Conference Room B"
  },
  "attendees": [
    {
      "emailAddress": {
        "address": "alice@contoso.com",
        "name": "Alice"
      },
      "type": "required"
    },
    {
      "emailAddress": {
        "address": "bob@contoso.com",
        "name": "Bob"
      },
      "type": "optional"
    }
  ],
  "isOnlineMeeting": true,
  "onlineMeetingProvider": "teamsForBusiness",
  "allowNewTimeProposals": true
}
```

### Field Reference

| Field | Type | Required | Notes |
|-------|------|----------|-------|
| `subject` | string | Yes | Event title |
| `body` | object | No | `{contentType, content}` |
| `start` | object | Yes | `{dateTime, timeZone}` |
| `end` | object | Yes | `{dateTime, timeZone}` |
| `location` | object | No | `{displayName}` |
| `attendees` | array | No | Each: `{emailAddress, type}` |
| `isOnlineMeeting` | boolean | No | Adds Teams link |
| `onlineMeetingProvider` | string | No | `"teamsForBusiness"` |
| `isAllDay` | boolean | No | All-day event |
| `recurrence` | object | No | Recurrence pattern |
| `allowNewTimeProposals` | boolean | No | Let attendees propose new times |
| `importance` | string | No | `"low"`, `"normal"`, `"high"` |
| `sensitivity` | string | No | `"normal"`, `"personal"`, `"private"`, `"confidential"` |
| `showAs` | string | No | `"free"`, `"tentative"`, `"busy"`, `"oof"`, `"workingElsewhere"` |
| `reminderMinutesBeforeStart` | int | No | Default 15 |

### Attendee Types

| Type | Meaning |
|------|---------|
| `required` | Must attend |
| `optional` | Optional attendee |
| `resource` | Room or equipment |

### Time Zone Values

Common values for the `timeZone` field:

- `Pacific Standard Time`
- `Mountain Standard Time`
- `Central Standard Time`
- `Eastern Standard Time`
- `UTC`
- `GMT Standard Time` (UK)
- `Central European Standard Time`
- `India Standard Time`
- `China Standard Time`
- `Tokyo Standard Time`

> **Always specify timeZone.** If omitted, the API uses UTC and the event may appear
> at the wrong time for the user.

---

## Update an Event

```
PATCH /me/events/{event-id}
Content-Type: application/json
```

```json
{
  "subject": "Updated: Project Kickoff",
  "start": {
    "dateTime": "2024-07-20T11:00:00",
    "timeZone": "Pacific Standard Time"
  },
  "end": {
    "dateTime": "2024-07-20T12:00:00",
    "timeZone": "Pacific Standard Time"
  }
}
```

Only include the fields you want to change. Attendees will receive an update notification.

---

## Delete an Event

```
DELETE /me/events/{event-id}
```

Returns `204 No Content`. Attendees receive a cancellation.

---

## RSVP to an Event

### Accept

```
POST /me/events/{event-id}/accept
Content-Type: application/json
```

```json
{
  "comment": "I'll be there!",
  "sendResponse": true
}
```

### Decline

```
POST /me/events/{event-id}/decline
Content-Type: application/json
```

```json
{
  "comment": "Sorry, I have a conflict.",
  "sendResponse": true
}
```

### Tentatively Accept

```
POST /me/events/{event-id}/tentativelyAccept
Content-Type: application/json
```

```json
{
  "comment": "I'll try to make it.",
  "sendResponse": true
}
```

### RSVP Fields

| Field | Type | Notes |
|-------|------|-------|
| `comment` | string | Optional message to organizer |
| `sendResponse` | boolean | Whether to notify the organizer |

---

## Find Meeting Times

```
POST /me/findMeetingTimes
Content-Type: application/json
```

```json
{
  "attendees": [
    {
      "emailAddress": {
        "address": "alice@contoso.com",
        "name": "Alice"
      },
      "type": "required"
    },
    {
      "emailAddress": {
        "address": "bob@contoso.com",
        "name": "Bob"
      },
      "type": "optional"
    }
  ],
  "timeConstraint": {
    "timeslots": [
      {
        "start": {
          "dateTime": "2024-07-20T09:00:00",
          "timeZone": "Pacific Standard Time"
        },
        "end": {
          "dateTime": "2024-07-20T17:00:00",
          "timeZone": "Pacific Standard Time"
        }
      }
    ]
  },
  "meetingDuration": "PT1H",
  "maxCandidates": 5,
  "isOrganizerOptional": false
}
```

### Response

```json
{
  "meetingTimeSuggestions": [
    {
      "confidence": 100.0,
      "meetingTimeSlot": {
        "start": {
          "dateTime": "2024-07-20T10:00:00.0000000",
          "timeZone": "Pacific Standard Time"
        },
        "end": {
          "dateTime": "2024-07-20T11:00:00.0000000",
          "timeZone": "Pacific Standard Time"
        }
      },
      "attendeeAvailability": [
        {
          "attendee": { "emailAddress": { "address": "alice@contoso.com" } },
          "availability": "free"
        }
      ]
    }
  ]
}
```

### Duration Format

ISO 8601 duration: `PT30M` (30 min), `PT1H` (1 hour), `PT1H30M` (90 min).

---

## Free/Busy (Get Schedule)

```
POST /me/calendar/getSchedule
Content-Type: application/json
```

```json
{
  "schedules": [
    "alice@contoso.com",
    "bob@contoso.com"
  ],
  "startTime": {
    "dateTime": "2024-07-20T09:00:00",
    "timeZone": "Pacific Standard Time"
  },
  "endTime": {
    "dateTime": "2024-07-20T17:00:00",
    "timeZone": "Pacific Standard Time"
  },
  "availabilityViewInterval": 30
}
```

### Response

```json
{
  "value": [
    {
      "scheduleId": "alice@contoso.com",
      "availabilityView": "0000220000000000",
      "scheduleItems": [
        {
          "status": "busy",
          "start": { "dateTime": "2024-07-20T11:00:00", "timeZone": "Pacific Standard Time" },
          "end": { "dateTime": "2024-07-20T12:00:00", "timeZone": "Pacific Standard Time" },
          "subject": "Sprint Review"
        }
      ]
    }
  ]
}
```

### Availability View Codes

| Code | Meaning |
|------|---------|
| `0` | Free |
| `1` | Tentative |
| `2` | Busy |
| `3` | Out of office |
| `4` | Working elsewhere |

---

## List Calendars

```
GET /me/calendars
```

```json
{
  "value": [
    {
      "id": "AAMkAGI2...",
      "name": "Calendar",
      "color": "auto",
      "isDefaultCalendar": true,
      "canEdit": true,
      "owner": {
        "name": "Marcus",
        "address": "marcus@contoso.com"
      }
    }
  ]
}
```

### Events from a Specific Calendar

```
GET /me/calendars/{calendar-id}/events
GET /me/calendars/{calendar-id}/calendarView?startDateTime=...&endDateTime=...
```

---

## Recurrence Pattern

For creating recurring events:

```json
{
  "recurrence": {
    "pattern": {
      "type": "weekly",
      "interval": 1,
      "daysOfWeek": ["monday", "wednesday", "friday"]
    },
    "range": {
      "type": "endDate",
      "startDate": "2024-07-15",
      "endDate": "2024-12-31"
    }
  }
}
```

### Pattern Types

| Type | Fields | Example |
|------|--------|---------|
| `daily` | `interval` | Every 2 days |
| `weekly` | `interval`, `daysOfWeek` | Every Mon/Wed/Fri |
| `absoluteMonthly` | `interval`, `dayOfMonth` | 15th of every month |
| `relativeMonthly` | `interval`, `daysOfWeek`, `index` | 2nd Tuesday monthly |
| `absoluteYearly` | `interval`, `dayOfMonth`, `month` | Jan 15 yearly |

### Range Types

| Type | Fields | Notes |
|------|--------|-------|
| `endDate` | `startDate`, `endDate` | Ends on a specific date |
| `noEnd` | `startDate` | Repeats forever |
| `numbered` | `startDate`, `numberOfOccurrences` | Fixed number of occurrences |

---

## Useful $select Fields

```
$select=id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeetingUrl,bodyPreview,responseStatus,isCancelled
```
