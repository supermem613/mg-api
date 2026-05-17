# Meeting Time Discovery

Find common availability across attendees with findMeetingTimes and getSchedule.

## Implemented commands

| Task | Command |
|------|---------|
| Find 30-min slots across attendees | `mg-api calendar find-times --attendees alice@example.com,bob@example.com --duration PT30M --start 2026-01-15T09:00:00 --end 2026-01-15T18:00:00` |

`mg-api calendar find-times` routes to `graph` and POSTs to `/me/findMeetingTimes`. `--attendees` is coerced into the attendee shape (`{ emailAddress: { address }, type: 'required' }`) automatically. Use `--body` for full control (resource attendees, location constraints, business-hours filter).

`getSchedule` (free/busy) is **planned-only**. Use it via `--body` against the `findMeetingTimes` verb only as a last resort, otherwise wait for the dedicated verb.

Inspect the live contract first:

```bash
mg-api schema calendar find-times
```

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
