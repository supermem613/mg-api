# Create, Update, Delete

POST/PATCH/DELETE on `/me/events`. Includes recurrence pattern construction.

## Implemented commands

| Task | Command |
|------|---------|
| Create a 30-min event | `mg-api calendar create --subject Standup --start 2026-01-15T09:00:00 --end 2026-01-15T09:30:00 --attendees alice@example.com,bob@example.com` |
| Update a field | `mg-api calendar update --event-id AAMkAGI... --body '{"subject":"Standup (renamed)"}'` |
| Delete an event | `mg-api calendar delete --event-id AAMkAGI...` |

All write verbs route to `graph`. `--attendees` is a comma-separated address list that is coerced into the Graph attendee shape (`{ emailAddress: { address }, type: 'required' }`) automatically. For raw control over recipients, types, recurrence, or location use `--body` with a full JSON event payload (see `mg-api schema calendar create` for the schema).

Inspect the live contract first:

```bash
mg-api schema calendar create
mg-api schema calendar update
mg-api schema calendar delete
```

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

---

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

---

---

## Delete an Event

```
DELETE /me/events/{event-id}
```

Returns `204 No Content`. Attendees receive a cancellation.

---

---

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

---
