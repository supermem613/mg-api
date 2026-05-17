# Calendars

List and target calendars — primary, secondary, and shared.

## CLI status

Calendar enumeration is **not yet a dedicated verb**. The implemented `mg-api calendar list|view|get` verbs operate on `/me/events` (the primary calendar). To target a secondary calendar, fetch its id with the REST shape below and call `mg-api calendar list --body` with `endpoint` override (advanced) — or wait for the planned `calendar instances` verb.

For the typical case (primary calendar only), no extra step is required:

```bash
mg-api calendar list --top 20
mg-api calendar view --start ... --end ...
```

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
