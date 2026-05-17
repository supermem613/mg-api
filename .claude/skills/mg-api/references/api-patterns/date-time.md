# Date, Time, and Time Zones

ISO 8601 timestamps and the Prefer header for time zones on calendar reads.

## Date/Time Format

Graph uses **ISO 8601** format throughout.

### Timestamps (UTC)

```
2024-07-15T14:30:00Z
2024-07-15T09:00:00.0000000Z
```

### Calendar DateTimeTimeZone Object

Calendar operations use a composite object:

```json
{
  "dateTime": "2024-07-15T09:00:00",
  "timeZone": "Pacific Standard Time"
}
```

> **The dateTime value does NOT include a Z suffix** when timeZone is specified.
> The timeZone field provides the context.

### Date-Only (for $filter)

```
$filter=receivedDateTime ge 2024-07-01T00:00:00Z
$filter=start/dateTime ge '2024-07-15T00:00:00'
```

### ISO 8601 Duration (for meetingDuration)

```
PT30M      → 30 minutes
PT1H       → 1 hour
PT1H30M    → 1 hour 30 minutes
PT2H       → 2 hours
```

---

---

## Time Zones

### When to Specify Time Zones

- **Calendar events:** Always specify `timeZone` in `start` and `end` objects
- **calendarView:** `startDateTime` and `endDateTime` parameters use UTC (Z suffix)
- **getSchedule:** Specify in `startTime.timeZone` and `endTime.timeZone`
- **Messages/users:** Timestamps are always UTC

### Common Time Zone Values

| Value | UTC Offset | Region |
|-------|-----------|--------|
| `Pacific Standard Time` | UTC-8 / UTC-7 | US West Coast |
| `Mountain Standard Time` | UTC-7 / UTC-6 | US Mountain |
| `Central Standard Time` | UTC-6 / UTC-5 | US Central |
| `Eastern Standard Time` | UTC-5 / UTC-4 | US East Coast |
| `UTC` | UTC+0 | Universal |
| `GMT Standard Time` | UTC+0 / UTC+1 | UK, Ireland, Portugal |
| `Central European Standard Time` | UTC+1 / UTC+2 | Western Europe |
| `E. Europe Standard Time` | UTC+2 / UTC+3 | Eastern Europe |
| `India Standard Time` | UTC+5:30 | India |
| `China Standard Time` | UTC+8 | China, Singapore |
| `Tokyo Standard Time` | UTC+9 | Japan, Korea |
| `AUS Eastern Standard Time` | UTC+10 / UTC+11 | Australia East |

> These are **Windows time zone names**, not IANA. Graph accepts both but Windows
> names are more reliably supported.

### Prefer Header for Time Zone

You can set a default time zone for the entire request:

```
Prefer: outlook.timezone="Pacific Standard Time"
```

---
