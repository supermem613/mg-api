# Useful $select Fields

Common `$select` recipes that keep event responses small but useful.

## CLI usage

Pass these via `--select` on any read verb:

```bash
mg-api calendar list --select subject,start,end,attendees,location
mg-api calendar view --start ... --end ... --select subject,start,end,isOnlineMeeting,onlineMeetingUrl
mg-api calendar get --event-id AAMkAGI... --select id,subject,start,end,organizer,attendees,bodyPreview,responseStatus,isCancelled
```

---

## Useful $select Fields

```
$select=id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeetingUrl,bodyPreview,responseStatus,isCancelled
```
