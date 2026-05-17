# RSVP

Accept, tentatively accept, or decline an existing event invitation.

## Implemented commands

| Task | Command |
|------|---------|
| Accept an invitation | `mg-api calendar accept --event-id AAMkAGI... --comment "See you there" --send-response true` |
| Decline an invitation | `mg-api calendar decline --event-id AAMkAGI... --comment "Conflict, can't make it" --send-response true` |

Both verbs route to `graph` and POST to `/me/events/{id}/accept` or `/decline`. `--send-response false` performs the RSVP silently. Tentative accept is **planned-only**.

Inspect the live contract first:

```bash
mg-api schema calendar accept
mg-api schema calendar decline
```

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

---
