# Send, Reply, Forward

Compose new mail and respond to existing threads via the Outlook REST base.

## Implemented commands

| Task | Command |
|------|---------|
| Send a new message | `mg-api email send --to alice@example.com,bob@example.com --subject "Hello" --body "Hi there"` |
| Send with cc/bcc | `mg-api email send --to alice@example.com --cc team@example.com --bcc audit@example.com --subject "Update" --body "..."` |
| Reply to a message | `mg-api email reply --message-id AAMkAGI... --comment "Thanks, approved"` |

Both verbs are registered with `token: outlook` and `base: outlook`, so the CLI uses the cached Outlook bearer against `https://outlook.office.com/api/v2.0`. `--to`, `--cc`, and `--bcc` accept comma-separated addresses and are coerced into the Outlook recipient shape automatically.

Forward and reply-all are planned-only — see `mg-api schema` for the planned list.

Inspect the live contract first:

```bash
mg-api schema email send
mg-api schema email reply
```

---

## Send Email

**Endpoint:** Outlook REST API (NOT Graph — uses Outlook token)

```
POST https://outlook.office.com/api/v2.0/me/sendmail
Content-Type: application/json
Authorization: Bearer {outlook-token}
```

### Request Body

```json
{
  "Message": {
    "Subject": "Project Update",
    "Body": {
      "ContentType": "HTML",
      "Content": "<p>Hi team,</p><p>Here is the latest update.</p>"
    },
    "ToRecipients": [
      {
        "EmailAddress": {
          "Address": "alice@contoso.com",
          "Name": "Alice"
        }
      }
    ],
    "CcRecipients": [
      {
        "EmailAddress": {
          "Address": "bob@contoso.com"
        }
      }
    ],
    "Importance": "Normal"
  },
  "SaveToSentItems": true
}
```

### Field Reference

| Field | Type | Required | Notes |
|-------|------|----------|-------|
| `Message.Subject` | string | Yes | Email subject line |
| `Message.Body.ContentType` | string | Yes | `"HTML"` or `"Text"` |
| `Message.Body.Content` | string | Yes | The body content |
| `Message.ToRecipients` | array | Yes | At least one recipient |
| `Message.CcRecipients` | array | No | CC recipients |
| `Message.BccRecipients` | array | No | BCC recipients |
| `Message.Importance` | string | No | `"Low"`, `"Normal"`, `"High"` |
| `SaveToSentItems` | boolean | No | Default `true` |

### Success Response

```
HTTP 202 Accepted
```

No response body — 202 means the message was accepted for delivery.

> **Token note:** `mg-api email send` is registered with `token: outlook` so the CLI uses the cached Outlook token automatically. You should never set `Authorization` yourself.

---

---

## Reply to a Message

```
POST /me/messages/{message-id}/reply
Content-Type: application/json
```

```json
{
  "comment": "Thanks, I'll review this today."
}
```

### Reply All

```
POST /me/messages/{message-id}/replyAll
```

Same body format — `comment` is added above the quoted original.

---

---

## Forward a Message

```
POST /me/messages/{message-id}/forward
Content-Type: application/json
```

```json
{
  "comment": "FYI — see the thread below.",
  "toRecipients": [
    {
      "emailAddress": {
        "address": "charlie@contoso.com",
        "name": "Charlie"
      }
    }
  ]
}
```

---
