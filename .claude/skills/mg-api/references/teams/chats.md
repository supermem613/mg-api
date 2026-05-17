# Chats

One-to-one and group chats, chat message threads, and chat membership.

## Implemented commands

| Task | Command |
|------|---------|
| List my chats | `mg-api chats list --top 50` |
| Read recent messages in a chat | `mg-api chats messages --chat-id 19:abc...@thread.v2 --top 50` |
| Send a chat message | `mg-api chats send --chat-id 19:abc...@thread.v2 --content "Heads up"` |

`chats list` and `chats messages` route to `token: outlook` (Outlook REST is the supported audience for chat reads in many tenants). `chats send` routes to `token: chat` (Graph with chat scopes). Chat get-by-id and member listing are **planned-only**.

Inspect the live contract first:

```bash
mg-api schema chats list
mg-api schema chats messages
mg-api schema chats send
```

---

## Chats (1:1 and Group)

### List Chats

```
GET /me/chats
```

> **Token note:** `mg-api chats list|messages` is registered with `token: outlook`. `mg-api chats send` and `mg-api teams send-channel-message` use `token: chat`. The CLI selects the right cached token from `~/.mg-api/auth.json` — you should never set `Authorization` yourself.

### Query Parameters

| Parameter | Example | Notes |
|-----------|---------|-------|
| `$top` | `$top=50` | Page size |
| `$expand` | `$expand=members` | Include chat members |
| `$filter` | `$filter=chatType eq 'oneOnOne'` | Filter by type |
| `$orderby` | `$orderby=lastMessagePreview/createdDateTime desc` | Sort by recent |

### Response

```json
{
  "value": [
    {
      "id": "19:abc123...@thread.v2",
      "chatType": "oneOnOne",
      "topic": null,
      "createdDateTime": "2024-06-01T10:00:00Z",
      "lastUpdatedDateTime": "2024-07-15T14:30:00Z",
      "lastMessagePreview": {
        "body": {
          "content": "Sounds good, let's sync tomorrow."
        },
        "createdDateTime": "2024-07-15T14:30:00Z",
        "from": {
          "user": {
            "displayName": "Alice"
          }
        }
      }
    }
  ]
}
```

### Chat Types

| Type | Description |
|------|-------------|
| `oneOnOne` | 1:1 chat |
| `group` | Group chat |
| `meeting` | Meeting chat |

### Get a Specific Chat

```
GET /me/chats/{chat-id}
```

---

---

## Chat Messages

### List Messages in a Chat

```
GET /me/chats/{chat-id}/messages
```

### Query Parameters

| Parameter | Example | Notes |
|-----------|---------|-------|
| `$top` | `$top=50` | Page size (default 20) |
| `$orderby` | `$orderby=createdDateTime desc` | Sort order |

### Response

```json
{
  "value": [
    {
      "id": "message-id",
      "messageType": "message",
      "createdDateTime": "2024-07-15T14:30:00Z",
      "from": {
        "user": {
          "id": "user-guid",
          "displayName": "Alice"
        }
      },
      "body": {
        "contentType": "text",
        "content": "Hey, are you free at 3?"
      }
    }
  ]
}
```

### Send a Chat Message

```
POST /me/chats/{chat-id}/messages
Content-Type: application/json
```

```json
{
  "body": {
    "content": "Sure, let's do 3pm!",
    "contentType": "text"
  }
}
```

---

---

## Chat Members

### List Chat Members

```
GET /me/chats/{chat-id}/members
```

### Add a Member to a Group Chat

```
POST /me/chats/{chat-id}/members
Content-Type: application/json
```

```json
{
  "@odata.type": "#microsoft.graph.aadUserConversationMember",
  "roles": ["member"],
  "user@odata.bind": "https://graph.microsoft.com/v1.0/users('user-guid')"
}
```

---
