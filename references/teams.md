# Microsoft Graph — Teams Reference

Lazy-loaded by the agent when handling Teams channels, chats, and messaging operations.

---

## List Joined Teams

```
GET /me/joinedTeams
```

### Response

```json
{
  "value": [
    {
      "id": "team-guid-here",
      "displayName": "Engineering",
      "description": "Engineering team workspace",
      "isArchived": false
    }
  ]
}
```

> Returns only teams the authenticated user is a member of.

---

## Get a Specific Team

```
GET /teams/{team-id}
```

### Response

```json
{
  "id": "team-guid-here",
  "displayName": "Engineering",
  "description": "Engineering team workspace",
  "isArchived": false,
  "memberSettings": {
    "allowCreateUpdateChannels": true,
    "allowDeleteChannels": false
  },
  "messagingSettings": {
    "allowUserEditMessages": true,
    "allowUserDeleteMessages": true
  }
}
```

---

## Channels

### List Channels

```
GET /teams/{team-id}/channels
```

### Response

```json
{
  "value": [
    {
      "id": "channel-id-here",
      "displayName": "General",
      "description": "General discussion",
      "membershipType": "standard"
    },
    {
      "id": "channel-id-2",
      "displayName": "Design Reviews",
      "description": null,
      "membershipType": "private"
    }
  ]
}
```

### Channel Membership Types

| Type | Description |
|------|-------------|
| `standard` | Visible to all team members |
| `private` | Visible only to channel members |
| `shared` | Shared across teams/orgs |

### Get a Specific Channel

```
GET /teams/{team-id}/channels/{channel-id}
```

---

## Channel Messages

### Send a Channel Message

```
POST /teams/{team-id}/channels/{channel-id}/messages
Content-Type: application/json
```

```json
{
  "body": {
    "content": "<p>Hey team, the deploy is complete! 🚀</p>",
    "contentType": "html"
  }
}
```

### Body Content Types

| Type | Notes |
|------|-------|
| `html` | Supports basic HTML: `<p>`, `<b>`, `<i>`, `<a>`, `<br>`, `<ul>`, `<li>` |
| `text` | Plain text only |

### Send with Mention

```json
{
  "body": {
    "content": "<at id=\"0\">Alice</at> can you review this?",
    "contentType": "html"
  },
  "mentions": [
    {
      "id": 0,
      "mentionText": "Alice",
      "mentioned": {
        "user": {
          "id": "user-guid",
          "displayName": "Alice",
          "userIdentityType": "aadUser"
        }
      }
    }
  ]
}
```

### List Channel Messages

```
GET /teams/{team-id}/channels/{channel-id}/messages
```

> ⚠️ **Permission note:** Reading channel messages requires `ChannelMessage.Read.All`,
> which needs **admin consent**. The delegated user cannot grant this themselves.
> This is a common blocker — if you get 403, this is likely why.

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
          "displayName": "Jane Doe"
        }
      },
      "body": {
        "contentType": "html",
        "content": "<p>Deploy is done!</p>"
      },
      "attachments": [],
      "reactions": []
    }
  ]
}
```

### List Replies to a Message

```
GET /teams/{team-id}/channels/{channel-id}/messages/{message-id}/replies
```

### Reply to a Channel Message

```
POST /teams/{team-id}/channels/{channel-id}/messages/{message-id}/replies
Content-Type: application/json
```

```json
{
  "body": {
    "content": "Great work! 🎉",
    "contentType": "text"
  }
}
```

---

## Chats (1:1 and Group)

### List Chats

```
GET /me/chats
```

> **Token note:** Chat operations use the **Outlook token** scope, not the standard
> Graph token. The agent's dual-token model handles this, but be aware if debugging
> auth issues.

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

## Permissions Summary

| Operation | Permission Required | Admin Consent |
|-----------|-------------------|---------------|
| List joined teams | `Team.ReadBasic.All` | No |
| List channels | `Channel.ReadBasic.All` | No |
| Send channel message | `ChannelMessage.Send` | No |
| **Read channel messages** | **`ChannelMessage.Read.All`** | **Yes** |
| List chats | `Chat.Read` | No |
| Read chat messages | `Chat.Read` | No |
| Send chat message | `Chat.ReadWrite` | No |

> **Key gotcha:** `ChannelMessage.Read.All` is the most common permission blocker.
> If the tenant admin hasn't granted it, reading channel messages will return 403.
> Chat messages (1:1 and group) do NOT require admin consent.

---

## Message Types

The `messageType` field in message responses:

| Type | Description |
|------|-------------|
| `message` | Regular user message |
| `chatEvent` | System event (member added/removed, topic changed) |
| `typing` | Typing indicator (real-time only) |
| `unknownFutureValue` | Future-proofing |

---

## Hosted Content (Inline Images)

If a message contains inline images, they appear as hosted content:

```
GET /teams/{team-id}/channels/{channel-id}/messages/{message-id}/hostedContents/{content-id}/$value
```

The message body HTML will reference them as:
```html
<img src="../hostedContents/{content-id}/$value">
```

---

## Important Notes

1. **Rate limiting:** Teams APIs are more aggressively throttled than other Graph APIs.
   Expect 429 responses more frequently. Always handle `Retry-After`.

2. **Delta queries:** For channel messages, use delta queries to get only new messages:
   ```
   GET /teams/{team-id}/channels/{channel-id}/messages/delta
   ```

3. **Notifications:** To get real-time message notifications, use Graph subscriptions
   (webhooks), not polling.

4. **Message size limit:** Channel and chat messages have a 28 KB body content limit.
