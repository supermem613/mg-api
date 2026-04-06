# Microsoft Graph — Email Reference

Lazy-loaded by the agent when handling email/mail/message operations.

---

## List Messages

```
GET /me/messages
```

### Query Parameters

| Parameter | Example | Notes |
|-----------|---------|-------|
| `$top` | `$top=25` | Page size (default 10, max 1000) |
| `$skip` | `$skip=10` | Offset-based paging (prefer `@odata.nextLink`) |
| `$select` | `$select=subject,from,receivedDateTime,isRead` | Reduce payload |
| `$filter` | `$filter=isRead eq false` | OData filter |
| `$orderby` | `$orderby=receivedDateTime desc` | Sort order |
| `$search` | `$search="quarterly report"` | KQL keyword search |
| `$count` | `$count=true` | Include total count |

### Common $filter Patterns

```
# Unread messages
$filter=isRead eq false

# High importance
$filter=importance eq 'high'

# From a specific sender
$filter=from/emailAddress/address eq 'boss@contoso.com'

# Received after a date
$filter=receivedDateTime ge 2024-06-01T00:00:00Z

# Has attachments
$filter=hasAttachments eq true

# Combine filters
$filter=isRead eq false and importance eq 'high'

# Subject contains (use $search instead for keyword matching)
$search="subject:budget"
```

### $search KQL Syntax

```
$search="from:john"
$search="subject:meeting"
$search="body:quarterly"
$search="hasAttachment:true"
$search="budget report"          # searches across all fields
```

> **Note:** `$search` and `$filter` can be combined but `$search` and `$orderby`
> cannot be combined — search results are ranked by relevance.

### Response Shape

```json
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('me')/messages",
  "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=10",
  "value": [
    {
      "id": "AAMkAGI2...",
      "subject": "Q3 Budget Review",
      "bodyPreview": "Please find attached...",
      "from": {
        "emailAddress": {
          "name": "Jane Doe",
          "address": "jane@contoso.com"
        }
      },
      "toRecipients": [
        {
          "emailAddress": {
            "name": "Marcus",
            "address": "marcus@contoso.com"
          }
        }
      ],
      "receivedDateTime": "2024-07-15T14:30:00Z",
      "isRead": false,
      "importance": "normal",
      "hasAttachments": true
    }
  ]
}
```

---

## Read a Single Message

```
GET /me/messages/{message-id}
```

Optional: `$select=subject,body,from,toRecipients,ccRecipients,receivedDateTime`

To get the full HTML body:

```
GET /me/messages/{message-id}?$select=body
```

The `body` field contains `{ contentType: "html"|"text", content: "..." }`.

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

> **Token note:** sendMail uses the **Outlook token**, not the Graph token.
> The agent's dual-token model handles this automatically.

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

## Attachments

### List Attachments

```
GET /me/messages/{message-id}/attachments
```

Response:

```json
{
  "value": [
    {
      "@odata.type": "#microsoft.graph.fileAttachment",
      "id": "AAMkAGI2...",
      "name": "report.pdf",
      "contentType": "application/pdf",
      "size": 234567,
      "isInline": false,
      "contentBytes": "base64-encoded-content..."
    }
  ]
}
```

### Get a Specific Attachment

```
GET /me/messages/{message-id}/attachments/{attachment-id}
```

### Download Raw Attachment Content

```
GET /me/messages/{message-id}/attachments/{attachment-id}/$value
```

Returns the raw binary content.

### Attachment Types

| Type | Description |
|------|-------------|
| `#microsoft.graph.fileAttachment` | File attachment — `contentBytes` has base64 data |
| `#microsoft.graph.itemAttachment` | Attached Outlook item (message, event) |
| `#microsoft.graph.referenceAttachment` | Link to a file (OneDrive, SharePoint) |

> **Large attachments:** For attachments > 3MB, use the upload session API.

---

## Move a Message

```
POST /me/messages/{message-id}/move
Content-Type: application/json
```

```json
{
  "destinationId": "AAMkAGI2..."
}
```

Common well-known folder names that can be used as `destinationId`:

- `inbox`
- `drafts`
- `sentitems`
- `deleteditems`
- `archive`
- `junkemail`

Example — move to archive:

```json
{
  "destinationId": "archive"
}
```

---

## Delete a Message

```
DELETE /me/messages/{message-id}
```

Returns `204 No Content`. Message goes to Deleted Items.

---

## Mail Folders

### List Folders

```
GET /me/mailFolders
```

### Response

```json
{
  "value": [
    {
      "id": "AAMkAGI2...",
      "displayName": "Inbox",
      "parentFolderId": "AAMkAGI2...",
      "childFolderCount": 2,
      "unreadItemCount": 5,
      "totalItemCount": 142
    }
  ]
}
```

### List Messages in a Specific Folder

```
GET /me/mailFolders/{folder-id}/messages
GET /me/mailFolders/inbox/messages
GET /me/mailFolders/drafts/messages
```

### Get Child Folders

```
GET /me/mailFolders/{folder-id}/childFolders
```

---

## Mark as Read / Unread

```
PATCH /me/messages/{message-id}
Content-Type: application/json
```

```json
{
  "isRead": true
}
```

---

## Useful $select Fields

Keep payloads small by selecting only needed fields:

```
$select=id,subject,from,toRecipients,receivedDateTime,isRead,importance,hasAttachments,bodyPreview
```

Full list of commonly used fields:

| Field | Type | Notes |
|-------|------|-------|
| `id` | string | Message ID |
| `subject` | string | Subject line |
| `bodyPreview` | string | First 255 chars of body |
| `body` | object | `{contentType, content}` — can be large |
| `from` | object | `{emailAddress: {name, address}}` |
| `toRecipients` | array | Array of `{emailAddress}` |
| `ccRecipients` | array | Array of `{emailAddress}` |
| `receivedDateTime` | string | ISO 8601 timestamp |
| `sentDateTime` | string | ISO 8601 timestamp |
| `isRead` | boolean | Read status |
| `importance` | string | `low`, `normal`, `high` |
| `hasAttachments` | boolean | Whether message has attachments |
| `flag` | object | `{flagStatus: "notFlagged"|"flagged"|"complete"}` |
| `inferenceClassification` | string | `focused` or `other` |
