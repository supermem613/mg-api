# List and Read Messages

Read mail with `$filter`, `$search`, and `$select`; mark as read/unread; move and delete.

## Implemented commands

| Task | Command |
|------|---------|
| List recent messages | `mg-api email list --top 10 --select subject,from,receivedDateTime --orderby "receivedDateTime desc"` |
| Read a single message | `mg-api email get --message-id AAMkAGI... --select subject,body,from,toRecipients` |
| Keyword search | `mg-api email search --query "quarterly review" --top 25` |
| Move to a folder | `mg-api email move --message-id AAMkAGI... --destination-id inbox` |
| Soft-delete | `mg-api email delete --message-id AAMkAGI...` |
| List attachments | `mg-api email attachments --message-id AAMkAGI...` |

All message read/move/delete verbs route to `outlook` (`https://outlook.office.com/api/v2.0`). Inspect the live contract first:

```bash
mg-api schema email list
mg-api schema email search
mg-api schema email move
```

The reference below covers the underlying REST shape for cases where the schema does not spell it out (custom `$filter`/`$search`, response payloads, mark-as-read patch body).

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

---

## Delete a Message

```
DELETE /me/messages/{message-id}
```

Returns `204 No Content`. Message goes to Deleted Items.

---

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
