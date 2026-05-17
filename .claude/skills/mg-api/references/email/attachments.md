# Attachments

List, download, and add attachments to mail messages.

## Implemented commands

| Task | Command |
|------|---------|
| List attachments on a message | `mg-api email attachments --message-id AAMkAGI...` |

Routes to `graph` and returns the `attachments` collection (`fileAttachment`, `itemAttachment`, `referenceAttachment`). Downloading the raw `contentBytes` of a large attachment, and adding attachments to drafts via upload sessions, are planned-only (`email large-attach` in `mg-api schema`).

Inspect the live contract first:

```bash
mg-api schema email attachments
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
