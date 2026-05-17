# Mail Folders

Browse and target mail folders by well-known id or by id from `/me/mailFolders`.

## CLI status

Folder browsing is **planned-only** (`mg-api schema` lists it under `email.folders`). Today, agents that need to target a folder use the well-known ids (`inbox`, `archive`, `sentitems`, `drafts`, `deleteditems`) as `--destination-id` on `mg-api email move`:

```bash
mg-api email move --message-id AAMkAGI... --destination-id archive
mg-api email move --message-id AAMkAGI... --destination-id sentitems
```

Use the REST shape below only when you need to enumerate custom folders or look up a non-well-known id, then pass that id to `mg-api email move`.

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
