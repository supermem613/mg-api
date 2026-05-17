# Directory Search

Tenant-wide directory lookups via `/users` with `$filter`, `$search`, and `$orderby`.

## CLI status

Tenant-wide directory listing is **planned-only** (`users list` in `mg-api schema`). For now:

- **Look up a known user** by id or UPN: `mg-api users get --user-id alice@example.com`
- **Find someone you collaborate with**: `mg-api users search --query "Martinez"` (relevance-ranked via `/me/people`)

Use the REST shape below to plan the eventual `users list` verb or to enumerate users from a script. Prefer adding the verb over open-coding `$filter`/`$search` against `/users` from agent code.

---

## List Users (Directory Search)

```
GET /users?$top=10
```

### Search Users in the Directory

```
GET /users?$filter=startsWith(displayName,'Alice')&$top=10
GET /users?$filter=department eq 'Engineering'&$top=25
GET /users?$filter=mail eq 'alice@contoso.com'
```

### Common $filter Patterns

```
# By display name prefix
$filter=startsWith(displayName,'Mar')

# By department
$filter=department eq 'Engineering'

# By job title
$filter=jobTitle eq 'Software Engineer'

# By company
$filter=companyName eq 'Contoso'

# Combine filters
$filter=department eq 'Engineering' and jobTitle eq 'Software Engineer'

# Account enabled
$filter=accountEnabled eq true
```

### $search (requires ConsistencyLevel header)

```
GET /users?$search="displayName:alice"&$count=true
```

> **Required header:** `ConsistencyLevel: eventual`
> Also requires `$count=true` in the query string.

```
# Search by display name
$search="displayName:alice johnson"

# Search by mail
$search="mail:alice"

# Search across multiple fields
$search="displayName:alice" OR "department:engineering"
```

### Query Parameters

| Parameter | Example | Notes |
|-----------|---------|-------|
| `$top` | `$top=25` | Page size (max 999) |
| `$select` | `$select=displayName,mail,department` | Reduce payload |
| `$filter` | `$filter=department eq 'Eng'` | OData filter |
| `$search` | `$search="displayName:alice"` | Requires ConsistencyLevel header |
| `$count` | `$count=true` | Required when using $search |
| `$orderby` | `$orderby=displayName` | Sort order |

---
