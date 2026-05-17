# Headers, Versions, and Patterns

Required headers, v1.0 vs beta, and small recipes that show up everywhere.

## Headers

### Standard Request Headers

```
Authorization: Bearer {token}
Content-Type: application/json
```

### Optional Headers

| Header | Value | When to Use |
|--------|-------|-------------|
| `Prefer: outlook.timezone="..."` | Time zone name | Calendar responses in local time |
| `Prefer: outlook.body-content-type="text"` | `text` or `html` | Control body format |
| `ConsistencyLevel: eventual` | `eventual` | Required for $search on /users |
| `Prefer: odata.maxpagesize=50` | Number | Request specific page size |
| `If-Match: {etag}` | ETag value | Optimistic concurrency |

---

---

## API Versions

| Version | Base URL | Notes |
|---------|----------|-------|
| `v1.0` | `graph.microsoft.com/v1.0` | Production — use this |
| `beta` | `graph.microsoft.com/beta` | Preview features, may change |

> **Always use v1.0** unless you need a beta-only feature. Beta endpoints
> can change without notice and may be removed.

---

---

## Useful Patterns

### Get Latest N Items

```
GET /me/messages?$top=5&$orderby=receivedDateTime desc&$select=subject,from,receivedDateTime
```

### Check If Something Exists

```
GET /me/messages?$filter=subject eq 'Exact Subject'&$top=1&$select=id
```

If `value` array is empty, it doesn't exist.

### Get Only Count

```
GET /me/messages/$count
```

Requires `ConsistencyLevel: eventual` header. Returns a plain integer.

### Combine $filter and $select

```
GET /me/messages?$filter=isRead eq false&$select=id,subject,from,receivedDateTime&$top=10&$orderby=receivedDateTime desc
```

### Expand Related Data

```
GET /me/messages/{id}?$expand=attachments($select=id,name,size)
GET /me/events/{id}?$expand=attachments
```
