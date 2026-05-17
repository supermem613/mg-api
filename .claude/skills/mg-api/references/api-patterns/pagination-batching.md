# Pagination and Batching

@odata.nextLink paging plus $batch for combining requests.

## Pagination

### @odata.nextLink Pattern

When a response has more results than the page size, Graph returns `@odata.nextLink`:

```json
{
  "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=10&$top=10",
  "value": [ ... ]
}
```

**How to paginate:**

1. Make the initial request
2. Process the `value` array
3. If `@odata.nextLink` exists, GET that URL for the next page
4. Repeat until no `@odata.nextLink` is present

```
// Pseudocode
let url = '/me/messages?$top=25'
while (url) {
  const response = await graph.get(url)
  processResults(response.value)
  url = response['@odata.nextLink'] || null
}
```

> **The nextLink URL is opaque** — don't parse or modify it. Just GET it as-is.

### @odata.deltaLink (Change Tracking)

Some endpoints support delta queries for incremental sync:

```
GET /me/messages/delta
```

Returns `@odata.deltaLink` on the last page. Store it and use it later to get only
changes since the last sync.

---

---

## Batch Requests

Combine multiple Graph calls into a single HTTP request.

```
POST https://graph.microsoft.com/v1.0/$batch
Content-Type: application/json
```

### Request Body

```json
{
  "requests": [
    {
      "id": "1",
      "method": "GET",
      "url": "/me/messages?$top=5&$select=subject,from"
    },
    {
      "id": "2",
      "method": "GET",
      "url": "/me/events?$top=5&$select=subject,start,end"
    }
  ]
}
```

### Response

```json
{
  "responses": [
    {
      "id": "1",
      "status": 200,
      "headers": { "Content-Type": "application/json" },
      "body": {
        "value": [ ... ]
      }
    },
    {
      "id": "2",
      "status": 200,
      "body": {
        "value": [ ... ]
      }
    },
    {
      "id": "3",
      "status": 200,
      "body": {
        "availability": "Available",
        "activity": "Available"
      }
    }
  ]
}
```

### Batch Rules

- Maximum **20 requests** per batch
- Requests can be **independent** or **dependent** (using `dependsOn`)
- Each request in the batch can succeed or fail independently
- Response order may differ from request order — match by `id`

### Dependent Requests

```json
{
  "requests": [
    {
      "id": "1",
      "method": "POST",
      "url": "/me/events",
      "body": { "subject": "New Event", ... },
      "headers": { "Content-Type": "application/json" }
    },
    {
      "id": "2",
      "dependsOn": ["1"],
      "method": "GET",
      "url": "/me/events?$top=1&$orderby=createdDateTime desc"
    }
  ]
}
```

---
