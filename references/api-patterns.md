# Microsoft Graph — Common API Patterns Reference

Lazy-loaded by the agent for cross-cutting patterns that apply to all Graph APIs.

---

## OData Query Parameters

All Graph API endpoints support a standard set of OData query parameters.

### $select — Choose Fields

Reduces payload size by returning only specified fields.

```
GET /me/messages?$select=subject,from,receivedDateTime,isRead
GET /me/events?$select=subject,start,end,location
GET /users?$select=displayName,mail,department
```

> **Always use $select.** Default responses include many fields you don't need.
> Smaller payloads = faster responses and lower token cost.

### $filter — Server-Side Filtering

OData filter expressions evaluated server-side.

```
# Equality
$filter=isRead eq false

# String comparison
$filter=importance eq 'high'

# Nested property
$filter=from/emailAddress/address eq 'alice@contoso.com'

# Date comparison
$filter=receivedDateTime ge 2024-07-01T00:00:00Z

# Boolean
$filter=hasAttachments eq true

# String functions
$filter=startsWith(displayName,'Alice')
$filter=contains(subject,'budget')

# Logical operators
$filter=isRead eq false and importance eq 'high'
$filter=department eq 'Engineering' or department eq 'Product'
$filter=not(isRead eq true)

# In operator
$filter=department in ('Engineering', 'Product', 'Design')
```

### Filter Operators

| Operator | Example | Notes |
|----------|---------|-------|
| `eq` | `status eq 'active'` | Equals |
| `ne` | `status ne 'deleted'` | Not equals |
| `gt` | `age gt 25` | Greater than |
| `ge` | `date ge 2024-01-01` | Greater than or equal |
| `lt` | `age lt 65` | Less than |
| `le` | `date le 2024-12-31` | Less than or equal |
| `and` | `a eq 1 and b eq 2` | Logical AND |
| `or` | `a eq 1 or a eq 2` | Logical OR |
| `not` | `not(isRead eq true)` | Logical NOT |
| `in` | `city in ('SEA','NYC')` | In list |
| `startsWith` | `startsWith(name,'A')` | String prefix |
| `contains` | `contains(subject,'Q3')` | String contains (limited support) |

> **Not all endpoints support all operators.** `contains` is not supported on
> message/event filters. Use `$search` for keyword matching instead.

### $orderby — Sorting

```
$orderby=receivedDateTime desc
$orderby=start/dateTime asc
$orderby=displayName
$orderby=createdDateTime desc,subject asc   # multiple sort keys
```

Supported directions: `asc` (default), `desc`.

### $top — Page Size

```
$top=25       # return 25 results per page
$top=1        # return just the first result
```

Defaults and maximums vary by endpoint:

| Endpoint | Default | Maximum |
|----------|---------|---------|
| `/me/messages` | 10 | 1000 |
| `/me/events` | 10 | — |
| `/me/calendarView` | — | — |
| `/users` | 100 | 999 |
| `/me/people` | 10 | 1000 |

### $skip — Offset Paging

```
$skip=10      # skip first 10 results
```

> **Prefer `@odata.nextLink` over $skip.** Manual skip-based paging can miss items
> or return duplicates if the underlying data changes between pages.

### $count — Total Count

```
$count=true
```

Returns `@odata.count` in the response. Some endpoints require `ConsistencyLevel: eventual` header.

### $search — Keyword Search

```
$search="quarterly report"
$search="from:alice"
$search="subject:budget"
```

> **$search uses KQL** (Keyword Query Language) for messages.
> **$search uses simple text matching** for users (requires `ConsistencyLevel: eventual`).
> **$search and $orderby cannot be combined** — search results are relevance-ranked.

### $expand — Include Related Entities

```
GET /me/messages/{id}?$expand=attachments
GET /me/events/{id}?$expand=attachments
GET /me/chats?$expand=members
```

Reduces round-trips by embedding related entities inline.

---

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

## Date/Time Format

Graph uses **ISO 8601** format throughout.

### Timestamps (UTC)

```
2024-07-15T14:30:00Z
2024-07-15T09:00:00.0000000Z
```

### Calendar DateTimeTimeZone Object

Calendar operations use a composite object:

```json
{
  "dateTime": "2024-07-15T09:00:00",
  "timeZone": "Pacific Standard Time"
}
```

> **The dateTime value does NOT include a Z suffix** when timeZone is specified.
> The timeZone field provides the context.

### Date-Only (for $filter)

```
$filter=receivedDateTime ge 2024-07-01T00:00:00Z
$filter=start/dateTime ge '2024-07-15T00:00:00'
```

### ISO 8601 Duration (for meetingDuration)

```
PT30M      → 30 minutes
PT1H       → 1 hour
PT1H30M    → 1 hour 30 minutes
PT2H       → 2 hours
```

---

## Time Zones

### When to Specify Time Zones

- **Calendar events:** Always specify `timeZone` in `start` and `end` objects
- **calendarView:** `startDateTime` and `endDateTime` parameters use UTC (Z suffix)
- **getSchedule:** Specify in `startTime.timeZone` and `endTime.timeZone`
- **Messages/users:** Timestamps are always UTC

### Common Time Zone Values

| Value | UTC Offset | Region |
|-------|-----------|--------|
| `Pacific Standard Time` | UTC-8 / UTC-7 | US West Coast |
| `Mountain Standard Time` | UTC-7 / UTC-6 | US Mountain |
| `Central Standard Time` | UTC-6 / UTC-5 | US Central |
| `Eastern Standard Time` | UTC-5 / UTC-4 | US East Coast |
| `UTC` | UTC+0 | Universal |
| `GMT Standard Time` | UTC+0 / UTC+1 | UK, Ireland, Portugal |
| `Central European Standard Time` | UTC+1 / UTC+2 | Western Europe |
| `E. Europe Standard Time` | UTC+2 / UTC+3 | Eastern Europe |
| `India Standard Time` | UTC+5:30 | India |
| `China Standard Time` | UTC+8 | China, Singapore |
| `Tokyo Standard Time` | UTC+9 | Japan, Korea |
| `AUS Eastern Standard Time` | UTC+10 / UTC+11 | Australia East |

> These are **Windows time zone names**, not IANA. Graph accepts both but Windows
> names are more reliably supported.

### Prefer Header for Time Zone

You can set a default time zone for the entire request:

```
Prefer: outlook.timezone="Pacific Standard Time"
```

---

## Error Handling

### Standard Error Format

```json
{
  "error": {
    "code": "ErrorItemNotFound",
    "message": "The specified object was not found in the store.",
    "innerError": {
      "date": "2024-07-15T14:30:00",
      "request-id": "guid-here",
      "client-request-id": "guid-here"
    }
  }
}
```

### Common Error Codes

| HTTP Status | Error Code | Meaning | Action |
|-------------|-----------|---------|--------|
| 400 | `BadRequest` | Malformed request | Fix query syntax |
| 401 | `Unauthorized` | Token expired or invalid | Refresh token |
| 403 | `AccessDenied` | Insufficient permissions | Check required scopes |
| 404 | `ErrorItemNotFound` | Resource doesn't exist | Handle gracefully |
| 409 | `Conflict` | Resource conflict | Retry with fresh data |
| 429 | `TooManyRequests` | Throttled | Retry after delay |
| 500 | `InternalServerError` | Service error | Retry with backoff |
| 503 | `ServiceUnavailable` | Service temporarily down | Retry with backoff |
| 504 | `GatewayTimeout` | Upstream timeout | Retry |

### 403 Troubleshooting

The most common 403 causes:

1. **Missing scope:** The token doesn't have the required permission
2. **Admin consent required:** The permission needs tenant admin approval
3. **Resource access policy:** The resource has access restrictions
4. **Conditional access:** Tenant CA policies blocking access

---

## Throttling

### 429 Too Many Requests

```
HTTP/1.1 429 Too Many Requests
Retry-After: 30
```

**Always honor the `Retry-After` header.** It specifies seconds to wait.

### Throttling Limits (Approximate)

| Resource | Limit |
|----------|-------|
| Per app, per user | ~10,000 requests / 10 min |
| Mail send | ~30 messages / minute |
| Teams messages | More aggressive — lower limits |
| Batch requests | 20 requests per batch |

### Retry Strategy

```
1. Get 429 response
2. Read Retry-After header (seconds)
3. Wait that many seconds
4. Retry the request
5. If 429 again, use exponential backoff
```

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

## The Dual-Token Model

The agent uses **two tokens** for different API surfaces:

### Graph Token

- **Scope:** `https://graph.microsoft.com/.default`
- **Used for:** Most operations — mail read, calendar, users, teams channels
- **Endpoints:** `https://graph.microsoft.com/v1.0/...`

### Outlook Token

- **Scope:** `https://outlook.office.com/.default`
- **Used for:** Sending email, chat operations
- **Endpoints:** `https://outlook.office.com/api/v2.0/...`

### Which Token for Which Operation

| Operation | Token | Endpoint Base |
|-----------|-------|---------------|
| List/read messages | Graph | `graph.microsoft.com/v1.0` |
| **Send email** | **Outlook** | `outlook.office.com/api/v2.0` |
| Calendar operations | Graph | `graph.microsoft.com/v1.0` |
| User/people operations | Graph | `graph.microsoft.com/v1.0` |
| Teams channels | Graph | `graph.microsoft.com/v1.0` |
| **Chat operations** | **Graph Chat** | `graph.microsoft.com/v1.0` |

> The agent handles token selection automatically based on the operation.
> This table is for debugging auth failures — if a sendMail gets 401,
> check that the Outlook token (not Graph token) is being used.

---

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

## API Versions

| Version | Base URL | Notes |
|---------|----------|-------|
| `v1.0` | `graph.microsoft.com/v1.0` | Production — use this |
| `beta` | `graph.microsoft.com/beta` | Preview features, may change |

> **Always use v1.0** unless you need a beta-only feature. Beta endpoints
> can change without notice and may be removed.

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
