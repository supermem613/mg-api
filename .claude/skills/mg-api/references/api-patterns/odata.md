# OData Query Parameters

$select, $filter, $orderby, $expand, $search — what Graph accepts and where it diverges from Outlook REST.

## OData Query Parameters

All Graph API endpoints support a standard set of OData query parameters.

> **⚠️ Key Convention: Include the `$` prefix for OData params, omit it for non-OData params.**
> The mg-api registry stores option templates exactly as they appear in the URL.
> OData params need `$`: `$filter`, `$select`, `$top`, `$orderby`, `$search`.
> Non-OData params (like `startDateTime`, `endDateTime`) must NOT have `$`.
> Helpers like `mg-api calendar view` and `mg-api email list` already use the correct shape for you.

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

---

## Endpoints with Non-OData Parameters

Some Graph endpoints require parameters that are NOT OData query params. These must NOT have a `$` prefix.

### calendarView — Required Non-OData Params

`/me/calendarView` requires `startDateTime` and `endDateTime` as **non-OData** query params:

```
# ✅ CORRECT — no $ prefix on startDateTime/endDateTime
GET /me/calendarView?startDateTime=2024-07-15T00:00:00Z&endDateTime=2024-07-15T23:59:59Z&$select=subject,start,end&$top=50

# ❌ WRONG — $startDateTime causes HTTP 400
GET /me/calendarView?$startDateTime=2024-07-15T00:00:00Z&$endDateTime=2024-07-15T23:59:59Z
```

Using `mg-api calendar view`:
```sh
mg-api calendar view \
  --start 2024-07-15T00:00:00Z \
  --end 2024-07-15T23:59:59Z \
  --select subject,start,end \
  --top 50
```

For raw exploration with `mg-api schema calendar view`, the registry stores `startDateTime` and `endDateTime` as non-OData query template keys, and `$select`, `$top` with the `$` prefix.

> **Rule of thumb:** If the param appears in the [OData v4 spec](https://docs.oasis-open.org/odata/odata/v4.0/) with a `$` prefix, include `$`. Otherwise (like `startDateTime`, `endDateTime`), don't.

---
