# Errors and Throttling

Graph and Outlook error envelopes, retry-after handling, and common codes.

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
