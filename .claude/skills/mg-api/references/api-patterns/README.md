# Microsoft Graph — Common API Patterns Reference

Cross-cutting reference for all `mg-api` capabilities. Covers OData syntax, calendar-style non-OData parameters, pagination, dates and time zones, error shapes, throttling, the Graph vs Outlook token routing, and the JSON batch protocol. Load this file when a capability help page is not enough.

---

Cross-cutting Graph + Outlook REST conventions split per concern.

- [OData Query Parameters](odata.md) — OData operators, function-style endpoints, escape rules.
- [Pagination and Batching](pagination-batching.md) — skip tokens, deltaLink, batch envelopes.
- [Date, Time, and Time Zones](date-time.md) — `Prefer: outlook.timezone="Pacific Standard Time"`, IANA vs Windows ids.
- [Errors and Throttling](errors.md) — error shape, 429 + Retry-After, retry policy.
- [Token Model](tokens.md) — `token` / `base` per verb, chat fallback, scope expectations.
- [Headers, Versions, and Patterns](conventions.md) — ConsistencyLevel, Prefer, x-anchorMailbox, beta fallback.
