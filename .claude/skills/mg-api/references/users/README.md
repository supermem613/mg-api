# Microsoft Graph — Users and People Reference

Deep reference for the `mg-api users` capability. The CLI exposes semantic verbs (`mg-api users me|search|get`) that build these endpoints for you. Load this file when you need OData $select field lists, people scoring nuances, or directory lookup semantics that the help text does not spell out.

> **Reminder:** People search results expose `scoredEmailAddresses`, not `emailAddresses`. The `mg-api users search` verb selects the correct field by default.

---

Pick the sibling that matches the task at hand.

- [User Profiles](profiles.md) — `/me`, `/users/{id|upn}`, `/me/manager`, `/me/directReports`, `/me/photo`, free/busy cross-ref.
- [People Search](people-search.md) — `mg-api users search`, `scoredEmailAddresses`, person classes.
- [Directory Search](directory-search.md) — ConsistencyLevel header, filter recipes, search syntax.
