# People Search

Relevance-ranked people search via `/me/people` — backed by the user's communication graph.

## Implemented commands

| Task | Command |
|------|---------|
| Relevance-ranked people search | `mg-api users search --query "Martinez" --top 10` |

Routes to `graph` and wraps `GET /me/people?$search="..."`. Results include `scoredEmailAddresses` (not `emailAddresses`) — the verb already requests the right field set. Use `--select` to override.

Inspect the live contract first:

```bash
mg-api schema users search
```

---

## Search People (Relevance-Ranked)

```
GET /me/people
```

Returns people ranked by relevance to the authenticated user (based on communication
patterns, org chart proximity, and collaboration signals).

### Search by Name

```
GET /me/people?$search="alice"
```

### Filter and Limit

```
GET /me/people?$top=10&$search="john"
GET /me/people?$filter=personType/class eq 'Person' and personType/subclass eq 'OrganizationUser'
```

### Response

```json
{
  "value": [
    {
      "id": "person-id",
      "displayName": "Alice Johnson",
      "givenName": "Alice",
      "surname": "Johnson",
      "scoredEmailAddresses": [
        {
          "address": "alice@contoso.com",
          "relevanceScore": 8.0
        }
      ],
      "jobTitle": "Program Manager",
      "department": "Product",
      "officeLocation": "Building 25",
      "personType": {
        "class": "Person",
        "subclass": "OrganizationUser"
      },
      "phones": [
        {
          "type": "business",
          "number": "+1-555-0102"
        }
      ]
    }
  ]
}
```

### Person Types

| Class | Subclass | Description |
|-------|----------|-------------|
| `Person` | `OrganizationUser` | Internal user in the organization |
| `Person` | `PersonalContact` | Personal contact from Outlook |
| `Person` | `ExternalUser` | External collaborator |
| `Person` | `ImplicitContact` | Inferred from communication patterns |
| `Group` | `UnifiedGroup` | Microsoft 365 group |

> **People API vs Users API:** Use `/me/people` when you want relevance-ranked results
> (e.g., "who does the user frequently work with named John?"). Use `/users` when you
> need to look up a specific user by exact ID or email.

---
