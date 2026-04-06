# Microsoft Graph — Users & People Reference

Lazy-loaded by the agent when handling user profile and people search operations.

---

## Get My Profile

```
GET /me
```

### With $select (recommended)

```
GET /me?$select=displayName,mail,jobTitle,department,officeLocation,mobilePhone,userPrincipalName
```

### Response

```json
{
  "id": "user-guid",
  "displayName": "Marcus Miller",
  "givenName": "Marcus",
  "surname": "Miller",
  "mail": "marcus@contoso.com",
  "userPrincipalName": "marcus@contoso.com",
  "jobTitle": "Software Engineer",
  "department": "Engineering",
  "officeLocation": "Building 25, Room 3042",
  "mobilePhone": "+1-555-0100",
  "businessPhones": ["+1-555-0101"],
  "preferredLanguage": "en-US",
  "companyName": "Contoso"
}
```

### Available Profile Fields

| Field | Type | Notes |
|-------|------|-------|
| `id` | string | Azure AD object ID (GUID) |
| `displayName` | string | Full display name |
| `givenName` | string | First name |
| `surname` | string | Last name |
| `mail` | string | Primary email |
| `userPrincipalName` | string | UPN (often same as email) |
| `jobTitle` | string | Job title |
| `department` | string | Department |
| `officeLocation` | string | Office location |
| `mobilePhone` | string | Mobile phone |
| `businessPhones` | array | Office phone numbers |
| `companyName` | string | Company name |
| `preferredLanguage` | string | Language preference |
| `city` | string | City |
| `state` | string | State/province |
| `country` | string | Country/region |
| `postalCode` | string | Postal/zip code |
| `streetAddress` | string | Street address |
| `employeeId` | string | Employee ID |
| `employeeType` | string | Employee type |
| `accountEnabled` | boolean | Whether account is active |
| `createdDateTime` | string | Account creation date |

---

## Get Another User by ID or UPN

### By Object ID

```
GET /users/{user-guid}
```

### By User Principal Name (email)

```
GET /users/alice@contoso.com
```

### With $select

```
GET /users/alice@contoso.com?$select=displayName,mail,jobTitle,department
```

### Response

Same shape as `/me` response.

> **Permission note:** Getting another user's profile requires `User.Read.All` or
> `User.ReadBasic.All`. `User.ReadBasic.All` only returns basic properties
> (displayName, mail, givenName, surname, userPrincipalName, id).

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

## Get My Manager

```
GET /me/manager
```

### Response

```json
{
  "@odata.type": "#microsoft.graph.user",
  "id": "manager-guid",
  "displayName": "Jane Doe",
  "mail": "jane@contoso.com",
  "jobTitle": "Engineering Manager"
}
```

---

## Get Direct Reports

```
GET /me/directReports
```

Returns an array of user objects for people who report to the authenticated user.

```
GET /users/{user-id}/directReports
```

---

## Get User's Photo

### Get Photo Metadata

```
GET /me/photo
GET /users/{user-id}/photo
```

### Download Photo Binary

```
GET /me/photo/$value
GET /users/{user-id}/photo/$value
```

Returns raw image bytes (JPEG typically). Content-Type header indicates the format.

### Photo Sizes

```
GET /me/photos/{size}/$value
```

Available sizes: `48x48`, `64x64`, `96x96`, `120x120`, `240x240`, `360x360`, `432x432`, `504x504`, `648x648`.

> **404 handling:** Not all users have photos. A 404 response is expected and normal —
> fall back to initials or a default avatar.

---

## Get User's Calendar (Cross-Reference)

When you need another user's availability, don't use their calendar directly.
Use the free/busy API instead:

```
POST /me/calendar/getSchedule
```

See the **calendar.md** reference for details.

---

## Org Chart Traversal

### Get Manager Chain (Recursive)

Not directly supported in one call. Walk up:

```
GET /users/{user-id}/manager
GET /users/{manager-id}/manager
```

### Get Org Chart Down

```
GET /users/{user-id}/directReports
```

Then recurse into each direct report.

---

## Permissions Summary

| Operation | Permission Required | Admin Consent |
|-----------|-------------------|---------------|
| Get own profile | `User.Read` | No |
| Get other user's profile | `User.Read.All` or `User.ReadBasic.All` | Depends on tenant |
| Search people | `People.Read` | No |
| List users | `User.Read.All` | Yes (typically) |
| Get photo | `User.Read` (own) / `User.Read.All` (others) | No / Depends |
| Get manager | `User.Read.All` | Depends |
| Get direct reports | `User.Read.All` | Depends |

> **User.ReadBasic.All vs User.Read.All:** `ReadBasic` returns only basic properties
> and does NOT require admin consent. `Read.All` returns all properties but may
> require admin consent depending on the tenant configuration.
