# User Profiles

Read the signed-in user, look up other users by id or UPN, walk the org chart, and fetch profile photos.

## Implemented commands

| Task | Command |
|------|---------|
| Get my profile | `mg-api users me --select displayName,mail,jobTitle,department,officeLocation` |
| Get another user by id or UPN | `mg-api users get --user-id alice@example.com --select displayName,mail,jobTitle` |

Both verbs route to `graph`. Manager / direct reports / photo download / org-chart walking are **planned-only** (`users list`, `users photo` in `mg-api schema`). Use the REST shape below to construct a manual call only when blocked, and prefer adding a verb instead.

Inspect the live contract first:

```bash
mg-api schema users me
mg-api schema users get
```

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

---

## Get User's Calendar (Cross-Reference)

When you need another user's availability, don't use their calendar directly.
Use the free/busy API instead:

```
POST /me/calendar/getSchedule
```

See the **calendar.md** reference for details.

---

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
