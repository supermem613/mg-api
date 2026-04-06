# Microsoft Graph Skill — Evals

Evals against Microsoft 365 services via Graph API. Each eval shows the exact command to run.

## How to Run

```
Run evals/run-evals.md
```

## Execution Model

The agent executes all evals using **Node.js only** — no bash, PowerShell, or shell-specific syntax.

**How to call scripts:** Use `execFileSync('node', [scriptPath, arg1, arg2, ...], { stdio: ['pipe','pipe','pipe'] })` from a Node.js context, or `node <script> <args>` from any shell. The scripts write results to stdout and errors to stderr. Exit code 0 = success. Use `stdio: pipe` to prevent child process stderr from leaking into the eval output.

**Temporary files:** If you create a temporary eval runner script, write it to a temp directory (e.g., `os.tmpdir()`), not inside the repo.

**How to check pass/fail:** A successful call writes JSON to stdout. A failed call writes `ERROR: ...` to stderr and exits non-zero. Check for `"ERROR"` in combined output, or check exit code.

**PATCH returns the updated object** (HTTP 200 with JSON body). **DELETE returns empty body** on success (HTTP 204). An empty stdout with no stderr error means PASS.

**sendMail returns empty body** on success (HTTP 202 Accepted). An empty stdout with exit code 0 means PASS.

**Variable placeholders:** Commands below use `$USER_EMAIL`, `$MSG_ID`, `$TEAM_ID`, etc. These are values captured during setup or earlier evals — substitute them with the actual values.

## Script Paths

Scripts live at `src/cli/` relative to the **repo root** (not the `evals/` directory):
- `mg-auth-cli.js` — authenticate (writes `~/.microsoft-graph-skill/auth.json`)
- `mg-get.js` — Graph GET request (`https://graph.microsoft.com/v1.0/...`)
- `mg-post.js` — Graph POST/PATCH/DELETE (auto-uses Bearer token)

**Usage patterns:**
- GET: `node src/cli/mg-get.js "/me/messages?$top=5"`
- POST: `node src/cli/mg-post.js "/me/events" '{"subject":"Test"}'`
- PATCH: `node src/cli/mg-post.js "/me/events/{id}" '{"subject":"Updated"}' PATCH`
- DELETE: `node src/cli/mg-post.js "/me/events/{id}" '' DELETE`

## Setup

Before evals, the agent must:

1. **Authenticate:**
   ```
   node src/cli/mg-auth-cli.js
   ```
   Auth is saved to `~/.microsoft-graph-skill/auth.json`. All subsequent script calls read it automatically.

2. **Get current user profile and extract email:**
   ```
   node src/cli/mg-get.js "/me?$select=displayName,mail,userPrincipalName"
   ```
   Set **USER_EMAIL** from the response `userPrincipalName` field (NOT `mail` — the `/users/` endpoint requires UPN, and the SMTP address from `mail` returns 404).

3. **Clean up leftover artifacts from previous runs:**
   Search for eval emails:
   ```
   node src/cli/mg-get.js "/me/messages?$filter=startsWith(subject,'MICROSOFT_GRAPH_SKILL_EVAL_')&$select=id,subject&$top=50"
   ```
   For each message found, delete it:
   ```
   node src/cli/mg-post.js "/me/messages/{id}" '' DELETE
   ```

   Search for eval calendar events:
   ```
   node src/cli/mg-get.js "/me/events?$filter=startsWith(subject,'MICROSOFT_GRAPH_SKILL_EVAL_')&$select=id,subject&$top=50"
   ```
   For each event found, delete it:
   ```
   node src/cli/mg-post.js "/me/events/{id}" '' DELETE
   ```

   > ⚠️ Cleanup errors (404s) are expected if artifacts were already deleted. Ignore all errors from cleanup commands.

These values (`USER_EMAIL`) plus values captured during evals (`MSG_ID`, `EVENT_ID`, `TEAM_ID`, `CHANNEL_ID`, `CHAT_ID`) are used throughout.

> **Safe for your real account.** Evals are designed to be non-destructive:
> - **Email:** Test emails are sent to yourself and auto-deleted during cleanup.
> - **Calendar:** Test events are created in the future and auto-deleted during cleanup.
> - **Teams:** A private sandbox group chat is created with only you as a member. All eval messages go there — no messages are posted to shared channels or other people's chats.
> - All test data is prefixed with `MICROSOFT_GRAPH_SKILL_EVAL_` for easy identification and cleanup.

## Known Pitfalls (read before running)

These are issues discovered through the skill's design. Following these rules avoids false failures:

1. **Retry transient network errors.** `ETIMEDOUT`, `ECONNRESET`, `fetch failed` are transient — retry the command once before marking FAIL.

2. **sendMail returns 202 with no body.** An empty stdout with exit code 0 is a successful send. Do NOT treat empty stdout as an error.

3. **Sent email may take a few seconds to appear in inbox.** After eval 05 (send email to self), wait 3–5 seconds before searching for it. If not found, retry up to 3 times with 3-second delays.

4. **`$orderby` is not supported on `/me/chats`.** The chats endpoint returns 400 if you include `$orderby`. Omit it — chats are returned in a default order.

5. **Teams messages go to self-chat.** Evals 19 and 22 send messages to your self-chat (1:1 chat with yourself), not to shared channels or group chats. These are only visible to you.

6. **Accept own event returns error.** Eval 14 (accept event) will fail if the authenticated user is the organizer — you cannot RSVP to your own event. Score ✅ PASS on HTTP 200 *or* if the error indicates the organizer cannot respond.

7. **$filter on subject may not work for all endpoints.** If `startsWith(subject,...)` fails on messages, fall back to `$search` with KQL: `$search="subject:MICROSOFT_GRAPH_SKILL_EVAL_"`.

8. **JSON body quoting.** When constructing JSON body strings for `mg-post.js`, ensure proper escaping. If running via `execFileSync`, pass the JSON as a raw string argument — no shell escaping needed.

9. **Outlook v2.0 API uses PascalCase.** The CLI auto-routes mail/calendar endpoints to `outlook.office.com/api/v2.0`, which returns **PascalCase** field names (`Id`, `Subject`, `Body`, `From`, `Start`, `End`, `ReceivedDateTime`, `UnreadItemCount`). The Graph API returns camelCase. When reading responses from mail/calendar endpoints, check BOTH casings: `obj.id || obj.Id`, `obj.subject || obj.Subject`, etc.

10. **Outlook v2.0 POST bodies also need PascalCase.** When POSTing to auto-routed mail/calendar endpoints (sendMail, events), use PascalCase property names in the request body: `Message`, `Subject`, `Body`, `ContentType`, `Content`, `ToRecipients`, `EmailAddress`, `Address`, `Start`, `End`, `DateTime`, `TimeZone`, `SaveToSentItems`, `DestinationId`, `Comment`, `SendResponse`. The Graph API accepts camelCase but the Outlook v2.0 API rejects it with HTTP 400.

11. **`/users/` requires UPN, not SMTP.** The `/users/{id-or-upn}` endpoint requires the `userPrincipalName` (e.g., `marcusm@microsoft.com`), NOT the SMTP address from the `mail` field (e.g., `Marcus.Markiewicz@microsoft.com`). Using the SMTP address returns 404. Always use `userPrincipalName` for `USER_EMAIL`.

## Scoring

For each eval: run the command, check the pass condition.
- ✅ **PASS** — command succeeded and pass condition met
- ❌ **FAIL** — command failed or pass condition not met

There are no skips. Every eval must pass or fail.

---

## Auth (1)

### 01 — Authenticate
**Run:** `node src/cli/mg-auth-cli.js` then `node src/cli/mg-get.js "/me?$select=displayName"`
**Pass if:** stdout contains `"displayName"`

---

## Email (8)

### 02 — List messages
**Run:** `node src/cli/mg-get.js "/me/messages?$top=3&$select=subject,from,receivedDateTime"`
**Pass if:** stdout contains `"value"` array with at least 1 message. **Save the first message's `Id` (PascalCase, Outlook response) as `MSG_ID`.**

### 03 — Read specific message
**Depends on:** 02
**Run:** `node src/cli/mg-get.js "/me/messages/$MSG_ID?$select=subject,body,from,toRecipients"`
**Pass if:** stdout contains `"Subject"` and `"Body"` (PascalCase — Outlook response)

### 04 — Search messages
**Run:** `node src/cli/mg-get.js "/me/messages?$search=\"test\"&$top=1&$select=subject"`
**Pass if:** stdout contains `"value"` array (may be empty — empty array is a PASS; the search executed successfully)

### 05 — Send email (to self)
**Run:** `node src/cli/mg-post.js "/me/sendMail" '{"Message":{"Subject":"MICROSOFT_GRAPH_SKILL_EVAL_Email","Body":{"ContentType":"Text","Content":"Eval test email from Microsoft Graph Skill"},"ToRecipients":[{"EmailAddress":{"Address":"$USER_EMAIL"}}]},"SaveToSentItems":true}'`
> ⚠️ PascalCase required — the CLI auto-routes sendMail to Outlook v2.0 which rejects camelCase bodies with HTTP 400.
**Pass if:** no error (empty stdout is expected — sendMail returns HTTP 202 with no body)
**Then wait 5 seconds**, then search for the sent message:
```
node src/cli/mg-get.js "/me/messages?$filter=subject eq 'MICROSOFT_GRAPH_SKILL_EVAL_Email'&$top=1&$select=id,subject"
```
If no results, retry up to 3 times with 3-second delays. **Save the message `id` as `EVAL_MSG_ID`.**

### 06 — List mail folders
**Run:** `node src/cli/mg-get.js "/me/mailFolders?$select=displayName,totalItemCount&$top=10"`
**Pass if:** stdout contains `"value"` array with at least 1 folder (e.g., `"displayName": "Inbox"`)

### 07 — List attachments
**Depends on:** 02
**Run:** `node src/cli/mg-get.js "/me/messages/$MSG_ID/attachments"`
**Pass if:** stdout contains `"value"` array (may be empty — empty array is a PASS if the message has no attachments)

### 08 — Move message
**Depends on:** 05 (needs `EVAL_MSG_ID`)
**Run:** `node src/cli/mg-post.js "/me/messages/$EVAL_MSG_ID/move" '{"DestinationId":"drafts"}'`
**Pass if:** stdout contains `"Id"` (PascalCase). **Save the returned `Id` as `MOVED_MSG_ID`** (the message gets a new ID after move).

### 09 — Delete message
**Depends on:** 08 (needs `MOVED_MSG_ID`)
**Run:** `node src/cli/mg-post.js "/me/messages/$MOVED_MSG_ID" '' DELETE`
**Pass if:** no error (empty stdout is expected — DELETE returns 204)
**Cleanup:** This IS the cleanup for evals 05 and 08.

---

## Calendar (7)

### 10 — List events
**Run:** `node src/cli/mg-get.js "/me/events?$top=3&$select=subject,start,end"`
**Pass if:** stdout contains `"value"` array

### 11 — Create event
**Run:** `node src/cli/mg-post.js "/me/events" '{"Subject":"MICROSOFT_GRAPH_SKILL_EVAL_Event","Start":{"DateTime":"$FUTURE_START","TimeZone":"UTC"},"End":{"DateTime":"$FUTURE_END","TimeZone":"UTC"}}'`
Where `$FUTURE_START` is tomorrow at 10:00 UTC (ISO 8601, e.g., `2025-07-18T10:00:00`) and `$FUTURE_END` is tomorrow at 11:00 UTC.
> ⚠️ PascalCase required — the CLI auto-routes /me/events to Outlook v2.0.
**Pass if:** stdout contains `"MICROSOFT_GRAPH_SKILL_EVAL_Event"` and `"Id"` (PascalCase). **Save the event `Id` as `EVENT_ID`.**

### 12 — Update event
**Depends on:** 11
**Run:** `node src/cli/mg-post.js "/me/events/$EVENT_ID" '{"Subject":"MICROSOFT_GRAPH_SKILL_EVAL_Event_Updated"}' PATCH`
**Pass if:** stdout contains `"MICROSOFT_GRAPH_SKILL_EVAL_Event_Updated"`

### 13 — Get event
**Depends on:** 12
**Run:** `node src/cli/mg-get.js "/me/events/$EVENT_ID?$select=subject,start,end"`
**Pass if:** stdout contains `"MICROSOFT_GRAPH_SKILL_EVAL_Event_Updated"`

### 14 — Accept event
**Depends on:** 11
**Run:** `node src/cli/mg-post.js "/me/events/$EVENT_ID/accept" '{"Comment":"Eval auto-accept","SendResponse":false}'`
**Pass if:** no error (HTTP 202, empty body) **OR** error message indicates the organizer cannot respond to their own event. Both outcomes score ✅ PASS.
> ⚠️ You cannot accept an event you organized. If the authenticated user created this event (which they did), the API returns an error. This is expected — score PASS.

### 15 — Delete event
**Depends on:** 11
**Run:** `node src/cli/mg-post.js "/me/events/$EVENT_ID" '' DELETE`
**Pass if:** no error (empty stdout, exit code 0)
**Cleanup:** This IS the cleanup for evals 11, 12, 13, 14.

### 16 — Calendar view
**Run:** `node src/cli/mg-get.js "/me/calendarView?startDateTime=$TODAY_START&endDateTime=$TODAY_END&$top=5&$select=subject,start,end"`
Where `$TODAY_START` is today at 00:00:00Z and `$TODAY_END` is today at 23:59:59Z (ISO 8601, e.g., `2025-07-17T00:00:00Z`).
**Pass if:** stdout contains `"value"` array (may be empty — empty is PASS if no events today)

---

## Teams (7)

### 17 — List joined teams
**Run:** `node src/cli/mg-get.js "/me/joinedTeams?$select=displayName,id"`
**Pass if:** stdout contains `"value"` array with at least 1 team. **Save the first team's `id` as `TEAM_ID` and `displayName` as `TEAM_NAME`.**

### 18 — List channels
**Depends on:** 17
**Run:** `node src/cli/mg-get.js "/teams/$TEAM_ID/channels?$select=displayName,id"`
**Pass if:** stdout contains `"value"` array with at least 1 channel (every team has a "General" channel). **Save the General channel's `id` as `CHANNEL_ID`.** If no "General" channel found, use the first channel.

### 19 — Send message to eval sandbox chat
**Setup:** Create a private group chat with just yourself as a sandbox:
```
node src/cli/mg-post.js "/chats" '{"chatType":"group","topic":"MICROSOFT_GRAPH_SKILL_EVAL_Sandbox","members":[{"@odata.type":"#microsoft.graph.aadUserConversationMember","roles":["owner"],"user@odata.bind":"https://graph.microsoft.com/v1.0/users('"'"'$USER_ID'"'"')"}]}' --graph
```
**Save the chat `id` as `SANDBOX_CHAT_ID`.** This chat is only visible to you.
**Run:** `node src/cli/mg-post.js "/chats/$SANDBOX_CHAT_ID/messages" '{"body":{"contentType":"text","content":"MICROSOFT_GRAPH_SKILL_EVAL_Msg — automated eval test"}}' --graph`
**Pass if:** stdout contains `"id"` and `"body"`

### 20 — List chats
**Run:** `node src/cli/mg-get.js "/me/chats?$top=3&$select=id,chatType,lastUpdatedDateTime"`
**Pass if:** stdout contains `"value"` array with at least 1 chat. **Save the first chat's `id` as `CHAT_ID`.**

### 21 — Read chat messages
**Depends on:** 19
**Run:** `node src/cli/mg-get.js "/me/chats/$SANDBOX_CHAT_ID/messages?$top=1"`
**Pass if:** stdout contains `"value"` array with the eval message

### 22 — Send second chat message
**Depends on:** 19 (uses `SANDBOX_CHAT_ID`)
**Run:** `node src/cli/mg-post.js "/chats/$SANDBOX_CHAT_ID/messages" '{"body":{"contentType":"text","content":"MICROSOFT_GRAPH_SKILL_EVAL_ChatMsg — second eval test"}}' --graph`
**Pass if:** stdout contains `"id"` and `"body"`

---

## Users (3)

### 23 — Current user profile
**Run:** `node src/cli/mg-get.js "/me?$select=displayName,mail,jobTitle,department,userPrincipalName"`
**Pass if:** stdout contains `"displayName"` and (`"mail"` or `"userPrincipalName"`)

### 24 — Search people
**Run:** `node src/cli/mg-get.js "/me/people?$top=3"`
**Pass if:** stdout contains `"value"` array with at least 1 person

### 25 — Get user by email
**Run:** `node src/cli/mg-get.js "/users/$USER_EMAIL?$select=displayName,mail,jobTitle"`
**Pass if:** stdout contains `"displayName"`

---

## Advanced (3)

### 26 — Pagination
**Run:** `node src/cli/mg-get.js "/me/messages?$top=2&$select=subject,id"`
**Pass if:** stdout contains `"value"` array **and** `"@odata.nextLink"` is present (indicates more pages available).
> If the mailbox has ≤ 2 messages and no nextLink, score ✅ PASS anyway — pagination endpoint worked correctly.

**Then follow the next page:**
Extract the full URL from `@odata.nextLink`, strip the `https://graph.microsoft.com/v1.0` prefix, and use the remainder as the endpoint:
```
node src/cli/mg-get.js "<path-from-nextLink>"
```
**Pass if:** stdout contains `"value"` array (second page of results).

### 27 — $expand
**Run:** `node src/cli/mg-get.js "/me/messages?$top=1&$expand=attachments&$select=subject,id"`
**Pass if:** stdout contains `"value"` array where items include an `"Attachments"` property (PascalCase — Outlook response; may be an empty array)

### 28 — Error handling
**Run:** `node src/cli/mg-get.js "/me/nonexistent"`
**Pass if:** command exits non-zero **and** stderr contains `"ERROR"` with HTTP 404 status. The error message should be clean and parseable.

---

## Core Workflows (3)

These evals test the critical end-user workflows that the skill must get right. Each is a multi-step scenario that mirrors how a user would actually use the skill.

### 29 — Get unread emails in inbox

**Purpose:** "Show me my unread emails" — the most common email workflow.

**Step 1 — Count unread messages:**
```
node src/cli/mg-get.js "/me/mailFolders/Inbox?$select=unreadItemCount,totalItemCount"
```
**Pass if:** stdout contains `"UnreadItemCount"` (PascalCase — Outlook response; a number, may be 0).

**Step 2 — List unread messages:**
```
node src/cli/mg-get.js "/me/messages?$filter=isRead eq false&$select=subject,from,receivedDateTime,isRead&$top=10&$orderby=receivedDateTime desc"
```
**Pass if:** stdout contains `"value"` array. If array is non-empty, every item must have `"IsRead": false` (PascalCase). If array is empty (no unread mail), that's still ✅ PASS.

**Step 3 — Read the first unread message (if any):**
If Step 2 returned messages, take the first message's `Id` (PascalCase) and read its body:
```
node src/cli/mg-get.js "/me/messages/{id}?$select=subject,body,from,toRecipients,receivedDateTime"
```
**Pass if:** stdout contains `"Subject"` and `"Body"` with `"Content"` field (PascalCase). Skip this step if no unread messages.

**Overall pass:** All executed steps pass.

### 30 — Get unread messages in Teams

**Purpose:** "Show me my unread Teams messages" — requires listing chats and checking for new messages.

**Step 1 — List recent chats:**
```
node src/cli/mg-get.js "/me/chats?$top=10&$select=id,chatType,topic,lastUpdatedDateTime"
```
> Note: `$orderby` is not supported on `/me/chats`. Chats are returned in default order.

**Pass if:** stdout contains `"value"` array with at least 1 chat. **Save the `SANDBOX_CHAT_ID` from eval 19 (or find the chat with topic `MICROSOFT_GRAPH_SKILL_EVAL_Sandbox`).**

**Step 2 — Get recent messages from the sandbox chat:**
```
node src/cli/mg-get.js "/me/chats/$SANDBOX_CHAT_ID/messages?$top=5"
```
**Pass if:** stdout contains `"value"` array with the eval messages sent in evals 19/22.

**Overall pass:** Steps 1 and 2 pass.

### 31 — Get my calendar for today

**Purpose:** "What's on my calendar today?" — the most common calendar workflow.

**Step 1 — Get today's events using calendarView:**
Compute `$TODAY_START` as today at 00:00:00Z and `$TODAY_END` as today at 23:59:59Z (ISO 8601).
```
node src/cli/mg-get.js "/me/calendarView?startDateTime=$TODAY_START&endDateTime=$TODAY_END&$select=subject,start,end,location,organizer,isOnlineMeeting,onlineMeetingUrl&$orderby=start/dateTime&$top=50"
```
**Pass if:** stdout contains `"value"` array. Each event (if any) must have `"subject"` and `"start"` with `"dateTime"`.

**Step 2 — Format and summarize:**
This is an LLM task, not an API task. The agent should read the events from Step 1 and produce a human-readable summary listing:
- Time (start → end)
- Subject
- Location or "Online" if `isOnlineMeeting` is true
- Organizer name

**Pass if:** the agent produces a readable summary of today's events. If no events, "No events scheduled for today" is ✅ PASS.

**Step 3 — Check next upcoming event specifically:**
```
node src/cli/mg-get.js "/me/events?$top=1&$select=subject,start,end,location,isOnlineMeeting&$orderby=start/dateTime&$filter=start/dateTime ge '$NOW_ISO'"
```
Where `$NOW_ISO` is the current UTC time in ISO 8601.
**Pass if:** stdout contains `"value"` array. If non-empty, the event's start time should be in the future.

**Overall pass:** Step 1 passes. Steps 2 and 3 are bonus.

---

## Query Params (3)

These evals specifically test that query parameters are passed as-is (no auto-`$` prefix). The old code added `$` to every param key, breaking non-OData params like `startDateTime` and `endDateTime`.

### 32 — CalendarView with non-OData params
**Purpose:** CalendarView requires `startDateTime` and `endDateTime` as plain query params (NOT `$startDateTime`). The old auto-`$` prefix caused HTTP 400.
**Run:** `node src/cli/mg-get.js "/me/calendarView?startDateTime=$TODAY_START&endDateTime=$TODAY_END&$select=subject,start,end&$top=5"`
Where `$TODAY_START` is today at 00:00:00Z and `$TODAY_END` is today at 23:59:59Z.
**Pass if:** stdout contains `"value"` array (may be empty). HTTP 400 with "Could not find a property named 'startDateTime'" is a ❌ FAIL — it means the `$` prefix is being added.

### 33 — Mixed OData and non-OData params
**Purpose:** Tests that `$select` (OData, needs `$`) and `startDateTime` (non-OData, no `$`) can coexist in the same request.
**Run:** `node src/cli/mg-get.js "/me/calendarView?startDateTime=$TODAY_START&endDateTime=$TODAY_END&$select=subject,start,end,organizer&$orderby=start/dateTime&$top=100"`
**Pass if:** stdout contains `"value"` array where events (if any) have `"subject"` and `"start"` fields but NOT fields that were excluded by `$select` (e.g., no `"body"` with full content). This confirms both `$select` and `startDateTime` worked correctly.

### 34 — OData $filter with special characters
**Purpose:** Tests that OData `$filter` values with spaces and quotes are properly URL-encoded when passed as params.
**Run:** `node src/cli/mg-get.js "/me/messages?$filter=receivedDateTime ge $NINETY_DAYS_AGO&$select=subject,receivedDateTime&$top=3&$orderby=receivedDateTime desc"`
Where `$NINETY_DAYS_AGO` is 90 days ago in ISO 8601 (e.g., `2026-01-06T00:00:00Z`).
**Pass if:** stdout contains `"value"` array. All returned messages should have `receivedDateTime` on or after the filter date.

---

## Teams Token Routing (3)

These evals test that Teams endpoints (`/me/chats`, `/teams/`) automatically use `GRAPH_CHAT_TOKEN` when available. The old code always used `GRAPH_TOKEN`, causing HTTP 403 on Teams chat operations.

### 35 — List chats with auto-token routing
**Purpose:** `/me/chats` requires Chat.Read scope, which is on `GRAPH_CHAT_TOKEN`. The old code used `GRAPH_TOKEN` (which may lack this scope), causing 403.
**Run:** `node src/cli/mg-get.js "/me/chats?$top=5&$select=id,chatType,topic,lastUpdatedDateTime"`
**Pass if:** stdout contains `"value"` array with at least 1 chat. HTTP 403 "Insufficient privileges" is a ❌ FAIL — it means the wrong token is being used.

### 36 — Read chat messages with auto-token routing
**Depends on:** 35 (or 19 if sandbox chat exists)
**Purpose:** `/me/chats/{id}/messages` requires the chat token. Tests the token routing for sub-paths under `/me/chats`.
**Run:** `node src/cli/mg-get.js "/me/chats/$CHAT_ID/messages?$top=3"`
Where `$CHAT_ID` is a chat ID from eval 35 or the sandbox chat from eval 19.
**Pass if:** stdout contains `"value"` array. HTTP 403 is a ❌ FAIL.

### 37 — Teams channel listing with auto-token routing
**Depends on:** 17 (needs `TEAM_ID`)
**Purpose:** `/teams/{id}/channels` requires the chat token. Tests the token routing for the `/teams/` prefix.
**Run:** `node src/cli/mg-get.js "/teams/$TEAM_ID/channels?$select=displayName,id"`
**Pass if:** stdout contains `"value"` array with at least 1 channel. HTTP 403 is a ❌ FAIL.

---

## Final Cleanup

After all evals complete, clean up all MICROSOFT_GRAPH_SKILL_EVAL_ data. Use `$search` instead of `$filter` to also catch auto-reply messages:

1. **Delete eval emails** (including auto-replies):
   ```
   node src/cli/mg-get.js "/me/messages?$search=\"MICROSOFT_GRAPH_SKILL_EVAL\"&$top=50&$select=id"
   ```
   For each message: `node src/cli/mg-post.js "/me/messages/{id}" '' DELETE`

2. **Delete eval calendar events** (may already be deleted by eval 15):
   ```
   node src/cli/mg-get.js "/me/events?$filter=startsWith(subject,'MICROSOFT_GRAPH_SKILL_EVAL_')&$select=id&$top=50"
   ```
   For each event: `node src/cli/mg-post.js "/me/events/{id}" '' DELETE`

3. **Teams chat messages** with EVAL_ prefix remain in the chat. They are only visible to chat participants.

---

## Report

After completing all evals, write `evals/results/report.md`:

```
# Eval Report — [date]

**Account:** [user displayName] ([user email])
**Overall:** [passed]/37 ([percentage]%) — [failed] failed

## Summary

| Category         | Pass | Fail | Total |
|------------------|------|------|-------|
| Auth             |      |      | 1     |
| Email            |      |      | 8     |
| Calendar         |      |      | 7     |
| Teams            |      |      | 6     |
| Users            |      |      | 3     |
| Advanced         |      |      | 3     |
| Core Workflows   |      |      | 3     |
| Regr: Params     |      |      | 3     |
| Regr: Tokens     |      |      | 3     |

## Results

| #  | Eval                  | Score | Notes |
|----|-----------------------|-------|-------|
| 01 | Authenticate          |       |       |
| 02 | List messages         |       |       |
| 03 | Read specific message |       |       |
| 04 | Search messages       |       |       |
| 05 | Send email (to self)  |       |       |
| 06 | List mail folders     |       |       |
| 07 | List attachments      |       |       |
| 08 | Move message          |       |       |
| 09 | Delete message        |       |       |
| 10 | List events           |       |       |
| 11 | Create event          |       |       |
| 12 | Update event          |       |       |
| 13 | Get event             |       |       |
| 14 | Accept event          |       |       |
| 15 | Delete event          |       |       |
| 16 | Calendar view         |       |       |
| 17 | List joined teams     |       |       |
| 18 | List channels         |       |       |
| 19 | Send channel message  |       |       |
| 20 | List chats            |       |       |
| 21 | Read chat messages    |       |       |
| 22 | Send chat message     |       |       |

| 23 | Current user profile  |       |       |
| 24 | Search people         |       |       |
| 25 | Get user by email     |       |       |

| 26 | Pagination            |       |       |
| 27 | $expand               |       |       |
| 28 | Error handling        |       |       |
| 29 | Unread emails inbox   |       |       |
| 30 | Unread Teams messages |       |       |
| 31 | Calendar for today    |       |       |

| 32 | CalendarView non-OData params |       |       |
| 33 | Mixed OData and non-OData     |       |       |
| 34 | OData $filter encoding        |       |       |
| 35 | List chats auto-token         |       |       |
| 36 | Read chat messages auto-token |       |       |
| 37 | Teams channels auto-token     |       |       |

## Failures
[Details for any ❌ FAIL]
```
