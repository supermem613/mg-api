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
   Set **USER_EMAIL** from the response `mail` or `userPrincipalName` field (whichever is populated).

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
**Pass if:** stdout contains `"value"` array with at least 1 message. **Save the first message's `id` as `MSG_ID`.**

### 03 — Read specific message
**Depends on:** 02
**Run:** `node src/cli/mg-get.js "/me/messages/$MSG_ID?$select=subject,body,from,toRecipients"`
**Pass if:** stdout contains `"subject"` and `"body"`

### 04 — Search messages
**Run:** `node src/cli/mg-get.js "/me/messages?$search=\"test\"&$top=1&$select=subject"`
**Pass if:** stdout contains `"value"` array (may be empty — empty array is a PASS; the search executed successfully)

### 05 — Send email (to self)
**Run:** `node src/cli/mg-post.js "/me/sendMail" '{"message":{"subject":"MICROSOFT_GRAPH_SKILL_EVAL_Email","body":{"contentType":"Text","content":"Eval test email from Microsoft Graph Skill"},"toRecipients":[{"emailAddress":{"address":"$USER_EMAIL"}}]},"saveToSentItems":true}'`
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
**Run:** `node src/cli/mg-post.js "/me/messages/$EVAL_MSG_ID/move" '{"destinationId":"drafts"}'`
**Pass if:** stdout contains `"id"`. **Save the returned `id` as `MOVED_MSG_ID`** (the message gets a new ID after move).

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
**Run:** `node src/cli/mg-post.js "/me/events" '{"subject":"MICROSOFT_GRAPH_SKILL_EVAL_Event","start":{"dateTime":"$FUTURE_START","timeZone":"UTC"},"end":{"dateTime":"$FUTURE_END","timeZone":"UTC"}}'`
Where `$FUTURE_START` is tomorrow at 10:00 UTC (ISO 8601, e.g., `2025-07-18T10:00:00`) and `$FUTURE_END` is tomorrow at 11:00 UTC.
**Pass if:** stdout contains `"MICROSOFT_GRAPH_SKILL_EVAL_Event"` and `"id"`. **Save the event `id` as `EVENT_ID`.**

### 12 — Update event
**Depends on:** 11
**Run:** `node src/cli/mg-post.js "/me/events/$EVENT_ID" '{"subject":"MICROSOFT_GRAPH_SKILL_EVAL_Event_Updated"}' PATCH`
**Pass if:** stdout contains `"MICROSOFT_GRAPH_SKILL_EVAL_Event_Updated"`

### 13 — Get event
**Depends on:** 12
**Run:** `node src/cli/mg-get.js "/me/events/$EVENT_ID?$select=subject,start,end"`
**Pass if:** stdout contains `"MICROSOFT_GRAPH_SKILL_EVAL_Event_Updated"`

### 14 — Accept event
**Depends on:** 11
**Run:** `node src/cli/mg-post.js "/me/events/$EVENT_ID/accept" '{"comment":"Eval auto-accept","sendResponse":false}'`
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
**Pass if:** stdout contains `"value"` array where items include an `"attachments"` property (may be an empty array)

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
**Pass if:** stdout contains `"unreadItemCount"` (a number, may be 0).

**Step 2 — List unread messages:**
```
node src/cli/mg-get.js "/me/messages?$filter=isRead eq false&$select=subject,from,receivedDateTime,isRead&$top=10&$orderby=receivedDateTime desc"
```
**Pass if:** stdout contains `"value"` array. If array is non-empty, every item must have `"isRead": false`. If array is empty (no unread mail), that's still ✅ PASS.

**Step 3 — Read the first unread message (if any):**
If Step 2 returned messages, take the first message's `id` and read its body:
```
node src/cli/mg-get.js "/me/messages/{id}?$select=subject,body,from,toRecipients,receivedDateTime"
```
**Pass if:** stdout contains `"subject"` and `"body"` with `"content"` field. Skip this step if no unread messages.

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
**Overall:** [passed]/31 ([percentage]%) — [failed] failed

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

## Failures
[Details for any ❌ FAIL]
```
