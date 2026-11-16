# Feed format

Everything lives in one folder in your OneDrive, by default `Documents/Cowork/COS Comms Dashboard`.
Each file has exactly one writer, so the automations and the dashboard never overwrite each other.

| File | Written by | Holds |
|---|---|---|
| `comms-feed.json` | Your Cowork automations (append only) | Every item that needs you, plus the run log |
| `dashboard-state.json` | The dashboard | What you sent, completed, deleted or edited |
| `my-tasks.json` | The dashboard | My tasks |
| `teams-sweep.json` | The Teams sweep | Where the last sweep ended |

Empty starter copies are in [`starter-files/`](../starter-files). All times are ISO 8601 with a UTC offset,
for example `2026-10-14T15:42:00-04:00`.

## comms-feed.json

```json
{
  "version": 1,
  "description": "Private feed of messages that meet my Chief of Staff constitution.",
  "items": [ ... ],
  "runsDescription": "Every COS check sets lastRun and appends one entry to recentRuns.",
  "recentRuns": [ ... ],
  "lastRun": { ... }
}
```

### Items

One object per message (or burst of messages) that needs you. Items are only ever appended.

| Field | Type | Notes |
|---|---|---|
| `id` | string | `email:<message id>`, `teams:<chat id>:<message id>` or `teams:<channel id>:<message id>`. Unique. |
| `channel` | `"email"` or `"teams"` | |
| `receivedAt` | ISO time | When the message arrived. Items are grouped by this day. |
| `addedAt` | ISO time | When the automation added it. |
| `from` | `{ name, email }` | The sender. |
| `to`, `cc` | `[{ name, email }]` | Up to 10 each. For Teams, the other chat members. |
| `toMore` | number | How many more recipients weren't listed. |
| `subject` | string | Email subject, chat topic, or `Team > Channel`. |
| `summary` | string | At most two sentences: who wants what, by when. |
| `whyItMatters` | string | At most eight words, shown in bold above the subject. |
| `priority` | `"high"`, `"medium"` or `"low"` | |
| `draft` | string | The proposed reply, plain text, ending with your signature line. |
| `research` | `{ notes, sources: [{ title, url }] }` | Private notes for you; shown under **Research notes**, never sent. Optional. |
| `reply` | object | Where Send goes; see below. |
| `webLink` | URL | Opens the original in Outlook or Teams. |

`reply` for email:

```json
{ "messageId": "...", "internetMessageId": "<...@...>", "conversationId": "...", "replyAll": true, "audience": "Dana Whitfield + 3 others" }
```

`reply` for a Teams chat:

```json
{ "chatId": "19:...@thread.v2", "messageId": "1793...", "audience": "Teams chat with Jordan Lee" }
```

`reply` for a Teams channel thread:

```json
{ "teamId": "...", "channelId": "19:...@thread.tacv2", "parentMessageId": "...", "messageId": "...", "audience": "Team > Channel thread" }
```

### Run log

Every automation run, including runs that add nothing, appends one entry to `recentRuns` (newest last,
at most 200) and copies it to `lastRun`. The dashboard shows it as **Latest check** and **Last refreshed**,
and warns when no check has landed for two business hours.

```json
{
  "at": "2026-11-06T12:27:56-05:00",
  "source": "email | teams-chat | teams-sweep",
  "outcome": "added | skipped | error",
  "subject": "Teams sweep: 4 chats reviewed, 1 added",
  "from": "",
  "note": "Added Jordan Lee's renewal question"
}
```

## dashboard-state.json

Written only by the dashboard. Keys are item ids.

```json
{
  "version": 1,
  "completed": { "<item id>": "<when>" },
  "deleted": { "<item id>": "<when>" },
  "sent": { "<item id>": { "at": "<when>", "text": "<what was sent>", "via": "reply | replyAll | newEmail | teams | manual", "added": ["<names added to a sender-only reply>"] } },
  "drafts": { "<item id>": "<your edited draft>" }
}
```

`manual` means you sent the reply yourself in Outlook or Teams and marked it replied on the dashboard.

## my-tasks.json

```json
{ "version": 1, "tasks": [ { "id": "<uuid>", "text": "Book travel for the onsite", "createdAt": "<when>", "completedAt": "<when, if done>" } ] }
```

## teams-sweep.json

```json
{ "version": 1, "description": "Sweep marker for the COS Comms Dashboard Teams sweep.", "lastSweepAt": "<when the last sweep started>" }
```

The sweep only moves `lastSweepAt` forward.
