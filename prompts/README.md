# Setup prompts

Copy-ready prompts for building your own COS Communications Dashboard in Cowork, in order.
The easiest way to use them is the setup guide (`setup.html`), which fills in your details for you.
Here, anything in [BRACKETS] is a placeholder to replace before you paste.

| Step | Prompt | What it does |
|---|---|---|
| Step 1 | [Install your Chief of Staff constitution](01-chief-of-staff-constitution.md) | Every automation below works "under my approved Chief of Staff constitution". This is where you define it. Skip this if you already have one in Cowork. |
| Step 2 | [Create your private feed](02-create-the-private-feed.md) | The dashboard reads three small files in your OneDrive. Your automations write the feed; the dashboard keeps your tasks and what you've sent, completed or deleted. |
| Step 3 (optional) | [Or build the same dashboard inside Cowork](03-optional-build-the-app-inside-cowork.md) | Prefer an app that lives in Cowork instead of this website? Cowork can rebuild the original app from this repository. Its Send button uses your own Outlook and Teams connections, so no IT app registration is needed. |
| Step 4 | [Screen every email into the dashboard](04-email-screening.md) | Every new email gets one decision: add it with a researched draft, or skip it with a reason. Marketing and no-ask mail is marked read; customer, personal and still-open items stay unread. |
| Step 5 | [Never miss a customer email](05-customer-email-screening.md) | Customer email gets its own Outlook folder and its own automation, so every customer message reaches the dashboard and nothing from a customer is quietly marked as read. |
| Step 6 | [Instant alerts for the people who can't wait](06-key-person-chat-alerts.md) | One automation per key person. Their 1:1 Teams messages land on the dashboard right away, with a private heads-up in your own Teams self-chat. |
| Step 7 | [Sweep every Teams chat on a schedule](07-teams-sweep.md) | Group, meeting and customer chats move fast. The sweep picks up where the last run ended, adds anything that needs you, every customer message and every @mention. |
| Step 8 (only if you wrote your own rules) | [Research every draft and sign it](08-upgrade-research-and-signature.md) | Already built into Steps 4 to 7. Use this to upgrade automations you created another way, so every draft is researched, answers point by point, uses placeholders instead of guessing, and carries your signature line. |
| Step 9 (only if you wrote your own rules) | [Prove it's running, and never stall](09-heartbeat-and-no-stalls.md) | Already built into Steps 4 to 7. A visible Last refreshed time proves the system is alive on quiet days, and direct file edits keep background runs from waiting hours for an approval nobody sees. |
| Step 10 | [Audit the whole system](10-audit.md) | A regular audit catches wrong skips, missed customer messages and weak drafts before they cost trust. |
| Keep improving | [Tune the dashboard every week](11-weekly-tuning.md) | The dashboard gets sharper every week as Cowork learns what you act on, edit and ignore. |

The automation prompts (Steps 4 to 7) include the exact feed format the dashboard reads; see `docs/feed-format.md`.
