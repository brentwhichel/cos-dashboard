# Facilitator guide: walking someone through their own COS dashboard

A 60-minute working session that leaves the person with a live dashboard, their email and Teams
automations running, and a plan for the first week. Use the setup guide (`setup.html`) on screen; it fills
their details into every prompt.

## Send this a few days before

- Confirm they have Microsoft 365 Copilot with Cowork, including automations (event triggers and scheduled tasks).
- Ask them to sync OneDrive to their computer and use Microsoft Edge or Google Chrome.
- Ask them to bring: their top 3 to 5 business priorities, their customer accounts and email domains, the
  3 to 6 people whose requests matter most, and the 2 to 4 people whose messages can't wait.
- Optional: if they want **Send** in the browser, ask IT for the one-time app registration in
  [`microsoft-365-sign-in.md`](microsoft-365-sign-in.md). Otherwise they can use the folder mode or build the
  app inside Cowork (Step 3).

## Agenda

| Time | What | Done when |
|---|---|---|
| 0:00 | **Show the end state.** Open the dashboard with `?demo`. Walk through My tasks, Latest check, a high-priority customer item, its research notes, editing a draft, Send with undo, Mark Complete and the filters. | They can say what they'll get. |
| 0:05 | **Step 1: constitution.** New Cowork task named *COS Email and Teams Dashboard*; paste Prompt 1 and refine it together. | Cowork confirms how it will operate. |
| 0:15 | **Customize once + Step 2.** Fill in their details on the setup page, then create the private feed. | The three files exist in OneDrive. |
| 0:20 | **Step 3: open the dashboard.** Connect with the OneDrive folder (or sign in). | The dashboard shows "No checks recorded yet". |
| 0:25 | **Steps 4 and 5: email.** Create the email screening and the customer rule + screening. Send them a test email that asks a question. | Latest check shows the test email, with a draft. |
| 0:40 | **Steps 6 and 7: Teams.** Create the key-person alerts and the scheduled sweep. Ask a colleague to message them. | The alert lands in their self-chat and on the dashboard. |
| 0:55 | **Wrap up.** Editing pauses automations; how to read Recent checks; Copilot Credits and `/cost`; book a 30-minute follow-up in a week for Step 10 (audit) and tuning. | Follow-up is on the calendar. |

## What good looks like

- Every email produces a check within a minute or two, even when it's skipped.
- On weekdays, **Last refreshed** never goes more than about 30 minutes without moving (the sweep).
- Drafts answer each question, use `[CONFIRM: ...]` placeholders instead of guessing, and end with the
  signature line.
- Nothing is ever sent without the person pressing **Send** or sending it themselves.

## Common snags

| Snag | Fix |
|---|---|
| A run waits for an approval and the dashboard goes stale | Approve it in Cowork, then run Step 9 so future runs edit files directly. |
| An automation stopped after an edit | Editing pauses automations. Re-activate it in Cowork Automations. |
| "Couldn't find your feed" | They picked the wrong folder, or Step 2 didn't run. Pick `Documents › Cowork › COS Comms Dashboard`. |
| Too many low-value items | Run the weekly tuning prompt and tighten the SKIP rules. |
| Credits higher than expected | Keep marketing out of the Inbox with Outlook rules, or run the Teams sweep hourly. |

## After the session

- **Day 1:** they check Recent checks at lunch and end of day; anything wrong goes into the tuning prompt.
- **Week 1:** run Step 10 (audit) together and apply the top three fixes.
- **Every week:** run the tuning prompt; every month, re-run the audit.
