# Build the dashboard inside Cowork (optional)

This folder holds the source of the original **COS Communications Dashboard** app, built in Cowork with
App Builder. Use it if you'd rather have the dashboard as an app inside Cowork than as this website. Its
**Send** button uses your own Office 365 Outlook and Microsoft Teams connections, so no IT app registration
is needed.

The website (`index.html`) and this app read the same feed, so you can use both.

## Files

| File | Goes to | What it is |
|---|---|---|
| `src/pages/home.tsx` | `src/pages/home.tsx` | The dashboard screen: My tasks, Latest check, filters and the day-by-day timeline |
| `src/lib/comms.ts` | `src/lib/comms.ts` | Reads the feed, saves your state and tasks, sends replies |
| `src/lib/header-slot.ts` | `src/lib/header-slot.ts` | Lets the page show "Last refreshed" in the header |
| `src/components/app-shell.tsx` | `src/components/app-shell.tsx` | Header with the app name |
| `theme.css` | the `:root` and `.dark` blocks of `src/index.css` | The teal theme |

## Settings to change

At the top of `src/lib/comms.ts`:

```ts
export const OWNER_EMAIL = "you@contoso.com";      // your work email
export const TIME_ZONE = "America/New_York";       // your IANA time zone
export const INTERNAL_DOMAINS: string[] = [];      // your company's email domains (empty = your own domain)
```

and, a few lines below, `FOLDER` if your feed isn't in `/Documents/Cowork/COS Comms Dashboard`.

## Ask Cowork to build it

Use Step 3 of the setup guide (`setup.html`), which fills in your details, or paste
[`prompts/03-optional-build-the-app-inside-cowork.md`](../prompts/03-optional-build-the-app-inside-cowork.md)
into the same Cowork task you use for your automations. Cowork adds the three data sources (OneDrive for
Business, Office 365 Outlook, Microsoft Teams), copies these files, type-checks, publishes, and gives you
the link.

## Notes

- Your organization may block apps from emailing people outside the company. If **Send** is blocked for an
  external reply, the draft is unchanged; send that one from Outlook.
- Nothing is sent unless you press **Send** on a specific item, and every send has a 5-second undo.
