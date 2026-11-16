# Step 3 (optional): Or build the same dashboard inside Cowork

Prefer an app that lives in Cowork instead of this website? Cowork can rebuild the original app from this repository. Its Send button uses your own Outlook and Teams connections, so no IT app registration is needed.

**Where:** Stay in the same Cowork task and paste this prompt. Approve the Outlook, Teams and OneDrive connections when Cowork asks.

```text
Build my COS Communications Dashboard as an App Builder app named "COS Communications Dashboard".

1. Add three data sources: OneDrive for Business, Office 365 Outlook and Microsoft Teams.
2. Use the source in https://github.com/brentwhichel/cos-dashboard/tree/main/cowork-app exactly. Copy these files into the app, replacing the template's versions:
   - cowork-app/src/pages/home.tsx -> src/pages/home.tsx
   - cowork-app/src/lib/comms.ts -> src/lib/comms.ts
   - cowork-app/src/lib/header-slot.ts -> src/lib/header-slot.ts
   - cowork-app/src/components/app-shell.tsx -> src/components/app-shell.tsx
   - the theme colors in cowork-app/theme.css -> the :root and .dark blocks of src/index.css
3. In src/lib/comms.ts set OWNER_EMAIL to "[YOUR_EMAIL]", TIME_ZONE to "[YOUR_TIME_ZONE, e.g. America/New_York]", INTERNAL_DOMAINS to my company email domains, and FOLDER to "/Documents/Cowork/COS Comms Dashboard".
4. Type-check it, publish it, and give me the link. Keep it private to me.

Nothing is ever sent unless I press Send on a specific item.
```

**After Cowork answers:** Open the link Cowork returns and bookmark it. The app reads the same feed, so you can use it and the website side by side.
