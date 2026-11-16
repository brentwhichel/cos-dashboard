# Microsoft 365 sign-in (optional)

The dashboard works without any sign-in: choose **Open my OneDrive folder** and pick the synced
`Documents › Cowork › COS Comms Dashboard` folder in Microsoft Edge or Google Chrome.

Signing in adds two things:

- The feed is read straight from OneDrive, so the dashboard works on any computer, with nothing synced.
- **Send** works. Replies go out from your own Outlook and Teams after a 5-second undo, exactly like the
  original Cowork app.

Sign-in needs one app registration in your organization's Microsoft Entra ID. Someone who is allowed to
register apps does this once, for everyone in the organization.

## 1. Register the app (about 5 minutes)

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com) and go to
   **Identity › Applications › App registrations › New registration**.
2. **Name:** `COS Communications Dashboard`.
3. **Supported account types:** *Accounts in this organizational directory only*.
4. **Redirect URI:** choose the platform **Single-page application (SPA)** and enter the dashboard's address,
   for example `https://brentwhichel.github.io/cos-dashboard/` for the shared copy, or `https://<github-user>.github.io/<repo>/` for your own
   (keep the trailing slash).
5. Select **Register**, then copy the **Application (client) ID** and the **Directory (tenant) ID**.

If people also open the page as `.../index.html`, add that address as a second SPA redirect URI under
**Authentication**.

## 2. Add the permissions

Go to **API permissions › Add a permission › Microsoft Graph › Delegated permissions** and add:

| Permission | Why the dashboard needs it |
|---|---|
| `User.Read` | Sign you in and read your email address, so you show as "You" and are never added to reply-all lists |
| `Files.ReadWrite` | Read `comms-feed.json`, and save `dashboard-state.json` and `my-tasks.json` in your OneDrive |
| `Mail.Read` | Find the original email by its Internet message ID before replying in the same thread |
| `Mail.Send` | Send the reply you approved: Reply, Reply all, or a new email when the original was moved |
| `ChatMessage.Send` | Post your approved reply in a Teams chat |
| `ChannelMessage.Send` | Post your approved reply in a Teams channel thread |

These are **delegated** permissions: the page can only do what the signed-in person can already do, and
only when they press a button. If your tenant requires it, select **Grant admin consent for your
organization**.

## 3. Point the dashboard at your app

Choose one:

- **For everyone:** in your copy of this repository, edit `config.js` and set `clientId` (and `tenantId` to
  your tenant ID or domain). Commit, and GitHub Pages republishes in a minute or two.
- **Per person:** open the dashboard, select the gear button, and paste the IDs under
  **Microsoft 365 sign-in**. They're saved only in that browser.

Then choose **Sign in with Microsoft 365** on the dashboard.

## What the page does with your data

- There is no server. The page is static HTML and JavaScript; it runs in your browser and talks only to
  `login.microsoftonline.com` (sign-in) and `graph.microsoft.com` (your OneDrive, mail and Teams).
- Sign-in uses the Microsoft Authentication Library (MSAL.js) with the authorization code flow and PKCE.
  Tokens are kept in your browser's local storage until you sign out from **Settings › Sign out**.
- The page reads only the dashboard folder and writes only `dashboard-state.json` and `my-tasks.json`.
  It sends a message only when you press **Send** on one item, after a 5-second undo.
- MSAL.js is loaded from jsDelivr (`https://cdn.jsdelivr.net/npm/@azure/msal-browser@4/+esm`). If your
  organization prefers to host it, download `@azure/msal-browser` from npm, put its ES module build next to
  the page, and set `msalUrl` in `config.js` to that file.

## Troubleshooting

| What you see | What to do |
|---|---|
| `AADSTS50011` (redirect URI mismatch) | Add the exact address shown in the error as an SPA redirect URI. |
| `AADSTS65001`, "Need admin approval" | Ask an admin to grant consent for the permissions above. |
| "Couldn't find your feed" after sign-in | Check **Settings › Feed folder in OneDrive**. The default is `Documents/Cowork/COS Comms Dashboard`. |
| Send fails for one email | The draft is unchanged. If the original was moved or deleted, use **Send as new email**; otherwise reply from Outlook. |
