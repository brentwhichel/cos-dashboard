/*
 * COS Communications Dashboard - site settings.
 *
 * Everything here is optional. Each person can also change these from the
 * gear button on the dashboard; their choices are saved in their own browser.
 * Never put personal data, passwords or secrets in this file - it is public.
 */
window.COS_DASHBOARD_CONFIG = {
  // Name shown in the header.
  appTitle: "COS Communications",

  // Where each person's COS automations keep the feed, relative to the root of
  // their OneDrive. Cowork's own folder is Documents/Cowork.
  feedFolder: "Documents/Cowork/COS Comms Dashboard",

  // Microsoft 365 sign-in (optional - needed only to send replies from the
  // dashboard). Paste the Application (client) ID of your organization's app
  // registration here, or let each person enter it under Settings.
  // See docs/microsoft-365-sign-in.md.
  clientId: "",
  // "organizations" lets any work account sign in. Use your tenant ID or
  // domain (for example "contoso.onmicrosoft.com") to limit it to your company.
  tenantId: "organizations",

  // Email domains that count as "inside the company". Anyone else is flagged
  // External before a reply goes out. Leave empty to use each person's own
  // email domain.
  internalDomains: [],

  // IANA time zone such as "America/New_York". Leave empty to use the
  // reader's own browser time zone.
  timeZone: "",

  // How often the page re-reads the feed, in seconds.
  pollSeconds: 60,
  // Seconds of "Undo" before a reply is actually sent.
  undoSeconds: 5,
  // Weekday business hours (24-hour clock). If no check lands for staleHours
  // inside these hours, the dashboard warns that the automations may be paused.
  businessHoursStart: 8,
  businessHoursEnd: 18,
  staleHours: 2,

  // Show the research notes your automations save with each draft (only you
  // ever see them). Set to false to hide them.
  showResearchNotes: true,
};
