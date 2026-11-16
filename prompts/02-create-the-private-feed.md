# Step 2: Create your private feed

The dashboard reads three small files in your OneDrive. Your automations write the feed; the dashboard keeps your tasks and what you've sent, completed or deleted.

**Where:** Stay in the COS Email and Teams Dashboard task and paste this prompt.

```text
Set up the private feed for my COS Communications Dashboard.

In my cross-session Cowork folder (OneDrive: Documents/Cowork), create a folder named "COS Comms Dashboard" containing the three files below, exactly as shown. Create them by editing files directly in my Cowork folder, never by uploading a downloaded copy. If any of them already exists, leave it unchanged and tell me.

1. comms-feed.json
{
 "version": 1,
 "description": "Private feed of messages that meet my Chief of Staff constitution. Appended by my COS automations; read by my COS Communications Dashboard. Do not edit by hand.",
 "items": [],
 "runsDescription": "Every COS check (email screening, key Teams chat, Teams sweep) sets lastRun and appends one entry to recentRuns, even when nothing is added, so the dashboard always shows when it was last refreshed. recentRuns keeps the newest 200 entries.",
 "recentRuns": [],
 "lastRun": null
}

2. teams-sweep.json
{
 "version": 1,
 "description": "Sweep marker for the COS Comms Dashboard Teams sweep. Only lastSweepAt is updated by each sweep run.",
 "lastSweepAt": "<the current time, ISO 8601 with my local UTC offset>"
}

3. my-tasks.json
{
 "version": 1,
 "tasks": []
}

Then tell me the OneDrive path of each file. Do not create any automations yet.
```

**After Cowork answers:** Open OneDrive and check that Documents › Cowork › COS Comms Dashboard holds comms-feed.json, teams-sweep.json and my-tasks.json.
