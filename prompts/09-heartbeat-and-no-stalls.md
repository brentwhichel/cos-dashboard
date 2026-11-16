# Step 9 (only if you wrote your own rules): Prove it's running, and never stall

Already built into Steps 4 to 7. A visible Last refreshed time proves the system is alive on quiet days, and direct file edits keep background runs from waiting hours for an approval nobody sees.

**Where:** Stay in the same task. Re-activate each automation after Cowork updates it.

```text
Make my Email and Teams Dashboard prove it is running, and make its background runs finish on their own.

Every run of every automation (both email screenings, each key-person alert and the Teams sweep) must:
1. Set lastRun in the feed: the time, the source, the outcome (added or skipped) and the subject.
2. Add one entry to the feed's recentRuns log, even when nothing is added. Keep the newest 200 entries.
3. Never claim a refresh that was not saved. If a save fails, say so in the run result.

Some ways of saving files ask for my approval, and a background run can stall for hours waiting for it. Update every automation so that:
- The feed and the sweep marker are saved by editing the files in my Cowork folder directly, never by downloading and re-uploading them.
- The sweep marker never moves backwards.
- Runs still never send, post, reply or delete anything.

Show me the updated instructions for each automation before you save them.
```

**After Cowork answers:** Re-activate each automation in Cowork Automations after Cowork updates it.
