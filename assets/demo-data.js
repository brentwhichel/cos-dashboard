/*
 * Demo data for the COS Communications Dashboard.
 *
 * Everything here is fictional (Contoso and its customers are Microsoft's
 * sample companies). Times are generated relative to "now" so the demo always
 * looks live. Demo mode never contacts Microsoft 365 and never sends anything.
 */
(function (root) {
  "use strict";

  const OWNER = { name: "Alex Morgan", email: "alex.morgan@contoso.com" };
  const SIGNATURE = "This message was reviewed by Alex, but sent by their Cowork Chief of Staff!";

  const P = {
    alex: OWNER,
    jordan: { name: "Jordan Lee", email: "jordan.lee@contoso.com" },
    morgan: { name: "Morgan Diaz", email: "morgan.diaz@contoso.com" },
    priya: { name: "Priya Shah", email: "priya.shah@contoso.com" },
    sam: { name: "Sam Rivera", email: "sam.rivera@contoso.com" },
    taylor: { name: "Taylor Kim", email: "taylor.kim@contoso.com" },
    chris: { name: "Chris Gomez", email: "chris.gomez@contoso.com" },
    dana: { name: "Dana Whitfield", email: "dana.whitfield@fabrikam.com" },
    marcus: { name: "Marcus Chen", email: "marcus.chen@fabrikam.com" },
    lena: { name: "Lena Ortiz", email: "lena.ortiz@fabrikam.com" },
    ken: { name: "Ken Ito", email: "ken.ito@northwindtraders.com" },
    aiko: { name: "Aiko Tanaka", email: "aiko.tanaka@northwindtraders.com" },
  };

  function signed(text) {
    return text + "\n\n" + SIGNATURE;
  }

  function build(now) {
    const t = typeof now === "number" ? now : Date.now();
    const at = function (minutesAgo) {
      return new Date(t - minutesAgo * 60 * 1000).toISOString();
    };

    const items = [
      {
        id: "email:demo-fabrikam-rollout",
        channel: "email",
        receivedAt: at(24),
        addedAt: at(21),
        from: P.dana,
        to: [P.alex, P.priya],
        cc: [P.marcus, P.lena],
        toMore: 0,
        subject: "Fabrikam rollout: three questions before Thursday's steering committee",
        summary:
          "Dana wants the wave 2 start date, the frontline manager training plan and whether usage reports can be split by business unit before Thursday's 9:00 AM steering committee. Marcus and Lena from Fabrikam are copied.",
        whyItMatters: "Fabrikam - needs rollout answers by Thu",
        priority: "high",
        draft: signed(
          "Hi Dana,\n\nThanks for sending these ahead of Thursday.\n\n1. Wave 2 start date: wave 1 wraps up on the 14th, and the plan we reviewed on the 2nd starts wave 2 the following Monday. [CONFIRM: wave 2 start date with Priya]\n2. Frontline manager training: we proposed two 45-minute live sessions plus the self-paced adoption kit. Priya is holding the week of the 21st for the live sessions. [CONFIRM: session dates]\n3. Usage reports by business unit: yes. The usage reports in the Microsoft 365 admin center can be filtered by department when that attribute is filled in for your users. [CONFIRM: whether Fabrikam's directory has department populated]\n\nThanks,\nAlex",
        ),
        research: {
          notes:
            "Checked the rollout plan from the Nov 2 review, Priya's wave schedule thread and Microsoft's admin center documentation. The wave 2 and training dates still need Priya's confirmation; whether Fabrikam fills in department is unknown.",
          sources: [
            { title: "Fabrikam rollout plan v3.pptx", url: "" },
            { title: "RE: Wave schedule for Fabrikam", url: "" },
            {
              title: "Microsoft 365 reports in the admin center",
              url: "https://learn.microsoft.com/microsoft-365/admin/activity-reports/activity-reports",
            },
          ],
        },
        reply: {
          messageId: "demo-message-1",
          internetMessageId: "<demo-message-1@contoso.example>",
          conversationId: "demo-conversation-1",
          replyAll: true,
          audience: "Dana Whitfield + 3 others",
        },
        webLink: "https://outlook.office.com/mail/",
      },
      {
        id: "teams:19:demo-jordan@unq.gbl.spaces:1001",
        channel: "teams",
        receivedAt: at(52),
        addedAt: at(51),
        from: P.jordan,
        to: [],
        cc: [],
        toMore: 0,
        subject: "Chat with Jordan Lee",
        summary:
          "Jordan needs a one-line status on the three at-risk Q3 renewals before the 3:00 PM forecast call, and wants to know which one needs executive help.",
        whyItMatters: "Jordan - needs renewal risks by 3 PM",
        priority: "high",
        draft: signed(
          "Jordan, here's where the three at-risk renewals stand:\n\n1. Northwind Traders: renewal paperwork is with their procurement team and signature is expected this week. Low risk.\n2. Tailwind Traders: usage is down 18% since September and their sponsor changed roles. This is the one that needs executive help. [CONFIRM: ask Morgan to call their new CIO?]\n3. Fabrikam: tied to Thursday's steering committee and on track if the wave 2 dates hold.\n\n[CONFIRM: do you want me on the 3:00 PM call?]",
        ),
        research: {
          notes:
            "Read the renewal tracker, last week's Tailwind usage report and Thursday's Fabrikam thread. Tailwind's sponsor change came from the account team's Teams chat on Monday.",
          sources: [
            { title: "Q3 renewal tracker.xlsx", url: "" },
            { title: "Tailwind Traders usage report - October", url: "" },
            { title: "Tailwind account team (Teams chat)", url: "" },
          ],
        },
        reply: {
          chatId: "19:demo-jordan@unq.gbl.spaces",
          messageId: "1001",
          audience: "Teams chat with Jordan Lee",
        },
        webLink: "https://teams.microsoft.com/",
      },
      {
        id: "email:demo-northwind-agenda",
        channel: "email",
        receivedAt: at(130),
        addedAt: at(127),
        from: P.sam,
        to: [P.alex],
        cc: [P.priya],
        toMore: 0,
        subject: "Agenda review: Northwind kickoff workshop",
        summary:
          "Sam shared the draft agenda for next Tuesday's Northwind kickoff workshop and asked you to confirm the order of the executive and technical sessions.",
        whyItMatters: "Sam - agenda sign-off for Northwind",
        priority: "medium",
        draft: signed(
          "Sam,\n\nThanks for pulling the agenda together. The order works. Let's keep the executive session first so their sponsor can leave after 10:30.\n\nThanks,\nAlex",
        ),
        research: null,
        reply: {
          messageId: "demo-message-3",
          internetMessageId: "<demo-message-3@contoso.example>",
          conversationId: "demo-conversation-3",
          replyAll: false,
          audience: "Sam Rivera",
        },
        webLink: "https://outlook.office.com/mail/",
      },
      {
        id: "teams:19:demo-northwind@thread.v2:2002",
        channel: "teams",
        receivedAt: at(185),
        addedAt: at(181),
        from: P.ken,
        to: [P.sam, P.priya, P.aiko],
        cc: [],
        toMore: 2,
        subject: "Northwind Traders rollout",
        summary:
          "Ken thanked the team for turning the pilot report around so quickly. There's no ask of you, and you already replied in the chat.",
        whyItMatters: "Northwind - thanks for quick turnaround",
        priority: "low",
        draft: signed("Ken, glad it helped! Shout if anything in the report needs a second look."),
        research: null,
        reply: {
          chatId: "19:demo-northwind@thread.v2",
          messageId: "2002",
          audience: "Northwind Traders rollout group chat (7 people, 3 external)",
        },
        webLink: "https://teams.microsoft.com/",
      },
      {
        id: "teams:19:demo-channel@thread.tacv2:3003",
        channel: "teams",
        receivedAt: at(262),
        addedAt: at(259),
        from: P.taylor,
        to: [P.chris, P.priya],
        cc: [],
        toMore: 0,
        subject: "Customer Success > Wins and Lessons",
        summary:
          "Taylor @mentioned you asking for the adoption framework one-pager you presented last week, so the team can reuse it with two new accounts.",
        whyItMatters: "Taylor - wants your adoption one-pager",
        priority: "medium",
        draft: signed(
          "Taylor, happy to share! Here's the one-pager: [LINK: Adoption framework one-pager]. The second page has the 30-60-90 day checklist the accounts liked most.",
        ),
        research: {
          notes: "Found the one-pager in your OneDrive (updated last Thursday). It's internal, and the thread is internal only.",
          sources: [{ title: "Adoption framework one-pager.docx", url: "" }],
        },
        reply: {
          teamId: "demo-team",
          channelId: "19:demo-channel@thread.tacv2",
          parentMessageId: "3000",
          messageId: "3003",
          audience: "Customer Success > Wins and Lessons thread",
        },
        webLink: "https://teams.microsoft.com/",
      },
      {
        id: "email:demo-tailwind-deck",
        channel: "email",
        receivedAt: at(22 * 60),
        addedAt: at(22 * 60 - 4),
        from: P.morgan,
        to: [P.alex],
        cc: [P.jordan],
        toMore: 0,
        subject: "Tailwind Traders exec briefing deck: final review",
        summary:
          "Morgan asked for your final edits on the Tailwind Traders executive briefing deck before it goes to their CIO on Friday.",
        whyItMatters: "Morgan - final deck edits before Friday",
        priority: "high",
        draft: signed(
          "Morgan,\n\nMy edits are in the deck. The two changes that matter: the adoption chart now shows the last 90 days, and the ask on slide 6 is a single named sponsor.\n\nThanks,\nAlex",
        ),
        research: null,
        reply: {
          messageId: "demo-message-6",
          internetMessageId: "<demo-message-6@contoso.example>",
          conversationId: "demo-conversation-6",
          replyAll: false,
          audience: "Morgan Diaz",
        },
        webLink: "https://outlook.office.com/mail/",
      },
      {
        id: "email:demo-tailwind-qbr",
        channel: "email",
        receivedAt: at(26 * 60),
        addedAt: at(26 * 60 - 3),
        from: P.priya,
        to: [P.alex],
        cc: [],
        toMore: 0,
        subject: "Can you cover the usage section of the Tailwind QBR?",
        summary:
          "Priya asked whether you can present the usage and adoption section of the Tailwind Traders QBR on Wednesday at 2:00 PM. Your calendar is open then.",
        whyItMatters: "Priya - QBR usage section on Wed",
        priority: "medium",
        draft: signed(
          "Priya,\n\nThanks for thinking of me. Wednesday at 2:00 PM is open on my calendar. [CONFIRM: yes, I'll cover the usage section]\n\nCould you send the latest usage export by Tuesday so I can build the two charts?\n\nThanks,\nAlex",
        ),
        research: {
          notes: "Checked your calendar: Wednesday 2:00-3:00 PM is free. The last QBR deck is in the Tailwind account folder.",
          sources: [
            { title: "Calendar: Wednesday afternoon", url: "" },
            { title: "Tailwind Traders QBR - Q2.pptx", url: "" },
          ],
        },
        reply: {
          messageId: "demo-message-7",
          internetMessageId: "<demo-message-7@contoso.example>",
          conversationId: "demo-conversation-7",
          replyAll: false,
          audience: "Priya Shah",
        },
        webLink: "https://outlook.office.com/mail/",
      },
      {
        id: "email:demo-fabrikam-contacts",
        channel: "email",
        receivedAt: at(50 * 60),
        addedAt: at(50 * 60 - 5),
        from: P.lena,
        to: [P.alex, P.sam],
        cc: [],
        toMore: 0,
        subject: "Updated Fabrikam project contact list",
        summary: "Lena shared Fabrikam's updated project contact list. There's no ask; it adds two new IT leads.",
        whyItMatters: "Fabrikam - FYI new project contacts",
        priority: "low",
        draft: signed("Hi Lena,\n\nThanks for the updated list. We'll add the two new IT leads to next week's invites.\n\nThanks,\nAlex"),
        research: null,
        reply: {
          messageId: "demo-message-8",
          internetMessageId: "<demo-message-8@contoso.example>",
          conversationId: "demo-conversation-8",
          replyAll: true,
          audience: "Lena Ortiz + 1 other",
        },
        webLink: "https://outlook.office.com/mail/",
      },
    ];

    function run(minutesAgo, source, outcome, subject, from, note) {
      return { at: at(minutesAgo), source: source, outcome: outcome, subject: subject, from: from, note: note };
    }

    const recentRuns = [
      run(280, "email", "skipped", "Your weekly analytics digest", "Viva Insights", "Automated digest; no ask; marked read"),
      run(262, "teams-sweep", "added", "Teams sweep: 6 chats reviewed, 1 added", "", "Added Taylor Kim @mention in Wins and Lessons"),
      run(230, "email", "skipped", "Lunch and learn: prompting tips", "Contoso Learning", "Broad optional invite; marked read"),
      run(184, "teams-sweep", "added", "Teams sweep: 4 chats reviewed, 1 added", "", "Added Ken Ito message in Northwind Traders rollout"),
      run(128, "email", "added", "Agenda review: Northwind kickoff workshop", "Sam Rivera", "Agenda sign-off requested; draft ready"),
      run(100, "email", "skipped", "Expense report approved", "Contoso Expense", "Automated notice; no ask; marked read"),
      run(80, "teams-sweep", "skipped", "Teams sweep: 3 chats reviewed, 0 added", "", "No chats needed you"),
      run(51, "teams-chat", "added", "Chat with Jordan Lee", "Jordan Lee", "Renewal risks needed by 3 PM; alert sent"),
      run(50, "teams-sweep", "skipped", "Teams sweep: 2 chats reviewed, 0 added", "", "Jordan's chat already captured"),
      run(21, "email", "added", "Fabrikam rollout: three questions before Thursday's steering committee", "Dana Whitfield", "Customer asks before Thursday; left unread"),
      run(20, "teams-sweep", "skipped", "Teams sweep: 5 chats reviewed, 0 added", "", "No chats needed you"),
      run(9, "email", "skipped", "Flash sale: 30% off travel gear", "Adventure Works", "Marketing; no ask; marked read"),
      run(3, "email", "skipped", "Your order has shipped", "Proseware Store", "Personal notice; left unread"),
    ];

    const feed = {
      version: 1,
      description: "Demo feed (fictional). Appended by the COS automations; read by the COS Communications Dashboard.",
      items: items,
      runsDescription:
        "Every COS check (email screening, key Teams chat, Teams sweep) sets lastRun and appends one entry to recentRuns, even when nothing is added.",
      recentRuns: recentRuns,
      lastRun: recentRuns[recentRuns.length - 1],
    };

    const state = {
      version: 1,
      completed: { "email:demo-tailwind-deck": at(20 * 60) },
      deleted: {},
      sent: {
        "email:demo-northwind-agenda": {
          at: at(95),
          text: signed(
            "Sam,\n\nThanks for pulling the agenda together. The order works. Let's keep the executive session first so their sponsor can leave after 10:30.\n\nThanks,\nAlex",
          ),
          via: "reply",
        },
        "email:demo-tailwind-deck": {
          at: at(20 * 60 + 10),
          text: signed(
            "Morgan,\n\nMy edits are in the deck. The two changes that matter: the adoption chart now shows the last 90 days, and the ask on slide 6 is a single named sponsor.\n\nThanks,\nAlex",
          ),
          via: "reply",
        },
      },
      drafts: {},
    };

    const tasks = {
      version: 1,
      tasks: [
        { id: "demo-task-1", text: "Send Fabrikam steering committee pre-read", createdAt: at(170) },
        { id: "demo-task-2", text: "Ask Morgan about an exec sponsor call for Tailwind Traders", createdAt: at(300) },
        { id: "demo-task-3", text: "Book travel for the Northwind onsite", createdAt: at(26 * 60) },
        { id: "demo-task-4", text: "Update the Tailwind Traders account plan", createdAt: at(48 * 60), completedAt: at(21 * 60) },
        { id: "demo-task-5", text: "Share the Q3 renewal forecast with Jordan", createdAt: at(52 * 60), completedAt: at(30 * 60) },
      ],
    };

    return { feed: feed, state: state, tasks: tasks };
  }

  root.COS_DEMO = { owner: OWNER, internalDomains: ["contoso.com"], build: build };
})(typeof window !== "undefined" ? window : globalThis);
