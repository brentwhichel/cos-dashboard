/*
 * COS Communications Dashboard - core helpers (no DOM).
 *
 * Data rules, formatting and recipient logic shared by the dashboard page.
 * Ported from the original COS Communications Dashboard app so every item,
 * check and task renders exactly the same way.
 */
(function (root) {
  "use strict";

  // ---------------------------------------------------------------------------
  // Small utilities
  // ---------------------------------------------------------------------------

  function str(value) {
    return typeof value === "string" ? value : "";
  }

  function stripBom(text) {
    return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  }

  function parseJsonText(text) {
    return JSON.parse(stripBom(String(text || "").trim()));
  }

  // ---------------------------------------------------------------------------
  // Feed normalisation - the feed is written by an automation, so validate
  // every field and drop items that lack the essentials.
  // ---------------------------------------------------------------------------

  function person(value) {
    if (!value || typeof value !== "object") return null;
    const name = str(value.name).trim();
    const email = str(value.email).trim();
    if (!name && !email) return null;
    return { name: name || email, email: email };
  }

  function people(value) {
    if (!Array.isArray(value)) return [];
    return value.map(person).filter(Boolean);
  }

  function priority(value) {
    const v = str(value).toLowerCase();
    return v === "high" || v === "medium" || v === "low" ? v : null;
  }

  const REPLY_KEYS = [
    "messageId",
    "internetMessageId",
    "conversationId",
    "chatId",
    "teamId",
    "channelId",
    "parentMessageId",
    "audience",
  ];

  function replyTarget(value) {
    if (!value || typeof value !== "object") return {};
    const out = {};
    for (const key of REPLY_KEYS) {
      const s = str(value[key]).trim();
      if (s) out[key] = s;
    }
    if (typeof value.replyAll === "boolean") out.replyAll = value.replyAll;
    return out;
  }

  function researchOf(value) {
    if (!value || typeof value !== "object") return null;
    const notes = str(value.notes).trim();
    const sources = Array.isArray(value.sources)
      ? value.sources
          .map(function (s) {
            if (!s || typeof s !== "object") return null;
            const title = str(s.title).trim();
            const url = str(s.url).trim();
            return title || url ? { title: title || url, url: url } : null;
          })
          .filter(Boolean)
          .slice(0, 10)
      : [];
    if (!notes && sources.length === 0) return null;
    return { notes: notes, sources: sources };
  }

  function normalizeItem(raw) {
    if (!raw || typeof raw !== "object") return null;
    const id = str(raw.id).trim();
    const channel = str(raw.channel).toLowerCase();
    const receivedAt = str(raw.receivedAt).trim();
    const receivedMs = Date.parse(receivedAt);
    if (!id || (channel !== "email" && channel !== "teams") || Number.isNaN(receivedMs)) {
      return null;
    }
    const from = person(raw.from) || { name: "Unknown sender", email: "" };
    const toMore = typeof raw.toMore === "number" && raw.toMore > 0 ? Math.floor(raw.toMore) : 0;
    return {
      id: id,
      channel: channel,
      receivedAt: receivedAt,
      receivedMs: receivedMs,
      from: from,
      to: people(raw.to),
      cc: people(raw.cc),
      toMore: toMore,
      subject: str(raw.subject).trim() || "(no subject)",
      summary: str(raw.summary).trim(),
      whyItMatters: str(raw.whyItMatters).trim(),
      priority: priority(raw.priority),
      draft: str(raw.draft),
      research: researchOf(raw.research),
      reply: replyTarget(raw.reply),
      webLink: str(raw.webLink).trim() || undefined,
    };
  }

  function runSource(value) {
    const v = str(value).trim().toLowerCase().replace(/[\s_]+/g, "-");
    if (v === "email" || v === "mail") return "email";
    if (v === "teams-chat" || v === "teams" || v === "chat") return "teams-chat";
    if (v === "teams-sweep" || v === "sweep") return "teams-sweep";
    return "other";
  }

  function runOutcome(value) {
    const v = str(value).trim().toLowerCase();
    if (v.indexOf("add") === 0) return "added";
    if (/^(err|fail|block|incomplete|stall)/.test(v)) return "error";
    return "skipped";
  }

  function normalizeRun(raw) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
    const at = str(raw.at).trim();
    const atMs = Date.parse(at);
    if (!at || Number.isNaN(atMs)) return null;
    return {
      at: at,
      atMs: atMs,
      source: runSource(raw.source),
      outcome: runOutcome(raw.outcome),
      subject: str(raw.subject).trim(),
      from: str(raw.from).trim(),
      note: str(raw.note).trim(),
    };
  }

  function normalizeFeed(json) {
    const rootObj = json && typeof json === "object" && !Array.isArray(json) ? json : {};
    if (!Array.isArray(rootObj.items)) throw new Error("Feed file has no items array");
    const seen = new Set();
    const items = [];
    for (const raw of rootObj.items) {
      const item = normalizeItem(raw);
      if (!item || seen.has(item.id)) continue;
      seen.add(item.id);
      items.push(item);
    }
    items.sort(function (a, b) {
      return b.receivedMs - a.receivedMs;
    });

    // Checks are appended oldest -> newest; reverse first so same-minute
    // entries keep newest-first order through the (stable) sort.
    const rawRuns = Array.isArray(rootObj.recentRuns) ? rootObj.recentRuns.slice().reverse() : [];
    const runKeys = new Set();
    const runs = [];
    for (const raw of rawRuns) {
      const run = normalizeRun(raw);
      if (!run) continue;
      const key = run.atMs + "|" + run.source + "|" + run.outcome + "|" + run.subject;
      if (runKeys.has(key)) continue;
      runKeys.add(key);
      runs.push(run);
    }
    runs.sort(function (a, b) {
      return b.atMs - a.atMs;
    });
    const declared = normalizeRun(rootObj.lastRun);
    const newest = runs[0] || null;
    const lastRun = declared && (!newest || declared.atMs >= newest.atMs) ? declared : newest;
    return { items: items, runs: runs, lastRun: lastRun };
  }

  // ---------------------------------------------------------------------------
  // Dashboard state (completed / deleted / sent / edited drafts)
  // ---------------------------------------------------------------------------

  function emptyState() {
    return { version: 1, completed: {}, deleted: {}, sent: {}, drafts: {} };
  }

  function cloneState(state) {
    return {
      version: 1,
      completed: Object.assign({}, state.completed),
      deleted: Object.assign({}, state.deleted),
      sent: Object.assign({}, state.sent),
      drafts: Object.assign({}, state.drafts),
    };
  }

  const SEND_VIAS = ["reply", "replyAll", "newEmail", "teams", "manual"];

  function normalizeState(json) {
    const state = emptyState();
    if (!json || typeof json !== "object") return state;
    function stringMap(value) {
      const out = {};
      if (value && typeof value === "object" && !Array.isArray(value)) {
        for (const key of Object.keys(value)) {
          if (typeof value[key] === "string") out[key] = value[key];
        }
      }
      return out;
    }
    state.completed = stringMap(json.completed);
    state.deleted = stringMap(json.deleted);
    state.drafts = stringMap(json.drafts);
    if (json.sent && typeof json.sent === "object" && !Array.isArray(json.sent)) {
      for (const key of Object.keys(json.sent)) {
        const r = json.sent[key];
        if (!r || typeof r !== "object") continue;
        if (typeof r.at !== "string" || typeof r.text !== "string") continue;
        const via = str(r.via);
        const added = Array.isArray(r.added)
          ? r.added.filter(function (n) {
              return typeof n === "string" && n.trim() !== "";
            })
          : [];
        const record = { at: r.at, text: r.text, via: SEND_VIAS.indexOf(via) >= 0 ? via : "reply" };
        if (added.length > 0) record.added = added;
        state.sent[key] = record;
      }
    }
    return state;
  }

  // ---------------------------------------------------------------------------
  // My tasks
  // ---------------------------------------------------------------------------

  const MAX_TASK_LENGTH = 500;

  function cleanTaskText(value) {
    return String(value || "")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, MAX_TASK_LENGTH);
  }

  function newTaskId() {
    if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
      return crypto.randomUUID();
    }
    return Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 10);
  }

  function normalizeTasks(json) {
    const list = Array.isArray(json)
      ? json
      : json && typeof json === "object" && Array.isArray(json.tasks)
        ? json.tasks
        : [];
    const seen = new Set();
    const out = [];
    for (const raw of list) {
      if (!raw || typeof raw !== "object" || Array.isArray(raw)) continue;
      const id = str(raw.id).trim();
      const text = cleanTaskText(str(raw.text));
      if (!id || !text || seen.has(id)) continue;
      seen.add(id);
      const createdAt = str(raw.createdAt).trim();
      const completedAt = str(raw.completedAt).trim();
      const task = { id: id, text: text, createdAt: createdAt };
      if (completedAt) task.completedAt = completedAt;
      out.push(task);
    }
    return out;
  }

  // ---------------------------------------------------------------------------
  // Formatting - always 12-hour time with AM/PM in the reader's time zone.
  // ---------------------------------------------------------------------------

  function resolveTimeZone(wanted) {
    const tz = String(wanted || "").trim();
    if (tz) {
      try {
        new Intl.DateTimeFormat("en-US", { timeZone: tz });
        return tz;
      } catch (e) {
        /* fall through to the browser's zone */
      }
    }
    try {
      return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
    } catch (e) {
      return "UTC";
    }
  }

  function makeFormat(timeZone, hours) {
    const tz = resolveTimeZone(timeZone);
    const startHour = hours && typeof hours.start === "number" ? hours.start : 8;
    const endHour = hours && typeof hours.end === "number" ? hours.end : 18;
    const dayKeyFormat = new Intl.DateTimeFormat("en-CA", {
      timeZone: tz,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    });
    const dayLabelFormat = new Intl.DateTimeFormat("en-US", {
      timeZone: tz,
      weekday: "long",
      month: "long",
      day: "numeric",
      year: "numeric",
    });
    const timeFormat = new Intl.DateTimeFormat("en-US", {
      timeZone: tz,
      hour: "numeric",
      minute: "2-digit",
      hour12: true,
      timeZoneName: "short",
    });
    const shortTimeFormat = new Intl.DateTimeFormat("en-US", {
      timeZone: tz,
      hour: "numeric",
      minute: "2-digit",
      hour12: true,
    });
    const compactDayFormat = new Intl.DateTimeFormat("en-US", {
      timeZone: tz,
      weekday: "short",
      month: "short",
      day: "numeric",
    });
    const monthDayFormat = new Intl.DateTimeFormat("en-US", { timeZone: tz, month: "short", day: "numeric" });
    const weekdayHourFormat = new Intl.DateTimeFormat("en-US", {
      timeZone: tz,
      weekday: "short",
      hour: "numeric",
      hourCycle: "h23",
    });

    function dayKey(ms) {
      return dayKeyFormat.format(ms);
    }
    function dayLabel(ms) {
      return dayLabelFormat.format(ms);
    }
    function relativeDay(ms, now) {
      const key = dayKey(ms);
      if (key === dayKey(now)) return "Today";
      if (key === dayKey(now - 24 * 60 * 60 * 1000)) return "Yesterday";
      return null;
    }
    function timeLabel(ms) {
      return timeFormat.format(ms);
    }
    function shortTime(ms) {
      return shortTimeFormat.format(ms);
    }
    /* "Thursday, October 15, 2026 · 9:48 AM EDT" */
    function fullStamp(ms) {
      return dayLabelFormat.format(ms) + " \u00b7 " + timeFormat.format(ms);
    }
    /* "Thu, Oct 15 · 9:48 AM EDT" - for narrow screens. */
    function compactStamp(ms) {
      return compactDayFormat.format(ms) + " \u00b7 " + timeFormat.format(ms);
    }
    function relativeAge(ms, now) {
      const minutes = Math.floor(Math.max(0, now - ms) / 60000);
      if (minutes < 1) return "just now";
      if (minutes < 60) return minutes + " min ago";
      const hrs = Math.floor(minutes / 60);
      if (hrs < 24) return hrs + " hr ago";
      const days = Math.floor(hrs / 24);
      return days === 1 ? "1 day ago" : days + " days ago";
    }
    /* Weekdays during business hours, when checks should arrive steadily. */
    function isBusinessHours(ms) {
      const parts = weekdayHourFormat.formatToParts(ms);
      let weekday = "";
      let hour = 0;
      for (const p of parts) {
        if (p.type === "weekday") weekday = p.value;
        if (p.type === "hour") hour = Number(p.value) % 24;
      }
      return weekday !== "Sat" && weekday !== "Sun" && hour >= startHour && hour < endHour;
    }
    function stampLabel(iso, now) {
      const ms = Date.parse(iso);
      if (Number.isNaN(ms)) return "";
      const today = dayKey(ms) === dayKey(typeof now === "number" ? now : Date.now());
      return today ? timeLabel(ms) : monthDayFormat.format(ms) + ", " + timeLabel(ms);
    }

    return {
      timeZone: tz,
      dayKey: dayKey,
      dayLabel: dayLabel,
      relativeDay: relativeDay,
      timeLabel: timeLabel,
      shortTime: shortTime,
      fullStamp: fullStamp,
      compactStamp: compactStamp,
      relativeAge: relativeAge,
      isBusinessHours: isBusinessHours,
      stampLabel: stampLabel,
    };
  }

  // ---------------------------------------------------------------------------
  // People and recipients
  // ---------------------------------------------------------------------------

  const EMAIL_PATTERN = /^[^\s@;,<>"]+@[^\s@;,<>"]+\.[^\s@;,<>"]+$/;

  function isValidEmail(value) {
    return EMAIL_PATTERN.test(String(value || "").trim());
  }

  function domainOf(email) {
    const parts = String(email || "")
      .trim()
      .toLowerCase()
      .split("@");
    return parts.length === 2 ? parts[1] : "";
  }

  /**
   * ownerEmail: the dashboard owner's address (shown as "You", never replied to).
   * internalDomains: addresses on these domains (and their subdomains) are
   * "inside the company"; anyone else is flagged External before sending.
   */
  function makePeople(ownerEmail, internalDomains) {
    const owner = String(ownerEmail || "")
      .trim()
      .toLowerCase();
    const domains = (Array.isArray(internalDomains) ? internalDomains : [])
      .map(function (d) {
        return String(d || "")
          .trim()
          .toLowerCase()
          .replace(/^@/, "");
      })
      .filter(Boolean);
    if (domains.length === 0 && domainOf(owner)) domains.push(domainOf(owner));

    function isOwner(email) {
      return !!owner && String(email || "").trim().toLowerCase() === owner;
    }

    function displayName(p) {
      return p.email && isOwner(p.email) ? "You" : p.name;
    }

    function recipientsText(list, more) {
      if (list.length === 0 && more === 0) return "You";
      const names = list.map(displayName).join(", ");
      if (more <= 0) return names;
      return names + (names ? ", " : "") + "+" + more + " more";
    }

    function addresses(list) {
      return list
        .map(function (p) {
          return p.email.trim();
        })
        .filter(function (e) {
          return e && !isOwner(e);
        });
    }

    function replyAllRecipients(item) {
      const all = [item.from.email].concat(addresses(item.to), addresses(item.cc));
      const out = [];
      const seen = new Set();
      for (const e of all) {
        if (!e) continue;
        const key = e.toLowerCase();
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(key);
      }
      return out;
    }

    /* Anyone outside the company, flagged so customers stand out before sending. */
    function isExternal(email) {
      const domain = domainOf(email);
      if (!domain || domains.length === 0) return false;
      for (const d of domains) {
        if (domain === d || domain.slice(-(d.length + 1)) === "." + d) return false;
      }
      return true;
    }

    /* Thread participants who can be added to a sender-only reply (never you or the sender). */
    function replyCandidates(item) {
      const seen = new Set([owner, item.from.email.trim().toLowerCase()]);
      const out = [];
      const lists = [
        [item.to, "to"],
        [item.cc, "cc"],
      ];
      for (const pair of lists) {
        for (const p of pair[0]) {
          const email = p.email.trim();
          const key = email.toLowerCase();
          if (!email || seen.has(key)) continue;
          seen.add(key);
          out.push({ name: p.name.trim() || email, email: email, line: pair[1] });
        }
      }
      return out;
    }

    function cleanExtras(item, extras) {
      const seen = new Set([owner, item.from.email.trim().toLowerCase()]);
      const out = [];
      for (const r of extras || []) {
        const email = String(r.email || "").trim();
        const key = email.toLowerCase();
        if (!isValidEmail(email) || seen.has(key)) continue;
        seen.add(key);
        out.push({ name: r.name, email: email, line: r.line === "cc" ? "cc" : "to" });
      }
      return out;
    }

    return {
      owner: owner,
      domains: domains,
      isOwner: isOwner,
      displayName: displayName,
      recipientsText: recipientsText,
      replyAllRecipients: replyAllRecipients,
      isExternal: isExternal,
      replyCandidates: replyCandidates,
      cleanExtras: cleanExtras,
    };
  }

  /* The owner is on nearly every email item's To/Cc line; pick the most common address. */
  function inferOwnerEmail(items) {
    const counts = new Map();
    let emailItems = 0;
    for (const item of items || []) {
      if (item.channel !== "email") continue;
      emailItems += 1;
      const seen = new Set();
      for (const p of item.to.concat(item.cc)) {
        const key = p.email.trim().toLowerCase();
        if (!key || seen.has(key)) continue;
        seen.add(key);
        counts.set(key, (counts.get(key) || 0) + 1);
      }
    }
    let best = "";
    let bestCount = 0;
    counts.forEach(function (count, key) {
      if (count > bestCount) {
        best = key;
        bestCount = count;
      }
    });
    return emailItems >= 2 && bestCount >= Math.ceil(emailItems * 0.6) ? best : "";
  }

  // ---------------------------------------------------------------------------
  // Sending helpers
  // ---------------------------------------------------------------------------

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function toHtml(text) {
    return String(text)
      .trim()
      .split(/\r?\n/)
      .map(escapeHtml)
      .join("<br>");
  }

  function viaLabel(via, channel) {
    switch (via) {
      case "replyAll":
        return " \u00b7 reply all";
      case "newEmail":
        return " \u00b7 as a new email";
      case "teams":
        return " \u00b7 in Teams";
      case "manual":
        return channel === "teams" ? " \u00b7 from Teams" : " \u00b7 from Outlook";
      default:
        return " \u00b7 reply";
    }
  }

  function sentLabel(record, channel) {
    const base = viaLabel(record.via, channel);
    const added = record.via === "reply" || record.via === "newEmail" ? record.added || [] : [];
    if (added.length === 0) return base;
    return added.length <= 2 ? base + " + " + added.join(", ") : base + " + " + added.length + " others";
  }

  // ---------------------------------------------------------------------------
  // Link safety - only http(s), mailto and tel links ever become clickable.
  // ---------------------------------------------------------------------------

  const SAFE_SCHEMES = ["http:", "https:", "mailto:", "tel:"];

  function safeUrl(url) {
    if (!url) return undefined;
    const trimmed = String(url).trim();
    if (!trimmed) return undefined;
    try {
      const base = typeof location !== "undefined" && location.origin && location.origin !== "null" ? location.origin : "https://example.invalid";
      const parsed = new URL(trimmed, base);
      if (!/^[a-z][a-z0-9+.-]*:/i.test(trimmed)) return undefined; // only absolute links from the feed
      return SAFE_SCHEMES.indexOf(parsed.protocol) >= 0 ? trimmed : undefined;
    } catch (e) {
      return undefined;
    }
  }

  const api = {
    str: str,
    parseJsonText: parseJsonText,
    normalizeItem: normalizeItem,
    normalizeRun: normalizeRun,
    normalizeFeed: normalizeFeed,
    emptyState: emptyState,
    cloneState: cloneState,
    normalizeState: normalizeState,
    MAX_TASK_LENGTH: MAX_TASK_LENGTH,
    cleanTaskText: cleanTaskText,
    newTaskId: newTaskId,
    normalizeTasks: normalizeTasks,
    resolveTimeZone: resolveTimeZone,
    makeFormat: makeFormat,
    isValidEmail: isValidEmail,
    domainOf: domainOf,
    makePeople: makePeople,
    inferOwnerEmail: inferOwnerEmail,
    escapeHtml: escapeHtml,
    toHtml: toHtml,
    viaLabel: viaLabel,
    sentLabel: sentLabel,
    safeUrl: safeUrl,
  };

  if (typeof module === "object" && module && module.exports) module.exports = api;
  root.COSCore = api;
})(typeof window !== "undefined" ? window : globalThis);
