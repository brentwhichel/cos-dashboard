/*
 * COS Communications Dashboard - setup guide.
 * Renders the "Customize once" fields and every copy-ready prompt.
 */
(function () {
  "use strict";

  const P = window.COS_PROMPTS;
  const STORE_KEY = "cos-dashboard.setup.v1";

  // ------------------------------------------------------------- helpers
  function h(tag, props) {
    const el = document.createElement(tag);
    if (props) {
      for (const key of Object.keys(props)) {
        const v = props[key];
        if (v === undefined || v === null || v === false) continue;
        if (key === "class") el.className = v;
        else if (key.slice(0, 2) === "on" && typeof v === "function") el.addEventListener(key.slice(2).toLowerCase(), v);
        else if (typeof v === "boolean") el.setAttribute(key, "");
        else el.setAttribute(key, String(v));
      }
    }
    for (let i = 2; i < arguments.length; i += 1) add(el, arguments[i]);
    return el;
  }
  function add(el, child) {
    if (child === null || child === undefined || child === false) return;
    if (Array.isArray(child)) return child.forEach((c) => add(el, c));
    el.append(child instanceof Node ? child : document.createTextNode(String(child)));
  }
  const ICON = {
    copy: '<rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>',
    check: '<path d="M20 6 9 17l-5-5"/>',
    arrow: '<path d="M5 12h14"/><path d="m12 5 7 7-7 7"/>',
    folder:
      '<path d="m6 14 1.5-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.54 6a2 2 0 0 1-1.95 1.5H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 1.67.9H18a2 2 0 0 1 2 2v2"/>',
    login: '<path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" x2="3" y1="12" y2="12"/>',
    app: '<rect width="18" height="18" x="3" y="3" rx="2"/><path d="M3 9h18"/><path d="M9 21V9"/>',
  };
  function icon(name) {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("fill", "none");
    svg.setAttribute("stroke", "currentColor");
    svg.setAttribute("stroke-width", "2");
    svg.setAttribute("stroke-linecap", "round");
    svg.setAttribute("stroke-linejoin", "round");
    svg.setAttribute("aria-hidden", "true");
    svg.setAttribute("class", "lucide");
    svg.innerHTML = ICON[name];
    return svg;
  }
  function copy(text) {
    const area = document.createElement("textarea");
    area.value = text;
    area.setAttribute("readonly", "");
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.appendChild(area);
    area.select();
    let ok = false;
    try {
      ok = document.execCommand("copy");
    } catch (e) {
      ok = false;
    }
    area.remove();
    if (!ok && navigator.clipboard) {
      navigator.clipboard.writeText(text).catch(() => {});
      ok = true;
    }
    return ok;
  }
  function toast(text) {
    const host = document.getElementById("toaster");
    const el = h("div", { class: "toast", role: "status" }, h("div", { class: "toast-title" }, text));
    host.append(el);
    setTimeout(() => el.remove(), 2500);
  }

  // ------------------------------------------------------------- fields
  const FIELDS = [
    { key: "YOUR_NAME", label: "Your name", placeholder: "Alex Morgan" },
    { key: "YOUR_EMAIL", label: "Your work email", placeholder: "alex.morgan@contoso.com", type: "email" },
    { key: "YOUR_ROLE", label: "Your role", placeholder: "Director, Customer Success" },
    { key: "YOUR_ORG", label: "Company", placeholder: "Contoso" },
    { key: "TIME_ZONE", label: "Time zone", placeholder: "America/New_York", auto: () => Intl.DateTimeFormat().resolvedOptions().timeZone || "" },
    { key: "FEED_FOLDER", label: "Dashboard folder (inside your OneDrive Cowork folder)", placeholder: "COS Comms Dashboard", auto: () => "COS Comms Dashboard" },
    { key: "CUSTOMER_NAMES", label: "Customer accounts", placeholder: "Fabrikam, Northwind Traders, Tailwind Traders" },
    { key: "CUSTOMER_DOMAINS", label: "Customer email domains", placeholder: "fabrikam.com, northwindtraders.com, tailwindtraders.com" },
    {
      key: "PRIORITY_PEOPLE",
      label: "Priority people (one per line: name, email, role)",
      placeholder: "Jordan Lee (jordan.lee@contoso.com, my manager)\nMorgan Diaz (morgan.diaz@contoso.com, my skip-level)",
      multi: true,
      join: ", ",
    },
    { key: "KEY_PEOPLE", label: "Key people for instant alerts (one per line)", placeholder: "Jordan Lee\nMorgan Diaz", multi: true },
    {
      key: "FOCUS_TOPICS",
      label: "What matters most for your accounts",
      placeholder: "licensing, adoption, executive sponsorship or expansion opportunities",
      wide: true,
    },
    { key: "SWEEP_SCHEDULE", label: "Teams sweep schedule", placeholder: "every 30 minutes on weekdays, from 7:00 AM to 7:30 PM ET" },
    {
      key: "SIGNATURE",
      label: "Draft signature line",
      placeholder: "This message was reviewed by Alex, but sent by their Cowork Chief of Staff!",
      auto: () => "This message was reviewed by [FIRST_NAME], but sent by their Cowork Chief of Staff!",
    },
    { key: "PRIORITIES", label: "Constitution: top business priorities (Step 1)", placeholder: "1. Grow adoption at our top 10 accounts\n2. ...", multi: true },
    { key: "STAKEHOLDERS", label: "Constitution: key stakeholders (Step 1)", placeholder: "My manager, my skip-level, the CIO at each customer...", multi: true },
    {
      key: "GUARDRAILS",
      label: "Constitution: guardrails (Step 1)",
      placeholder: "- Draft emails for me, but never send them.",
      multi: true,
      wide: true,
      auto: () => P.DEFAULTS.GUARDRAILS,
    },
    { key: "REPO_URL", label: "This repository's GitHub address (for the optional Cowork app)", placeholder: "https://github.com/you/cos-dashboard", auto: repoUrl, wide: true },
  ];

  /* On GitHub Pages, open docs on github.com so Markdown renders; locally, use the file. */
  function repoLink(path) {
    const base = repoUrl();
    if (!base) return path;
    const isFile = /\.[a-z0-9]+$/i.test(path);
    return base + (isFile ? "/blob/main/" : "/tree/main/") + path.replace(/\/+$/, "");
  }

  function repoUrl() {
    const host = window.location.hostname;
    if (/\.github\.io$/i.test(host)) {
      const owner = host.split(".")[0];
      const repo = window.location.pathname.split("/").filter(Boolean)[0];
      return repo ? "https://github.com/" + owner + "/" + repo : "https://github.com/" + owner + "/" + owner + ".github.io";
    }
    return "";
  }

  function loadValues() {
    let saved = {};
    try {
      saved = JSON.parse(localStorage.getItem(STORE_KEY) || "{}") || {};
    } catch (e) {
      saved = {};
    }
    const values = {};
    for (const f of FIELDS) {
      if (typeof saved[f.key] === "string") values[f.key] = saved[f.key];
      else if (f.auto) values[f.key] = f.auto();
      else values[f.key] = "";
    }
    return values;
  }
  let values = loadValues();
  function save() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify(values));
    } catch (e) {
      /* private mode */
    }
  }

  const TZ_SHORT = {
    "America/New_York": "ET",
    "America/Detroit": "ET",
    "America/Toronto": "ET",
    "America/Chicago": "CT",
    "America/Winnipeg": "CT",
    "America/Denver": "MT",
    "America/Edmonton": "MT",
    "America/Phoenix": "MST",
    "America/Los_Angeles": "PT",
    "America/Vancouver": "PT",
    "Europe/London": "UK time",
    "Europe/Dublin": "Irish time",
    "Europe/Paris": "CET",
    "Europe/Berlin": "CET",
    "Europe/Amsterdam": "CET",
    "Asia/Kolkata": "IST",
    "Asia/Singapore": "SGT",
    "Asia/Tokyo": "JST",
    "Australia/Sydney": "AET",
  };
  function tzLabel(tz) {
    if (!tz) return "";
    if (TZ_SHORT[tz]) return TZ_SHORT[tz];
    try {
      const part = new Intl.DateTimeFormat("en-US", { timeZone: tz, timeZoneName: "short" }).formatToParts(new Date()).find((p) => p.type === "timeZoneName");
      return part ? part.value : tz;
    } catch (e) {
      return tz;
    }
  }

  function tokenValues() {
    const v = {};
    for (const f of FIELDS) {
      let value = String(values[f.key] || "").trim();
      if (f.multi && f.join) {
        value = value
          .split(/\n+/)
          .map((s) => s.trim())
          .filter(Boolean)
          .join(f.join);
      }
      v[f.key] = value;
    }
    const first = (v.YOUR_NAME || "").split(/\s+/)[0];
    v.FIRST_NAME = first;
    const label = tzLabel(v.TIME_ZONE);
    v.TZ_LABEL = label;
    if (!v.SWEEP_SCHEDULE) v.SWEEP_SCHEDULE = "every 30 minutes on weekdays, from 7:00 AM to 7:30 PM " + (label || "[TZ]");
    if (v.SIGNATURE) v.SIGNATURE = v.SIGNATURE.split("[FIRST_NAME]").join(first || "[FIRST_NAME]").split("[YOUR_NAME]").join(v.YOUR_NAME || "[YOUR_NAME]");
    if (v.REPO_URL) v.REPO_URL = v.REPO_URL.replace(/\/+$/, "");
    return v;
  }

  function renderFields() {
    const form = document.getElementById("fields");
    form.replaceChildren();
    for (const f of FIELDS) {
      const id = "f-" + f.key.toLowerCase();
      const input = f.multi
        ? h("textarea", { id: id, class: "textarea", placeholder: f.placeholder, rows: "3" })
        : h("input", { id: id, class: "input", type: f.type || "text", placeholder: f.placeholder });
      input.value = values[f.key] || "";
      input.addEventListener("input", () => {
        values[f.key] = input.value;
        save();
        renderPrompts();
      });
      form.append(h("div", { class: "field" + (f.multi || f.wide ? " wide" : "") }, h("label", { for: id, class: "label" }, f.label), input));
    }
    form.addEventListener("submit", (e) => e.preventDefault());
  }

  // ------------------------------------------------------------- steps

  const promptEls = [];

  function promptBox(prompt) {
    const pre = h("pre", { class: "prompt", tabindex: "0" });
    const expand = h("button", { type: "button", class: "btn btn-ghost btn-sm" }, "Show all");
    expand.addEventListener("click", () => {
      const open = pre.classList.toggle("is-open");
      expand.textContent = open ? "Show less" : "Show all";
    });
    const copyBtn = h("button", { type: "button", class: "btn btn-default btn-sm" }, icon("copy"), "Copy prompt");
    copyBtn.addEventListener("click", () => {
      const ok = copy(pre.textContent);
      toast(ok ? "Copied. Paste it into Cowork." : "Select the prompt and copy it yourself.");
      if (ok) {
        copyBtn.replaceChildren(icon("check"), "Copied");
        setTimeout(() => copyBtn.replaceChildren(icon("copy"), "Copy prompt"), 2000);
      }
    });
    promptEls.push({ prompt: prompt, pre: pre, expand: expand });
    return h(
      "div",
      { class: "prompt-wrap" },
      h("div", { class: "prompt-bar" }, h("span", { class: "label" }, "Copy into Cowork"), h("div", { class: "prompt-bar-actions" }, expand, copyBtn)),
      pre,
    );
  }

  function afterNote(id) {
    const prompt = P.PROMPTS.find((p) => p.id === id);
    if (!prompt || !prompt.after) return null;
    const el = h("p", { class: "after", "data-after": id }, icon("arrow"), h("span", null, ""));
    return el;
  }

  function stepCard(prompt, extra) {
    return h(
      "section",
      { class: "card", id: "step-" + prompt.id, "aria-labelledby": "title-" + prompt.id },
      h("div", { class: "step-head" }, h("p", { class: "eyebrow" }, prompt.step)),
      h("h2", { class: "section-title", id: "title-" + prompt.id }, prompt.title),
      h("p", { class: "section-desc" }, prompt.why),
      extra || null,
      h("p", { class: "where" }, h("strong", null, "Where:"), h("span", null, prompt.where)),
      promptBox(prompt),
      afterNote(prompt.id),
    );
  }

  function dashboardCard() {
    const appPrompt = P.PROMPTS.find((p) => p.id === "cowork-app");
    return h(
      "section",
      { class: "card", id: "step-dashboard", "aria-labelledby": "title-dashboard" },
      h("p", { class: "eyebrow" }, "Step 3"),
      h("h2", { class: "section-title", id: "title-dashboard" }, "Open your dashboard"),
      h(
        "p",
        { class: "section-desc" },
        "The dashboard is this website's home page. It looks and works like the original Cowork app, and it shows only your own feed: it runs in your browser and talks only to your own Microsoft 365. Choose how it connects:",
      ),
      h(
        "div",
        { class: "options" },
        h(
          "div",
          { class: "option" },
          h("span", { class: "option-icon" }, icon("folder")),
          h("h3", null, "OneDrive folder (no IT needed)"),
          h(
            "p",
            null,
            "Sync your OneDrive to your computer, open the dashboard in Edge or Chrome, choose \u201cOpen my OneDrive folder\u201d and pick Documents \u203a Cowork \u203a COS Comms Dashboard. \u201cCopy & open\u201d copies each reply so you can paste it in Outlook or Teams.",
          ),
        ),
        h(
          "div",
          { class: "option" },
          h("span", { class: "option-icon" }, icon("login")),
          h("h3", null, "Microsoft 365 sign-in (full Send)"),
          h(
            "p",
            null,
            "Your IT admin registers the dashboard once (about 5 minutes) and gives you an Application ID. Add it in the dashboard's Settings, sign in, and Send works from your own Outlook and Teams.",
          ),
          h("a", { class: "link", href: repoLink("docs/microsoft-365-sign-in.md") }, "Registration steps for IT"),
        ),
        h(
          "div",
          { class: "option" },
          h("span", { class: "option-icon" }, icon("app")),
          h("h3", null, "Inside Cowork (optional)"),
          h(
            "p",
            null,
            "Prefer the original app inside Cowork? The prompt below has Cowork rebuild it from this repository. It uses your own Outlook and Teams connections, so Send works with no IT registration.",
          ),
        ),
      ),
      h(
        "div",
        { class: "hero-actions" },
        h("a", { class: "btn btn-default btn-md", href: "index.html" }, "Open my dashboard"),
        h("a", { class: "btn btn-outline btn-md", href: "index.html?demo" }, "See the demo first"),
      ),
      h("p", { class: "where" }, h("strong", null, "Optional:"), h("span", null, appPrompt.why + " " + appPrompt.where)),
      promptBox(appPrompt),
      afterNote("cowork-app"),
    );
  }

  function emailRules() {
    return h(
      "div",
      { class: "cards-2" },
      h("div", { class: "rule" }, h("h3", null, "Add to the dashboard"), "Direct asks, commitments, customer or executive issues, deadlines, decisions and risks. Each gets a priority, a summary, why it matters and a researched draft."),
      h("div", { class: "rule" }, h("h3", null, "Skip and mark read"), "Marketing, newsletters, broad announcements, automated notices and thank-you replies that ask nothing of you."),
      h("div", { class: "rule" }, h("h3", null, "Skip but leave unread"), "Anything from a customer, anything personal, and anything with an action still open, so you still see it."),
    );
  }

  function renderSteps() {
    const host = document.getElementById("steps");
    const byId = {};
    P.PROMPTS.forEach((p) => (byId[p.id] = p));
    const cards = [
      stepCard(byId.constitution),
      stepCard(byId.feed),
      dashboardCard(),
      stepCard(byId.email, emailRules()),
      stepCard(byId.customer),
      stepCard(byId.keychat),
      stepCard(byId.sweep),
      stepCard(byId.research),
      stepCard(byId.heartbeat),
      stepCard(byId.audit),
      stepCard(byId.tune),
    ];
    host.replaceChildren(h("div", { class: "guide" }, cards));
  }

  function renderPrompts() {
    const v = tokenValues();
    for (const p of promptEls) p.pre.textContent = P.fill(p.prompt.text, v);
    document.querySelectorAll("[data-after]").forEach((el) => {
      const prompt = P.PROMPTS.find((p) => p.id === el.getAttribute("data-after"));
      el.lastChild.textContent = P.fill(prompt ? prompt.after : "", v);
    });
    document.querySelectorAll("[data-token]").forEach((el) => {
      const key = el.getAttribute("data-token");
      el.textContent = v[key] || P.DEFAULTS[key] || "";
    });
  }

  document.getElementById("clear-fields").addEventListener("click", () => {
    try {
      localStorage.removeItem(STORE_KEY);
    } catch (e) {
      /* ignore */
    }
    values = loadValues();
    renderFields();
    renderPrompts();
  });

  document.querySelectorAll("a[data-repo-path]").forEach((a) => {
    a.setAttribute("href", repoLink(a.getAttribute("data-repo-path")));
  });

  renderFields();
  renderSteps();
  renderPrompts();
})();
