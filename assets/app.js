/*
 * COS Communications Dashboard - page logic.
 *
 * Reads the private feed your COS automations write to OneDrive and shows it
 * exactly like the original COS Communications Dashboard app: My tasks, the
 * latest check, and a day-by-day timeline of emails and Teams chats with a
 * researched draft reply for each one.
 *
 * Three ways to connect, all private to the person using the page:
 *   - folder: the OneDrive folder synced to this computer (no sign-in)
 *   - graph:  Microsoft 365 sign-in (reads OneDrive, sends from Outlook/Teams)
 *   - demo:   fictional sample data, nothing saved or sent
 */
(function () {
  "use strict";

  const C = window.COSCore;
  const DEMO = window.COS_DEMO;

  // ---------------------------------------------------------------------------
  // Configuration and per-browser settings
  // ---------------------------------------------------------------------------

  const DEFAULTS = {
    appTitle: "COS Communications",
    feedFolder: "Documents/Cowork/COS Comms Dashboard",
    clientId: "",
    tenantId: "organizations",
    redirectUri: "",
    internalDomains: [],
    timeZone: "",
    pollSeconds: 60,
    undoSeconds: 5,
    businessHoursStart: 8,
    businessHoursEnd: 18,
    staleHours: 2,
    showResearchNotes: true,
    msalUrl: "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4/+esm",
    setupGuideUrl: "setup.html",
  };
  const CONFIG = Object.assign({}, DEFAULTS, window.COS_DASHBOARD_CONFIG || {});

  const SETTINGS_KEY = "cos-dashboard.settings.v1";
  const FEED_NAME = "comms-feed.json";
  const STATE_NAME = "dashboard-state.json";
  const TASKS_NAME = "my-tasks.json";
  const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
  const GRAPH_SCOPES = ["User.Read", "Files.ReadWrite", "Mail.Read", "Mail.Send", "ChatMessage.Send", "ChannelMessage.Send"];

  function readSettings() {
    try {
      const parsed = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}");
      return parsed && typeof parsed === "object" ? parsed : {};
    } catch (e) {
      return {};
    }
  }
  let settings = readSettings();

  function writeSettings(patch) {
    settings = Object.assign({}, settings, patch);
    try {
      localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
    } catch (e) {
      /* private mode - keep settings for this visit only */
    }
  }

  function pick(key) {
    const v = settings[key];
    return v === undefined || v === null || v === "" ? CONFIG[key] : v;
  }

  function domainsSetting() {
    const v = pick("internalDomains");
    const list = Array.isArray(v) ? v : String(v || "").split(/[\s,;]+/);
    return list.map((d) => String(d).trim().replace(/^@/, "")).filter(Boolean);
  }

  function showResearch() {
    const v = settings.showResearchNotes;
    return typeof v === "boolean" ? v : CONFIG.showResearchNotes !== false;
  }

  function undoSeconds() {
    const n = Number(CONFIG.undoSeconds);
    return Number.isFinite(n) && n >= 0 ? Math.floor(n) : 5;
  }

  function makeFmt() {
    return C.makeFormat(pick("timeZone"), {
      start: Number(CONFIG.businessHoursStart),
      end: Number(CONFIG.businessHoursEnd),
    });
  }
  let fmt = makeFmt();

  // ---------------------------------------------------------------------------
  // Tones, labels
  // ---------------------------------------------------------------------------

  const PRIORITY = {
    high: { label: "High priority", tone: "danger" },
    medium: { label: "Medium priority", tone: "warning" },
    low: { label: "Low priority", tone: "neutral" },
  };
  const STATUS = {
    open: { label: "Open", tone: "info" },
    replied: { label: "Replied", tone: "success" },
    completed: { label: "Completed", tone: "neutral" },
  };
  const OUTCOME = {
    added: { label: "Added", tone: "success" },
    skipped: { label: "Skipped", tone: "neutral" },
    error: { label: "Needs attention", tone: "danger" },
  };
  const SOURCE_LABEL = {
    email: "Email",
    "teams-chat": "Teams chat",
    "teams-sweep": "Teams sweep",
    other: "Check",
  };

  // ---------------------------------------------------------------------------
  // Icons (Lucide, MIT)
  // ---------------------------------------------------------------------------

  const ICONS = {
    inbox:
      '<polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>',
    mail: '<rect width="20" height="16" x="2" y="4" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/>',
    "message-square-text":
      '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M13 8H7"/><path d="M17 12H7"/>',
    "circle-check": '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/>',
    "chevron-down": '<path d="m6 9 6 6 6-6"/>',
    "external-link": '<path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>',
    plus: '<path d="M5 12h14"/><path d="M12 5v14"/>',
    "refresh-cw":
      '<path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/>',
    "rotate-ccw": '<path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/>',
    send:
      '<path d="M14.536 21.686a.5.5 0 0 0 .937-.024l6.5-19a.496.496 0 0 0-.635-.635l-19 6.5a.5.5 0 0 0-.024.937l7.93 3.18a2 2 0 0 1 1.112 1.11z"/><path d="m21.854 2.147-10.94 10.939"/>',
    "trash-2":
      '<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/><line x1="10" x2="10" y1="11" y2="17"/><line x1="14" x2="14" y1="11" y2="17"/>',
    "triangle-alert":
      '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
    "user-plus":
      '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><line x1="19" x2="19" y1="8" y2="14"/><line x1="22" x2="16" y1="11" y2="11"/>',
    "loader-circle": '<path d="M21 12a9 9 0 1 1-6.219-8.56"/>',
    "octagon-x":
      '<path d="m15 9-6 6"/><path d="M2.586 16.726A2 2 0 0 1 2 15.312V8.688a2 2 0 0 1 .586-1.414l4.688-4.688A2 2 0 0 1 8.688 2h6.624a2 2 0 0 1 1.414.586l4.688 4.688A2 2 0 0 1 22 8.688v6.624a2 2 0 0 1-.586 1.414l-4.688 4.688a2 2 0 0 1-1.414.586H8.688a2 2 0 0 1-1.414-.586z"/><path d="m9 9 6 6"/>',
    info: '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/>',
    settings:
      '<path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>',
    "folder-open":
      '<path d="m6 14 1.5-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.54 6a2 2 0 0 1-1.95 1.5H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 1.67.9H18a2 2 0 0 1 2 2v2"/>',
    "log-in":
      '<path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" x2="3" y1="12" y2="12"/>',
    play: '<polygon points="6 3 20 12 6 21 6 3"/>',
    copy: '<rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>',
    search: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
    "book-open":
      '<path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/>',
  };

  function icon(name, extraClass) {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("fill", "none");
    svg.setAttribute("stroke", "currentColor");
    svg.setAttribute("stroke-width", "2");
    svg.setAttribute("stroke-linecap", "round");
    svg.setAttribute("stroke-linejoin", "round");
    svg.setAttribute("aria-hidden", "true");
    svg.setAttribute("focusable", "false");
    svg.setAttribute("class", "lucide lucide-" + name + (extraClass ? " " + extraClass : ""));
    svg.innerHTML = ICONS[name] || "";
    return svg;
  }

  function spinner(label) {
    const svg = icon("loader-circle", "spin");
    svg.removeAttribute("aria-hidden");
    svg.setAttribute("role", "status");
    svg.setAttribute("aria-label", label || "Loading");
    return svg;
  }

  // ---------------------------------------------------------------------------
  // DOM helpers - every piece of feed text goes in as text, never as HTML.
  // ---------------------------------------------------------------------------

  function h(tag, props) {
    const el = document.createElement(tag);
    if (props) {
      for (const key of Object.keys(props)) {
        const v = props[key];
        if (v === undefined || v === null || v === false) continue;
        if (key === "class") el.className = v;
        else if (key.length > 2 && key.slice(0, 2) === "on" && typeof v === "function") {
          el.addEventListener(key.slice(2).toLowerCase(), v);
        } else if (typeof v === "boolean") {
          if (key in el) el[key] = v;
          else el.setAttribute(key, "");
        } else el.setAttribute(key, String(v));
      }
    }
    for (let i = 2; i < arguments.length; i += 1) appendChild(el, arguments[i]);
    return el;
  }

  function appendChild(el, child) {
    if (child === null || child === undefined || child === false) return;
    if (Array.isArray(child)) {
      for (const c of child) appendChild(el, c);
      return;
    }
    el.append(child instanceof Node ? child : document.createTextNode(String(child)));
  }

  /* Replace an element's children, skipping null / undefined / false (never rendered as text). */
  function fill(el) {
    el.replaceChildren();
    for (let i = 1; i < arguments.length; i += 1) appendChild(el, arguments[i]);
    return el;
  }

  /* Put `nodes` into `container` in order, moving as few existing nodes as possible. */
  function reconcile(container, nodes) {
    let ref = container.firstChild;
    for (const node of nodes) {
      if (node === ref) {
        ref = ref.nextSibling;
        continue;
      }
      container.insertBefore(node, ref);
    }
    while (ref) {
      const next = ref.nextSibling;
      container.removeChild(ref);
      ref = next;
    }
  }

  function keepFocus(fn) {
    const active = document.activeElement;
    const start = active && "selectionStart" in active ? active.selectionStart : null;
    const end = active && "selectionEnd" in active ? active.selectionEnd : null;
    fn();
    if (active && active !== document.body && active.isConnected && document.activeElement !== active) {
      try {
        active.focus({ preventScroll: true });
        if (start !== null && typeof active.setSelectionRange === "function") active.setSelectionRange(start, end);
      } catch (e) {
        /* element no longer focusable */
      }
    }
  }

  let uid = 0;
  function nextId(prefix) {
    uid += 1;
    return prefix + "-" + uid;
  }

  function copyText(text) {
    let ok = false;
    const area = document.createElement("textarea");
    area.value = text;
    area.setAttribute("readonly", "");
    area.style.position = "fixed";
    area.style.opacity = "0";
    area.style.top = "0";
    document.body.appendChild(area);
    area.select();
    try {
      ok = document.execCommand("copy");
    } catch (e) {
      ok = false;
    }
    area.remove();
    if (!ok && navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(function () {});
      ok = true;
    }
    return ok;
  }

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  // ---------------------------------------------------------------------------
  // Toasts (same look and placement as the original app)
  // ---------------------------------------------------------------------------

  const toaster = document.getElementById("toaster");

  function toast(title, opts) {
    const o = opts || {};
    const el = h("div", { class: "toast", role: o.kind === "error" ? "alert" : "status" });
    if (o.kind) {
      const name = o.kind === "success" ? "circle-check" : o.kind === "error" ? "octagon-x" : "info";
      el.append(h("span", { class: "toast-icon" }, icon(name)));
    }
    el.append(
      h(
        "div",
        { class: "toast-content" },
        h("div", { class: "toast-title" }, title),
        o.description ? h("div", { class: "toast-desc" }, o.description) : null,
      ),
    );
    let closed = false;
    let timer = null;
    const close = () => {
      if (closed) return;
      closed = true;
      clearTimeout(timer);
      el.classList.add("is-leaving");
      setTimeout(() => el.remove(), 200);
    };
    if (o.action) {
      el.append(
        h(
          "button",
          {
            type: "button",
            class: "toast-action",
            onclick: () => {
              close();
              o.action.onClick();
            },
          },
          o.action.label,
        ),
      );
    }
    toaster.append(el);
    while (toaster.children.length > 3) toaster.firstChild.remove();
    const duration = o.duration || 4000;
    timer = setTimeout(close, duration);
    el.addEventListener("mouseenter", () => clearTimeout(timer));
    el.addEventListener("mouseleave", () => {
      clearTimeout(timer);
      timer = setTimeout(close, 2000);
    });
    el.addEventListener("focusin", () => clearTimeout(timer));
  }
  toast.success = (title, opts) => toast(title, Object.assign({ kind: "success" }, opts));
  toast.error = (title, opts) => toast(title, Object.assign({ kind: "error" }, opts));

  // ---------------------------------------------------------------------------
  // Errors
  // ---------------------------------------------------------------------------

  class DashError extends Error {
    constructor(kind, message) {
      super(message);
      this.kind = kind;
    }
  }

  function conflictError() {
    const e = new Error("conflict");
    e.conflict = true;
    return e;
  }

  // ---------------------------------------------------------------------------
  // IndexedDB - remembers which OneDrive folder this browser may open.
  // ---------------------------------------------------------------------------

  function idbOpen() {
    return new Promise((resolve, reject) => {
      if (!("indexedDB" in window)) return reject(new Error("no indexedDB"));
      const req = indexedDB.open("cos-dashboard", 1);
      req.onupgradeneeded = () => req.result.createObjectStore("kv");
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function idbGet(key) {
    const db = await idbOpen();
    return new Promise((resolve, reject) => {
      const req = db.transaction("kv", "readonly").objectStore("kv").get(key);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function idbSet(key, value) {
    const db = await idbOpen();
    return new Promise((resolve, reject) => {
      const tx = db.transaction("kv", "readwrite");
      tx.objectStore("kv").put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }
  async function idbDel(key) {
    const db = await idbOpen();
    return new Promise((resolve) => {
      const tx = db.transaction("kv", "readwrite");
      tx.objectStore("kv").delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    });
  }

  // ---------------------------------------------------------------------------
  // Backends. Each one reads the feed, reads/writes the dashboard's own two
  // files, and (where it can) sends replies:
  //   readFeed(prevVersion) -> { json, version } | { unchanged, version }
  //   readFile(name)        -> { json, version } | null when the file is missing
  //   writeFile(name, value, expectedVersion) -> newVersion (throws conflict)
  // ---------------------------------------------------------------------------

  function emailHtml(text) {
    return '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' + C.toHtml(text) + "</div>";
  }

  class DemoBackend {
    constructor() {
      this.kind = "demo";
      this.sendMode = "send";
      this.label = "Demo data (fictional)";
      this.data = DEMO.build(Date.now());
    }
    async ownerEmail() {
      return DEMO.owner.email;
    }
    async readFeed(prev) {
      await sleep(250);
      if (prev === "demo") return { unchanged: true, version: "demo" };
      return { json: this.data.feed, version: "demo" };
    }
    async readFile(name) {
      const raw = sessionStorage.getItem("cos-demo:" + name);
      if (raw) return { json: JSON.parse(raw), version: "v" + raw.length };
      if (name === STATE_NAME) return { json: this.data.state, version: "seed" };
      if (name === TASKS_NAME) return { json: this.data.tasks, version: "seed" };
      return null;
    }
    async writeFile(name, value) {
      const text = JSON.stringify(value, null, 1);
      sessionStorage.setItem("cos-demo:" + name, text);
      return "v" + text.length;
    }
    async sendEmailReply(item, text, replyAll) {
      await sleep(700);
      return replyAll ? "replyAll" : "reply";
    }
    async sendEmailAsNew() {
      await sleep(700);
      return "newEmail";
    }
    async sendTeamsReply() {
      await sleep(700);
      return "teams";
    }
    forget() {
      sessionStorage.removeItem("cos-demo:" + STATE_NAME);
      sessionStorage.removeItem("cos-demo:" + TASKS_NAME);
    }
  }

  class FolderBackend {
    constructor(dir) {
      this.kind = "folder";
      this.sendMode = "copy";
      this.dir = dir;
      this.label = "OneDrive folder \u201c" + dir.name + "\u201d on this computer";
    }
    async ownerEmail() {
      return "";
    }
    async handle(name, create) {
      try {
        return await this.dir.getFileHandle(name, { create: !!create });
      } catch (e) {
        if (e && (e.name === "NotFoundError" || e.name === "TypeMismatchError")) return null;
        throw e;
      }
    }
    async readFeed(prev) {
      const fh = await this.handle(FEED_NAME);
      if (!fh) throw new DashError("feed", "feed-missing");
      const file = await fh.getFile();
      const version = file.lastModified + ":" + file.size;
      if (prev && prev === version) return { unchanged: true, version: version };
      return { json: C.parseJsonText(await file.text()), version: version };
    }
    async readFile(name) {
      const fh = await this.handle(name);
      if (!fh) return null;
      const file = await fh.getFile();
      const text = await file.text();
      return { json: text.trim() ? C.parseJsonText(text) : null, version: file.lastModified + ":" + file.size };
    }
    async writeFile(name, value, expected) {
      const fh = await this.handle(name, true);
      if (expected) {
        const current = await fh.getFile();
        if (current.lastModified + ":" + current.size !== expected) throw conflictError();
      }
      const writable = await fh.createWritable();
      await writable.write(JSON.stringify(value, null, 1));
      await writable.close();
      const saved = await fh.getFile();
      return saved.lastModified + ":" + saved.size;
    }
  }

  // ----- Microsoft 365 (MSAL + Microsoft Graph) -----

  let msalLib = null;
  let msalApp = null;

  function redirectUri() {
    if (CONFIG.redirectUri) return CONFIG.redirectUri;
    return window.location.origin + window.location.pathname.replace(/index\.html$/, "");
  }

  async function getMsalApp() {
    if (msalApp) return msalApp;
    const clientId = String(pick("clientId") || "").trim();
    if (!clientId) throw new DashError("auth", "no-client-id");
    const mod = await import(CONFIG.msalUrl);
    const lib = mod && mod.PublicClientApplication ? mod : mod && mod.default ? mod.default : mod;
    const tenant = String(pick("tenantId") || "organizations").trim() || "organizations";
    const app = new lib.PublicClientApplication({
      auth: {
        clientId: clientId,
        authority: "https://login.microsoftonline.com/" + encodeURIComponent(tenant),
        redirectUri: redirectUri(),
      },
      cache: { cacheLocation: "localStorage" },
    });
    await app.initialize();
    msalLib = lib;
    msalApp = app;
    return app;
  }

  function needsInteraction(error) {
    if (!error) return false;
    if (msalLib && msalLib.InteractionRequiredAuthError && error instanceof msalLib.InteractionRequiredAuthError) return true;
    const code = String(error.errorCode || error.name || "");
    return /interaction_required|login_required|consent_required|no_tokens_found|InteractionRequired/i.test(code);
  }

  function encodePath(path) {
    return String(path || "")
      .split("/")
      .filter(Boolean)
      .map(encodeURIComponent)
      .join("/");
  }

  class GraphBackend {
    constructor(app, account) {
      this.kind = "graph";
      this.sendMode = "send";
      this.app = app;
      this.account = account;
      this.label = "Microsoft 365 as " + (account.username || account.name || "you");
    }
    async token(force) {
      try {
        const result = await this.app.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: this.account, forceRefresh: !!force });
        return result.accessToken;
      } catch (e) {
        if (needsInteraction(e)) {
          await this.app.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account: this.account });
          return new Promise(function () {}); // the page is navigating to sign-in
        }
        throw e;
      }
    }
    async call(path, opts, attempt) {
      const n = attempt || 0;
      const o = opts || {};
      const token = await this.token(n > 0 && o._retry401);
      const headers = Object.assign({ Authorization: "Bearer " + token }, o.headers || {});
      const res = await fetch(path.indexOf("http") === 0 ? path : GRAPH_ROOT + path, {
        method: o.method || "GET",
        headers: headers,
        body: o.body,
        cache: "no-store",
      });
      if ((res.status === 429 || res.status === 503 || res.status === 504) && n < 3) {
        const wait = Number(res.headers.get("Retry-After")) || Math.pow(2, n + 1);
        await sleep(Math.min(wait, 30) * 1000);
        return this.call(path, o, n + 1);
      }
      if (res.status === 401 && n === 0) return this.call(path, Object.assign({}, o, { _retry401: true }), 1);
      return res;
    }
    folderPath() {
      return encodePath(pick("feedFolder"));
    }
    itemPath(name) {
      return "/me/drive/root:/" + this.folderPath() + "/" + encodeURIComponent(name);
    }
    async ownerEmail() {
      const res = await this.call("/me?$select=mail,userPrincipalName");
      if (!res.ok) return "";
      const me = await res.json();
      return String(me.mail || me.userPrincipalName || "");
    }
    async meta(name, kind) {
      const res = await this.call(this.itemPath(name));
      if (res.status === 404) return null;
      if (!res.ok) throw new DashError(kind, "read-failed-" + res.status);
      return res.json();
    }
    async download(meta, name, kind) {
      const url = meta["@microsoft.graph.downloadUrl"];
      const res = url ? await fetch(url, { cache: "no-store" }) : await this.call(this.itemPath(name) + ":/content");
      if (!res.ok) throw new DashError(kind, "download-failed-" + res.status);
      const text = await res.text();
      return text.trim() ? C.parseJsonText(text) : null;
    }
    async readFeed(prev) {
      const meta = await this.meta(FEED_NAME, "feed");
      if (!meta) throw new DashError("feed", "feed-missing");
      const version = meta.eTag || meta.lastModifiedDateTime || "";
      if (prev && version && prev === version) return { unchanged: true, version: version };
      return { json: await this.download(meta, FEED_NAME, "feed"), version: version };
    }
    async readFile(name) {
      const meta = await this.meta(name, "state");
      if (!meta) return null;
      return { json: await this.download(meta, name, "state"), version: meta.eTag || "" };
    }
    async writeFile(name, value, expected) {
      const headers = { "Content-Type": "application/json" };
      let url = this.itemPath(name) + ":/content";
      if (expected) headers["If-Match"] = expected;
      else url += "?@microsoft.graph.conflictBehavior=fail";
      const res = await this.call(url, { method: "PUT", headers: headers, body: JSON.stringify(value, null, 1) });
      if (res.status === 412 || res.status === 409) throw conflictError();
      if (!res.ok) throw new DashError("save", "save-failed-" + res.status);
      const saved = await res.json();
      return saved.eTag || "";
    }
    async resolveMessageId(item) {
      const imid = item.reply.internetMessageId;
      if (imid) {
        const filter = "internetMessageId eq '" + imid.replace(/'/g, "''") + "'";
        const res = await this.call("/me/messages?$filter=" + encodeURIComponent(filter) + "&$select=id&$top=1");
        if (res.ok) {
          const data = await res.json();
          if (data && Array.isArray(data.value) && data.value[0] && data.value[0].id) return data.value[0].id;
        }
      }
      return item.reply.messageId || null;
    }
    async sendEmailReply(item, text, replyAll, extras) {
      const id = await this.resolveMessageId(item);
      if (!id) throw new DashError("not-found", "no-message-id");
      const added = replyAll ? [] : extras || [];
      const body = { comment: emailHtml(text) };
      if (added.length > 0) {
        const rcpt = (p) => ({ emailAddress: { address: p.email, name: p.name } });
        body.message = {
          toRecipients: [rcpt(item.from)].concat(added.filter((p) => p.line === "to").map(rcpt)),
          ccRecipients: added.filter((p) => p.line === "cc").map(rcpt),
        };
      }
      const res = await this.call("/me/messages/" + encodeURIComponent(id) + (replyAll ? "/replyAll" : "/reply"), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      if (res.status === 404) throw new DashError("not-found", "original-missing");
      if (res.status === 400) {
        const detail = await res.text().catch(() => "");
        if (/ErrorInvalidIdMalformed|ErrorItemNotFound|ResourceNotFound/i.test(detail)) {
          throw new DashError("not-found", "original-missing");
        }
      }
      if (!res.ok) throw new DashError("send", "send-failed-" + res.status);
      return replyAll ? "replyAll" : "reply";
    }
    async sendEmailAsNew(item, text, replyAll, extras) {
      const to = item.from.email.trim();
      if (!to) throw new DashError("send", "no-recipient");
      const added = replyAll ? [] : extras || [];
      const toList = [to].concat(added.filter((p) => p.line === "to").map((p) => p.email));
      const cc = replyAll
        ? S.people.replyAllRecipients(item).filter((e) => e !== to.toLowerCase())
        : added.filter((p) => p.line === "cc").map((p) => p.email);
      const subject = /^re:/i.test(item.subject) ? item.subject : "RE: " + item.subject;
      const rcpt = (address) => ({ emailAddress: { address: address } });
      const res = await this.call("/me/sendMail", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          message: {
            subject: subject,
            body: { contentType: "HTML", content: emailHtml(text) },
            toRecipients: toList.map(rcpt),
            ccRecipients: cc.map(rcpt),
          },
          saveToSentItems: true,
        }),
      });
      if (!res.ok) throw new DashError("send", "send-failed-" + res.status);
      return "newEmail";
    }
    async sendTeamsReply(item, text) {
      const r = item.reply;
      const payload = JSON.stringify({ body: { contentType: "html", content: C.toHtml(text) } });
      let path = null;
      if (r.teamId && r.channelId && r.parentMessageId) {
        path =
          "/teams/" +
          encodeURIComponent(r.teamId) +
          "/channels/" +
          encodeURIComponent(r.channelId) +
          "/messages/" +
          encodeURIComponent(r.parentMessageId) +
          "/replies";
      } else if (r.chatId) {
        path = "/chats/" + encodeURIComponent(r.chatId) + "/messages";
      }
      if (!path) throw new DashError("not-found", "no-chat");
      const res = await this.call(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: payload });
      if (!res.ok) throw new DashError("send", "send-failed-" + res.status);
      return "teams";
    }
  }

  // ---------------------------------------------------------------------------
  // Application state
  // ---------------------------------------------------------------------------

  function freshQuery() {
    return { data: null, error: null, pending: true, fetching: false, version: null, updatedAt: 0, gen: 0 };
  }

  const S = {
    backend: null,
    feed: freshQuery(),
    saved: freshQuery(),
    tasks: freshQuery(),
    filter: "all",
    localSent: {},
    pendingWrites: 0,
    pendingTaskWrites: 0,
    writeChain: Promise.resolve(),
    taskChain: Promise.resolve(),
    ownerEmail: "",
    inferredOwner: "",
    people: C.makePeople("", domainsSetting()),
    connectError: null,
    reconnectHandle: null,
    timers: [],
  };

  function refreshPeople() {
    if (S.backend && S.backend.kind === "demo") {
      S.people = C.makePeople(DEMO.owner.email, DEMO.internalDomains);
      return;
    }
    const owner = String(settings.ownerEmail || S.ownerEmail || S.inferredOwner || "").trim();
    S.people = C.makePeople(owner, domainsSetting());
  }

  function stateData() {
    return S.saved.data || C.emptyState();
  }

  function statusOf(item) {
    const st = stateData();
    if (st.completed[item.id]) return "completed";
    if (st.sent[item.id] || S.localSent[item.id]) return "replied";
    return "open";
  }

  // ----- reads -----

  async function loadFeed(force) {
    const q = S.feed;
    if (!S.backend || q.fetching) return;
    const backend = S.backend;
    q.fetching = true;
    scheduleRender();
    try {
      const res = await backend.readFeed(force ? null : q.version);
      if (backend !== S.backend) return;
      if (!res.unchanged) {
        try {
          q.data = C.normalizeFeed(res.json);
        } catch (e) {
          throw new DashError("feed", "feed-unreadable");
        }
      }
      q.version = res.version;
      q.error = null;
      q.updatedAt = Date.now();
      if (!settings.ownerEmail && !S.ownerEmail && q.data) {
        const inferred = C.inferOwnerEmail(q.data.items);
        if (inferred !== S.inferredOwner) {
          S.inferredOwner = inferred;
          refreshPeople();
        }
      }
    } catch (e) {
      if (backend !== S.backend) return;
      console.error("Feed read failed", e);
      q.error = e;
      if (!q.data) q.version = null;
    } finally {
      if (backend === S.backend) {
        q.fetching = false;
        q.pending = false;
        scheduleRender();
      }
    }
  }

  async function loadSaved() {
    const q = S.saved;
    if (!S.backend || q.fetching) return;
    const backend = S.backend;
    const gen = q.gen;
    q.fetching = true;
    scheduleRender();
    try {
      const res = await backend.readFile(STATE_NAME);
      if (backend !== S.backend) return;
      if (S.pendingWrites === 0 && gen === q.gen) q.data = C.normalizeState(res ? res.json : null);
      q.error = null;
    } catch (e) {
      if (backend !== S.backend) return;
      console.error("State read failed", e);
      q.error = e;
    } finally {
      if (backend === S.backend) {
        q.fetching = false;
        q.pending = false;
        scheduleRender();
      }
    }
  }

  async function loadTasks() {
    const q = S.tasks;
    if (!S.backend || q.fetching) return;
    const backend = S.backend;
    const gen = q.gen;
    q.fetching = true;
    scheduleRender();
    try {
      const res = await backend.readFile(TASKS_NAME);
      if (backend !== S.backend) return;
      if (S.pendingTaskWrites === 0 && gen === q.gen) q.data = C.normalizeTasks(res ? res.json : null);
      q.error = null;
    } catch (e) {
      if (backend !== S.backend) return;
      console.error("Tasks read failed", e);
      q.error = e;
    } finally {
      if (backend === S.backend) {
        q.fetching = false;
        q.pending = false;
        scheduleRender();
      }
    }
  }

  function refreshAll() {
    loadFeed(true);
    if (S.pendingWrites === 0) loadSaved();
    if (S.pendingTaskWrites === 0) loadTasks();
  }

  // ----- writes: optimistic change, then a serialized read-latest / apply / save -----

  async function mutateFile(name, normalize, change, wrap) {
    for (let attempt = 0; ; attempt += 1) {
      const latest = await S.backend.readFile(name);
      const next = change(normalize(latest ? latest.json : null));
      try {
        const version = await S.backend.writeFile(name, wrap(next), latest ? latest.version : null);
        return { value: next, version: version };
      } catch (e) {
        if (e && e.conflict && attempt < 2) continue;
        throw e;
      }
    }
  }

  function applyChange(change, failMessage) {
    const fn = (s) => {
      const next = C.cloneState(s);
      change(next);
      return next;
    };
    S.saved.data = fn(stateData());
    S.saved.gen += 1;
    S.pendingWrites += 1;
    scheduleRender();
    const backend = S.backend;
    const run = S.writeChain.then(() => mutateFile(STATE_NAME, C.normalizeState, fn, (v) => v));
    S.writeChain = run.catch(() => undefined);
    return run
      .then((result) => {
        if (backend !== S.backend) return false;
        S.pendingWrites -= 1;
        if (S.pendingWrites === 0) S.saved.data = result.value;
        scheduleRender();
        return true;
      })
      .catch((e) => {
        if (backend !== S.backend) return false;
        console.error("Dashboard save failed", e);
        S.pendingWrites -= 1;
        toast.error(failMessage);
        loadSaved();
        return false;
      });
  }

  function applyTaskChange(change, failMessage) {
    S.tasks.data = change((S.tasks.data || []).map((t) => Object.assign({}, t)));
    S.tasks.gen += 1;
    S.pendingTaskWrites += 1;
    scheduleRender();
    const backend = S.backend;
    const run = S.taskChain.then(() =>
      mutateFile(TASKS_NAME, C.normalizeTasks, (list) => change(list.map((t) => Object.assign({}, t))), (list) => ({
        version: 1,
        tasks: list,
      })),
    );
    S.taskChain = run.catch(() => undefined);
    return run
      .then((result) => {
        if (backend !== S.backend) return false;
        S.pendingTaskWrites -= 1;
        if (S.pendingTaskWrites === 0) S.tasks.data = result.value;
        scheduleRender();
        return true;
      })
      .catch((e) => {
        if (backend !== S.backend) return false;
        console.error("Task save failed", e);
        S.pendingTaskWrites -= 1;
        toast.error(failMessage);
        loadTasks();
        return false;
      });
  }

  const actions = {
    markComplete(item) {
      applyChange((s) => {
        s.completed[item.id] = new Date().toISOString();
      }, "Couldn't mark that complete. Try again.");
    },
    reopen(item) {
      applyChange((s) => {
        delete s.completed[item.id];
      }, "Couldn't reopen that item. Try again.");
    },
    remove(item) {
      applyChange((s) => {
        s.deleted[item.id] = new Date().toISOString();
      }, "Couldn't delete that item. Try again.");
      toast("Removed from your dashboard", {
        description: item.subject,
        action: {
          label: "Undo",
          onClick: () =>
            applyChange((s) => {
              delete s.deleted[item.id];
            }, "Couldn't restore that item. Try again."),
        },
      });
    },
    saveDraft(item, text) {
      applyChange((s) => {
        s.drafts[item.id] = text;
      }, "Couldn't save your edit. Copy it before leaving.");
    },
    recordSent(item, record) {
      S.localSent[item.id] = record;
      applyChange((s) => {
        s.sent[item.id] = record;
      }, "Reply sent, but its status couldn't be saved.");
    },
    addTask(text) {
      const clean = C.cleanTaskText(text);
      if (!clean) return;
      const task = { id: C.newTaskId(), text: clean, createdAt: new Date().toISOString() };
      applyTaskChange(
        (list) => (list.some((t) => t.id === task.id) ? list : [task].concat(list)),
        "Couldn't save that task. Try again.",
      );
    },
    toggleTask(task, done) {
      const at = new Date().toISOString();
      applyTaskChange(
        (list) =>
          list.map((t) => {
            if (t.id !== task.id) return t;
            return done ? Object.assign({}, t, { completedAt: t.completedAt || at }) : { id: t.id, text: t.text, createdAt: t.createdAt };
          }),
        done ? "Couldn't mark that task complete. Try again." : "Couldn't reopen that task. Try again.",
      );
    },
    deleteTask(task) {
      const index = Math.max(0, (S.tasks.data || []).findIndex((t) => t.id === task.id));
      applyTaskChange((list) => list.filter((t) => t.id !== task.id), "Couldn't delete that task. Try again.");
      toast("Task deleted", {
        description: task.text,
        action: {
          label: "Undo",
          onClick: () =>
            applyTaskChange(
              (list) => (list.some((t) => t.id === task.id) ? list : list.slice(0, index).concat([task], list.slice(index))),
              "Couldn't restore that task. Try again.",
            ),
        },
      });
    },
  };

  // ---------------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------------

  const mainEl = document.getElementById("main");
  const headerSlot = document.getElementById("headerSlot");
  let view = null;
  let renderQueued = false;

  function scheduleRender() {
    if (renderQueued) return;
    renderQueued = true;
    const run = () => {
      renderQueued = false;
      render();
    };
    if (document.hidden) setTimeout(run, 0);
    else requestAnimationFrame(run);
  }

  function render() {
    if (!S.backend) {
      if (!view || view.kind !== "connect") {
        view = ConnectView();
        fill(mainEl, view.root);
      }
    } else if (!view || view.kind !== "dashboard") {
      view = DashboardView();
      fill(mainEl, view.root);
    }
    keepFocus(() => view.update());
    renderHeader();
  }

  // ----- header: Last refreshed stamp, demo badge, settings -----

  const header = (function () {
    const dot = h("span", { class: "stamp-dot", "aria-hidden": "true" });
    const compact = h("span", { class: "stamp-compact" });
    const full = h("span", { class: "stamp-full" });
    const time = h("time", { class: "stamp-time" }, compact, full);
    const none = h("span", { class: "stamp-time" }, "Not recorded yet");
    const stamp = h(
      "div",
      { class: "stamp" },
      dot,
      h("p", null, h("span", { class: "stamp-label" }, "Last refreshed"), time, none),
    );
    const demo = h("span", { class: "badge tone-warning header-demo", title: "Sample data from a fictional company" }, "Demo data");
    const gear = h(
      "button",
      { type: "button", class: "btn btn-ghost btn-icon-sm", "aria-label": "Dashboard settings", title: "Settings", onclick: () => openSettings() },
      icon("settings"),
    );
    headerSlot.append(stamp, demo, gear);
    return { stamp: stamp, dot: dot, time: time, compact: compact, full: full, none: none, demo: demo };
  })();

  function staleNow() {
    const lastRun = S.feed.data ? S.feed.data.lastRun : null;
    const now = Date.now();
    const staleMs = Math.max(0.25, Number(CONFIG.staleHours) || 2) * 60 * 60 * 1000;
    return lastRun !== null && fmt.isBusinessHours(now) && now - lastRun.atMs > staleMs;
  }

  function freshnessTone(lastRun, stale) {
    if (!lastRun) return "neutral";
    return stale ? "warning" : "success";
  }

  function renderHeader() {
    const data = S.backend ? S.feed.data : null;
    header.stamp.hidden = !data;
    header.demo.hidden = !(S.backend && S.backend.kind === "demo");
    if (!data) return;
    const lastRun = data.lastRun;
    const stale = staleNow();
    header.dot.className = "stamp-dot dot-" + freshnessTone(lastRun, stale);
    header.time.hidden = !lastRun;
    header.none.hidden = !!lastRun;
    if (lastRun) {
      header.time.dateTime = lastRun.at;
      header.compact.textContent = fmt.compactStamp(lastRun.atMs);
      header.full.textContent = fmt.fullStamp(lastRun.atMs);
    }
  }

  // ----- connect screen -----

  function ConnectView() {
    const folderOk = typeof window.showDirectoryPicker === "function";
    const errorHost = h("div");
    const folderBtn = h("button", { type: "button", class: "btn btn-default btn-md", onclick: () => chooseFolder() });
    const signInBtn = h("button", { type: "button", class: "btn btn-default btn-md", onclick: () => signIn() }, icon("log-in"), "Sign in");
    const signInNote = h("p", { class: "option-note" });
    const settingsLink = h(
      "button",
      { type: "button", class: "btn btn-outline btn-sm", onclick: () => openSettings() },
      icon("settings"),
      "Add app ID",
    );
    const demoBtn = h("button", { type: "button", class: "btn btn-outline btn-md", onclick: () => startDemo(true) }, icon("play"), "Open demo");

    const root = h(
      "div",
      { class: "page" },
      h(
        "section",
        { class: "card connect", "aria-labelledby": "connect-heading" },
        h("h2", { id: "connect-heading", class: "h-lg" }, "Connect your COS dashboard"),
        h(
          "p",
          { class: "connect-intro" },
          "Your COS automations in Cowork write a private feed of the emails and Teams chats that need you to your OneDrive. Choose how this page reads it. Your data goes only between this browser and your own Microsoft 365; nothing is stored on this website.",
        ),
        h(
          "div",
          { class: "connect-grid" },
          h(
            "div",
            { class: "option" },
            h("span", { class: "option-icon" }, icon("folder-open")),
            h("h3", null, "Open my OneDrive folder"),
            h(
              "p",
              null,
              "No sign-in or IT setup. Pick the synced Documents \u203a Cowork \u203a COS Comms Dashboard folder on this computer. Replies are copied so you can paste and send them in Outlook or Teams.",
            ),
            folderBtn,
            folderOk ? null : h("p", { class: "option-note" }, "Needs Microsoft Edge or Google Chrome on a computer."),
          ),
          h(
            "div",
            { class: "option" },
            h("span", { class: "option-icon" }, icon("log-in")),
            h("h3", null, "Sign in with Microsoft 365"),
            h(
              "p",
              null,
              "The full experience: reads your feed straight from OneDrive and sends replies from your own Outlook and Teams after a 5-second undo.",
            ),
            h("div", { style: "display:flex;flex-wrap:wrap;gap:0.5rem;align-items:center" }, signInBtn, settingsLink),
            signInNote,
          ),
          h(
            "div",
            { class: "option" },
            h("span", { class: "option-icon" }, icon("play")),
            h("h3", null, "Explore the demo"),
            h("p", null, "See exactly how the dashboard works with sample data from a fictional company. Nothing is saved or sent."),
            demoBtn,
          ),
        ),
        errorHost,
        h(
          "p",
          { class: "connect-foot" },
          "New here? Turn on your COS automations first with the ",
          h("a", { href: CONFIG.setupGuideUrl }, "setup guide"),
          ".",
        ),
      ),
    );

    function update() {
      const reconnect = S.reconnectHandle;
      fill(folderBtn, icon("folder-open"), reconnect ? "Reconnect \u201c" + reconnect.name + "\u201d" : "Choose folder");
      folderBtn.disabled = !folderOk;
      const hasClient = !!String(pick("clientId") || "").trim();
      signInBtn.disabled = !hasClient;
      settingsLink.hidden = hasClient;
      signInNote.textContent = hasClient
        ? ""
        : "Needs a one-time app registration from your IT admin. Add its Application (client) ID in Settings.";
      signInNote.hidden = hasClient;
      fill(errorHost, 
        S.connectError
          ? h("div", { class: "notice items-start connect-error", role: "alert" }, icon("triangle-alert"), h("span", null, S.connectError))
          : "",
      );
    }
    return { kind: "connect", root: root, update: update };
  }

  // ----- dashboard -----

  function DashboardView() {
    const tasksPanel = TasksPanel();
    const checkLog = CheckLog();
    const toolbar = Toolbar();
    const savedAlert = SavedAlert();
    const feedView = FeedView();
    const root = h("div", { class: "page" }, tasksPanel.root, checkLog.root, toolbar.root, savedAlert.root, feedView.root);
    function update() {
      tasksPanel.update();
      checkLog.update();
      const counts = feedView.update();
      toolbar.update(counts);
      savedAlert.update();
    }
    return { kind: "dashboard", root: root, update: update };
  }

  function TasksPanel() {
    let completedOpen = false;
    let sig = null;
    let ready = false;
    const input = h("input", { id: "new-task", class: "input", placeholder: "Add a task", maxlength: "500", autocomplete: "off" });
    const addBtn = h("button", { type: "submit", class: "btn btn-default btn-md" }, icon("plus"), "Add");
    const form = h(
      "form",
      { class: "tasks-form" },
      h("label", { for: "new-task", class: "sr-only" }, "New task"),
      input,
      addBtn,
    );
    const count = h("span", { class: "text-sm tabular-nums muted" });
    const body = h("div");
    const root = h(
      "section",
      { class: "card tasks-panel", "aria-labelledby": "tasks-heading" },
      h("div", { class: "row-between" }, h("h2", { id: "tasks-heading", class: "h-lg" }, "My tasks"), count),
      form,
      body,
    );

    function syncForm() {
      input.disabled = !ready;
      addBtn.disabled = !ready || !C.cleanTaskText(input.value);
    }
    input.addEventListener("input", syncForm);
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const clean = C.cleanTaskText(input.value);
      if (!clean || !ready) return;
      actions.addTask(clean);
      input.value = "";
      syncForm();
      input.focus();
    });

    function row(task) {
      const done = Boolean(task.completedAt);
      const boxId = "task-" + task.id.replace(/[^a-zA-Z0-9_-]/g, "");
      const stamp = done
        ? "Completed " + fmt.stampLabel(task.completedAt || "")
        : task.createdAt
          ? "Added " + fmt.stampLabel(task.createdAt)
          : "";
      return h(
        "li",
        { class: "task-row" },
        h("input", {
          type: "checkbox",
          class: "checkbox",
          id: boxId,
          checked: done,
          onchange: (e) => actions.toggleTask(task, e.target.checked),
        }),
        h(
          "div",
          { class: "task-body" },
          h("label", { for: boxId, class: "task-text" + (done ? " is-done" : "") }, task.text),
          stamp ? h("p", { class: "task-stamp" }, stamp) : null,
        ),
        h(
          "button",
          {
            type: "button",
            class: "btn btn-ghost btn-icon-sm btn-danger-text",
            "aria-label": "Delete task: " + task.text,
            title: "Delete task",
            onclick: () => actions.deleteTask(task),
          },
          icon("trash-2"),
        ),
      );
    }

    function update() {
      const q = S.tasks;
      ready = Array.isArray(q.data);
      syncForm();
      const list = q.data || [];
      const nextSig = JSON.stringify([ready, q.pending, !!q.error, completedOpen, list, fmt.timeZone, fmt.dayKey(Date.now())]);
      if (nextSig === sig) return;
      sig = nextSig;
      const open = list.filter((t) => !t.completedAt);
      const done = list
        .filter((t) => t.completedAt)
        .sort((a, b) => (Date.parse(b.completedAt || "") || 0) - (Date.parse(a.completedAt || "") || 0));
      count.textContent = list.length > 0 ? open.length + " open" + (done.length > 0 ? " \u00b7 " + done.length + " completed" : "") : "";
      count.hidden = list.length === 0;

      if (q.pending && !ready) {
        fill(body, 
          h(
            "div",
            { role: "status", "aria-label": "Loading tasks", class: "tasks-loading" },
            h("div", { class: "skeleton", style: "height:1.25rem;width:66%" }),
            h("div", { class: "skeleton", style: "height:1.25rem;width:50%" }),
          ),
        );
        return;
      }
      if (q.error && !ready) {
        fill(body, 
          h(
            "div",
            { role: "alert", class: "tasks-failed" },
            icon("triangle-alert"),
            h("span", null, "Couldn't load your tasks."),
            h("button", { type: "button", class: "btn btn-outline btn-sm", onclick: () => loadTasks() }, "Retry"),
          ),
        );
        return;
      }
      if (list.length === 0) {
        fill(body, h("p", { class: "tasks-empty" }, "No tasks yet."));
        return;
      }
      const nodes = [];
      if (open.length > 0) nodes.push(h("ul", { "aria-label": "Open tasks", class: "tasks-list divide-y" }, open.map(row)));
      else nodes.push(h("p", { class: "tasks-empty" }, "All tasks completed."));
      if (done.length > 0) {
        const doneList = h("ul", { id: "completed-tasks", "aria-label": "Completed tasks", class: "divide-y", hidden: !completedOpen }, done.map(row));
        const trigger = h(
          "button",
          {
            type: "button",
            class: "btn btn-ghost btn-sm -ml-2",
            "aria-expanded": String(completedOpen),
            "aria-controls": "completed-tasks",
            onclick: () => {
              completedOpen = !completedOpen;
              trigger.setAttribute("aria-expanded", String(completedOpen));
              doneList.hidden = !completedOpen;
              sig = null;
            },
          },
          "Completed",
          h("span", { class: "tabular-nums" }, String(done.length)),
          icon("chevron-down", "chevron"),
        );
        nodes.push(h("div", { class: "tasks-done" }, trigger, doneList));
      }
      fill(body, nodes);
    }
    return { root: root, update: update };
  }

  function RunLine(run) {
    const outcome = OUTCOME[run.outcome];
    const detail = [run.from, run.note].filter(Boolean).join(" \u00b7 ");
    return [
      h(
        "div",
        { class: "run-line" },
        h("span", { class: "badge tone-" + outcome.tone }, outcome.label),
        h("span", { class: "muted" }, SOURCE_LABEL[run.source]),
        h("span", { class: "run-subject" }, run.subject || "(no subject)"),
      ),
      detail ? h("p", { class: "run-detail" }, detail) : null,
    ];
  }

  function CheckLog() {
    let open = false;
    let lineSig = null;
    let listSig = null;
    const dot = h("span", { class: "checklog-dot", "aria-hidden": "true" });
    const meta = h("span");
    const title = h("h2", { id: "latest-check", class: "checklog-title" }, "Latest check", meta);
    const line = h("div");
    const countEl = h("span", { class: "tabular-nums" });
    const trigger = h(
      "button",
      {
        type: "button",
        class: "btn btn-outline btn-sm",
        "aria-controls": "recent-checks",
        "aria-expanded": "false",
        onclick: () => {
          open = !open;
          trigger.setAttribute("aria-expanded", String(open));
          list.hidden = !open;
        },
      },
      "Recent checks",
      countEl,
      icon("chevron-down", "chevron"),
    );
    const staleText = h("span");
    const staleBox = h("div", { role: "status", class: "notice items-start" }, icon("triangle-alert"), staleText);
    const list = h("ol", { id: "recent-checks", "aria-label": "Recent checks, newest first", class: "runs divide-y", hidden: true });
    const root = h(
      "section",
      { class: "card checklog", "aria-labelledby": "latest-check" },
      h("div", { class: "checklog-row" }, dot, h("div", { class: "checklog-main" }, title, line), trigger),
      staleBox,
      list,
    );

    function update() {
      const data = S.feed.data;
      root.hidden = !data;
      if (!data) return;
      const now = Date.now();
      const lastRun = data.lastRun;
      const runs = data.runs;
      const stale = staleNow();
      dot.className = "checklog-dot dot-" + freshnessTone(lastRun, stale);
      if (lastRun) {
        fill(meta, 
          " \u00b7 ",
          h("time", { datetime: lastRun.at }, fmt.stampLabel(lastRun.at, now)),
          " (" + fmt.relativeAge(lastRun.atMs, now) + ")",
        );
      } else fill(meta);
      const nextLineSig = JSON.stringify([lastRun, fmt.timeZone]);
      if (nextLineSig !== lineSig) {
        lineSig = nextLineSig;
        fill(line,
          lastRun
            ? RunLine(lastRun).filter(Boolean)
            : [
                h(
                  "p",
                  { class: "muted" },
                  "No checks recorded yet. Every email and Teams check will show here, even when nothing needs you.",
                ),
              ],
        );
      }
      trigger.hidden = runs.length === 0;
      countEl.textContent = String(runs.length);
      staleBox.hidden = !(stale && lastRun);
      if (stale && lastRun) {
        staleText.textContent =
          "No checks recorded since " +
          fmt.stampLabel(lastRun.at, now) +
          ". Your COS checks may be paused or waiting for your approval in Cowork.";
      }
      const nextListSig = JSON.stringify([runs.length, runs[0] ? runs[0].at : "", runs[runs.length - 1] ? runs[runs.length - 1].at : "", fmt.timeZone, fmt.dayKey(now)]);
      if (nextListSig !== listSig) {
        listSig = nextListSig;
        fill(list,
          runs.map((run) =>
            h(
              "li",
              { class: "run-row" },
              h("time", { datetime: run.at }, fmt.stampLabel(run.at, now)),
              h("div", { style: "min-width:0" }, RunLine(run)),
            ),
          ),
        );
      }
    }
    return { root: root, update: update };
  }

  function Toolbar() {
    const items = {};
    const group = h("div", { class: "toggle-group", role: "group", "aria-label": "Show items" });
    [
      ["all", "All"],
      ["open", "Open"],
      ["completed", "Completed"],
    ].forEach(function (pair) {
      const value = pair[0];
      const count = h("span", { class: "tabular-nums" });
      const btn = h(
        "button",
        {
          type: "button",
          class: "toggle-item",
          "aria-pressed": "false",
          onclick: () => {
            S.filter = value;
            scheduleRender();
          },
        },
        pair[1],
        count,
      );
      items[value] = { btn: btn, count: count };
      group.append(btn);
    });
    const synced = h("span", { class: "text-sm muted", title: "When this screen last pulled your feed" });
    let refreshSig = null;
    const refreshBtn = h("button", { type: "button", class: "btn btn-outline btn-sm", onclick: () => refreshAll() });
    const root = h("div", { class: "toolbar" }, group, h("div", { class: "toolbar-right" }, synced, refreshBtn));

    function update(counts) {
      const c = counts || { all: 0, open: 0, completed: 0 };
      for (const key of Object.keys(items)) {
        items[key].count.textContent = String(c[key] || 0);
        items[key].btn.setAttribute("aria-pressed", String(S.filter === key));
      }
      synced.hidden = !(S.feed.updatedAt > 0);
      if (S.feed.updatedAt > 0) synced.textContent = "Synced " + fmt.timeLabel(S.feed.updatedAt);
      const fetching = S.feed.fetching || S.saved.fetching || S.tasks.fetching;
      if (refreshSig !== fetching) {
        refreshSig = fetching;
        fill(refreshBtn, fetching ? spinner("Refreshing") : icon("refresh-cw"), "Refresh");
      }
      refreshBtn.disabled = fetching;
    }
    return { root: root, update: update };
  }

  function SavedAlert() {
    const root = h(
      "div",
      { role: "alert", class: "notice saved-alert" },
      icon("triangle-alert"),
      h("span", null, "Saved statuses couldn't be loaded, so completed and deleted items may reappear."),
      h("button", { type: "button", class: "btn btn-outline btn-sm", onclick: () => loadSaved() }, "Retry"),
    );
    function update() {
      root.hidden = !S.saved.error;
    }
    return { root: root, update: update };
  }

  function FeedSkeleton() {
    const pills = () =>
      h(
        "div",
        { class: "feed-skeleton-pills" },
        h("div", { class: "skeleton", style: "height:1.25rem;width:4rem;border-radius:9999px" }),
        h("div", { class: "skeleton", style: "height:1.25rem;width:6rem;border-radius:9999px" }),
        h("div", { class: "skeleton", style: "height:1.25rem;width:3.5rem;border-radius:9999px" }),
      );
    const rowEl = () =>
      h(
        "div",
        { class: "feed-skeleton-row" },
        h("div", { class: "skeleton", style: "height:1.25rem;width:5rem" }),
        h(
          "div",
          { class: "feed-skeleton-card" },
          pills(),
          h("div", { class: "skeleton", style: "height:1.25rem;width:66%" }),
          h("div", { class: "skeleton", style: "height:1rem;width:50%" }),
          h("div", { class: "skeleton", style: "height:6rem;width:100%" }),
        ),
      );
    return h(
      "div",
      { role: "status", "aria-label": "Loading messages", class: "feed-skeleton" },
      h("div", { class: "skeleton", style: "height:1.5rem;width:18rem;max-width:100%" }),
      rowEl(),
      rowEl(),
      rowEl(),
    );
  }

  function StateBlock(kind, title, description, actionsList) {
    return h(
      "div",
      { class: "state" },
      kind === "empty" ? icon("inbox", "state-icon") : null,
      h("p", { class: "state-title" + (kind === "error" ? " is-error" : "") }, title),
      description ? h("p", { class: "state-desc" }, description) : null,
      actionsList && actionsList.length ? h("div", { style: "display:flex;flex-wrap:wrap;gap:0.5rem;justify-content:center" }, actionsList) : null,
    );
  }

  function FeedView() {
    const root = h("section", { id: "feed", tabindex: "-1", "aria-label": "Messages", class: "feed" });
    const groupsHost = h("div");
    const comms = new Map();
    const groupEls = new Map();
    let mode = null;
    let stateSig = null;

    function setMode(next, node) {
      if (mode === next && next === "list") return;
      mode = next;
      fill(root, node);
    }

    function group(key) {
      let g = groupEls.get(key);
      if (g) return g;
      const headingId = "day-" + key;
      const h2 = h("h2", { id: headingId });
      const count = h("span");
      const list = h("ol");
      const section = h("section", { class: "day-group", "aria-labelledby": headingId }, h("div", { class: "day-head" }, h2, count), list);
      g = { section: section, h2: h2, count: count, list: list };
      groupEls.set(key, g);
      return g;
    }

    function update() {
      const q = S.feed;
      const data = q.data;
      const st = stateData();
      if (!data) {
        if (q.error) {
          const missing = q.error instanceof DashError && q.error.message === "feed-missing";
          const sig = "error:" + (missing ? "missing" : "other");
          if (sig !== stateSig) {
            stateSig = sig;
            const acts = [h("button", { type: "button", class: "btn btn-outline btn-md", onclick: () => loadFeed(true) }, "Try again")];
            if (missing) acts.push(h("a", { class: "btn btn-outline btn-md", href: CONFIG.setupGuideUrl }, icon("book-open"), "Setup guide"));
            if (S.backend && S.backend.kind === "folder") {
              acts.push(h("button", { type: "button", class: "btn btn-outline btn-md", onclick: () => chooseFolder() }, icon("folder-open"), "Choose a different folder"));
            }
            setMode(
              sig,
              StateBlock(
                "error",
                missing ? "Couldn't find your feed" : "Couldn't load your messages",
                missing
                  ? "There's no comms-feed.json in " +
                      (S.backend && S.backend.kind === "folder" ? "this folder" : pick("feedFolder")) +
                      " yet. Turn on your COS automations with the setup guide, then try again."
                  : "Check your connection, then try again.",
                acts,
              ),
            );
          }
        } else if (mode !== "loading") {
          stateSig = "loading";
          setMode("loading", FeedSkeleton());
        }
        return { all: 0, open: 0, completed: 0 };
      }
      stateSig = null;

      const live = data.items.filter((item) => !st.deleted[item.id]);
      const liveIds = new Set(live.map((i) => i.id));
      comms.forEach((ci, id) => {
        if (!liveIds.has(id)) {
          ci.destroy();
          comms.delete(id);
        }
      });

      let completed = 0;
      const statuses = new Map();
      for (const item of live) {
        const status = statusOf(item);
        statuses.set(item.id, status);
        if (status === "completed") completed += 1;
      }
      const counts = { all: live.length, open: live.length - completed, completed: completed };

      const visible = live.filter((item) => {
        const status = statuses.get(item.id);
        if (S.filter === "open") return status !== "completed";
        if (S.filter === "completed") return status === "completed";
        return true;
      });

      if (visible.length === 0) {
        const sig = "empty:" + S.filter;
        if (mode !== sig) {
          setMode(
            sig,
            StateBlock(
              "empty",
              S.filter === "completed" ? "No completed items" : S.filter === "open" ? "Nothing open" : "Nothing needs your attention",
              S.filter === "all" ? "Emails and Teams chats that meet your COS criteria will appear here." : "",
            ),
          );
        }
        return counts;
      }
      setMode("list", groupsHost);

      const groups = [];
      for (const item of visible) {
        const key = fmt.dayKey(item.receivedMs);
        const last = groups[groups.length - 1];
        if (last && last.key === key) last.items.push(item);
        else groups.push({ key: key, ms: item.receivedMs, items: [item] });
      }
      const now = Date.now();
      const used = new Set();
      const sections = groups.map((g) => {
        used.add(g.key);
        const el = group(g.key);
        const rel = fmt.relativeDay(g.ms, now);
        el.h2.textContent = (rel ? rel + " \u00b7 " : "") + fmt.dayLabel(g.ms);
        el.count.textContent = g.items.length + " " + (g.items.length === 1 ? "item" : "items");
        const nodes = g.items.map((item) => {
          let ci = comms.get(item.id);
          if (!ci) {
            ci = new CommItem(item);
            comms.set(item.id, ci);
          }
          ci.update({
            item: item,
            status: statuses.get(item.id),
            sent: st.sent[item.id] || S.localSent[item.id],
            savedDraft: st.drafts[item.id],
          });
          return ci.root;
        });
        reconcile(el.list, nodes);
        return el.section;
      });
      reconcile(groupsHost, sections);
      Array.from(groupEls.keys()).forEach((key) => {
        if (!used.has(key)) groupEls.delete(key);
      });
      // Browsers without CSS field-sizing: size each new draft box once it's on the page.
      comms.forEach((ci) => {
        if (!ci.sized && ci.root.isConnected) {
          ci.sized = true;
          ci.autosize();
        }
      });
      return counts;
    }

    return { root: root, update: update };
  }

  // ----- one email or Teams item -----

  class CommItem {
    constructor(item) {
      this.item = item;
      this.domId = nextId("c");
      this.text = null;
      this.baseline = null;
      this.dirty = false;
      this.replyAll = item.reply.replyAll === true;
      this.phase = { kind: "idle" };
      this.secondsLeft = undoSeconds();
      this.timer = null;
      this.textAtSend = "";
      this.extrasAtSend = [];
      this.manual = [];
      this.chosen = {};
      this.addError = null;
      this.addOpen = false;
      this.researchOpen = false;
      this.completedOpen = false;
      this.props = null;
      this.rows = new Map();
      this.sig = {};
      this.build();
    }

    build() {
      const id = this.domId;
      this.timeEl = h("time");
      this.dot = h("span", { class: "timeline-dot", "aria-hidden": "true" });
      this.badges = h("div", { class: "badges" });
      this.content = h("div");

      this.researchBtn = h("button", {
        type: "button",
        class: "btn btn-ghost btn-sm -ml-2",
        "aria-expanded": "false",
        "aria-controls": id + "-research",
        onclick: () => {
          this.researchOpen = !this.researchOpen;
          this.paintResearch();
        },
      });
      this.researchBody = h("div", { id: id + "-research", class: "research-box", hidden: true });
      this.researchWrap = h("div", { class: "collapsible" }, this.researchBtn, this.researchBody);

      this.draftLabel = h("label", { for: id + "-draft", class: "label" });
      this.audienceEl = h("span", { class: "draft-audience" });
      this.textarea = h("textarea", { id: id + "-draft", class: "textarea" });
      this.textarea.addEventListener("input", () => {
        this.dirty = true;
        this.text = this.textarea.value;
        this.autosize();
        this.paintActions();
      });
      this.textarea.addEventListener("blur", () => this.commitDraft());

      this.sendBtn = h("button", { type: "button", class: "btn btn-default btn-md", onclick: () => this.startSend("reply") });
      this.undoBtn = h("button", { type: "button", class: "btn btn-outline btn-md", hidden: true, onclick: () => this.undoSend() });
      this.replyAllBox = h("input", { type: "checkbox", class: "checkbox", id: id + "-replyall" });
      this.replyAllBox.addEventListener("change", () => {
        this.replyAll = this.replyAllBox.checked;
        this.paintDraft();
      });
      this.replyAllRow = h(
        "div",
        { class: "inline-check" },
        this.replyAllBox,
        h("label", { for: id + "-replyall", class: "label" }, "Reply all"),
      );
      this.actionsEl = h("div", { class: "draft-actions" }, this.sendBtn, this.undoBtn, this.replyAllRow);

      this.addTrigger = h("button", {
        type: "button",
        class: "btn btn-ghost btn-sm -ml-2",
        "aria-expanded": "false",
        "aria-controls": id + "-pick",
        onclick: () => {
          this.addOpen = !this.addOpen;
          this.paintAddPeople();
        },
      });
      this.recipientList = h("ul", { class: "divide-y" });
      this.moreNote = h("p", { class: "recipients-more" });
      this.addInput = h("input", {
        type: "email",
        inputmode: "email",
        autocomplete: "off",
        spellcheck: "false",
        class: "input",
        id: id + "-pick-add",
      });
      this.addInput.addEventListener("input", () => {
        if (this.addError) this.addError = null;
        this.paintAddPeople();
      });
      this.addBtn = h("button", { type: "submit", class: "btn btn-outline btn-md" }, icon("plus"), "Add");
      this.addErrorEl = h("p", { id: id + "-pick-add-error", role: "alert", class: "field-error", hidden: true });
      this.addForm = h(
        "form",
        { class: "add-form", novalidate: true },
        h("div", { class: "add-field" }, h("label", { for: id + "-pick-add", class: "label" }, "Add by email"), this.addInput),
        this.addBtn,
      );
      this.addForm.addEventListener("submit", (event) => {
        event.preventDefault();
        this.addPerson();
      });
      this.fieldset = h(
        "fieldset",
        { class: "recipients" },
        h("legend", null, "Recipients"),
        this.recipientList,
        this.moreNote,
        this.addForm,
        this.addErrorEl,
      );
      this.addContent = h("div", { id: id + "-pick", hidden: true }, this.fieldset);
      this.addWrap = h("div", { class: "collapsible" }, this.addTrigger, this.addContent);
      this.statusLine = h("div", { class: "status-line", "aria-live": "polite" });

      this.draftPanel = h(
        "section",
        { class: "draft-panel" },
        h("div", { class: "draft-head" }, this.draftLabel, this.audienceEl),
        h("div", { class: "draft-body" }, this.textarea, this.actionsEl),
        this.addWrap,
        this.statusLine,
      );
      this.draftHost = h("div");
      this.completedTrigger = h("button", {
        type: "button",
        class: "btn btn-ghost btn-sm -ml-2",
        "aria-expanded": "false",
        "aria-controls": id + "-done",
        onclick: () => {
          this.completedOpen = !this.completedOpen;
          this.paintHost();
          if (this.completedOpen) this.autosize();
        },
      });
      this.completedContent = h("div", { id: id + "-done", hidden: true });
      this.completedWrap = h("div", { class: "collapsible" }, this.completedTrigger, this.completedContent);

      this.completeBtn = h("button", { type: "button", class: "btn btn-outline btn-sm" });
      this.completeBtn.addEventListener("click", () => {
        if (this.props.status === "completed") actions.reopen(this.item);
        else actions.markComplete(this.item);
      });
      this.deleteBtn = h(
        "button",
        { type: "button", class: "btn btn-outline btn-sm btn-danger-text", onclick: () => actions.remove(this.item) },
        icon("trash-2"),
        "Delete",
      );
      this.openLink = h("a", { class: "link", target: "_blank", rel: "noopener noreferrer" });
      this.footer = h("div", { class: "card-footer" }, this.completeBtn, this.deleteBtn, this.openLink);

      this.article = h(
        "article",
        { class: "card comm-card", "aria-labelledby": id + "-subject" },
        this.badges,
        this.content,
        this.researchWrap,
        this.draftHost,
        this.footer,
      );
      this.root = h(
        "li",
        { class: "comm-item" },
        h("div", { class: "time-col" }, this.timeEl),
        h("div", { class: "rail" }, h("span", { class: "rail-line", "aria-hidden": "true" }), this.dot, this.article),
      );
    }

    update(props) {
      this.props = props;
      this.item = props.item;
      const baseline = props.savedDraft !== undefined ? props.savedDraft : props.item.draft;
      if (this.text === null) {
        this.text = baseline;
        this.baseline = baseline;
      } else if (baseline !== this.baseline) {
        this.baseline = baseline;
        if (!this.dirty) this.text = baseline;
      }
      this.paintStatic();
      this.paint();
    }

    changed(key, value) {
      const next = typeof value === "string" ? value : JSON.stringify(value);
      if (this.sig[key] === next) return false;
      this.sig[key] = next;
      return true;
    }

    busy() {
      return this.phase.kind === "countdown" || this.phase.kind === "sending";
    }

    people() {
      const candidates = S.people.replyCandidates(this.item);
      const known = new Set(candidates.map((p) => p.email.toLowerCase()));
      return candidates.concat(this.manual.filter((p) => !known.has(p.email.toLowerCase())));
    }

    extras() {
      if (this.item.channel !== "email" || this.replyAll || S.backend.sendMode !== "send") return [];
      const out = [];
      for (const p of this.people()) {
        const line = this.chosen[p.email.toLowerCase()];
        if (line) out.push({ name: p.name, email: p.email, line: line });
      }
      return out;
    }

    audience() {
      const item = this.item;
      if (item.channel === "teams") return item.reply.audience || "This Teams chat";
      const all = S.people.replyAllRecipients(item);
      if (this.replyAll && all.length > 1) {
        const n = all.length - 1;
        return "Reply all \u00b7 " + item.from.name + " + " + n + " " + (n === 1 ? "other" : "others");
      }
      const extras = this.extras();
      if (extras.length > 0) {
        return "Reply to " + item.from.name + " + " + extras.length + " " + (extras.length === 1 ? "other" : "others");
      }
      return "Reply to " + item.from.name;
    }

    paintStatic() {
      const item = this.item;
      const id = this.domId;
      const key = [
        item.receivedAt,
        item.whyItMatters,
        item.subject,
        item.from,
        item.to,
        item.cc,
        item.toMore,
        item.summary,
        item.webLink,
        item.channel,
        fmt.timeZone,
        S.people.owner,
      ];
      if (!this.changed("static", key)) return;
      this.timeEl.dateTime = item.receivedAt;
      this.timeEl.textContent = fmt.timeLabel(item.receivedMs);
      const nodes = [];
      if (item.whyItMatters) nodes.push(h("p", { class: "why" }, item.whyItMatters));
      nodes.push(h("h3", { id: id + "-subject", class: "subject" }, item.subject));
      const dl = h(
        "dl",
        { class: "meta" },
        h("dt", null, "From"),
        h(
          "dd",
          null,
          h("span", { class: "font-medium" }, item.from.name),
          item.from.email && item.from.email !== item.from.name ? h("span", { class: "muted" }, " \u00b7 " + item.from.email) : null,
        ),
        h("dt", null, "To"),
        h("dd", null, S.people.recipientsText(item.to, item.toMore)),
      );
      if (item.cc.length > 0) dl.append(h("dt", null, "Cc"), h("dd", null, S.people.recipientsText(item.cc, 0)));
      nodes.push(dl);
      if (item.summary) nodes.push(h("p", { class: "summary" }, item.summary));
      fill(this.content, nodes);

      const link = C.safeUrl(item.webLink);
      this.openLink.hidden = !link;
      if (link) {
        this.openLink.href = link;
        fill(this.openLink, "Open in " + (item.channel === "email" ? "Outlook" : "Teams"), icon("external-link"));
      }
      this.rows.clear();
      this.sig.recipients = null;
    }

    paint() {
      const props = this.props;
      const item = props.item;
      const status = props.status;
      const meta = STATUS[status];
      this.dot.className = "timeline-dot dot-" + meta.tone;
      this.article.classList.toggle("is-completed", status === "completed");
      if (this.changed("badges", [status, item.channel, item.priority, item.receivedAt, fmt.timeZone])) {
        const pr = item.priority ? PRIORITY[item.priority] : null;
        fill(this.badges, 
          h(
            "span",
            { class: "badge badge-outline" },
            icon(item.channel === "email" ? "mail" : "message-square-text"),
            item.channel === "email" ? "Email" : "Teams",
          ),
          pr ? h("span", { class: "badge tone-" + pr.tone }, pr.label) : "",
          h("span", { class: "badge tone-" + meta.tone }, status === "completed" ? icon("circle-check") : null, meta.label),
          h("time", { class: "mobile-time", datetime: item.receivedAt }, fmt.timeLabel(item.receivedMs)),
        );
      }
      this.paintResearch();
      this.paintHost();
      this.paintDraft();
      if (this.changed("footer", status === "completed" ? "done" : "open")) {
        const done = status === "completed";
        fill(this.completeBtn, icon(done ? "rotate-ccw" : "circle-check"), done ? "Reopen" : "Mark Complete");
      }
    }

    paintResearch() {
      const r = this.item.research;
      const show = showResearch() && !!r;
      this.researchWrap.hidden = !show;
      if (!show) return;
      this.researchBtn.setAttribute("aria-expanded", String(this.researchOpen));
      this.researchBody.hidden = !this.researchOpen;
      if (this.changed("research", [r, this.researchOpen])) {
        const n = r.sources.length;
        fill(this.researchBtn, 
          icon("search"),
          "Research notes",
          n > 0 ? h("span", { class: "tabular-nums muted" }, "\u00b7 " + n + (n === 1 ? " source" : " sources")) : null,
          icon("chevron-down", "chevron"),
        );
        const parts = [];
        if (r.notes) parts.push(h("p", null, r.notes));
        if (n > 0) {
          parts.push(
            h(
              "ul",
              { "aria-label": "Sources" },
              r.sources.map((s) => {
                const url = C.safeUrl(s.url);
                return h(
                  "li",
                  null,
                  url ? h("a", { href: url, target: "_blank", rel: "noopener noreferrer" }, s.title) : h("span", null, s.title),
                );
              }),
            ),
          );
        }
        parts.push(h("p", { class: "research-private" }, "Only you see these notes. They're never sent."));
        fill(this.researchBody, parts);
      }
    }

    paintHost() {
      const completed = this.props.status === "completed";
      if (completed) {
        if (this.draftHost.firstChild !== this.completedWrap) fill(this.draftHost, this.completedWrap);
        if (this.draftPanel.parentNode !== this.completedContent) fill(this.completedContent, this.draftPanel);
        this.completedTrigger.setAttribute("aria-expanded", String(this.completedOpen));
        if (this.changed("doneTrigger", this.props.sent ? "sent" : "draft")) {
          fill(this.completedTrigger, this.props.sent ? "Show sent reply" : "Show draft", icon("chevron-down", "chevron"));
        }
        this.completedContent.hidden = !this.completedOpen;
      } else if (this.draftHost.firstChild !== this.draftPanel) {
        fill(this.draftHost, this.draftPanel);
      }
    }

    paintDraft() {
      const sent = this.props.sent;
      this.draftPanel.setAttribute("aria-label", sent ? "Sent reply" : "Draft reply");
      this.draftLabel.textContent = sent ? "Sent reply" : "Draft reply";
      this.audienceEl.textContent = this.audience();
      const value = sent ? sent.text : this.text;
      if (this.textarea.value !== value) {
        this.textarea.value = value;
        this.autosize();
      }
      this.textarea.readOnly = Boolean(sent) || this.busy();
      this.actionsEl.hidden = Boolean(sent);
      this.paintActions();
      this.paintAddPeople();
      this.paintStatus();
    }

    paintActions() {
      if (this.props.sent) return;
      const mode = S.backend.sendMode;
      const counting = this.phase.kind === "countdown";
      const sending = this.phase.kind === "sending";
      this.undoBtn.hidden = !counting;
      this.sendBtn.hidden = counting;
      if (counting) fill(this.undoBtn, icon("rotate-ccw"), "Undo (" + this.secondsLeft + ")");
      if (this.changed("send", [mode, sending])) {
        fill(this.sendBtn, 
          sending ? spinner("Sending") : icon(mode === "copy" ? "copy" : "send"),
          mode === "copy" ? "Copy & open" : "Send",
        );
        this.sendBtn.title = mode === "copy" ? "Copy this reply and open the message so you can paste and send it" : "";
      }
      this.sendBtn.disabled = !this.text.trim() || sending;
      this.replyAllRow.hidden = this.item.channel !== "email";
      this.replyAllBox.checked = this.replyAll;
      this.replyAllBox.disabled = this.busy();
    }

    paintAddPeople() {
      const item = this.item;
      const show = item.channel === "email" && !this.props.sent && !this.replyAll && S.backend.sendMode === "send";
      this.addWrap.hidden = !show;
      if (!show) return;
      const extras = this.extras();
      this.addTrigger.setAttribute("aria-expanded", String(this.addOpen));
      if (this.changed("addTrigger", extras.length)) {
        fill(this.addTrigger, 
          icon("user-plus"),
          "Add people",
          extras.length > 0 ? h("span", { class: "tabular-nums muted" }, "\u00b7 " + extras.length + " added") : null,
          icon("chevron-down", "chevron"),
        );
      }
      this.addContent.hidden = !this.addOpen;
      this.fieldset.disabled = this.busy();
      this.paintRecipients();
      this.moreNote.hidden = !(item.toMore > 0);
      this.moreNote.textContent =
        item.toMore === 1 ? "1 more recipient on the original isn't listed." : item.toMore + " more recipients on the original aren't listed.";
      this.addBtn.disabled = !this.addInput.value.trim();
      this.addErrorEl.hidden = !this.addError;
      this.addErrorEl.textContent = this.addError || "";
      if (this.addError) {
        this.addInput.setAttribute("aria-invalid", "true");
        this.addInput.setAttribute("aria-describedby", this.domId + "-pick-add-error");
      } else {
        this.addInput.removeAttribute("aria-invalid");
        this.addInput.removeAttribute("aria-describedby");
      }
    }

    senderRow() {
      const item = this.item;
      const boxId = this.domId + "-pick-sender";
      return h(
        "li",
        { class: "recipient" },
        h("input", { type: "checkbox", class: "checkbox", id: boxId, checked: true, disabled: true }),
        h(
          "label",
          { for: boxId, class: "label" },
          h("span", { class: "who" }, item.from.name),
          item.from.email && item.from.email !== item.from.name ? h("span", { class: "muted" }, " \u00b7 " + item.from.email) : null,
        ),
        S.people.isExternal(item.from.email) ? h("span", { class: "badge tone-warning" }, "External") : null,
        h("span", { class: "badge badge-outline" }, "Sender"),
      );
    }

    personRow(person) {
      const key = person.email.toLowerCase();
      const boxId = nextId(this.domId + "-pick");
      const box = h("input", { type: "checkbox", class: "checkbox", id: boxId });
      box.addEventListener("change", () => this.togglePerson(person, box.checked));
      const toBtn = h("button", { type: "button", class: "toggle-item", "aria-pressed": "false", onclick: () => this.setLine(person, "to") }, "To");
      const ccBtn = h("button", { type: "button", class: "toggle-item", "aria-pressed": "false", onclick: () => this.setLine(person, "cc") }, "Cc");
      const toggle = h("div", { class: "toggle-group toggle-sm", role: "group", "aria-label": person.name + ": To or Cc" }, toBtn, ccBtn);
      const li = h(
        "li",
        { class: "recipient" },
        box,
        h(
          "label",
          { for: boxId, class: "label" },
          h("span", { class: "who" }, person.name),
          person.email !== person.name ? h("span", { class: "muted" }, " \u00b7 " + person.email) : null,
        ),
        S.people.isExternal(person.email) ? h("span", { class: "badge tone-warning" }, "External") : null,
        toggle,
      );
      return { key: key, li: li, box: box, toBtn: toBtn, ccBtn: ccBtn, toggle: toggle };
    }

    paintRecipients() {
      const list = this.people();
      if (this.changed("recipients", [list.map((p) => p.email), S.people.owner])) {
        const rows = new Map();
        for (const p of list) {
          const key = p.email.toLowerCase();
          rows.set(key, this.rows.get(key) || this.personRow(p));
        }
        this.rows = rows;
        this.senderEl = this.senderRow();
        reconcile(this.recipientList, [this.senderEl].concat(Array.from(rows.values()).map((r) => r.li)));
      }
      this.rows.forEach((row, key) => {
        const line = this.chosen[key];
        row.box.checked = Boolean(line);
        row.toggle.hidden = !line;
        row.toBtn.setAttribute("aria-pressed", String(line === "to"));
        row.ccBtn.setAttribute("aria-pressed", String(line === "cc"));
      });
    }

    togglePerson(person, on) {
      const key = person.email.toLowerCase();
      if (on) this.chosen[key] = this.chosen[key] || person.line;
      else delete this.chosen[key];
      this.paintDraft();
    }

    setLine(person, line) {
      this.chosen[person.email.toLowerCase()] = line;
      this.paintDraft();
    }

    addPerson() {
      const email = this.addInput.value.trim();
      const key = email.toLowerCase();
      if (!C.isValidEmail(email)) this.addError = "Enter a valid email address.";
      else if (S.people.isOwner(key)) this.addError = "That's your own address.";
      else if (key === this.item.from.email.trim().toLowerCase()) this.addError = this.item.from.name + " is already on this reply.";
      if (this.addError) {
        this.paintAddPeople();
        return;
      }
      const existing = this.people().find((p) => p.email.toLowerCase() === key);
      if (!existing) this.manual.push({ name: email, email: email, line: "to" });
      this.chosen[key] = this.chosen[key] || (existing ? existing.line : "to");
      this.addInput.value = "";
      this.addError = null;
      this.paintDraft();
    }

    paintStatus() {
      const sent = this.props.sent;
      const p = this.phase.kind;
      const key = sent ? ["sent", sent.at, sent.via, sent.added || [], fmt.timeZone] : [p, p === "countdown" ? this.secondsLeft : 0];
      if (!this.changed("status", key)) return;
      let node = null;
      if (sent) {
        node = h(
          "span",
          { class: "status-sent" },
          icon("circle-check", "size-4"),
          (sent.via === "manual" ? "Marked replied " : "Sent ") + fmt.stampLabel(sent.at) + C.sentLabel(sent, this.item.channel),
        );
      } else if (p === "countdown") node = h("span", null, "Sending in " + this.secondsLeft + "s");
      else if (p === "sending") node = h("span", null, "Sending\u2026");
      else if (p === "error") node = h("span", { class: "status-error" }, "Couldn't send. Your draft is unchanged \u2014 try again.");
      else if (p === "missing") {
        node = h(
          "span",
          { class: "status-missing" },
          "The original email was moved or deleted, so it can't be answered in-thread.",
          h("button", { type: "button", class: "btn btn-outline btn-sm", onclick: () => this.startSend("newEmail") }, icon("mail"), "Send as new email"),
        );
      }
      fill(this.statusLine, node);
    }

    startSend(mode) {
      const body = this.text.trim();
      if (!body || this.busy()) return;
      if (S.backend.sendMode === "copy") {
        this.copyAndOpen(body);
        return;
      }
      this.textAtSend = body;
      this.extrasAtSend = this.extras();
      if (this.dirty) {
        this.dirty = false;
        actions.saveDraft(this.item, this.text);
      }
      let left = undoSeconds();
      if (left <= 0) {
        this.doSend(mode);
        return;
      }
      this.secondsLeft = left;
      this.phase = { kind: "countdown", mode: mode };
      this.paintDraft();
      this.timer = window.setInterval(() => {
        left -= 1;
        this.secondsLeft = left;
        if (left <= 0) {
          window.clearInterval(this.timer);
          this.timer = null;
          this.doSend(mode);
        } else this.paintDraft();
      }, 1000);
    }

    undoSend() {
      if (this.timer !== null) window.clearInterval(this.timer);
      this.timer = null;
      this.phase = { kind: "idle" };
      this.paintDraft();
    }

    async doSend(mode) {
      this.phase = { kind: "sending" };
      this.paintDraft();
      const item = this.item;
      const body = this.textAtSend;
      const added = this.extrasAtSend;
      const backend = S.backend;
      try {
        let via;
        if (item.channel === "teams") via = await backend.sendTeamsReply(item, body);
        else if (mode === "newEmail") via = await backend.sendEmailAsNew(item, body, this.replyAll, added);
        else via = await backend.sendEmailReply(item, body, this.replyAll, added);
        const names = via === "reply" || via === "newEmail" ? added.map((p) => p.name) : [];
        const record = { at: new Date().toISOString(), text: body, via: via };
        if (names.length > 0) record.added = names;
        this.phase = { kind: "idle" };
        actions.recordSent(item, record);
        this.props = Object.assign({}, this.props, { sent: record });
        this.paintDraft();
        toast.success(
          item.channel === "teams" ? "Posted in Teams" : "Reply sent",
          backend.kind === "demo" ? { description: "Demo mode: nothing actually left this page." } : undefined,
        );
      } catch (e) {
        console.error("Send failed", e);
        this.phase = e instanceof DashError && e.kind === "not-found" && item.channel === "email" ? { kind: "missing" } : { kind: "error" };
        this.paintDraft();
      }
    }

    copyAndOpen(body) {
      const item = this.item;
      const ok = copyText(body);
      const link = C.safeUrl(item.webLink);
      if (link) window.open(link, "_blank", "noopener,noreferrer");
      if (this.dirty) {
        this.dirty = false;
        actions.saveDraft(item, this.text);
      }
      const where = item.channel === "email" ? "Outlook" : "Teams";
      const how = item.channel === "email" ? (this.replyAll ? "a Reply all" : "a Reply") : "the chat";
      toast(ok ? "Reply copied" : "Couldn't copy the reply", {
        kind: ok ? "success" : "error",
        description: ok
          ? "Paste it into " + how + " in " + where + ", then mark it replied here."
          : "Select the draft and copy it yourself, then reply in " + where + ".",
        duration: 12000,
        action: {
          label: "Mark replied",
          onClick: () => actions.recordSent(item, { at: new Date().toISOString(), text: body, via: "manual" }),
        },
      });
    }

    commitDraft() {
      if (!this.dirty) return;
      this.dirty = false;
      if (this.text !== this.baseline) actions.saveDraft(this.item, this.text);
    }

    autosize() {
      if (window.CSS && CSS.supports && CSS.supports("field-sizing", "content")) return;
      const ta = this.textarea;
      if (!ta.isConnected) return;
      ta.style.height = "auto";
      ta.style.height = ta.scrollHeight + 2 + "px";
    }

    destroy() {
      if (this.timer !== null) window.clearInterval(this.timer);
      this.timer = null;
    }
  }

  // ---------------------------------------------------------------------------
  // Settings dialog
  // ---------------------------------------------------------------------------

  /* On GitHub Pages, open docs on github.com so Markdown renders; locally, use the file. */
  function repoLink(path) {
    const host = window.location.hostname;
    if (!/\.github\.io$/i.test(host)) return path;
    const owner = host.split(".")[0];
    const repo = window.location.pathname.split("/").filter(Boolean)[0] || owner + ".github.io";
    return "https://github.com/" + owner + "/" + repo + "/blob/main/" + path;
  }

  function field(label, input, help) {
    return h("div", { class: "field" }, h("label", { for: input.id, class: "label" }, label), input, help ? h("p", { class: "field-help" }, help) : null);
  }

  function openSettings() {
    const owner = h("input", {
      id: "set-owner",
      class: "input",
      type: "email",
      autocomplete: "email",
      placeholder: S.ownerEmail || S.inferredOwner || "you@company.com",
      value: settings.ownerEmail || "",
    });
    const domains = h("input", {
      id: "set-domains",
      class: "input",
      placeholder: S.people.domains.join(", ") || "contoso.com",
      value: Array.isArray(settings.internalDomains) ? settings.internalDomains.join(", ") : settings.internalDomains || "",
    });
    const tz = h("input", { id: "set-tz", class: "input", placeholder: C.resolveTimeZone(""), value: settings.timeZone || "" });
    const research = h("input", { id: "set-research", type: "checkbox", class: "checkbox", checked: showResearch() });
    const clientId = h("input", {
      id: "set-client",
      class: "input",
      placeholder: "00000000-0000-0000-0000-000000000000",
      value: settings.clientId || "",
      autocomplete: "off",
      spellcheck: "false",
    });
    const tenant = h("input", {
      id: "set-tenant",
      class: "input",
      placeholder: CONFIG.tenantId || "organizations",
      value: settings.tenantId || "",
      autocomplete: "off",
      spellcheck: "false",
    });
    const folder = h("input", { id: "set-folder", class: "input", placeholder: CONFIG.feedFolder, value: settings.feedFolder || "" });
    const error = h("p", { class: "field-error", role: "alert", hidden: true });

    const connection = S.backend
      ? h(
          "div",
          { class: "connection-line" },
          h("span", null, "Connected to ", h("strong", { class: "font-medium" }, S.backend.label)),
          h(
            "button",
            { type: "button", class: "btn btn-outline btn-sm", onclick: () => disconnect(dlg) },
            S.backend.kind === "graph" ? "Sign out" : S.backend.kind === "demo" ? "Leave demo" : "Disconnect",
          ),
        )
      : h("p", { class: "field-help" }, "Not connected yet.");

    const form = h(
      "form",
      { method: "dialog" },
      h("h2", { id: "settings-title", class: "dialog-title" }, "Dashboard settings"),
      h("p", { class: "dialog-desc" }, "Saved in this browser only. Nothing here is shared or sent anywhere."),
      h(
        "div",
        { class: "dialog-body" },
        connection,
        h(
          "div",
          { class: "dialog-section" },
          h("h3", null, "You"),
          field("Your email address", owner, "Shown as \u201cYou\u201d and never added to reply-all lists. Leave blank to detect it."),
          field("Company email domains", domains, "Comma-separated. Anyone else is flagged External before you send."),
          field("Time zone", tz, "For example America/New_York. Leave blank to use this browser's time zone."),
          h(
            "div",
            { class: "inline-check" },
            research,
            h("label", { for: "set-research", class: "label", style: "font-weight:400" }, "Show the research notes saved with each draft"),
          ),
        ),
        h(
          "div",
          { class: "dialog-section" },
          h("h3", null, "Microsoft 365 sign-in"),
          h(
            "p",
            { class: "field-help" },
            "Needed only to read your feed from OneDrive and send replies from this page. ",
            h("a", { href: repoLink("docs/microsoft-365-sign-in.md"), class: "link", style: "font-size:inherit", target: "_blank", rel: "noopener" }, "How to get an app ID"),
          ),
          field("Application (client) ID", clientId),
          field("Directory (tenant) ID or domain", tenant, "Leave blank to allow any work account."),
          field("Feed folder in OneDrive", folder, "Path from the top of your OneDrive."),
        ),
        error,
      ),
      h(
        "div",
        { class: "dialog-footer" },
        h("button", { type: "button", class: "btn btn-outline btn-md", onclick: () => dlg.close() }, "Cancel"),
        h("button", { type: "submit", class: "btn btn-default btn-md", value: "save" }, "Save"),
      ),
    );
    const dlg = h("dialog", { class: "dialog", "aria-labelledby": "settings-title" }, form);
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const ownerValue = owner.value.trim();
      if (ownerValue && !C.isValidEmail(ownerValue)) {
        error.hidden = false;
        error.textContent = "Enter a valid email address, or leave it blank.";
        owner.focus();
        return;
      }
      const tzValue = tz.value.trim();
      if (tzValue && C.resolveTimeZone(tzValue) !== tzValue) {
        error.hidden = false;
        error.textContent = "That time zone isn't recognized. Try a name like America/Chicago.";
        tz.focus();
        return;
      }
      const clientChanged = clientId.value.trim() !== (settings.clientId || "") || tenant.value.trim() !== (settings.tenantId || "");
      writeSettings({
        ownerEmail: ownerValue,
        internalDomains: domains.value
          .split(/[\s,;]+/)
          .map((d) => d.trim().replace(/^@/, ""))
          .filter(Boolean),
        timeZone: tzValue,
        showResearchNotes: research.checked,
        clientId: clientId.value.trim(),
        tenantId: tenant.value.trim(),
        feedFolder: folder.value.trim().replace(/^\/+|\/+$/g, ""),
      });
      if (clientChanged) msalApp = null;
      fmt = makeFmt();
      refreshPeople();
      dlg.close();
      if (S.backend && S.backend.kind === "graph") loadFeed(true);
      scheduleRender();
    });
    dlg.addEventListener("close", () => dlg.remove());
    document.body.append(dlg);
    dlg.showModal();
  }

  // ---------------------------------------------------------------------------
  // Connecting
  // ---------------------------------------------------------------------------

  function stopPolling() {
    S.timers.forEach((t) => clearInterval(t));
    S.timers = [];
  }

  function startPolling() {
    const ms = Math.max(15, Number(CONFIG.pollSeconds) || 60) * 1000;
    S.timers.push(
      setInterval(() => {
        loadFeed();
        if (S.pendingWrites === 0) loadSaved();
        if (S.pendingTaskWrites === 0) loadTasks();
      }, ms),
    );
    S.timers.push(setInterval(scheduleRender, 30000));
  }

  function onFocus() {
    if (!S.backend) return;
    if (Date.now() - (S.feed.updatedAt || 0) > 30000) {
      loadFeed();
      if (S.pendingWrites === 0) loadSaved();
      if (S.pendingTaskWrites === 0) loadTasks();
    }
  }
  window.addEventListener("focus", onFocus);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) onFocus();
  });

  async function connect(backend) {
    stopPolling();
    S.backend = backend;
    S.feed = freshQuery();
    S.saved = freshQuery();
    S.tasks = freshQuery();
    S.filter = "all";
    S.localSent = {};
    S.pendingWrites = 0;
    S.pendingTaskWrites = 0;
    S.writeChain = Promise.resolve();
    S.taskChain = Promise.resolve();
    S.ownerEmail = "";
    S.inferredOwner = "";
    S.connectError = null;
    S.reconnectHandle = null;
    view = null;
    refreshPeople();
    scheduleRender();
    backend
      .ownerEmail()
      .then((email) => {
        if (backend !== S.backend || !email) return;
        S.ownerEmail = email;
        refreshPeople();
        scheduleRender();
      })
      .catch(() => {});
    loadFeed(true);
    loadSaved();
    loadTasks();
    startPolling();
  }

  async function disconnect(dlg) {
    const backend = S.backend;
    if (dlg) dlg.close();
    stopPolling();
    S.backend = null;
    S.feed = freshQuery();
    view = null;
    writeSettings({ mode: "" });
    if (backend && backend.kind === "folder") await idbDel("folder").catch(() => {});
    if (backend && backend.kind === "demo") backend.forget();
    scheduleRender();
    if (backend && backend.kind === "graph" && msalApp) {
      try {
        await msalApp.logoutRedirect({ account: backend.account, postLogoutRedirectUri: redirectUri() });
      } catch (e) {
        console.error("Sign-out failed", e);
      }
    }
  }

  function startDemo(remember) {
    if (remember) writeSettings({ mode: "demo" });
    connect(new DemoBackend());
  }

  async function locateFeedFolder(dir) {
    const leaf = String(pick("feedFolder")).split("/").filter(Boolean).pop() || "COS Comms Dashboard";
    const paths = [[], [leaf], ["Cowork", leaf], ["Documents", "Cowork", leaf]];
    for (const path of paths) {
      try {
        let d = dir;
        for (const segment of path) d = await d.getDirectoryHandle(segment);
        await d.getFileHandle(FEED_NAME);
        return d;
      } catch (e) {
        /* try the next likely location */
      }
    }
    return null;
  }

  async function chooseFolder() {
    if (typeof window.showDirectoryPicker !== "function") {
      S.connectError = "Opening a folder needs Microsoft Edge or Google Chrome on a computer. Use Microsoft 365 sign-in or the demo instead.";
      scheduleRender();
      return;
    }
    const saved = S.reconnectHandle;
    if (saved) {
      try {
        const perm = await saved.requestPermission({ mode: "readwrite" });
        if (perm === "granted") {
          writeSettings({ mode: "folder" });
          connect(new FolderBackend(saved));
          return;
        }
      } catch (e) {
        /* fall through to a fresh pick */
      }
    }
    let dir;
    try {
      dir = await window.showDirectoryPicker({ id: "cos-dashboard", mode: "readwrite", startIn: "documents" });
    } catch (e) {
      if (e && e.name === "AbortError") return;
      S.connectError = "Couldn't open that folder. Try again.";
      scheduleRender();
      return;
    }
    const found = await locateFeedFolder(dir);
    if (!found) {
      S.connectError =
        "There's no comms-feed.json in \u201c" +
        dir.name +
        "\u201d. Choose the COS Comms Dashboard folder inside Documents \u203a Cowork in your synced OneDrive, or follow the setup guide to create it.";
      if (S.backend) toast.error("That folder doesn't have a COS feed.");
      scheduleRender();
      return;
    }
    await idbSet("folder", found).catch(() => {});
    writeSettings({ mode: "folder" });
    connect(new FolderBackend(found));
  }

  async function signIn() {
    try {
      const app = await getMsalApp();
      writeSettings({ mode: "graph" });
      await app.loginRedirect({ scopes: GRAPH_SCOPES, prompt: "select_account", redirectUri: redirectUri() });
    } catch (e) {
      console.error("Sign-in failed", e);
      S.connectError =
        e instanceof DashError && e.message === "no-client-id"
          ? "Add your organization's Application (client) ID in Settings first."
          : "Couldn't start sign-in. Check the app ID and tenant in Settings, and that this page's address is registered as a redirect URI.";
      scheduleRender();
    }
  }

  async function boot() {
    document.title = CONFIG.appTitle;
    const titleEl = document.getElementById("appTitle");
    if (titleEl) titleEl.textContent = CONFIG.appTitle;
    const params = new URLSearchParams(window.location.search);
    if (params.has("demo")) {
      startDemo(false);
      return;
    }
    const mode = settings.mode;
    if (mode === "demo") {
      startDemo(false);
      return;
    }
    if (mode === "graph" && String(pick("clientId") || "").trim()) {
      try {
        const app = await getMsalApp();
        const result = await app.handleRedirectPromise();
        const account = (result && result.account) || app.getActiveAccount() || app.getAllAccounts()[0];
        if (account) {
          app.setActiveAccount(account);
          connect(new GraphBackend(app, account));
          return;
        }
      } catch (e) {
        console.error("Sign-in failed", e);
        S.connectError = "We couldn't finish signing you in. Check the app ID in Settings, then try again.";
      }
    }
    if (mode === "folder") {
      try {
        const handle = await idbGet("folder");
        if (handle) {
          const perm = await handle.queryPermission({ mode: "readwrite" });
          if (perm === "granted") {
            connect(new FolderBackend(handle));
            return;
          }
          S.reconnectHandle = handle;
        }
      } catch (e) {
        console.error("Folder reconnect failed", e);
      }
    }
    scheduleRender();
  }

  boot();
})();
