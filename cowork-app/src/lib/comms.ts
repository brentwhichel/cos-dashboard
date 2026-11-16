// Data layer for the COS Communications Dashboard.
//
// - The FEED (qualifying emails / Teams chats, each with a proposed reply) is
//   appended by your Chief of Staff automations to a JSON file in your
//   OneDrive. The app only ever READS the feed.
// - The app's own STATE (completed / deleted / sent / edited drafts) lives in a
//   separate JSON file the app alone writes, so the two writers never collide.
// - Send uses your own Outlook / Teams connections and only runs when you
//   press Send on a specific item.

import {
  MicrosoftTeamsService,
  Office365OutlookService,
  OneDriveforBusinessService,
} from "../../generated";

// ---- Set these for the person who owns this dashboard -----------------------
/** Your work email. Shown as "You" and never added to reply-all lists. */
export const OWNER_EMAIL = "you@contoso.com";
/** Your IANA time zone, for example "America/New_York" or "Europe/London". */
export const TIME_ZONE = "America/New_York";
/** Email domains inside your company. Anyone else is flagged External. Empty = your own domain. */
export const INTERNAL_DOMAINS: string[] = [];
// -----------------------------------------------------------------------------

const FOLDER = "/Documents/Cowork/COS Comms Dashboard";
const FEED_PATH = `${FOLDER}/comms-feed.json`;
const STATE_NAME = "dashboard-state.json";
const STATE_PATH = `${FOLDER}/${STATE_NAME}`;
// Your own tasks live in their own file so no other writer (or an older
// copy of this app still open in a tab) can ever overwrite them.
const TASKS_NAME = "my-tasks.json";
const TASKS_PATH = `${FOLDER}/${TASKS_NAME}`;
const MAX_TASK_LENGTH = 500;

export type Channel = "email" | "teams";
export type Priority = "high" | "medium" | "low";

export interface Person {
  name: string;
  email: string;
}

export type RecipientLine = "to" | "cc";

/** Someone you add to a reply that otherwise goes to the sender only. */
export interface ExtraRecipient {
  name: string;
  email: string;
  line: RecipientLine;
}

export interface ReplyTarget {
  messageId?: string;
  internetMessageId?: string;
  conversationId?: string;
  replyAll?: boolean;
  chatId?: string;
  teamId?: string;
  channelId?: string;
  parentMessageId?: string;
  audience?: string;
}

/** Private research notes the automations save with each draft (never sent). */
export interface Research {
  notes: string;
  sources: { title: string; url: string }[];
}

export interface FeedItem {
  id: string;
  channel: Channel;
  receivedAt: string;
  receivedMs: number;
  from: Person;
  to: Person[];
  cc: Person[];
  toMore: number;
  subject: string;
  summary: string;
  whyItMatters: string;
  priority: Priority | null;
  draft: string;
  research: Research | null;
  reply: ReplyTarget;
  webLink?: string;
}

// Every COS check (email screening, key Teams chat, Teams sweep) records itself
// in the feed — lastRun plus a bounded recentRuns log — even when nothing is
// added, so the dashboard always shows when it was last refreshed.
export type RunSource = "email" | "teams-chat" | "teams-sweep" | "other";
export type RunOutcome = "added" | "skipped" | "error";

export interface RunEntry {
  at: string;
  atMs: number;
  source: RunSource;
  outcome: RunOutcome;
  subject: string;
  from: string;
  note: string;
}

export interface FeedSnapshot {
  items: FeedItem[];
  /** Newest check first. */
  runs: RunEntry[];
  /** The most recent check, or null when none has been recorded yet. */
  lastRun: RunEntry | null;
}

export type SendVia = "reply" | "replyAll" | "newEmail" | "teams";

export interface SentRecord {
  at: string;
  text: string;
  via: SendVia;
  /** Names of the people you added to a sender-only reply. */
  added?: string[];
}

export interface DashboardState {
  version: 1;
  completed: Record<string, string>;
  deleted: Record<string, string>;
  sent: Record<string, SentRecord>;
  drafts: Record<string, string>;
}

export class CommsError extends Error {
  readonly kind: "feed" | "state" | "save" | "send" | "not-found";
  constructor(kind: CommsError["kind"], message: string) {
    super(message);
    this.kind = kind;
  }
}

export function emptyState(): DashboardState {
  return { version: 1, completed: {}, deleted: {}, sent: {}, drafts: {} };
}

// ---------------------------------------------------------------------------
// Payload decoding — the file-content operation can surface raw bytes, base64
// text, a `$content` envelope, or (when the service infers JSON) an already
// parsed object. Accept exactly those shapes; anything else is an error.
// ---------------------------------------------------------------------------

function bytesToText(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}

function base64ToText(value: string): string {
  const binary = atob(value.replace(/\s+/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytesToText(bytes);
}

function decodeJson(payload: unknown): unknown {
  if (payload instanceof Uint8Array) return JSON.parse(stripBom(bytesToText(payload)));
  if (payload instanceof ArrayBuffer) {
    return JSON.parse(stripBom(bytesToText(new Uint8Array(payload))));
  }
  if (typeof payload === "string") {
    const trimmed = stripBom(payload.trim());
    if (trimmed.startsWith("{") || trimmed.startsWith("[")) return JSON.parse(trimmed);
    return JSON.parse(stripBom(base64ToText(trimmed)));
  }
  if (payload && typeof payload === "object") {
    const envelope = payload as Record<string, unknown>;
    if (typeof envelope["$content"] === "string") return decodeJson(envelope["$content"]);
    return payload;
  }
  throw new Error("Unsupported file payload");
}

function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

function errorStatus(error: unknown): number | undefined {
  if (error && typeof error === "object" && "status" in error) {
    const status = (error as { status?: unknown }).status;
    if (typeof status === "number") return status;
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Feed normalisation — the feed is written by an automation, so validate every
// field and drop items that lack the essentials rather than rendering garbage.
// ---------------------------------------------------------------------------

function str(value: unknown): string {
  return typeof value === "string" ? value : "";
}

function person(value: unknown): Person | null {
  if (!value || typeof value !== "object") return null;
  const v = value as Record<string, unknown>;
  const name = str(v.name).trim();
  const email = str(v.email).trim();
  if (!name && !email) return null;
  return { name: name || email, email };
}

function people(value: unknown): Person[] {
  if (!Array.isArray(value)) return [];
  return value.map(person).filter((p): p is Person => p !== null);
}

function priority(value: unknown): Priority | null {
  const v = str(value).toLowerCase();
  return v === "high" || v === "medium" || v === "low" ? v : null;
}

function reply(value: unknown): ReplyTarget {
  if (!value || typeof value !== "object") return {};
  const v = value as Record<string, unknown>;
  const out: ReplyTarget = {};
  for (const key of [
    "messageId",
    "internetMessageId",
    "conversationId",
    "chatId",
    "teamId",
    "channelId",
    "parentMessageId",
    "audience",
  ] as const) {
    const s = str(v[key]).trim();
    if (s) out[key] = s;
  }
  if (typeof v.replyAll === "boolean") out.replyAll = v.replyAll;
  return out;
}

function research(value: unknown): Research | null {
  if (!value || typeof value !== "object") return null;
  const v = value as Record<string, unknown>;
  const notes = str(v.notes).trim();
  const sources: { title: string; url: string }[] = [];
  if (Array.isArray(v.sources)) {
    for (const raw of v.sources) {
      if (!raw || typeof raw !== "object") continue;
      const s = raw as Record<string, unknown>;
      const title = str(s.title).trim();
      const url = str(s.url).trim();
      if (title || url) sources.push({ title: title || url, url });
    }
  }
  if (!notes && sources.length === 0) return null;
  return { notes, sources: sources.slice(0, 10) };
}

function normalizeItem(raw: unknown): FeedItem | null {
  if (!raw || typeof raw !== "object") return null;
  const v = raw as Record<string, unknown>;
  const id = str(v.id).trim();
  const channel = str(v.channel).toLowerCase();
  const receivedAt = str(v.receivedAt).trim();
  const receivedMs = Date.parse(receivedAt);
  if (!id || (channel !== "email" && channel !== "teams") || Number.isNaN(receivedMs)) {
    return null;
  }
  const from = person(v.from) ?? { name: "Unknown sender", email: "" };
  const toMore = typeof v.toMore === "number" && v.toMore > 0 ? Math.floor(v.toMore) : 0;
  return {
    id,
    channel,
    receivedAt,
    receivedMs,
    from,
    to: people(v.to),
    cc: people(v.cc),
    toMore,
    subject: str(v.subject).trim() || "(no subject)",
    summary: str(v.summary).trim(),
    whyItMatters: str(v.whyItMatters).trim(),
    priority: priority(v.priority),
    draft: str(v.draft),
    research: research(v.research),
    reply: reply(v.reply),
    webLink: str(v.webLink).trim() || undefined,
  };
}

function runSource(value: unknown): RunSource {
  const v = str(value).trim().toLowerCase().replace(/[\s_]+/g, "-");
  if (v === "email" || v === "mail") return "email";
  if (v === "teams-chat" || v === "teams" || v === "chat") return "teams-chat";
  if (v === "teams-sweep" || v === "sweep") return "teams-sweep";
  return "other";
}

function runOutcome(value: unknown): RunOutcome {
  const v = str(value).trim().toLowerCase();
  if (v.startsWith("add")) return "added";
  if (/^(err|fail|block|incomplete|stall)/.test(v)) return "error";
  return "skipped";
}

function normalizeRun(raw: unknown): RunEntry | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const v = raw as Record<string, unknown>;
  const at = str(v.at).trim();
  const atMs = Date.parse(at);
  if (!at || Number.isNaN(atMs)) return null;
  return {
    at,
    atMs,
    source: runSource(v.source),
    outcome: runOutcome(v.outcome),
    subject: str(v.subject).trim(),
    from: str(v.from).trim(),
    note: str(v.note).trim(),
  };
}

function normalizeFeed(json: unknown): FeedSnapshot {
  const root: Record<string, unknown> =
    json && typeof json === "object" && !Array.isArray(json) ? (json as Record<string, unknown>) : {};
  const rawItems = Array.isArray(root.items) ? (root.items as unknown[]) : null;
  if (rawItems === null) throw new Error("Feed file has no items array");
  const seen = new Set<string>();
  const items: FeedItem[] = [];
  for (const raw of rawItems) {
    const item = normalizeItem(raw);
    if (!item || seen.has(item.id)) continue;
    seen.add(item.id);
    items.push(item);
  }
  items.sort((a, b) => b.receivedMs - a.receivedMs);

  // Checks are appended oldest → newest; reverse first so same-minute entries
  // keep newest-first order through the (stable) sort.
  const rawRuns = Array.isArray(root.recentRuns) ? [...(root.recentRuns as unknown[])].reverse() : [];
  const runKeys = new Set<string>();
  const runs: RunEntry[] = [];
  for (const raw of rawRuns) {
    const run = normalizeRun(raw);
    if (!run) continue;
    const key = `${run.atMs}|${run.source}|${run.outcome}|${run.subject}`;
    if (runKeys.has(key)) continue;
    runKeys.add(key);
    runs.push(run);
  }
  runs.sort((a, b) => b.atMs - a.atMs);
  const declared = normalizeRun(root.lastRun);
  const newest = runs[0] ?? null;
  const lastRun = declared && (!newest || declared.atMs >= newest.atMs) ? declared : newest;
  return { items, runs, lastRun };
}

function normalizeState(json: unknown): DashboardState {
  const state = emptyState();
  if (!json || typeof json !== "object") return state;
  const v = json as Record<string, unknown>;
  const stringMap = (value: unknown): Record<string, string> => {
    const out: Record<string, string> = {};
    if (value && typeof value === "object" && !Array.isArray(value)) {
      for (const [k, val] of Object.entries(value as Record<string, unknown>)) {
        if (typeof val === "string") out[k] = val;
      }
    }
    return out;
  };
  state.completed = stringMap(v.completed);
  state.deleted = stringMap(v.deleted);
  state.drafts = stringMap(v.drafts);
  if (v.sent && typeof v.sent === "object" && !Array.isArray(v.sent)) {
    for (const [k, val] of Object.entries(v.sent as Record<string, unknown>)) {
      if (!val || typeof val !== "object") continue;
      const r = val as Record<string, unknown>;
      const via = str(r.via);
      if (typeof r.at === "string" && typeof r.text === "string") {
        const added = Array.isArray(r.added)
          ? r.added.filter((n): n is string => typeof n === "string" && n.trim() !== "")
          : [];
        state.sent[k] = {
          at: r.at,
          text: r.text,
          via: via === "replyAll" || via === "newEmail" || via === "teams" ? via : "reply",
          ...(added.length > 0 ? { added } : {}),
        };
      }
    }
  }
  return state;
}

// ---------------------------------------------------------------------------
// Reads
// ---------------------------------------------------------------------------

export async function fetchFeed(): Promise<FeedSnapshot> {
  const result = await OneDriveforBusinessService.GetFileContentByPath(FEED_PATH, false);
  if (!result.success) {
    console.error("Feed read failed", result.error);
    throw new CommsError("feed", "feed-unavailable");
  }
  try {
    return normalizeFeed(decodeJson(result.data));
  } catch (error) {
    console.error("Feed parse failed", error);
    throw new CommsError("feed", "feed-unreadable");
  }
}

export interface StateSnapshot {
  state: DashboardState;
  fileId: string | null;
}

export async function fetchState(): Promise<StateSnapshot> {
  const meta = await OneDriveforBusinessService.GetFileMetadataByPath(STATE_PATH);
  if (!meta.success) {
    if (errorStatus(meta.error) === 404) return { state: emptyState(), fileId: null };
    console.error("State lookup failed", meta.error);
    throw new CommsError("state", "state-unavailable");
  }
  const fileId = meta.data?.Id ?? null;
  const content = await OneDriveforBusinessService.GetFileContentByPath(STATE_PATH, false);
  if (!content.success) {
    console.error("State read failed", content.error);
    throw new CommsError("state", "state-unavailable");
  }
  try {
    return { state: normalizeState(decodeJson(content.data)), fileId };
  } catch (error) {
    console.error("State parse failed", error);
    throw new CommsError("state", "state-unreadable");
  }
}

// ---------------------------------------------------------------------------
// Writes — always read the latest saved state first, apply one change, save.
// ---------------------------------------------------------------------------

async function writeState(state: DashboardState, fileId: string | null): Promise<string> {
  const body = JSON.stringify(state, null, 1);
  if (fileId) {
    const updated = await OneDriveforBusinessService.UpdateFile(fileId, body);
    if (!updated.success) {
      console.error("State save failed", updated.error);
      throw new CommsError("save", "save-failed");
    }
    return fileId;
  }
  const created = await OneDriveforBusinessService.CreateFile(FOLDER, STATE_NAME, body);
  if (!created.success || !created.data?.Id) {
    console.error("State create failed", created.error);
    throw new CommsError("save", "save-failed");
  }
  return created.data.Id;
}

export async function mutateState(
  change: (state: DashboardState) => void,
): Promise<StateSnapshot> {
  const latest = await fetchState();
  const next: DashboardState = {
    version: 1,
    completed: { ...latest.state.completed },
    deleted: { ...latest.state.deleted },
    sent: { ...latest.state.sent },
    drafts: { ...latest.state.drafts },
  };
  change(next);
  const fileId = await writeState(next, latest.fileId);
  return { state: next, fileId };
}

// ---------------------------------------------------------------------------
// My tasks — your own to-dos, kept in their own OneDrive file. Same rule as
// the state file: read the latest, apply one change, save.
// ---------------------------------------------------------------------------

export interface TaskItem {
  id: string;
  text: string;
  createdAt: string;
  completedAt?: string;
}

export interface TasksSnapshot {
  tasks: TaskItem[];
  fileId: string | null;
}

export function cleanTaskText(value: string): string {
  return value.replace(/\s+/g, " ").trim().slice(0, MAX_TASK_LENGTH);
}

export function newTaskId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function normalizeTasks(json: unknown): TaskItem[] {
  const list =
    Array.isArray(json)
      ? json
      : json && typeof json === "object" && Array.isArray((json as Record<string, unknown>).tasks)
        ? ((json as Record<string, unknown>).tasks as unknown[])
        : [];
  const seen = new Set<string>();
  const out: TaskItem[] = [];
  for (const raw of list) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) continue;
    const v = raw as Record<string, unknown>;
    const id = str(v.id).trim();
    const text = cleanTaskText(str(v.text));
    if (!id || !text || seen.has(id)) continue;
    seen.add(id);
    const createdAt = str(v.createdAt).trim();
    const completedAt = str(v.completedAt).trim();
    out.push({ id, text, createdAt, ...(completedAt ? { completedAt } : {}) });
  }
  return out;
}

export async function fetchTasks(): Promise<TasksSnapshot> {
  const meta = await OneDriveforBusinessService.GetFileMetadataByPath(TASKS_PATH);
  if (!meta.success) {
    if (errorStatus(meta.error) === 404) return { tasks: [], fileId: null };
    console.error("Tasks lookup failed", meta.error);
    throw new CommsError("state", "tasks-unavailable");
  }
  const fileId = meta.data?.Id ?? null;
  const content = await OneDriveforBusinessService.GetFileContentByPath(TASKS_PATH, false);
  if (!content.success) {
    console.error("Tasks read failed", content.error);
    throw new CommsError("state", "tasks-unavailable");
  }
  try {
    return { tasks: normalizeTasks(decodeJson(content.data)), fileId };
  } catch (error) {
    console.error("Tasks parse failed", error);
    throw new CommsError("state", "tasks-unreadable");
  }
}

async function writeTasks(tasks: TaskItem[], fileId: string | null): Promise<string> {
  const body = JSON.stringify({ version: 1, tasks }, null, 1);
  if (fileId) {
    const updated = await OneDriveforBusinessService.UpdateFile(fileId, body);
    if (!updated.success) {
      console.error("Tasks save failed", updated.error);
      throw new CommsError("save", "save-failed");
    }
    return fileId;
  }
  const created = await OneDriveforBusinessService.CreateFile(FOLDER, TASKS_NAME, body);
  if (!created.success || !created.data?.Id) {
    console.error("Tasks create failed", created.error);
    throw new CommsError("save", "save-failed");
  }
  return created.data.Id;
}

/** `change` must be pure and safe to apply twice (optimistically, then to the saved copy). */
export async function mutateTasks(
  change: (tasks: TaskItem[]) => TaskItem[],
): Promise<TasksSnapshot> {
  const latest = await fetchTasks();
  const next = change(latest.tasks.map((t) => ({ ...t })));
  const fileId = await writeTasks(next, latest.fileId);
  return { tasks: next, fileId };
}

// ---------------------------------------------------------------------------
// Send — only ever called from an explicit Send press on one item.
// ---------------------------------------------------------------------------

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function toHtml(text: string): string {
  return text
    .trim()
    .split(/\r?\n/)
    .map((line) => escapeHtml(line))
    .join("<br>");
}

function addresses(list: Person[]): string[] {
  return list
    .map((p) => p.email.trim())
    .filter((e) => e && e.toLowerCase() !== OWNER_EMAIL.toLowerCase());
}

export function replyAllRecipients(item: FeedItem): string[] {
  const all = [item.from.email, ...addresses(item.to), ...addresses(item.cc)];
  return Array.from(new Set(all.filter(Boolean).map((e) => e.toLowerCase())));
}

const EMAIL_PATTERN = /^[^\s@;,<>"]+@[^\s@;,<>"]+\.[^\s@;,<>"]+$/;

export function isValidEmail(value: string): boolean {
  return EMAIL_PATTERN.test(value.trim());
}

/** Anyone outside your company, flagged so customers stand out before sending. */
export function isExternal(email: string): boolean {
  const domain = email.trim().toLowerCase().split("@")[1] ?? "";
  if (!domain) return false;
  const internal =
    INTERNAL_DOMAINS.length > 0
      ? INTERNAL_DOMAINS.map((d) => d.trim().toLowerCase()).filter(Boolean)
      : [OWNER_EMAIL.trim().toLowerCase().split("@")[1] ?? ""];
  return !internal.some((d) => d !== "" && (domain === d || domain.endsWith(`.${d}`)));
}

/** Thread participants you can add to a sender-only reply (never yourself or the sender). */
export function replyCandidates(item: FeedItem): ExtraRecipient[] {
  const seen = new Set([OWNER_EMAIL.toLowerCase(), item.from.email.trim().toLowerCase()]);
  const out: ExtraRecipient[] = [];
  const lists: [Person[], RecipientLine][] = [
    [item.to, "to"],
    [item.cc, "cc"],
  ];
  for (const [list, line] of lists) {
    for (const p of list) {
      const email = p.email.trim();
      const key = email.toLowerCase();
      if (!email || seen.has(key)) continue;
      seen.add(key);
      out.push({ name: p.name.trim() || email, email, line });
    }
  }
  return out;
}

function cleanExtras(item: FeedItem, extras: ExtraRecipient[]): ExtraRecipient[] {
  const seen = new Set([OWNER_EMAIL.toLowerCase(), item.from.email.trim().toLowerCase()]);
  const out: ExtraRecipient[] = [];
  for (const r of extras) {
    const email = r.email.trim();
    const key = email.toLowerCase();
    if (!isValidEmail(email) || seen.has(key)) continue;
    seen.add(key);
    out.push({ ...r, email });
  }
  return out;
}

export async function sendEmailReply(
  item: FeedItem,
  text: string,
  replyAll: boolean,
  extras: ExtraRecipient[] = [],
): Promise<SendVia> {
  const messageId = item.reply.messageId;
  if (!messageId) throw new CommsError("not-found", "no-message-id");
  // A sender-only reply always reaches the sender; anyone you added is
  // included on To or Cc alongside them.
  const added = replyAll ? [] : cleanExtras(item, extras);
  const to = added.filter((r) => r.line === "to").map((r) => r.email);
  const cc = added.filter((r) => r.line === "cc").map((r) => r.email);
  const result = await Office365OutlookService.ReplyToV3(messageId, {
    Body: `<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">${toHtml(text)}</div>`,
    ReplyAll: replyAll,
    Importance: "Normal",
    ...(to.length > 0 ? { To: to.join(";") } : {}),
    ...(cc.length > 0 ? { Cc: cc.join(";") } : {}),
  });
  if (!result.success) {
    console.error("Email reply failed", result.error);
    if (errorStatus(result.error) === 404) throw new CommsError("not-found", "original-missing");
    throw new CommsError("send", "send-failed");
  }
  return replyAll ? "replyAll" : "reply";
}

export async function sendEmailAsNew(
  item: FeedItem,
  text: string,
  replyAll: boolean,
  extras: ExtraRecipient[] = [],
): Promise<SendVia> {
  const to = item.from.email.trim();
  if (!to) throw new CommsError("send", "no-recipient");
  const added = replyAll ? [] : cleanExtras(item, extras);
  const toList = [to, ...added.filter((r) => r.line === "to").map((r) => r.email)];
  const cc = replyAll
    ? replyAllRecipients(item).filter((e) => e !== to.toLowerCase())
    : added.filter((r) => r.line === "cc").map((r) => r.email);
  const subject = /^re:/i.test(item.subject) ? item.subject : `RE: ${item.subject}`;
  const result = await Office365OutlookService.SendEmailV2({
    To: toList.join(";"),
    Cc: cc.length ? cc.join(";") : undefined,
    Subject: subject,
    Body: `<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">${toHtml(text)}</div>`,
    Importance: "Normal",
  });
  if (!result.success) {
    console.error("Email send failed", result.error);
    throw new CommsError("send", "send-failed");
  }
  return "newEmail";
}

export async function sendTeamsReply(item: FeedItem, text: string): Promise<SendVia> {
  const { chatId, teamId, channelId, parentMessageId } = item.reply;
  const messageBody = toHtml(text);
  const result =
    teamId && channelId && parentMessageId
      ? await MicrosoftTeamsService.ReplyWithMessageToConversation(
          { recipient: { groupId: teamId, channelId }, parentMessageId, messageBody },
          "User",
          "Channel",
        )
      : chatId
        ? await MicrosoftTeamsService.PostMessageToConversation(
            { recipient: chatId, messageBody },
            "User",
            "Group chat",
          )
        : null;
  if (!result) throw new CommsError("not-found", "no-chat");
  if (!result.success) {
    console.error("Teams post failed", result.error);
    throw new CommsError("send", "send-failed");
  }
  return "teams";
}

// ---------------------------------------------------------------------------
// Formatting (your time zone; always 12-hour with AM/PM)
// ---------------------------------------------------------------------------

const dayKeyFormat = new Intl.DateTimeFormat("en-CA", {
  timeZone: TIME_ZONE,
  year: "numeric",
  month: "2-digit",
  day: "2-digit",
});

const dayLabelFormat = new Intl.DateTimeFormat("en-US", {
  timeZone: TIME_ZONE,
  weekday: "long",
  month: "long",
  day: "numeric",
  year: "numeric",
});

const timeFormat = new Intl.DateTimeFormat("en-US", {
  timeZone: TIME_ZONE,
  hour: "numeric",
  minute: "2-digit",
  hour12: true,
  timeZoneName: "short",
});

const shortTimeFormat = new Intl.DateTimeFormat("en-US", {
  timeZone: TIME_ZONE,
  hour: "numeric",
  minute: "2-digit",
  hour12: true,
});

export function dayKey(ms: number): string {
  return dayKeyFormat.format(ms);
}

export function dayLabel(ms: number): string {
  return dayLabelFormat.format(ms);
}

export function relativeDay(ms: number, now: number): string | null {
  const key = dayKey(ms);
  if (key === dayKey(now)) return "Today";
  if (key === dayKey(now - 24 * 60 * 60 * 1000)) return "Yesterday";
  return null;
}

export function timeLabel(ms: number): string {
  return timeFormat.format(ms);
}

export function shortTime(ms: number): string {
  return shortTimeFormat.format(ms);
}

const compactDayFormat = new Intl.DateTimeFormat("en-US", {
  timeZone: TIME_ZONE,
  weekday: "short",
  month: "short",
  day: "numeric",
});

const weekdayHourFormat = new Intl.DateTimeFormat("en-US", {
  timeZone: TIME_ZONE,
  weekday: "short",
  hour: "numeric",
  hourCycle: "h23",
});

/** "Thursday, October 15, 2026 · 9:48 AM EDT" */
export function fullStamp(ms: number): string {
  return `${dayLabelFormat.format(ms)} · ${timeFormat.format(ms)}`;
}

/** "Thu, Oct 15 · 9:48 AM EDT" — for narrow screens. */
export function compactStamp(ms: number): string {
  return `${compactDayFormat.format(ms)} · ${timeFormat.format(ms)}`;
}

export function relativeAge(ms: number, now: number): string {
  const minutes = Math.floor(Math.max(0, now - ms) / 60_000);
  if (minutes < 1) return "just now";
  if (minutes < 60) return `${minutes} min ago`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `${hours} hr ago`;
  const days = Math.floor(hours / 24);
  return days === 1 ? "1 day ago" : `${days} days ago`;
}

/** Weekdays 8:00 AM – 6:00 PM in your time zone, when checks should arrive steadily. */
export function isBusinessHours(ms: number): boolean {
  const parts = weekdayHourFormat.formatToParts(ms);
  const weekday = parts.find((p) => p.type === "weekday")?.value ?? "";
  const hour = Number(parts.find((p) => p.type === "hour")?.value ?? "0") % 24;
  return weekday !== "Sat" && weekday !== "Sun" && hour >= 8 && hour < 18;
}

export function stampLabel(iso: string): string {
  const ms = Date.parse(iso);
  if (Number.isNaN(ms)) return "";
  const today = dayKey(ms) === dayKey(Date.now());
  return today
    ? timeLabel(ms)
    : `${new Intl.DateTimeFormat("en-US", { timeZone: TIME_ZONE, month: "short", day: "numeric" }).format(ms)}, ${timeLabel(ms)}`;
}
