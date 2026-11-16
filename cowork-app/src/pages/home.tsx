import { useCallback, useContext, useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import { createPortal } from "react-dom";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import {
  CheckCircle2Icon,
  ChevronDownIcon,
  ExternalLinkIcon,
  InboxIcon,
  MailIcon,
  MessageSquareTextIcon,
  PlusIcon,
  RefreshCwIcon,
  RotateCcwIcon,
  SearchIcon,
  SendIcon,
  Trash2Icon,
  TriangleAlertIcon,
  UserPlusIcon,
} from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from "@/components/ui/collapsible";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Skeleton } from "@/components/ui/skeleton";
import { Spinner } from "@/components/ui/spinner";
import { Textarea } from "@/components/ui/textarea";
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle-group";
import { EmptyState, ErrorState } from "@/components/states";
import { cn } from "@/lib/utils";
import { safeUrl } from "@/lib/safe-url";
import { HeaderSlotContext } from "@/lib/header-slot";
import {
  CommsError,
  OWNER_EMAIL,
  cleanTaskText,
  compactStamp,
  dayKey,
  dayLabel,
  emptyState,
  fetchFeed,
  fetchState,
  fetchTasks,
  fullStamp,
  isBusinessHours,
  isExternal,
  isValidEmail,
  mutateState,
  mutateTasks,
  newTaskId,
  relativeAge,
  relativeDay,
  replyAllRecipients,
  replyCandidates,
  sendEmailAsNew,
  sendEmailReply,
  sendTeamsReply,
  stampLabel,
  timeLabel,
  type DashboardState,
  type ExtraRecipient,
  type FeedItem,
  type FeedSnapshot,
  type Person,
  type Priority,
  type RecipientLine,
  type RunEntry,
  type RunOutcome,
  type RunSource,
  type SendVia,
  type SentRecord,
  type StateSnapshot,
  type TaskItem,
  type TasksSnapshot,
} from "@/lib/comms";

type Filter = "all" | "open" | "completed";
type Tone = "success" | "warning" | "danger" | "info" | "neutral";
type Status = "open" | "replied" | "completed";

// One state-tone map, reused by badges and timeline dots.
const TONE_BADGE: Record<Tone, string> = {
  success: "bg-green-100 text-green-800 dark:bg-green-950 dark:text-green-200",
  warning: "bg-amber-100 text-amber-900 dark:bg-amber-950 dark:text-amber-200",
  danger: "bg-red-100 text-red-800 dark:bg-red-950 dark:text-red-200",
  info: "bg-sky-100 text-sky-800 dark:bg-sky-950 dark:text-sky-200",
  neutral: "bg-muted text-foreground",
};

const TONE_DOT: Record<Tone, string> = {
  success: "bg-green-700 dark:bg-green-400",
  warning: "bg-amber-600 dark:bg-amber-400",
  danger: "bg-red-700 dark:bg-red-400",
  info: "bg-sky-700 dark:bg-sky-400",
  neutral: "bg-zinc-500 dark:bg-zinc-400",
};

const PRIORITY: Record<Priority, { label: string; tone: Tone }> = {
  high: { label: "High priority", tone: "danger" },
  medium: { label: "Medium priority", tone: "warning" },
  low: { label: "Low priority", tone: "neutral" },
};

const STATUS: Record<Status, { label: string; tone: Tone }> = {
  open: { label: "Open", tone: "info" },
  replied: { label: "Replied", tone: "success" },
  completed: { label: "Completed", tone: "neutral" },
};

const OUTCOME: Record<RunOutcome, { label: string; tone: Tone }> = {
  added: { label: "Added", tone: "success" },
  skipped: { label: "Skipped", tone: "neutral" },
  error: { label: "Needs attention", tone: "danger" },
};

const SOURCE_LABEL: Record<RunSource, string> = {
  email: "Email",
  "teams-chat": "Teams chat",
  "teams-sweep": "Teams sweep",
  other: "Check",
};

const UNDO_SECONDS = 5;
const POLL_MS = 60_000;
// During weekday business hours a check normally lands at least every 30
// minutes (the Teams sweep), so two quiet hours means something is stuck.
const STALE_MS = 2 * 60 * 60 * 1000;

function cloneState(state: DashboardState): DashboardState {
  return {
    version: 1,
    completed: { ...state.completed },
    deleted: { ...state.deleted },
    sent: { ...state.sent },
    drafts: { ...state.drafts },
  };
}

function displayName(p: Person): string {
  return p.email && p.email.toLowerCase() === OWNER_EMAIL.toLowerCase() ? "You" : p.name;
}

function recipientsText(list: Person[], more: number): string {
  if (list.length === 0 && more === 0) return "You";
  const names = list.map(displayName).join(", ");
  if (more <= 0) return names;
  return `${names}${names ? ", " : ""}+${more} more`;
}

function viaLabel(via: SendVia): string {
  switch (via) {
    case "replyAll":
      return " · reply all";
    case "newEmail":
      return " · as a new email";
    case "teams":
      return " · in Teams";
    default:
      return " · reply";
  }
}

function sentLabel(record: SentRecord): string {
  const base = viaLabel(record.via);
  const added = record.via === "reply" || record.via === "newEmail" ? (record.added ?? []) : [];
  if (added.length === 0) return base;
  return added.length <= 2 ? `${base} + ${added.join(", ")}` : `${base} + ${added.length} others`;
}

export function HomePage() {
  const queryClient = useQueryClient();
  const headerSlot = useContext(HeaderSlotContext);
  const [filter, setFilter] = useState<Filter>("all");
  const [localSent, setLocalSent] = useState<Record<string, SentRecord>>({});
  const [pendingWrites, setPendingWrites] = useState(0);
  const writeChain = useRef<Promise<unknown>>(Promise.resolve());
  const pendingRef = useRef(0);
  const [pendingTaskWrites, setPendingTaskWrites] = useState(0);
  const taskChain = useRef<Promise<unknown>>(Promise.resolve());
  const pendingTaskRef = useRef(0);

  const feed = useQuery({
    queryKey: ["feed"],
    queryFn: fetchFeed,
    refetchInterval: POLL_MS,
    refetchOnWindowFocus: true,
    staleTime: 30_000,
  });
  const saved = useQuery({
    queryKey: ["state"],
    queryFn: fetchState,
    refetchInterval: pendingWrites > 0 ? false : POLL_MS,
    refetchOnWindowFocus: pendingWrites === 0,
    staleTime: 30_000,
  });
  const myTasks = useQuery({
    queryKey: ["tasks"],
    queryFn: fetchTasks,
    refetchInterval: pendingTaskWrites > 0 ? false : POLL_MS,
    refetchOnWindowFocus: pendingTaskWrites === 0,
    staleTime: 30_000,
  });

  const state = saved.data?.state ?? emptyState();

  // Tasks: optimistic change, then a serialized read-latest / apply / save of
  // the tasks file. Each change is pure, so applying it twice is safe.
  const applyTaskChange = useCallback(
    async (change: (list: TaskItem[]) => TaskItem[], failMessage: string): Promise<boolean> => {
      await queryClient.cancelQueries({ queryKey: ["tasks"] });
      queryClient.setQueryData<TasksSnapshot>(["tasks"], (prev) => {
        const base = prev ?? { tasks: [], fileId: null };
        return { tasks: change(base.tasks.map((t) => ({ ...t }))), fileId: base.fileId };
      });
      pendingTaskRef.current += 1;
      setPendingTaskWrites(pendingTaskRef.current);
      const run = taskChain.current.then(() => mutateTasks(change));
      taskChain.current = run.catch(() => undefined);
      return run
        .then((snapshot) => {
          pendingTaskRef.current -= 1;
          setPendingTaskWrites(pendingTaskRef.current);
          if (pendingTaskRef.current === 0) queryClient.setQueryData(["tasks"], snapshot);
          return true;
        })
        .catch((error: unknown) => {
          console.error("Task save failed", error);
          pendingTaskRef.current -= 1;
          setPendingTaskWrites(pendingTaskRef.current);
          toast.error(failMessage);
          void queryClient.invalidateQueries({ queryKey: ["tasks"] });
          return false;
        });
    },
    [queryClient],
  );

  const addTask = (text: string) => {
    const clean = cleanTaskText(text);
    if (!clean) return;
    const task: TaskItem = { id: newTaskId(), text: clean, createdAt: new Date().toISOString() };
    void applyTaskChange(
      (list) => (list.some((t) => t.id === task.id) ? list : [task, ...list]),
      "Couldn't save that task. Try again.",
    );
  };

  const toggleTask = (task: TaskItem, done: boolean) => {
    const at = new Date().toISOString();
    void applyTaskChange(
      (list) =>
        list.map((t) => {
          if (t.id !== task.id) return t;
          return done
            ? { ...t, completedAt: t.completedAt ?? at }
            : { id: t.id, text: t.text, createdAt: t.createdAt };
        }),
      done ? "Couldn't mark that task complete. Try again." : "Couldn't reopen that task. Try again.",
    );
  };

  const deleteTask = (task: TaskItem) => {
    const index = Math.max(0, (myTasks.data?.tasks ?? []).findIndex((t) => t.id === task.id));
    void applyTaskChange(
      (list) => list.filter((t) => t.id !== task.id),
      "Couldn't delete that task. Try again.",
    );
    toast("Task deleted", {
      description: task.text,
      action: {
        label: "Undo",
        onClick: () =>
          void applyTaskChange(
            (list) =>
              list.some((t) => t.id === task.id)
                ? list
                : [...list.slice(0, index), task, ...list.slice(index)],
            "Couldn't restore that task. Try again.",
          ),
      },
    });
  };

  // Optimistic change, then a serialized read-latest / apply / save on OneDrive.
  const applyChange = useCallback(
    (change: (s: DashboardState) => void, failMessage: string): Promise<boolean> => {
      queryClient.setQueryData<StateSnapshot>(["state"], (prev) => {
        const base = prev ?? { state: emptyState(), fileId: null };
        const next = cloneState(base.state);
        change(next);
        return { state: next, fileId: base.fileId };
      });
      pendingRef.current += 1;
      setPendingWrites(pendingRef.current);
      const run = writeChain.current.then(() => mutateState(change));
      writeChain.current = run.catch(() => undefined);
      return run
        .then((snapshot) => {
          pendingRef.current -= 1;
          setPendingWrites(pendingRef.current);
          if (pendingRef.current === 0) queryClient.setQueryData(["state"], snapshot);
          return true;
        })
        .catch((error: unknown) => {
          console.error("Dashboard save failed", error);
          pendingRef.current -= 1;
          setPendingWrites(pendingRef.current);
          toast.error(failMessage);
          void queryClient.invalidateQueries({ queryKey: ["state"] });
          return false;
        });
    },
    [queryClient],
  );

  const statusOf = useCallback(
    (item: FeedItem): Status => {
      if (state.completed[item.id]) return "completed";
      if (state.sent[item.id] || localSent[item.id]) return "replied";
      return "open";
    },
    [state, localSent],
  );

  const live = useMemo(
    () => (feed.data?.items ?? []).filter((item) => !state.deleted[item.id]),
    [feed.data, state.deleted],
  );

  const counts = useMemo(() => {
    let completed = 0;
    for (const item of live) if (statusOf(item) === "completed") completed += 1;
    return { all: live.length, open: live.length - completed, completed };
  }, [live, statusOf]);

  const groups = useMemo(() => {
    const visible = live.filter((item) => {
      const status = statusOf(item);
      if (filter === "open") return status !== "completed";
      if (filter === "completed") return status === "completed";
      return true;
    });
    const out: { key: string; ms: number; items: FeedItem[] }[] = [];
    for (const item of visible) {
      const key = dayKey(item.receivedMs);
      const last = out[out.length - 1];
      if (last && last.key === key) last.items.push(item);
      else out.push({ key, ms: item.receivedMs, items: [item] });
    }
    return out;
  }, [live, filter, statusOf]);

  const markComplete = (item: FeedItem) =>
    void applyChange((s) => {
      s.completed[item.id] = new Date().toISOString();
    }, "Couldn't mark that complete. Try again.");

  const reopen = (item: FeedItem) =>
    void applyChange((s) => {
      delete s.completed[item.id];
    }, "Couldn't reopen that item. Try again.");

  const remove = (item: FeedItem) => {
    void applyChange((s) => {
      s.deleted[item.id] = new Date().toISOString();
    }, "Couldn't delete that item. Try again.");
    toast("Removed from your dashboard", {
      description: item.subject,
      action: {
        label: "Undo",
        onClick: () =>
          void applyChange((s) => {
            delete s.deleted[item.id];
          }, "Couldn't restore that item. Try again."),
      },
    });
  };

  const saveDraft = (item: FeedItem, text: string) =>
    void applyChange((s) => {
      s.drafts[item.id] = text;
    }, "Couldn't save your edit. Copy it before leaving.");

  const recordSent = (item: FeedItem, record: SentRecord) => {
    setLocalSent((prev) => ({ ...prev, [item.id]: record }));
    void applyChange((s) => {
      s.sent[item.id] = record;
    }, "Reply sent, but its status couldn't be saved.");
  };

  const fetching = feed.isFetching || saved.isFetching || myTasks.isFetching;
  const refresh = () => {
    void feed.refetch();
    void saved.refetch();
    if (pendingTaskRef.current === 0) void myTasks.refetch();
  };
  const now = Date.now();
  const lastRun = feed.data?.lastRun ?? null;
  const stale = lastRun !== null && isBusinessHours(now) && now - lastRun.atMs > STALE_MS;

  return (
    <div className="mx-auto w-full max-w-5xl px-4 pb-20 pt-6 sm:px-6">
      {headerSlot && feed.data && createPortal(<HeaderStamp lastRun={lastRun} stale={stale} />, headerSlot)}
      <TasksPanel
        tasks={myTasks.data?.tasks}
        loading={myTasks.isPending}
        failed={myTasks.isError}
        onRetry={() => void myTasks.refetch()}
        onAdd={addTask}
        onToggle={toggleTask}
        onDelete={deleteTask}
      />
      {feed.data && <CheckLog snapshot={feed.data} now={now} stale={stale} />}
      <div className="flex flex-wrap items-center justify-between gap-3">
        <ToggleGroup
          type="single"
          variant="outline"
          value={filter}
          onValueChange={(value) => value && setFilter(value as Filter)}
          aria-label="Show items"
        >
          <ToggleGroupItem value="all" className="gap-2">
            All <span className="tabular-nums">{counts.all}</span>
          </ToggleGroupItem>
          <ToggleGroupItem value="open" className="gap-2">
            Open <span className="tabular-nums">{counts.open}</span>
          </ToggleGroupItem>
          <ToggleGroupItem value="completed" className="gap-2">
            Completed <span className="tabular-nums">{counts.completed}</span>
          </ToggleGroupItem>
        </ToggleGroup>
        <div className="flex items-center gap-3">
          {feed.dataUpdatedAt > 0 && (
            <span className="text-sm text-muted-foreground" title="When this screen last pulled your feed">
              Synced {timeLabel(feed.dataUpdatedAt)}
            </span>
          )}
          <Button variant="outline" size="sm" onClick={refresh} disabled={fetching}>
            {fetching ? <Spinner aria-label="Refreshing" /> : <RefreshCwIcon aria-hidden="true" />}
            Refresh
          </Button>
        </div>
      </div>

      {saved.isError && (
        <div
          role="alert"
          className="mt-4 flex flex-wrap items-center gap-2 rounded-md border border-amber-300 bg-amber-50 p-3 text-sm text-amber-900 dark:border-amber-800 dark:bg-amber-950 dark:text-amber-100"
        >
          <TriangleAlertIcon className="size-4 shrink-0" aria-hidden="true" />
          <span>Saved statuses couldn't be loaded, so completed and deleted items may reappear.</span>
          <Button variant="outline" size="sm" onClick={() => void saved.refetch()}>
            Retry
          </Button>
        </div>
      )}

      <section id="feed" tabIndex={-1} aria-label="Messages" className="mt-6 outline-none">
        {feed.isPending ? (
          <FeedSkeleton />
        ) : feed.isError ? (
          <ErrorState
            title="Couldn't load your messages"
            description="Check your connection, then try again."
            action={
              <Button variant="outline" onClick={() => void feed.refetch()}>
                Try again
              </Button>
            }
          />
        ) : groups.length === 0 ? (
          <EmptyState
            icon={InboxIcon}
            title={
              filter === "completed"
                ? "No completed items"
                : filter === "open"
                  ? "Nothing open"
                  : "Nothing needs your attention"
            }
            description={
              filter === "all"
                ? "Emails and Teams chats that meet your COS criteria will appear here."
                : undefined
            }
          />
        ) : (
          groups.map((group) => {
            const rel = relativeDay(group.ms, now);
            const headingId = `day-${group.key}`;
            return (
              <section key={group.key} aria-labelledby={headingId} className="mb-8">
                <div className="sticky top-14 z-20 -mx-2 mb-2 flex flex-wrap items-baseline gap-x-3 bg-background/95 px-2 py-2 backdrop-blur supports-[backdrop-filter]:bg-background/85">
                  <h2 id={headingId} className="text-lg font-semibold tracking-tight">
                    {rel ? `${rel} · ` : ""}
                    {dayLabel(group.ms)}
                  </h2>
                  <span className="text-sm text-muted-foreground">
                    {group.items.length} {group.items.length === 1 ? "item" : "items"}
                  </span>
                </div>
                <ol>
                  {group.items.map((item) => (
                    <CommItem
                      key={item.id}
                      item={item}
                      status={statusOf(item)}
                      sent={state.sent[item.id] ?? localSent[item.id]}
                      savedDraft={state.drafts[item.id]}
                      onComplete={() => markComplete(item)}
                      onReopen={() => reopen(item)}
                      onDelete={() => remove(item)}
                      onSaveDraft={(text) => saveDraft(item, text)}
                      onSent={(record) => recordSent(item, record)}
                    />
                  ))}
                </ol>
              </section>
            );
          })
        )}
      </section>
    </div>
  );
}

type Phase =
  | { kind: "idle" }
  | { kind: "countdown"; mode: "reply" | "newEmail" }
  | { kind: "sending" }
  | { kind: "error" }
  | { kind: "missing" };

function CommItem({
  item,
  status,
  sent,
  savedDraft,
  onComplete,
  onReopen,
  onDelete,
  onSaveDraft,
  onSent,
}: {
  item: FeedItem;
  status: Status;
  sent: SentRecord | undefined;
  savedDraft: string | undefined;
  onComplete: () => void;
  onReopen: () => void;
  onDelete: () => void;
  onSaveDraft: (text: string) => void;
  onSent: (record: SentRecord) => void;
}) {
  const domId = useMemo(() => item.id.replace(/[^a-zA-Z0-9_-]/g, "").slice(-40), [item.id]);
  const baseline = savedDraft ?? item.draft;
  const [text, setText] = useState(baseline);
  const dirty = useRef(false);
  const [replyAll, setReplyAll] = useState(item.reply.replyAll ?? false);
  const [phase, setPhase] = useState<Phase>({ kind: "idle" });
  const [secondsLeft, setSecondsLeft] = useState(UNDO_SECONDS);
  const timer = useRef<number | null>(null);
  const textAtSend = useRef("");
  // Sender-only replies: people you add from the thread (or by address),
  // keyed by lower-cased email, each placed on To or Cc.
  const [manual, setManual] = useState<ExtraRecipient[]>([]);
  const [chosen, setChosen] = useState<Record<string, RecipientLine>>({});
  const [addValue, setAddValue] = useState("");
  const [addError, setAddError] = useState<string | null>(null);
  const extrasAtSend = useRef<ExtraRecipient[]>([]);

  useEffect(() => {
    if (!dirty.current) setText(baseline);
  }, [baseline]);

  useEffect(
    () => () => {
      if (timer.current !== null) window.clearInterval(timer.current);
    },
    [],
  );

  const busy = phase.kind === "countdown" || phase.kind === "sending";
  const statusMeta = STATUS[status];
  const priorityMeta = item.priority ? PRIORITY[item.priority] : null;
  const link = safeUrl(item.webLink);
  const allRecipients = replyAllRecipients(item);
  const candidates = useMemo(() => replyCandidates(item), [item]);
  const people = useMemo(() => {
    const known = new Set(candidates.map((p) => p.email.toLowerCase()));
    return [...candidates, ...manual.filter((p) => !known.has(p.email.toLowerCase()))];
  }, [candidates, manual]);
  const senderKey = item.from.email.trim().toLowerCase();
  const extras: ExtraRecipient[] =
    item.channel !== "email" || replyAll
      ? []
      : people.flatMap((p) => {
          const line = chosen[p.email.toLowerCase()];
          return line ? [{ ...p, line }] : [];
        });

  const audience =
    item.channel === "teams"
      ? item.reply.audience || "This Teams chat"
      : replyAll && allRecipients.length > 1
        ? `Reply all · ${item.from.name} + ${allRecipients.length - 1} ${allRecipients.length - 1 === 1 ? "other" : "others"}`
        : extras.length > 0
          ? `Reply to ${item.from.name} + ${extras.length} ${extras.length === 1 ? "other" : "others"}`
          : `Reply to ${item.from.name}`;

  const togglePerson = (person: ExtraRecipient, on: boolean) =>
    setChosen((prev) => {
      const key = person.email.toLowerCase();
      const next = { ...prev };
      if (on) next[key] = prev[key] ?? person.line;
      else delete next[key];
      return next;
    });

  const setLine = (person: ExtraRecipient, line: RecipientLine) =>
    setChosen((prev) => ({ ...prev, [person.email.toLowerCase()]: line }));

  const addPerson = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const email = addValue.trim();
    const key = email.toLowerCase();
    if (!isValidEmail(email)) {
      setAddError("Enter a valid email address.");
      return;
    }
    if (key === OWNER_EMAIL.toLowerCase()) {
      setAddError("That's your own address.");
      return;
    }
    if (key === senderKey) {
      setAddError(`${item.from.name} is already on this reply.`);
      return;
    }
    const existing = people.find((p) => p.email.toLowerCase() === key);
    if (!existing) setManual((prev) => [...prev, { name: email, email, line: "to" }]);
    setChosen((prev) => ({ ...prev, [key]: prev[key] ?? existing?.line ?? "to" }));
    setAddValue("");
    setAddError(null);
  };

  const doSend = async (mode: "reply" | "newEmail") => {
    setPhase({ kind: "sending" });
    const body = textAtSend.current;
    const added = extrasAtSend.current;
    try {
      let via: SendVia;
      if (item.channel === "teams") via = await sendTeamsReply(item, body);
      else if (mode === "newEmail") via = await sendEmailAsNew(item, body, replyAll, added);
      else via = await sendEmailReply(item, body, replyAll, added);
      const names = via === "reply" || via === "newEmail" ? added.map((p) => p.name) : [];
      onSent({
        at: new Date().toISOString(),
        text: body,
        via,
        ...(names.length > 0 ? { added: names } : {}),
      });
      setPhase({ kind: "idle" });
      toast.success(item.channel === "teams" ? "Posted in Teams" : "Reply sent");
    } catch (error) {
      if (error instanceof CommsError && error.kind === "not-found" && item.channel === "email") {
        setPhase({ kind: "missing" });
      } else {
        setPhase({ kind: "error" });
      }
    }
  };

  const startSend = (mode: "reply" | "newEmail") => {
    const body = text.trim();
    if (!body || busy) return;
    textAtSend.current = body;
    extrasAtSend.current = extras;
    if (dirty.current) {
      dirty.current = false;
      onSaveDraft(text);
    }
    let left = UNDO_SECONDS;
    setSecondsLeft(left);
    setPhase({ kind: "countdown", mode });
    timer.current = window.setInterval(() => {
      left -= 1;
      setSecondsLeft(left);
      if (left <= 0) {
        if (timer.current !== null) window.clearInterval(timer.current);
        timer.current = null;
        void doSend(mode);
      }
    }, 1000);
  };

  const undoSend = () => {
    if (timer.current !== null) window.clearInterval(timer.current);
    timer.current = null;
    setPhase({ kind: "idle" });
  };

  const commitDraft = () => {
    if (!dirty.current) return;
    dirty.current = false;
    if (text !== baseline) onSaveDraft(text);
  };

  const draftId = `draft-${domId}`;
  const replyAllId = `replyall-${domId}`;
  const pickId = `pick-${domId}`;

  const draftPanel = (
    <section aria-label={sent ? "Sent reply" : "Draft reply"} className="mt-4 rounded-md border bg-background p-3 sm:p-4">
      <div className="flex flex-wrap items-baseline justify-between gap-x-3 gap-y-1">
        <Label htmlFor={draftId} className="text-sm font-semibold">
          {sent ? "Sent reply" : "Draft reply"}
        </Label>
        <span className="text-sm text-muted-foreground">{audience}</span>
      </div>
      <div className="mt-2 flex flex-col gap-3 md:flex-row md:items-start">
        <Textarea
          id={draftId}
          value={sent ? sent.text : text}
          readOnly={Boolean(sent) || busy}
          onChange={(event) => {
            dirty.current = true;
            setText(event.target.value);
          }}
          onBlur={commitDraft}
          className="min-h-32 flex-1 text-sm leading-relaxed"
        />
        {!sent && (
          <div className="flex flex-row flex-wrap items-center gap-3 md:w-36 md:flex-col md:items-stretch">
            {phase.kind === "countdown" ? (
              <Button variant="outline" onClick={undoSend}>
                <RotateCcwIcon aria-hidden="true" />
                Undo ({secondsLeft})
              </Button>
            ) : (
              <Button
                onClick={() => startSend("reply")}
                disabled={!text.trim() || phase.kind === "sending"}
              >
                {phase.kind === "sending" ? (
                  <Spinner aria-label="Sending" />
                ) : (
                  <SendIcon aria-hidden="true" />
                )}
                Send
              </Button>
            )}
            {item.channel === "email" && (
              <div className="flex items-center gap-2">
                <Checkbox
                  id={replyAllId}
                  checked={replyAll}
                  disabled={busy}
                  onCheckedChange={(value) => setReplyAll(value === true)}
                />
                <Label htmlFor={replyAllId} className="text-sm font-normal">
                  Reply all
                </Label>
              </div>
            )}
          </div>
        )}
      </div>
      {item.channel === "email" && !sent && !replyAll && (
        <Collapsible className="mt-3">
          <CollapsibleTrigger asChild>
            <Button variant="ghost" size="sm" className="group -ml-2">
              <UserPlusIcon aria-hidden="true" />
              Add people
              {extras.length > 0 && (
                <span className="tabular-nums text-muted-foreground">· {extras.length} added</span>
              )}
              <ChevronDownIcon
                aria-hidden="true"
                className="transition-transform group-data-[state=open]:rotate-180"
              />
            </Button>
          </CollapsibleTrigger>
          <CollapsibleContent>
            <fieldset disabled={busy} className="mt-2 min-w-0 rounded-md border px-3 pb-3 pt-1">
              <legend className="px-1 text-sm font-semibold">Recipients</legend>
              <ul className="divide-y">
                <li className="flex flex-wrap items-center gap-x-3 gap-y-1 py-2">
                  <Checkbox id={`${pickId}-sender`} checked disabled />
                  <Label htmlFor={`${pickId}-sender`} className="min-w-0 flex-1 font-normal break-words">
                    <span className="font-medium">{item.from.name}</span>
                    {item.from.email && item.from.email !== item.from.name && (
                      <span className="text-muted-foreground"> · {item.from.email}</span>
                    )}
                  </Label>
                  {isExternal(item.from.email) && (
                    <Badge className={TONE_BADGE.warning}>External</Badge>
                  )}
                  <Badge variant="outline">Sender</Badge>
                </li>
                {people.map((person, index) => {
                  const key = person.email.toLowerCase();
                  const line = chosen[key];
                  const boxId = `${pickId}-${index}`;
                  return (
                    <li key={key} className="flex flex-wrap items-center gap-x-3 gap-y-1 py-2">
                      <Checkbox
                        id={boxId}
                        checked={Boolean(line)}
                        onCheckedChange={(value) => togglePerson(person, value === true)}
                      />
                      <Label htmlFor={boxId} className="min-w-0 flex-1 font-normal break-words">
                        <span className="font-medium">{person.name}</span>
                        {person.email !== person.name && (
                          <span className="text-muted-foreground"> · {person.email}</span>
                        )}
                      </Label>
                      {isExternal(person.email) && (
                        <Badge className={TONE_BADGE.warning}>External</Badge>
                      )}
                      {line && (
                        <ToggleGroup
                          type="single"
                          variant="outline"
                          size="sm"
                          value={line}
                          onValueChange={(value) => value && setLine(person, value as RecipientLine)}
                          aria-label={`${person.name}: To or Cc`}
                        >
                          <ToggleGroupItem value="to">To</ToggleGroupItem>
                          <ToggleGroupItem value="cc">Cc</ToggleGroupItem>
                        </ToggleGroup>
                      )}
                    </li>
                  );
                })}
              </ul>
              {item.toMore > 0 && (
                <p className="mt-1 text-sm text-muted-foreground">
                  {item.toMore === 1
                    ? "1 more recipient on the original isn't listed."
                    : `${item.toMore} more recipients on the original aren't listed.`}
                </p>
              )}
              <form noValidate onSubmit={addPerson} className="mt-3 flex flex-wrap items-end gap-2">
                <div className="flex min-w-0 flex-1 flex-col gap-1.5 sm:max-w-sm">
                  <Label htmlFor={`${pickId}-add`}>Add by email</Label>
                  <Input
                    id={`${pickId}-add`}
                    type="email"
                    inputMode="email"
                    autoComplete="off"
                    spellCheck={false}
                    value={addValue}
                    onChange={(event) => {
                      setAddValue(event.target.value);
                      if (addError) setAddError(null);
                    }}
                    aria-invalid={addError ? true : undefined}
                    aria-describedby={addError ? `${pickId}-add-error` : undefined}
                  />
                </div>
                <Button type="submit" variant="outline" disabled={!addValue.trim()}>
                  <PlusIcon aria-hidden="true" />
                  Add
                </Button>
              </form>
              {addError && (
                <p
                  id={`${pickId}-add-error`}
                  role="alert"
                  className="mt-1 text-sm font-medium text-red-700 dark:text-red-400"
                >
                  {addError}
                </p>
              )}
            </fieldset>
          </CollapsibleContent>
        </Collapsible>
      )}
      <div aria-live="polite" className="mt-2 text-sm empty:hidden">
        {sent ? (
          <span className="inline-flex items-center gap-1.5 font-medium text-green-800 dark:text-green-300">
            <CheckCircle2Icon className="size-4" aria-hidden="true" />
            Sent {stampLabel(sent.at)}
            {sentLabel(sent)}
          </span>
        ) : phase.kind === "countdown" ? (
          <span>Sending in {secondsLeft}s</span>
        ) : phase.kind === "sending" ? (
          <span>Sending…</span>
        ) : phase.kind === "error" ? (
          <span className="font-medium text-red-700 dark:text-red-400">
            Couldn't send. Your draft is unchanged — try again.
          </span>
        ) : phase.kind === "missing" ? (
          <span className="flex flex-wrap items-center gap-2 font-medium text-red-700 dark:text-red-400">
            The original email was moved or deleted, so it can't be answered in-thread.
            <Button variant="outline" size="sm" onClick={() => startSend("newEmail")}>
              <MailIcon aria-hidden="true" />
              Send as new email
            </Button>
          </span>
        ) : null}
      </div>
    </section>
  );

  return (
    <li className="relative grid grid-cols-1 pb-6 sm:grid-cols-[6.5rem_1fr] sm:gap-4">
      <div className="hidden pt-4 text-right sm:block">
        <time dateTime={item.receivedAt} className="text-sm font-semibold tabular-nums">
          {timeLabel(item.receivedMs)}
        </time>
      </div>
      <div className="relative sm:pl-6">
        <span aria-hidden="true" className="absolute inset-y-0 left-0 hidden w-px bg-border sm:block" />
        <span
          aria-hidden="true"
          className={cn(
            "absolute -left-[6px] top-5 hidden size-3 rounded-full ring-4 ring-background sm:block forced-colors:bg-[CanvasText]",
            TONE_DOT[statusMeta.tone],
          )}
        />
        <article
          aria-labelledby={`subject-${domId}`}
          className={cn(
            "rounded-lg border bg-card p-4 text-card-foreground shadow-xs sm:p-5",
            status === "completed" && "bg-muted/40",
          )}
        >
          <div className="flex flex-wrap items-center gap-2">
            <Badge variant="outline" className="gap-1">
              {item.channel === "email" ? (
                <MailIcon aria-hidden="true" />
              ) : (
                <MessageSquareTextIcon aria-hidden="true" />
              )}
              {item.channel === "email" ? "Email" : "Teams"}
            </Badge>
            {priorityMeta && (
              <Badge className={TONE_BADGE[priorityMeta.tone]}>{priorityMeta.label}</Badge>
            )}
            <Badge className={TONE_BADGE[statusMeta.tone]}>
              {status === "completed" && <CheckCircle2Icon aria-hidden="true" />}
              {statusMeta.label}
            </Badge>
            <time dateTime={item.receivedAt} className="text-sm text-muted-foreground sm:hidden">
              {timeLabel(item.receivedMs)}
            </time>
          </div>

          {item.whyItMatters && (
            <p className="mt-3 text-sm font-semibold">{item.whyItMatters}</p>
          )}
          <h3 id={`subject-${domId}`} className="mt-1 text-base font-semibold leading-snug break-words">
            {item.subject}
          </h3>

          <dl className="mt-3 grid grid-cols-[2.5rem_1fr] gap-x-3 gap-y-1 text-sm">
            <dt className="text-muted-foreground">From</dt>
            <dd className="min-w-0 break-words">
              <span className="font-medium">{item.from.name}</span>
              {item.from.email && item.from.email !== item.from.name && (
                <span className="text-muted-foreground"> · {item.from.email}</span>
              )}
            </dd>
            <dt className="text-muted-foreground">To</dt>
            <dd className="min-w-0 break-words">{recipientsText(item.to, item.toMore)}</dd>
            {item.cc.length > 0 && (
              <>
                <dt className="text-muted-foreground">Cc</dt>
                <dd className="min-w-0 break-words">{recipientsText(item.cc, 0)}</dd>
              </>
            )}
          </dl>

          {item.summary && <p className="mt-3 text-sm leading-relaxed">{item.summary}</p>}

          {item.research && (
            <Collapsible className="mt-3">
              <CollapsibleTrigger asChild>
                <Button variant="ghost" size="sm" className="group -ml-2">
                  <SearchIcon aria-hidden="true" />
                  Research notes
                  {item.research.sources.length > 0 && (
                    <span className="tabular-nums text-muted-foreground">
                      · {item.research.sources.length} {item.research.sources.length === 1 ? "source" : "sources"}
                    </span>
                  )}
                  <ChevronDownIcon
                    aria-hidden="true"
                    className="transition-transform group-data-[state=open]:rotate-180"
                  />
                </Button>
              </CollapsibleTrigger>
              <CollapsibleContent>
                <div className="mt-2 rounded-md border border-dashed bg-muted/60 p-3 text-sm leading-normal">
                  {item.research.notes && <p>{item.research.notes}</p>}
                  {item.research.sources.length > 0 && (
                    <ul aria-label="Sources" className="mt-2 space-y-1 text-muted-foreground">
                      {item.research.sources.map((source, index) => {
                        const href = safeUrl(source.url);
                        return (
                          <li key={`${source.title}-${index}`} className="break-words">
                            {"• "}
                            {href ? (
                              <a
                                href={href}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-primary underline-offset-4 hover:underline"
                              >
                                {source.title}
                              </a>
                            ) : (
                              source.title
                            )}
                          </li>
                        );
                      })}
                    </ul>
                  )}
                  <p className="mt-2 text-xs text-muted-foreground">Only you see these notes. They're never sent.</p>
                </div>
              </CollapsibleContent>
            </Collapsible>
          )}

          {status === "completed" ? (
            <Collapsible className="mt-3">
              <CollapsibleTrigger asChild>
                <Button variant="ghost" size="sm" className="group -ml-2">
                  {sent ? "Show sent reply" : "Show draft"}
                  <ChevronDownIcon
                    aria-hidden="true"
                    className="transition-transform group-data-[state=open]:rotate-180"
                  />
                </Button>
              </CollapsibleTrigger>
              <CollapsibleContent>{draftPanel}</CollapsibleContent>
            </Collapsible>
          ) : (
            draftPanel
          )}

          <div className="mt-4 flex flex-wrap items-center gap-2 border-t pt-3">
            {status === "completed" ? (
              <Button variant="outline" size="sm" onClick={onReopen}>
                <RotateCcwIcon aria-hidden="true" />
                Reopen
              </Button>
            ) : (
              <Button variant="outline" size="sm" onClick={onComplete}>
                <CheckCircle2Icon aria-hidden="true" />
                Mark Complete
              </Button>
            )}
            <Button
              variant="outline"
              size="sm"
              onClick={onDelete}
              className="text-red-700 hover:text-red-800 dark:text-red-400 dark:hover:text-red-300"
            >
              <Trash2Icon aria-hidden="true" />
              Delete
            </Button>
            {link && (
              <a
                href={link}
                target="_blank"
                rel="noopener noreferrer"
                className="ml-auto inline-flex items-center gap-1 rounded-sm text-sm font-medium text-primary underline-offset-4 hover:underline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring"
              >
                Open in {item.channel === "email" ? "Outlook" : "Teams"}
                <ExternalLinkIcon className="size-3.5" aria-hidden="true" />
              </a>
            )}
          </div>
        </article>
      </div>
    </li>
  );
}

function freshnessDot(lastRun: RunEntry | null, stale: boolean): string {
  if (!lastRun) return TONE_DOT.neutral;
  return stale ? TONE_DOT.warning : TONE_DOT.success;
}

function HeaderStamp({ lastRun, stale }: { lastRun: RunEntry | null; stale: boolean }) {
  return (
    <div className="flex items-center gap-2 text-right">
      <span aria-hidden="true" className={cn("size-2 shrink-0 rounded-full", freshnessDot(lastRun, stale))} />
      <p className="leading-tight">
        <span className="block text-[11px] text-muted-foreground sm:text-xs">Last refreshed</span>
        {lastRun ? (
          <time dateTime={lastRun.at} className="block whitespace-nowrap text-xs font-medium tabular-nums sm:text-sm">
            <span className="sm:hidden">{compactStamp(lastRun.atMs)}</span>
            <span className="hidden sm:inline">{fullStamp(lastRun.atMs)}</span>
          </time>
        ) : (
          <span className="block text-xs font-medium sm:text-sm">Not recorded yet</span>
        )}
      </p>
    </div>
  );
}

function RunLine({ run }: { run: RunEntry }) {
  const outcome = OUTCOME[run.outcome];
  const detail = [run.from, run.note].filter(Boolean).join(" · ");
  return (
    <>
      <div className="flex flex-wrap items-center gap-x-2 gap-y-1">
        <Badge className={TONE_BADGE[outcome.tone]}>{outcome.label}</Badge>
        <span className="text-muted-foreground">{SOURCE_LABEL[run.source]}</span>
        <span className="min-w-0 break-words font-medium">{run.subject || "(no subject)"}</span>
      </div>
      {detail && <p className="mt-0.5 break-words text-muted-foreground">{detail}</p>}
    </>
  );
}

function CheckLog({ snapshot, now, stale }: { snapshot: FeedSnapshot; now: number; stale: boolean }) {
  const [open, setOpen] = useState(false);
  const { runs, lastRun } = snapshot;
  return (
    <section aria-labelledby="latest-check" className="mb-6 rounded-lg border bg-card text-card-foreground">
      <Collapsible open={open} onOpenChange={setOpen}>
        <div className="flex flex-wrap items-start gap-x-3 gap-y-2 p-3 text-sm sm:p-4">
          <span aria-hidden="true" className={cn("mt-1.5 size-2.5 shrink-0 rounded-full", freshnessDot(lastRun, stale))} />
          <div className="min-w-0 flex-1">
            <h2 id="latest-check" className="mb-1 font-semibold">
              Latest check
              {lastRun && (
                <span className="font-normal text-muted-foreground">
                  {" · "}
                  <time dateTime={lastRun.at}>{stampLabel(lastRun.at)}</time>
                  {` (${relativeAge(lastRun.atMs, now)})`}
                </span>
              )}
            </h2>
            {lastRun ? (
              <RunLine run={lastRun} />
            ) : (
              <p className="text-muted-foreground">
                No checks recorded yet. Every email and Teams check will show here, even when nothing needs you.
              </p>
            )}
          </div>
          {runs.length > 0 && (
            <CollapsibleTrigger asChild>
              <Button variant="outline" size="sm" aria-controls="recent-checks">
                Recent checks <span className="tabular-nums">{runs.length}</span>
                <ChevronDownIcon aria-hidden="true" className={cn("transition-transform", open && "rotate-180")} />
              </Button>
            </CollapsibleTrigger>
          )}
        </div>
        {stale && lastRun && (
          <div
            role="status"
            className="mx-3 mb-3 flex items-start gap-2 rounded-md border border-amber-300 bg-amber-50 p-3 text-sm text-amber-900 sm:mx-4 sm:mb-4 dark:border-amber-800 dark:bg-amber-950 dark:text-amber-100"
          >
            <TriangleAlertIcon className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
            <span>
              No checks recorded since {stampLabel(lastRun.at)}. Your COS checks may be paused or waiting for your
              approval in Cowork.
            </span>
          </div>
        )}
        <CollapsibleContent id="recent-checks">
          <ol aria-label="Recent checks, newest first" className="divide-y border-t">
            {runs.map((run, index) => (
              <li
                key={`${run.at}-${index}`}
                className="grid gap-1 px-3 py-2.5 text-sm sm:grid-cols-[10rem_1fr] sm:gap-3 sm:px-4"
              >
                <time dateTime={run.at} className="tabular-nums text-muted-foreground">
                  {stampLabel(run.at)}
                </time>
                <div className="min-w-0">
                  <RunLine run={run} />
                </div>
              </li>
            ))}
          </ol>
        </CollapsibleContent>
      </Collapsible>
    </section>
  );
}

function TasksPanel({
  tasks,
  loading,
  failed,
  onRetry,
  onAdd,
  onToggle,
  onDelete,
}: {
  tasks: TaskItem[] | undefined;
  loading: boolean;
  failed: boolean;
  onRetry: () => void;
  onAdd: (text: string) => void;
  onToggle: (task: TaskItem, done: boolean) => void;
  onDelete: (task: TaskItem) => void;
}) {
  const [text, setText] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);
  const ready = tasks !== undefined;
  const list = tasks ?? [];
  const open = list.filter((t) => !t.completedAt);
  const done = list
    .filter((t) => t.completedAt)
    .sort((a, b) => (Date.parse(b.completedAt ?? "") || 0) - (Date.parse(a.completedAt ?? "") || 0));

  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const clean = cleanTaskText(text);
    if (!clean || !ready) return;
    onAdd(clean);
    setText("");
    inputRef.current?.focus();
  };

  return (
    <section
      aria-labelledby="tasks-heading"
      className="mb-6 rounded-lg border bg-card p-4 text-card-foreground shadow-xs sm:p-5"
    >
      <div className="flex flex-wrap items-baseline justify-between gap-x-3 gap-y-1">
        <h2 id="tasks-heading" className="text-lg font-semibold tracking-tight">
          My tasks
        </h2>
        {list.length > 0 && (
          <span className="text-sm tabular-nums text-muted-foreground">
            {open.length} open
            {done.length > 0 ? ` · ${done.length} completed` : ""}
          </span>
        )}
      </div>

      <form onSubmit={submit} className="mt-3 flex gap-2">
        <Label htmlFor="new-task" className="sr-only">
          New task
        </Label>
        <Input
          id="new-task"
          ref={inputRef}
          value={text}
          onChange={(event) => setText(event.target.value)}
          placeholder="Add a task"
          maxLength={500}
          autoComplete="off"
          disabled={!ready}
        />
        <Button type="submit" disabled={!ready || !cleanTaskText(text)}>
          <PlusIcon aria-hidden="true" />
          Add
        </Button>
      </form>

      {loading && !ready ? (
        <div role="status" aria-label="Loading tasks" className="mt-4 space-y-3">
          <Skeleton className="h-5 w-2/3" />
          <Skeleton className="h-5 w-1/2" />
        </div>
      ) : failed && !ready ? (
        <div role="alert" className="mt-3 flex flex-wrap items-center gap-2 text-sm">
          <TriangleAlertIcon className="size-4 shrink-0 text-red-700 dark:text-red-400" aria-hidden="true" />
          <span>Couldn't load your tasks.</span>
          <Button variant="outline" size="sm" onClick={onRetry}>
            Retry
          </Button>
        </div>
      ) : list.length === 0 ? (
        <p className="mt-3 text-sm text-muted-foreground">No tasks yet.</p>
      ) : (
        <>
          {open.length > 0 ? (
            <ul aria-label="Open tasks" className="mt-2 divide-y">
              {open.map((task) => (
                <TaskRow key={task.id} task={task} onToggle={onToggle} onDelete={onDelete} />
              ))}
            </ul>
          ) : (
            <p className="mt-3 text-sm text-muted-foreground">All tasks completed.</p>
          )}
          {done.length > 0 && (
            <Collapsible className="mt-2 border-t pt-2">
              <CollapsibleTrigger asChild>
                <Button variant="ghost" size="sm" className="group -ml-2">
                  Completed <span className="tabular-nums">{done.length}</span>
                  <ChevronDownIcon
                    aria-hidden="true"
                    className="transition-transform group-data-[state=open]:rotate-180"
                  />
                </Button>
              </CollapsibleTrigger>
              <CollapsibleContent>
                <ul aria-label="Completed tasks" className="divide-y">
                  {done.map((task) => (
                    <TaskRow key={task.id} task={task} onToggle={onToggle} onDelete={onDelete} />
                  ))}
                </ul>
              </CollapsibleContent>
            </Collapsible>
          )}
        </>
      )}
    </section>
  );
}

function TaskRow({
  task,
  onToggle,
  onDelete,
}: {
  task: TaskItem;
  onToggle: (task: TaskItem, done: boolean) => void;
  onDelete: (task: TaskItem) => void;
}) {
  const done = Boolean(task.completedAt);
  const boxId = `task-${task.id.replace(/[^a-zA-Z0-9_-]/g, "")}`;
  const stamp = done
    ? `Completed ${stampLabel(task.completedAt ?? "")}`
    : task.createdAt
      ? `Added ${stampLabel(task.createdAt)}`
      : "";
  return (
    <li className="flex items-start gap-3 py-2.5">
      <Checkbox
        id={boxId}
        checked={done}
        onCheckedChange={(value) => onToggle(task, value === true)}
        className="mt-0.5"
      />
      <div className="min-w-0 flex-1">
        <Label
          htmlFor={boxId}
          className={cn(
            "block text-sm font-normal leading-snug break-words",
            done && "text-muted-foreground line-through",
          )}
        >
          {task.text}
        </Label>
        {stamp && <p className="mt-0.5 text-xs text-muted-foreground">{stamp}</p>}
      </div>
      <Button
        variant="ghost"
        size="icon-sm"
        aria-label={`Delete task: ${task.text}`}
        title="Delete task"
        onClick={() => onDelete(task)}
        className="-my-1 shrink-0 text-red-700 hover:text-red-800 dark:text-red-400 dark:hover:text-red-300"
      >
        <Trash2Icon aria-hidden="true" />
      </Button>
    </li>
  );
}

function FeedSkeleton() {
  return (
    <div role="status" aria-label="Loading messages" className="space-y-6">
      <Skeleton className="h-6 w-72" />
      {[0, 1, 2].map((row) => (
        <div key={row} className="grid grid-cols-1 gap-4 sm:grid-cols-[6.5rem_1fr]">
          <Skeleton className="hidden h-5 w-20 justify-self-end sm:block" />
          <div className="space-y-3 rounded-lg border p-5">
            <div className="flex gap-2">
              <Skeleton className="h-5 w-16 rounded-full" />
              <Skeleton className="h-5 w-24 rounded-full" />
              <Skeleton className="h-5 w-14 rounded-full" />
            </div>
            <Skeleton className="h-5 w-2/3" />
            <Skeleton className="h-4 w-1/2" />
            <Skeleton className="h-24 w-full" />
          </div>
        </div>
      ))}
    </div>
  );
}
