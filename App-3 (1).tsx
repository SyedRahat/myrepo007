// App.tsx - Welcome Desk: employee onboarding (Power Apps code app)
//
// UI ONLY: no SharePoint and no flows. The 8 sample joiners live in the app's memory.
// Your changes stay until you refresh the page. After a refresh, the samples come back.
// The joiner fields match the columns of your "Onboarding" SharePoint list
// (Title, PersonalEmail, Department, JobTitle, JoiningDate, ManagerEmail, Stage, WelcomeEmailSent, Notes),
// so the app can be connected to that list later.
//
// How it adapts to the screen:
//   Desktop: a sidebar, and the joiner list with the chosen joiner's checklist beside it
//   Tablet:  a slim rail, and the checklist slides in from the right
//   Phone:   big buttons at the bottom, the checklist fills the screen, one-tap "Done" on every card
//
// The file has 5 parts:
//   PART 1: Types and settings        PART 2: Sample joiners        PART 3: Helper functions
//   PART 4: Small building blocks     PART 5: The app (2 screens, the side panel, the popup and the undo bar)
import { useEffect, useRef, useState, type ReactNode } from "react";
import "./index.css";

// ---------- PART 1: Types and settings ----------

type Team = "IT" | "HR" | "Manager";
type TaskId = "laptop" | "email" | "access" | "offer" | "idproof" | "bank" | "plan" | "buddy" | "welcome";
type Stage = "IT Setup" | "HR Documents" | "Manager Welcome" | "Completed"; // the same 4 stages as before
type Status = "late" | "risk" | "track" | "ready";
type Screen = "runway" | "tasks";
type ScreenSize = "phone" | "tablet" | "desktop";
type Show = "all" | "attention" | "track" | "ready";

// One new joiner. The comment after each field is its SharePoint column.
type Joiner = {
  id: string; // ID
  name: string; // Title
  personalEmail: string; // PersonalEmail
  department: string; // Department
  jobTitle: string; // JobTitle
  joiningDate: string; // JoiningDate, like 2026-12-01
  managerEmail: string; // ManagerEmail
  notes: string; // Notes
  done: Partial<Record<TaskId, number>>; // the tasks that are done, and when (a new column, like Tasks, when you connect)
};

// One task on the checklist
type TaskInfo = { id: TaskId; team: Team; name: string; hint: string };

// The form fields (like the controls on a canvas screen)
type JoinerForm = { name: string; personalEmail: string; department: string; jobTitle: string; joiningDate: string; managerEmail: string; notes: string };
type Errors = Partial<Record<keyof JoinerForm, string>>;

// What the side panel shows (null = the panel is closed)
type Panel = { kind: "form" } | { kind: "details"; id: string } | null;
type Popup = { id: number; tone: Status; icon: IconName; title: string; lines: string[] } | null;
// The bar at the bottom after a tick, with an Undo button
type Toast = { id: number; text: string; joinerId: string; task: TaskId; before: number | undefined } | null;

const TEAMS: { key: Team; name: string; short: string; css: string; does: string }[] = [
  { key: "IT", name: "IT team", short: "IT", css: "it", does: "laptop, work email and system access" },
  { key: "HR", name: "HR team", short: "HR", css: "hr", does: "offer letter and documents" },
  { key: "Manager", name: "Manager", short: "manager", css: "mgr", does: "first-day plan, buddy and welcome email" },
];

// The 9 tasks every joiner needs before Day 1: 3 for each team
const TASKS: TaskInfo[] = [
  { id: "laptop", team: "IT", name: "Laptop ready", hint: "Set up, tagged and charged" },
  { id: "email", team: "IT", name: "Work email created", hint: "Login details sent to the manager" },
  { id: "access", team: "IT", name: "System access given", hint: "VPN, Teams and the HR portal" },
  { id: "offer", team: "HR", name: "Offer letter signed", hint: "Signed copy received" },
  { id: "idproof", team: "HR", name: "ID proof received", hint: "A copy of an ID card or passport" },
  { id: "bank", team: "HR", name: "Bank details received", hint: "Needed for the first salary" },
  { id: "plan", team: "Manager", name: "First-day plan shared", hint: "Who to meet and where to sit" },
  { id: "buddy", team: "Manager", name: "Buddy assigned", hint: "A colleague who helps in week 1" },
  { id: "welcome", team: "Manager", name: "Welcome email sent", hint: "Sent to the personal email" },
];

const STAGE_OF_TEAM: Record<Team, Stage> = { IT: "IT Setup", HR: "HR Documents", Manager: "Manager Welcome" };

// The 4 labels, and the rule behind each one
const STATUS_INFO: Record<Status, { name: string; rule: string }> = {
  late: { name: "Overdue", rule: "The joining day has come and tasks are still open." },
  risk: { name: "At risk", rule: "Joins in 3 days or less and tasks are still open." },
  track: { name: "On track", rule: "Joins in more than 3 days. There is still time." },
  ready: { name: "Ready", rule: "Every task is done." },
};
const STATUS_ORDER: Status[] = ["late", "risk", "track", "ready"];

// The buttons above the list that filter it
const SHOWS: { key: Show; label: string; css: string }[] = [
  { key: "all", label: "All", css: "all" },
  { key: "attention", label: "Need attention", css: "late" },
  { key: "track", label: "On track", css: "track" },
  { key: "ready", label: "Ready", css: "ready" },
];

const DEPARTMENTS = ["HR", "IT", "Finance", "Sales", "Operations"];
const RISK_DAYS = 3; // "At risk" when the joining day is this close and tasks are open
const RECENT_DAYS = 14; // joiners stay on the list for 14 days after they join
const NOTES_MAX = 300;
const POPUP_SECONDS = 4; // the popup closes by itself after this many seconds
const TOAST_SECONDS = 6; // the undo bar hides itself after this many seconds
const DAY_MS = 24 * 60 * 60 * 1000;
const HOUR_MS = 60 * 60 * 1000;
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

const EMPTY_FORM: JoinerForm = { name: "", personalEmail: "", department: "", jobTitle: "", joiningDate: "", managerEmail: "", notes: "" };

// The name of each field, in screen order (used in the red error list)
const FIELD_LABELS: Record<keyof JoinerForm, string> = {
  name: "Full name",
  personalEmail: "Personal email",
  department: "Department",
  joiningDate: "Joining date",
  jobTitle: "Job title",
  managerEmail: "Manager email",
  notes: "Notes",
};

// ---------- PART 2: Sample joiners ----------

// The moment the app opened. The sample dates are counted from today, so they always look current.
const OPENED = Date.now();
const TODAY = dateValue(OPENED);
const ALL_TASKS: TaskId[] = TASKS.map((t) => t.id);

// Marks some tasks as done, a few hours apart
function doneAt(hoursAgo: number, ids: TaskId[]): Partial<Record<TaskId, number>> {
  const done: Partial<Record<TaskId, number>> = {};
  ids.forEach((id, i) => {
    done[id] = OPENED - (hoursAgo + i * 5) * HOUR_MS;
  });
  return done;
}

// workdayFrom(TODAY, 2) = 2 working days from today. A minus number = in the past.
const SAMPLE_JOINERS: Joiner[] = [
  {
    id: "j1",
    name: "Ravi Kumar",
    personalEmail: "ravi.kumar@gmail.com",
    department: "IT",
    jobTitle: "Power Platform Developer",
    joiningDate: workdayFrom(TODAY, 2),
    managerEmail: "priya.sharma@yourcompany.com",
    notes: "Needs a laptop with 16 GB RAM.",
    done: doneAt(4, ["laptop", "email", "offer", "idproof"]),
  },
  {
    id: "j2",
    name: "Ananya Iyer",
    personalEmail: "ananya.iyer@outlook.com",
    department: "Finance",
    jobTitle: "Accounts Executive",
    joiningDate: workdayFrom(TODAY, 0),
    managerEmail: "suresh.nair@yourcompany.com",
    notes: "",
    done: doneAt(20, ALL_TASKS),
  },
  {
    id: "j3",
    name: "Mohammed Imran",
    personalEmail: "imran.mohd@gmail.com",
    department: "Sales",
    jobTitle: "Sales Manager",
    joiningDate: workdayFrom(TODAY, 6),
    managerEmail: "kavya.reddy@yourcompany.com",
    notes: "Works from the Hyderabad office.",
    done: doneAt(30, ["offer", "idproof", "bank", "laptop"]),
  },
  {
    id: "j4",
    name: "Sneha Patil",
    personalEmail: "sneha.patil@yahoo.com",
    department: "HR",
    jobTitle: "HR Executive",
    joiningDate: workdayFrom(TODAY, 9),
    managerEmail: "priya.sharma@yourcompany.com",
    notes: "",
    done: doneAt(50, ["offer"]),
  },
  {
    id: "j5",
    name: "Arjun Reddy",
    personalEmail: "arjun.reddy@gmail.com",
    department: "Operations",
    jobTitle: "Operations Analyst",
    joiningDate: workdayFrom(TODAY, 3),
    managerEmail: "vikram.rao@yourcompany.com",
    notes: "",
    done: doneAt(8, ["laptop", "email", "access", "offer", "idproof", "bank", "plan"]),
  },
  {
    id: "j6",
    name: "Fatima Khan",
    personalEmail: "fatima.khan@gmail.com",
    department: "IT",
    jobTitle: "QA Engineer",
    joiningDate: workdayFrom(TODAY, -2),
    managerEmail: "priya.sharma@yourcompany.com",
    notes: "Buddy not chosen yet.",
    done: doneAt(60, ALL_TASKS.filter((id) => id !== "buddy")),
  },
  {
    id: "j7",
    name: "Kiran Rao",
    personalEmail: "kiran.rao@outlook.com",
    department: "Finance",
    jobTitle: "Financial Analyst",
    joiningDate: workdayFrom(TODAY, 15),
    managerEmail: "suresh.nair@yourcompany.com",
    notes: "",
    done: {},
  },
  {
    id: "j8",
    name: "Lakshmi Menon",
    personalEmail: "lakshmi.menon@gmail.com",
    department: "Operations",
    jobTitle: "Supply Chain Lead",
    joiningDate: workdayFrom(TODAY, -4),
    managerEmail: "vikram.rao@yourcompany.com",
    notes: "",
    done: doneAt(150, ALL_TASKS),
  },
];

// On a desktop, the first joiner who needs attention is shown when the app opens
const START_ID =
  [...SAMPLE_JOINERS].sort(byJoiningDate).find((j) => statusOf(j, TODAY) === "late" || statusOf(j, TODAY) === "risk")?.id ?? SAMPLE_JOINERS[0].id;

// ---------- PART 3: Helper functions ----------

// The time right now, in milliseconds. Only the buttons and the clock call it, never the drawing of the screen.
function timeNow(): number {
  return Date.now();
}

// Makes an id for a new joiner, like "j1731830000000-1"
let lastId = 0;
function newId(prefix: string): string {
  lastId += 1;
  return `${prefix}${timeNow()}-${lastId}`;
}

function pad(n: number): string {
  return String(n).padStart(2, "0");
}

// A time as a date, like 2026-11-17 (in your own time zone)
function dateValue(ms: number): string {
  const d = new Date(ms);
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

// 2026-11-17 as a number of milliseconds at midnight UTC, so counting days is always exact
function utcDay(value: string): number {
  const [y, m, d] = value.split("-").map(Number);
  return Date.UTC(y, m - 1, d);
}

// How many days from today to the date (0 = today, a minus number = in the past)
function daysUntil(value: string, today: string): number {
  return Math.round((utcDay(value) - utcDay(today)) / DAY_MS);
}

function weekdayName(value: string): string {
  return WEEKDAYS[new Date(utcDay(value)).getUTCDay()];
}

// True when the date is a Saturday or a Sunday
function isWeekend(value: string): boolean {
  const day = new Date(utcDay(value)).getUTCDay();
  return day === 0 || day === 6;
}

// Shows 2026-11-19 as 19-11-2026
function formatDate(value: string): string {
  const [y, m, d] = value.split("-");
  return `${d}-${m}-${y}`;
}

// Shows 2026-11-19 as "Thu 19 Nov"
function shortDate(value: string): string {
  const d = new Date(utcDay(value));
  return `${WEEKDAYS[d.getUTCDay()].slice(0, 3)} ${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]}`;
}

// 14:05 becomes "2:05 PM"
function formatTime(ms: number): string {
  const d = new Date(ms);
  const h = d.getHours();
  return `${h % 12 === 0 ? 12 : h % 12}:${pad(d.getMinutes())} ${h < 12 ? "AM" : "PM"}`;
}

// A working day n days away from `start`. Saturdays and Sundays are skipped.
function workdayFrom(start: string, n: number): string {
  const step = n >= 0 ? 1 : -1;
  const weekend = (ms: number) => {
    const day = new Date(ms).getUTCDay();
    return day === 0 || day === 6;
  };
  let day = utcDay(start);
  while (weekend(day)) day += step * DAY_MS;
  let left = Math.abs(n);
  while (left > 0) {
    day += step * DAY_MS;
    if (!weekend(day)) left -= 1;
  }
  const d = new Date(day);
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
}

// "Joins today", "Joins in 5 days" or "Joined 2 days ago"
function joinText(value: string, today: string): string {
  const d = daysUntil(value, today);
  if (d === 0) return "Joins today";
  if (d === 1) return "Joins tomorrow";
  if (d > 1) return `Joins in ${d} days`;
  if (d === -1) return "Joined yesterday";
  return `Joined ${-d} days ago`;
}

// The countdown on the left of each card: a big number and a small word
function countdown(value: string, today: string): { big: string; small: string } {
  const d = daysUntil(value, today);
  if (d === 0) return { big: "Today", small: "" };
  if (d > 0) return { big: String(d), small: d === 1 ? "day" : "days" };
  return { big: String(-d), small: d === -1 ? "day ago" : "days ago" };
}

// "Done today, 10:42 AM", "Done yesterday, 4:05 PM" or "Done on 12-11-2026"
function doneText(ms: number, today: string): string {
  const day = dateValue(ms);
  if (day === today) return `Done today, ${formatTime(ms)}`;
  if (daysUntil(day, today) === -1) return `Done yesterday, ${formatTime(ms)}`;
  return `Done on ${formatDate(day)}`;
}

// "Ravi Kumar" becomes "RK"
function initials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  return (parts[0][0] + (parts.length > 1 ? parts[parts.length - 1][0] : "")).toUpperCase();
}

// A simple email check: has an @ and a dot, and no spaces
function looksLikeEmail(value: string): boolean {
  const v = value.trim();
  return v.includes("@") && v.includes(".") && !v.includes(" ") && v.indexOf("@") > 0;
}

// Earliest joining date first
function byJoiningDate(a: Joiner, b: Joiner): number {
  return a.joiningDate.localeCompare(b.joiningDate) || a.name.localeCompare(b.name);
}

function doneCount(j: Joiner): number {
  return TASKS.filter((t) => j.done[t.id]).length;
}

// The first task that is still open (IT first, then HR, then the manager)
function nextTask(j: Joiner): TaskInfo | undefined {
  return TASKS.find((t) => !j.done[t.id]);
}

// The label of a joiner is worked out, never stored (like a formula in a canvas app)
function statusOf(j: Joiner, today: string): Status {
  if (TASKS.every((t) => j.done[t.id])) return "ready";
  const d = daysUntil(j.joiningDate, today);
  if (d <= 0) return "late";
  if (d <= RISK_DAYS) return "risk";
  return "track";
}

// The stage for your SharePoint "Stage" column: the first team that still has open tasks
function stageOf(j: Joiner): Stage {
  const team = TEAMS.find((tm) => TASKS.some((t) => t.team === tm.key && !j.done[t.id]));
  return team ? STAGE_OF_TEAM[team.key] : "Completed";
}

// Checks every field (like If(IsBlank(...)) checks in a canvas app)
function validate(f: JoinerForm, joiners: Joiner[], editingId: string, originalDate: string, today: string): Errors {
  const e: Errors = {};
  const email = f.personalEmail.trim().toLowerCase();
  if (f.name.trim().length < 2) e.name = "Enter the new joiner's full name.";
  if (!looksLikeEmail(f.personalEmail)) e.personalEmail = "Enter a valid email, like ravi.kumar@gmail.com.";
  else if (joiners.some((j) => j.id !== editingId && j.personalEmail.toLowerCase() === email)) e.personalEmail = "Another joiner already has this email.";
  if (!f.department) e.department = "Choose a department.";
  if (!f.joiningDate) e.joiningDate = "Choose the joining date.";
  else if (isWeekend(f.joiningDate)) e.joiningDate = `This is a ${weekdayName(f.joiningDate)}. Choose a day from Monday to Friday.`;
  else if (f.joiningDate < today && f.joiningDate !== originalDate) e.joiningDate = "Choose today or a later date.";
  if (!f.jobTitle.trim()) e.jobTitle = "Enter the job title.";
  if (!looksLikeEmail(f.managerEmail)) e.managerEmail = "Enter a valid manager email, like priya.sharma@yourcompany.com.";
  if (f.notes.length > NOTES_MAX) e.notes = `Keep notes under ${NOTES_MAX} letters.`;
  return e;
}

// Works out the layout from the screen width (like App.SizeBreakpoints in a canvas app)
function getScreenSize(): ScreenSize {
  const width = typeof window === "undefined" ? 1280 : window.innerWidth;
  if (width < 600) return "phone";
  if (width < 1024) return "tablet";
  return "desktop";
}

// Checks the layout again when the window is resized or the phone is turned
function useScreenSize(): ScreenSize {
  const [size, setSize] = useState<ScreenSize>(getScreenSize);
  useEffect(() => {
    const onResize = () => setSize(getScreenSize());
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);
  return size;
}

// ---------- PART 4: Small building blocks ----------

// Simple line icons, drawn with SVG (no icon library needed)
type IconName = "runway" | "tasks" | "plus" | "close" | "search" | "check" | "mail" | "edit" | "trash" | "badge" | "info";
const ICONS: Record<IconName, ReactNode> = {
  runway: (
    <>
      <path d="M6 3v18" />
      <circle cx="6" cy="7" r="2" />
      <circle cx="6" cy="17" r="2" />
      <path d="M11 7h9M11 12h6M11 17h9" />
    </>
  ),
  tasks: (
    <>
      <rect x="3" y="4" width="18" height="16" rx="2" />
      <path d="m7 9 2 2 3-3M7 15l2 2 3-3M15 10h3M15 16h3" />
    </>
  ),
  plus: <path d="M12 5v14M5 12h14" />,
  close: <path d="M6 6l12 12M18 6 6 18" />,
  search: (
    <>
      <circle cx="11" cy="11" r="7" />
      <path d="m20 20-4-4" />
    </>
  ),
  check: <path d="m5 12.5 4.5 4.5L19 7.5" />,
  mail: (
    <>
      <rect x="3" y="5" width="18" height="14" rx="2" />
      <path d="m3 7 9 6 9-6" />
    </>
  ),
  edit: (
    <>
      <path d="M4 20h4L19 9l-4-4L4 16z" />
      <path d="m13.5 6.5 4 4" />
    </>
  ),
  trash: <path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3" />,
  badge: (
    <>
      <rect x="5" y="3" width="14" height="18" rx="2.5" />
      <path d="M10 3v3h4V3" />
      <circle cx="12" cy="11" r="2.5" />
      <path d="M8.5 17.5c.8-1.6 2-2.4 3.5-2.4s2.7.8 3.5 2.4" />
    </>
  ),
  info: (
    <>
      <circle cx="12" cy="12" r="9" />
      <path d="M12 11v6M12 7.5h.01" />
    </>
  ),
};

function Icon({ name, size = 20 }: { name: IconName; size?: number }) {
  return (
    <svg className="icon" width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
      {ICONS[name]}
    </svg>
  );
}

// The readiness bar: 9 blocks, 3 for each team. A filled block = a task that is done.
function Strip({ joiner, big = false }: { joiner: Joiner; big?: boolean }) {
  return (
    <span className={big ? "strip big" : "strip"} aria-hidden="true">
      {TEAMS.map((team) => (
        <span key={team.key} className={`strip-team ${team.css}`}>
          {TASKS.filter((t) => t.team === team.key).map((t) => (
            <i key={t.id} className={joiner.done[t.id] ? "on" : undefined} />
          ))}
        </span>
      ))}
    </span>
  );
}

function StatusPill({ status }: { status: Status }) {
  return <span className={`pill ${status}`}>{STATUS_INFO[status].name}</span>;
}

// One form field: the label, the box, then a hint or the red error
function Field({ id, label, error, hint, children }: { id: string; label: string; error?: string; hint?: string; children: ReactNode }) {
  return (
    <div className="field">
      <label htmlFor={id}>{label}</label>
      {children}
      {hint && !error && <span className="hint">{hint}</span>}
      {error && (
        <span className="error" id={`${id}-error`}>
          {error}
        </span>
      )}
    </div>
  );
}

// ---------- PART 5: The app ----------

export default function App() {
  const size = useScreenSize();
  const [screen, setScreen] = useState<Screen>("runway"); // like Navigate() between screens
  const [now, setNow] = useState(OPENED); // the clock, updated every minute (like a Timer control)
  const [joiners, setJoiners] = useState<Joiner[]>(SAMPLE_JOINERS);
  const [selectedId, setSelectedId] = useState(START_ID); // desktop: the joiner shown beside the list
  const [search, setSearch] = useState("");
  const [show, setShow] = useState<Show>("all");
  const [showJoined, setShowJoined] = useState(false); // joiners who joined and are ready are folded away
  const [team, setTeam] = useState<Team>("IT"); // the team on the Tasks screen
  const [showDone, setShowDone] = useState(false);
  const [panel, setPanel] = useState<Panel>(null);
  const [form, setForm] = useState<JoinerForm>(EMPTY_FORM);
  const [editingId, setEditingId] = useState(""); // "" = adding a new joiner
  const [backToDetails, setBackToDetails] = useState(false); // after Save, show the joiner's checklist again
  const [errors, setErrors] = useState<Errors>({});
  const [triedSave, setTriedSave] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [popup, setPopup] = useState<Popup>(null);
  const [toast, setToast] = useState<Toast>(null);
  const pageRef = useRef<HTMLDivElement>(null);
  const panelRef = useRef<HTMLElement>(null);
  const paneRef = useRef<HTMLElement>(null);
  const popupRef = useRef<HTMLDivElement>(null);

  // The clock: every minute the countdowns update by themselves (also after midnight)
  useEffect(() => {
    const timer = window.setInterval(() => setNow(timeNow()), 60 * 1000);
    return () => window.clearInterval(timer);
  }, []);

  // While the side panel or the popup is open, lock the page behind it.
  // When it closes, unlock the page and put the focus back on the button you clicked.
  const layerOpen = panel !== null || popup !== null;
  useEffect(() => {
    if (!layerOpen) return;
    const opener = document.activeElement as HTMLElement | null;
    const page = pageRef.current;
    page?.setAttribute("inert", ""); // nothing behind can be clicked or typed in
    document.body.style.overflow = "hidden";
    return () => {
      page?.removeAttribute("inert");
      document.body.style.overflow = "";
      if (opener && document.body.contains(opener)) opener.focus();
      else document.getElementById("screen-title")?.focus();
    };
  }, [layerOpen]);

  // The side panel: move the focus into it, and close it with the Esc key
  const panelKind = panel?.kind ?? "";
  useEffect(() => {
    if (!panelKind) return;
    if (panelKind === "form") document.getElementById("field-name")?.focus();
    else panelRef.current?.focus();
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        setPanel(null);
        setConfirmDelete(false);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [panelKind]);

  // The popup: move the focus into it, then close it after a few seconds
  useEffect(() => {
    if (!popup) return;
    popupRef.current?.focus();
    const timer = window.setTimeout(() => setPopup(null), POPUP_SECONDS * 1000);
    return () => window.clearTimeout(timer);
  }, [popup]);

  // Desktop: when another joiner is chosen, show the top of their checklist
  useEffect(() => {
    paneRef.current?.querySelector(".d-body")?.scrollTo({ top: 0 });
  }, [selectedId]);

  // The undo bar hides itself after a few seconds
  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(null), TOAST_SECONDS * 1000);
    return () => window.clearTimeout(timer);
  }, [toast]);

  // ----- Values that update by themselves (like formulas in a canvas app) -----
  const today = dateValue(now);
  const statusFor = (j: Joiner) => statusOf(j, today);
  const needsAttention = (j: Joiner) => statusFor(j) === "late" || statusFor(j) === "risk";
  const sorted = [...joiners].sort(byJoiningDate);
  // Joiners who joined more than 14 days ago and are ready leave the list
  const onList = sorted.filter((j) => daysUntil(j.joiningDate, today) >= -RECENT_DAYS || statusFor(j) !== "ready");
  const counts: Record<Show, number> = {
    all: onList.length,
    attention: onList.filter(needsAttention).length,
    track: onList.filter((j) => statusFor(j) === "track").length,
    ready: onList.filter((j) => statusFor(j) === "ready").length,
  };
  const words = search.trim().toLowerCase();
  const shown = onList.filter(
    (j) =>
      (show === "all" || (show === "attention" ? needsAttention(j) : statusFor(j) === show)) &&
      (!words || `${j.name} ${j.jobTitle} ${j.department} ${j.personalEmail} ${j.managerEmail}`.toLowerCase().includes(words))
  );
  const joined = shown.filter((j) => daysUntil(j.joiningDate, today) < 0);
  const joinedOpen = joined.filter((j) => statusFor(j) !== "ready");
  const joinedReady = joined.filter((j) => statusFor(j) === "ready");
  const joiningToday = shown.filter((j) => daysUntil(j.joiningDate, today) === 0);
  const nextWeek = shown.filter((j) => daysUntil(j.joiningDate, today) >= 1 && daysUntil(j.joiningDate, today) <= 7);
  const later = shown.filter((j) => daysUntil(j.joiningDate, today) > 7);
  const selected = shown.find((j) => j.id === selectedId) ?? shown[0];
  const openTasks = (tm: Team) => joiners.reduce((sum, j) => sum + TASKS.filter((t) => t.team === tm && !j.done[t.id]).length, 0);
  const openTotal = TEAMS.reduce((sum, tm) => sum + openTasks(tm.key), 0);

  // ----- Actions (like the OnSelect of buttons) -----

  function showPopup(tone: Status, icon: IconName, title: string, lines: string[]) {
    setPopup({ id: timeNow(), tone, icon, title, lines });
  }

  function closePanel() {
    setPanel(null);
    setConfirmDelete(false);
  }

  function go(to: Screen) {
    setScreen(to);
    window.scrollTo(0, 0);
  }

  // Shows a joiner's checklist: beside the list on a desktop, in the side panel everywhere else
  function openDetails(j: Joiner) {
    setSelectedId(j.id);
    setConfirmDelete(false);
    if (!(size === "desktop" && screen === "runway")) setPanel({ kind: "details", id: j.id });
  }

  // UPDATE: ticks a task, or takes the tick away (like Patch(Onboarding, ThisItem, {...}))
  function toggleTask(j: Joiner, task: TaskId) {
    const before = j.done[task];
    const done = { ...j.done };
    if (before) delete done[task];
    else done[task] = timeNow();
    setJoiners((list) => list.map((x) => (x.id === j.id ? { ...x, done } : x)));
    const info = TASKS.find((t) => t.id === task);
    const nowReady = !before && TASKS.every((t) => done[t.id]);
    if (nowReady) {
      setToast(null);
      showPopup("ready", "check", `${j.name} is ready for Day 1`, ["Every task is done.", `${joinText(j.joiningDate, today)}, ${weekdayName(j.joiningDate)} ${formatDate(j.joiningDate)}.`]);
    } else {
      setToast({ id: timeNow(), joinerId: j.id, task, before, text: `${info?.name ?? "Task"}: ${before ? "not done" : "done"} for ${j.name}.` });
    }
  }

  // Puts the task back the way it was before the last tick
  function undo() {
    if (!toast) return;
    setJoiners((list) =>
      list.map((x) => {
        if (x.id !== toast.joinerId) return x;
        const done = { ...x.done };
        if (toast.before) done[toast.task] = toast.before;
        else delete done[toast.task];
        return { ...x, done };
      })
    );
    setToast(null);
  }

  // UI only: no real email is sent. When you connect your welcome email flow, call it here.
  function sendWelcome(j: Joiner) {
    const done = { ...j.done, welcome: j.done.welcome ?? timeNow() };
    setJoiners((list) => list.map((x) => (x.id === j.id ? { ...x, done } : x)));
    setToast(null);
    const ready = TASKS.every((t) => done[t.id]);
    showPopup(ready ? "ready" : "track", "mail", "Welcome email sent", [`To ${j.personalEmail}`, ready ? `Every task is done. ${j.name} is ready for Day 1.` : `${joinText(j.joiningDate, today)}.`]);
  }

  // Opens the empty form
  function openAddForm() {
    setForm(EMPTY_FORM);
    setEditingId("");
    setBackToDetails(false);
    setErrors({});
    setTriedSave(false);
    setConfirmDelete(false);
    setPanel({ kind: "form" });
  }

  // Puts a joiner's details in the form, to change them
  function startEdit(j: Joiner, fromPanel: boolean) {
    setForm({ name: j.name, personalEmail: j.personalEmail, department: j.department, jobTitle: j.jobTitle, joiningDate: j.joiningDate, managerEmail: j.managerEmail, notes: j.notes });
    setEditingId(j.id);
    setBackToDetails(fromPanel);
    setErrors({});
    setTriedSave(false);
    setConfirmDelete(false);
    setPanel({ kind: "form" });
  }

  // Changes one form field. After a failed Save, it checks the form again at once.
  function update<K extends keyof JoinerForm>(field: K, value: JoinerForm[K]) {
    const next: JoinerForm = { ...form, [field]: value };
    setForm(next);
    if (triedSave) setErrors(validate(next, joiners, editingId, joiners.find((j) => j.id === editingId)?.joiningDate ?? "", today));
    else if (field === "joiningDate" && typeof value === "string" && value && isWeekend(value)) setErrors((e) => ({ ...e, joiningDate: `This is a ${weekdayName(value)}. Choose a day from Monday to Friday.` }));
    else setErrors((e) => ({ ...e, [field]: undefined }));
  }

  function closeForm() {
    if (backToDetails && editingId) setPanel({ kind: "details", id: editingId });
    else closePanel();
  }

  // CREATE or UPDATE: saves the form (like Patch(Onboarding, Defaults(Onboarding) or ThisItem, {...}))
  function saveJoiner() {
    const original = joiners.find((j) => j.id === editingId);
    const found = validate(form, joiners, editingId, original?.joiningDate ?? "", today);
    setErrors(found);
    const wrong = (Object.keys(FIELD_LABELS) as (keyof JoinerForm)[]).filter((key) => found[key]);
    if (wrong.length > 0) {
      setTriedSave(true);
      document.getElementById(`field-${wrong[0]}`)?.focus();
      return;
    }
    const clean = {
      name: form.name.trim(),
      personalEmail: form.personalEmail.trim(),
      department: form.department,
      jobTitle: form.jobTitle.trim(),
      joiningDate: form.joiningDate,
      managerEmail: form.managerEmail.trim(),
      notes: form.notes.trim(),
    };
    if (original) {
      setJoiners((list) => list.map((j) => (j.id === original.id ? { ...j, ...clean } : j)));
      if (backToDetails) setPanel({ kind: "details", id: original.id });
      else closePanel();
      showPopup("track", "check", "Changes saved", [`${clean.name}'s details are updated.`]);
    } else {
      const id = newId("j");
      setJoiners((list) => [...list, { id, ...clean, done: {} }]);
      setSelectedId(id);
      setShow("all");
      setSearch("");
      closePanel();
      showPopup("track", "check", "Joiner added", [
        `${clean.name} joins on ${weekdayName(clean.joiningDate)} ${formatDate(clean.joiningDate)}.`,
        `${TASKS.length} tasks were added: 3 each for IT, HR and the manager.`,
      ]);
    }
  }

  // DELETE: removes a joiner and their tasks (like Remove(Onboarding, ThisItem))
  function deleteJoiner(j: Joiner) {
    setJoiners((list) => list.filter((x) => x.id !== j.id));
    if (toast?.joinerId === j.id) setToast(null);
    closePanel();
    showPopup("track", "trash", "Joiner deleted", [`${j.name} was deleted, with all their tasks.`]);
  }

  // ----- Pieces of screen used in several places -----

  const brand = (
    <div className="brand">
      <span className="brand-mark" aria-hidden="true">
        <Icon name="badge" size={24} />
      </span>
      <span className="brand-text">
        <strong>Welcome Desk</strong>
        <span>Employee onboarding</span>
      </span>
    </div>
  );

  function screenHead(title: string, text: string, actions?: ReactNode) {
    return (
      <header className="screen-head">
        <div>
          <h1 id="screen-title" tabIndex={-1}>
            {title}
          </h1>
          <p>{text}</p>
        </div>
        {actions && <div className="head-actions">{actions}</div>}
      </header>
    );
  }

  // The red list at the top of the form after a failed save
  function errorList() {
    const wrong = (Object.keys(FIELD_LABELS) as (keyof JoinerForm)[]).filter((key) => errors[key]);
    if (!triedSave || wrong.length === 0) return null;
    return (
      <p className="notice" role="alert">
        Please fix {wrong.length} {wrong.length === 1 ? "field" : "fields"}: {wrong.map((key) => FIELD_LABELS[key]).join(", ")}.
      </p>
    );
  }

  // One task with its tick box. The welcome email task also has a Send button.
  function taskRow(j: Joiner, t: TaskInfo) {
    const when = j.done[t.id];
    return (
      <li key={t.id} className={when ? "task-row is-done" : "task-row"}>
        <label className="task">
          <input type="checkbox" checked={!!when} onChange={() => toggleTask(j, t.id)} />
          <span className="task-text">
            <strong>{t.name}</strong>
            <small>{when ? doneText(when, today) : t.hint}</small>
          </span>
        </label>
        {t.id === "welcome" && !when && (
          <button type="button" className="mini" onClick={() => sendWelcome(j)} aria-label={`Send the welcome email to ${j.name}`}>
            <Icon name="mail" size={16} />
            <span>Send</span>
          </button>
        )}
      </li>
    );
  }

  // One joiner on the runway: countdown, readiness bar and the next task with a Done button
  function runwayItem(j: Joiner) {
    const status = statusFor(j);
    const cd = countdown(j.joiningDate, today);
    const next = nextTask(j);
    const nextTeam = next ? TEAMS.find((tm) => tm.key === next.team) : undefined;
    const isSelected = size === "desktop" && selected?.id === j.id;
    return (
      <li key={j.id} className={`t-item ${status}${isSelected ? " selected" : ""}`}>
        <div className="t-when" aria-hidden="true">
          <b className={cd.small ? undefined : "word"}>{cd.big}</b>
          {cd.small && <span>{cd.small}</span>}
          <small>{shortDate(j.joiningDate)}</small>
        </div>
        <article className="t-card">
          <div className="t-top">
            <h3>
              <button type="button" className="t-open" aria-current={isSelected ? "true" : undefined} onClick={() => openDetails(j)}>
                {j.name}
              </button>
            </h3>
            <StatusPill status={status} />
          </div>
          <p className="t-meta">
            {j.jobTitle}, {j.department}
            <span className="sr-only">
              . {joinText(j.joiningDate, today)}, {weekdayName(j.joiningDate)} {formatDate(j.joiningDate)}.
            </span>
          </p>
          <div className="t-progress">
            <Strip joiner={j} />
            <span>
              {doneCount(j)} of {TASKS.length} done
            </span>
          </div>
          {next && nextTeam ? (
            <div className={`t-next ${nextTeam.css}`}>
              <span className="t-next-text">
                Next for {nextTeam.name}: <b>{next.name}</b>
              </span>
              <button type="button" className="done-btn" aria-label={`Mark ${next.name} as done for ${j.name}`} onClick={() => toggleTask(j, next.id)}>
                <Icon name="check" size={18} />
                <span>Done</span>
              </button>
            </div>
          ) : (
            <p className="t-ready">
              <Icon name="check" size={16} />
              Every task is done. Ready for Day 1.
            </p>
          )}
        </article>
      </li>
    );
  }

  // A group on the runway, like "Next 7 days"
  function group(key: string, title: string, list: Joiner[], extra?: ReactNode) {
    return (
      <section className="group" aria-labelledby={`group-${key}`}>
        <h2 className="group-title" id={`group-${key}`}>
          {title}
        </h2>
        {list.length > 0 && <ol className="timeline">{list.map(runwayItem)}</ol>}
        {extra}
      </section>
    );
  }

  // A joiner's full checklist and details. Used beside the list (desktop) and in the side panel.
  function detailsView(j: Joiner, inPanel: boolean) {
    const status = statusFor(j);
    return (
      <>
        <header className="d-head">
          <span className="avatar" aria-hidden="true">
            {initials(j.name)}
          </span>
          <div className="d-title">
            <h2 id={inPanel ? "panel-title" : "pane-title"}>{j.name}</h2>
            <p>
              {j.jobTitle}, {j.department}
            </p>
          </div>
          {inPanel && (
            <button type="button" className="close" aria-label="Close" onClick={closePanel}>
              <Icon name="close" />
            </button>
          )}
        </header>
        <div className="d-body">
          <div className={`d-when ${status}`}>
            <div>
              <b>
                {weekdayName(j.joiningDate)} {formatDate(j.joiningDate)}
              </b>
              <span>
                {joinText(j.joiningDate, today)}. Stage: {stageOf(j)}.
              </span>
            </div>
            <StatusPill status={status} />
          </div>
          <div className="d-progress">
            <Strip joiner={j} big />
            <span>
              {doneCount(j)} of {TASKS.length} tasks done
            </span>
          </div>
          {TEAMS.map((tm) => {
            const tasks = TASKS.filter((t) => t.team === tm.key);
            const n = tasks.filter((t) => j.done[t.id]).length;
            return (
              <section key={tm.key} className={`team-box ${tm.css}`} aria-labelledby={`team-${tm.css}-${inPanel ? "panel" : "pane"}`}>
                <h3 id={`team-${tm.css}-${inPanel ? "panel" : "pane"}`}>
                  <span className="team-dot" aria-hidden="true" />
                  {tm.name}
                  <small>
                    {n} of {tasks.length} done
                  </small>
                </h3>
                <ul className="checklist">{tasks.map((t) => taskRow(j, t))}</ul>
              </section>
            );
          })}
          <dl className="details">
            <div>
              <dt>Personal email</dt>
              <dd>{j.personalEmail}</dd>
            </div>
            <div>
              <dt>Manager</dt>
              <dd>{j.managerEmail}</dd>
            </div>
            {j.notes && (
              <div>
                <dt>Notes</dt>
                <dd>{j.notes}</dd>
              </div>
            )}
          </dl>
          {confirmDelete && (
            <div className="confirm-box">
              <p>
                Delete {j.name} and all their tasks? This cannot be undone.
              </p>
              <div className="row-actions">
                <button type="button" className="danger" onClick={() => deleteJoiner(j)}>
                  Yes, delete
                </button>
                <button type="button" onClick={() => setConfirmDelete(false)}>
                  No
                </button>
              </div>
            </div>
          )}
        </div>
        {!confirmDelete && (
          <footer className="d-foot">
            <button type="button" onClick={() => startEdit(j, inPanel)}>
              <Icon name="edit" size={18} />
              <span>Edit details</span>
            </button>
            <button type="button" className="danger-text" onClick={() => setConfirmDelete(true)}>
              <Icon name="trash" size={18} />
              <span>Delete</span>
            </button>
          </footer>
        )}
      </>
    );
  }

  // ===== SCREEN 1: Runway =====
  function runwayScreen() {
    const upcoming = joiningToday.length + nextWeek.length + later.length;
    const list = (
      <div className="runway-list">
        {screenHead("Runway", `Every joiner by joining date, with what is still open before Day 1. ${counts.attention} need attention.`)}
        <label className="search">
          <Icon name="search" size={18} />
          <span className="sr-only">Search joiners</span>
          <input type="search" value={search} placeholder="Search name, job or department" autoComplete="off" onChange={(e) => setSearch(e.target.value)} />
        </label>
        <div className="chips" role="group" aria-label="Show">
          {SHOWS.map((s) => (
            <button key={s.key} type="button" className={`chip ${s.css}`} aria-pressed={show === s.key} onClick={() => setShow(s.key)}>
              <i className="chip-dot" aria-hidden="true" />
              {s.label} <b>{counts[s.key]}</b>
            </button>
          ))}
        </div>
        <details className="rules">
          <summary>
            <Icon name="info" size={18} />
            What do the labels mean?
          </summary>
          <ul>
            {STATUS_ORDER.map((s) => (
              <li key={s}>
                <StatusPill status={s} />
                <span>{STATUS_INFO[s].rule}</span>
              </li>
            ))}
          </ul>
        </details>

        {shown.length === 0 ? (
          <div className="empty box">
            <p>{onList.length === 0 ? "No joiners yet. Press Add joiner to add the first one." : "No joiner matches. Clear the search or choose All."}</p>
          </div>
        ) : (
          <>
            {joined.length > 0 &&
              group(
                "joined",
                "Joined",
                showJoined ? [...joinedOpen, ...joinedReady] : joinedOpen,
                joinedReady.length > 0 && (
                  <button type="button" className="quiet fold" aria-expanded={showJoined} onClick={() => setShowJoined((v) => !v)}>
                    {showJoined ? "Hide the joiners who are ready" : `Show ${joinedReady.length} who joined and ${joinedReady.length === 1 ? "is" : "are"} ready`}
                  </button>
                )
              )}
            <div className="today-line" role="separator" aria-label={`Today, ${weekdayName(today)} ${formatDate(today)}`}>
              <span>
                Today, {weekdayName(today)} {formatDate(today)}
              </span>
            </div>
            {joiningToday.length > 0 && group("today", "Joining today", joiningToday)}
            {nextWeek.length > 0 && group("week", "In the next 7 days", nextWeek)}
            {later.length > 0 && group("later", "Later", later)}
            {upcoming === 0 && <p className="empty">No one in this list joins after today.</p>}
          </>
        )}
      </div>
    );
    if (size !== "desktop") return list;
    // Desktop: the list and the chosen joiner's checklist side by side
    return (
      <div className="runway-split">
        {list}
        <aside ref={paneRef} className="pane" aria-labelledby={selected ? "pane-title" : undefined} aria-label={selected ? undefined : "Joiner details"}>
          {selected ? (
            detailsView(selected, false)
          ) : (
            <div className="empty pane-empty">
              <p>Choose a joiner to see their checklist.</p>
            </div>
          )}
        </aside>
      </div>
    );
  }

  // ===== SCREEN 2: Tasks for one team =====
  function tasksScreen() {
    const info = TEAMS.find((tm) => tm.key === team) ?? TEAMS[0];
    const teamTasks = TASKS.filter((t) => t.team === team);
    const openGroups = sorted.map((j) => ({ j, tasks: teamTasks.filter((t) => !j.done[t.id]) })).filter((g) => g.tasks.length > 0);
    const doneGroups = sorted.map((j) => ({ j, tasks: teamTasks.filter((t) => j.done[t.id]) })).filter((g) => g.tasks.length > 0);
    const doneTotal = doneGroups.reduce((sum, g) => sum + g.tasks.length, 0);
    const taskGroup = (j: Joiner, tasks: TaskInfo[]) => {
      const status = statusFor(j);
      return (
        <section key={j.id} className={`task-group ${status}`} aria-labelledby={`tg-${j.id}`}>
          <header className="tg-head">
            <div>
              <h3 id={`tg-${j.id}`}>
                <button type="button" className="link" onClick={() => openDetails(j)}>
                  {j.name}
                </button>
              </h3>
              <p>
                {joinText(j.joiningDate, today)}, {shortDate(j.joiningDate)}
              </p>
            </div>
            <StatusPill status={status} />
          </header>
          <ul className="checklist">{tasks.map((t) => taskRow(j, t))}</ul>
        </section>
      );
    };
    return (
      <>
        {screenHead("Tasks", "Tick what your team has finished. Each joiner moves forward by themselves.")}
        <div className="team-switch" role="group" aria-label="Team">
          {TEAMS.map((tm) => (
            <button
              key={tm.key}
              type="button"
              className={`team-tab ${tm.css}`}
              aria-pressed={team === tm.key}
              onClick={() => {
                setTeam(tm.key);
                setShowDone(false);
              }}
            >
              <span className="team-dot" aria-hidden="true" />
              <span className="team-tab-text">
                <strong>{tm.name}</strong>
                <span>{openTasks(tm.key)} open</span>
              </span>
            </button>
          ))}
        </div>
        <p className="team-does">
          {info.name}: {info.does}.
        </p>
        {openGroups.length === 0 ? (
          <div className="empty box">
            <p>No open {info.short} tasks. Each new joiner adds 3 tasks here.</p>
          </div>
        ) : (
          <div className={`task-groups ${info.css}`}>{openGroups.map((g) => taskGroup(g.j, g.tasks))}</div>
        )}
        {doneTotal > 0 && (
          <button type="button" className="quiet fold" aria-expanded={showDone} onClick={() => setShowDone((v) => !v)}>
            {showDone ? "Hide done tasks" : `Show done tasks (${doneTotal})`}
          </button>
        )}
        {showDone && <div className={`task-groups ${info.css}`}>{doneGroups.map((g) => taskGroup(g.j, g.tasks))}</div>}
      </>
    );
  }

  // The add and edit form
  function formPanel() {
    const date = form.joiningDate;
    const dateHint = date && !isWeekend(date) ? `${weekdayName(date)}. ${joinText(date, today)}.` : undefined;
    return (
      <>
        <header className="d-head plain">
          <div className="d-title">
            <h2 id="panel-title">{editingId ? "Edit joiner" : "Add a new joiner"}</h2>
            <p>{editingId ? "Change the details, then save. The ticks stay as they are." : `${TASKS.length} tasks are added for the new joiner: 3 each for IT, HR and the manager.`}</p>
          </div>
          <button type="button" className="close" aria-label="Close" onClick={closeForm}>
            <Icon name="close" />
          </button>
        </header>
        <div className="d-body">
          {errorList()}
          <Field id="field-name" label="Full name" error={errors.name}>
            <input id="field-name" value={form.name} placeholder="Ravi Kumar" autoComplete="off" aria-invalid={!!errors.name} aria-describedby={errors.name ? "field-name-error" : undefined} onChange={(e) => update("name", e.target.value)} />
          </Field>
          <Field id="field-personalEmail" label="Personal email" error={errors.personalEmail}>
            <input
              id="field-personalEmail"
              type="email"
              inputMode="email"
              value={form.personalEmail}
              placeholder="ravi.kumar@gmail.com"
              autoComplete="off"
              aria-invalid={!!errors.personalEmail}
              aria-describedby={errors.personalEmail ? "field-personalEmail-error" : undefined}
              onChange={(e) => update("personalEmail", e.target.value)}
            />
          </Field>
          <div className="pair">
            <Field id="field-department" label="Department" error={errors.department}>
              <select id="field-department" value={form.department} aria-invalid={!!errors.department} aria-describedby={errors.department ? "field-department-error" : undefined} onChange={(e) => update("department", e.target.value)}>
                <option value="">Choose</option>
                {DEPARTMENTS.map((d) => (
                  <option key={d} value={d}>
                    {d}
                  </option>
                ))}
              </select>
            </Field>
            <Field id="field-joiningDate" label="Joining date" error={errors.joiningDate} hint={dateHint}>
              <input id="field-joiningDate" type="date" value={form.joiningDate} aria-invalid={!!errors.joiningDate} aria-describedby={errors.joiningDate ? "field-joiningDate-error" : undefined} onChange={(e) => update("joiningDate", e.target.value)} />
            </Field>
          </div>
          <Field id="field-jobTitle" label="Job title" error={errors.jobTitle}>
            <input id="field-jobTitle" value={form.jobTitle} placeholder="Power Platform Developer" autoComplete="off" aria-invalid={!!errors.jobTitle} aria-describedby={errors.jobTitle ? "field-jobTitle-error" : undefined} onChange={(e) => update("jobTitle", e.target.value)} />
          </Field>
          <Field id="field-managerEmail" label="Manager email" error={errors.managerEmail}>
            <input
              id="field-managerEmail"
              type="email"
              inputMode="email"
              value={form.managerEmail}
              placeholder="priya.sharma@yourcompany.com"
              autoComplete="off"
              aria-invalid={!!errors.managerEmail}
              aria-describedby={errors.managerEmail ? "field-managerEmail-error" : undefined}
              onChange={(e) => update("managerEmail", e.target.value)}
            />
          </Field>
          <Field id="field-notes" label="Notes (optional)" error={errors.notes}>
            <textarea id="field-notes" rows={3} maxLength={NOTES_MAX} value={form.notes} placeholder="Needs a laptop with 16 GB RAM" aria-invalid={!!errors.notes} onChange={(e) => update("notes", e.target.value)} />
            <span className="counter">
              {form.notes.length}/{NOTES_MAX}
            </span>
          </Field>
        </div>
        <footer className="d-foot">
          <button type="button" id="save-button" className="primary" onClick={saveJoiner}>
            {editingId ? "Save changes" : "Add joiner"}
          </button>
          <button type="button" onClick={closeForm}>
            Cancel
          </button>
        </footer>
      </>
    );
  }

  const panelJoiner = panel?.kind === "details" ? joiners.find((j) => j.id === panel.id) : undefined;

  return (
    <>
      <div ref={pageRef} className={`app ${size}`}>
        {/* ----- Desktop: sidebar. Tablet: slim rail with icons ----- */}
        {size !== "phone" && (
          <nav className="side" aria-label="Main">
            {brand}
            <button type="button" className="primary add-btn" onClick={openAddForm}>
              <Icon name="plus" />
              <span>Add joiner</span>
            </button>
            <ul className="side-list">
              <li>
                <button type="button" className="side-link" aria-current={screen === "runway" ? "page" : undefined} onClick={() => go("runway")}>
                  <Icon name="runway" size={22} />
                  <span>Runway</span>
                  {counts.attention > 0 && (
                    <span className="badge alert" aria-label={`${counts.attention} need attention`}>
                      {counts.attention}
                    </span>
                  )}
                </button>
              </li>
              <li>
                <button type="button" className="side-link" aria-current={screen === "tasks" ? "page" : undefined} onClick={() => go("tasks")}>
                  <Icon name="tasks" size={22} />
                  <span>Tasks</span>
                  {openTotal > 0 && (
                    <span className="badge" aria-label={`${openTotal} open tasks`}>
                      {openTotal}
                    </span>
                  )}
                </button>
              </li>
            </ul>
            {size === "desktop" && (
              <div className="side-today">
                <b>{weekdayName(today)}</b>
                <span>{formatDate(today)}</span>
              </div>
            )}
          </nav>
        )}

        <main className="main">
          {/* ----- Phone: a small bar at the top ----- */}
          {size === "phone" && (
            <header className="phone-top">
              {brand}
              <p className="phone-date">{shortDate(today)}</p>
            </header>
          )}
          <div className="screen">{screen === "runway" ? runwayScreen() : tasksScreen()}</div>
        </main>

        {/* ----- Phone: big buttons at the bottom, near your thumb ----- */}
        {size === "phone" && (
          <nav className="tabbar" aria-label="Main">
            <button type="button" className="tab" aria-current={screen === "runway" ? "page" : undefined} onClick={() => go("runway")}>
              <Icon name="runway" size={24} />
              <span>Runway</span>
              {counts.attention > 0 && <span className="badge alert">{counts.attention}</span>}
            </button>
            <button type="button" className="primary tab-add" onClick={openAddForm}>
              <Icon name="plus" size={26} />
              <span>Add joiner</span>
            </button>
            <button type="button" className="tab" aria-current={screen === "tasks" ? "page" : undefined} onClick={() => go("tasks")}>
              <Icon name="tasks" size={24} />
              <span>Tasks</span>
              {openTotal > 0 && <span className="badge">{openTotal}</span>}
            </button>
          </nav>
        )}
      </div>

      {/* ----- Side panel: the form, or one joiner's checklist ----- */}
      {panel && (
        <div
          className={`scrim ${size}`}
          onMouseDown={(e) => {
            // A click on the dark area closes the checklist, but never the form (so typing is not lost)
            if (e.target === e.currentTarget && panel.kind === "details") closePanel();
          }}
        >
          <aside ref={panelRef} className="panel" role="dialog" aria-modal="true" aria-labelledby="panel-title" tabIndex={-1}>
            {panel.kind === "form" && formPanel()}
            {panel.kind === "details" && panelJoiner && detailsView(panelJoiner, true)}
            {panel.kind === "details" && !panelJoiner && (
              <>
                <header className="d-head plain">
                  <div className="d-title">
                    <h2 id="panel-title">Joiner not found</h2>
                    <p>This joiner is no longer in the list.</p>
                  </div>
                  <button type="button" className="close" aria-label="Close" onClick={closePanel}>
                    <Icon name="close" />
                  </button>
                </header>
                <div className="d-body" />
              </>
            )}
          </aside>
        </div>
      )}

      {/* ----- Undo bar after a tick ----- */}
      {toast && (
        <div key={toast.id} className={`toast ${size}`} role="status">
          <span>{toast.text}</span>
          <button type="button" onClick={undo}>
            Undo
          </button>
        </div>
      )}

      {/* ----- Popup after bigger actions: covers the screen and closes by itself ----- */}
      {popup && (
        <div className="overlay" onMouseDown={(e) => e.preventDefault()}>
          <div
            ref={popupRef}
            key={popup.id}
            className={`popup ${popup.tone}`}
            role="dialog"
            aria-modal="true"
            aria-labelledby="popup-title"
            tabIndex={-1}
            onKeyDown={(e) => {
              if (e.key === "Escape") setPopup(null);
              else if (e.key === "Tab") e.preventDefault();
            }}
          >
            <div className="popup-icon" aria-hidden="true">
              <Icon name={popup.icon} size={28} />
            </div>
            <h2 id="popup-title">{popup.title}</h2>
            {popup.lines.map((line) => (
              <p key={line}>{line}</p>
            ))}
            <p className="popup-note">This message closes by itself in {POPUP_SECONDS} seconds.</p>
            <div className="popup-bar" aria-hidden="true">
              <span style={{ animationDuration: `${POPUP_SECONDS}s` }} />
            </div>
          </div>
        </div>
      )}
    </>
  );
}
