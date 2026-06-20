/* eslint-disable @typescript-eslint/no-explicit-any */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend, LineChart, Line,
  ScatterChart, Scatter, ZAxis,
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
} from "recharts";
import { PublicClientApplication } from "@azure/msal-browser";

// ---------- SSO provider types ----------
type SSOProvider = {
  id: "azure" | "google" | string;
  label: string;
  client_id?: string;
  tenant_id?: string;
};
type ProvidersInfo = {
  local_login: boolean;
  providers: SSOProvider[];
  reachable: boolean;
};

// google identity services loader
let googleScriptPromise: Promise<void> | null = null;
function loadGoogleScript(): Promise<void> {
  if (typeof window === "undefined") return Promise.resolve();
  if ((window as any).google?.accounts?.oauth2) return Promise.resolve();
  if (googleScriptPromise) return googleScriptPromise;
  googleScriptPromise = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://accounts.google.com/gsi/client";
    s.async = true; s.defer = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("Failed to load Google script"));
    document.head.appendChild(s);
  });
  return googleScriptPromise;
}

type ChartType = "bar" | "pie" | "line" | "scatter" | "heatmap" | "radar";
const CHART_TYPES: ChartType[] = ["bar", "pie", "line", "scatter", "heatmap", "radar"];

// ---------- Types ----------
type Totals = { events: number; violations: number; alerts: number; agents: number };
type Summary = {
  totals: Totals;
  events_by_module: { module: string; count: number }[];
  alerts_by_severity: { severity: string; count: number }[];
  top_alerting_agents: { agent_id: string; hostname: string; violation_count: number }[];
  hourly_violations: { hour: string; count: number }[];
};
type EventRow = {
  timestamp: string; hostname: string; agent_id: string; module: string;
  action: string; violation: boolean; control_ref?: string;
};
type AlertRow = {
  timestamp: string; rule: string; severity: "CRITICAL" | "HIGH" | "MEDIUM" | "LOW";
  hostname: string; agent_id: string;
};
type Agent = {
  agent_id: string; hostname: string; os_type: string; last_seen: string;
  event_count: number; ip_address?: string;
};

const MODULES = ["usb_control", "dlp_monitor", "secure_access", "storage_control", "policy_engine"];
const MOD_COLORS: Record<string, string> = {
  usb_control: "#00c8ff", dlp_monitor: "#9b59ff", storage_control: "#ffcc00",
  secure_access: "#ff3355", policy_engine: "#00ff88",
};
const SEV_COLORS: Record<string, string> = {
  CRITICAL: "#ff3355", HIGH: "#ffcc00", MEDIUM: "#00c8ff", LOW: "#4a6070",
};

// ---------- Demo data ----------
function rand(min: number, max: number) { return Math.floor(Math.random() * (max - min + 1)) + min; }
function pick<T>(arr: T[]) { return arr[rand(0, arr.length - 1)]; }
const DEMO_HOSTS = ["WIN-DESK-01", "MAC-LAPTOP-04", "LIN-WS-09", "WIN-LAPTOP-12"];
const DEMO_ACTIONS = ["usb_inserted", "file_copied", "login_attempt", "policy_blocked", "scan_complete", "drive_mounted", "network_block"];
const DEMO_CONTROLS = ["ISO27001-A.8.3", "ISO27001-A.9.4", "NIST-AC-17", "PCI-DSS-3.4", "HIPAA-164.312"];

function demoSummary(): Summary {
  return {
    totals: { events: 8421 + rand(-50, 50), violations: 312 + rand(-10, 10), alerts: 47 + rand(-3, 3), agents: 4 },
    events_by_module: MODULES.map((m) => ({ module: m, count: rand(40, 400) })),
    alerts_by_severity: ["CRITICAL", "HIGH", "MEDIUM", "LOW"].map((s) => ({ severity: s, count: rand(2, 20) })),
    top_alerting_agents: DEMO_HOSTS.map((h, i) => ({
      agent_id: `agt-${1000 + i}`, hostname: h, violation_count: rand(20, 90),
    })),
    hourly_violations: Array.from({ length: 24 }, (_, i) => ({
      hour: `${String(i).padStart(2, "0")}:00`, count: rand(0, 30),
    })),
  };
}
function demoEvents(limit: number, offset: number): { events: EventRow[]; total: number } {
  const total = 1234;
  const events: EventRow[] = Array.from({ length: limit }, (_, i) => {
    const n = offset + i;
    const ts = new Date(Date.now() - n * 60_000 - rand(0, 60_000));
    return {
      timestamp: ts.toISOString(),
      hostname: pick(DEMO_HOSTS),
      agent_id: `agt-${1000 + (n % 4)}`,
      module: pick(MODULES),
      action: pick(DEMO_ACTIONS),
      violation: Math.random() < 0.25,
      control_ref: pick(DEMO_CONTROLS),
    };
  });
  return { events, total };
}
function demoAlerts(): AlertRow[] {
  const sevs: AlertRow["severity"][] = ["CRITICAL", "HIGH", "MEDIUM", "LOW"];
  return Array.from({ length: 30 }, (_, i) => ({
    timestamp: new Date(Date.now() - i * 600_000).toISOString(),
    rule: pick(["USB Unauthorized", "Brute Force", "Mass File Copy", "Unauthorized Mount", "Suspicious Login", "Policy Override"]),
    severity: pick(sevs),
    hostname: pick(DEMO_HOSTS),
    agent_id: `agt-${1000 + rand(0, 3)}`,
  }));
}
function demoAgents(): Agent[] {
  return DEMO_HOSTS.map((h, i) => ({
    agent_id: `agt-${1000 + i}`,
    hostname: h,
    os_type: ["Windows 11", "macOS 14", "Ubuntu 22.04", "Windows 10"][i],
    last_seen: new Date(Date.now() - (i < 2 ? rand(1, 8) * 60_000 : rand(40, 1200) * 60_000)).toISOString(),
    event_count: rand(120, 2400),
    ip_address: `192.168.1.${20 + i}`,
  }));
}
function demoRange(range: "day" | "week" | "month"): { hour: string; count: number }[] {
  const n = range === "day" ? 24 : range === "week" ? 7 : 30;
  const fmt = (i: number) => range === "day" ? `${String(i).padStart(2, "0")}:00`
    : `${range === "week" ? "D" : ""}${i + 1}`;
  return Array.from({ length: n }, (_, i) => ({ hour: fmt(i), count: rand(0, range === "day" ? 30 : 120) }));
}
function demoTimeline(): { date: string; events: number; violations: number }[] {
  return Array.from({ length: 30 }, (_, i) => {
    const d = new Date(); d.setDate(d.getDate() - (29 - i));
    return { date: d.toISOString().slice(5, 10), events: rand(80, 400), violations: rand(5, 60) };
  });
}

// ---------- Utils ----------
function parseUtc(ts: string): Date {
  if (!ts) return new Date();
  if (/[zZ]|[+-]\d{2}:?\d{2}$/.test(ts)) return new Date(ts);
  return new Date(ts + "Z");
}
function fmtDateTime(ts: string) {
  const d = parseUtc(ts);
  if (isNaN(+d)) return ts;
  return d.toLocaleString();
}
function fmtRelative(ts: string) {
  const d = parseUtc(ts); const diff = (Date.now() - +d) / 1000;
  if (diff < 60) return `${Math.floor(diff)}s ago`;
  if (diff < 3600) return `${Math.floor(diff / 60)}m ago`;
  if (diff < 86400) return `${Math.floor(diff / 3600)}h ago`;
  return `${Math.floor(diff / 86400)}d ago`;
}
function formatNum(n: number) { return n.toLocaleString(); }

// ---------- Toast ----------
type Toast = { id: number; type: "ok" | "err"; msg: string };

// ---------- Component ----------
export default function EndguardX() {
  // theme
  const [theme, setTheme] = useState<"dark" | "light">(() => {
    if (typeof window === "undefined") return "dark";
    return (localStorage.getItem("gx-theme") as any) || "dark";
  });
  useEffect(() => {
    document.documentElement.setAttribute("data-gx-theme", theme);
    try { localStorage.setItem("gx-theme", theme); } catch { /* ignore */ }
  }, [theme]);

  // managers + url
  const [managerUrl, setManagerUrl] = useState<string>(() => {
    if (typeof window === "undefined") return "https://192.168.1.5:8443";
    return localStorage.getItem("gx-last-manager") || "https://192.168.1.5:8443";
  });
  const [managers, setManagers] = useState<string[]>(() => {
    if (typeof window === "undefined") return [];
    try { return JSON.parse(localStorage.getItem("gx-managers") || "[]"); } catch { return []; }
  });
  useEffect(() => { try { localStorage.setItem("gx-managers", JSON.stringify(managers)); } catch { /* ignore */ } }, [managers]);
  useEffect(() => { try { localStorage.setItem("gx-last-manager", managerUrl); } catch { /* ignore */ } }, [managerUrl]);

  // auth / connection (restore from sessionStorage)
  const [token, setToken] = useState<string>(() => {
    if (typeof window === "undefined") return "";
    return sessionStorage.getItem("gx-dash-token") || "";
  });
  const [authed, setAuthed] = useState<boolean>(() => {
    if (typeof window === "undefined") return false;
    return !!sessionStorage.getItem("gx-dash-token");
  });
  const [sessionUser, setSessionUser] = useState<string>(() => {
    if (typeof window === "undefined") return "";
    return sessionStorage.getItem("gx-dash-user") || "";
  });
  const [demoMode, setDemoMode] = useState<boolean>(false);
  const [connecting, setConnecting] = useState(false);
  const [connStatus, setConnStatus] = useState<"offline" | "live" | "demo">(() => {
    if (typeof window !== "undefined" && sessionStorage.getItem("gx-dash-token")) return "live";
    return "offline";
  });
  const [errBanner, setErrBanner] = useState<string>("");
  const [lastSync, setLastSync] = useState<string>("--");

  // data
  const [summary, setSummary] = useState<Summary | null>(null);
  const [prevTotals, setPrevTotals] = useState<Totals | null>(null);
  const [rangeSel, setRangeSel] = useState<"day" | "week" | "month">("day");
  const [rangeData, setRangeData] = useState<{ hour: string; count: number }[]>([]);
  const [timeline, setTimeline] = useState<{ date: string; events: number; violations: number }[]>([]);

  // events tab
  const [evFilters, setEvFilters] = useState({ module: "", violation: "", agent_id: "" });
  const [evLimit, setEvLimit] = useState(25);
  const [evOffset, setEvOffset] = useState(0);
  const [events, setEvents] = useState<EventRow[]>([]);
  const [eventsTotal, setEventsTotal] = useState(0);
  const [eventsLoading, setEventsLoading] = useState(false);

  // alerts
  const [alerts, setAlerts] = useState<AlertRow[]>([]);
  const [alertsPage, setAlertsPage] = useState(1);
  const [alertFilters, setAlertFilters] = useState({ severity: "", search: "" });
  const ALERTS_PER_PAGE = 25;

  // agents
  const [agents, setAgents] = useState<Agent[]>([]);
  const [agentFilters, setAgentFilters] = useState({ status: "", search: "" });
  const [agentsPage, setAgentsPage] = useState(1);
  const AGENTS_PER_PAGE = 10;
  const [scanTarget, setScanTarget] = useState<string>(() => {
    if (typeof window === "undefined") return "";
    return localStorage.getItem("gx-scan-target") || "";
  });
  useEffect(() => { try { localStorage.setItem("gx-scan-target", scanTarget); } catch { /* ignore */ } }, [scanTarget]);

  // tab
  const [tab, setTab] = useState<"events" | "alerts" | "agents">("events");

  // chart types per panel
  const [violationsChart, setViolationsChart] = useState<ChartType>("bar");
  const [modulesChart, setModulesChart] = useState<ChartType>("pie");
  const [severityChart, setSeverityChart] = useState<ChartType>("pie");
  const [topAgentsChart, setTopAgentsChart] = useState<ChartType>("bar");
  const [timelineChart, setTimelineChart] = useState<ChartType>("line");

  // toasts
  const [toasts, setToasts] = useState<Toast[]>([]);
  const pushToast = useCallback((type: Toast["type"], msg: string) => {
    const id = Date.now() + Math.random();
    setToasts((t) => [...t, { id, type, msg }]);
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 3000);
  }, []);

  // ---------- API helper ----------
  const api = useCallback(async (path: string, opts: RequestInit = {}) => {
    const res = await fetch(`${managerUrl}${path}`, {
      ...opts,
      headers: {
        "Content-Type": "application/json",
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
        ...(opts.headers || {}),
      },
    });
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      const err: any = new Error(`HTTP ${res.status} ${text}`);
      err.status = res.status;
      throw err;
    }
    return res.status === 204 ? null : res.json();
  }, [managerUrl, token]);

  // ---------- Fetch all ----------
  const fetchAll = useCallback(async () => {
    if (demoMode) {
      const s = demoSummary();
      setPrevTotals(summary?.totals || null);
      setSummary(s);
      setRangeData(demoRange(rangeSel));
      setTimeline(demoTimeline());
      setAlerts(demoAlerts());
      setAgents(demoAgents());
      const d = demoEvents(evLimit, evOffset);
      setEvents(d.events); setEventsTotal(d.total);
      setLastSync(new Date().toLocaleTimeString());
      setConnStatus("demo");
      return;
    }
    try {
      const [s, r, t, al, ag, ev] = await Promise.all([
        api("/api/v1/dashboard/summary"),
        api(`/api/v1/dashboard/violations/range?range=${rangeSel}`),
        api("/api/v1/dashboard/timeline?days=30"),
        api("/api/v1/dashboard/alerts?limit=50"),
        api("/api/v1/dashboard/agents"),
        api(`/api/v1/dashboard/events?limit=${evLimit}&offset=${evOffset}` +
          (evFilters.module ? `&module=${encodeURIComponent(evFilters.module)}` : "") +
          (evFilters.violation ? `&violation=${evFilters.violation}` : "") +
          (evFilters.agent_id ? `&agent_id=${encodeURIComponent(evFilters.agent_id)}` : "")),
      ]);
      setPrevTotals(summary?.totals || null);
      setSummary(s);
      setRangeData(Array.isArray(r) ? r : (r?.data || []));
      setTimeline(Array.isArray(t) ? t : (t?.data || []));
      setAlerts(Array.isArray(al) ? al : (al?.alerts || []));
      setAgents(Array.isArray(ag) ? ag : (ag?.agents || []));
      setEvents(ev?.events || []);
      setEventsTotal(ev?.total || 0);
      setLastSync(new Date().toLocaleTimeString());
      setConnStatus("live");
      setErrBanner("");
    } catch (e: any) {
      if (e?.status === 401) {
        setErrBanner("Invalid token - check dashboard_tokens.json on manager");
      } else {
        setErrBanner(`Cannot reach ${managerUrl}. Is server.py running?`);
      }
      setConnStatus("offline");
    }
  }, [api, demoMode, rangeSel, evLimit, evOffset, evFilters, managerUrl, summary?.totals]);

  // ---------- Fetch events only (filter/page changes) ----------
  const fetchEvents = useCallback(async () => {
    if (demoMode) {
      setEventsLoading(true);
      const d = demoEvents(evLimit, evOffset);
      setEvents(d.events); setEventsTotal(d.total);
      setEventsLoading(false);
      return;
    }
    setEventsLoading(true);
    try {
      const ev = await api(`/api/v1/dashboard/events?limit=${evLimit}&offset=${evOffset}` +
        (evFilters.module ? `&module=${encodeURIComponent(evFilters.module)}` : "") +
        (evFilters.violation ? `&violation=${evFilters.violation}` : "") +
        (evFilters.agent_id ? `&agent_id=${encodeURIComponent(evFilters.agent_id)}` : ""));
      setEvents(ev?.events || []);
      setEventsTotal(ev?.total || 0);
    } catch {
      pushToast("err", "Failed to load events");
    } finally { setEventsLoading(false); }
  }, [api, demoMode, evLimit, evOffset, evFilters, pushToast]);

  // refetch events on filter / page changes (after first connect)
  const initRef = useRef(false);
  useEffect(() => {
    if (!authed && !demoMode) return;
    if (!initRef.current) { initRef.current = true; return; }
    void fetchEvents();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [evFilters.module, evFilters.violation, evLimit, evOffset, rangeSel]);

  // debounced agent_id filter
  useEffect(() => {
    if (!authed && !demoMode) return;
    const t = setTimeout(() => { void fetchEvents(); }, 500);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [evFilters.agent_id]);

  // auto-refresh
  useEffect(() => {
    if (!authed && !demoMode) return;
    const id = setInterval(() => { void fetchAll(); }, 30_000);
    return () => clearInterval(id);
  }, [authed, demoMode, fetchAll]);

  // refetch range when changed
  useEffect(() => {
    if (!authed && !demoMode) return;
    if (demoMode) { setRangeData(demoRange(rangeSel)); return; }
    api(`/api/v1/dashboard/violations/range?range=${rangeSel}`)
      .then((r) => setRangeData(Array.isArray(r) ? r : (r?.data || [])))
      .catch(() => { /* ignore */ });
  }, [rangeSel, authed, demoMode, api]);

  // ---------- Login ----------
  const [loginUser, setLoginUser] = useState("");
  const [loginPass, setLoginPass] = useState("");
  const [loginErr, setLoginErr] = useState("");
  const [loggingIn, setLoggingIn] = useState(false);
  const [ssoBusy, setSsoBusy] = useState<string>("");

  // providers info (fetched when login screen is shown)
  const [providersInfo, setProvidersInfo] = useState<ProvidersInfo>({
    local_login: true, providers: [], reachable: true,
  });
  const [providersLoading, setProvidersLoading] = useState(false);

  // clock on login screen
  const [nowStr, setNowStr] = useState<string>(() => new Date().toLocaleString());
  useEffect(() => {
    const id = setInterval(() => setNowStr(new Date().toLocaleString()), 1000);
    return () => clearInterval(id);
  }, []);

  const fetchProviders = useCallback(async () => {
    if (!managerUrl) return;
    setProvidersLoading(true);
    try {
      const res = await fetch(`${managerUrl}/api/v1/auth/providers`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setProvidersInfo({
        local_login: data?.local_login !== false,
        providers: Array.isArray(data?.providers) ? data.providers : [],
        reachable: true,
      });
    } catch {
      setProvidersInfo({ local_login: true, providers: [], reachable: false });
    } finally {
      setProvidersLoading(false);
    }
  }, [managerUrl]);

  // fetch providers whenever login screen is visible / manager url changes
  useEffect(() => {
    if (authed) return;
    void fetchProviders();
  }, [authed, fetchProviders]);

  // persist session helper
  const persistSession = useCallback((tok: string, user: string) => {
    try {
      sessionStorage.setItem("gx-dash-token", tok);
      sessionStorage.setItem("gx-dash-manager", managerUrl);
      sessionStorage.setItem("gx-dash-user", user || "");
    } catch { /* ignore */ }
  }, [managerUrl]);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoginErr(""); setLoggingIn(true);
    try {
      const res = await fetch(`${managerUrl}/api/v1/auth/login`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username: loginUser, password: loginPass }),
      });
      if (res.status === 401) { setLoginErr("Invalid username or password"); return; }
      if (!res.ok) { setLoginErr(`Login failed (HTTP ${res.status})`); return; }
      const data = await res.json();
      setToken(data.token || "");
      setSessionUser(data.username || loginUser);
      persistSession(data.token || "", data.username || loginUser);
      setAuthed(true);
      setDemoMode(false);
      setConnStatus("live");
      setTimeout(() => void fetchAll(), 0);
    } catch {
      setLoginErr("Cannot reach manager. Check URL.");
    } finally { setLoggingIn(false); }
  };

  const handleAzureLogin = async (p: SSOProvider) => {
    setLoginErr(""); setSsoBusy("azure");
    try {
      if (!p.client_id || !p.tenant_id) throw new Error("Azure provider misconfigured");
      const msal = new PublicClientApplication({
        auth: {
          clientId: p.client_id,
          authority: `https://login.microsoftonline.com/${p.tenant_id}`,
          redirectUri: window.location.origin,
        },
      });
      await msal.initialize();
      const result = await msal.loginPopup({ scopes: ["openid", "profile", "email"] });
      const accessToken = result.accessToken || (result as any).idToken;
      const res = await fetch(`${managerUrl}/api/v1/auth/sso/verify`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ provider: "azure", access_token: accessToken }),
      });
      if (!res.ok) { setLoginErr(`SSO failed (HTTP ${res.status})`); return; }
      const data = await res.json();
      setToken(data.token || "");
      setSessionUser(data.username || "");
      persistSession(data.token || "", data.username || "");
      setAuthed(true); setDemoMode(false); setConnStatus("live");
      setTimeout(() => void fetchAll(), 0);
    } catch (err: any) {
      setLoginErr(err?.message || "Microsoft sign-in failed");
    } finally { setSsoBusy(""); }
  };

  const handleGoogleLogin = async (p: SSOProvider) => {
    setLoginErr(""); setSsoBusy("google");
    try {
      if (!p.client_id) throw new Error("Google provider misconfigured");
      await loadGoogleScript();
      const google = (window as any).google;
      const accessToken: string = await new Promise((resolve, reject) => {
        const client = google.accounts.oauth2.initTokenClient({
          client_id: p.client_id,
          scope: "openid email profile",
          callback: (resp: any) => {
            if (resp?.error) reject(new Error(resp.error));
            else resolve(resp.access_token);
          },
        });
        client.requestAccessToken();
      });
      const res = await fetch(`${managerUrl}/api/v1/auth/sso/verify`, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ provider: "google", access_token: accessToken }),
      });
      if (!res.ok) { setLoginErr(`SSO failed (HTTP ${res.status})`); return; }
      const data = await res.json();
      setToken(data.token || "");
      setSessionUser(data.username || "");
      persistSession(data.token || "", data.username || "");
      setAuthed(true); setDemoMode(false); setConnStatus("live");
      setTimeout(() => void fetchAll(), 0);
    } catch (err: any) {
      setLoginErr(err?.message || "Google sign-in failed");
    } finally { setSsoBusy(""); }
  };

  const enterDemo = () => {
    setDemoMode(true); setAuthed(true); setConnStatus("demo"); setErrBanner("");
    setTimeout(() => void fetchAll(), 0);
  };

  const handleConnect = () => {
    if (!token) { enterDemo(); return; }
    setConnecting(true);
    fetchAll().finally(() => setConnecting(false));
  };

  const logout = () => {
    setAuthed(false); setToken(""); setSessionUser(""); setDemoMode(false); setConnStatus("offline");
    setSummary(null); setEvents([]); setAlerts([]); setAgents([]); setLastSync("--");
    try {
      sessionStorage.removeItem("gx-dash-token");
      sessionStorage.removeItem("gx-dash-manager");
      sessionStorage.removeItem("gx-dash-user");
    } catch { /* ignore */ }
  };

  // ---------- Manager helpers ----------
  const addManager = () => {
    if (!managerUrl) return;
    if (managers.includes(managerUrl)) { pushToast("err", "Already saved"); return; }
    setManagers([...managers, managerUrl]); pushToast("ok", "Manager saved");
  };
  const removeManager = () => {
    if (!managers.includes(managerUrl)) return;
    setManagers(managers.filter((m) => m !== managerUrl));
    pushToast("ok", "Manager removed");
  };

  // ---------- Trend helper ----------
  const trend = (k: keyof Totals): { dir: "up" | "down" | null; diff: number } => {
    if (!summary || !prevTotals) return { dir: null, diff: 0 };
    const diff = summary.totals[k] - prevTotals[k];
    if (diff === 0) return { dir: null, diff: 0 };
    return { dir: diff > 0 ? "up" : "down", diff: Math.abs(diff) };
  };

  // ---------- Derived ----------
  const onlineAgents = useMemo(() => {
    const now = Date.now();
    return agents.filter((a) => now - +parseUtc(a.last_seen) < 12 * 60_000);
  }, [agents]);

  const alertsFiltered = useMemo(() => {
    const q = alertFilters.search.trim().toLowerCase();
    return alerts.filter((a) =>
      (!alertFilters.severity || a.severity === alertFilters.severity) &&
      (!q || a.rule.toLowerCase().includes(q) || a.hostname.toLowerCase().includes(q) || a.agent_id.toLowerCase().includes(q))
    );
  }, [alerts, alertFilters]);
  const alertsPaged = useMemo(() => {
    const start = (alertsPage - 1) * ALERTS_PER_PAGE;
    return alertsFiltered.slice(start, start + ALERTS_PER_PAGE);
  }, [alertsFiltered, alertsPage]);
  const alertsTotalPages = Math.max(1, Math.ceil(alertsFiltered.length / ALERTS_PER_PAGE));

  const agentsFiltered = useMemo(() => {
    const q = agentFilters.search.trim().toLowerCase();
    const now = Date.now();
    return agents.filter((a) => {
      const online = now - +parseUtc(a.last_seen) < 12 * 60_000;
      if (agentFilters.status === "online" && !online) return false;
      if (agentFilters.status === "offline" && online) return false;
      if (q && !a.hostname.toLowerCase().includes(q) && !a.agent_id.toLowerCase().includes(q) && !(a.ip_address || "").toLowerCase().includes(q) && !a.os_type.toLowerCase().includes(q)) return false;
      return true;
    });
  }, [agents, agentFilters]);
  const agentsPaged = useMemo(() => {
    const start = (agentsPage - 1) * AGENTS_PER_PAGE;
    return agentsFiltered.slice(start, start + AGENTS_PER_PAGE);
  }, [agentsFiltered, agentsPage]);
  const agentsTotalPages = Math.max(1, Math.ceil(agentsFiltered.length / AGENTS_PER_PAGE));

  useEffect(() => { setAlertsPage(1); }, [alertFilters]);
  useEffect(() => { setAgentsPage(1); }, [agentFilters]);

  // ---------- Scan actions ----------
  const scanAgent = async (a: Agent) => {
    if (!scanTarget) { pushToast("err", "Set SCAN TARGET URL first"); return; }
    if (!a.ip_address) { pushToast("err", "Agent has no IP"); return; }
    if (!confirm(`Dispatch scan on ${a.hostname} (${a.ip_address})?`)) return;
    if (demoMode) { pushToast("ok", `Scan dispatched to ${a.hostname} (demo)`); return; }
    try {
      await api("/api/v1/scan/dispatch", {
        method: "POST",
        body: JSON.stringify({ agent_id: a.agent_id, ip: a.ip_address, hostname: a.hostname, scan_target: scanTarget }),
      });
      pushToast("ok", `Scan dispatched to ${a.hostname}`);
    } catch { pushToast("err", `Scan failed for ${a.hostname}`); }
  };
  const scanAll = async () => {
    if (!scanTarget) { pushToast("err", "Set SCAN TARGET URL first"); return; }
    const targets = onlineAgents.filter((a) => a.ip_address);
    if (!targets.length) { pushToast("err", "No online agents with IPs"); return; }
    if (!confirm(`Dispatch scan to ${targets.length} online agent(s)?`)) return;
    if (demoMode) { pushToast("ok", `${targets.length} dispatched (demo)`); return; }
    let ok = 0, fail = 0;
    await Promise.all(targets.map(async (a) => {
      try {
        await api("/api/v1/scan/dispatch", {
          method: "POST",
          body: JSON.stringify({ agent_id: a.agent_id, ip: a.ip_address, hostname: a.hostname, scan_target: scanTarget }),
        });
        ok++;
      } catch { fail++; }
    }));
    pushToast(fail ? "err" : "ok", `${ok} dispatched, ${fail} failed`);
  };
  const deleteAgent = async (a: Agent) => {
    if (!confirm(`Delete agent ${a.agent_id}?`)) return;
    if (demoMode) {
      setAgents((arr) => arr.filter((x) => x.agent_id !== a.agent_id));
      pushToast("ok", "Agent deleted (demo)"); return;
    }
    try {
      await api(`/api/v1/agents/${encodeURIComponent(a.agent_id)}`, { method: "DELETE" });
      setAgents((arr) => arr.filter((x) => x.agent_id !== a.agent_id));
      pushToast("ok", "Agent deleted");
    } catch { pushToast("err", "Delete failed"); }
  };

  // ---------- Chart theme colors ----------
  const isDark = theme === "dark";
  const axisColor = isDark ? "#9bb0c2" : "#5a6a7c";
  const gridColor = isDark ? "rgba(0,200,255,0.12)" : "rgba(0,0,0,0.06)";
  const tooltipStyle = {
    backgroundColor: isDark ? "#0d1117" : "#fff",
    border: `1px solid ${isDark ? "rgba(0,200,255,0.2)" : "rgba(0,0,0,0.15)"}`,
    color: isDark ? "#c8d8e8" : "#1a2233",
    fontSize: 11,
    fontFamily: "Share Tech Mono, monospace",
  };

  // ---------- Render ----------
  const showLogin = !authed;

  return (
    <div className="gx-root">
      <div className="gx-scanlines" />

      {showLogin && (() => {
        const azureP = providersInfo.providers.find((x) => x.id === "azure");
        const googleP = providersInfo.providers.find((x) => x.id === "google");
        const showLocal = providersInfo.local_login || !providersInfo.reachable;
        const showDivider = showLocal && (azureP || googleP);
        return (
          <div className="gx-login-overlay">
            <form className="gx-login-card" onSubmit={handleLogin}>
              <h2>Endguard<span style={{ color: "var(--gx-green)" }}>X</span></h2>
              <div className="sub">SECURE ACCESS</div>

              {/* status + clock */}
              <div style={{
                display: "flex", justifyContent: "space-between", alignItems: "center",
                fontSize: 11, letterSpacing: "0.15em", margin: "8px 0 12px",
                color: "var(--gx-fg-dim, #8aa0b4)", fontFamily: "Share Tech Mono, monospace",
              }}>
                <span style={{
                  color: providersInfo.reachable ? "var(--gx-green, #29d398)" : "var(--gx-red, #ff5a6e)",
                }}>
                  ● {providersLoading ? "CHECKING..." : providersInfo.reachable ? "MANAGER ONLINE" : "MANAGER OFFLINE"}
                </span>
                <span>{nowStr}</span>
              </div>

              {/* theme toggle */}
              <div
                onClick={() => setTheme(theme === "dark" ? "light" : "dark")}
                style={{
                  cursor: "pointer", fontSize: 10, letterSpacing: "0.2em",
                  textAlign: "right", marginBottom: 10,
                  color: "var(--gx-cyan, #00c8ff)",
                }}
                title="Toggle theme"
              >
                ☾ {theme.toUpperCase()} MODE
              </div>

              <label>MANAGER URL</label>
              <input
                type="text" value={managerUrl}
                onChange={(e) => setManagerUrl(e.target.value)}
                onBlur={() => void fetchProviders()}
                placeholder="https://manager:8443"
              />

              {!providersInfo.reachable && (
                <div style={{
                  fontSize: 10, color: "var(--gx-amber, #ffb454)", margin: "4px 0 8px",
                  letterSpacing: "0.1em",
                }}>
                  ⚠ MANAGER UNREACHABLE — LOCAL LOGIN ONLY
                </div>
              )}

              {showLocal && (
                <>
                  <label>USERNAME</label>
                  <input type="text" value={loginUser} onChange={(e) => setLoginUser(e.target.value)} autoFocus />
                  <label>PASSWORD</label>
                  <input type="password" value={loginPass} onChange={(e) => setLoginPass(e.target.value)} />
                  <button type="submit" disabled={loggingIn || !!ssoBusy}>{loggingIn ? "SIGNING IN..." : "SIGN IN"}</button>
                </>
              )}

              {showDivider && (
                <div style={{
                  display: "flex", alignItems: "center", gap: 8, margin: "14px 0 10px",
                  color: "var(--gx-fg-dim, #6b7b8d)", fontSize: 10, letterSpacing: "0.3em",
                }}>
                  <div style={{ flex: 1, height: 1, background: "currentColor", opacity: 0.3 }} />
                  OR
                  <div style={{ flex: 1, height: 1, background: "currentColor", opacity: 0.3 }} />
                </div>
              )}

              {azureP && (
                <button
                  type="button"
                  onClick={() => handleAzureLogin(azureP)}
                  disabled={!!ssoBusy || loggingIn}
                  style={{ marginTop: 6 }}
                >
                  {ssoBusy === "azure" ? "OPENING MICROSOFT..." : (azureP.label || "LOGIN WITH MICROSOFT").toUpperCase()}
                </button>
              )}
              {googleP && (
                <button
                  type="button"
                  onClick={() => handleGoogleLogin(googleP)}
                  disabled={!!ssoBusy || loggingIn}
                  style={{ marginTop: 6 }}
                >
                  {ssoBusy === "google" ? "OPENING GOOGLE..." : (googleP.label || "LOGIN WITH GOOGLE").toUpperCase()}
                </button>
              )}

              <button type="button" style={{ marginTop: 8 }} onClick={enterDemo}>ENTER DEMO MODE</button>
              <div className="gx-login-err">{loginErr}</div>
            </form>
          </div>
        );
      })()}

      {/* Topbar */}
      <header className="gx-topbar">
        <div>
          <div className="gx-logo">
            <span style={{ fontSize: 26 }}>■</span>Endguard<span style={{ color: "var(--gx-green)", fontSize: 28 }}>X</span>
          </div>
          <div className="gx-logo-sub">ENDPOINT CONTROL PLATFORM</div>
        </div>

        <div className="gx-topbar-right">
          <div className="gx-theme-toggle" onClick={() => setTheme(theme === "dark" ? "light" : "dark")} title="Toggle theme">
            <span className="gx-toggle-lbl">{theme.toUpperCase()}</span>
            <div className="gx-toggle-track"><div className="gx-toggle-knob" /></div>
          </div>

          <div className="gx-conn-box">
            <label>MANAGER</label><div className="gx-conn-sep" />
            <select value="" onChange={(e) => { if (e.target.value) setManagerUrl(e.target.value); }}>
              <option value="">-- saved --</option>
              {managers.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
            <input type="text" value={managerUrl} onChange={(e) => setManagerUrl(e.target.value)}
              placeholder="https://manager:8443" />
            <button className="gx-btn-mini" type="button" onClick={addManager}>+ ADD</button>
            <button className="gx-btn-mini danger" type="button" onClick={removeManager}>DEL</button>
            <div className="gx-conn-sep" />
            <label>TOKEN</label><div className="gx-conn-sep" />
            <input type="password" value={token} onChange={(e) => setToken(e.target.value)}
              placeholder="paste token (empty = demo)" style={{ width: 180 }} />
          </div>

          <button
            className={`gx-btn-connect ${connStatus === "live" ? "connected" : ""} ${connStatus === "demo" ? "demo" : ""}`}
            onClick={handleConnect} disabled={connecting}>
            {connecting ? "CONNECTING..." : connStatus === "live" ? "CONNECTED" : connStatus === "demo" ? "DEMO" : "CONNECT"}
          </button>

          <div className="gx-status-pill">
            <div className={`gx-dot ${connStatus === "live" ? "gx-dot-live" : connStatus === "demo" ? "gx-dot-demo" : "gx-dot-dead"}`} />
            <span style={{
              color: connStatus === "live" ? "var(--gx-green)" : connStatus === "demo" ? "var(--gx-yellow)" : "var(--gx-muted)",
            }}>{connStatus.toUpperCase()}</span>
          </div>

          <div className="gx-sync-time">{lastSync}</div>

          {authed && (
            <button className="gx-btn-mini danger" onClick={logout}>LOGOUT</button>
          )}
        </div>
      </header>

      <main className="gx-main">
        {demoMode && (
          <div className="gx-demo-banner">
            <span>⚠ DEMO MODE — Showing simulated data. Enter manager URL and token to see real data.</span>
            <button className="gx-btn-sm" onClick={logout}>Connect to Real Manager</button>
          </div>
        )}
        {errBanner && <div className="gx-err-banner">{errBanner}</div>}

        {/* Stat cards */}
        <div className="gx-stat-row">
          <StatCard color="cyan" label="Total Events" value={summary?.totals.events}
            sub={summary ? `${summary.totals.agents} agents reporting` : "connect to load"}
            trend={trend("events")} />
          <StatCard color="red" label="Violations" value={summary?.totals.violations}
            sub="policy breaches" trend={trend("violations")} />
          <StatCard color="yellow" label="Fired Alerts" value={summary?.totals.alerts}
            sub="rule engine hits" trend={trend("alerts")} />
          <StatCard color="green" label="Active Agents" value={summary?.totals.agents}
            sub="reporting endpoints" trend={trend("agents")} />
        </div>

        {/* Row 1: violations bar + modules doughnut */}
        <div className="gx-grid2">
          <Panel
            title={
              <span>Violations
                <select value={rangeSel} onChange={(e) => setRangeSel(e.target.value as any)}
                  style={{ background: "var(--gx-input-bg)", color: "var(--gx-text)", border: "1px solid var(--gx-border2)",
                    padding: "2px 6px", fontFamily: "var(--gx-font)", fontSize: 11, marginLeft: 6 }}>
                  <option value="day">Last 24h (hourly)</option>
                  <option value="week">Last 7 days</option>
                  <option value="month">Last 30 days</option>
                </select>
              </span>
            }
            right={
              <>
                <span>{rangeData.length ? `${rangeData.reduce((s, r) => s + r.count, 0)} total` : "--"}</span>
                <ChartTypeSelect value={violationsChart} onChange={setViolationsChart} />
              </>
            }
          >
            <div className="gx-chart-wrap">
              <FlexChart
                type={violationsChart}
                data={rangeData.map((r) => ({ name: r.hour, value: r.count }))}
                colorFor={() => "#ff3355"}
                baseColor="#ff3355"
                axisColor={axisColor} gridColor={gridColor} tooltipStyle={tooltipStyle}
              />
            </div>
          </Panel>

          <Panel
            title="Events by module"
            right={
              <>
                <span>{summary ? `${summary.events_by_module.length} modules • ${summary.totals.events.toLocaleString()} total events` : "--"}</span>
                <ChartTypeSelect value={modulesChart} onChange={setModulesChart} />
              </>
            }
          >
            <div className="gx-chart-wrap">
              <FlexChart
                type={modulesChart}
                data={(summary?.events_by_module || []).map((m) => ({ name: m.module, value: m.count }))}
                colorFor={(n) => MOD_COLORS[n] || "#00c8ff"}
                baseColor="#00c8ff"
                axisColor={axisColor} gridColor={gridColor} tooltipStyle={tooltipStyle}
              />
            </div>
          </Panel>
        </div>

        {/* Row 2: severity + top agents + timeline */}
        <div className="gx-grid3">
          <Panel
            title="Alert Severity"
            right={
              <>
                <span>{summary ? `${summary.totals.alerts.toLocaleString()} total alerts` : "--"}</span>
                <ChartTypeSelect value={severityChart} onChange={setSeverityChart} />
              </>
            }
          >
            <div className="gx-chart-wrap">
              <FlexChart
                type={severityChart}
                data={(summary?.alerts_by_severity || []).map((s) => ({ name: s.severity, value: s.count }))}
                colorFor={(n) => SEV_COLORS[n] || "#4a6070"}
                baseColor="#ff3355"
                axisColor={axisColor} gridColor={gridColor} tooltipStyle={tooltipStyle}
              />
            </div>
          </Panel>

          <Panel
            title="Top Agents by Events"
            right={
              <>
                <span>top 6</span>
                <ChartTypeSelect value={topAgentsChart} onChange={setTopAgentsChart} />
              </>
            }
          >
            <div className="gx-chart-wrap">
              <FlexChart
                type={topAgentsChart}
                data={(summary?.top_alerting_agents || []).slice(0, 6).map((a) => ({ name: a.hostname, value: a.violation_count }))}
                colorFor={() => "#00c8ff"}
                baseColor="#00c8ff"
                horizontal
                axisColor={axisColor} gridColor={gridColor} tooltipStyle={tooltipStyle}
              />
            </div>
          </Panel>

          <Panel
            title="Events Timeline - 30 Days"
            right={
              <>
                <span>events vs violations</span>
                <ChartTypeSelect value={timelineChart} onChange={setTimelineChart} />
              </>
            }
          >
            <div className="gx-chart-wrap">
              <FlexTimeline
                type={timelineChart}
                data={timeline}
                axisColor={axisColor} gridColor={gridColor} tooltipStyle={tooltipStyle}
              />
            </div>
          </Panel>
        </div>


        {/* Tabs */}
        <div className="gx-panel" style={{ marginBottom: 24 }}>
          <div className="gx-tabs">
            <button className={`gx-tab-btn ${tab === "events" ? "active" : ""}`} onClick={() => setTab("events")}>Events</button>
            <button className={`gx-tab-btn ${tab === "alerts" ? "active" : ""}`} onClick={() => setTab("alerts")}>Alerts</button>
            <button className={`gx-tab-btn ${tab === "agents" ? "active" : ""}`} onClick={() => setTab("agents")}>Agents</button>
          </div>

          {tab === "events" && (
            <>
              <div className="gx-filter-bar">
                <span className="gx-filter-label">MODULE</span>
                <select value={evFilters.module} onChange={(e) => { setEvFilters({ ...evFilters, module: e.target.value }); setEvOffset(0); }}>
                  <option value="">All modules</option>
                  {MODULES.map((m) => <option key={m} value={m}>{m}</option>)}
                </select>
                <span className="gx-filter-label">STATUS</span>
                <select value={evFilters.violation} onChange={(e) => { setEvFilters({ ...evFilters, violation: e.target.value }); setEvOffset(0); }}>
                  <option value="">All events</option>
                  <option value="true">Violations only</option>
                  <option value="false">Clean only</option>
                </select>
                <span className="gx-filter-label">AGENT ID</span>
                <input type="text" value={evFilters.agent_id}
                  onChange={(e) => { setEvFilters({ ...evFilters, agent_id: e.target.value }); setEvOffset(0); }}
                  placeholder="filter…" />
                <span className="gx-filter-label">PAGE SIZE</span>
                <select value={evLimit} onChange={(e) => { setEvLimit(Number(e.target.value)); setEvOffset(0); }}>
                  <option value={25}>25</option><option value={50}>50</option><option value={100}>100</option>
                </select>
                <button className="gx-btn-sm" onClick={() => void fetchEvents()}>Refresh</button>
              </div>

              <div style={{ overflowX: "auto" }}>
                <table className="gx-tbl">
                  <thead><tr>
                    <th>Time</th><th>Host</th><th>Agent ID</th><th>Module</th>
                    <th>Action</th><th>Status</th><th>Control Ref</th>
                  </tr></thead>
                  <tbody>
                    {eventsLoading ? <tr><td colSpan={7}><div className="gx-loading">LOADING...</div></td></tr> :
                      events.length === 0 ? <tr><td colSpan={7}><div className="gx-empty">Connect to manager to see live data</div></td></tr> :
                        events.map((r, i) => (
                          <tr key={i}>
                            <td>{fmtDateTime(r.timestamp)}</td>
                            <td>{r.hostname}</td>
                            <td style={{ color: "var(--gx-muted)" }}>{r.agent_id}</td>
                            <td><span className="gx-mod-tag">{r.module}</span></td>
                            <td>{r.action}</td>
                            <td><span className={r.violation ? "gx-badge gx-badge-v" : "gx-badge gx-badge-ok"}>{r.violation ? "VIOLATION" : "OK"}</span></td>
                            <td style={{ color: "var(--gx-muted)", fontSize: 10 }}>{r.control_ref || ""}</td>
                          </tr>
                        ))}
                  </tbody>
                </table>
              </div>

              <div className="gx-pager">
                <span className="info">
                  {eventsTotal === 0 ? "0 events" :
                    `${formatNum(evOffset + 1)}–${formatNum(Math.min(evOffset + evLimit, eventsTotal))} of ${formatNum(eventsTotal)} events`}
                </span>
                {(() => {
                  const totalPages = Math.max(1, Math.ceil(eventsTotal / evLimit));
                  const curPage = Math.floor(evOffset / evLimit) + 1;
                  return (
                    <>
                      <button disabled={evOffset === 0} onClick={() => setEvOffset(0)}>« First</button>
                      <button disabled={evOffset === 0} onClick={() => setEvOffset(Math.max(0, evOffset - evLimit))}>← Prev</button>
                      <span>Page {curPage} / {totalPages}</span>
                      <button disabled={evOffset + evLimit >= eventsTotal} onClick={() => setEvOffset(evOffset + evLimit)}>Next →</button>
                      <button disabled={evOffset + evLimit >= eventsTotal} onClick={() => setEvOffset((totalPages - 1) * evLimit)}>Last »</button>
                    </>
                  );
                })()}
              </div>
            </>
          )}

          {tab === "alerts" && (
            <>
              <div className="gx-filter-bar">
                <span className="gx-filter-label">SEVERITY</span>
                <select value={alertFilters.severity} onChange={(e) => setAlertFilters({ ...alertFilters, severity: e.target.value })}>
                  <option value="">All</option>
                  <option value="CRITICAL">CRITICAL</option>
                  <option value="HIGH">HIGH</option>
                  <option value="MEDIUM">MEDIUM</option>
                  <option value="LOW">LOW</option>
                </select>
                <span className="gx-filter-label">SEARCH</span>
                <input type="text" value={alertFilters.search}
                  onChange={(e) => setAlertFilters({ ...alertFilters, search: e.target.value })}
                  placeholder="rule, host, agent id…" style={{ width: 220 }} />
                <button className="gx-btn-sm" onClick={() => setAlertFilters({ severity: "", search: "" })}>Clear</button>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table className="gx-tbl">
                  <thead><tr>
                    <th>Time</th><th>Rule</th><th>Severity</th><th>Host</th><th>Agent ID</th>
                  </tr></thead>
                  <tbody>
                    {alertsPaged.length === 0 ? <tr><td colSpan={5}><div className="gx-empty">No alerts match filters</div></td></tr> :
                      alertsPaged.map((a, i) => (
                        <tr key={i}>
                          <td>{fmtDateTime(a.timestamp)}</td>
                          <td>{a.rule}</td>
                          <td><span className={`gx-badge gx-badge-${a.severity}`}>{a.severity}</span></td>
                          <td>{a.hostname}</td>
                          <td style={{ color: "var(--gx-muted)" }}>{a.agent_id}</td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
              <div className="gx-pager">
                <span className="info">{alertsFiltered.length} of {alerts.length} alerts</span>
                <button disabled={alertsPage === 1} onClick={() => setAlertsPage(1)}>« First</button>
                <button disabled={alertsPage === 1} onClick={() => setAlertsPage(alertsPage - 1)}>← Prev</button>
                <span>Page {alertsPage} / {alertsTotalPages}</span>
                <button disabled={alertsPage >= alertsTotalPages} onClick={() => setAlertsPage(alertsPage + 1)}>Next →</button>
                <button disabled={alertsPage >= alertsTotalPages} onClick={() => setAlertsPage(alertsTotalPages)}>Last »</button>
              </div>
            </>
          )}

          {tab === "agents" && (
            <>
              <div className="gx-filter-bar">
                <button className="gx-btn-sm" onClick={() => void fetchAll()}>Refresh</button>
                <span className="gx-filter-label">STATUS</span>
                <select value={agentFilters.status} onChange={(e) => setAgentFilters({ ...agentFilters, status: e.target.value })}>
                  <option value="">All</option>
                  <option value="online">Online</option>
                  <option value="offline">Offline</option>
                </select>
                <span className="gx-filter-label">SEARCH</span>
                <input type="text" value={agentFilters.search}
                  onChange={(e) => setAgentFilters({ ...agentFilters, search: e.target.value })}
                  placeholder="host, id, ip, os…" style={{ width: 200 }} />
                <button className="gx-btn-sm" onClick={() => setAgentFilters({ status: "", search: "" })}>Clear</button>
                <span className="gx-filter-label">SCAN TARGET</span>
                <input type="text" value={scanTarget} onChange={(e) => setScanTarget(e.target.value)}
                  placeholder="https://target-to-scan" style={{ width: 220 }} />
                <button className="gx-btn-sm" onClick={() => void scanAll()}>SCAN ALL IPS</button>
              </div>
              <div className="gx-agent-summary">
                <span className="on">{onlineAgents.length} online</span>
                <span className="off">{agents.length - onlineAgents.length} offline</span>
                <span style={{ marginLeft: "auto" }}>Showing {agentsFiltered.length} of {agents.length}</span>
              </div>
              {agentsFiltered.length === 0 ? <div className="gx-empty">{agents.length === 0 ? "Connect to manager to see live data" : "No agents match filters"}</div> :
                agentsPaged.map((a) => {
                  const online = onlineAgents.includes(a);
                  return (
                    <div className="gx-agent-row" key={a.agent_id}>
                      <div className="gx-agent-left">
                        <div className="gx-status-indicator">
                          <div className={online ? "gx-dot-on" : "gx-dot-off"} />
                          <span className={online ? "gx-status-on" : "gx-status-off"}>{online ? "ONLINE" : "OFFLINE"}</span>
                        </div>
                        <div>
                          <div className="gx-agent-name">{a.hostname}</div>
                          <div className="gx-agent-id-text">{a.agent_id}</div>
                        </div>
                      </div>
                      <div className="gx-agent-right">
                        <span style={{ color: "var(--gx-muted)", fontSize: 10 }}>{a.os_type}</span>
                        <span style={{ color: "var(--gx-cyan-text)", fontSize: 11 }}>{formatNum(a.event_count)} events</span>
                        <span style={{ color: "var(--gx-muted)", fontSize: 10 }}>{fmtRelative(a.last_seen)}</span>
                        <span style={{ color: "var(--gx-muted)", fontSize: 10, fontFamily: "var(--gx-font)" }}>{a.ip_address || "—"}</span>
                        <button className="gx-btn-scan" onClick={() => void scanAgent(a)}>SCAN</button>
                        <button className="gx-btn-danger" onClick={() => void deleteAgent(a)}>DELETE</button>
                      </div>
                    </div>
                  );
                })}
              {agentsFiltered.length > 0 && (
                <div className="gx-pager">
                  <span className="info">{agentsFiltered.length} agents</span>
                  <button disabled={agentsPage === 1} onClick={() => setAgentsPage(1)}>« First</button>
                  <button disabled={agentsPage === 1} onClick={() => setAgentsPage(agentsPage - 1)}>← Prev</button>
                  <span>Page {agentsPage} / {agentsTotalPages}</span>
                  <button disabled={agentsPage >= agentsTotalPages} onClick={() => setAgentsPage(agentsPage + 1)}>Next →</button>
                  <button disabled={agentsPage >= agentsTotalPages} onClick={() => setAgentsPage(agentsTotalPages)}>Last »</button>
                </div>
              )}
            </>
          )}
        </div>
      </main>

      {/* Toasts */}
      <div className="gx-toast-wrap">
        {toasts.map((t) => <div key={t.id} className={`gx-toast ${t.type}`}>{t.msg}</div>)}
      </div>
    </div>
  );
}

// ---------- Sub-components ----------
function StatCard({ color, label, value, sub, trend }: {
  color: "cyan" | "red" | "yellow" | "green"; label: string;
  value: number | undefined; sub: string; trend: { dir: "up" | "down" | null; diff: number };
}) {
  return (
    <div className={`gx-stat-card c-${color}`}>
      <div className="gx-stat-label">{label}</div>
      <div className="gx-stat-value">{value === undefined || value === null ? "--" : formatNum(value)}</div>
      <div className="gx-stat-sub">{sub}</div>
      <div className={`gx-stat-trend ${trend.dir || ""}`}>
        {trend.dir === "up" ? `▲ +${trend.diff} since last sync` :
          trend.dir === "down" ? `▼ -${trend.diff} since last sync` : ""}
      </div>
    </div>
  );
}

function Panel({ title, right, children }: { title: React.ReactNode; right?: React.ReactNode; children: React.ReactNode }) {
  return (
    <div className="gx-panel">
      <div className="gx-panel-header">
        <span>{title}</span>
        <span className="gx-panel-header-right">{right}</span>
      </div>
      <div className="gx-panel-body">{children}</div>
    </div>
  );
}

function ChartTypeSelect({ value, onChange }: { value: ChartType; onChange: (v: ChartType) => void }) {
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value as ChartType)}
      title="Chart type"
      style={{
        background: "var(--gx-input-bg)", color: "var(--gx-text)",
        border: "1px solid var(--gx-border2)", padding: "2px 6px",
        fontFamily: "var(--gx-font)", fontSize: 10, marginLeft: 8,
        textTransform: "uppercase", letterSpacing: "0.1em", cursor: "pointer",
      }}
    >
      {CHART_TYPES.map((t) => (
        <option key={t} value={t}>{t.toUpperCase()}</option>
      ))}
    </select>
  );
}

function HeatmapGrid({ data, colorFor }: {
  data: { name: string; value: number }[];
  colorFor: (n: string) => string;
}) {
  const max = Math.max(1, ...data.map((d) => d.value));
  const cols = Math.min(Math.max(2, Math.ceil(Math.sqrt(data.length))), 8);
  return (
    <div style={{
      display: "grid", gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: 4,
      padding: 8, height: "100%", overflow: "auto",
    }}>
      {data.map((d) => {
        const intensity = d.value / max;
        return (
          <div key={d.name} title={`${d.name}: ${d.value}`}
            style={{
              background: colorFor(d.name),
              opacity: 0.25 + intensity * 0.75,
              borderRadius: 4, padding: "6px 8px",
              display: "flex", flexDirection: "column", justifyContent: "space-between",
              fontFamily: "Share Tech Mono, monospace", fontSize: 10, color: "#fff",
              minHeight: 48, boxShadow: "inset 0 0 0 1px rgba(255,255,255,0.05)",
            }}>
            <span style={{ opacity: 0.95, fontSize: 9, textTransform: "uppercase", letterSpacing: "0.05em" }}>{d.name}</span>
            <span style={{ fontWeight: 700, fontSize: 16 }}>{d.value}</span>
          </div>
        );
      })}
    </div>
  );
}

function FlexChart({ type, data, colorFor, baseColor, horizontal, axisColor, gridColor, tooltipStyle }: {
  type: ChartType;
  data: { name: string; value: number }[];
  colorFor: (n: string) => string;
  baseColor: string;
  horizontal?: boolean;
  axisColor: string; gridColor: string; tooltipStyle: React.CSSProperties;
}) {
  if (!data.length) return <div className="gx-empty">No data</div>;
  if (type === "heatmap") return <HeatmapGrid data={data} colorFor={colorFor} />;
  if (type === "radar") {
    return (
      <ResponsiveContainer><RadarChart data={data} outerRadius="75%">
        <PolarGrid stroke={gridColor} />
        <PolarAngleAxis dataKey="name" tick={{ fontSize: 10, fill: axisColor }} stroke={axisColor} />
        <PolarRadiusAxis tick={{ fontSize: 9, fill: axisColor }} stroke={axisColor} />
        <Tooltip contentStyle={tooltipStyle} />
        <Radar dataKey="value" stroke={baseColor} fill={baseColor} fillOpacity={0.35} />
      </RadarChart></ResponsiveContainer>
    );
  }
  if (type === "pie") {
    return (
      <ResponsiveContainer><PieChart>
        <Pie data={data} dataKey="value" nameKey="name" innerRadius={45} outerRadius={75} stroke="none">
          {data.map((d) => <Cell key={d.name} fill={colorFor(d.name)} />)}
        </Pie>
        <Tooltip contentStyle={tooltipStyle} />
        <Legend wrapperStyle={{ fontSize: 11, fontFamily: "Share Tech Mono, monospace", color: axisColor }} />
      </PieChart></ResponsiveContainer>
    );
  }
  if (type === "line") {
    return (
      <ResponsiveContainer><LineChart data={data} margin={{ top: 5, right: 8, left: -20, bottom: 0 }}>
        <CartesianGrid stroke={gridColor} vertical={false} />
        <XAxis dataKey="name" stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} />
        <YAxis stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} allowDecimals={false} />
        <Tooltip contentStyle={tooltipStyle} />
        <Line type="monotone" dataKey="value" stroke={baseColor} strokeWidth={2} dot={false} />
      </LineChart></ResponsiveContainer>
    );
  }
  if (type === "scatter") {
    const scat = data.map((d, i) => ({ ...d, idx: i }));
    return (
      <ResponsiveContainer><ScatterChart margin={{ top: 5, right: 8, left: -20, bottom: 0 }}>
        <CartesianGrid stroke={gridColor} />
        <XAxis type="number" dataKey="idx" stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} />
        <YAxis type="number" dataKey="value" stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} allowDecimals={false} />
        <ZAxis range={[60, 60]} />
        <Tooltip
          contentStyle={tooltipStyle}
          cursor={{ strokeDasharray: "3 3" }}
          formatter={(v: any, _n: any, p: any) => [v, p?.payload?.name]}
        />
        <Scatter data={scat} fill={baseColor}>
          {scat.map((d) => <Cell key={d.name} fill={colorFor(d.name)} />)}
        </Scatter>
      </ScatterChart></ResponsiveContainer>
    );
  }
  // bar
  if (horizontal) {
    return (
      <ResponsiveContainer><BarChart data={data} layout="vertical" margin={{ top: 5, right: 10, left: 10, bottom: 0 }}>
        <CartesianGrid stroke={gridColor} horizontal={false} />
        <XAxis type="number" stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} allowDecimals={false} />
        <YAxis type="category" dataKey="name" stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} width={90} />
        <Tooltip contentStyle={tooltipStyle} cursor={{ fill: "rgba(0,200,255,0.06)" }} />
        <Bar dataKey="value" fill={baseColor} fillOpacity={0.7}>
          {data.map((d) => <Cell key={d.name} fill={colorFor(d.name)} />)}
        </Bar>
      </BarChart></ResponsiveContainer>
    );
  }
  return (
    <ResponsiveContainer><BarChart data={data} margin={{ top: 5, right: 8, left: -20, bottom: 0 }}>
      <CartesianGrid stroke={gridColor} vertical={false} />
      <XAxis dataKey="name" stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} />
      <YAxis stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} allowDecimals={false} />
      <Tooltip contentStyle={tooltipStyle} cursor={{ fill: "rgba(255,51,85,0.08)" }} />
      <Bar dataKey="value" fill={baseColor} fillOpacity={0.6}>
        {data.map((d) => <Cell key={d.name} fill={colorFor(d.name)} />)}
      </Bar>
    </BarChart></ResponsiveContainer>
  );
}

function FlexTimeline({ type, data, axisColor, gridColor, tooltipStyle }: {
  type: ChartType;
  data: { date: string; events: number; violations: number }[];
  axisColor: string; gridColor: string; tooltipStyle: React.CSSProperties;
}) {
  if (!data.length) return <div className="gx-empty">No data</div>;
  if (type === "heatmap") {
    const max = Math.max(1, ...data.map((d) => d.events));
    return (
      <div style={{ padding: 8, height: "100%", overflow: "auto" }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(10, 1fr)", gap: 3 }}>
          {data.map((d) => {
            const t = d.events / max;
            return (
              <div key={d.date} title={`${d.date} • events ${d.events} • viol ${d.violations}`}
                style={{
                  background: "#00c8ff", opacity: 0.18 + t * 0.82,
                  borderRadius: 3, minHeight: 38,
                  display: "flex", flexDirection: "column", justifyContent: "space-between",
                  padding: "4px 5px", fontFamily: "Share Tech Mono, monospace",
                  fontSize: 9, color: "#001018",
                }}>
                <span>{d.date}</span>
                <span style={{ fontWeight: 700, color: "#fff" }}>{d.events}</span>
              </div>
            );
          })}
        </div>
      </div>
    );
  }
  if (type === "radar") {
    return (
      <ResponsiveContainer><RadarChart data={data} outerRadius="75%">
        <PolarGrid stroke={gridColor} />
        <PolarAngleAxis dataKey="date" tick={{ fontSize: 9, fill: axisColor }} stroke={axisColor} />
        <PolarRadiusAxis tick={{ fontSize: 9, fill: axisColor }} stroke={axisColor} />
        <Tooltip contentStyle={tooltipStyle} />
        <Legend wrapperStyle={{ fontSize: 11, fontFamily: "Share Tech Mono, monospace", color: axisColor }} />
        <Radar name="events" dataKey="events" stroke="#00c8ff" fill="#00c8ff" fillOpacity={0.3} />
        <Radar name="violations" dataKey="violations" stroke="#ff3355" fill="#ff3355" fillOpacity={0.3} />
      </RadarChart></ResponsiveContainer>
    );
  }
  if (type === "pie") {
    const totals = [
      { name: "events", value: data.reduce((s, d) => s + d.events, 0) },
      { name: "violations", value: data.reduce((s, d) => s + d.violations, 0) },
    ];
    const colors: Record<string, string> = { events: "#00c8ff", violations: "#ff3355" };
    return (
      <ResponsiveContainer><PieChart>
        <Pie data={totals} dataKey="value" nameKey="name" innerRadius={45} outerRadius={75} stroke="none">
          {totals.map((d) => <Cell key={d.name} fill={colors[d.name]} />)}
        </Pie>
        <Tooltip contentStyle={tooltipStyle} />
        <Legend wrapperStyle={{ fontSize: 11, fontFamily: "Share Tech Mono, monospace", color: axisColor }} />
      </PieChart></ResponsiveContainer>
    );
  }
  if (type === "scatter") {
    const ev = data.map((d, i) => ({ idx: i, value: d.events, date: d.date }));
    const vi = data.map((d, i) => ({ idx: i, value: d.violations, date: d.date }));
    return (
      <ResponsiveContainer><ScatterChart margin={{ top: 5, right: 8, left: -20, bottom: 0 }}>
        <CartesianGrid stroke={gridColor} />
        <XAxis type="number" dataKey="idx" stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} />
        <YAxis type="number" dataKey="value" stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} allowDecimals={false} />
        <ZAxis range={[50, 50]} />
        <Tooltip contentStyle={tooltipStyle} cursor={{ strokeDasharray: "3 3" }} />
        <Legend wrapperStyle={{ fontSize: 11, fontFamily: "Share Tech Mono, monospace", color: axisColor }} />
        <Scatter name="events" data={ev} fill="#00c8ff" />
        <Scatter name="violations" data={vi} fill="#ff3355" />
      </ScatterChart></ResponsiveContainer>
    );
  }
  if (type === "bar") {
    return (
      <ResponsiveContainer><BarChart data={data} margin={{ top: 5, right: 8, left: -20, bottom: 0 }}>
        <CartesianGrid stroke={gridColor} vertical={false} />
        <XAxis dataKey="date" stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} />
        <YAxis stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} allowDecimals={false} />
        <Tooltip contentStyle={tooltipStyle} />
        <Legend wrapperStyle={{ fontSize: 11, fontFamily: "Share Tech Mono, monospace", color: axisColor }} />
        <Bar dataKey="events" fill="#00c8ff" fillOpacity={0.6} />
        <Bar dataKey="violations" fill="#ff3355" fillOpacity={0.7} />
      </BarChart></ResponsiveContainer>
    );
  }
  // line (default)
  return (
    <ResponsiveContainer><LineChart data={data} margin={{ top: 5, right: 8, left: -20, bottom: 0 }}>
      <CartesianGrid stroke={gridColor} vertical={false} />
      <XAxis dataKey="date" stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} />
      <YAxis stroke={axisColor} tick={{ fontSize: 10, fill: axisColor }} allowDecimals={false} />
      <Tooltip contentStyle={tooltipStyle} />
      <Legend wrapperStyle={{ fontSize: 11, fontFamily: "Share Tech Mono, monospace", color: axisColor }} />
      <Line type="monotone" dataKey="events" stroke="#00c8ff" strokeWidth={2} dot={false} />
      <Line type="monotone" dataKey="violations" stroke="#ff3355" strokeWidth={2} dot={false} />
    </LineChart></ResponsiveContainer>
  );
}
