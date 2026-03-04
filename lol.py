import { useState, useCallback, useMemo, useRef, useEffect } from “react”;
import * as Papa from “papaparse”;
import * as XLSX from “sheetjs”;
import {
LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, AreaChart, Area,
XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
RadialBarChart, RadialBar, ComposedChart
} from “recharts”;
import {
Upload, FileSpreadsheet, TrendingUp, Users, Calendar, BarChart3,
ChevronDown, ChevronUp, AlertTriangle, CheckCircle, Clock, Zap,
ArrowUpRight, ArrowDownRight, Minus, Filter, Download, RefreshCw,
Layers, Activity, Target, Award, Sun, Moon
} from “lucide-react”;

// ─── COLOUR PALETTE ───
const P = {
bg: “#0B0F1A”,
card: “#111827”,
cardHover: “#1a2235”,
border: “#1E293B”,
borderLight: “#334155”,
text: “#E2E8F0”,
textMuted: “#94A3B8”,
textDim: “#64748B”,
accent: “#3B82F6”,
accentGlow: “rgba(59,130,246,0.15)”,
green: “#10B981”,
greenGlow: “rgba(16,185,129,0.15)”,
amber: “#F59E0B”,
amberGlow: “rgba(245,158,11,0.15)”,
red: “#EF4444”,
redGlow: “rgba(239,68,68,0.15)”,
purple: “#8B5CF6”,
purpleGlow: “rgba(139,92,246,0.15)”,
cyan: “#06B6D4”,
white: “#FFFFFF”,
gradient1: “linear-gradient(135deg, #3B82F6 0%, #8B5CF6 100%)”,
gradient2: “linear-gradient(135deg, #10B981 0%, #06B6D4 100%)”,
gradient3: “linear-gradient(135deg, #F59E0B 0%, #EF4444 100%)”,
};

const CHART_COLORS = [”#3B82F6”, “#10B981”, “#F59E0B”, “#EF4444”, “#8B5CF6”, “#06B6D4”, “#EC4899”, “#14B8A6”];

// ─── MOCK DATA GENERATOR ───
function generateMockData() {
const teams = [“Alpha-01”, “Bravo-02”, “Charlie-03”, “Delta-04”, “Echo-05”, “Foxtrot-06”];
const statuses = [“Site Clear”, “In Progress”, “Scheduled”, “On Hold”, “Cancelled”];
const workTypes = [“New Connection”, “Repair”, “Maintenance”, “Emergency”, “Upgrade”, “Inspection”];
const pms = [“Colin Donnelly”, “John Ashton”, “Ian Jones”, “Christine McNally”, “Steven Norton”, “James Tattersfield”];
const postcodes = [“LL57”, “LL55”, “CH5”, “CH6”, “CW1”, “WA1”, “BL1”, “BL3”, “M1”, “M4”];
const contracts = [“SPEN”, “ENW”];

const weeks = [];
const baseDate = new Date(2025, 1, 3);

for (let w = 0; w < 8; w++) {
const weekDate = new Date(baseDate);
weekDate.setDate(weekDate.getDate() + w * 7);
const jobCount = 80 + Math.floor(Math.random() * 60);
const jobs = [];

```
for (let j = 0; j < jobCount; j++) {
  const statusWeights = [0.35 + w * 0.02, 0.25, 0.2 - w * 0.01, 0.12, 0.08];
  let r = Math.random(), statusIdx = 0, cum = 0;
  for (let s = 0; s < statusWeights.length; s++) {
    cum += statusWeights[s];
    if (r < cum) { statusIdx = s; break; }
  }

  jobs.push({
    "Job ID": `JOB-${(w * 1000 + j + 10000).toString()}`,
    "Job Status": statuses[statusIdx],
    "Gang Ref": teams[Math.floor(Math.random() * teams.length)],
    "Work Type": workTypes[Math.floor(Math.random() * workTypes.length)],
    "Project Manager": pms[Math.floor(Math.random() * pms.length)],
    "Postcode": postcodes[Math.floor(Math.random() * postcodes.length)],
    "Contract": contracts[Math.floor(Math.random() * contracts.length)],
    "Date From": weekDate.toISOString().split("T")[0],
  });
}
weeks.push({ date: weekDate, jobs, label: weekDate.toLocaleDateString("en-GB", { day: "2-digit", month: "short" }) });
```

}
return weeks;
}

// ─── STAT CARD ───
function StatCard({ title, value, subtitle, icon: Icon, color, glow, trend, trendValue }) {
return (
<div style={{
background: P.card, borderRadius: 16, padding: “20px 24px”, border: `1px solid ${P.border}`,
position: “relative”, overflow: “hidden”, transition: “all 0.3s ease”,
cursor: “default”, minWidth: 0,
}}
onMouseEnter={e => { e.currentTarget.style.borderColor = color; e.currentTarget.style.boxShadow = `0 0 30px ${glow}`; }}
onMouseLeave={e => { e.currentTarget.style.borderColor = P.border; e.currentTarget.style.boxShadow = “none”; }}
>
<div style={{ position: “absolute”, top: -20, right: -20, width: 80, height: 80, borderRadius: “50%”, background: glow }} />
<div style={{ display: “flex”, justifyContent: “space-between”, alignItems: “flex-start”, position: “relative”, zIndex: 1 }}>
<div style={{ minWidth: 0 }}>
<div style={{ fontSize: 12, color: P.textMuted, textTransform: “uppercase”, letterSpacing: 1.5, marginBottom: 8, fontWeight: 600 }}>{title}</div>
<div style={{ fontSize: 32, fontWeight: 800, color: P.white, letterSpacing: -1, lineHeight: 1 }}>{value}</div>
{subtitle && <div style={{ fontSize: 12, color: P.textDim, marginTop: 6 }}>{subtitle}</div>}
{trend && (
<div style={{ display: “flex”, alignItems: “center”, gap: 4, marginTop: 8, fontSize: 12, fontWeight: 600,
color: trend === “up” ? P.green : trend === “down” ? P.red : P.textMuted }}>
{trend === “up” ? <ArrowUpRight size={14} /> : trend === “down” ? <ArrowDownRight size={14} /> : <Minus size={14} />}
{trendValue}
</div>
)}
</div>
<div style={{ width: 44, height: 44, borderRadius: 12, background: glow, display: “flex”, alignItems: “center”, justifyContent: “center”, flexShrink: 0 }}>
<Icon size={22} color={color} />
</div>
</div>
</div>
);
}

// ─── CHART CARD WRAPPER ───
function ChartCard({ title, subtitle, children, span = 1 }) {
return (
<div style={{
background: P.card, borderRadius: 16, border: `1px solid ${P.border}`, padding: 24,
gridColumn: `span ${span}`, minHeight: 320,
}}>
<div style={{ marginBottom: 20 }}>
<div style={{ fontSize: 16, fontWeight: 700, color: P.white }}>{title}</div>
{subtitle && <div style={{ fontSize: 12, color: P.textDim, marginTop: 4 }}>{subtitle}</div>}
</div>
{children}
</div>
);
}

// ─── MINI SPARKLINE ───
function Sparkline({ data, color, height = 40 }) {
if (!data || data.length < 2) return null;
const max = Math.max(…data);
const min = Math.min(…data);
const range = max - min || 1;
const w = 120, h = height;
const points = data.map((v, i) => `${(i / (data.length - 1)) * w},${h - ((v - min) / range) * h}`).join(” “);
return (
<svg width={w} height={h} style={{ display: “block” }}>
<polyline points={points} fill="none" stroke={color} strokeWidth={2} strokeLinejoin="round" strokeLinecap="round" />
</svg>
);
}

// ─── MAIN APP ───
export default function PlannedJobsAnalyser() {
const [data, setData] = useState(null);
const [selectedWeekIdx, setSelectedWeekIdx] = useState(null);
const [filterTeam, setFilterTeam] = useState(“All”);
const [filterPM, setFilterPM] = useState(“All”);
const [filterContract, setFilterContract] = useState(“All”);
const [activeTab, setActiveTab] = useState(“overview”);
const [isLoading, setIsLoading] = useState(false);
const fileInputRef = useRef(null);

// Load demo data
const loadDemo = useCallback(() => {
setIsLoading(true);
setTimeout(() => {
setData(generateMockData());
setIsLoading(false);
}, 600);
}, []);

// Parse uploaded files
const handleFiles = useCallback(async (e) => {
const files = Array.from(e.target.files);
if (!files.length) return;
setIsLoading(true);

```
try {
  const weeks = [];
  for (const file of files) {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    // Extract date from filename or first row
    let fileDate = new Date();
    const dateMatch = file.name.match(/(\d{4}[-_]?\d{2}[-_]?\d{2})/);
    if (dateMatch) {
      fileDate = new Date(dateMatch[1].replace(/_/g, "-"));
    } else if (json[0]?.["Date From"]) {
      fileDate = new Date(json[0]["Date From"]);
    }

    weeks.push({
      date: fileDate,
      jobs: json.map(row => {
        const clean = {};
        Object.keys(row).forEach(k => { clean[k.trim()] = row[k]; });
        return clean;
      }),
      label: fileDate.toLocaleDateString("en-GB", { day: "2-digit", month: "short" }),
      filename: file.name,
    });
  }

  weeks.sort((a, b) => a.date - b.date);
  setData(weeks);
} catch (err) {
  console.error("Parse error:", err);
}
setIsLoading(false);
```

}, []);

// ─── DERIVED ANALYTICS ───
const analytics = useMemo(() => {
if (!data) return null;

```
const allJobs = data.flatMap(w => w.jobs);
const totalJobs = allJobs.length;
const teams = [...new Set(allJobs.map(j => j["Gang Ref"]).filter(Boolean))].sort();
const pms = [...new Set(allJobs.map(j => j["Project Manager"]).filter(Boolean))].sort();
const contracts = [...new Set(allJobs.map(j => j["Contract"]).filter(Boolean))].sort();
const statuses = [...new Set(allJobs.map(j => j["Job Status"]).filter(Boolean))].sort();

// Apply filters
const filtered = allJobs.filter(j =>
  (filterTeam === "All" || j["Gang Ref"] === filterTeam) &&
  (filterPM === "All" || j["Project Manager"] === filterPM) &&
  (filterContract === "All" || j["Contract"] === filterContract)
);

// Status counts
const statusCounts = {};
filtered.forEach(j => { statusCounts[j["Job Status"]] = (statusCounts[j["Job Status"]] || 0) + 1; });

const completed = statusCounts["Site Clear"] || 0;
const inProgress = statusCounts["In Progress"] || 0;
const scheduled = statusCounts["Scheduled"] || 0;
const completionRate = filtered.length > 0 ? ((completed / filtered.length) * 100).toFixed(1) : 0;

// Weekly trend data
const weeklyTrend = data.map((w, i) => {
  const wFiltered = w.jobs.filter(j =>
    (filterTeam === "All" || j["Gang Ref"] === filterTeam) &&
    (filterPM === "All" || j["Project Manager"] === filterPM) &&
    (filterContract === "All" || j["Contract"] === filterContract)
  );
  const wStatuses = {};
  wFiltered.forEach(j => { wStatuses[j["Job Status"]] = (wStatuses[j["Job Status"]] || 0) + 1; });
  return {
    name: w.label,
    total: wFiltered.length,
    completed: wStatuses["Site Clear"] || 0,
    inProgress: wStatuses["In Progress"] || 0,
    scheduled: wStatuses["Scheduled"] || 0,
    onHold: wStatuses["On Hold"] || 0,
    completionRate: wFiltered.length > 0 ? +((wStatuses["Site Clear"] || 0) / wFiltered.length * 100).toFixed(1) : 0,
  };
});

// Trend direction
const recentTotal = weeklyTrend.slice(-2);
const totalTrend = recentTotal.length === 2
  ? (recentTotal[1].total > recentTotal[0].total ? "up" : recentTotal[1].total < recentTotal[0].total ? "down" : "flat")
  : "flat";
const totalTrendVal = recentTotal.length === 2
  ? `${Math.abs(recentTotal[1].total - recentTotal[0].total)} vs prev week`
  : "";

const compTrend = recentTotal.length === 2
  ? (recentTotal[1].completionRate > recentTotal[0].completionRate ? "up" : recentTotal[1].completionRate < recentTotal[0].completionRate ? "down" : "flat")
  : "flat";
const compTrendVal = recentTotal.length === 2
  ? `${Math.abs(recentTotal[1].completionRate - recentTotal[0].completionRate).toFixed(1)}pp vs prev`
  : "";

// Team performance
const teamPerf = {};
filtered.forEach(j => {
  const t = j["Gang Ref"] || "Unassigned";
  if (!teamPerf[t]) teamPerf[t] = { name: t, total: 0, completed: 0, inProgress: 0, weeklyJobs: {} };
  teamPerf[t].total++;
  if (j["Job Status"] === "Site Clear") teamPerf[t].completed++;
  if (j["Job Status"] === "In Progress") teamPerf[t].inProgress++;
});
// Build weekly sparkline per team
data.forEach((w, wi) => {
  w.jobs.forEach(j => {
    const t = j["Gang Ref"] || "Unassigned";
    if (teamPerf[t]) {
      teamPerf[t].weeklyJobs[wi] = (teamPerf[t].weeklyJobs[wi] || 0) + 1;
    }
  });
});
const teamData = Object.values(teamPerf).map(t => ({
  ...t,
  completionRate: t.total > 0 ? +((t.completed / t.total) * 100).toFixed(1) : 0,
  sparkline: data.map((_, wi) => t.weeklyJobs[wi] || 0),
})).sort((a, b) => b.total - a.total);

// PM performance
const pmPerf = {};
filtered.forEach(j => {
  const p = j["Project Manager"] || "Unassigned";
  if (!pmPerf[p]) pmPerf[p] = { name: p, total: 0, completed: 0 };
  pmPerf[p].total++;
  if (j["Job Status"] === "Site Clear") pmPerf[p].completed++;
});
const pmData = Object.values(pmPerf).map(p => ({
  ...p,
  completionRate: p.total > 0 ? +((p.completed / p.total) * 100).toFixed(1) : 0,
})).sort((a, b) => b.total - a.total);

// Work type breakdown
const workTypes = {};
filtered.forEach(j => {
  const wt = j["Work Type"] || "Other";
  workTypes[wt] = (workTypes[wt] || 0) + 1;
});
const workTypeData = Object.entries(workTypes).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);

// Status pie
const statusPie = Object.entries(statusCounts).map(([name, value]) => ({ name, value }));
const statusColorMap = { "Site Clear": P.green, "In Progress": P.accent, "Scheduled": P.amber, "On Hold": P.red, "Cancelled": P.textDim };

// Postcode heatmap data
const postcodeData = {};
filtered.forEach(j => {
  const pc = (j["Postcode"] || "").substring(0, 3) || "Unknown";
  postcodeData[pc] = (postcodeData[pc] || 0) + 1;
});
const postcodeChart = Object.entries(postcodeData).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value).slice(0, 12);

return {
  totalJobs, filtered: filtered.length, completed, inProgress, scheduled, completionRate,
  teams, pms, contracts, statuses, weeklyTrend, totalTrend, totalTrendVal, compTrend, compTrendVal,
  teamData, pmData, workTypeData, statusPie, statusColorMap, postcodeChart,
  dateRange: data.length > 1
    ? `${data[0].date.toLocaleDateString("en-GB")} — ${data[data.length - 1].date.toLocaleDateString("en-GB")}`
    : data[0]?.date.toLocaleDateString("en-GB") || "",
  weeksCount: data.length,
};
```

}, [data, filterTeam, filterPM, filterContract]);

// ─── RENDER: UPLOAD SCREEN ───
if (!data) {
return (
<div style={{
minHeight: “100vh”, background: P.bg, display: “flex”, flexDirection: “column”,
alignItems: “center”, justifyContent: “center”, fontFamily: “‘Segoe UI’, system-ui, sans-serif”, padding: 20,
}}>
{/* Animated background */}
<div style={{ position: “fixed”, inset: 0, overflow: “hidden”, pointerEvents: “none”, zIndex: 0 }}>
<div style={{
position: “absolute”, width: 600, height: 600, borderRadius: “50%”,
background: “radial-gradient(circle, rgba(59,130,246,0.08) 0%, transparent 70%)”,
top: “10%”, left: “20%”, animation: “float 20s ease-in-out infinite”,
}} />
<div style={{
position: “absolute”, width: 400, height: 400, borderRadius: “50%”,
background: “radial-gradient(circle, rgba(139,92,246,0.06) 0%, transparent 70%)”,
bottom: “20%”, right: “15%”, animation: “float 15s ease-in-out infinite reverse”,
}} />
</div>

```
    <div style={{ position: "relative", zIndex: 1, textAlign: "center", maxWidth: 560 }}>
      <div style={{
        width: 80, height: 80, borderRadius: 24, background: P.accentGlow, border: `2px solid ${P.accent}`,
        display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 32px",
        boxShadow: `0 0 60px ${P.accentGlow}`,
      }}>
        <BarChart3 size={36} color={P.accent} />
      </div>

      <h1 style={{ fontSize: 42, fontWeight: 800, color: P.white, margin: 0, letterSpacing: -1.5, lineHeight: 1.1 }}>
        Multi-Week<br />
        <span style={{ background: P.gradient1, WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>
          Jobs Analyser
        </span>
      </h1>
      <p style={{ fontSize: 16, color: P.textMuted, marginTop: 16, lineHeight: 1.6 }}>
        Drop your weekly Excel exports to unlock trend analytics,<br />team performance tracking, and completion forecasting.
      </p>

      <div style={{ display: "flex", gap: 16, marginTop: 40, justifyContent: "center", flexWrap: "wrap" }}>
        <button onClick={() => fileInputRef.current?.click()} style={{
          padding: "16px 32px", borderRadius: 14, border: "none", background: P.gradient1,
          color: P.white, fontSize: 16, fontWeight: 700, cursor: "pointer",
          display: "flex", alignItems: "center", gap: 10,
          boxShadow: "0 4px 20px rgba(59,130,246,0.3)", transition: "transform 0.2s, box-shadow 0.2s",
        }}
          onMouseEnter={e => { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = "0 8px 30px rgba(59,130,246,0.4)"; }}
          onMouseLeave={e => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "0 4px 20px rgba(59,130,246,0.3)"; }}
        >
          <Upload size={20} /> Upload Excel Files
        </button>

        <button onClick={loadDemo} style={{
          padding: "16px 32px", borderRadius: 14, border: `2px solid ${P.border}`, background: "transparent",
          color: P.textMuted, fontSize: 16, fontWeight: 600, cursor: "pointer",
          display: "flex", alignItems: "center", gap: 10, transition: "all 0.2s",
        }}
          onMouseEnter={e => { e.currentTarget.style.borderColor = P.accent; e.currentTarget.style.color = P.white; }}
          onMouseLeave={e => { e.currentTarget.style.borderColor = P.border; e.currentTarget.style.color = P.textMuted; }}
        >
          <Zap size={20} /> Load Demo Data
        </button>
      </div>

      <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" multiple
        onChange={handleFiles} style={{ display: "none" }} />

      <div style={{ display: "flex", gap: 32, marginTop: 48, justifyContent: "center", color: P.textDim, fontSize: 13 }}>
        {[["8-Week Trends", TrendingUp], ["Team Analytics", Users], ["Completion Rates", Target]].map(([label, Icon]) => (
          <div key={label} style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <Icon size={14} /> {label}
          </div>
        ))}
      </div>
    </div>

    {isLoading && (
      <div style={{
        position: "fixed", inset: 0, background: "rgba(11,15,26,0.9)", display: "flex",
        alignItems: "center", justifyContent: "center", zIndex: 100,
      }}>
        <div style={{ textAlign: "center" }}>
          <div style={{
            width: 48, height: 48, border: `3px solid ${P.border}`, borderTop: `3px solid ${P.accent}`,
            borderRadius: "50%", animation: "spin 0.8s linear infinite", margin: "0 auto 16px",
          }} />
          <div style={{ color: P.textMuted, fontSize: 14 }}>Processing files...</div>
        </div>
      </div>
    )}

    <style>{`
      @keyframes float { 0%, 100% { transform: translate(0, 0); } 50% { transform: translate(30px, -30px); } }
      @keyframes spin { to { transform: rotate(360deg); } }
    `}</style>
  </div>
);
```

}

// ─── RENDER: DASHBOARD ───
const a = analytics;
const tabs = [
{ id: “overview”, label: “Overview”, icon: Layers },
{ id: “trends”, label: “Trends”, icon: TrendingUp },
{ id: “teams”, label: “Teams”, icon: Users },
{ id: “managers”, label: “Managers”, icon: Award },
{ id: “breakdown”, label: “Breakdown”, icon: BarChart3 },
];

const CustomTooltip = ({ active, payload, label }) => {
if (!active || !payload?.length) return null;
return (
<div style={{
background: “#1E293B”, border: `1px solid ${P.border}`, borderRadius: 10, padding: “10px 14px”,
boxShadow: “0 8px 24px rgba(0,0,0,0.4)”,
}}>
<div style={{ fontSize: 12, color: P.textMuted, marginBottom: 6 }}>{label}</div>
{payload.map((p, i) => (
<div key={i} style={{ display: “flex”, alignItems: “center”, gap: 8, fontSize: 13, color: P.text, marginTop: 2 }}>
<div style={{ width: 8, height: 8, borderRadius: “50%”, background: p.color }} />
{p.name}: <span style={{ fontWeight: 700, color: P.white }}>{p.value}</span>
</div>
))}
</div>
);
};

return (
<div style={{
minHeight: “100vh”, background: P.bg, fontFamily: “‘Segoe UI’, system-ui, sans-serif”, color: P.text,
}}>
{/* ─── HEADER ─── */}
<div style={{
background: “rgba(17,24,39,0.8)”, backdropFilter: “blur(20px)”, borderBottom: `1px solid ${P.border}`,
position: “sticky”, top: 0, zIndex: 50, padding: “0 24px”,
}}>
<div style={{ maxWidth: 1400, margin: “0 auto”, display: “flex”, alignItems: “center”, justifyContent: “space-between”, height: 64 }}>
<div style={{ display: “flex”, alignItems: “center”, gap: 12 }}>
<div style={{ width: 36, height: 36, borderRadius: 10, background: P.accentGlow, border: `1.5px solid ${P.accent}`,
display: “flex”, alignItems: “center”, justifyContent: “center” }}>
<BarChart3 size={18} color={P.accent} />
</div>
<div>
<div style={{ fontSize: 16, fontWeight: 700, color: P.white }}>Jobs Analyser</div>
<div style={{ fontSize: 11, color: P.textDim }}>{a.dateRange} · {a.weeksCount} weeks · {a.filtered.toLocaleString()} jobs</div>
</div>
</div>

```
      {/* Filters */}
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        {[
          { label: "Team", value: filterTeam, setter: setFilterTeam, options: a.teams },
          { label: "PM", value: filterPM, setter: setFilterPM, options: a.pms },
          { label: "Contract", value: filterContract, setter: setFilterContract, options: a.contracts },
        ].map(f => (
          <select key={f.label} value={f.value} onChange={e => f.setter(e.target.value)} style={{
            background: "#1E293B", border: `1px solid ${P.border}`, borderRadius: 8, padding: "6px 12px",
            color: P.text, fontSize: 12, cursor: "pointer", outline: "none",
          }}>
            <option value="All">All {f.label}s</option>
            {f.options.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        ))}

        <button onClick={() => { setData(null); setFilterTeam("All"); setFilterPM("All"); setFilterContract("All"); }} style={{
          background: "transparent", border: `1px solid ${P.border}`, borderRadius: 8,
          padding: "6px 12px", color: P.textMuted, fontSize: 12, cursor: "pointer",
          display: "flex", alignItems: "center", gap: 4,
        }}>
          <RefreshCw size={12} /> New Analysis
        </button>
      </div>
    </div>
  </div>

  {/* ─── TABS ─── */}
  <div style={{ maxWidth: 1400, margin: "0 auto", padding: "16px 24px 0" }}>
    <div style={{ display: "flex", gap: 4, borderBottom: `1px solid ${P.border}`, paddingBottom: 0 }}>
      {tabs.map(t => (
        <button key={t.id} onClick={() => setActiveTab(t.id)} style={{
          padding: "10px 20px", border: "none", borderBottom: activeTab === t.id ? `2px solid ${P.accent}` : "2px solid transparent",
          background: "transparent", color: activeTab === t.id ? P.white : P.textDim,
          fontSize: 13, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: 6,
          transition: "all 0.2s",
        }}>
          <t.icon size={15} /> {t.label}
        </button>
      ))}
    </div>
  </div>

  {/* ─── CONTENT ─── */}
  <div style={{ maxWidth: 1400, margin: "0 auto", padding: "24px" }}>

    {/* OVERVIEW TAB */}
    {activeTab === "overview" && (
      <>
        {/* Stat cards */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 16, marginBottom: 24 }}>
          <StatCard title="Total Jobs" value={a.filtered.toLocaleString()} subtitle={`Across ${a.weeksCount} weeks`}
            icon={FileSpreadsheet} color={P.accent} glow={P.accentGlow} trend={a.totalTrend} trendValue={a.totalTrendVal} />
          <StatCard title="Completion Rate" value={`${a.completionRate}%`} subtitle={`${a.completed} completed`}
            icon={Target} color={P.green} glow={P.greenGlow} trend={a.compTrend} trendValue={a.compTrendVal} />
          <StatCard title="In Progress" value={a.inProgress.toLocaleString()} subtitle="Active right now"
            icon={Activity} color={P.amber} glow={P.amberGlow} />
          <StatCard title="Active Teams" value={a.teamData.length} subtitle="Across all weeks"
            icon={Users} color={P.purple} glow={P.purpleGlow} />
        </div>

        {/* Charts row */}
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 16, marginBottom: 16 }}>
          <ChartCard title="Weekly Job Volume & Completion" subtitle="Stacked area showing job pipeline over time">
            <ResponsiveContainer width="100%" height={260}>
              <AreaChart data={a.weeklyTrend}>
                <defs>
                  <linearGradient id="gComplete" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor={P.green} stopOpacity={0.3} />
                    <stop offset="95%" stopColor={P.green} stopOpacity={0} />
                  </linearGradient>
                  <linearGradient id="gProgress" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor={P.accent} stopOpacity={0.3} />
                    <stop offset="95%" stopColor={P.accent} stopOpacity={0} />
                  </linearGradient>
                  <linearGradient id="gScheduled" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor={P.amber} stopOpacity={0.3} />
                    <stop offset="95%" stopColor={P.amber} stopOpacity={0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis dataKey="name" tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <YAxis tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Area type="monotone" dataKey="completed" stackId="1" stroke={P.green} fill="url(#gComplete)" name="Completed" />
                <Area type="monotone" dataKey="inProgress" stackId="1" stroke={P.accent} fill="url(#gProgress)" name="In Progress" />
                <Area type="monotone" dataKey="scheduled" stackId="1" stroke={P.amber} fill="url(#gScheduled)" name="Scheduled" />
              </AreaChart>
            </ResponsiveContainer>
          </ChartCard>

          <ChartCard title="Status Distribution" subtitle="Current split across all weeks">
            <ResponsiveContainer width="100%" height={260}>
              <PieChart>
                <Pie data={a.statusPie} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={55} outerRadius={90}
                  paddingAngle={3} stroke="none">
                  {a.statusPie.map((entry, i) => (
                    <Cell key={i} fill={a.statusColorMap[entry.name] || CHART_COLORS[i % CHART_COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 11 }} />
              </PieChart>
            </ResponsiveContainer>
          </ChartCard>
        </div>

        {/* Bottom row */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
          <ChartCard title="Top Teams by Volume" subtitle="Total jobs assigned per team">
            <ResponsiveContainer width="100%" height={260}>
              <BarChart data={a.teamData.slice(0, 8)} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis type="number" tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <YAxis dataKey="name" type="category" tick={{ fill: P.textDim, fontSize: 11 }} width={80} axisLine={{ stroke: P.border }} />
                <Tooltip content={<CustomTooltip />} />
                <Bar dataKey="completed" stackId="a" fill={P.green} name="Completed" radius={[0, 0, 0, 0]} />
                <Bar dataKey="inProgress" stackId="a" fill={P.accent} name="In Progress" />
                <Bar dataKey="total" fill="none" name="" />
              </BarChart>
            </ResponsiveContainer>
          </ChartCard>

          <ChartCard title="Jobs by Postcode Area" subtitle="Geographic distribution of work">
            <ResponsiveContainer width="100%" height={260}>
              <BarChart data={a.postcodeChart}>
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis dataKey="name" tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <YAxis tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <Tooltip content={<CustomTooltip />} />
                <Bar dataKey="value" name="Jobs" radius={[6, 6, 0, 0]}>
                  {a.postcodeChart.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </ChartCard>
        </div>
      </>
    )}

    {/* TRENDS TAB */}
    {activeTab === "trends" && (
      <>
        <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 16, marginBottom: 16 }}>
          <ChartCard title="Completion Rate Trend" subtitle="Weekly completion percentage trajectory" span={1}>
            <ResponsiveContainer width="100%" height={300}>
              <ComposedChart data={a.weeklyTrend}>
                <defs>
                  <linearGradient id="gRate" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor={P.green} stopOpacity={0.3} />
                    <stop offset="95%" stopColor={P.green} stopOpacity={0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis dataKey="name" tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <YAxis yAxisId="left" tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <YAxis yAxisId="right" orientation="right" tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} domain={[0, 100]} unit="%" />
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Bar yAxisId="left" dataKey="total" fill={P.accent} name="Total Jobs" radius={[4, 4, 0, 0]} opacity={0.6} />
                <Line yAxisId="right" type="monotone" dataKey="completionRate" stroke={P.green} strokeWidth={3}
                  name="Completion %" dot={{ fill: P.green, r: 5, strokeWidth: 2, stroke: P.card }} />
              </ComposedChart>
            </ResponsiveContainer>
          </ChartCard>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
          <ChartCard title="Job Volume Trend" subtitle="Total jobs per reporting period">
            <ResponsiveContainer width="100%" height={260}>
              <LineChart data={a.weeklyTrend}>
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis dataKey="name" tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <YAxis tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <Tooltip content={<CustomTooltip />} />
                <Line type="monotone" dataKey="total" stroke={P.accent} strokeWidth={2.5} name="Total"
                  dot={{ fill: P.accent, r: 4, strokeWidth: 2, stroke: P.card }} />
                <Line type="monotone" dataKey="completed" stroke={P.green} strokeWidth={2} name="Completed" strokeDasharray="5 5"
                  dot={{ fill: P.green, r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          </ChartCard>

          <ChartCard title="On Hold / Cancelled Trend" subtitle="Tracking blockers over time">
            <ResponsiveContainer width="100%" height={260}>
              <AreaChart data={a.weeklyTrend}>
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis dataKey="name" tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <YAxis tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <Tooltip content={<CustomTooltip />} />
                <Area type="monotone" dataKey="onHold" stroke={P.red} fill={P.redGlow} name="On Hold" />
              </AreaChart>
            </ResponsiveContainer>
          </ChartCard>
        </div>

        {/* Week-over-week table */}
        <div style={{ marginTop: 16 }}>
          <ChartCard title="Week-over-Week Comparison" subtitle="Detailed weekly breakdown">
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0, fontSize: 13 }}>
                <thead>
                  <tr>
                    {["Week", "Total", "Completed", "In Progress", "Scheduled", "On Hold", "Completion %", "Δ Total", "Δ Rate"].map(h => (
                      <th key={h} style={{
                        padding: "10px 14px", textAlign: "left", borderBottom: `2px solid ${P.border}`,
                        color: P.textMuted, fontSize: 11, textTransform: "uppercase", letterSpacing: 1,
                        fontWeight: 600, whiteSpace: "nowrap",
                      }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {a.weeklyTrend.map((w, i) => {
                    const prev = i > 0 ? a.weeklyTrend[i - 1] : null;
                    const deltaTotal = prev ? w.total - prev.total : 0;
                    const deltaRate = prev ? (w.completionRate - prev.completionRate).toFixed(1) : 0;
                    return (
                      <tr key={i} style={{ borderBottom: `1px solid ${P.border}` }}
                        onMouseEnter={e => e.currentTarget.style.background = "rgba(59,130,246,0.05)"}
                        onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                        <td style={{ padding: "10px 14px", fontWeight: 600, color: P.white }}>{w.name}</td>
                        <td style={{ padding: "10px 14px", fontWeight: 700, color: P.white }}>{w.total}</td>
                        <td style={{ padding: "10px 14px", color: P.green }}>{w.completed}</td>
                        <td style={{ padding: "10px 14px", color: P.accent }}>{w.inProgress}</td>
                        <td style={{ padding: "10px 14px", color: P.amber }}>{w.scheduled}</td>
                        <td style={{ padding: "10px 14px", color: P.red }}>{w.onHold}</td>
                        <td style={{ padding: "10px 14px" }}>
                          <span style={{
                            padding: "2px 10px", borderRadius: 12, fontSize: 12, fontWeight: 600,
                            background: w.completionRate >= 40 ? P.greenGlow : w.completionRate >= 25 ? P.amberGlow : P.redGlow,
                            color: w.completionRate >= 40 ? P.green : w.completionRate >= 25 ? P.amber : P.red,
                          }}>{w.completionRate}%</span>
                        </td>
                        <td style={{ padding: "10px 14px", color: deltaTotal > 0 ? P.green : deltaTotal < 0 ? P.red : P.textDim, fontWeight: 600 }}>
                          {i > 0 ? (deltaTotal > 0 ? "+" : "") + deltaTotal : "—"}
                        </td>
                        <td style={{ padding: "10px 14px", color: deltaRate > 0 ? P.green : deltaRate < 0 ? P.red : P.textDim, fontWeight: 600 }}>
                          {i > 0 ? (deltaRate > 0 ? "+" : "") + deltaRate + "pp" : "—"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </ChartCard>
        </div>
      </>
    )}

    {/* TEAMS TAB */}
    {activeTab === "teams" && (
      <>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
          <ChartCard title="Team Completion Rates" subtitle="Percentage of jobs completed per team">
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={a.teamData} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis type="number" domain={[0, 100]} tick={{ fill: P.textDim, fontSize: 11 }} unit="%" axisLine={{ stroke: P.border }} />
                <YAxis dataKey="name" type="category" tick={{ fill: P.textDim, fontSize: 11 }} width={90} axisLine={{ stroke: P.border }} />
                <Tooltip content={<CustomTooltip />} />
                <Bar dataKey="completionRate" name="Completion %" radius={[0, 6, 6, 0]}>
                  {a.teamData.map((t, i) => (
                    <Cell key={i} fill={t.completionRate >= 40 ? P.green : t.completionRate >= 25 ? P.amber : P.red} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </ChartCard>

          <ChartCard title="Team Workload Distribution" subtitle="Jobs per team breakdown">
            <ResponsiveContainer width="100%" height={280}>
              <PieChart>
                <Pie data={a.teamData} dataKey="total" nameKey="name" cx="50%" cy="50%" innerRadius={50} outerRadius={95}
                  paddingAngle={2} stroke="none">
                  {a.teamData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Pie>
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 11 }} />
              </PieChart>
            </ResponsiveContainer>
          </ChartCard>
        </div>

        {/* Team detail table */}
        <ChartCard title="Team Performance Detail" subtitle="Full breakdown with sparkline trends">
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "separate", borderSpacing: 0, fontSize: 13 }}>
              <thead>
                <tr>
                  {["Team", "Total Jobs", "Completed", "In Progress", "Completion %", "Trend"].map(h => (
                    <th key={h} style={{
                      padding: "10px 14px", textAlign: "left", borderBottom: `2px solid ${P.border}`,
                      color: P.textMuted, fontSize: 11, textTransform: "uppercase", letterSpacing: 1, fontWeight: 600,
                    }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {a.teamData.map((t, i) => (
                  <tr key={i} onMouseEnter={e => e.currentTarget.style.background = "rgba(59,130,246,0.05)"}
                    onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                    <td style={{ padding: "10px 14px", fontWeight: 600, color: P.white }}>{t.name}</td>
                    <td style={{ padding: "10px 14px", fontWeight: 700, color: P.white }}>{t.total}</td>
                    <td style={{ padding: "10px 14px", color: P.green }}>{t.completed}</td>
                    <td style={{ padding: "10px 14px", color: P.accent }}>{t.inProgress}</td>
                    <td style={{ padding: "10px 14px" }}>
                      <span style={{
                        padding: "2px 10px", borderRadius: 12, fontSize: 12, fontWeight: 600,
                        background: t.completionRate >= 40 ? P.greenGlow : t.completionRate >= 25 ? P.amberGlow : P.redGlow,
                        color: t.completionRate >= 40 ? P.green : t.completionRate >= 25 ? P.amber : P.red,
                      }}>{t.completionRate}%</span>
                    </td>
                    <td style={{ padding: "10px 14px" }}>
                      <Sparkline data={t.sparkline} color={CHART_COLORS[i % CHART_COLORS.length]} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </ChartCard>
      </>
    )}

    {/* MANAGERS TAB */}
    {activeTab === "managers" && (
      <>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
          <ChartCard title="PM Job Volumes" subtitle="Total jobs assigned to each project manager">
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={a.pmData}>
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis dataKey="name" tick={{ fill: P.textDim, fontSize: 10 }} axisLine={{ stroke: P.border }} angle={-20} textAnchor="end" height={60} />
                <YAxis tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <Tooltip content={<CustomTooltip />} />
                <Bar dataKey="total" name="Total Jobs" radius={[6, 6, 0, 0]}>
                  {a.pmData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </ChartCard>

          <ChartCard title="PM Completion Rates" subtitle="Percentage completed per manager">
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={a.pmData} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis type="number" domain={[0, 100]} unit="%" tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <YAxis dataKey="name" type="category" tick={{ fill: P.textDim, fontSize: 10 }} width={130} axisLine={{ stroke: P.border }} />
                <Tooltip content={<CustomTooltip />} />
                <Bar dataKey="completionRate" name="Completion %" radius={[0, 6, 6, 0]}>
                  {a.pmData.map((p, i) => (
                    <Cell key={i} fill={p.completionRate >= 40 ? P.green : p.completionRate >= 25 ? P.amber : P.red} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </ChartCard>
        </div>

        {/* PM Leaderboard */}
        <ChartCard title="Project Manager Leaderboard" subtitle="Ranked by completion rate">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 12, marginTop: 8 }}>
            {a.pmData.sort((a, b) => b.completionRate - a.completionRate).map((pm, i) => (
              <div key={pm.name} style={{
                background: i === 0 ? "rgba(16,185,129,0.08)" : "rgba(30,41,59,0.5)",
                border: `1px solid ${i === 0 ? "rgba(16,185,129,0.3)" : P.border}`,
                borderRadius: 12, padding: 16, display: "flex", alignItems: "center", gap: 14,
              }}>
                <div style={{
                  width: 36, height: 36, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center",
                  background: i === 0 ? P.greenGlow : i === 1 ? P.accentGlow : i === 2 ? P.amberGlow : "rgba(100,116,139,0.1)",
                  color: i === 0 ? P.green : i === 1 ? P.accent : i === 2 ? P.amber : P.textDim,
                  fontWeight: 800, fontSize: 14,
                }}>{i + 1}</div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: P.white, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{pm.name}</div>
                  <div style={{ fontSize: 11, color: P.textDim }}>{pm.total} jobs · {pm.completed} completed</div>
                </div>
                <div style={{
                  fontSize: 16, fontWeight: 800,
                  color: pm.completionRate >= 40 ? P.green : pm.completionRate >= 25 ? P.amber : P.red,
                }}>{pm.completionRate}%</div>
              </div>
            ))}
          </div>
        </ChartCard>
      </>
    )}

    {/* BREAKDOWN TAB */}
    {activeTab === "breakdown" && (
      <>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
          <ChartCard title="Work Type Distribution" subtitle="Jobs categorised by work type">
            <ResponsiveContainer width="100%" height={280}>
              <PieChart>
                <Pie data={a.workTypeData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={95}
                  paddingAngle={2} stroke="none">
                  {a.workTypeData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Pie>
                <Tooltip content={<CustomTooltip />} />
                <Legend wrapperStyle={{ fontSize: 11 }} />
              </PieChart>
            </ResponsiveContainer>
          </ChartCard>

          <ChartCard title="Work Type Volume" subtitle="Bar breakdown of work categories">
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={a.workTypeData}>
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis dataKey="name" tick={{ fill: P.textDim, fontSize: 10 }} axisLine={{ stroke: P.border }} angle={-15} textAnchor="end" height={60} />
                <YAxis tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <Tooltip content={<CustomTooltip />} />
                <Bar dataKey="value" name="Jobs" radius={[6, 6, 0, 0]}>
                  {a.workTypeData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </ChartCard>
        </div>

        {/* Geographic & Contract */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
          <ChartCard title="Geographic Spread" subtitle="Jobs by postcode area">
            <ResponsiveContainer width="100%" height={260}>
              <BarChart data={a.postcodeChart} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" stroke={P.border} />
                <XAxis type="number" tick={{ fill: P.textDim, fontSize: 11 }} axisLine={{ stroke: P.border }} />
                <YAxis dataKey="name" type="category" tick={{ fill: P.textDim, fontSize: 11 }} width={50} axisLine={{ stroke: P.border }} />
                <Tooltip content={<CustomTooltip />} />
                <Bar dataKey="value" name="Jobs" radius={[0, 6, 6, 0]}>
                  {a.postcodeChart.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </ChartCard>

          <ChartCard title="Weekly Volume Heatmap" subtitle="Job counts per week summary">
            <div style={{ display: "grid", gridTemplateColumns: `repeat(${Math.min(a.weeklyTrend.length, 8)}, 1fr)`, gap: 8, marginTop: 8 }}>
              {a.weeklyTrend.map((w, i) => {
                const maxTotal = Math.max(...a.weeklyTrend.map(x => x.total));
                const intensity = maxTotal > 0 ? w.total / maxTotal : 0;
                return (
                  <div key={i} style={{
                    borderRadius: 12, padding: 14, textAlign: "center",
                    background: `rgba(59,130,246,${0.05 + intensity * 0.2})`,
                    border: `1px solid rgba(59,130,246,${0.1 + intensity * 0.3})`,
                  }}>
                    <div style={{ fontSize: 11, color: P.textDim, marginBottom: 6 }}>{w.name}</div>
                    <div style={{ fontSize: 24, fontWeight: 800, color: P.white }}>{w.total}</div>
                    <div style={{ fontSize: 11, color: P.green, marginTop: 4 }}>{w.completionRate}%</div>
                  </div>
                );
              })}
            </div>
          </ChartCard>
        </div>
      </>
    )}
  </div>

  {/* Footer */}
  <div style={{
    textAlign: "center", padding: "24px", color: P.textDim, fontSize: 11,
    borderTop: `1px solid ${P.border}`, marginTop: 24,
  }}>
    Multi-Week Jobs Analyser v2.0 · Built for SPEN & ENW Operations
  </div>
</div>
```

);
}
