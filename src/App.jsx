import { useState, useEffect, useCallback } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { loginRequest } from "./services/authConfig";
import { fetchProjects } from "./services/powerbiApi";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";

const T = {
  navy: "#0D1F3C", navyMid: "#1a3260",
  gold: "#C9A84C", red: "#C0392B",
  green: "#1D9E75", amber: "#E67E22",
  bg: "#F4F5F7", surface: "#FFFFFF",
  border: "#E2E4E9", text: "#1a1a2e", muted: "#6B7280"
};

const fmt$ = n => Number(n) > 0 ? "$" + Number(n).toLocaleString() : "—";
const fmtPct = n => Math.round(Number(n) || 0) + "%";
const num = v => Number(v) || 0;

function PhaseBadge({ phase }) {
  const s = { CONSTRUCTION: { background: "#FEF5E7", color: "#7D4E0F" }, PRECON: { background: "#EBF0F8", color: "#1a3260" } }[phase] || { background: "#F1F3F5", color: T.muted };
  return <span style={{ ...s, padding: "2px 8px", borderRadius: 3, fontSize: 10, fontWeight: 700, textTransform: "uppercase", display: "inline-block" }}>{phase || "—"}</span>;
}

function HealthDot({ health }) {
  const c = { Green: T.green, Red: T.red, Amber: T.amber }[health] || "#9CA3AF";
  return <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 8, height: 8, borderRadius: "50%", background: c, display: "inline-block" }} />{health || "—"}</span>;
}

function KpiCard({ label, value, sub, accent }) {
  return (
    <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, padding: "16px 18px", position: "relative", overflow: "hidden" }}>
      <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: 3, background: accent || T.navy }} />
      <div style={{ fontSize: 10, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.6, color: T.muted, marginBottom: 8 }}>{label}</div>
      <div style={{ fontSize: 32, fontWeight: 300, color: accent || T.navy, lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: T.muted, marginTop: 6 }}>{sub}</div>}
    </div>
  );
}

function Panel({ title, count, children, noPad }) {
  return (
    <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, overflow: "hidden" }}>
      <div style={{ background: T.navy, color: "white", padding: "10px 16px", fontSize: 11, fontWeight: 600, letterSpacing: 0.8, textTransform: "uppercase", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        {title}
        {count !== undefined && <span style={{ background: T.gold, color: T.navy, fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 10 }}>{count}</span>}
      </div>
      <div style={noPad ? {} : { padding: "14px 16px" }}>{children}</div>
    </div>
  );
}

function ProgressBar({ pct }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
      <div style={{ flex: 1, height: 4, background: T.bg, borderRadius: 2, overflow: "hidden" }}>
        <div style={{ width: `${Math.min(pct, 100)}%`, height: "100%", background: T.navy, borderRadius: 2 }} />
      </div>
      <span style={{ fontSize: 11, fontWeight: 600, minWidth: 30 }}>{fmtPct(pct)}</span>
    </div>
  );
}

function LoginPage({ onLogin }) {
  return (
    <div style={{ minHeight: "100vh", background: T.navy, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 24 }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 3, alignItems: "center" }}>
        <span style={{ display: "block", width: 40, height: 16, background: T.gold, borderRadius: 2 }} />
        <span style={{ display: "block", width: 28, height: 10, background: T.gold, borderRadius: 2 }} />
        <span style={{ display: "block", width: 16, height: 6, background: T.gold, borderRadius: 2 }} />
      </div>
      <div style={{ textAlign: "center" }}>
        <div style={{ fontSize: 24, fontWeight: 300, color: "white", letterSpacing: 4, textTransform: "uppercase", marginBottom: 8 }}>Blue Heron</div>
        <div style={{ fontSize: 13, color: "rgba(255,255,255,0.5)", letterSpacing: 2, textTransform: "uppercase" }}>Project Dashboard</div>
      </div>
      <button onClick={onLogin} style={{ background: T.gold, color: T.navy, border: "none", padding: "12px 32px", borderRadius: 4, fontSize: 13, fontWeight: 600, cursor: "pointer", letterSpacing: 1, textTransform: "uppercase", fontFamily: "inherit", marginTop: 16 }}>
        Sign in with Microsoft
      </button>
      <div style={{ fontSize: 11, color: "rgba(255,255,255,0.3)" }}>Use your Blue Heron Microsoft account</div>
    </div>
  );
}

function Header({ onRefresh, lastUpdated, onLogout, userName }) {
  return (
    <header style={{ background: T.navy, height: 58, padding: "0 28px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 100, boxShadow: "0 2px 12px rgba(0,0,0,0.3)" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
          <span style={{ display: "block", width: 20, height: 8, background: T.gold, borderRadius: 1 }} />
          <span style={{ display: "block", width: 14, height: 5, background: T.gold, borderRadius: 1 }} />
          <span style={{ display: "block", width: 8, height: 3, background: T.gold, borderRadius: 1 }} />
        </div>
        <span style={{ fontSize: 15, fontWeight: 600, color: "white", letterSpacing: 3, textTransform: "uppercase" }}>Blue Heron</span>
        <span style={{ width: 1, height: 24, background: "rgba(255,255,255,0.2)", margin: "0 12px" }} />
        <span style={{ fontSize: 12, color: "rgba(255,255,255,0.6)", letterSpacing: 1, textTransform: "uppercase" }}>Project Dashboard</span>
        {lastUpdated && <span style={{ fontSize: 10, color: "rgba(255,255,255,0.4)", marginLeft: 8 }}>Updated: {lastUpdated}</span>}
      </div>
      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        {userName && <span style={{ fontSize: 11, color: "rgba(255,255,255,0.5)" }}>{userName}</span>}
        <button onClick={onRefresh} style={{ background: "transparent", border: `1px solid rgba(201,168,76,0.4)`, color: T.gold, padding: "5px 14px", borderRadius: 3, fontSize: 11, fontWeight: 500, letterSpacing: 0.8, textTransform: "uppercase", cursor: "pointer", fontFamily: "inherit" }}>↻ Refresh</button>
        <button onClick={onLogout} style={{ background: "transparent", border: `1px solid rgba(255,255,255,0.2)`, color: "rgba(255,255,255,0.5)", padding: "5px 14px", borderRadius: 3, fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>Sign out</button>
      </div>
    </header>
  );
}

export default function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [projects, setProjects] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [selected, setSelected] = useState(null);
  const [phaseFilter, setPhaseFilter] = useState("all");
  const [healthFilter, setHealthFilter] = useState("all");
  const [lastUpdated, setLastUpdated] = useState(null);

  // Handle redirect response on page load
  useEffect(() => {
    instance.handleRedirectPromise().catch(console.error);
  }, [instance]);

  const handleLogin = async () => {
    try {
      await instance.loginRedirect({
        ...loginRequest,
        redirectUri: "https://thankful-pond-052fbeb0f.7.azurestaticapps.net"
      });
    } catch (e) {
      console.error(e);
    }
  };

  const handleLogout = () => {
    instance.logoutRedirect({
      postLogoutRedirectUri: "https://thankful-pond-052fbeb0f.7.azurestaticapps.net"
    });
  };

  const load = useCallback(async () => {
    if (!isAuthenticated) return;
    setLoading(true); setError(null);
    try {
      const data = await fetchProjects(instance);
      setProjects(data);
      setLastUpdated(new Date().toLocaleTimeString());
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        await instance.loginRedirect({
          ...loginRequest,
          redirectUri: "https://thankful-pond-052fbeb0f.7.azurestaticapps.net"
        });
      } else {
        setError(e.message);
      }
    } finally {
      setLoading(false);
    }
  }, [isAuthenticated, instance]);

  useEffect(() => {
    if (isAuthenticated) {
      load();
      const interval = setInterval(load, 5 * 60 * 1000);
      return () => clearInterval(interval);
    }
  }, [isAuthenticated, load]);

  if (!isAuthenticated) return <LoginPage onLogin={handleLogin} />;

  const userName = accounts[0]?.name || accounts[0]?.username;

  const filtered = projects.filter(p => {
    if (phaseFilter !== "all" && p.current_phase !== phaseFilter) return false;
    if (healthFilter !== "all" && p.project_health !== healthFilter) return false;
    return true;
  });

  const total   = filtered.length;
  const inCon   = filtered.filter(p => p.in_construction === "Yes").length;
  const inPre   = filtered.filter(p => p.in_pre_con === "Yes").length;
  const overdue = filtered.reduce((a, p) => a + num(p.tasks_overdue), 0);
  const issues  = filtered.reduce((a, p) => a + num(p.open_issues), 0);

  const phaseData = ["CONSTRUCTION", "PRECON"].map(ph => ({
    name: ph === "CONSTRUCTION" ? "Construction" : "Pre-Con",
    value: filtered.filter(p => p.current_phase === ph).length
  })).filter(d => d.value > 0);

  const overdueData = filtered
    .filter(p => num(p.tasks_overdue) > 0)
    .sort((a, b) => num(b.tasks_overdue) - num(a.tasks_overdue))
    .slice(0, 8)
    .map(p => ({ name: (p.project_name || "").split(" ")[0], value: num(p.tasks_overdue) }));

  if (loading) return (
    <div style={{ minHeight: "100vh", background: T.bg, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 16 }}>
      <div style={{ width: 40, height: 40, border: `3px solid ${T.border}`, borderTop: `3px solid ${T.navy}`, borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
      <div style={{ color: T.muted, fontSize: 13 }}>Loading live project data…</div>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } } * { box-sizing: border-box; margin: 0; padding: 0; }`}</style>
    </div>
  );

  if (error) return (
    <div style={{ minHeight: "100vh", background: T.bg, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 12, padding: 24 }}>
      <Header onRefresh={load} lastUpdated={lastUpdated} onLogout={handleLogout} userName={userName} />
      <div style={{ fontSize: 13, color: T.red, background: "#FDECEA", border: `1px solid ${T.red}`, borderRadius: 6, padding: "16px 24px", maxWidth: 600, textAlign: "center" }}>
        <strong>Error loading data</strong><br /><br />{error}
      </div>
      <button onClick={load} style={{ background: T.navy, color: "white", border: "none", padding: "8px 20px", borderRadius: 4, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}>Retry</button>
    </div>
  );

  if (selected) {
    const p = selected;
    const ahead = !(p.schedule_variance || "").startsWith("-");
    return (
      <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Segoe UI', system-ui, sans-serif" }}>
        <style>{`* { box-sizing: border-box; margin: 0; padding: 0; }`}</style>
        <Header onRefresh={load} lastUpdated={lastUpdated} onLogout={handleLogout} userName={userName} />
        <div style={{ padding: "20px 28px" }}>
          <button onClick={() => setSelected(null)} style={{ background: "transparent", border: `1px solid ${T.border}`, padding: "6px 14px", borderRadius: 4, fontSize: 12, color: T.muted, cursor: "pointer", marginBottom: 16, fontFamily: "inherit" }}>← All Projects</button>
          <div style={{ background: "#EBF0F8", borderLeft: `3px solid ${T.navy}`, borderRadius: "0 4px 4px 0", padding: "10px 14px", marginBottom: 16, fontSize: 12 }}>
            <strong>{p.project_name}</strong> — {p.current_phase || "Pre-planning"}, {fmtPct(p.project_completion)} complete.{" "}
            {num(p.tasks_overdue) > 0 ? <span style={{ color: T.red }}>{num(p.tasks_overdue)} overdue tasks. </span> : "No overdue tasks. "}
            {num(p.open_issues) > 0 ? `${num(p.open_issues)} open issue(s).` : "No open issues."}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 14, marginBottom: 20 }}>
            <KpiCard label="Completion" value={fmtPct(p.project_completion)} sub={`${num(p.tasks_completed)} tasks done`} />
            <KpiCard label="Initial Budget" value={fmt$(p.initial_budget)} sub={`Lot: ${fmt$(p.lot_cost)}`} accent={T.gold} />
            <KpiCard label="Overdue Tasks" value={num(p.tasks_overdue)} sub={`${num(p.tasks_in_progress)} in progress`} accent={num(p.tasks_overdue) > 0 ? T.red : T.green} />
            <KpiCard label="Open Issues" value={num(p.open_issues)} sub={`${num(p.open_risks)} risk(s)`} accent={num(p.open_issues) > 0 ? T.red : T.green} />
            <KpiCard label="Health" value={p.project_health || "—"} sub={`${p.division || ""} division`} accent={p.project_health === "Green" ? T.green : p.project_health === "Red" ? T.red : T.amber} />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
            <Panel title="Project Info">
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                {[["Community", p.community], ["Lot No.", p.lot_no], ["Address", p.lot_address_full || "—"], ["Division", p.division], ["Project No.", p.project_no], ["In Construction", p.in_construction], ["In Pre-Con", p.in_pre_con]].map(([k, v]) => (
                  <tr key={k}><td style={{ padding: "6px 0", borderBottom: `1px solid ${T.border}`, color: T.muted, fontSize: 11, width: "45%" }}>{k}</td><td style={{ padding: "6px 0", borderBottom: `1px solid ${T.border}`, fontWeight: 600 }}>{v || "—"}</td></tr>
                ))}
              </table>
            </Panel>
            <Panel title="Budget">
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                {[["Initial Budget", fmt$(p.initial_budget)], ["Lot Cost", fmt$(p.lot_cost)], ["Home Base Price", fmt$(p.home_base_price)], ["Total Expenses", fmt$(p.total_project_expenses)], ["$/SF", num(p.current_sqft) > 0 && num(p.initial_budget) > 0 ? "$" + Math.round(num(p.initial_budget) / num(p.current_sqft)) + "/sf" : "—"]].map(([k, v]) => (
                  <tr key={k}><td style={{ padding: "6px 0", borderBottom: `1px solid ${T.border}`, color: T.muted, fontSize: 11, width: "45%" }}>{k}</td><td style={{ padding: "6px 0", borderBottom: `1px solid ${T.border}`, fontWeight: 600 }}>{v}</td></tr>
                ))}
              </table>
            </Panel>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
            <Panel title="Schedule Variance">
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
                {[["Baseline Start", p.baseline_start_date], ["Baseline Finish", p.baseline_end_date], null, ["Actual Start", p.actual_start_date], ["Forecasted Turnover", p.current_forcasted_turnover]].map((item, i) =>
                  i === 2 ? (
                    <div key="var" style={{ textAlign: "center" }}>
                      <div style={{ fontSize: 10, textTransform: "uppercase", color: T.muted, marginBottom: 5, fontWeight: 500 }}>Variance</div>
                      <span style={{ background: ahead ? T.green : T.red, color: "white", padding: "5px 12px", borderRadius: 4, fontSize: 12, fontWeight: 700, display: "inline-block" }}>{ahead ? "+" : ""}{p.schedule_variance || "0d"}</span>
                    </div>
                  ) : item ? (
                    <div key={item[0]} style={{ textAlign: "center" }}>
                      <div style={{ fontSize: 10, textTransform: "uppercase", color: T.muted, marginBottom: 5, fontWeight: 500 }}>{item[0]}</div>
                      <div style={{ fontSize: 12, fontWeight: 600 }}>{item[1] || "—"}</div>
                    </div>
                  ) : null
                )}
              </div>
            </Panel>
            <Panel title="Open Activities" noPad>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)" }}>
                {[["Overdue", num(p.tasks_overdue), T.red], ["Risks", num(p.open_risks), T.amber], ["Issues", num(p.open_issues), num(p.open_issues) > 0 ? T.red : T.green]].map(([label, val, color]) => (
                  <div key={label} style={{ textAlign: "center", padding: "16px 12px", borderRight: `1px solid ${T.border}` }}>
                    <div style={{ fontSize: 30, fontWeight: 300, color }}>{val}</div>
                    <div style={{ fontSize: 10, textTransform: "uppercase", color: T.muted, marginTop: 3, fontWeight: 500 }}>{label}</div>
                  </div>
                ))}
              </div>
            </Panel>
          </div>
          <Panel title="Property Details">
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12 }}>
              {[["Lot No.", p.lot_no], ["Bedrooms", p.of_bedrooms], ["Bathrooms", p.of_bathrooms], ["Sq Ft", num(p.current_sqft) > 0 ? num(p.current_sqft).toLocaleString() : "—"]].map(([label, val]) => (
                <div key={label}>
                  <div style={{ fontSize: 10, textTransform: "uppercase", color: T.muted, fontWeight: 500 }}>{label}</div>
                  <div style={{ fontSize: 18, fontWeight: 600, marginTop: 4 }}>{val || "—"}</div>
                </div>
              ))}
            </div>
          </Panel>
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'Segoe UI', system-ui, sans-serif" }}>
      <style>{`* { box-sizing: border-box; margin: 0; padding: 0; } tr:hover td { background: #F8F9FB !important; }`}</style>
      <Header onRefresh={load} lastUpdated={lastUpdated} onLogout={handleLogout} userName={userName} />
      <div style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, padding: "10px 28px", display: "flex", alignItems: "center", gap: 20, flexWrap: "wrap" }}>
        {[["Project", <select key="p" onChange={e => e.target.value !== "all" && setSelected(projects[parseInt(e.target.value)])} style={{ fontFamily: "inherit", fontSize: 13, fontWeight: 500, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "5px 10px", cursor: "pointer", minWidth: 200 }}>
          <option value="all">All Projects</option>
          {projects.map((p, i) => <option key={i} value={i}>{p.project_name}{p.lot_no ? ` (Lot ${p.lot_no})` : ""}</option>)}
        </select>],
        ["Phase", <select key="ph" value={phaseFilter} onChange={e => setPhaseFilter(e.target.value)} style={{ fontFamily: "inherit", fontSize: 13, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "5px 10px", cursor: "pointer" }}>
          <option value="all">All Phases</option>
          <option value="CONSTRUCTION">Construction</option>
          <option value="PRECON">Pre-Con</option>
        </select>],
        ["Health", <select key="h" value={healthFilter} onChange={e => setHealthFilter(e.target.value)} style={{ fontFamily: "inherit", fontSize: 13, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 4, padding: "5px 10px", cursor: "pointer" }}>
          <option value="all">All</option>
          <option value="Green">Green</option>
          <option value="Red">Red</option>
        </select>]].map(([label, control]) => (
          <div key={label} style={{ display: "flex", flexDirection: "column", gap: 3 }}>
            <span style={{ fontSize: 10, textTransform: "uppercase", letterSpacing: 0.6, color: T.muted, fontWeight: 500 }}>{label}</span>
            {control}
          </div>
        ))}
      </div>
      <div style={{ padding: "20px 28px" }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 14, marginBottom: 20 }}>
          <KpiCard label="Total Projects" value={total} sub="Ascaya community" />
          <KpiCard label="In Construction" value={inCon} sub="Active builds" accent={T.gold} />
          <KpiCard label="In Pre-Con" value={inPre} sub="Planning phase" accent={T.navyMid} />
          <KpiCard label="Overdue Tasks" value={overdue} sub="Across all projects" accent={T.red} />
          <KpiCard label="Open Issues" value={issues} sub={issues > 0 ? "Needs attention" : "All clear"} accent={issues > 0 ? T.red : T.green} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 20 }}>
          <Panel title="Projects by Phase">
            <ResponsiveContainer width="100%" height={180}>
              <PieChart>
                <Pie data={phaseData} cx="50%" cy="50%" outerRadius={70} dataKey="value" label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`} labelLine={false} fontSize={11}>
                  {phaseData.map((_, i) => <Cell key={i} fill={[T.gold, T.navy][i % 2]} />)}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          </Panel>
          <Panel title="Overdue Tasks by Project">
            <ResponsiveContainer width="100%" height={180}>
              <BarChart data={overdueData} layout="vertical" margin={{ left: 10, right: 20 }}>
                <XAxis type="number" tick={{ fontSize: 10 }} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 10 }} width={70} />
                <Tooltip />
                <Bar dataKey="value" fill={T.red} radius={[0, 3, 3, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </Panel>
        </div>
        <Panel title="All Projects" count={filtered.length}>
          <div style={{ overflowX: "auto", margin: "-14px -16px" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr>{["Project", "Lot", "Phase", "Health", "Completion", "Variance", "Overdue", "Issues", "Risks", "Budget"].map(h => (
                  <th key={h} style={{ fontSize: 10, textTransform: "uppercase", color: T.muted, fontWeight: 600, padding: "8px 12px", borderBottom: `1px solid ${T.border}`, textAlign: "left", whiteSpace: "nowrap", background: T.surface }}>{h}</th>
                ))}</tr>
              </thead>
              <tbody>
                {filtered.map((p, i) => (
                  <tr key={i} onClick={() => setSelected(p)} style={{ cursor: "pointer" }}>
                    <td style={{ padding: "10px 12px", borderBottom: `1px solid ${T.border}` }}>
                      <div style={{ fontWeight: 600 }}>{p.project_name}</div>
                      {p.lot_address_full && <div style={{ fontSize: 10, color: T.muted, marginTop: 1 }}>{p.lot_address_full}</div>}
                    </td>
                    <td style={{ padding: "10px 12px", borderBottom: `1px solid ${T.border}` }}>{p.lot_no || "—"}</td>
                    <td style={{ padding: "10px 12px", borderBottom: `1px solid ${T.border}` }}><PhaseBadge phase={p.current_phase} /></td>
                    <td style={{ padding: "10px 12px", borderBottom: `1px solid ${T.border}` }}><HealthDot health={p.project_health} /></td>
                    <td style={{ padding: "10px 12px", borderBottom: `1px solid ${T.border}`, minWidth: 130 }}><ProgressBar pct={num(p.project_completion)} /></td>
                    <td style={{ padding: "10px 12px", borderBottom: `1px solid ${T.border}`, fontWeight: 600, color: (p.schedule_variance || "").startsWith("-") ? T.red : T.green }}>{p.schedule_variance || "—"}</td>
                    <td style={{ padding: "10px 12px", borderBottom: `1px solid ${T.border}`, fontWeight: 700, color: num(p.tasks_overdue) > 0 ? T.red : T.text }}>{num(p.tasks_overdue)}</td>
                    <td style={{ padding: "10px 12px", borderBottom: `1px solid ${T.border}`, fontWeight: 700, color: num(p.open_issues) > 0 ? T.red : T.text }}>{num(p.open_issues)}</td>
                    <td style={{ padding: "10px 12px", borderBottom: `1px solid ${T.border}`, color: T.amber, fontWeight: 600 }}>{num(p.open_risks)}</td>
                    <td style={{ padding: "10px 12px", borderBottom: `1px solid ${T.border}` }}>{fmt$(p.initial_budget)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Panel>
      </div>
    </div>
  );
}
