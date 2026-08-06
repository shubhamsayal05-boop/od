import { useCallback, useEffect, useMemo, useState } from "react";
import {
  ScatterChart, Scatter, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer,
} from "recharts";
import { api, fmt, getChannel, CELL_COLORS } from "@/lib/api";
import { useWorkbook } from "@/components/Workbook";

const TEAL = "#215967";
const HBLUE = "#538DD5";
const CRIT_FILL = { 1: "#FFC7CE", 2: "#FFEB9C", 3: "#C6EFCE" };

export default function SdvSheet({ name }) {
  const { openTab, refresh, setSelection } = useWorkbook();
  const [data, setData] = useState(null);
  const [err, setErr] = useState(null);
  const [part, setPart] = useState("driv"); // driv | dyn
  const [colorFilter, setColorFilter] = useState(null); // null | RED | RY
  const [gearFilter, setGearFilter] = useState(null);
  const [showC3, setShowC3] = useState(false);
  const [busy, setBusy] = useState(false);

  const load = useCallback(async () => {
    try {
      setErr(null);
      const res = await api.get(`/sdv/${encodeURIComponent(name)}`);
      setData(res.data);
    } catch (e) { setErr(e.response?.data?.detail || e.message); }
  }, [name]);
  useEffect(() => { load(); }, [load]);

  const update = async () => {
    setBusy(true);
    try { await api.post("/rating/calculate"); await load(); await refresh(); }
    catch (e) { alert(e.response?.data?.detail || e.message); }
    finally { setBusy(false); }
  };

  const criteria = useMemo(() => {
    if (!data) return [];
    const key = part === "driv" ? "driv" : "resp";
    return (data.structure?.criteria || [])
      .map((c) => ({ name: c.name, crit: data.targets?.[c.name]?.[key] ?? null,
        wl: data.targets?.[c.name]?.wl, t: data.targets?.[c.name]?.t }))
      .filter((c) => c.crit !== null);
  }, [data, part]);

  const visibleCriteria = criteria.filter((c) => showC3 || Number(c.crit) !== 3);
  const dataCols = (data?.structure?.data || []).slice(0, 8);

  const events = useMemo(() => {
    if (!data) return [];
    let evs = data.events.map((ev) => ({ ...ev, comp: ev[part] }));
    if (colorFilter === "RED") evs = evs.filter((e) => e.comp?.color === "RED");
    if (colorFilter === "RY") evs = evs.filter((e) => ["RED", "YELLOW"].includes(e.comp?.color));
    if (gearFilter != null) {
      evs = evs.filter((e) => String(getChannel(e.channels, "gear new") ?? getChannel(e.channels, "Gear New") ?? "") === String(gearFilter));
    }
    return evs;
  }, [data, part, colorFilter, gearFilter]);

  if (err) return <div style={{ padding: 30, color: "#a00" }}>{err}</div>;
  if (!data) return <div style={{ padding: 30, color: "#777" }}>Loading {name}…</div>;

  const res = data.result || {};
  const pr = res[part] || {};
  const counts = pr.counts || {};
  const idxLabel = part === "driv" ? "Drivability Index" : "Responsiveness Index";
  const summaryLabel = part === "driv" ? "DRIVABILITY SUMMARY" : "RESPONSIVENESS SUMMARY";
  const idxCell = part === "driv" ? "J5" : "BQ5";

  const chart = (data.charts || []).find((c) => c.active) || data.charts?.[0];
  const xName = chart?.x || "Vehicle Speed";
  const yName = chart?.y || "AccelerationChassis";
  const points = { RED: [], YELLOW: [], GREEN: [] };
  data.events.forEach((ev) => {
    const x = Number(getChannel(ev.channels, xName));
    const y = Number(getChannel(ev.channels, yName));
    const c = ev[part]?.color;
    if (!Number.isNaN(x) && !Number.isNaN(y) && points[c]) points[c].push({ x, y, name: String(getChannel(ev.channels, "Sub Event Name") ?? "") });
  });

  return (
    <div style={{ padding: "8px 14px 50px", minWidth: 1100 }}>
      {/* toolbar */}
      <div className="xl-ribbon" style={{ border: "1px solid #D4D4D4", marginBottom: 8 }}>
        <button className="xl-btn primary" onClick={update} disabled={busy} data-testid="sdv-update-btn">
          {busy ? "UPDATING…" : "UPDATE"}
        </button>
        <button className="xl-btn" onClick={load} data-testid="sdv-targets-btn" title="Re-pull targets & recolor">TARGETS</button>
        <span className="sep" />
        <button className="xl-btn" style={colorFilter === "RED" ? { background: "#FBC7C7" } : null}
          onClick={() => setColorFilter(colorFilter === "RED" ? null : "RED")} data-testid="sdv-red-only-btn">RED ONLY</button>
        <button className="xl-btn" style={colorFilter === "RY" ? { background: "#FBE9A0" } : null}
          onClick={() => setColorFilter(colorFilter === "RY" ? null : "RY")} data-testid="sdv-yellow-red-btn">YELLOW + RED</button>
        <button className="xl-btn" onClick={() => { setColorFilter(null); setGearFilter(null); }} data-testid="sdv-filters-off-btn">FILTERS OFF</button>
        <span className="sep" />
        <span style={{ fontSize: 11, color: "#666" }}>Gear New =</span>
        {[1, 2, 3, 4, 5, 6, 7, 8].map((g) => (
          <button key={g} className="xl-btn" style={{ padding: "2px 7px", ...(gearFilter === g ? { background: "#CCE4F7" } : {}) }}
            onClick={() => setGearFilter(gearFilter === g ? null : g)} data-testid={`sdv-gear-${g}`}>{g}</button>
        ))}
        <span className="sep" />
        <button className="xl-btn" onClick={() => setShowC3(!showC3)} data-testid="sdv-c3-btn">
          {showC3 ? "HIDE C3" : "Show C3"}
        </button>
        <span className="sep" />
        <button className="xl-btn" onClick={() => setPart(part === "driv" ? "dyn" : "driv")} data-testid="sdv-dyn-toggle">
          {part === "driv" ? "→ dyn" : "→ driv"}
        </button>
        <button className="xl-btn" onClick={() => openTab("HOME")} data-testid="sdv-home-btn">🏠 HOME</button>
        <button className="xl-btn" onClick={() => openTab("RATING")} data-testid="sdv-rating-btn">Q RATING</button>
        <span className="hint">{events.length}/{data.events.length} events shown</span>
      </div>

      {/* banner + summary */}
      <div style={{ display: "flex", gap: 14, alignItems: "flex-start", flexWrap: "wrap" }}>
        <div style={{ flex: "0 0 auto" }}>
          <div style={{ background: TEAL, color: "#fff", fontWeight: 700, fontSize: 15, padding: "6px 14px" }} data-testid="sdv-banner">
            {name} — {summaryLabel}
          </div>
          <table className="xl-grid" style={{ marginTop: 4 }} data-testid="sdv-summary-table">
            <thead>
              <tr>
                {["Color", "Priority (Px)", "Events", "% of total"].map((h) => (
                  <th key={h} style={{ background: HBLUE, color: "#fff" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr style={{ fontWeight: 700 }}>
                <td>TOTAL</td><td>—</td>
                <td className="num">{counts.total ?? 0}</td><td className="num">100%</td>
              </tr>
              {["GREEN", "YELLOW", "RED"].map((color) =>
                [1, 2, 3].map((p, i) => (
                  <tr key={color + p}>
                    {i === 0 && (
                      <td rowSpan={3} style={{ background: { GREEN: "#007F00", YELLOW: "#FFFF00", RED: "#FF0000" }[color], color: color === "YELLOW" ? "#333" : "#fff", fontWeight: 700 }}>
                        {color[0] + color.slice(1).toLowerCase()}
                      </td>
                    )}
                    <td>P{p}</td>
                    <td className="num" data-testid={`count-${color}-P${p}`}>{counts[`${color}_P${p}`] ?? 0}</td>
                    <td className="num">
                      {counts.total ? (((counts[`${color}_P${p}`] || 0) / counts.total) * 100).toFixed(1) + "%" : "0%"}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {/* index cells */}
        <div style={{ display: "flex", gap: 8 }}>
          <IndexCell label={idxLabel} addr={idxCell} value={pr.index} big setSelection={setSelection} testid="sdv-index" />
          <IndexCell label="Target Index" addr="K5" value={pr.target_index} setSelection={setSelection} testid="sdv-target-index" />
        </div>

        {/* status dots */}
        <table className="xl-grid">
          <thead>
            <tr><th style={{ background: HBLUE, color: "#fff" }} colSpan={3}>Status (current)</th>
              <th style={{ background: HBLUE, color: "#fff" }} colSpan={3}>SOPM prediction</th></tr>
            <tr>{["P1", "P2", "P3", "P1", "P2", "P3"].map((p, i) => <th key={i}>{p}</th>)}</tr>
          </thead>
          <tbody>
            <tr>
              {[["status", 1], ["status", 2], ["status", 3], ["status_pred", 1], ["status_pred", 2], ["status_pred", 3]].map(([k, p], i) => {
                const st = pr[k]?.[String(p)] || "NONE";
                return (
                  <td key={i} style={{ textAlign: "center" }} data-testid={`sdv-${k}-P${p}`}>
                    <span style={{ color: { RED: "#FF0000", ORANGE: "#FFC000", GREEN: "#00B050", NONE: "#BFBFBF" }[st], fontSize: 17 }}>●</span>
                  </td>
                );
              })}
            </tr>
          </tbody>
        </table>
      </div>

      {/* event grid */}
      <div style={{ marginTop: 14, overflowX: "auto" }}>
        <table className="xl-grid" data-testid="sdv-event-table">
          <thead>
            {/* criticity row (row 5) */}
            <tr>
              <th colSpan={4 + dataCols.length} style={{ textAlign: "right", background: "#fff", border: "none" }}>Criticity →</th>
              {visibleCriteria.map((c) => (
                <th key={c.name} style={{ background: CRIT_FILL[Number(c.crit)] || "#eee" }} data-testid={`crit-${c.name}`}>
                  {c.crit}
                </th>
              ))}
              <th style={{ background: "#fff", border: "none" }} colSpan={2} />
            </tr>
            {/* header row 6 */}
            <tr>
              <th className="rowhead">#</th>
              <th>Criticality</th>
              <th>Priority</th>
              <th>Color</th>
              {dataCols.map((d) => <th key={d.name}>{d.name}</th>)}
              {visibleCriteria.map((c) => <th key={c.name} style={{ background: HBLUE, color: "#fff", minWidth: 56 }}>{c.name}</th>)}
              <th>Indice occurrencé</th>
              <th>Id BdD</th>
            </tr>
            {/* waterline / target rows */}
            <tr style={{ background: "#FBE2D5" }}>
              <th colSpan={4 + dataCols.length} style={{ textAlign: "right" }}>Waterline</th>
              {visibleCriteria.map((c) => <td key={c.name} className="num">{fmt(c.wl)}</td>)}
              <td colSpan={2} />
            </tr>
            <tr style={{ background: "#A7D5AB" }}>
              <th colSpan={4 + dataCols.length} style={{ textAlign: "right" }}>Target</th>
              {visibleCriteria.map((c) => <td key={c.name} className="num">{fmt(c.t)}</td>)}
              <td colSpan={2} />
            </tr>
          </thead>
          <tbody>
            {events.map((ev, i) => (
              <EventRow key={ev.id} i={i + 1} ev={ev} dataCols={dataCols}
                criteria={visibleCriteria} setSelection={setSelection} />
            ))}
          </tbody>
        </table>
      </div>

      {/* charts */}
      <div style={{ marginTop: 20, display: "flex", gap: 20, flexWrap: "wrap" }}>
        <div style={{ width: 640, border: "1px solid #ccc", padding: 6 }} data-testid="sdv-chart">
          <div style={{ fontSize: 12, fontWeight: 700, color: TEAL, marginBottom: 2 }}>
            {chart?.name || "Graphique 1"} — {yName} vs {xName}
          </div>
          <ResponsiveContainer width="100%" height={320} minWidth={300}>
            <ScatterChart margin={{ top: 8, right: 16, bottom: 18, left: 0 }}>
              <CartesianGrid strokeDasharray="2 2" />
              <XAxis type="number" dataKey="x" name={xName}
                domain={[numOr(chart?.x_min, "auto"), numOr(chart?.x_max, "auto")]}
                label={{ value: xName, position: "insideBottom", offset: -8, fontSize: 11 }} tick={{ fontSize: 10 }} />
              <YAxis type="number" dataKey="y" name={yName}
                domain={[numOr(chart?.y_min, "auto"), numOr(chart?.y_max, "auto")]}
                tick={{ fontSize: 10 }} />
              <Tooltip cursor={{ strokeDasharray: "3 3" }}
                formatter={(v) => fmt(v, 2)} labelFormatter={() => ""} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              {["GREEN", "YELLOW", "RED"].map((c) => (
                <Scatter key={c} name={c} data={points[c]} fill={CELL_COLORS[c]}
                  stroke="#333" strokeWidth={0.6} />
              ))}
            </ScatterChart>
          </ResponsiveContainer>
        </div>
        <div style={{ fontSize: 11.5, color: "#666", maxWidth: 320 }}>
          <b>Chart parameters (PARAMETRES GRAPH)</b>
          <table className="xl-grid" style={{ marginTop: 4 }}>
            <tbody>
              <tr><td className="rowhead">Abscisse</td><td>{xName}</td></tr>
              <tr><td className="rowhead">Ordonnée</td><td>{yName}</td></tr>
              <tr><td className="rowhead">X range</td><td>{String(chart?.x_min ?? "Auto")} → {String(chart?.x_max ?? "Auto")}</td></tr>
              <tr><td className="rowhead">Y range</td><td>{String(chart?.y_min ?? "Auto")} → {String(chart?.y_max ?? "Auto")}</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function numOr(v, dflt) {
  const f = Number(v);
  return v === null || v === undefined || v === "" || Number.isNaN(f) || String(v).toLowerCase() === "automatique" ? dflt : f;
}

function IndexCell({ label, addr, value, big, setSelection, testid }) {
  return (
    <div onClick={() => setSelection({ addr, value: value ?? "" })} data-testid={testid}
      style={{ border: "2px solid " + TEAL, minWidth: 120, textAlign: "center", cursor: "pointer" }}>
      <div style={{ background: TEAL, color: "#fff", fontSize: 11, padding: "3px 8px" }}>{label}</div>
      <div style={{ fontSize: big ? 30 : 24, fontWeight: 700, padding: "8px 10px", color: "#1E2336" }}>
        {fmt(value)}
      </div>
    </div>
  );
}

function EventRow({ i, ev, dataCols, criteria, setSelection }) {
  const comp = ev.comp;
  const colorBg = { RED: "#FF0000", YELLOW: "#FFFF00", GREEN: "#00B050" }[comp?.color];
  return (
    <tr data-testid={`event-row-${ev.id.slice(0, 8)}`}>
      <td className="rowhead">{i}</td>
      <td style={{ fontSize: 11 }}>{comp?.criticity ?? "-"}</td>
      <td style={{ textAlign: "center" }}>{comp ? `P${comp.priority}` : "-"}</td>
      <td style={{ background: colorBg, color: comp?.color === "YELLOW" ? "#333" : "#fff", textAlign: "center", fontSize: 10.5 }}>
        {comp?.color ?? ""}
      </td>
      {dataCols.map((d) => {
        const v = getChannel(ev.channels, d.import || d.name) ?? getChannel(ev.channels, d.name);
        return <td key={d.name} className="num" onClick={() => setSelection({ addr: d.name, value: v ?? "" })}>{typeof v === "number" ? fmt(v, 1) : String(v ?? "")}</td>;
      })}
      {criteria.map((c) => {
        const v = getChannel(ev.channels, c.name);
        const cc = comp?.crit_colors?.[c.name] || comp?.crit_colors?.[c.name.replace(/\./g, "\u00b7")];
        const bg = { RED: "#FFC7CE", YELLOW: "#FFEB9C", GREEN: "#C6EFCE" }[cc];
        return (
          <td key={c.name} className="num" style={{ background: bg }}
            onClick={() => setSelection({ addr: c.name, value: v ?? "" })}>
            {typeof v === "number" ? fmt(v, 1) : String(v ?? "")}
          </td>
        );
      })}
      <td className="num" style={{ fontWeight: 700, color: (comp?.indice_occ ?? 0) < 0 ? "#c00" : "#070" }}>
        {fmt(comp?.indice_occ, 3)}
      </td>
      <td style={{ fontFamily: "monospace", fontSize: 10.5 }}>{ev.id.slice(0, 8)}</td>
    </tr>
  );
}
