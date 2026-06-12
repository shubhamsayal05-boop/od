import { useEffect, useState } from "react";
import { api, fmt, DOT_COLORS, VERDICT_COLORS } from "@/lib/api";
import { useWorkbook } from "@/components/Workbook";

const HEADBLUE = "#17375E";

export default function RatingSheet() {
  const { openTab, refresh, setSelection } = useWorkbook();
  const [data, setData] = useState(null);
  const [maskEmpty, setMaskEmpty] = useState(true);
  const [busy, setBusy] = useState(null);

  const load = async () => {
    const res = await api.get("/rating");
    setData(res.data);
  };
  useEffect(() => { load(); }, []);

  const calculate = async () => {
    setBusy("calc");
    try { await api.post("/rating/calculate"); await load(); await refresh(); }
    catch (e) { alert(e.response?.data?.detail || e.message); }
    finally { setBusy(null); }
  };

  const setAsTarget = async () => {
    if (!window.confirm("Promote the current vehicle's results to TARGET VEHICLE (benchmark)?")) return;
    setBusy("target");
    try { await api.post("/rating/set-as-target"); await load(); alert("Current results are now the TARGET VEHICLE benchmark."); }
    catch (e) { alert(e.response?.data?.detail || e.message); }
    finally { setBusy(null); }
  };

  if (!data) return <div style={{ padding: 30, color: "#777" }}>Loading RATING…</div>;
  const { project, global: glob, rows, catalog } = data;
  const resultBySdv = Object.fromEntries((rows || []).map((r) => [r.name, r]));
  const tvName = project?.target_vehicle || "P8 MHEV MDL2";

  // build catalog-ordered display rows with group bands
  const groups = [];
  (catalog || []).forEach((c) => {
    let g = groups.find((x) => x.name === c.group);
    if (!g) { g = { name: c.group, sdvs: [] }; groups.push(g); }
    g.sdvs.push(c.name);
  });

  const worst = (rows || [])
    .filter((r) => r.driv?.index != null)
    .sort((a, b) => a.driv.index - b.driv.index)
    .slice(0, 3);

  return (
    <div style={{ padding: "10px 16px 40px", minWidth: 1280 }}>
      {/* title */}
      <div style={{ display: "flex", alignItems: "flex-start", gap: 16 }}>
        <div style={{ fontSize: 15, fontWeight: 700, whiteSpace: "pre-line", flex: 1 }} data-testid="rating-title">
          DRIVABILITY VERIFICATION — SCORECARD{"\n"}
          <span style={{ fontSize: 10.5, fontWeight: 400, color: "#666" }}>
            (reflects only vehicle drivability under standard ambient conditions at sea level.
            Cold/hot, grades or high altitude behaviour is not covered by this status)
          </span>
        </div>
        <div style={{ color: "#c00", fontWeight: 700, fontSize: 13 }}>Confidential</div>
      </div>

      {/* meta + actions */}
      <div style={{ display: "flex", gap: 24, margin: "10px 0", alignItems: "flex-start", flexWrap: "wrap" }}>
        <table className="xl-grid">
          <tbody>
            <tr><td className="rowhead">Application :</td><td data-testid="rating-application">{project?.name_code || "-"}</td></tr>
            <tr><td className="rowhead">Stage :</td><td>{project?.odriv_milestone || "-"}</td></tr>
            <tr><td className="rowhead">Odriv Version :</td><td>{data.version}</td></tr>
          </tbody>
        </table>
        <table className="xl-grid">
          <tbody>
            <tr><td className="rowhead" rowSpan={3} style={{ verticalAlign: "top" }}>Top Areas for Improvement:</td>
              {worst[0] ? <td>1. {worst[0].name} ({fmt(worst[0].driv.index)})</td> : <td>1. Op mode (criteria)</td>}</tr>
            <tr>{worst[1] ? <td>2. {worst[1].name} ({fmt(worst[1].driv.index)})</td> : <td>2. Op mode (criteria)</td>}</tr>
            <tr>{worst[2] ? <td>3. {worst[2].name} ({fmt(worst[2].driv.index)})</td> : <td>3. Op mode (criteria)</td>}</tr>
          </tbody>
        </table>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <button className="xl-btn primary" onClick={calculate} disabled={!!busy} data-testid="rating-calculate-btn">
            {busy === "calc" ? "Calculating…" : "CALCULATE RATING"}
          </button>
          <button className="xl-btn" onClick={setAsTarget} disabled={!!busy || !glob} data-testid="rating-set-as-target-btn">
            SET AS TARGET — FOR CALIBRATION TEAM REVIEW
          </button>
          <label className="xl-btn" style={{ display: "flex", alignItems: "center", gap: 5 }}>
            <input type="checkbox" checked={maskEmpty} onChange={(e) => setMaskEmpty(e.target.checked)}
              data-testid="rating-mask-empty" />
            Mask empty SdV
          </label>
        </div>
      </div>

      {/* the two risk blocks */}
      <RiskBlock label={"DRIVABILITY\ncomfort/disturbances (bump, shock, jerk, …)"} g={glob?.driv} tvName={tvName} testid="driv" setSelection={setSelection} />
      <RiskBlock label={"RESPONSIVENESS\nPerformance feel (response delay, vehicle Agility, …)"} g={glob?.dyn} tvName={tvName} testid="dyn" setSelection={setSelection} />

      {/* use-case table */}
      <table className="xl-grid" style={{ marginTop: 18 }} data-testid="rating-table">
        <thead>
          <tr>
            <th rowSpan={3} style={{ minWidth: 230, background: HEADBLUE, color: "#fff" }}>USE CASE</th>
            <th colSpan={9} style={{ background: HEADBLUE, color: "#fff" }}>Drivability</th>
            <th colSpan={9} style={{ background: HEADBLUE, color: "#fff" }}>Responsiveness</th>
          </tr>
          <tr>
            <th colSpan={3} style={{ background: HEADBLUE, color: "#fff" }}>Current Status</th>
            <th colSpan={3} style={{ background: HEADBLUE, color: "#fff" }}>SOPM Prediction</th>
            <th style={{ background: HEADBLUE, color: "#fff" }}>Driveability Index</th>
            <th style={{ background: "#D9D9D9" }}>{tvName}</th>
            <th style={{ background: HEADBLUE, color: "#fff", minWidth: 150 }}>Drivability Lowest Events</th>
            <th colSpan={3} style={{ background: HEADBLUE, color: "#fff" }}>Current Status</th>
            <th colSpan={3} style={{ background: HEADBLUE, color: "#fff" }}>SOPM Prediction</th>
            <th style={{ background: HEADBLUE, color: "#fff" }}>Responsiveness Index</th>
            <th style={{ background: "#D9D9D9" }}>{tvName}</th>
            <th style={{ background: HEADBLUE, color: "#fff", minWidth: 150 }}>Responsiveness Lowest Events</th>
          </tr>
          <tr>
            {["P1", "P2", "P3", "P1", "P2", "P3"].map((p, i) => <th key={"d" + i}>{p}</th>)}
            <th /><th /><th />
            {["P1", "P2", "P3", "P1", "P2", "P3"].map((p, i) => <th key={"r" + i}>{p}</th>)}
            <th /><th /><th />
          </tr>
        </thead>
        <tbody>
          {groups.map((g) => {
            const visible = g.sdvs.filter((s) => !maskEmpty || resultBySdv[s]);
            if (!visible.length) return null;
            return [
              <tr key={g.name}>
                <td colSpan={19} style={{ background: HEADBLUE, color: "#fff", fontWeight: 700 }} data-testid={`rating-group-${g.name}`}>
                  {g.name}
                </td>
              </tr>,
              ...visible.map((sdvName) => {
                const r = resultBySdv[sdvName];
                return (
                  <SdvRow key={sdvName} name={sdvName} r={r} openTab={openTab} />
                );
              }),
            ];
          })}
        </tbody>
      </table>
    </div>
  );
}

function RiskBlock({ label, g, tvName, testid, setSelection }) {
  return (
    <div style={{ display: "flex", gap: 0, marginTop: 10, alignItems: "stretch" }}>
      <div style={{ background: "#fff", border: "1px solid #999", padding: "8px 12px", width: 290, whiteSpace: "pre-line", fontWeight: 700, fontSize: 13 }}>
        {label}
      </div>
      <table className="xl-grid" style={{ borderLeft: "none" }}>
        <thead>
          <tr>
            <th style={{ background: HEADBLUE, color: "#fff", minWidth: 130 }}>Current status</th>
            <th style={{ background: HEADBLUE, color: "#fff", minWidth: 150 }}>Forecast status @ SOPM</th>
            <th style={{ minWidth: 90 }}>index</th>
            <th style={{ minWidth: 160 }}>Weighted % of events below target</th>
            <th style={{ background: "#D9D9D9", minWidth: 110 }}>Tested vehicle</th>
            <th style={{ background: "#D9D9D9", minWidth: 110 }}>{tvName}</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td data-testid={`verdict-${testid}-current`} onClick={() => setSelection({ addr: "E11", value: g?.verdict || "" })}
              style={{ background: VERDICT_COLORS[g?.verdict] || "#F2F2F2", color: "#fff", fontWeight: 700, textAlign: "center" }}>
              {g?.verdict || "—"}
            </td>
            <td data-testid={`verdict-${testid}-forecast`}
              style={{ background: VERDICT_COLORS[g?.verdict_pred] || "#F2F2F2", color: "#fff", fontWeight: 700, textAlign: "center" }}>
              {g?.verdict_pred || "—"}
            </td>
            <td className="num" data-testid={`index-${testid}`}>{fmt(g?.index)}</td>
            <td className="num" data-testid={`ratelow-${testid}`}>{g ? (g.rate_low * 100).toFixed(2) + " %" : "-"}</td>
            <td className="num">{fmt(g?.index)}</td>
            <td className="num">{fmt(g?.target_index)}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function Dot({ status }) {
  return (
    <span style={{ color: DOT_COLORS[status || "NONE"], fontSize: 15, lineHeight: 1 }}>●</span>
  );
}

function SdvRow({ name, r, openTab }) {
  const d = r?.driv;
  const dy = r?.dyn;
  const dots = (part, key) =>
    [1, 2, 3].map((p) => (
      <td key={key + p} style={{ textAlign: "center" }}>
        <Dot status={part ? part[key]?.[String(p)] : "NONE"} />
      </td>
    ));
  return (
    <tr data-testid={`rating-row-${name}`} style={{ background: r ? "#fff" : "#FAFAFA" }}>
      <td>
        <span
          onClick={() => r && openTab(name)}
          style={{ color: r ? "#0563C1" : "#999", textDecoration: r ? "underline" : "none", cursor: r ? "pointer" : "default", paddingLeft: 14 }}
          data-testid={`rating-link-${name}`}
        >
          {name}
        </span>
      </td>
      {dots(d, "status")}
      {dots(d, "status_pred")}
      <td className="num" data-testid={`driv-index-${name}`}>{fmt(d?.index)}</td>
      <td className="num" style={{ background: "#EFEFEF" }}>{fmt(d?.target_index)}</td>
      <td style={{ fontSize: 11 }}>{d?.lowest_event || ""}</td>
      {dots(dy, "status")}
      {dots(dy, "status_pred")}
      <td className="num">{fmt(dy?.index)}</td>
      <td className="num" style={{ background: "#EFEFEF" }}>{fmt(dy?.target_index)}</td>
      <td style={{ fontSize: 11 }}>{dy?.lowest_event || ""}</td>
    </tr>
  );
}
