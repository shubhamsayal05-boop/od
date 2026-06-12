import { useRef, useState } from "react";
import { api } from "@/lib/api";
import { useWorkbook } from "@/components/Workbook";
import { NewProjectModal, DocVersionsModal, OpenDatabaseModal } from "@/components/modals";

const NAVY = "#1E2336";
const SLATE = "#556F81";

const FIELDS = [
  ["MODE", "mode", ["AUTO", "ECO", "SPORT", "MANUAL", "Hybrid", "ZEV", "eSAVE"]],
  ["FUEL", "fuel", ["GASOLINE", "DIESEL"]],
  ["GEARS", "gears", ["5", "6", "7", "8", "9", "10"]],
  ["SOFTWARE MILESTONE", "software_milestone", null],
  ["PRIORITY", "priority", ["PREMIUM", "MAINSTREAM"]],
  ["VERSION", "version", ["4.2", "4.6"]],
  ["ODRIV MILESTONE", "odriv_milestone", null],
  ["AREA", "area", null],
  ["TARGET VEHICLE", "target_vehicle", "text"],
  ["NUMBER OF GEARS", "number_of_gears", ["5", "6", "7", "8", "9", "10"]],
];

export default function HomeSheet() {
  const { state, refresh, openTab, setSelection } = useWorkbook();
  const project = state?.project;
  const [modal, setModal] = useState(null);
  const [busyMsg, setBusyMsg] = useState(null);
  const [lists, setLists] = useState(null);
  const fileRef = useRef(null);
  const lastLog = state?.log_tail?.[0]?.message || "";

  const ensureLists = async () => {
    if (lists) return lists;
    const res = await api.get("/config/configurations");
    setLists(res.data);
    return res.data;
  };

  const run = async (label, fn) => {
    setBusyMsg(label);
    try { await fn(); await refresh(); }
    catch (e) { alert(e.response?.data?.detail || e.message); }
    finally { setBusyMsg(null); }
  };

  const onField = (key) => async (e) => {
    await run("Updating project…", () => api.put("/project", { [key]: e.target.value }));
  };

  const addFile = () => {
    if (!project) return alert("Project information missing. Use NEW PROJECT first.");
    fileRef.current?.click();
  };

  const onFilePicked = async (e) => {
    const f = e.target.files?.[0];
    e.target.value = "";
    if (!f) return;
    const fd = new FormData();
    fd.append("file", f);
    await run("LoadData : importing acquisition file…", async () => {
      const res = await api.post("/import/file", fd);
      const d = res.data;
      alert(`File added to your project.\nImported: ${d.imported}\nClassified: ${d.classified}\nUnclassified: ${d.unclassified}\nSDV sheets: ${Object.keys(d.per_sdv).length}`);
    });
  };

  const demo = () =>
    run("Generating demo acquisition…", async () => {
      const res = await api.post("/import/demo");
      const d = res.data;
      alert(`Demo acquisition added.\nImported: ${d.imported} events across ${Object.keys(d.per_sdv).length} SDV sheets.`);
    });

  const calculate = () =>
    run("CALCULATE RATING : Afficher_Calcul running…", async () => {
      await api.post("/rating/calculate");
      openTab("RATING");
    });

  const eraseAll = () => {
    if (!project) return alert("No project to erase.");
    if (!window.confirm("This operation will delete all your data. Continue?")) return;
    run("Erase_All2…", () => api.delete("/project"));
  };

  const cellSel = (addr, value) => () => setSelection({ addr, value });

  return (
    <div style={{ minHeight: "100%", background: "#fff" }}>
      {/* row 1 banner */}
      <div style={{ background: NAVY, color: "#fff", padding: "14px 30px", display: "flex", alignItems: "baseline", gap: 30 }}>
        <span style={{ fontSize: 34, fontWeight: 700, letterSpacing: 6 }} data-testid="home-title">ODRIV</span>
        <span style={{ fontSize: 13, opacity: 0.8 }}>{state?.version}</span>
        <span style={{ marginLeft: "auto", fontSize: 11, opacity: 0.6 }}>
          Objective DRIVability — Stellantis/PSA toolchain (Python port)
        </span>
      </div>

      <div style={{ display: "flex", gap: 26, padding: "20px 30px", flexWrap: "wrap" }}>
        {/* PROJECT SETTINGS panel */}
        <div style={{ background: SLATE, padding: "12px 16px 18px", minWidth: 480, color: "#fff" }}>
          <div style={{ fontWeight: 700, fontSize: 14, letterSpacing: 1, marginBottom: 10 }}>PROJECT SETTINGS</div>

          <Label text="ID" />
          <Cell value={project?.id?.slice(0, 13) || ""} readOnly testid="home-id" onClick={cellSel("C7", project?.id || "")} />

          <Label text="NAME / CODE" />
          <Cell value={project?.name_code || ""} readOnly testid="home-name_code"
            onClick={cellSel("C9", project?.name_code || "")} wide />

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", columnGap: 18 }}>
            {FIELDS.map(([label, key, opts]) => (
              <div key={key}>
                <Label text={label} />
                {project ? (
                  opts === "text" ? (
                    <input
                      defaultValue={project[key] || ""} onBlur={onField(key)}
                      data-testid={`home-field-${key}`}
                      style={inputStyle} onFocus={cellSel(label, project[key] || "")}
                    />
                  ) : (
                    <select
                      value={project[key] || ""} onChange={onField(key)}
                      data-testid={`home-field-${key}`} style={inputStyle}
                      onFocus={cellSel(label, project[key] || "")}
                    >
                      <option value="">—</option>
                      {(opts || dynamicOpts(key, lists)).map((o) => (
                        <option key={o} value={o}>{o}</option>
                      ))}
                    </select>
                  )
                ) : (
                  <Cell value="" readOnly testid={`home-field-${key}`} />
                )}
              </div>
            ))}
          </div>
          {!project && (
            <div style={{ marginTop: 12, fontSize: 12, background: "#FFFFFFAA", color: "#333", padding: "6px 8px" }}>
              No project. Click <b>NEW PROJECT</b> to start.
            </div>
          )}
        </div>

        {/* button stack */}
        <div style={{ display: "flex", flexDirection: "column", gap: 7, minWidth: 240 }}>
          <button className="xl-btn primary" data-testid="btn-new-project"
            onClick={async () => { await ensureLists(); setModal("new"); }}>NEW PROJECT</button>
          <button className="xl-btn" data-testid="btn-add-file" onClick={addFile}>ADD FILE TO DATABASE</button>
          <button className="xl-btn" data-testid="btn-add-subjective" disabled title="Empty stub in v29.2.1 AT">ADD FILE (SUBJECTIVE)</button>
          <button className="xl-btn" data-testid="btn-open-database" onClick={() => setModal("db")}>OPEN DATABASE</button>
          <button className="xl-btn primary" data-testid="btn-calculate-rating" onClick={calculate}
            disabled={!state?.total_events}>CALCULATE RATING</button>
          <button className="xl-btn" data-testid="btn-create-report" onClick={() => setModal("report")}
            disabled={!state?.has_rating}>CREATE REPORT</button>
          <button className="xl-btn danger" data-testid="btn-erase-all" onClick={eraseAll}>ERASE ALL DATA</button>
          <button className="xl-btn" data-testid="btn-versions" onClick={() => openTab("VERSIONS")}>VERSIONS</button>
          <div style={{ borderTop: "1px solid #ccc", margin: "6px 0" }} />
          <button className="xl-btn" data-testid="btn-demo-data" onClick={demo} disabled={!project}>
            ⚙ Generate demo acquisition
          </button>
          <a className="xl-btn" style={{ textAlign: "center", textDecoration: "none" }}
            href={`${api.defaults.baseURL}/import/sample`} data-testid="btn-download-sample">
            ⬇ Download sample TRIE file
          </a>
          <input ref={fileRef} type="file" accept=".xlsx,.xlsm" hidden onChange={onFilePicked}
            data-testid="file-input" />
        </div>

        {/* status panel */}
        <div style={{ minWidth: 280, flex: 1 }}>
          <table className="xl-grid" style={{ width: "100%" }}>
            <tbody>
              <tr><td className="rowhead">Events in DB</td><td className="num" data-testid="home-total-events">{state?.total_events ?? 0}</td></tr>
              <tr><td className="rowhead">SDV sheets generated</td><td className="num" data-testid="home-sheet-count">{state?.sheets?.length ?? 0}</td></tr>
              <tr><td className="rowhead">Rating calculated</td><td data-testid="home-has-rating">{state?.has_rating ? "Yes" : "No"}</td></tr>
            </tbody>
          </table>
          <div style={{ marginTop: 10, fontSize: 11, color: "#555" }}>Moniteur</div>
          <div className={`moniteur ${busyMsg ? "busy" : ""}`} data-testid="moniteur">
            {busyMsg ? <><span className="spin" /> {busyMsg}</> : lastLog || "Ready."}
          </div>
          <div style={{ marginTop: 14, fontSize: 11.5, color: "#777", lineHeight: 1.6 }}>
            Workflow: NEW PROJECT → ADD FILE TO DATABASE (acquisition .xlsx with a <b>TRIE</b> tab)
            → CALCULATE RATING → review the RATING scorecard &amp; SDV sheets → CREATE REPORT.
          </div>
        </div>
      </div>

      {modal === "new" && (
        <NewProjectModal lists={lists} onClose={() => setModal(null)}
          onCreated={async () => { setModal(null); await refresh(); }} />
      )}
      {modal === "db" && <OpenDatabaseModal onClose={() => setModal(null)} />}
      {modal === "report" && <DocVersionsModal onClose={() => setModal(null)} />}
    </div>
  );
}

const inputStyle = {
  width: "100%", border: "1px solid #9AB", padding: "4px 6px", fontSize: 12.5,
  background: "#fff", color: "#222", marginBottom: 2, borderRadius: 0,
};

function dynamicOpts(key, lists) {
  if (!lists) return ["MDL2"];
  if (key.includes("milestone")) return Object.keys(lists.milestones || {});
  if (key === "area") return lists.areas || [];
  return [];
}

function Label({ text }) {
  return <div style={{ fontSize: 10.5, letterSpacing: 0.5, margin: "8px 0 2px", opacity: 0.85 }}>{text}</div>;
}

function Cell({ value, readOnly, testid, onClick, wide }) {
  return (
    <div onClick={onClick} data-testid={testid}
      style={{ ...inputStyle, minHeight: 24, cursor: "default", width: wide ? "100%" : undefined, background: readOnly ? "#EDF1F4" : "#fff" }}>
      {value}
    </div>
  );
}
