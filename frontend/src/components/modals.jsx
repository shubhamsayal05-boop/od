import { useEffect, useRef, useState } from "react";
import { api, fmt } from "@/lib/api";

export function Modal({ title, onClose, children, footer, wide }) {
  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={wide ? { minWidth: 760 } : undefined}>
        <div className="modal-title">
          <span>{title}</span>
          <span className="x" onClick={onClose} data-testid="modal-close">✕</span>
        </div>
        <div className="modal-body">{children}</div>
        {footer && <div className="modal-footer">{footer}</div>}
      </div>
    </div>
  );
}

/* ---------------- unlock (password = UNLOCK, like the original) -------- */
export function UnlockModal({ onClose, onUnlock }) {
  const [pwd, setPwd] = useState("");
  const [err, setErr] = useState(false);
  const submit = () => {
    if (pwd.trim().toUpperCase() === "UNLOCK") onUnlock();
    else setErr(true);
  };
  return (
    <Modal
      title="Unlock sheets" onClose={onClose}
      footer={
        <>
          <button className="xl-btn" onClick={onClose} data-testid="unlock-cancel">Cancel</button>
          <button className="xl-btn primary" onClick={submit} data-testid="unlock-submit">OK</button>
        </>
      }
    >
      <div className="form-row">
        <label>Password</label>
        <input
          type="password" value={pwd} autoFocus data-testid="unlock-password"
          onChange={(e) => { setPwd(e.target.value); setErr(false); }}
          onKeyDown={(e) => e.key === "Enter" && submit()}
        />
      </div>
      <div style={{ fontSize: 11, color: err ? "#c00" : "#777" }}>
        {err ? "Wrong password." : "Hint: same password as the original tool (UNLOCK)."}
      </div>
    </Modal>
  );
}

/* ---------------- new project ---------------- */
export function NewProjectModal({ onClose, onCreated, lists }) {
  const [form, setForm] = useState({
    name_code: "", mode: "AUTO", fuel: "DIESEL", gears: "8",
    software_milestone: "MDL2", priority: "PREMIUM", version: "4.6",
    odriv_milestone: "MDL2", area: "EUROPE", target_vehicle: "P8 MHEV MDL2",
    number_of_gears: "8",
  });
  const [saving, setSaving] = useState(false);
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  const milestones = Object.keys(lists?.milestones || { MDL2: 4 });
  const areas = lists?.areas || ["EUROPE"];
  const create = async () => {
    if (!form.name_code.trim()) return;
    setSaving(true);
    try {
      await api.post("/project/new", form);
      onCreated();
    } finally { setSaving(false); }
  };
  const Sel = ({ k, opts }) => (
    <select value={form[k]} onChange={set(k)} data-testid={`np-${k}`}>
      {opts.map((o) => <option key={o} value={o}>{o}</option>)}
    </select>
  );
  return (
    <Modal
      title="NEW PROJECT" onClose={onClose}
      footer={
        <>
          <button className="xl-btn" onClick={onClose} data-testid="np-cancel">Cancel</button>
          <button className="xl-btn primary" onClick={create} disabled={saving || !form.name_code.trim()} data-testid="np-create">
            {saving ? "Creating…" : "Create project"}
          </button>
        </>
      }
    >
      <div style={{ fontSize: 11.5, color: "#a33", marginBottom: 10 }}>
        Creating a new project erases all existing data (Erase_All2).
      </div>
      <div className="form-row"><label>NAME / CODE</label>
        <input value={form.name_code} onChange={set("name_code")} autoFocus data-testid="np-name_code" placeholder="K0_2.2 diesel_AT8" /></div>
      <div className="form-row"><label>MODE</label><Sel k="mode" opts={["AUTO", "ECO", "SPORT", "MANUAL", "Hybrid", "ZEV", "eSAVE"]} /></div>
      <div className="form-row"><label>FUEL</label><Sel k="fuel" opts={["GASOLINE", "DIESEL"]} /></div>
      <div className="form-row"><label>GEARS</label><Sel k="gears" opts={["5", "6", "7", "8", "9", "10"]} /></div>
      <div className="form-row"><label>SOFTWARE MILESTONE</label><Sel k="software_milestone" opts={milestones} /></div>
      <div className="form-row"><label>PRIORITY (target range)</label><Sel k="priority" opts={["PREMIUM", "MAINSTREAM"]} /></div>
      <div className="form-row"><label>VERSION (drive)</label><Sel k="version" opts={["4.2", "4.6"]} /></div>
      <div className="form-row"><label>ODRIV MILESTONE</label><Sel k="odriv_milestone" opts={milestones} /></div>
      <div className="form-row"><label>AREA</label><Sel k="area" opts={areas} /></div>
      <div className="form-row"><label>TARGET VEHICLE</label>
        <input value={form.target_vehicle} onChange={set("target_vehicle")} data-testid="np-target_vehicle" /></div>
      <div className="form-row"><label>NUMBER OF GEARS</label><Sel k="number_of_gears" opts={["5", "6", "7", "8", "9", "10"]} /></div>
    </Modal>
  );
}

/* ---------------- DocVersions -> report ---------------- */
export function DocVersionsModal({ onClose }) {
  const [versions, setVersions] = useState(["", "", "", ""]);
  const [generating, setGenerating] = useState(null);
  const labels = ["Document version", "AVLD version", "Software version", "Calibration version"];
  const gen = async (fmtType) => {
    setGenerating(fmtType);
    try {
      const res = await api.post(`/report/${fmtType}`, { doc_versions: versions }, { responseType: "blob" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([res.data]));
      a.download = `ODRIV_report.${fmtType}`;
      document.body.appendChild(a); a.click(); a.remove();
    } catch (e) {
      alert("Report failed: " + (e.response?.status === 400 ? "calculate the rating first" : e.message));
    } finally { setGenerating(null); }
  };
  return (
    <Modal
      title="DocVersions — CREATE REPORT" onClose={onClose}
      footer={
        <>
          <button className="xl-btn" onClick={onClose} data-testid="docv-cancel">Cancel</button>
          <button className="xl-btn primary" onClick={() => gen("pptx")} disabled={!!generating} data-testid="docv-pptx">
            {generating === "pptx" ? "Generating…" : "OK, suite → PowerPoint (.pptx)"}
          </button>
          <button className="xl-btn primary" onClick={() => gen("pdf")} disabled={!!generating} data-testid="docv-pdf">
            {generating === "pdf" ? "Generating…" : "PDF report"}
          </button>
        </>
      }
    >
      <div style={{ fontSize: 11.5, color: "#555", marginBottom: 10 }}>
        Enter the document version numbers (DocVersions D7:D10), then continue to the report engine.
      </div>
      {labels.map((l, i) => (
        <div className="form-row" key={l}>
          <label>{l}</label>
          <input value={versions[i]} data-testid={`docv-field-${i}`}
            onChange={(e) => setVersions(versions.map((v, j) => (j === i ? e.target.value : v)))} />
        </div>
      ))}
    </Modal>
  );
}

/* ---------------- OPEN DATABASE (form.frm) ---------------- */
export function OpenDatabaseModal({ onClose }) {
  const [events, setEvents] = useState([]);
  const [total, setTotal] = useState(0);
  const [skip, setSkip] = useState(0);
  const [sdv, setSdv] = useState("");
  const [search, setSearch] = useState("");
  const [editing, setEditing] = useState(null);
  const limit = 25;

  const load = async (s = skip) => {
    const res = await api.get("/events", { params: { skip: s, limit, sdv: sdv || undefined, search: search || undefined } });
    setEvents(res.data.events);
    setTotal(res.data.total);
  };
  useEffect(() => { load(0); setSkip(0); }, [sdv]); // eslint-disable-line react-hooks/exhaustive-deps

  const del = async (id) => {
    if (!window.confirm("Delete this record from the database?")) return;
    await api.delete(`/events/${id}`);
    load();
  };

  return (
    <Modal title="OPEN DATABASE — _OdrivDB (events)" onClose={onClose} wide>
      <div className="cfg-toolbar">
        <input placeholder="Search…" value={search} onChange={(e) => setSearch(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && load(0)} data-testid="db-search" style={{ width: 200 }} />
        <input placeholder="Filter by SDV…" value={sdv} onChange={(e) => setSdv(e.target.value)}
          data-testid="db-sdv-filter" style={{ width: 220 }} />
        <button className="xl-btn" onClick={() => load(0)} data-testid="db-refresh">Refresh</button>
        <span style={{ fontSize: 11.5, color: "#666" }}>{total} records</span>
        <span style={{ marginLeft: "auto", display: "flex", gap: 4 }}>
          <button className="xl-btn" disabled={skip === 0} data-testid="db-prev"
            onClick={() => { const s = Math.max(0, skip - limit); setSkip(s); load(s); }}>◀</button>
          <button className="xl-btn" disabled={skip + limit >= total} data-testid="db-next"
            onClick={() => { const s = skip + limit; setSkip(s); load(s); }}>▶</button>
        </span>
      </div>
      <table className="xl-grid" style={{ width: "100%" }}>
        <thead>
          <tr><th>Id BdD</th><th>SDV</th><th>Sub Event</th><th>File</th><th>Driv</th><th>Dyn</th><th></th></tr>
        </thead>
        <tbody>
          {events.map((ev) => (
            <tr key={ev.id} data-testid={`db-row-${ev.id.slice(0, 8)}`}>
              <td style={{ fontFamily: "monospace", fontSize: 11 }}>{ev.id.slice(0, 8)}</td>
              <td>{ev.sdv || <i style={{ color: "#999" }}>unclassified</i>}</td>
              <td>{String(findChan(ev.channels, "Sub Event Name") ?? "")}</td>
              <td style={{ fontSize: 11 }}>{ev.file}</td>
              <td style={{ color: dotColor(ev.driv) }}>{ev.driv ? `${fmt(ev.driv.indice, 2)} P${ev.driv.priority}` : "-"}</td>
              <td style={{ color: dotColor(ev.dyn) }}>{ev.dyn ? `${fmt(ev.dyn.indice, 2)} P${ev.dyn.priority}` : "-"}</td>
              <td>
                <button className="xl-btn" style={{ padding: "1px 7px" }} onClick={() => setEditing(ev)} data-testid={`db-edit-${ev.id.slice(0, 8)}`}>Edit</button>{" "}
                <button className="xl-btn" style={{ padding: "1px 7px" }} onClick={() => del(ev.id)} data-testid={`db-del-${ev.id.slice(0, 8)}`}>Del</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {editing && (
        <EventEditModal event={editing} onClose={() => setEditing(null)}
          onSaved={() => { setEditing(null); load(); }} />
      )}
    </Modal>
  );
}

function dotColor(part) {
  if (!part) return "#999";
  return { RED: "#c00", YELLOW: "#b8860b", GREEN: "#070" }[part.color] || "#333";
}

function findChan(channels, name) {
  for (const [k, v] of Object.entries(channels || {})) {
    if (k.toLowerCase().includes(name.toLowerCase())) return v;
  }
  return null;
}

function EventEditModal({ event, onClose, onSaved }) {
  const [channels, setChannels] = useState({ ...event.channels });
  const [saving, setSaving] = useState(false);
  const save = async () => {
    setSaving(true);
    try {
      await api.put(`/events/${event.id}`, { channels });
      onSaved();
    } finally { setSaving(false); }
  };
  return (
    <Modal title={`Edit record ${event.id.slice(0, 8)} — ${event.sdv}`} onClose={onClose} wide
      footer={
        <>
          <button className="xl-btn" onClick={onClose}>Cancel</button>
          <button className="xl-btn primary" onClick={save} disabled={saving} data-testid="event-save">
            {saving ? "Saving…" : "Save record"}
          </button>
        </>
      }>
      <div style={{ maxHeight: 420, overflow: "auto" }}>
        <table className="xl-grid" style={{ width: "100%" }}>
          <thead><tr><th style={{ width: "55%" }}>Channel (Family, Channel)</th><th>Value</th></tr></thead>
          <tbody>
            {Object.entries(channels).map(([k, v]) => (
              <tr key={k}>
                <td style={{ fontSize: 11 }}>{k.replace(/\u00b7/g, ".")}</td>
                <td>
                  <input className="xl-cell-input" value={String(v ?? "")} data-testid={`chan-${k.slice(0, 12)}`}
                    onChange={(e) => setChannels({ ...channels, [k]: e.target.value })} />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}
