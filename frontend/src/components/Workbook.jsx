import { useCallback, useEffect, useState, createContext, useContext } from "react";
import { api } from "@/lib/api";
import HomeSheet from "@/sheets/HomeSheet";
import RatingSheet from "@/sheets/RatingSheet";
import SdvSheet from "@/sheets/SdvSheet";
import ConfigSheet from "@/sheets/ConfigSheet";
import LogSheet from "@/sheets/LogSheet";
import VersionsSheet from "@/sheets/VersionsSheet";
import { UnlockModal } from "@/components/modals";

export const WorkbookCtx = createContext(null);
export const useWorkbook = () => useContext(WorkbookCtx);

const CONFIG_TABS = [
  ["SETTINGS", "settings_blocks"],
  ["DEFINITION SDV", "definitions"],
  ["TARGETS", "targets"],
  ["CONFIGURATIONS SEETINGS", "priorisation"],
  ["CONFIGURATIONS", "configurations"],
  ["SDV MANAGER", "catalog"],
  ["TARGET VEHICLE", "target_vehicle"],
  ["Calculs", "calculs"],
  ["Graph_status", "thresholds"],
  ["cfg_criticity", "criticity"],
];

export default function Workbook() {
  const [state, setState] = useState(null);
  const [loadError, setLoadError] = useState(null);
  const [active, setActiveState] = useState(() => decodeURIComponent(window.location.hash.slice(1)) || "HOME");
  const [unlocked, setUnlocked] = useState(false);
  const [showUnlock, setShowUnlock] = useState(false);
  const [busy, setBusy] = useState(null);
  const [selection, setSelection] = useState({ addr: "A1", value: "" });

  const setActive = useCallback((name) => {
    window.location.hash = encodeURIComponent(name);
    setActiveState(name);
  }, []);

  useEffect(() => {
    const onHash = () => setActiveState(decodeURIComponent(window.location.hash.slice(1)) || "HOME");
    window.addEventListener("hashchange", onHash);
    return () => window.removeEventListener("hashchange", onHash);
  }, []);

  const refresh = useCallback(async () => {
    try {
      const res = await api.get("/state");
      setState(res.data);
      setLoadError(null);
    } catch (e) {
      console.error("state fetch failed", e);
      setLoadError(e.response?.data?.detail || e.message || "Backend unreachable");
    }
  }, []);

  useEffect(() => { refresh(); }, [refresh]);

  const openTab = useCallback((name) => setActive(name), [setActive]);

  const ctx = { state, refresh, openTab, busy, setBusy, setSelection, unlocked };

  const sdvSheets = (state?.sheets || []).map((s) => s.name);
  const configNames = CONFIG_TABS.map(([n]) => n);

  let content = null;
  if (!state && loadError) {
    content = (
      <div style={{ padding: 40, maxWidth: 560 }} data-testid="backend-error">
        <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 8 }}>DriveScope can’t reach the API</div>
        <div style={{ color: "#666", marginBottom: 14, lineHeight: 1.5 }}>
          {loadError}. Start the backend on port 8000 (see README), then retry.
        </div>
        <button className="xl-btn primary" onClick={refresh} data-testid="backend-retry-btn">Retry connection</button>
      </div>
    );
  } else if (!state) {
    content = <div style={{ padding: 40, color: "#666" }}>Loading DriveScope…</div>;
  } else if (active === "HOME") content = <HomeSheet />;
  else if (active === "RATING") content = <RatingSheet />;
  else if (active === "VERSIONS") content = <VersionsSheet />;
  else if (active === "Macro_Log") content = <LogSheet />;
  else if (configNames.includes(active)) {
    const section = CONFIG_TABS.find(([n]) => n === active)[1];
    content = <ConfigSheet sheetName={active} section={section} />;
  } else content = <SdvSheet key={active} name={active} />;

  return (
    <WorkbookCtx.Provider value={ctx}>
      <div className="xl-titlebar" data-testid="titlebar">
        <span className="logo">DriveScope</span>
        <span className="doc-name">ODRIV_v29_2_1_AT — drivability rating workbook</span>
        <span style={{ fontSize: 11, opacity: 0.85 }}>{state?.version || ""}</span>
      </div>
      <div className="xl-formulabar">
        <div className="xl-namebox" data-testid="name-box">{selection.addr}</div>
        <div className="xl-fx">fx</div>
        <div className="xl-formula" data-testid="formula-bar">{String(selection.value ?? "")}</div>
      </div>
      <div className="xl-sheet-area" data-testid="sheet-area">{content}</div>
      <div className="xl-tabstrip">
        <div className="xl-tabs" data-testid="tab-strip">
          {["HOME", "RATING"].map((t) => (
            <Tab key={t} name={t} active={active} onClick={openTab} />
          ))}
          {sdvSheets.map((t) => (
            <Tab key={t} name={t} active={active} onClick={openTab} />
          ))}
          {["VERSIONS", "Macro_Log"].map((t) => (
            <Tab key={t} name={t} active={active} onClick={openTab} />
          ))}
          {unlocked && CONFIG_TABS.map(([t]) => (
            <Tab key={t} name={t} active={active} onClick={openTab} config />
          ))}
        </div>
        <button
          className="xl-lock-btn" data-testid="unlock-sheets-btn"
          onClick={() => (unlocked ? setUnlocked(false) : setShowUnlock(true))}
        >
          {unlocked ? "🔓 Hide config sheets" : "🔒 Unhide sheets"}
        </button>
      </div>
      {showUnlock && (
        <UnlockModal
          onClose={() => setShowUnlock(false)}
          onUnlock={() => { setUnlocked(true); setShowUnlock(false); }}
        />
      )}
    </WorkbookCtx.Provider>
  );
}

function Tab({ name, active, onClick, config }) {
  return (
    <div
      className={`xl-tab ${active === name ? "active" : ""} ${config ? "config" : ""}`}
      onClick={() => onClick(name)}
      data-testid={`tab-${name}`}
      title={name}
    >
      {name.length > 26 ? name.slice(0, 24) + "…" : name}
    </div>
  );
}
