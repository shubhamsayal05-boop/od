import { useWorkbook } from "@/components/Workbook";

const CHANGELOG = [
  ["v29.2.1 AT (Python)", "Full Python/Web rebuild: FastAPI engine + React Excel-style UI. Import pipeline (TRIE), conditionSDV classifier (64 rules), agreement-index scoring, priorisation grids, criticity, status dots, NoteGlobale verdicts, PPTX/PDF reporting."],
  ["v29.2.1 AT", "Original Excel build — 30 sheets, 117 VBA modules, 40,468 lines (Stellantis/PSA GALIMA toolchain)."],
  ["v29.2.0 AT", "Automatic Transmission variant; SETTINGS milestone blocks per SDV."],
  ["…", "See the original VERSIONS sheet for the complete changelog history."],
];

export default function VersionsSheet() {
  const { openTab } = useWorkbook();
  return (
    <div className="cfg-sheet">
      <div style={{ display: "flex", gap: 14, alignItems: "center" }}>
        <h2>VERSIONS — tool changelog</h2>
        <button className="xl-btn" onClick={() => openTab("HOME")} data-testid="versions-exit">Exit</button>
      </div>
      <table className="xl-grid" style={{ marginTop: 8 }}>
        <thead><tr><th style={{ minWidth: 160 }}>Version</th><th style={{ minWidth: 620 }}>Changes</th></tr></thead>
        <tbody>
          {CHANGELOG.map(([v, c]) => (
            <tr key={v}><td style={{ fontWeight: 700 }}>{v}</td><td style={{ whiteSpace: "normal" }}>{c}</td></tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
