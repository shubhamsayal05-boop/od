export default function PptPanel({ tool }) {
  const hasEmbed = Boolean(tool.pptEmbedUrl);
  const hasFile = Boolean(tool.pptFile);

  return (
    <aside className="ppt-panel">
      <div className="ppt-panel-header">
        <h3>Presentation</h3>
      </div>
      <div className="ppt-panel-body">
        {hasEmbed ? (
          <iframe
            className="ppt-embed"
            src={tool.pptEmbedUrl}
            title={`${tool.name} tutorial presentation`}
            allowFullScreen
          />
        ) : hasFile ? (
          <div className="ppt-placeholder">
            <span style={{ fontSize: '2.5rem' }}>📑</span>
            <strong>{tool.name} Tutorial</strong>
            <span>
              Place your .pptx file at <code>public{tool.pptFile}</code> or set an embed URL in
              tools.json.
            </span>
            <a className="btn btn-primary" href={tool.pptFile} download>
              Download PPT
            </a>
          </div>
        ) : (
          <div className="ppt-placeholder">
            <span style={{ fontSize: '2.5rem' }}>📑</span>
            <strong>No presentation linked</strong>
            <span>Add pptFile or pptEmbedUrl in tools.json for this tool.</span>
          </div>
        )}
      </div>
    </aside>
  );
}
