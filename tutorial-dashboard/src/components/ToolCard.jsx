import { Link } from 'react-router-dom';

export default function ToolCard({ tool }) {
  return (
    <Link to={`/tool/${tool.id}`} className="tool-card">
      <span className="tool-card-icon" aria-hidden="true">
        {tool.icon}
      </span>
      <h2>{tool.name}</h2>
      <p>{tool.description}</p>
      <div className="tool-card-meta">
        <span className="badge">{tool.category}</span>
        <span className="badge badge-muted">{tool.steps.length} steps</span>
        {tool.estimatedMinutes && (
          <span className="badge badge-muted">~{tool.estimatedMinutes} min</span>
        )}
      </div>
    </Link>
  );
}
