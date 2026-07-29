import { useMemo, useState } from 'react';
import tools from '../data/tools.json';
import ToolCard from '../components/ToolCard';

export default function Home() {
  const [query, setQuery] = useState('');

  const filteredTools = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return tools;
    return tools.filter(
      (tool) =>
        tool.name.toLowerCase().includes(q) ||
        tool.description.toLowerCase().includes(q) ||
        tool.category.toLowerCase().includes(q)
    );
  }, [query]);

  return (
    <>
      <header className="app-header">
        <h1>Tool Tutorial Dashboard</h1>
        <p>
          Step-by-step guides for every tool your team uses. Pick a tool to open its tutorial
          dashboard.
        </p>
      </header>

      <div className="search-bar">
        <input
          type="search"
          className="search-input"
          placeholder="Search tools..."
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          aria-label="Search tools"
        />
      </div>

      {filteredTools.length === 0 ? (
        <div className="empty-state">
          <p>No tools match &ldquo;{query}&rdquo;</p>
        </div>
      ) : (
        <div className="tool-grid">
          {filteredTools.map((tool) => (
            <ToolCard key={tool.id} tool={tool} />
          ))}
        </div>
      )}
    </>
  );
}
