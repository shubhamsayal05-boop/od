import { useState } from 'react';
import { Link, useParams } from 'react-router-dom';
import tools from '../data/tools.json';
import StepList from '../components/StepList';
import PptPanel from '../components/PptPanel';

export default function ToolTutorial() {
  const { toolId } = useParams();
  const tool = tools.find((t) => t.id === toolId);
  const [activeStep, setActiveStep] = useState(0);

  if (!tool) {
    return (
      <div className="not-found">
        <h2>Tool not found</h2>
        <p>The tutorial you are looking for does not exist.</p>
        <Link to="/" className="btn btn-primary" style={{ marginTop: '1rem', display: 'inline-flex' }}>
          Back to dashboard
        </Link>
      </div>
    );
  }

  return (
    <>
      <Link to="/" className="back-link">
        ← All tools
      </Link>

      <header className="tool-detail-header">
        <span className="tool-detail-icon" aria-hidden="true">
          {tool.icon}
        </span>
        <div>
          <h1>{tool.name}</h1>
          <p>{tool.description}</p>
          <div className="tool-actions">
            <span className="badge">{tool.category}</span>
            {tool.estimatedMinutes && (
              <span className="badge badge-muted">~{tool.estimatedMinutes} min</span>
            )}
            {tool.pptFile && (
              <a className="btn btn-secondary" href={tool.pptFile} download>
                Download PPT
              </a>
            )}
          </div>
        </div>
      </header>

      <div className="dashboard-layout">
        <StepList
          steps={tool.steps}
          activeStep={activeStep}
          onStepChange={setActiveStep}
          onComplete={() => {}}
        />
        <PptPanel tool={tool} />
      </div>
    </>
  );
}
