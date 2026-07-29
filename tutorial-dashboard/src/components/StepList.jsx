import { useState } from 'react';

export default function StepList({ steps, activeStep, onStepChange, onComplete }) {
  const completedCount = steps.filter((_, i) => i < activeStep).length;

  return (
    <section className="steps-panel">
      <div className="steps-panel-header">
        <h2>Tutorial Steps</h2>
        <span className="progress-text">
          {completedCount}/{steps.length}
        </span>
      </div>
      <ol className="step-list">
        {steps.map((step, index) => (
          <li key={step.title} className="step-item">
            <button
              type="button"
              className={`step-button ${index === activeStep ? 'active' : ''} ${
                index < activeStep ? 'completed' : ''
              }`}
              onClick={() => onStepChange(index)}
            >
              <span className="step-number">{index + 1}</span>
              <div>
                <div className="step-preview-title">{step.title}</div>
                <div className="step-preview-desc">{step.description}</div>
              </div>
            </button>
          </li>
        ))}
      </ol>
      <div className="step-detail">
        <h3>{steps[activeStep].title}</h3>
        <p>{steps[activeStep].description}</p>
        {steps[activeStep].tips?.length > 0 && (
          <div className="step-tips">
            <h4>Tips</h4>
            <ul>
              {steps[activeStep].tips.map((tip) => (
                <li key={tip}>{tip}</li>
              ))}
            </ul>
          </div>
        )}
        <div className="step-nav">
          <button
            type="button"
            className="btn btn-secondary"
            disabled={activeStep === 0}
            onClick={() => onStepChange(activeStep - 1)}
          >
            ← Previous
          </button>
          {activeStep < steps.length - 1 ? (
            <button
              type="button"
              className="btn btn-primary"
              onClick={() => {
                onComplete(activeStep);
                onStepChange(activeStep + 1);
              }}
            >
              Next →
            </button>
          ) : (
            <button type="button" className="btn btn-primary" onClick={() => onComplete(activeStep)}>
              Finish ✓
            </button>
          )}
        </div>
      </div>
    </section>
  );
}
