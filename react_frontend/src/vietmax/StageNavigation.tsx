import type { StageDefinition, StageId } from './workflowStages';

export function StageNavigation({ stages, stage, busy, canEnterStage, goToStage }: { stages: StageDefinition[]; stage: StageId; busy: boolean; canEnterStage: (target: StageId) => boolean; goToStage: (target: StageId) => void }) {
  const purchaseStages = stages.filter((item) => item.phase === 'purchase');
  const salesStages = stages.filter((item) => item.phase === 'sales');
  const inventoryStages = stages.filter((item) => item.phase === 'inventory');
  const fastStages = stages.filter((item) => item.phase === 'fast');
  const shellStages = stages.filter((item) => item.phase === 'generic' || item.phase === 'price');

  if (purchaseStages.length || salesStages.length || inventoryStages.length || fastStages.length) {
    return (
      <div className="stage-groups" aria-label="Stage navigation">
        <StageGroup title="Mua vào" stages={purchaseStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />
        <StageGroup title="Bán ra" stages={salesStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />
        <StageGroup title="Tồn kho" stages={inventoryStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />
        <StageGroup title="Xuất FAST" stages={fastStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />
      </div>
    );
  }

  return (
    <div className="stage-groups" aria-label="Stage navigation">
      <StageGroup title="Profile" stages={shellStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />
    </div>
  );
}

function StageGroup({ title, stages, stage, busy, canEnterStage, goToStage }: { title: string; stages: StageDefinition[]; stage: StageId; busy: boolean; canEnterStage: (target: StageId) => boolean; goToStage: (target: StageId) => void }) {
  return (
    <div className="stage-group">
      <span className="stage-group-label">{title}</span>
      <div className="stage-group-pills">
        {stages.map((item) => (
          <button
            key={item.id}
            className={`step-pill ${item.id === stage ? 'active' : ''} ${item.phase}`}
            disabled={!canEnterStage(item.id) || busy}
            onClick={() => goToStage(item.id)}
            type="button"
            title={`${item.id}. ${item.label}`}
          >
            <span>{item.id}.</span> {item.short}
          </button>
        ))}
      </div>
    </div>
  );
}
