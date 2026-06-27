import type { OperationProgress, ProcessedFileStats, UploadSummary } from '../types';

export type GenericColumns = {
  company_col: string;
  mst_col: string;
  address_col: string;
  product_col: string;
  qty_col: string;
  price_col: string;
  output_col: string;
  invoice_status_col: string;
  invoice_status_skip_values: string[];
};

const purchaseColumnLetters: Record<string, string> = {
  'Tên công ty': 'F',
  MST: 'G',
  'Tên hàng': 'M',
  'Số lượng': 'O',
  'Đơn giá': 'P',
  'Mã VT': 'L',
};

const salesColumnLetters: Record<string, string> = {
  'Tên công ty': 'I',
  MST: 'J',
  'Tên hàng': 'M',
  'Số lượng': 'O',
  'Đơn giá': 'P',
  'Mã VT': 'L',
};

export function UploadStage({ title, summary, disabled, onUpload }: { title: string; summary: UploadSummary | null; disabled: boolean; onUpload: (file: File | undefined) => void }) {
  const stepNumber = title.includes('bán') ? 6 : 1;
  const uploadLabel = `Chọn ${title}`;
  return (
    <div className="center-stage">
      <span className="upload-step-badge">BƯỚC {stepNumber}</span>
      <p className="description">Tải lên file Excel <strong>.xls</strong>, <strong>.xlsx</strong> hoặc <strong>.xlsm</strong> cho <strong>{title}</strong>.</p>
      <label className={`upload-label ${summary ? 'has-file' : ''}`}>
        <input className="upload-input" type="file" accept=".xls,.xlsx,.xlsm" disabled={disabled} onChange={(event) => onUpload(event.currentTarget.files?.[0])} />
        <span className="upload-mark">XLSX</span>
        <strong>{summary?.original_name || 'Kéo thả hoặc chọn file Excel'}</strong>
        <small>{summary ? `${summary.columns.length} cột đã đọc` : 'Hệ thống sẽ đọc bảng, nhận diện cột và giữ nguyên file gốc.'}</small>
        <span className="upload-button">{summary ? `Đổi ${title}` : uploadLabel}</span>
      </label>
    </div>
  );
}

export function SalesEntryStage({ salesFile, processedPurchaseReady, processedPurchaseStats, disabled, onSalesUpload, onProcessedPurchaseUpload }: { salesFile: UploadSummary | null; processedPurchaseReady: boolean; processedPurchaseStats: ProcessedFileStats | null; disabled: boolean; onSalesUpload: (file: File | undefined) => void; onProcessedPurchaseUpload: (file: File | undefined) => void }) {
  return (
    <div className="skip-entry-stage">
      <section className="skip-required-panel">
        <div className="stage-toolbar">
          <h3>Bỏ qua mua vào</h3>
          <p>Upload file mua vào đã xử lý trước khi chạy các bước bán ra.</p>
        </div>
        <ProcessedFileUpload label="Mua vào đã xử lý" ready={processedPurchaseReady} stats={processedPurchaseStats} disabled={disabled} onUpload={onProcessedPurchaseUpload} />
      </section>
      <UploadStage title="HD bán ra" summary={salesFile} disabled={disabled} onUpload={onSalesUpload} />
    </div>
  );
}

export function ProcessedFileUpload({ label, ready, stats, disabled, onUpload }: { label: string; ready: boolean; stats: ProcessedFileStats | null; disabled: boolean; onUpload: (file: File | undefined) => void }) {
  return (
    <label className={`processed-file-upload ${ready ? 'has-file' : ''}`}>
      <input className="upload-input" type="file" accept=".xls,.xlsx,.xlsm" disabled={disabled} onChange={(event) => onUpload(event.currentTarget.files?.[0])} />
      <span className="upload-mark">FDI</span>
      <strong>{ready ? `${label} đã sẵn sàng` : `Chọn ${label}`}</strong>
      <small>{ready ? 'Có thể dùng file này cho stage sau.' : 'Bắt buộc khi bỏ qua các stage xử lý trước đó.'}</small>
      <ProcessedStatsSummary stats={stats} />
      <span className="upload-button">{ready ? `Đổi ${label}` : `Tải ${label}`}</span>
    </label>
  );
}

export function ProcessedStatsSummary({ stats }: { stats: ProcessedFileStats | null }) {
  if (!stats) return <div className="processed-stats muted">Chưa có thống kê file đã xử lý.</div>;
  return (
    <div className="processed-stats">
      <span>Công ty <strong>{formatCount(stats.processed_company_count)} / {formatCount(stats.company_count)}</strong></span>
      <span>Dòng hàng <strong>{formatCount(stats.processed_product_row_count)} / {formatCount(stats.product_row_count)}</strong></span>
    </div>
  );
}

export function MappingStage({ summary, phase, scope, setScope }: { summary: UploadSummary | null; phase: 'purchase' | 'sales'; scope: string; setScope: (value: string) => void }) {
  if (!summary) return <PlaceholderStage title="Chưa có file" detail="Quay lại stage tải file trước khi chọn cột." />;
  const previewKeys = summary.preview.length ? Object.keys(summary.preview[0]) : [];
  const columnLetters = phase === 'sales' ? salesColumnLetters : purchaseColumnLetters;
  return (
    <div className="stage-grid">
      <div className="form-panel">
        <p className="description left">File: <strong>{summary.original_name}</strong>. Mặc định Vietmax dùng cột P cho đơn giá bán/mua.</p>
        <label><span>Phạm vi so sánh</span><select value={scope} onChange={(event) => setScope(event.currentTarget.value)}><option value="all_companies">Nhiều công ty</option><option value="same_company">Chỉ cùng công ty</option></select></label>
        <div className="column-grid">{['Tên công ty', 'MST', 'Tên hàng', 'Số lượng', 'Đơn giá', 'Mã VT'].map((label) => {
          const defaultLetter = columnLetters[label];
          return <label key={label}><span>{label}</span><select defaultValue={defaultLetter}><option value={defaultLetter} title="Mặc định theo Vietmax">{defaultLetter}</option>{summary.columns.filter((column) => column.letter !== defaultLetter).map((column) => <option key={`${label}-${column.letter}`} value={column.letter} title={column.label}>{column.letter}</option>)}</select></label>;
        })}</div>
      </div>
      <div className="preview-panel">
        <h3>Xem trước dữ liệu</h3>
        <div className="preview-scroll"><table><thead><tr>{previewKeys.map((key) => <th key={key}>{key}</th>)}</tr></thead><tbody>{summary.preview.map((row, index) => <tr key={index}>{previewKeys.map((key) => <td key={key}>{row[key]}</td>)}</tr>)}</tbody></table></div>
      </div>
    </div>
  );
}

export function GenericMappingStage({ summary, columns, onColumnsChange }: { summary: UploadSummary | null; columns: GenericColumns; onColumnsChange: (update: Partial<GenericColumns>) => void }) {
  if (!summary) return <PlaceholderStage title="Chưa có file" detail="Quay lại stage tải file trước khi chọn cột." />;
  const previewKeys = summary.preview.length ? Object.keys(summary.preview[0]) : [];
  const columnFields: Array<{ key: keyof GenericColumns; label: string; allowBlank?: boolean }> = [
    { key: 'company_col', label: 'Tên công ty' },
    { key: 'mst_col', label: 'MST' },
    { key: 'address_col', label: 'Địa chỉ', allowBlank: true },
    { key: 'product_col', label: 'Tên hàng' },
    { key: 'qty_col', label: 'Số lượng', allowBlank: true },
    { key: 'price_col', label: 'Đơn giá', allowBlank: true },
    { key: 'output_col', label: 'Mã VT' },
    { key: 'invoice_status_col', label: 'Trạng thái HĐ', allowBlank: true },
  ];
  return (
    <div className="stage-grid">
      <div className="form-panel">
        <p className="description left">File: <strong>{summary.original_name}</strong>. Chọn đúng cột trước khi tải danh sách công ty.</p>
        <div className="column-grid">
          {columnFields.map((field) => (
            <label key={field.key}>
              <span>{field.label}</span>
              <select value={String(columns[field.key] || '')} onChange={(event) => onColumnsChange({ [field.key]: event.currentTarget.value } as Partial<GenericColumns>)}>
                {field.allowBlank && <option value="">Không dùng</option>}
                {summary.columns.map((column) => <option key={`${field.key}-${column.letter}`} value={column.letter} title={column.label}>{column.letter}</option>)}
              </select>
            </label>
          ))}
        </div>
      </div>
      <div className="preview-panel">
        <h3>Xem trước dữ liệu</h3>
        <div className="preview-scroll"><table><thead><tr>{previewKeys.map((key) => <th key={key}>{key}</th>)}</tr></thead><tbody>{summary.preview.map((row, index) => <tr key={index}>{previewKeys.map((key) => <td key={key}>{row[key]}</td>)}</tr>)}</tbody></table></div>
      </div>
    </div>
  );
}

export function ProcessStage({
  title,
  detail,
  buttonLabel = 'Xuất file',
  disabled,
  onProcess,
}: {
  title: string;
  detail: string;
  buttonLabel?: string;
  disabled: boolean;
  onProcess: () => void;
}) {
  return (
    <div className="placeholder-stage">
      <h3>{title}</h3>
      <p>{detail}</p>
      <div className="process-actions">
        <button type="button" disabled={disabled} onClick={onProcess}>{buttonLabel}</button>
      </div>
    </div>
  );
}

export function FastImportExportStage({
  processedPurchaseSavedName,
  processedSalesSavedName,
  processedPurchaseStats,
  processedSalesStats,
  busy,
  onProcessedPurchaseUpload,
  onProcessedSalesUpload,
  onDownload,
}: {
  processedPurchaseSavedName: string;
  processedSalesSavedName: string;
  processedPurchaseStats: ProcessedFileStats | null;
  processedSalesStats: ProcessedFileStats | null;
  busy: boolean;
  onProcessedPurchaseUpload: (file: File | undefined) => void;
  onProcessedSalesUpload: (file: File | undefined) => void;
  onDownload: () => void;
}) {
  const canExport = Boolean(processedPurchaseSavedName && processedSalesSavedName);
  const renderUpload = (
    label: string,
    ready: boolean,
    stats: ProcessedFileStats | null,
    onUpload: (file: File | undefined) => void,
  ) => (
    <label className={`inventory-file-status ${ready ? 'ready' : ''}`}>
      <input
        type="file"
        accept=".xls,.xlsx,.xlsm"
        disabled={busy}
        onChange={(event) => {
          onUpload(event.currentTarget.files?.[0]);
          event.currentTarget.value = '';
        }}
      />
      <span>{label}</span>
      <strong>{ready ? 'Đã có FDI đã xử lý' : 'Tải file FDI đã xử lý'}</strong>
      <small>{ready ? processedStatsSentence(stats) : 'Dùng khi muốn bỏ qua các stage trước và xuất FAST trực tiếp.'}</small>
    </label>
  );

  return (
    <div className="inventory-allocation-stage export-only-stage">
      <section className="inventory-allocation-card inventory-export-card">
        <div className="report-stage-heading">
          <span>Xuất FAST</span>
          <h3>Xuất workbook FAST 4 sheet</h3>
          <p>Workbook sẽ gồm Hoadonmuahang, Hoadonbanhang, DMvat_tu và DMkhachhang. Chỉ các cột có dữ liệu trong file mẫu FAST mới được điền.</p>
        </div>
        <div className="inventory-file-grid">
          {renderUpload('FDI mua vào đã xử lý', Boolean(processedPurchaseSavedName), processedPurchaseStats, onProcessedPurchaseUpload)}
          {renderUpload('FDI bán ra đã xử lý', Boolean(processedSalesSavedName), processedSalesStats, onProcessedSalesUpload)}
          <div className={`inventory-file-status ${canExport ? 'ready' : ''}`}>
            <span>Workbook xuất</span>
            <strong>{canExport ? 'Sẵn sàng xuất' : 'Cần đủ 2 file FDI'}</strong>
            <small>{canExport ? 'Một file .xls với các sheet FAST.' : 'Tải hoặc tạo cả FDI mua vào và FDI bán ra trước.'}</small>
          </div>
        </div>
        <div className="export-action-panel">
          <button type="button" disabled={busy || !canExport} onClick={onDownload}>Xuất workbook FAST</button>
        </div>
      </section>
    </div>
  );
}

export function PlaceholderStage({ title, detail }: { title: string; detail: string }) {
  return <div className="placeholder-stage"><h3>{title}</h3><p>{detail}</p></div>;
}

export function LoadingStage({ title, detail, progress }: { title: string; detail: string; progress?: OperationProgress | null }) {
  const percent = Math.max(0, Math.min(100, Number(progress?.percent ?? 0)));
  const hasRealRowProgress = Boolean(progress && Number(progress.total || 0) > 1);
  const rowProgress = hasRealRowProgress ? progress : null;
  return (
    <div className="loading-stage">
      <div className="loading-spinner" aria-hidden="true" />
      <h3>{title}</h3>
      <p>{progress?.label || detail}</p>
      {rowProgress && (
        <div className="loading-progress-block">
          <div className="loading-progress-bar" aria-label="Tiến trình xử lý"><span style={{ width: `${percent}%` }} /></div>
          <strong>{percent}%</strong>
          <small>{formatCount(rowProgress.done)} / {formatCount(rowProgress.total)} dòng</small>
        </div>
      )}
    </div>
  );
}

function processedStatsSentence(stats: ProcessedFileStats | null) {
  if (!stats) return '';
  return `CÃ´ng ty ${formatCount(stats.processed_company_count)}/${formatCount(stats.company_count)}, dÃ²ng hÃ ng ${formatCount(stats.processed_product_row_count)}/${formatCount(stats.product_row_count)}.`;
}

function formatCount(value: number | undefined) {
  return Number(value || 0).toLocaleString('en-US');
}
