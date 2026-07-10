import { useMemo, useState } from 'react';
import { analyzeEstimateWorkbook, exportEstimateWorkbook, uploadEstimateWorkbook } from '../api';
import type { EstimateAnalysis, EstimateColumnConfig, EstimateSheetInfo, EstimateSheetSelection, EstimateUploadSummary } from '../types';
import type { StageId } from '../vietmax/workflowStages';

type Props = {
  licenseReady: boolean;
  onStatus: (message: string) => void;
  stage: StageId;
  onStageChange: (stage: StageId) => void | Promise<void>;
};

type WarningSection = {
  key: string;
  label: string;
  detail: string;
};

type ColumnField = readonly [string, string];

const warningSections: WarningSection[] = [
  { key: 'identity_mismatches', label: 'Lệch Dự thầu / Chiết tính', detail: 'Dòng dự thầu không khớp tổng hoặc metadata từ block chiết tính.' },
  { key: 'calculation_mismatches', label: 'Lệch công thức', detail: 'Tổng các nhóm VL/NC/MTC... không khớp thành tiền dự thầu.' },
  { key: 'unclassified_rows', label: 'Chưa phân loại', detail: 'Dòng chiết tính chưa xác định được nhóm hao phí.' },
];

const bidColumnFields: readonly ColumnField[] = [
  ['stt', 'STT'],
  ['code', 'Ma so'],
  ['name', 'Ten cong tac'],
  ['unit', 'Don vi'],
  ['qty', 'Khoi luong'],
  ['vl', 'VL'],
  ['nc', 'NC'],
  ['mtc', 'MTC'],
  ['unit_total', 'Don gia'],
  ['amount_total', 'Thanh tien'],
];

const detailColumnFields: readonly ColumnField[] = [
  ['stt', 'STT'],
  ['code', 'Ma so'],
  ['name', 'Thanh phan hao phi'],
  ['unit', 'Don vi'],
  ['norm', 'Dinh muc'],
  ['price', 'Don gia'],
  ['coef', 'He so'],
  ['amount', 'Thanh tien'],
  ['helper', 'Helper'],
];

function formatNumber(value: number | undefined | null) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '0';
  return value.toLocaleString('en-US', { maximumFractionDigits: 3 });
}

function formatBytes(value: number) {
  if (!Number.isFinite(value) || value <= 0) return '0 KB';
  if (value < 1024 * 1024) return `${Math.round(value / 1024).toLocaleString('en-US')} KB`;
  return `${(value / (1024 * 1024)).toLocaleString('en-US', { maximumFractionDigits: 2 })} MB`;
}

function workbookStem(name: string) {
  return (name || 'du_toan').replace(/\.[^.]+$/, '') || 'du_toan';
}

function isSheetIndex(value: number | null | undefined): value is number {
  return typeof value === 'number' && Number.isInteger(value);
}

function isPositiveInteger(value: number | null | undefined): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value > 0;
}

function suggestedHeaderRow(value: unknown) {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : null;
}

function cleanColumnConfig(value: unknown): EstimateColumnConfig {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return {};
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>).map(([key, cell]) => [
      key,
      String(cell || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, ''),
    ]),
  );
}

function selectedSheetName(sheets: EstimateSheetInfo[], value: number | null | undefined) {
  return sheets.find((sheet) => sheet.index === value)?.name || '';
}

function buildInitialSelection(summary: EstimateUploadSummary): EstimateSheetSelection {
  const sheets = summary.sheets || [];
  const first = sheets[0]?.index ?? null;
  const second = sheets.find((sheet) => sheet.index !== first)?.index ?? null;
  const suggested = summary.suggested_sheets || {};
  return {
    bid_sheet_index: isSheetIndex(suggested.bid_sheet_index) ? suggested.bid_sheet_index : first,
    detail_sheet_index: isSheetIndex(suggested.detail_sheet_index) ? suggested.detail_sheet_index : second,
    bid_header_row: suggestedHeaderRow(suggested.bid_header_row) || 4,
    detail_header_row: suggestedHeaderRow(suggested.detail_header_row) || 4,
    bid_columns: cleanColumnConfig(suggested.bid_columns),
    detail_columns: cleanColumnConfig(suggested.detail_columns),
  };
}

type DesktopSaveResult = { saved?: boolean; cancelled?: boolean; path?: string; error?: string };
type SaveFilePickerWritable = { write: (data: Blob) => Promise<void>; close: () => Promise<void> };
type SaveFilePickerHandle = { createWritable: () => Promise<SaveFilePickerWritable> };

declare global {
  interface Window {
    pywebview?: { api?: { save_file?: (filename: string, dataBase64: string) => Promise<DesktopSaveResult> } };
    showSaveFilePicker?: (options: { suggestedName: string; types?: Array<{ description: string; accept: Record<string, string[]> }> }) => Promise<SaveFilePickerHandle>;
  }
}

async function saveBlob(blob: Blob, filename: string) {
  const desktopSave = window.pywebview?.api?.save_file;
  if (desktopSave) {
    const result = await desktopSave(filename, await blobToBase64(blob));
    if (result.error) throw new Error(result.error);
    return result.saved === true;
  }
  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [{ description: 'Excel workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }],
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return true;
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return false;
      throw error;
    }
  }
  downloadBlob(blob, filename);
  return true;
}

function blobToBase64(blob: Blob) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result;
      if (typeof result === 'string') resolve(result.split(',')[1] || '');
      else reject(new Error('Không đọc được file để lưu.'));
    };
    reader.onerror = () => reject(reader.error || new Error('Không đọc được file để lưu.'));
    reader.readAsDataURL(blob);
  });
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function warningCount(analysis: EstimateAnalysis | null) {
  if (!analysis) return 0;
  return warningSections.reduce((total, section) => total + ((analysis.warnings[section.key] as unknown[] | undefined)?.length || 0), 0);
}

function compactValue(value: unknown): string {
  if (value === null || value === undefined || value === '') return '—';
  if (typeof value === 'number') return formatNumber(value);
  if (typeof value === 'boolean') return value ? 'Có' : 'Không';
  if (typeof value === 'string') return value;
  if (Array.isArray(value)) return value.map(compactValue).join(', ');
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function recordRows(rows: unknown[]) {
  const objects = rows.filter((row): row is Record<string, unknown> => Boolean(row && typeof row === 'object' && !Array.isArray(row)));
  const columns = Array.from(new Set(objects.flatMap((row) => Object.keys(row)))).slice(0, 8);
  return { objects, columns };
}

function WarningCard({ section, rows }: { section: WarningSection; rows: unknown[] }) {
  const { objects, columns } = recordRows(rows);
  if (!rows.length) return null;
  return (
    <article className="estimate-warning-card">
      <header>
        <div>
          <strong>{section.label}</strong>
          <span>{section.detail}</span>
        </div>
        <b>{rows.length.toLocaleString('en-US')}</b>
      </header>
      {objects.length && columns.length ? (
        <div className="estimate-warning-table-wrap">
          <table className="estimate-warning-table">
            <thead>
              <tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr>
            </thead>
            <tbody>
              {objects.slice(0, 20).map((row, index) => (
                <tr key={`${section.key}-${index}`}>
                  {columns.map((column) => <td key={column} title={compactValue(row[column])}>{compactValue(row[column])}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
          {objects.length > 20 && <p className="estimate-muted">Đang hiển thị 20 dòng đầu. File xuất vẫn dùng toàn bộ dữ liệu.</p>}
        </div>
      ) : (
        <pre className="estimate-json">{JSON.stringify(rows.slice(0, 20), null, 2)}</pre>
      )}
    </article>
  );
}

function SheetSelect({
  label,
  value,
  sheets,
  disabled,
  onChange,
}: {
  label: string;
  value: number | null | undefined;
  sheets: EstimateSheetInfo[];
  disabled: boolean;
  onChange: (value: number | null) => void;
}) {
  return (
    <label className="estimate-sheet-field">
      <span>{label}</span>
      <select disabled={disabled} value={isSheetIndex(value) ? String(value) : ''} onChange={(event) => onChange(event.currentTarget.value === '' ? null : Number(event.currentTarget.value))}>
        <option value="">Chọn sheet</option>
        {sheets.map((sheet) => (
          <option key={sheet.index} value={sheet.index}>
            {sheet.name} ({sheet.rows.toLocaleString('en-US')} dòng, {sheet.cols.toLocaleString('en-US')} cột)
          </option>
        ))}
      </select>
    </label>
  );
}

function ColumnConfigPanel({
  title,
  headerRow,
  columns,
  fields,
  disabled,
  onHeaderRowChange,
  onColumnChange,
}: {
  title: string;
  headerRow: number | null | undefined;
  columns: EstimateColumnConfig | undefined;
  fields: readonly ColumnField[];
  disabled: boolean;
  onHeaderRowChange: (value: string) => void;
  onColumnChange: (field: string, value: string) => void;
}) {
  return (
    <section className="estimate-column-card">
      <header>
        <div>
          <strong>{title}</strong>
          <span>Nhap dong header va chu cot Excel. De trong cot tuy chon neu file khong co.</span>
        </div>
        <label className="estimate-column-field compact">
          <span>Header row</span>
          <input
            type="number"
            min={1}
            disabled={disabled}
            value={headerRow ?? ''}
            onChange={(event) => onHeaderRowChange(event.currentTarget.value)}
          />
        </label>
      </header>
      <div className="estimate-column-grid">
        {fields.map(([field, label]) => (
          <label className="estimate-column-field" key={field}>
            <span>{label}</span>
            <input
              disabled={disabled}
              value={columns?.[field] || ''}
              placeholder="A"
              onChange={(event) => onColumnChange(field, event.currentTarget.value)}
            />
          </label>
        ))}
      </div>
    </section>
  );
}

export function EstimateExtractorWorkflow({ licenseReady, onStatus, stage, onStageChange }: Props) {
  const [summary, setSummary] = useState<EstimateUploadSummary | null>(null);
  const [analysis, setAnalysis] = useState<EstimateAnalysis | null>(null);
  const [sheetSelection, setSheetSelection] = useState<EstimateSheetSelection>({});
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');

  const warnings = useMemo(() => warningCount(analysis), [analysis]);
  const sheets = summary?.sheets || [];
  const headerReady = isPositiveInteger(sheetSelection.bid_header_row) && isPositiveInteger(sheetSelection.detail_header_row);
  const selectionReady = isSheetIndex(sheetSelection.bid_sheet_index) && isSheetIndex(sheetSelection.detail_sheet_index) && sheetSelection.bid_sheet_index !== sheetSelection.detail_sheet_index && headerReady;
  const showUpload = !summary || stage === 1;
  const showSelection = Boolean(summary && !showUpload && (stage === 2 || !analysis));
  const showAnalysis = Boolean(summary && analysis && stage === 3);
  const canExport = Boolean(summary?.saved_name && analysis && !busy && licenseReady && selectionReady);

  function updateSelection(key: keyof EstimateSheetSelection, value: number | null) {
    setSheetSelection((current) => ({ ...current, [key]: value }));
    setAnalysis(null);
  }

  function updateHeaderRow(key: 'bid_header_row' | 'detail_header_row', value: string) {
    const trimmed = value.trim();
    const nextValue = trimmed ? Number(trimmed) : null;
    setSheetSelection((current) => ({ ...current, [key]: Number.isFinite(nextValue) ? nextValue : null }));
    setAnalysis(null);
  }

  function updateColumnConfig(section: 'bid_columns' | 'detail_columns', field: string, value: string) {
    const cleaned = value.toUpperCase().replace(/[^A-Z0-9]/g, '');
    setSheetSelection((current) => ({
      ...current,
      [section]: { ...(current[section] || {}), [field]: cleaned },
    }));
    setAnalysis(null);
  }

  async function handleUpload(file: File | null) {
    if (!file) return;
    setBusy(true);
    setError('');
    setAnalysis(null);
    onStatus('Đang đọc danh sách sheet trong file dự toán...');
    try {
      const nextSummary = await uploadEstimateWorkbook(file);
      const nextSelection = buildInitialSelection(nextSummary);
      setSummary(nextSummary);
      setSheetSelection(nextSelection);
      void onStageChange(2);
      onStatus(`Đã đọc ${file.name}. Chọn sheet Dự thầu và Chiết tính trước khi phân tích.`);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setError(message);
      onStatus(message);
    } finally {
      setBusy(false);
    }
  }

  async function refreshAnalysis() {
    if (!summary) return;
    if (!selectionReady) {
      const message = 'Cần chọn 2 sheet khác nhau cho Dự thầu và Chiết tính.';
      setError(message);
      onStatus(message);
      return;
    }
    setBusy(true);
    setError('');
    onStatus('Đang phân tích Dự thầu và Chiết tính theo sheet đã chọn...');
    try {
      const nextAnalysis = await analyzeEstimateWorkbook(summary.saved_name, sheetSelection);
      setAnalysis(nextAnalysis);
      void onStageChange(3);
      const count = warningCount(nextAnalysis);
      onStatus(count ? `Đã phân tích. Có ${count.toLocaleString('en-US')} cảnh báo cần xem trước khi xuất.` : 'Đã phân tích. Không có cảnh báo lớn.');
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setError(message);
      onStatus(message);
    } finally {
      setBusy(false);
    }
  }

  async function exportWorkbook() {
    if (!summary) return;
    if (!selectionReady) {
      const message = 'Cần chọn sheet Dự thầu và Chiết tính trước khi xuất.';
      setError(message);
      onStatus(message);
      return;
    }
    setBusy(true);
    setError('');
    onStatus('Đang tạo file bóc tách dự toán...');
    try {
      const blob = await exportEstimateWorkbook(summary.saved_name, summary.original_name, sheetSelection);
      const saved = await saveBlob(blob, `${workbookStem(summary.original_name)}_boc_tach.xlsx`);
      onStatus(saved ? 'Đã lưu file bóc tách dự toán. Hãy kiểm tra file vừa lưu trước khi dùng chính thức.' : 'Đã hủy lưu file bóc tách dự toán.');
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setError(message);
      onStatus(message);
    } finally {
      setBusy(false);
    }
  }

  return (
    <section className="estimate-workflow">
      {showUpload ? (
        <div className="estimate-upload-card">
          <span className="step-pill">Hồ Gươm</span>
          <h2>Tải file dự toán</h2>
          <p>Tải file có sheet Dự thầu và Chiết tính. App sẽ chỉ đọc danh sách sheet trước, sau đó bạn chọn đúng sheet ở stage kế tiếp.</p>
          <label className={`estimate-drop-zone ${!licenseReady || busy ? 'disabled' : ''}`}>
            <input type="file" accept=".xls,.xlsx,.xlsm" disabled={!licenseReady || busy} onChange={(event) => void handleUpload(event.currentTarget.files?.[0] || null)} />
            <strong>{busy ? 'Đang đọc file...' : 'Chọn file dự toán'}</strong>
            <span>Hỗ trợ .xls, .xlsx, .xlsm. File xuất của workflow bóc tách dùng định dạng .xlsx mới.</span>
          </label>
          {!licenseReady && <p className="estimate-error">Cần license cho profile Hồ Gươm trước khi tải file.</p>}
          {error && <p className="estimate-error">{error}</p>}
        </div>
      ) : (
        <div className="estimate-preview-layout">
          <section className="estimate-file-card">
            <div>
              <span className="section-eyebrow">File dự toán</span>
              <h3>{summary?.original_name}</h3>
              <p>{summary ? `${formatBytes(summary.size)} · Cache: ${summary.saved_name}` : ''}</p>
              <p>
                Dự thầu: <b>{selectedSheetName(sheets, sheetSelection.bid_sheet_index) || 'chưa chọn'}</b>
                {' · '}
                Chiết tính: <b>{selectedSheetName(sheets, sheetSelection.detail_sheet_index) || 'chưa chọn'}</b>
              </p>
            </div>
            <div className="estimate-actions">
              <label className="btn-secondary estimate-reupload">
                Đổi file
                <input type="file" accept=".xls,.xlsx,.xlsm" disabled={busy || !licenseReady} onChange={(event) => void handleUpload(event.currentTarget.files?.[0] || null)} />
              </label>
              {stage !== 2 && <button type="button" className="btn-secondary" disabled={busy || !licenseReady} onClick={() => void onStageChange(2)}>Chọn lại sheet</button>}
              {stage === 3 && <button type="button" className="btn-secondary" disabled={busy || !licenseReady} onClick={() => void refreshAnalysis()}>Phân tích lại</button>}
              {stage === 3 && <button type="button" disabled={!canExport} onClick={() => void exportWorkbook()}>{busy ? 'Đang xử lý...' : 'Xuất file bóc tách'}</button>}
            </div>
          </section>

          {error && <p className="estimate-error">{error}</p>}

          {showSelection && (
            <section className="estimate-sheet-panel">
              <header>
                <div>
                  <span className="section-eyebrow">Stage 2</span>
                  <h3>Chọn sheet để bóc tách</h3>
                  <p>Chọn sheet nguồn cho bảng Dự thầu và sheet giải thích Chiết tính. THVT sẽ được tạo ở file xuất, không cần upload sẵn.</p>
                </div>
              </header>
              <div className="estimate-sheet-grid">
                <SheetSelect label="Sheet Dự thầu" value={sheetSelection.bid_sheet_index} sheets={sheets} disabled={busy} onChange={(value) => updateSelection('bid_sheet_index', value)} />
                <SheetSelect label="Sheet Chiết tính" value={sheetSelection.detail_sheet_index} sheets={sheets} disabled={busy} onChange={(value) => updateSelection('detail_sheet_index', value)} />
              </div>
              <div className="estimate-column-layout">
                <ColumnConfigPanel
                  title="Cot Du thau"
                  headerRow={sheetSelection.bid_header_row}
                  columns={sheetSelection.bid_columns}
                  fields={bidColumnFields}
                  disabled={busy}
                  onHeaderRowChange={(value) => updateHeaderRow('bid_header_row', value)}
                  onColumnChange={(field, value) => updateColumnConfig('bid_columns', field, value)}
                />
                <ColumnConfigPanel
                  title="Cot Chiet tinh"
                  headerRow={sheetSelection.detail_header_row}
                  columns={sheetSelection.detail_columns}
                  fields={detailColumnFields}
                  disabled={busy}
                  onHeaderRowChange={(value) => updateHeaderRow('detail_header_row', value)}
                  onColumnChange={(field, value) => updateColumnConfig('detail_columns', field, value)}
                />
              </div>
              <div className="estimate-sheet-meta">
                <strong>{sheets.length.toLocaleString('en-US')} sheet trong file</strong>
                <span>{selectionReady ? 'Đã đủ lựa chọn để phân tích.' : 'Cần chọn 2 sheet khác nhau.'}</span>
              </div>
              <div className="estimate-actions">
                <button type="button" disabled={!selectionReady || busy || !licenseReady} onClick={() => void refreshAnalysis()}>{busy ? 'Đang phân tích...' : 'Phân tích file'}</button>
              </div>
            </section>
          )}

          {showAnalysis && analysis && (
            <>
              <section className="estimate-grid" aria-label="Tổng quan bóc tách">
                <div className="estimate-kpi"><span>Dự thầu</span><strong>{formatNumber(analysis.summary.bid_rows)}</strong><small>dòng công tác</small></div>
                <div className="estimate-kpi"><span>Chiết tính</span><strong>{formatNumber(analysis.summary.detail_blocks)}</strong><small>block đã ghép</small></div>
                <div className="estimate-kpi"><span>THVT mẫu</span><strong>{formatNumber(analysis.summary.thvt_rows)}</strong><small>dòng hiện có</small></div>
                <div className="estimate-kpi"><span>THVT tạo mới</span><strong>{formatNumber(analysis.summary.generated_thvt_rows)}</strong><small>dòng từ chiết tính</small></div>
                <div className={`estimate-kpi ${warnings ? 'warning' : 'ok'}`}><span>Cảnh báo</span><strong>{formatNumber(warnings)}</strong><small>{warnings ? 'cần rà soát' : 'ổn'}</small></div>
                <div className={`estimate-kpi ${analysis.summary.ok ? 'ok' : 'warning'}`}><span>Trạng thái</span><strong>{analysis.summary.ok ? 'OK' : 'Cần xem'}</strong><small>trước khi xuất</small></div>
              </section>

              <section className="estimate-preview-panel">
                <header>
                  <div>
                    <span className="section-eyebrow">Preview kiểm tra</span>
                    <h3>Đối chiếu dữ liệu bóc tách</h3>
                  </div>
                  <span className={warnings ? 'warning-text' : 'ok-text'}>{warnings ? `${warnings.toLocaleString('en-US')} cảnh báo` : 'Không có cảnh báo lớn'}</span>
                </header>

                {!warnings ? (
                  <div className="estimate-ok-card">
                    <strong>Sẵn sàng xuất file</strong>
                    <span>Dự thầu và Chiết tính đang khớp theo logic hiện tại. File xuất sẽ tạo thêm THVT và Tổng hợp THVT.</span>
                  </div>
                ) : (
                  <div className="estimate-warning-grid">
                    {warningSections.map((section) => <WarningCard key={section.key} section={section} rows={(analysis.warnings[section.key] as unknown[] | undefined) || []} />)}
                  </div>
                )}
              </section>
            </>
          )}
        </div>
      )}
    </section>
  );
}
