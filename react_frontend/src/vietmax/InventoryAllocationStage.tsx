import type { InventoryAllocationConfig, InventoryAllocationJob, InventoryAllocationMappingSection, InventoryAllocationResult, UploadSummary } from '../types';

type InventoryAllocationStageProps = {
  purchaseFile: UploadSummary | null;
  salesFile: UploadSummary | null;
  processedPurchaseSavedName: string;
  processedSalesSavedName: string;
  openingStockFile: File | null;
  config: InventoryAllocationConfig;
  job: InventoryAllocationJob | null;
  result: InventoryAllocationResult | null;
  busy: boolean;
  onOpeningStockFileChange: (file: File | null) => void;
  onConfigChange: (config: InventoryAllocationConfig) => void;
  onProcess: () => void;
  onDownload: () => void;
};

const mappingSections: Array<{ key: 'purchase' | 'sales' | 'opening'; label: string; showInvoice: boolean }> = [
  { key: 'purchase', label: 'File mua vào đã xử lý', showInvoice: true },
  { key: 'sales', label: 'File bán ra đã xử lý', showInvoice: true },
  { key: 'opening', label: 'Tồn đầu kỳ (tùy chọn)', showInvoice: false },
];

export function InventoryAllocationStage({ purchaseFile, salesFile, processedPurchaseSavedName, processedSalesSavedName, openingStockFile, config, job, result, busy, onOpeningStockFileChange, onConfigChange, onProcess, onDownload }: InventoryAllocationStageProps) {
  const ready = Boolean(processedPurchaseSavedName && processedSalesSavedName);
  const progress = Math.max(0, Math.min(100, Number(job?.progress ?? (result ? 100 : 0))));
  const activeResult = result ?? job?.result ?? null;

  return (
    <div className="inventory-allocation-stage">
      <section className="inventory-allocation-card inventory-input-card">
        <div className="stage-toolbar">
          <div>
            <h3>Stage 12: Phân bổ tồn kho</h3>
            <p>Tạo báo cáo phân bổ từ file mua vào và bán ra đã xử lý ở các stage trước.</p>
          </div>
          <button type="button" disabled={busy || !ready} onClick={onProcess}>Chạy phân bổ tồn kho</button>
        </div>

        <div className="inventory-file-grid">
          <FileStatus label="Mua vào đã xử lý" originalName={purchaseFile?.original_name} ready={Boolean(processedPurchaseSavedName)} />
          <FileStatus label="Bán ra đã xử lý" originalName={salesFile?.original_name} ready={Boolean(processedSalesSavedName)} />
          <label className={`inventory-opening-upload ${openingStockFile ? 'has-file' : ''}`}>
            <input type="file" accept=".xlsx,.xlsm" disabled={busy} onChange={(event) => onOpeningStockFileChange(event.currentTarget.files?.[0] ?? null)} />
            <span>Tồn đầu kỳ</span>
            <strong>{openingStockFile?.name || 'Không bắt buộc'}</strong>
            <small>Upload nếu có tồn đầu kỳ cần cộng vào phân bổ.</small>
          </label>
        </div>

        {!ready && <p className="warning-text">Cần xuất cả file mua vào và file bán ra trước khi chạy Stage 12.</p>}
      </section>

      <section className="inventory-allocation-card inventory-config-card">
        <div className="stage-toolbar"><h3>Cấu hình phân bổ</h3><p>Giữ mặc định Vietmax nếu file xuất không đổi cột.</p></div>
        <div className="inventory-policy-grid">
          <NullableNumber label="Lỗ tối đa (%)" value={config.policy.max_loss_percent} onChange={(value) => onConfigChange({ ...config, policy: { ...config.policy, max_loss_percent: value } })} />
          <NullableNumber label="Lãi tối đa (%)" value={config.policy.max_profit_percent} onChange={(value) => onConfigChange({ ...config, policy: { ...config.policy, max_profit_percent: value } })} />
          <label><span>Profile phân bổ</span><select value={config.policy.company_profile} disabled={busy} onChange={(event) => onConfigChange({ ...config, policy: { ...config.policy, company_profile: event.currentTarget.value } })}><option value="yen_thanh">Yến Thành</option><option value="son_phuong">Sơn Phương</option></select></label>
          <label><span>Cửa sổ mua sau ngày bán</span><input type="number" min={1} value={config.policy.future_purchase_window_days} disabled={busy || !config.policy.allow_future_purchase_reorder} onChange={(event) => onConfigChange({ ...config, policy: { ...config.policy, future_purchase_window_days: parseInt(event.currentTarget.value) || 31 } })} /></label>
          <label className="inline-check"><input type="checkbox" checked={config.policy.ignore_sale_suffix} disabled={busy} onChange={(event) => onConfigChange({ ...config, policy: { ...config.policy, ignore_sale_suffix: event.currentTarget.checked } })} /> Bỏ hậu tố mã bán ra</label>
          <label className="inline-check"><input type="checkbox" checked={config.policy.allow_negative_export} disabled={busy} onChange={(event) => onConfigChange({ ...config, policy: { ...config.policy, allow_negative_export: event.currentTarget.checked } })} /> Cho phép xuất âm KTP</label>
          <label className="inline-check"><input type="checkbox" checked={config.policy.allow_future_purchase_reorder} disabled={busy} onChange={(event) => onConfigChange({ ...config, policy: { ...config.policy, allow_future_purchase_reorder: event.currentTarget.checked } })} /> Đưa mua sau lên trước</label>
        </div>
        <div className="inventory-mapping-grid">
          {mappingSections.map((section) => <MappingEditor key={section.key} title={section.label} mapping={config.mapping[section.key]} showInvoice={section.showInvoice} busy={busy} onChange={(mapping) => onConfigChange({ ...config, mapping: { ...config.mapping, [section.key]: mapping } })} />)}
        </div>
      </section>

      <section className="inventory-allocation-card inventory-result-card">
        <div className="stage-toolbar">
          <h3>Tiến trình và kết quả</h3>
          <button type="button" className="btn-secondary" disabled={busy || !activeResult?.job_id} onClick={onDownload}>Tải báo cáo</button>
        </div>
        <div className="allocation-progress" aria-label="Tiến trình phân bổ"><span style={{ width: `${progress}%` }} /></div>
        <p className="muted">{job?.label || (activeResult ? 'Đã hoàn tất phân bổ tồn kho.' : 'Chưa chạy phân bổ.')}</p>
        {job?.error && <p className="warning-text">{job.error}</p>}
        {activeResult && <AllocationResult result={activeResult} />}
      </section>
    </div>
  );
}

function FileStatus({ label, originalName, ready }: { label: string; originalName?: string; ready: boolean }) {
  return <div className={`inventory-file-status ${ready ? 'ready' : ''}`}><span>{label}</span><strong>{originalName || 'Chưa có file'}</strong><small>{ready ? 'Đã lưu cache để phân bổ.' : 'Hãy xuất file ở stage trước.'}</small></div>;
}

function NullableNumber({ label, value, onChange }: { label: string; value: number | null; onChange: (value: number | null) => void }) {
  return <label><span>{label}</span><input type="number" value={value ?? ''} placeholder="Không giới hạn" onChange={(event) => onChange(event.currentTarget.value.trim() ? Number(event.currentTarget.value) : null)} /></label>;
}

function MappingEditor({ title, mapping, showInvoice, busy, onChange }: { title: string; mapping: InventoryAllocationMappingSection; showInvoice: boolean; busy: boolean; onChange: (mapping: InventoryAllocationMappingSection) => void }) {
  const update = (field: keyof InventoryAllocationMappingSection, value: string | number) => onChange({ ...mapping, [field]: value });
  return (
    <section className="inventory-mapping-card">
      <h4>{title}</h4>
      <div className="inventory-mapping-fields">
        <label><span>Sheet</span><input value={mapping.sheet} disabled={busy} placeholder="Sheet đầu tiên" onChange={(event) => update('sheet', event.currentTarget.value)} /></label>
        <label><span>Header</span><input type="number" min={1} value={mapping.header_row} disabled={busy} onChange={(event) => update('header_row', parseInt(event.currentTarget.value) || 1)} /></label>
        <label><span>Dòng dữ liệu</span><input type="number" min={1} value={mapping.data_start_row} disabled={busy} onChange={(event) => update('data_start_row', parseInt(event.currentTarget.value) || 1)} /></label>
        {showInvoice && <label><span>Số HĐ</span><input value={mapping.invoice_col} disabled={busy} onChange={(event) => update('invoice_col', event.currentTarget.value.toUpperCase())} /></label>}
        {showInvoice && <label><span>Ngày HĐ</span><input value={mapping.date_col} disabled={busy} onChange={(event) => update('date_col', event.currentTarget.value.toUpperCase())} /></label>}
        <label><span>Mã VT</span><input value={mapping.code_col} disabled={busy} onChange={(event) => update('code_col', event.currentTarget.value.toUpperCase())} /></label>
        <label><span>Tên hàng</span><input value={mapping.product_col} disabled={busy} onChange={(event) => update('product_col', event.currentTarget.value.toUpperCase())} /></label>
        <label><span>Số lượng</span><input value={mapping.qty_col} disabled={busy} onChange={(event) => update('qty_col', event.currentTarget.value.toUpperCase())} /></label>
        <label><span>Đơn giá</span><input value={mapping.price_col} disabled={busy} onChange={(event) => update('price_col', event.currentTarget.value.toUpperCase())} /></label>
      </div>
    </section>
  );
}

function AllocationResult({ result }: { result: InventoryAllocationResult }) {
  const summary = result.summary ?? {};
  const summaryItems = [
    ['Dòng phân bổ', result.allocation_count],
    ['Mã tồn kho', result.stock_count],
    ['SL tồn đầu kỳ', summary.opening_quantity],
    ['SL mua vào', summary.purchase_quantity],
    ['SL bán ra', summary.sales_quantity],
    ['SL lấy KHH', summary.material_quantity],
    ['SL lấy KTP', summary.finished_quantity],
    ['Mã chỉ bán ra', summary.sale_only_code_count],
  ];
  return (
    <div className="allocation-result-grid">
      <div className="allocation-summary-list">{summaryItems.map(([label, value]) => <div key={label}><span>{label}</span><strong>{formatNumber(value)}</strong></div>)}</div>
      <div className="allocation-verification">
        <h4>Đối chiếu</h4>
        {(result.verification ?? []).slice(0, 8).map((row, index) => <p key={`${row.group}-${row.check}-${index}`} className={row.status === 'OK' ? 'ok-text' : 'warning-text'}>{row.group}: {row.check} - {row.status}</p>)}
        {Boolean(result.warnings?.length) && <p className="warning-text">{result.warnings?.length} cảnh báo. Mở file báo cáo để xem chi tiết.</p>}
      </div>
    </div>
  );
}

function formatNumber(value: unknown) {
  if (typeof value === 'number') return value.toLocaleString('vi-VN');
  if (typeof value === 'string') return value;
  return '';
}
