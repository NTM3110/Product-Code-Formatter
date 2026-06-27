import { useEffect, useMemo, useRef, useState } from 'react';
import type { InventoryAllocationConfig, InventoryAllocationMappingSection, InventoryAllocationReportView, InventoryAllocationResult, InventoryLedgerDetailRow, InventorySalesDetailRow, InventorySalesSummaryRow, InventorySummaryRow, ProcessedFileStats, UploadSummary } from '../types';

type InventoryAllocationStageProps = {
  purchaseFile: UploadSummary | null;
  salesFile: UploadSummary | null;
  processedPurchaseSavedName: string;
  processedSalesSavedName: string;
  processedPurchaseStats: ProcessedFileStats | null;
  processedSalesStats: ProcessedFileStats | null;
  openingStockFile: File | null;
  config: InventoryAllocationConfig;
  busy: boolean;
  onProcessedPurchaseFileChange: (file: File | undefined) => void;
  onProcessedSalesFileChange: (file: File | undefined) => void;
  onOpeningStockFileChange: (file: File | null) => void;
  onConfigChange: (config: InventoryAllocationConfig) => void;
};

const mappingSections: Array<{ key: 'purchase' | 'sales' | 'opening'; label: string; showInvoice: boolean }> = [
  { key: 'purchase', label: 'File mua vào đã xử lý', showInvoice: true },
  { key: 'sales', label: 'File bán ra đã xử lý', showInvoice: true },
  { key: 'opening', label: 'Tồn đầu kỳ (tùy chọn)', showInvoice: false },
];
const REPORT_RENDER_BATCH_SIZE = 140;
const REPORT_LEDGER_SECTION_BATCH_SIZE = 12;
const REPORT_LEDGER_ROW_BATCH_SIZE = 80;
const REPORT_RENDER_CACHE_LIMIT = 80;

function rememberRenderedCount(cacheRef: { current: Map<string, number> } | undefined, cacheKey: string, count: number) {
  if (!cacheRef || !cacheKey) return;
  cacheRef.current.delete(cacheKey);
  cacheRef.current.set(cacheKey, count);
  while (cacheRef.current.size > REPORT_RENDER_CACHE_LIMIT) {
    const oldestKey = cacheRef.current.keys().next().value;
    if (!oldestKey) break;
    cacheRef.current.delete(oldestKey);
  }
}

function useProgressiveRows<T>(rows: T[], active = true, batchSize = REPORT_RENDER_BATCH_SIZE, cacheKey = '', cacheRef?: { current: Map<string, number> }) {
  const cachedCount = cacheKey && cacheRef ? cacheRef.current.get(cacheKey) : undefined;
  const initialCount = active ? Math.min(rows.length, Math.max(batchSize, cachedCount ?? 0)) : 0;
  const [visibleCount, setVisibleCount] = useState(initialCount);

  useEffect(() => {
    if (!active) {
      setVisibleCount(0);
      return;
    }
    const savedCount = cacheKey && cacheRef ? cacheRef.current.get(cacheKey) : undefined;
    const firstBatch = Math.min(rows.length, Math.max(batchSize, savedCount ?? 0));
    setVisibleCount(firstBatch);
    rememberRenderedCount(cacheRef, cacheKey, firstBatch);
    if (rows.length <= firstBatch) return;

    let cancelled = false;
    let timer: number | null = null;
    const showNextBatch = () => {
      if (cancelled) return;
      setVisibleCount((current) => {
        const next = Math.min(rows.length, current + batchSize);
        rememberRenderedCount(cacheRef, cacheKey, next);
        if (next < rows.length) timer = window.setTimeout(showNextBatch, 16);
        return next;
      });
    };
    timer = window.setTimeout(showNextBatch, 16);
    return () => {
      cancelled = true;
      if (timer !== null) window.clearTimeout(timer);
    };
  }, [active, batchSize, rows, cacheKey, cacheRef]);

  const count = active ? Math.min(visibleCount, rows.length) : 0;
  return {
    visibleRows: active && count < rows.length ? rows.slice(0, count) : rows,
    visibleCount: count,
    isRendering: active && count < rows.length,
  };
}

export function InventoryAllocationStage({ purchaseFile, salesFile, processedPurchaseSavedName, processedSalesSavedName, processedPurchaseStats, processedSalesStats, openingStockFile, config, busy, onProcessedPurchaseFileChange, onProcessedSalesFileChange, onOpeningStockFileChange, onConfigChange }: InventoryAllocationStageProps) {
  const ready = Boolean(processedPurchaseSavedName && processedSalesSavedName);
  const visibleMappingSections = openingStockFile ? mappingSections : mappingSections.filter((section) => section.key !== 'opening');

  return (
    <div className="inventory-allocation-stage">
      <section className="inventory-allocation-card inventory-input-card">
        <div className="stage-toolbar">
          <div>
            <h3>Stage 12: Phân bổ tồn kho</h3>
            <p>Kiểm tra file đã xử lý và cấu hình trước khi bấm Tiếp tục để chạy phân bổ.</p>
          </div>
        </div>

        <div className="inventory-file-grid">
          <ProcessedFileInput label="Mua vào đã xử lý" originalName={purchaseFile?.original_name} ready={Boolean(processedPurchaseSavedName)} stats={processedPurchaseStats} busy={busy} onFileChange={onProcessedPurchaseFileChange} />
          <ProcessedFileInput label="Bán ra đã xử lý" originalName={salesFile?.original_name} ready={Boolean(processedSalesSavedName)} stats={processedSalesStats} busy={busy} onFileChange={onProcessedSalesFileChange} />
          <label className={`inventory-opening-upload ${openingStockFile ? 'has-file' : ''}`}>
            <input type="file" accept=".xls,.xlsx,.xlsm" disabled={busy} onChange={(event) => onOpeningStockFileChange(event.currentTarget.files?.[0] ?? null)} />
            <span>Tồn đầu kỳ</span>
            <strong>{openingStockFile?.name || 'Bỏ qua nếu không có'}</strong>
            <small>Chỉ upload khi cần cộng tồn đầu kỳ vào phân bổ.</small>
          </label>
        </div>

        {!ready && <p className="warning-text">Cần tải hoặc tạo cả file mua vào và file bán ra đã xử lý trước khi chạy Stage 12.</p>}
      </section>

      <section className="inventory-allocation-card inventory-config-card">
        <div className="stage-toolbar"><h3>Cấu hình phân bổ</h3><p>Giữ mặc định Vietmax nếu file xuất không đổi cột.</p></div>
        <div className="inventory-policy-grid">
          <NullableNumber label="Lỗ tối đa (%)" value={config.policy.max_loss_percent} onChange={(value) => onConfigChange({ ...config, policy: { ...config.policy, max_loss_percent: value } })} />
          <NullableNumber label="Lãi tối đa (%)" value={config.policy.max_profit_percent} onChange={(value) => onConfigChange({ ...config, policy: { ...config.policy, max_profit_percent: value } })} />
          <label><span>Cửa sổ mua sau ngày bán</span><input type="number" min={1} value={config.policy.future_purchase_window_days} disabled={busy || !config.policy.allow_future_purchase_reorder} onChange={(event) => onConfigChange({ ...config, policy: { ...config.policy, future_purchase_window_days: parseInt(event.currentTarget.value) || 31 } })} /></label>
          <label className="inline-check"><input type="checkbox" checked={config.policy.ignore_sale_suffix} disabled={busy} onChange={(event) => onConfigChange({ ...config, policy: { ...config.policy, ignore_sale_suffix: event.currentTarget.checked } })} /> Bỏ hậu tố mã bán ra</label>
          <label className="inline-check"><input type="checkbox" checked={config.policy.allow_negative_export} disabled={busy} onChange={(event) => onConfigChange({ ...config, policy: { ...config.policy, allow_negative_export: event.currentTarget.checked } })} /> Cho phép xuất âm theo kho bán ra</label>
          <label className="inline-check"><input type="checkbox" checked={config.policy.allow_future_purchase_reorder} disabled={busy} onChange={(event) => onConfigChange({ ...config, policy: { ...config.policy, allow_future_purchase_reorder: event.currentTarget.checked } })} /> Đưa mua sau lên trước</label>
        </div>
        <div className="inventory-mapping-grid">
          {visibleMappingSections.map((section) => <MappingEditor key={section.key} title={section.label} mapping={config.mapping[section.key]} showInvoice={section.showInvoice} busy={busy} onChange={(mapping) => onConfigChange({ ...config, mapping: { ...config.mapping, [section.key]: mapping } })} />)}
        </div>
      </section>
    </div>
  );
}

function ProcessedFileInput({ label, originalName, ready, stats, busy, onFileChange }: { label: string; originalName?: string; ready: boolean; stats: ProcessedFileStats | null; busy: boolean; onFileChange: (file: File | undefined) => void }) {
  return (
    <label className={`inventory-file-status ${ready ? 'ready' : ''}`}>
      <input type="file" accept=".xls,.xlsx,.xlsm" disabled={busy} onChange={(event) => onFileChange(event.currentTarget.files?.[0])} />
      <span>{label}</span>
      <strong>{ready ? (originalName || 'Đã có cache') : 'Chưa có file đã xử lý'}</strong>
      <small>{ready ? 'Đã sẵn sàng để phân bổ.' : 'Upload file đã xử lý nếu bỏ qua stage trước.'}</small>
      <ProcessedStats stats={stats} />
    </label>
  );
}

function ProcessedStats({ stats }: { stats: ProcessedFileStats | null }) {
  if (!stats) return <small className="muted">Chưa có thống kê.</small>;
  return <small className="processed-stats inline-stats">Công ty <strong>{formatNumber(stats.processed_company_count)} / {formatNumber(stats.company_count)}</strong> · Dòng hàng <strong>{formatNumber(stats.processed_product_row_count)} / {formatNumber(stats.product_row_count)}</strong></small>;
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
    ['SL khớp mua vào', summary.material_quantity],
    ['SL còn lại theo kho bán ra', summary.finished_quantity],
    ['Mã chỉ bán ra', summary.sale_only_code_count],
  ];
  return (
    <div className="allocation-result-grid">
      <div className="allocation-summary-list">{summaryItems.map(([label, value]) => <div key={label}><span>{label}</span><strong>{formatNumber(value)}</strong></div>)}</div>
      <div className="allocation-verification">
        <h4>Đối chiếu</h4>
        {(result.verification ?? []).slice(0, 8).map((row, index) => <p key={`${row.group}-${row.check}-${index}`} className={row.status === 'OK' ? 'ok-text' : 'warning-text'}>{row.group}: {row.check} - {row.status}</p>)}
        {Boolean(result.warnings?.length) && <p className="warning-text">{result.warnings?.length} cảnh báo. Qua Stage 13 để xem chi tiết.</p>}
      </div>
    </div>
  );
}

export function InventoryAllocationReportStage({ result }: { result: InventoryAllocationResult | null; busy: boolean }) {
  if (!result?.report_view) {
    return (
      <div className="inventory-allocation-stage report-only-stage">
        <section className="inventory-allocation-card empty-report-stage">
          <div>
            <h3>Stage 13: Xem báo cáo tồn kho</h3>
            <p className="muted">Chưa có dữ liệu báo cáo. Quay lại Stage 12 và bấm Tiếp tục để chạy phân bổ trước.</p>
          </div>
        </section>
      </div>
    );
  }
  return (
    <div className="inventory-allocation-stage report-only-stage">
      <section className="inventory-allocation-card inventory-result-card">
        <AllocationReportViewer reportView={result.report_view} />
      </section>
    </div>
  );
}

export function InventoryAllocationExportStage({ result, busy, onDownload }: { result: InventoryAllocationResult | null; busy: boolean; onDownload: () => void }) {
  return (
    <div className="inventory-allocation-stage export-only-stage">
      <section className="inventory-allocation-card inventory-export-card">
        <div className="report-stage-heading">
          <span>Xuất file</span>
          <h3>Xuất báo cáo phân bổ tồn kho</h3>
          <p>File xuất sẽ giữ toàn bộ sheet báo cáo, sổ chi tiết, bảng kê hóa đơn và kiểm tra đối chiếu.</p>
        </div>
        {result && <AllocationResult result={result} />}
        {!result && <p className="warning-text">Chưa có kết quả phân bổ. Quay lại Stage 12 để chạy phân bổ trước.</p>}
        <div className="export-action-panel">
          <button type="button" disabled={busy || !result?.job_id} onClick={onDownload}>Xuất file phân bổ tồn kho</button>
          {result?.filename && <small className="muted">File: {result.filename}</small>}
        </div>
      </section>
    </div>
  );
}

function AllocationReportViewer({ reportView }: { reportView: InventoryAllocationReportView }) {
  const [tab, setTab] = useState<'sales' | 'ledger' | 'inventory'>('sales');
  const [warehouseFilter, setWarehouseFilter] = useState('');
  const [selectedSalesKey, setSelectedSalesKey] = useState('');
  const [selectedInventoryKey, setSelectedInventoryKey] = useState('');
  const initialFromDate = reportView.date_range?.from ?? '';
  const initialToDate = reportView.date_range?.to ?? '';
  const [draftFromDate, setDraftFromDate] = useState(initialFromDate);
  const [draftToDate, setDraftToDate] = useState(initialToDate);
  const [fromDate, setFromDate] = useState(initialFromDate);
  const [toDate, setToDate] = useState(initialToDate);
  useEffect(() => {
    setDraftFromDate(initialFromDate);
    setDraftToDate(initialToDate);
    setFromDate(initialFromDate);
    setToDate(initialToDate);
    setSelectedSalesKey('');
    setSelectedInventoryKey('');
  }, [initialFromDate, initialToDate]);
  const [isReportSwitching, setIsReportSwitching] = useState(false);
  const reportSwitchTimer = useRef<number | null>(null);
  const reportSettleTimer = useRef<number | null>(null);
  const reportRenderCache = useRef(new Map<string, number>());
  useEffect(() => {
    reportRenderCache.current.clear();
  }, [reportView]);
  useEffect(() => () => {
    if (reportSwitchTimer.current !== null) window.clearTimeout(reportSwitchTimer.current);
    if (reportSettleTimer.current !== null) window.clearTimeout(reportSettleTimer.current);
  }, []);

  const warehouses = reportView.warehouses ?? [];
  const salesDetailSource = reportView.sales_detail_rows ?? [];
  const ledgerDetailSource = reportView.ledger_detail_rows ?? [];
  const needsLedgerRows = tab === 'ledger' || tab === 'inventory' || Boolean(selectedSalesKey) || Boolean(selectedInventoryKey);
  const warehouseSalesDetails = useMemo(() => filterByWarehouse(salesDetailSource, warehouseFilter), [salesDetailSource, warehouseFilter]);
  const warehouseLedgerDetails = useMemo(() => needsLedgerRows ? filterByWarehouse(ledgerDetailSource, warehouseFilter) : [], [needsLedgerRows, ledgerDetailSource, warehouseFilter]);
  const filteredSalesDetails = useMemo(() => filterSalesDetailsByDate(warehouseSalesDetails, fromDate, toDate), [warehouseSalesDetails, fromDate, toDate]);
  const filteredLedgerDetails = useMemo(() => filterLedgerDetailsByDate(warehouseLedgerDetails, fromDate, toDate), [warehouseLedgerDetails, fromDate, toDate]);
  const salesSummaryRows = useMemo(() => buildSalesSummaryRows(filteredSalesDetails, salesDetailSource.length ? [] : filterByWarehouse(reportView.sales_summary_rows ?? [], warehouseFilter)), [filteredSalesDetails, salesDetailSource.length, reportView.sales_summary_rows, warehouseFilter]);
  const inventorySummaryRows = useMemo(() => tab === 'inventory' ? buildInventorySummaryRows(warehouseLedgerDetails, ledgerDetailSource.length ? [] : filterByWarehouse(reportView.inventory_summary_rows ?? [], warehouseFilter), fromDate, toDate) : [], [tab, warehouseLedgerDetails, ledgerDetailSource.length, reportView.inventory_summary_rows, warehouseFilter, fromDate, toDate]);
  const ledgerSections = useMemo(() => tab === 'ledger' ? buildLedgerSections(warehouseLedgerDetails, warehouses, fromDate, toDate) : [], [tab, warehouseLedgerDetails, warehouses, fromDate, toDate]);
  const inventoryDetailRows = useMemo(() => tab === 'inventory' && selectedInventoryKey ? filteredLedgerDetails.filter((row) => row.summary_key === selectedInventoryKey) : [], [tab, filteredLedgerDetails, selectedInventoryKey]);
  const reportTotals = useMemo(() => reportKpis(salesSummaryRows, inventorySummaryRows), [salesSummaryRows, inventorySummaryRows]);
  const selectedWarehouse = warehouses.find((warehouse) => warehouse.warehouse_code === warehouseFilter);
  const selectedSalesRow = useMemo(() => salesSummaryRows.find((row) => row.key === selectedSalesKey) ?? null, [salesSummaryRows, selectedSalesKey]);
  const selectedSalesDetails = useMemo(() => filteredSalesDetails.filter((row) => row.summary_key === selectedSalesKey), [filteredSalesDetails, selectedSalesKey]);
  const selectedLedgerRows = useMemo(() => {
    if (!selectedSalesRow) return [];
    const ledgerKey = inventoryReportSummaryKey(selectedSalesRow.warehouse_code, selectedSalesRow.variant_code);
    return filteredLedgerDetails.filter((row) => row.summary_key === ledgerKey);
  }, [filteredLedgerDetails, selectedSalesRow]);
  const ledgerRenderKey = ['ledger', fromDate || 'all-from', toDate || 'all-to', warehouseFilter || 'all-warehouses'].join('|');

  function runReportSwitch(action: () => void) {
    if (isReportSwitching) return;
    if (reportSwitchTimer.current !== null) window.clearTimeout(reportSwitchTimer.current);
    if (reportSettleTimer.current !== null) window.clearTimeout(reportSettleTimer.current);
    setIsReportSwitching(true);
    reportSwitchTimer.current = window.setTimeout(() => {
      action();
      reportSwitchTimer.current = null;
      reportSettleTimer.current = window.setTimeout(() => {
        setIsReportSwitching(false);
        reportSettleTimer.current = null;
      }, 140);
    }, 90);
  }

  function changeTab(nextTab: 'sales' | 'ledger' | 'inventory') {
    if (nextTab === tab) return;
    runReportSwitch(() => {
      setTab(nextTab);
      setSelectedSalesKey('');
      setSelectedInventoryKey('');
    });
  }


  function applyDateFilter() {
    const nextFrom = draftFromDate;
    const nextTo = draftToDate;
    runReportSwitch(() => {
      if (nextFrom && nextTo && nextFrom > nextTo) {
        setFromDate(nextTo);
        setToDate(nextFrom);
        setDraftFromDate(nextTo);
        setDraftToDate(nextFrom);
      } else {
        setFromDate(nextFrom);
        setToDate(nextTo);
      }
      setSelectedSalesKey('');
      setSelectedInventoryKey('');
    });
  }

  function changeWarehouse(value: string) {
    if (value === warehouseFilter) return;
    runReportSwitch(() => {
      setWarehouseFilter(value);
      setSelectedSalesKey('');
      setSelectedInventoryKey('');
    });
  }

  return (
    <div className="allocation-report-viewer">
      <div className="report-filter-band">
        <div>
          <strong>{selectedWarehouse ? `KHO: ${selectedWarehouse.warehouse_code} - ${selectedWarehouse.warehouse_name || ''}` : 'TẤT CẢ KHO'}</strong>
          <span>TỪ NGÀY: {dateDisplay(fromDate)} ĐẾN NGÀY: {dateDisplay(toDate)}</span>
        </div>
        <label><span>Từ ngày</span><input type="date" value={draftFromDate} disabled={isReportSwitching} onChange={(event) => setDraftFromDate(event.currentTarget.value)} /></label>
        <label><span>Đến ngày</span><input type="date" value={draftToDate} disabled={isReportSwitching} onChange={(event) => setDraftToDate(event.currentTarget.value)} /></label>
        <button type="button" className="btn-secondary" disabled={isReportSwitching} onClick={applyDateFilter}>{isReportSwitching ? 'Đang áp dụng...' : 'Áp dụng'}</button>
      </div>
      <div className="report-kpi-grid">
        <div><strong>{formatNumber(reportTotals.salesQuantity)}</strong><span>Số lượng bán ra</span></div>
        <div><strong>{formatNumber(reportTotals.saleAmount)}</strong><span>Tiền hàng</span></div>
        <div><strong>{formatNumber(reportTotals.costAmount)}</strong><span>Tiền vốn</span></div>
        <div><strong>{formatNumber(reportTotals.warehouseCount)}</strong><span>Kho có phát sinh</span></div>
      </div>
      <div className="allocation-report-toolbar">
        <div className="segmented-control">
          <button type="button" className={tab === 'sales' ? 'active' : ''} disabled={isReportSwitching} onClick={() => changeTab('sales')}>Báo cáo bán hàng</button>
          <button type="button" className={tab === 'ledger' ? 'active' : ''} disabled={isReportSwitching} onClick={() => changeTab('ledger')}>Sổ chi tiết</button>
          <button type="button" className={tab === 'inventory' ? 'active' : ''} disabled={isReportSwitching} onClick={() => changeTab('inventory')}>Tổng hợp NXT</button>
        </div>
        <div className="report-toolbar-right">
          <label className="report-warehouse-filter">
            <span>Kho</span>
            <select value={warehouseFilter} disabled={isReportSwitching} onChange={(event) => changeWarehouse(event.currentTarget.value)}>
              <option value="">Tất cả kho</option>
              {warehouses.map((warehouse) => <option key={warehouse.warehouse_code} value={warehouse.warehouse_code}>{warehouse.warehouse_code}{warehouse.warehouse_name ? ` - ${warehouse.warehouse_name}` : ''}</option>)}
            </select>
          </label>
        </div>
      </div>
      {tab === 'sales' && (
        <div className="allocation-report-layout single">
          <SalesSummaryPanel title={selectedWarehouse ? `Báo cáo bán hàng - ${selectedWarehouse.warehouse_code}${selectedWarehouse.warehouse_name ? ` - ${selectedWarehouse.warehouse_name}` : ''}` : 'Báo cáo bán hàng - tất cả kho'} rows={salesSummaryRows} onOpen={(row) => setSelectedSalesKey(row.key)} />
        </div>
      )}
      {tab === 'ledger' && (
        <div className="allocation-report-layout single">
          <LedgerSectionsView title={selectedWarehouse ? `Sổ chi tiết - ${selectedWarehouse.warehouse_code}${selectedWarehouse.warehouse_name ? ` - ${selectedWarehouse.warehouse_name}` : ''}` : 'Sổ chi tiết hàng hóa - tất cả kho'} sections={ledgerSections} cacheKey={ledgerRenderKey} cacheRef={reportRenderCache} />
        </div>
      )}
      {tab === 'inventory' && (
        <div className="allocation-report-layout single">
          <InventorySummaryPanel rows={inventorySummaryRows} selectedKey={selectedInventoryKey} onSelect={setSelectedInventoryKey} />
          {selectedInventoryKey && <InventoryDetailTable rows={inventoryDetailRows} />}
        </div>
      )}

      {isReportSwitching && (
        <div className="report-switch-overlay" aria-live="polite" aria-busy="true">
          <div className="loading-spinner" />
          <h3>Đang đổi báo cáo</h3>
          <p>Đang chuẩn bị dữ liệu và dựng lại bảng...</p>
        </div>
      )}

      {selectedSalesRow && (
        <SalesExplainModal
          row={selectedSalesRow}
          details={selectedSalesDetails}
          ledgerRows={selectedLedgerRows}
          onClose={() => setSelectedSalesKey('')}
        />
      )}
    </div>
  );
}

function ProgressiveRowsNotice({ colSpan, visibleCount, total }: { colSpan: number; visibleCount: number; total: number }) {
  return <tr className="report-progressive-row"><td colSpan={colSpan}>Đang hiển thị {formatNumber(visibleCount)} / {formatNumber(total)} dòng...</td></tr>;
}

function ProgressiveSectionsNotice({ visibleCount, total }: { visibleCount: number; total: number }) {
  return <div className="report-progressive-note">Đang hiển thị {formatNumber(visibleCount)} / {formatNumber(total)} nhóm sổ chi tiết...</div>;
}
function SalesSummaryPanel({ title, rows, onOpen }: { title: string; rows: InventorySalesSummaryRow[]; onOpen: (row: InventorySalesSummaryRow) => void }) {
  const totals = salesTotals(rows);
  const { visibleRows, visibleCount, isRendering } = useProgressiveRows(rows);
  return (
    <div className="report-table-panel">
      <h4>{title}</h4>
      <div className="report-table-scroll">
        <table className="data-table report-table sales-report-table">
          <thead><tr><th>Kho</th><th>Mã VT</th><th>Tên hàng</th><th>ĐVT</th><th>SL</th><th>Tiền vốn</th><th>Tiền hàng</th><th>Lãi/lỗ</th><th>%</th><th>Thuế</th><th>Tổng TT</th></tr></thead>
          <tbody>
            {visibleRows.map((row) => (
              <tr key={row.key} onDoubleClick={() => onOpen(row)} title="Double-click để xem giải thích giá vốn">
                <td>{row.warehouse_code}</td><td>{row.variant_code}</td><td>{row.product_name}</td><td>{row.unit_name}</td>
                <td className="number-cell">{formatNumber(row.quantity)}</td><td className="number-cell">{formatNumber(row.cost_amount)}</td><td className="number-cell">{formatNumber(row.sale_amount)}</td><td className="number-cell">{formatNumber(row.profit_amount)}</td><td className="number-cell">{formatPercent(row.margin_percent)}</td><td className="number-cell">{formatNumber(row.tax_amount)}</td><td className="number-cell">{formatNumber(row.total_amount)}</td>
              </tr>
            ))}
            {isRendering && <ProgressiveRowsNotice colSpan={11} visibleCount={visibleCount} total={rows.length} />}
            {!rows.length && <tr><td colSpan={11}>Chưa có dữ liệu báo cáo bán hàng.</td></tr>}
          </tbody>
          {!!rows.length && <tfoot><tr><td></td><td>Tổng cộng:</td><td></td><td></td><td className="number-cell">{formatNumber(totals.quantity)}</td><td className="number-cell">{formatNumber(totals.cost_amount)}</td><td className="number-cell">{formatNumber(totals.sale_amount)}</td><td className="number-cell">{formatNumber(totals.profit_amount)}</td><td className="number-cell">{formatPercent(totals.margin_percent)}</td><td className="number-cell">{formatNumber(totals.tax_amount)}</td><td className="number-cell">{formatNumber(totals.total_amount)}</td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

function InventorySummaryPanel({ rows, selectedKey, onSelect }: { rows: InventorySummaryRow[]; selectedKey: string; onSelect: (key: string) => void }) {
  const totals = inventoryTotals(rows);
  const { visibleRows, visibleCount, isRendering } = useProgressiveRows(rows);
  return (
    <div className="report-table-panel">
      <h4>Tổng hợp nhập xuất tồn</h4>
      <div className="report-table-scroll">
        <table className="data-table report-table inventory-report-table">
          <thead><tr><th>Kho</th><th>Mã VT</th><th>Tên hàng</th><th>ĐVT</th><th>Tồn đầu SL</th><th>Tồn đầu tiền</th><th>Nhập SL</th><th>Nhập tiền</th><th>Xuất SL</th><th>Xuất tiền</th><th>Tồn cuối SL</th><th>Tồn cuối tiền</th></tr></thead>
          <tbody>
            {visibleRows.map((row) => (
              <tr key={row.key} className={selectedKey === row.key ? 'selected-row' : ''} onDoubleClick={() => onSelect(row.key)}>
                <td>{row.warehouse_code}</td><td>{row.variant_code}</td><td>{row.product_name}</td><td>{row.unit_name}</td>
                <td className="number-cell">{formatNumber(row.opening_qty)}</td><td className="number-cell">{formatNumber(row.opening_amount)}</td><td className="number-cell">{formatNumber(row.in_qty)}</td><td className="number-cell">{formatNumber(row.in_amount)}</td><td className="number-cell">{formatNumber(row.out_qty)}</td><td className="number-cell">{formatNumber(row.out_amount)}</td><td className="number-cell">{formatNumber(row.ending_qty)}</td><td className="number-cell">{formatNumber(row.ending_amount)}</td>
              </tr>
            ))}
            {isRendering && <ProgressiveRowsNotice colSpan={12} visibleCount={visibleCount} total={rows.length} />}
            {!rows.length && <tr><td colSpan={12}>Chưa có dữ liệu tổng hợp nhập xuất tồn.</td></tr>}
          </tbody>
          {!!rows.length && <tfoot><tr><td></td><td>Tổng cộng:</td><td></td><td></td><td className="number-cell">{formatNumber(totals.opening_qty)}</td><td className="number-cell">{formatNumber(totals.opening_amount)}</td><td className="number-cell">{formatNumber(totals.in_qty)}</td><td className="number-cell">{formatNumber(totals.in_amount)}</td><td className="number-cell">{formatNumber(totals.out_qty)}</td><td className="number-cell">{formatNumber(totals.out_amount)}</td><td className="number-cell">{formatNumber(totals.ending_qty)}</td><td className="number-cell">{formatNumber(totals.ending_amount)}</td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

function SalesExplainModal({ row, details, ledgerRows, onClose }: { row: InventorySalesSummaryRow; details: InventorySalesDetailRow[]; ledgerRows: InventoryLedgerDetailRow[]; onClose: () => void }) {
  const totals = salesTotals([row]);
  const missingCount = Number(row.cost_missing_count || details.filter((detail) => detail.cost_missing).length || 0);
  return (
    <div className="modal-overlay report-modal-overlay" role="dialog" aria-modal="true">
      <div className="modal-content report-detail-modal">
        <div className="modal-header">
          <div>
            <h2>Giải thích báo cáo bán hàng</h2>
            <p className="muted">{row.warehouse_code} - {row.variant_code} - {row.product_name}</p>
          </div>
          <button type="button" className="modal-close" onClick={onClose}>×</button>
        </div>
        <div className="modal-body report-detail-body">
          <div className="report-kpi-grid compact-kpis">
            <div><strong>{formatNumber(row.quantity)}</strong><span>Số lượng</span></div>
            <div><strong>{formatNumber(row.sale_amount)}</strong><span>Tiền hàng</span></div>
            <div><strong>{formatNumber(row.cost_amount)}</strong><span>Tiền vốn từ sổ</span></div>
            <div><strong>{formatPercent(totals.margin_percent)}</strong><span>% lãi</span></div>
          </div>
          <div className="report-proof-note">
            <strong>Nguồn chứng minh:</strong> Tiền vốn của dòng này bằng tổng cột <b>Xuất thành tiền</b> trong Sổ chi tiết cùng kho/mã. {missingCount ? `${missingCount} dòng không có nhập/tồn trong kho nên giá vốn và % lãi được để 0 để không làm lệch tổng.` : 'Các dòng có nhập/tồn sẽ lấy đúng giá vốn từ dòng xuất trong sổ chi tiết.'}
          </div>
          <div className="modal-report-grid">
            <SalesDetailTable rows={details} />
            <LedgerProofTable rows={ledgerRows} />
          </div>
        </div>
        <div className="modal-footer"><button type="button" className="btn-secondary" onClick={onClose}>Đóng</button></div>
      </div>
    </div>
  );
}

function SalesDetailTable({ rows }: { rows: InventorySalesDetailRow[] }) {
  const totals = salesDetailTotals(rows);
  return (
    <div className="report-table-panel detail-panel">
      <h4>Chi tiết bán hàng</h4>
      <div className="report-table-scroll">
        <table className="data-table report-table detail-table sales-detail-table">
          <thead><tr><th>Ngày</th><th>Số HĐ</th><th>Khách hàng</th><th>Mã VT</th><th>Tên hàng</th><th>SL</th><th>Giá vốn</th><th>Tiền vốn</th><th>Giá bán</th><th>Tiền hàng</th><th>Thuế</th><th>Tổng TT</th><th>Ghi chú</th></tr></thead>
          <tbody>
            {rows.map((row) => <tr key={row.key}><td>{row.invoice_date}</td><td>{row.invoice_no}</td><td>{row.customer}</td><td>{row.variant_code}</td><td>{row.product_name}</td><td className="number-cell">{formatNumber(row.quantity)}</td><td className="number-cell">{formatNumber(row.cost_price)}</td><td className="number-cell">{formatNumber(row.cost_amount)}</td><td className="number-cell">{formatNumber(row.sale_price)}</td><td className="number-cell">{formatNumber(row.sale_amount)}</td><td className="number-cell">{formatNumber(row.tax_amount)}</td><td className="number-cell">{formatNumber(row.total_amount)}</td><td>{row.cost_missing ? 'Không có nhập/tồn trong kho, giá vốn tính 0' : ''}</td></tr>)}
            {!rows.length && <tr><td colSpan={13}>Không có dòng chi tiết bán hàng cho mã này.</td></tr>}
          </tbody>
          {!!rows.length && <tfoot><tr><td></td><td>Tổng cộng:</td><td></td><td></td><td></td><td className="number-cell">{formatNumber(totals.quantity)}</td><td></td><td className="number-cell">{formatNumber(totals.cost_amount)}</td><td></td><td className="number-cell">{formatNumber(totals.sale_amount)}</td><td className="number-cell">{formatNumber(totals.tax_amount)}</td><td className="number-cell">{formatNumber(totals.total_amount)}</td><td></td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

function LedgerProofTable({ rows }: { rows: InventoryLedgerDetailRow[] }) {
  return (
    <div className="report-table-panel detail-panel">
      <h4>Sổ chi tiết chứng minh giá vốn</h4>
      <div className="report-table-scroll">
        <table className="data-table report-table detail-table ledger-proof-table">
          <thead><tr><th>Ngày</th><th>Số</th><th>Khách hàng</th><th>Diễn giải</th><th>TK ĐƯ</th><th>Đơn giá vốn</th><th>Nhập SL</th><th>Nhập tiền</th><th>Xuất SL</th><th>Xuất tiền</th><th>Tồn SL</th><th>Tồn tiền</th><th>Ghi chú</th></tr></thead>
          <tbody>
            {rows.map((row) => <tr key={row.key} className={row.row_type === 'sale' ? 'ledger-sale-row' : row.row_type === 'purchase' || row.row_type === 'purchase_future_reorder' ? 'ledger-purchase-row' : ''}><td>{row.date}</td><td>{row.doc_no}</td><td>{row.customer}</td><td>{row.description}</td><td>{row.account}</td><td className="number-cell">{formatNumber(row.unit_price)}</td><td className="number-cell">{formatNumber(row.qty_in)}</td><td className="number-cell">{formatNumber(row.amount_in)}</td><td className="number-cell">{formatNumber(row.qty_out)}</td><td className="number-cell">{formatNumber(row.amount_out)}</td><td className="number-cell">{formatNumber(row.running_qty)}</td><td className="number-cell">{formatNumber(row.running_amount)}</td><td>{row.logic_note}</td></tr>)}
            {!rows.length && <tr><td colSpan={13}>Không có sổ chi tiết cho dòng này.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function InventoryDetailTable({ rows }: { rows: InventoryLedgerDetailRow[] }) {
  return (
    <div className="report-table-panel detail-panel">
      <h4>Chi tiết nhập xuất tồn</h4>
      <div className="report-table-scroll">
        <table className="data-table report-table detail-table">
          <thead><tr><th>Ngày</th><th>Số</th><th>Khách hàng</th><th>Diễn giải</th><th>TK ĐƯ</th><th>Nhập SL</th><th>Nhập tiền</th><th>Xuất SL</th><th>Xuất tiền</th><th>Tồn SL</th><th>Tồn tiền</th><th>Ghi chú</th></tr></thead>
          <tbody>
            {rows.map((row) => <tr key={row.key}><td>{row.date}</td><td>{row.doc_no}</td><td>{row.customer}</td><td>{row.description}</td><td>{row.account}</td><td className="number-cell">{formatNumber(row.qty_in)}</td><td className="number-cell">{formatNumber(row.amount_in)}</td><td className="number-cell">{formatNumber(row.qty_out)}</td><td className="number-cell">{formatNumber(row.amount_out)}</td><td className="number-cell">{formatNumber(row.running_qty)}</td><td className="number-cell">{formatNumber(row.running_amount)}</td><td>{row.logic_note}</td></tr>)}
            {!rows.length && <tr><td colSpan={12}>Double-click một dòng tổng hợp NXT để xem sổ chi tiết.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function LedgerSectionsView({ title, sections, cacheKey, cacheRef }: { title: string; sections: LedgerSection[]; cacheKey: string; cacheRef: { current: Map<string, number> } }) {
  const { visibleRows: visibleSections, visibleCount, isRendering } = useProgressiveRows(sections, true, REPORT_LEDGER_SECTION_BATCH_SIZE, `${cacheKey}|sections`, cacheRef);
  const visibleTotals = useMemo(() => ledgerSectionTotals(visibleSections), [visibleSections]);
  return (
    <div className="report-table-panel ledger-panel">
      <h4>{title}</h4>
      <div className="ledger-scroll-shell">
        <div className="report-table-scroll ledger-sections">
          {visibleSections.map((section) => <LedgerSectionTable key={section.key} section={section} cacheKey={cacheKey} cacheRef={cacheRef} />)}
          {!sections.length && <p className="muted">Chưa có dữ liệu sổ chi tiết.</p>}
        </div>
        {!!sections.length && <LedgerGrandTotal totals={visibleTotals} visibleCount={visibleCount} totalCount={sections.length} isRendering={isRendering} />}
      </div>
    </div>
  );
}

function LedgerSectionTable({ section, cacheKey, cacheRef }: { section: LedgerSection; cacheKey: string; cacheRef: { current: Map<string, number> } }) {
  const { visibleRows, visibleCount, isRendering } = useProgressiveRows(section.rows, true, REPORT_LEDGER_ROW_BATCH_SIZE, `${cacheKey}|row|${section.key}`, cacheRef);
  return (
    <section className="ledger-section">
      <div className="ledger-section-heading">
        <strong>KHO: {section.warehouse_code}{section.warehouse_name ? ` - ${section.warehouse_name}` : ''}</strong>
        <span>Vật tư: {section.variant_code} - {section.product_name || ''}, Đvt: {section.unit_name || '-'}, TK: {section.account || ''}</span>
      </div>
      <table className="data-table report-table detail-table ledger-section-table">
        <thead><tr><th>Ngày</th><th>Số</th><th>Khách hàng</th><th>Diễn giải</th><th>TK ĐƯ</th><th>Đơn giá</th><th>Nhập SL</th><th>Nhập tiền</th><th>Xuất SL</th><th>Xuất tiền</th><th>Tồn SL</th><th>Tồn tiền</th></tr></thead>
        <tbody>
          <LedgerSummaryRow label="Tồn đầu kỳ" qtyIn={section.opening_qty} amountIn={section.opening_amount} runningQty={section.opening_qty} runningAmount={section.opening_amount} />
          <LedgerSummaryRow label="Nhập trong kỳ" qtyIn={section.in_qty} amountIn={section.in_amount} />
          <LedgerSummaryRow label="Xuất trong kỳ" qtyOut={section.out_qty} amountOut={section.out_amount} />
          <LedgerSummaryRow label="Tồn cuối kỳ" qtyIn={section.ending_qty} amountIn={section.ending_amount} runningQty={section.ending_qty} runningAmount={section.ending_amount} />
          {visibleRows.map((row) => <tr key={row.key} className={row.row_type === 'sale' ? 'ledger-sale-row' : row.row_type === 'purchase' || row.row_type === 'purchase_future_reorder' ? 'ledger-purchase-row' : ''}><td>{row.date}</td><td>{row.doc_no}</td><td>{row.customer}</td><td>{row.description}</td><td>{row.account}</td><td className="number-cell">{formatNumber(row.unit_price)}</td><td className="number-cell">{formatNumber(row.qty_in)}</td><td className="number-cell">{formatNumber(row.amount_in)}</td><td className="number-cell">{formatNumber(row.qty_out)}</td><td className="number-cell">{formatNumber(row.amount_out)}</td><td className="number-cell">{formatNumber(row.running_qty)}</td><td className="number-cell">{formatNumber(row.running_amount)}</td></tr>)}
          {isRendering && <ProgressiveRowsNotice colSpan={12} visibleCount={visibleCount} total={section.rows.length} />}
        </tbody>
      </table>
    </section>
  );
}

function LedgerSummaryRow({ label, qtyIn = '', amountIn = '', qtyOut = '', amountOut = '', runningQty = '', runningAmount = '' }: { label: string; qtyIn?: number | string; amountIn?: number | string; qtyOut?: number | string; amountOut?: number | string; runningQty?: number | string; runningAmount?: number | string }) {
  return <tr className="ledger-summary-row"><td></td><td></td><td></td><td>{label}</td><td></td><td></td><td className="number-cell">{formatNumber(qtyIn)}</td><td className="number-cell">{formatNumber(amountIn)}</td><td className="number-cell">{formatNumber(qtyOut)}</td><td className="number-cell">{formatNumber(amountOut)}</td><td className="number-cell">{formatNumber(runningQty)}</td><td className="number-cell">{formatNumber(runningAmount)}</td></tr>;
}

function LedgerGrandTotal({ totals, visibleCount, totalCount, isRendering }: { totals: ReturnType<typeof ledgerSectionTotals>; visibleCount: number; totalCount: number; isRendering: boolean }) {
  const progressPercent = totalCount ? Math.min(100, Math.round((visibleCount / totalCount) * 100)) : 0;
  const inQty = formatNumber(totals.in_qty);
  const inAmount = formatNumber(totals.in_amount);
  const outQty = formatNumber(totals.out_qty);
  const outAmount = formatNumber(totals.out_amount);
  const endingQty = formatNumber(totals.ending_qty);
  const endingAmount = formatNumber(totals.ending_amount);
  return (
    <div className="ledger-grand-total" aria-live="polite">
      <div className="ledger-total-caption">
        <strong>{isRendering ? `Đang hiển thị ${formatNumber(visibleCount)} / ${formatNumber(totalCount)} nhóm sổ chi tiết` : `Đã hiển thị đủ ${formatNumber(totalCount)} nhóm sổ chi tiết`}</strong>
        <span>{isRendering ? `${progressPercent}%` : '100%'}</span>
      </div>
      <div className="ledger-render-progress" aria-hidden="true"><span style={{ width: `${progressPercent}%` }} /></div>
      <table className="data-table report-table detail-table ledger-section-table">
        <tbody>
          <tr>
            <td></td>
            <td></td>
            <td></td>
            <td>Tổng cộng đang hiển thị:</td>
            <td></td>
            <td></td>
            <td className="number-cell ledger-total-number" title={inQty}>{inQty}</td>
            <td className="number-cell ledger-total-number" title={inAmount}>{inAmount}</td>
            <td className="number-cell ledger-total-number" title={outQty}>{outQty}</td>
            <td className="number-cell ledger-total-number" title={outAmount}>{outAmount}</td>
            <td className="number-cell ledger-total-number" title={endingQty}>{endingQty}</td>
            <td className="number-cell ledger-total-number" title={endingAmount}>{endingAmount}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function filterByWarehouse<T extends { warehouse_code?: string }>(rows: T[], warehouseCode: string) {
  return warehouseCode ? rows.filter((row) => row.warehouse_code === warehouseCode) : rows;
}

function filterSalesDetailsByDate(rows: InventorySalesDetailRow[], fromDate: string, toDate: string) {
  return rows.filter((row) => dateInRange(reportDateValue(row.invoice_date_iso, row.invoice_date), fromDate, toDate));
}

function filterLedgerDetailsByDate(rows: InventoryLedgerDetailRow[], fromDate: string, toDate: string) {
  return rows.filter((row) => isOpeningLedgerRow(row) || dateInRange(ledgerRowDate(row), fromDate, toDate));
}

function isOpeningLedgerRow(row: InventoryLedgerDetailRow) {
  return row.row_type === 'opening' || row.description === 'Tồn đầu kỳ';
}

function ledgerRowDate(row: InventoryLedgerDetailRow) {
  return reportDateValue(row.date_iso, row.date);
}

function reportDateValue(primary?: string, fallback?: string) {
  return normalizeReportDate(primary) || normalizeReportDate(fallback);
}

function normalizeReportDate(value?: string) {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  const isoMatch = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2].padStart(2, '0')}-${isoMatch[3].padStart(2, '0')}`;
  const localMatch = raw.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (localMatch) return `${localMatch[3]}-${localMatch[2].padStart(2, '0')}-${localMatch[1].padStart(2, '0')}`;
  return raw;
}

function dateInRange(value: string | undefined, fromDate: string, toDate: string) {
  const date = normalizeReportDate(value);
  const from = normalizeReportDate(fromDate);
  const to = normalizeReportDate(toDate);
  if (!date) return true;
  return (!from || date >= from) && (!to || date <= to);
}

function isBeforeRange(value: string | undefined, fromDate: string) {
  const date = normalizeReportDate(value);
  const from = normalizeReportDate(fromDate);
  return Boolean(date && from && date < from);
}
function buildSalesSummaryRows(details: InventorySalesDetailRow[], fallback: InventorySalesSummaryRow[]): InventorySalesSummaryRow[] {
  if (!details.length) return fallback;
  const grouped = new Map<string, InventorySalesSummaryRow>();
  details.forEach((row) => {
    const key = row.summary_key || [row.warehouse_code, row.variant_code, row.product_name, row.unit_name].join('|||');
    const item = grouped.get(key) ?? {
      key,
      warehouse_code: row.warehouse_code || '',
      variant_code: row.variant_code || '',
      product_name: row.product_name || '',
      unit_name: row.unit_name || '',
      quantity: 0,
      cost_amount: 0,
      sale_amount: 0,
      profit_amount: 0,
      tax_amount: 0,
      total_amount: 0,
      row_count: 0,
      margin_sale_amount: 0,
      cost_missing_count: 0,
    };
    const saleAmount = Number(row.sale_amount || 0);
    const costAmount = Number(row.cost_amount || 0);
    const costMissing = Boolean(row.cost_missing);
    item.quantity += Number(row.quantity || 0);
    item.cost_amount += costAmount;
    item.sale_amount += saleAmount;
    item.profit_amount += costMissing ? 0 : Number(row.profit_amount ?? (saleAmount - costAmount));
    item.tax_amount += Number(row.tax_amount || 0);
    item.total_amount += Number(row.total_amount || 0);
    item.row_count = Number(item.row_count || 0) + 1;
    item.margin_sale_amount = Number(item.margin_sale_amount || 0) + (costMissing ? 0 : saleAmount);
    item.cost_missing_count = Number(item.cost_missing_count || 0) + (costMissing ? 1 : 0);
    item.margin_percent = marginPercentFromProfit(item.profit_amount, item.margin_sale_amount);
    grouped.set(key, item);
  });
  return Array.from(grouped.values()).map((row, index) => ({ ...row, index: index + 1 }));
}

function buildInventorySummaryRows(details: InventoryLedgerDetailRow[], fallback: InventorySummaryRow[], fromDate = '', toDate = ''): InventorySummaryRow[] {
  if (!details.length) return fallback;
  const grouped = new Map<string, InventorySummaryRow>();
  details.forEach((row) => {
    const key = row.summary_key || [row.warehouse_code, row.variant_code].join('|||');
    const item = grouped.get(key) ?? {
      key,
      warehouse_code: row.warehouse_code || '',
      account: row.account || '',
      variant_code: row.variant_code || '',
      product_name: row.product_name || '',
      unit_name: row.unit_name || '',
      opening_qty: 0,
      opening_amount: 0,
      in_qty: 0,
      in_amount: 0,
      out_qty: 0,
      out_amount: 0,
      ending_qty: 0,
      ending_amount: 0,
      row_count: 0,
    };
    const rowDate = ledgerRowDate(row);
    const isOpening = isOpeningLedgerRow(row) || isBeforeRange(rowDate, fromDate);
    if (isOpening) {
      item.opening_qty += Number(row.qty_in || 0) - Number(row.qty_out || 0);
      item.opening_amount += Number(row.amount_in || 0) - Number(row.amount_out || 0);
    } else if (dateInRange(rowDate, fromDate, toDate)) {
      item.in_qty += Number(row.qty_in || 0);
      item.in_amount += Number(row.amount_in || 0);
      item.out_qty += Number(row.qty_out || 0);
      item.out_amount += Number(row.amount_out || 0);
      item.row_count = Number(item.row_count || 0) + 1;
    } else {
      return;
    }
    item.ending_qty = item.opening_qty + item.in_qty - item.out_qty;
    item.ending_amount = item.opening_amount + item.in_amount - item.out_amount;
    grouped.set(key, item);
  });
  return Array.from(grouped.values()).map((row, index) => ({ ...row, index: index + 1 }));
}

type LedgerSection = {
  key: string;
  warehouse_code: string;
  warehouse_name?: string;
  account?: string;
  variant_code: string;
  product_name?: string;
  unit_name?: string;
  opening_qty: number;
  opening_amount: number;
  in_qty: number;
  in_amount: number;
  out_qty: number;
  out_amount: number;
  ending_qty: number;
  ending_amount: number;
  rows: InventoryLedgerDetailRow[];
};

function buildLedgerSections(rows: InventoryLedgerDetailRow[], warehouses: Array<{ warehouse_code: string; warehouse_name?: string; account?: string }>, fromDate = '', toDate = ''): LedgerSection[] {
  const warehouseMap = new Map(warehouses.map((warehouse) => [warehouse.warehouse_code, warehouse]));
  const grouped = new Map<string, LedgerSection>();
  rows.forEach((row) => {
    const key = row.summary_key || inventoryReportSummaryKey(row.warehouse_code || '', row.variant_code || '');
    const warehouse = warehouseMap.get(row.warehouse_code || '');
    const section = grouped.get(key) ?? {
      key,
      warehouse_code: row.warehouse_code || '',
      warehouse_name: warehouse?.warehouse_name || '',
      account: row.account || warehouse?.account || '',
      variant_code: row.variant_code || '',
      product_name: row.product_name || '',
      unit_name: row.unit_name || '',
      opening_qty: 0,
      opening_amount: 0,
      in_qty: 0,
      in_amount: 0,
      out_qty: 0,
      out_amount: 0,
      ending_qty: 0,
      ending_amount: 0,
      rows: [],
    };
    if (!section.product_name && row.product_name) section.product_name = row.product_name;
    if (!section.unit_name && row.unit_name) section.unit_name = row.unit_name;
    if (!section.account && row.account) section.account = row.account;
    const rowDate = ledgerRowDate(row);
    const isOpening = isOpeningLedgerRow(row) || isBeforeRange(rowDate, fromDate);
    if (isOpening) {
      section.opening_qty += Number(row.qty_in || 0) - Number(row.qty_out || 0);
      section.opening_amount += Number(row.amount_in || 0) - Number(row.amount_out || 0);
    } else if (dateInRange(rowDate, fromDate, toDate)) {
      section.in_qty += Number(row.qty_in || 0);
      section.in_amount += Number(row.amount_in || 0);
      section.out_qty += Number(row.qty_out || 0);
      section.out_amount += Number(row.amount_out || 0);
      section.rows.push(row);
    } else {
      return;
    }
    section.ending_qty = section.opening_qty + section.in_qty - section.out_qty;
    section.ending_amount = section.opening_amount + section.in_amount - section.out_amount;
    grouped.set(key, section);
  });
  return Array.from(grouped.values()).sort((a, b) => `${a.warehouse_code}|||${a.variant_code}`.localeCompare(`${b.warehouse_code}|||${b.variant_code}`));
}

function reportKpis(salesRows: InventorySalesSummaryRow[], inventoryRows: InventorySummaryRow[]) {
  const activeWarehouses = new Set<string>();
  salesRows.forEach((row) => row.warehouse_code && activeWarehouses.add(row.warehouse_code));
  inventoryRows.forEach((row) => row.warehouse_code && activeWarehouses.add(row.warehouse_code));
  return {
    salesQuantity: salesRows.reduce((sum, row) => sum + Number(row.quantity || 0), 0),
    saleAmount: salesRows.reduce((sum, row) => sum + Number(row.sale_amount || 0), 0),
    costAmount: salesRows.reduce((sum, row) => sum + Number(row.cost_amount || 0), 0),
    warehouseCount: activeWarehouses.size,
  };
}

function salesTotals(rows: InventorySalesSummaryRow[]) {
  const totals = rows.reduce((acc, row) => {
    acc.quantity += Number(row.quantity || 0);
    acc.cost_amount += Number(row.cost_amount || 0);
    acc.sale_amount += Number(row.sale_amount || 0);
    acc.profit_amount += Number(row.profit_amount || 0);
    acc.tax_amount += Number(row.tax_amount || 0);
    acc.total_amount += Number(row.total_amount || 0);
    acc.margin_sale_amount += Number(row.margin_sale_amount ?? (Number(row.cost_missing_count || 0) ? 0 : row.sale_amount || 0));
    return acc;
  }, { quantity: 0, cost_amount: 0, sale_amount: 0, profit_amount: 0, tax_amount: 0, total_amount: 0, margin_sale_amount: 0, margin_percent: 0 });
  totals.margin_percent = marginPercentFromProfit(totals.profit_amount, totals.margin_sale_amount);
  return totals;
}

function salesDetailTotals(rows: InventorySalesDetailRow[]) {
  return rows.reduce((acc, row) => {
    acc.quantity += Number(row.quantity || 0);
    acc.cost_amount += Number(row.cost_amount || 0);
    acc.sale_amount += Number(row.sale_amount || 0);
    acc.tax_amount += Number(row.tax_amount || 0);
    acc.total_amount += Number(row.total_amount || 0);
    return acc;
  }, { quantity: 0, cost_amount: 0, sale_amount: 0, tax_amount: 0, total_amount: 0 });
}

function inventoryTotals(rows: InventorySummaryRow[]) {
  return rows.reduce((acc, row) => {
    acc.opening_qty += Number(row.opening_qty || 0);
    acc.opening_amount += Number(row.opening_amount || 0);
    acc.in_qty += Number(row.in_qty || 0);
    acc.in_amount += Number(row.in_amount || 0);
    acc.out_qty += Number(row.out_qty || 0);
    acc.out_amount += Number(row.out_amount || 0);
    acc.ending_qty += Number(row.ending_qty || 0);
    acc.ending_amount += Number(row.ending_amount || 0);
    return acc;
  }, { opening_qty: 0, opening_amount: 0, in_qty: 0, in_amount: 0, out_qty: 0, out_amount: 0, ending_qty: 0, ending_amount: 0 });
}

function ledgerSectionTotals(sections: LedgerSection[]) {
  return sections.reduce((acc, section) => {
    acc.opening_qty += Number(section.opening_qty || 0);
    acc.opening_amount += Number(section.opening_amount || 0);
    acc.in_qty += Number(section.in_qty || 0);
    acc.in_amount += Number(section.in_amount || 0);
    acc.out_qty += Number(section.out_qty || 0);
    acc.out_amount += Number(section.out_amount || 0);
    acc.ending_qty += Number(section.ending_qty || 0);
    acc.ending_amount += Number(section.ending_amount || 0);
    return acc;
  }, { opening_qty: 0, opening_amount: 0, in_qty: 0, in_amount: 0, out_qty: 0, out_amount: 0, ending_qty: 0, ending_amount: 0 });
}

function inventoryReportSummaryKey(warehouseCode: string, variantCode: string) {
  return [warehouseCode || '', variantCode || ''].join('|||');
}

function marginPercentFromProfit(profitAmount: number, baseAmount: number) {
  if (!baseAmount) return 0;
  return (Number(profitAmount || 0) / Number(baseAmount || 0)) * 100;
}

function dateDisplay(value: string) {
  if (!value) return '';
  const [year, month, day] = value.split('-');
  return year && month && day ? `${day}/${month}/${year}` : value;
}

function formatNumber(value: unknown) {
  if (typeof value === 'number') return value.toLocaleString('en-US');
  if (typeof value === 'string') return value;
  return '';
}

function formatPercent(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '';
  return `${value.toLocaleString('en-US', { maximumFractionDigits: 2 })}%`;
}
