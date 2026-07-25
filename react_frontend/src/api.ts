import type { CompanyRow, EstimateAnalysis, EstimateSheetSelection, EstimateUploadSummary, FormatMappingDefaults, FormMappingPreset, GenericAnalyzeResult, InventoryAllocationConfig, InventoryAllocationJob, InventoryPair, InventoryRule, InvoiceStatusOption, LicenseStatus, MissingMstCompanyWarning, MatchRow, OperationProgress, ProcessResult, ProcessedFileStats, ReviewProduct, ReviewRow, UploadSummary, WorkflowJob, WorkflowSession, UpdateManifest } from './types';

async function parseJsonResponse<T>(response: Response): Promise<T> {
  const text = await response.text();
  const endpoint = response.url || 'API';
  const parsePayload = () => {
    if (!text.trim()) return { empty: true, value: null as unknown };
    try {
      return { empty: false, value: JSON.parse(text) as unknown };
    } catch {
      return { empty: false, value: null as unknown };
    }
  };
  const parsed = parsePayload();
  if (response.ok) {
    if (parsed.empty) {
      throw new Error(`Backend trả về phản hồi rỗng từ ${endpoint}. Vui lòng thử lại, nếu còn lỗi hãy gửi trạng thái debug cho dev.`);
    }
    return parsed.value as T;
  }
  const payload = parsed.value && typeof parsed.value === 'object' ? parsed.value as Record<string, unknown> : {};
  const detail = payload.detail || payload.error;
  if (detail && typeof detail === 'object') {
    const detailRecord = detail as Record<string, unknown>;
    const message = String(detailRecord.message || response.statusText);
    const code = String(detailRecord.code || 'API_ERROR');
    const operationId = String(detailRecord.operation_id || '-');
    throw new Error(`${message} [${code}; ${operationId}]`);
  }
  throw new Error(String(detail || text || `${response.status} ${response.statusText} từ ${endpoint}`));
}


async function fetchWithTimeout(input: RequestInfo | URL, init: RequestInit = {}, timeoutMs = 120000, timeoutMessage = 'Yêu cầu xử lý quá lâu và đã được dừng. Vui lòng thử lại hoặc kiểm tra file có đang mở trong Excel không.'): Promise<Response> {
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(input, { ...init, signal: controller.signal });
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') {
      throw new Error(timeoutMessage);
    }
    throw error;
  } finally {
    window.clearTimeout(timeout);
  }
}

const workbookTimeoutMessage = 'Tạo file quá 5 phút chưa xong. Vui lòng kiểm tra cấu hình cột/form mapping, hoặc đóng file Excel đang mở rồi thử lại.';

export async function uploadExcel(file: File, purpose = 'source'): Promise<UploadSummary> {
  const form = new FormData();
  form.append('file', file);
  form.append('purpose', purpose);
  return parseJsonResponse<UploadSummary>(await fetchWithTimeout('/api/files/upload', { method: 'POST', body: form }, 180000));
}

export async function uploadFormTemplate(file: File): Promise<UploadSummary> {
  const form = new FormData();
  form.append('file', file);
  return parseJsonResponse<UploadSummary>(await fetchWithTimeout('/api/templates/upload', { method: 'POST', body: form }, 180000));
}

export async function getWorkflowSession(): Promise<WorkflowSession> {
  return parseJsonResponse<WorkflowSession>(await fetch('/api/session/current'));
}

export async function closeWorkflowSession(): Promise<void> {
  await parseJsonResponse(await fetch('/api/session/close', { method: 'POST' }));
}

export async function startWorkflowProcessJob(payload: {
  savedName: string;
  originalName: string;
  processPayload: Record<string, unknown>;
  processor: 'vietmax' | 'generic';
  retry?: boolean;
}): Promise<WorkflowJob> {
  return parseJsonResponse<WorkflowJob>(await fetch('/api/workflow/process-jobs', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      saved_name: payload.savedName,
      original_name: payload.originalName,
      payload: payload.processPayload,
      processor: payload.processor,
      retry: Boolean(payload.retry),
    }),
  }));
}

export async function getWorkflowJob(jobId: string): Promise<WorkflowJob> {
  return parseJsonResponse<WorkflowJob>(await fetch(`/api/jobs/${encodeURIComponent(jobId)}`));
}

export async function waitForWorkflowJob(initialJob: WorkflowJob, onProgress?: (job: WorkflowJob) => void): Promise<WorkflowJob> {
  let job = initialJob;
  onProgress?.(job);
  while (job.status === 'queued' || job.status === 'running') {
    await new Promise((resolve) => window.setTimeout(resolve, 350));
    job = await getWorkflowJob(job.job_id);
    onProgress?.(job);
  }
  return job;
}

async function runJsonWorkflowJob(kind: string, payload: Record<string, unknown>, retry = false): Promise<WorkflowJob> {
  const initial = await parseJsonResponse<WorkflowJob>(await fetch('/api/workflow/json-jobs', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ kind, payload, retry }),
  }));
  const completed = await waitForWorkflowJob(initial);
  if (completed.status === 'failed') {
    const failure = completed.error;
    throw new Error(`${failure?.message || 'Không thể hoàn tất xử lý.'} [${failure?.code || 'WORKFLOW_ERROR'}; ${completed.operation_id}]`);
  }
  return completed;
}

export async function uploadEstimateWorkbook(file: File): Promise<EstimateUploadSummary> {
  const form = new FormData();
  form.append('file', file);
  return parseJsonResponse<EstimateUploadSummary>(await fetch('/api/estimate/upload', { method: 'POST', body: form }));
}

export async function analyzeEstimateWorkbook(savedName: string, selection: EstimateSheetSelection = {}): Promise<EstimateAnalysis> {
  return parseJsonResponse<EstimateAnalysis>(
    await fetch('/api/estimate/analyze', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, ...selection }),
    }),
  );
}

export async function exportEstimateWorkbook(savedName: string, originalName: string, selection: EstimateSheetSelection = {}): Promise<Blob> {
  const response = await fetch('/api/estimate/export', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ saved_name: savedName, original_name: originalName, ...selection }),
  });
  if (!response.ok) {
    const errorPayload = await response.json().catch(() => ({}));
    throw new Error(String(errorPayload.detail || errorPayload.error || response.statusText));
  }
  return response.blob();
}

export async function getAppConfig(): Promise<Record<string, unknown>> {
  return parseJsonResponse<Record<string, unknown>>(await fetch('/api/config'));
}

export async function analyzeGenericWorkbook(payload: Record<string, unknown>): Promise<GenericAnalyzeResult> {
  return parseJsonResponse<GenericAnalyzeResult>(
    await fetch('/api/check', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    }),
  );
}

export async function fetchInvoiceStatuses(savedName: string, invoiceStatusCol: string, invoiceStatusSkipValues: string[] = []): Promise<InvoiceStatusOption[]> {
  const result = await parseJsonResponse<{ invoice_statuses: InvoiceStatusOption[] }>(
    await fetch('/api/invoice_statuses', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, invoice_status_col: invoiceStatusCol, invoice_status_skip_values: invoiceStatusSkipValues }),
    }),
  );
  return result.invoice_statuses || [];
}

export async function previewGenericProductCodes(payload: { profile: string; products: string[]; word_rules?: Record<string, string>; first_word_rules?: Record<string, string>; repeated_phrase_removals?: string[]; product_code_replacements?: Record<string, string> }) {
  return parseJsonResponse<{ codes: Record<string, string> }>(
    await fetch('/api/product-preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    }),
  );
}

export async function processGenericWorkbook(payload: Record<string, unknown>): Promise<Blob> {
  const response = await fetchWithTimeout('/api/process', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  }, 300000, workbookTimeoutMessage);
  if (!response.ok) {
    const errorPayload = await response.json().catch(() => ({}));
    throw new Error(String(errorPayload.detail || errorPayload.error || response.statusText));
  }
  return response.blob();
}

export async function importProductCodeReplacements(file: File) {
  const form = new FormData();
  form.append('file', file);
  return parseJsonResponse<{ product_code_replacements: Record<string, string>; count: number }>(
    await fetchWithTimeout('/api/product-code-replacements/import', { method: 'POST', body: form }, 180000),
  );
}

export async function processGenericWorkbookCache(payload: Record<string, unknown>, operationId = ''): Promise<ProcessResult> {
  return parseJsonResponse<ProcessResult>(
    await fetchWithTimeout('/api/process', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ...payload, cache_only: true, operation_id: operationId }),
    }, 300000, workbookTimeoutMessage),
  );
}

export async function exportPriceReportWorkbook(payload: Record<string, unknown>): Promise<Blob> {
  const response = await fetch('/api/export_price_report', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    const errorPayload = await response.json().catch(() => ({}));
    throw new Error(String(errorPayload.detail || errorPayload.error || response.statusText));
  }
  return response.blob();
}

export async function inspectProcessedVietmaxFile(savedName: string, phase: 'purchase' | 'sales'): Promise<ProcessedFileStats> {
  return parseJsonResponse<ProcessedFileStats>(
    await fetchWithTimeout('/api/vietmax/processed-file-stats', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, phase }),
    }, 120000),
  );
}

export async function getVietmaxFormatMappingDefaults(profile?: string): Promise<FormatMappingDefaults> {
  const query = profile ? `?profile=${encodeURIComponent(profile)}` : '';
  return parseJsonResponse<FormatMappingDefaults>(await fetch(`/api/vietmax/format-mapping-defaults${query}`));
}

export async function validateFastImportProcessedFile(savedName: string, phase: 'purchase' | 'sales'): Promise<{ valid_rows: number; tk_vat_tu_col: string; ma_kho_col: string }> {
  return parseJsonResponse<{ valid_rows: number; tk_vat_tu_col: string; ma_kho_col: string }>(
    await fetchWithTimeout('/api/vietmax/validate-fast-import-processed', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, phase }),
    }, 120000),
  );
}

export async function getLicenseStatus(): Promise<LicenseStatus> {
  return parseJsonResponse<LicenseStatus>(await fetchWithTimeout(
    '/api/license/status',
    {},
    10000,
    'Kiểm tra license quá 10 giây chưa xong. Hãy kiểm tra kết nối tới license server rồi bấm Kiểm tra lại.',
  ));
}

export async function activateLicense(payload: { license_key: string; server_url?: string; account_id?: string }): Promise<LicenseStatus> {
  return parseJsonResponse<LicenseStatus>(
    await fetchWithTimeout('/api/license/activate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    }, 30000, 'Kích hoạt license quá 30 giây chưa xong. Hãy kiểm tra kết nối tới license server rồi thử lại.'),
  );
}

export async function reloadLicense(): Promise<LicenseStatus> {
  return parseJsonResponse<LicenseStatus>(await fetchWithTimeout(
    '/api/license/reload',
    { method: 'POST' },
    30000,
    'Tải lại license quá 30 giây chưa xong. Hãy kiểm tra kết nối tới license server rồi thử lại.',
  ));
}

export async function getOperationProgress(operationId: string): Promise<OperationProgress> {
  return parseJsonResponse<OperationProgress>(
    await fetchWithTimeout(`/api/progress/${encodeURIComponent(operationId)}`, {}, 15000, 'Không đọc được trạng thái xử lý từ backend.'),
  );
}

export async function createPurchaseReview(savedName: string, comparisonScope: string, wordRules: Record<string, string>, repeatedPhraseRemovals: string[], products: ReviewProduct[], operationId = '', phase: 'purchase' | 'sales' = 'purchase', allowSameCodeSplit = false) {
  const job = await runJsonWorkflowJob('vietmax-review', { saved_name: savedName, phase, comparison_scope: comparisonScope, price_col: 'P', word_rules: wordRules, repeated_phrase_removals: repeatedPhraseRemovals, products, operation_id: operationId, allow_same_code_split: allowSameCodeSplit });
  return job.result as unknown as { products: unknown[]; review_rows: unknown[] };
}

export async function createGenericReview(savedName: string, profile: string, comparisonScope: string, wordRules: Record<string, string>, firstWordRules: Record<string, string>, repeatedPhraseRemovals: string[], products: ReviewProduct[], operationId = '', phase: 'purchase' | 'sales' = 'purchase') {
  const job = await runJsonWorkflowJob('generic-review', { saved_name: savedName, profile, vietmax_phase: phase, comparison_scope: comparisonScope, word_rules: wordRules, first_word_rules: firstWordRules, repeated_phrase_removals: repeatedPhraseRemovals, products, operation_id: operationId });
  return job.result as unknown as { products: unknown[]; review_rows: unknown[] };
}

const vietmaxInvoiceStatusSkipValues = ['Hóa đơn đã bị hủy'];

type VietmaxColumnPayload = Partial<Record<'company_col' | 'mst_col' | 'address_col' | 'product_col' | 'qty_col' | 'price_col' | 'output_col' | 'invoice_status_col' | 'purchase_price_col', string>> & { invoice_status_skip_values?: string[] };

export async function analyzeVietmaxCompanies(savedName: string, phase: 'purchase' | 'sales' = 'purchase', columns: VietmaxColumnPayload = {}) {
  const company_col = columns.company_col || (phase === 'sales' ? 'I' : 'F');
  const mst_col = columns.mst_col || (phase === 'sales' ? 'J' : 'G');
  const address_col = columns.address_col || (phase === 'sales' ? 'K' : 'H');
  const invoiceStatusSkipValues = columns.invoice_status_skip_values || vietmaxInvoiceStatusSkipValues;
  return parseJsonResponse<{ rows_to_process: number; company_count: number; companies: CompanyRow[]; missing_mst_companies?: MissingMstCompanyWarning[]; missing_mst_row_count?: number; manual_code_overrides?: Record<string, string>; product_code_replacements?: Record<string, string>; word_rules?: Record<string, string>; repeated_phrase_removals?: string[]; inventory_pairs?: InventoryPair[]; use_default_inventory_pair?: boolean; default_inventory_pair_id?: string; inventory_pair_rules?: InventoryRule[]; sales_match_rules?: MatchRow[]; vietmax_mua_vao_internal_merges?: ReviewRow[]; vietmax_ban_ra_sales_internal_merges?: ReviewRow[]; include_company_prefix?: boolean; prefix_strategy?: string; prefix_mst_digits?: number; prefix_name_words?: number; prefix_name_chars?: number; prefix_missing_mst_strategy?: string; prefix_strategy_values?: Record<string, Record<string, string>>; processing_groups?: import('./types').ProcessingGroup[]; company_group_assignments?: Record<string, string>; form_mapping_presets?: import('./types').FormMappingPreset[]; columns?: Record<string, unknown> }>(
    await fetch('/api/vietmax/analyze', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, phase, company_col, mst_col, address_col, product_col: columns.product_col || 'M', qty_col: columns.qty_col || 'O', price_col: columns.price_col || 'P', invoice_status_col: columns.invoice_status_col || 'AJ', invoice_status_skip_values: invoiceStatusSkipValues }),
    }),
  );
}

export async function previewVietmaxProductCodes(products: string[], wordRules: Record<string, string>, repeatedPhraseRemovals: string[], phase: 'purchase' | 'sales' = 'purchase', productCodeReplacements: Record<string, string> = {}) {
  return parseJsonResponse<{ codes: Record<string, string> }>(
    await fetch('/api/vietmax/product-preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ phase, products, word_rules: wordRules, repeated_phrase_removals: repeatedPhraseRemovals, product_code_replacements: productCodeReplacements }),
    }),
  );
}


export async function saveVietmaxConfig(payload: Record<string, unknown>) {
  return parseJsonResponse<Record<string, unknown>>(
    await fetch('/api/vietmax/save-config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    }),
  );
}

export async function importVietmaxConfig(phase: 'purchase' | 'sales', payload: Record<string, unknown>) {
  return parseJsonResponse<Record<string, unknown>>(
    await fetch(`/api/vietmax/import-config/${phase}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    }),
  );
}

export async function processVietmaxPurchase(savedName: string, originalName: string, payload: Record<string, unknown>, options: { cacheOnly?: boolean; operationId?: string } = {}): Promise<ProcessResult> {
  const response = await fetchWithTimeout('/api/vietmax/process', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ saved_name: savedName, original_name: originalName, payload, operation_id: options.operationId || '', cache_only: Boolean(options.cacheOnly) }),
  }, 300000, workbookTimeoutMessage);
  if (!response.ok) {
    const errorPayload = await response.json().catch(() => ({}));
    throw new Error(String(errorPayload.detail || errorPayload.error || response.statusText));
  }
  if (options.cacheOnly) {
    const data = await response.json().catch(() => ({}));
    return {
      processedSavedName: String(data.processed_saved_name || response.headers.get('X-Processed-Saved-Name') || ''),
    };
  }
  return { blob: await response.blob(), processedSavedName: response.headers.get('X-Processed-Saved-Name') || '' };
}

export async function downloadCachedFile(savedName: string): Promise<Blob> {
  const response = await fetch(`/api/files/download/${encodeURIComponent(savedName)}`);
  if (!response.ok) {
    const errorPayload = await response.json().catch(() => ({}));
    throw new Error(String(errorPayload.detail || errorPayload.error || response.statusText));
  }
  return response.blob();
}

export type VietmaxFastImportMappingPayload = {
  profile?: string;
  purchaseOriginalSavedName?: string;
  salesOriginalSavedName?: string;
  purchaseFormMappingPresets?: FormMappingPreset[];
  salesFormMappingPresets?: FormMappingPreset[];
  purchaseCompanyGroupAssignments?: Record<string, string>;
  salesCompanyGroupAssignments?: Record<string, string>;
};

export async function createVietmaxFastImportPackage(purchaseSavedName: string, salesSavedName: string, operationId = '', mappingPayload: VietmaxFastImportMappingPayload = {}): Promise<Blob> {
  const job = await runJsonWorkflowJob('fast-export', {
    profile: mappingPayload.profile || 'vietmax',
    purchase_saved_name: purchaseSavedName,
    sales_saved_name: salesSavedName,
    purchase_original_saved_name: mappingPayload.purchaseOriginalSavedName || '',
    sales_original_saved_name: mappingPayload.salesOriginalSavedName || '',
    purchase_form_mapping_presets: mappingPayload.purchaseFormMappingPresets,
    sales_form_mapping_presets: mappingPayload.salesFormMappingPresets,
    purchase_company_group_assignments: mappingPayload.purchaseCompanyGroupAssignments,
    sales_company_group_assignments: mappingPayload.salesCompanyGroupAssignments,
    operation_id: operationId,
  }, true);
  const artifactId = job.result?.artifact?.artifact_id || '';
  if (!artifactId) throw new Error(`Job ${job.operation_id} không trả về workbook FAST.`);
  const response = await fetch(`/api/artifacts/${encodeURIComponent(artifactId)}`);
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(String(payload.detail || payload.error || response.statusText));
  }
  return response.blob();
}

const salesMatchTimeoutMessage = 'Khớp mua/bán quá 5 phút chưa xong nên đã dừng lần chạy này. Vui lòng kiểm tra cấu hình cột bán ra, file mua vào đã xử lý, hoặc bấm Khớp lại sau khi sửa.';

export async function createSalesMatches(salesSavedName: string, purchaseSavedName: string, comparisonScope: string, operationId = '', columns: VietmaxColumnPayload = {}) {
  const job = await runJsonWorkflowJob('sales-match', {
    sales_saved_name: salesSavedName,
    purchase_saved_name: purchaseSavedName,
    comparison_scope: comparisonScope,
    sales_price_col: columns.price_col || 'P',
    purchase_price_col: columns.purchase_price_col || 'P',
    sales_company_col: columns.company_col || 'I',
    sales_mst_col: columns.mst_col || 'J',
    product_col: columns.product_col || 'M',
    qty_col: columns.qty_col || 'O',
    invoice_status_col: columns.invoice_status_col || 'AJ',
    invoice_status_skip_values: columns.invoice_status_skip_values || vietmaxInvoiceStatusSkipValues,
    require_existing_purchase_code: true,
    operation_id: operationId,
  }, true);
  return job.result as unknown as { sales_products: unknown[]; purchase_products: unknown[]; exact_matches: MatchRow[]; matches: MatchRow[]; match_rules?: MatchRow[] };
}

export async function exportMatches(matches: MatchRow[]): Promise<Blob> {
  const response = await fetch('/api/vietmax/export-matches', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ matches }),
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(String(payload.detail || payload.error || response.statusText));
  }
  return response.blob();
}

export async function startInventoryAllocation(payload: { purchaseSavedName: string; salesSavedName: string; salesOriginalName: string; openingFile?: File | null; config: InventoryAllocationConfig }) {
  const form = new FormData();
  form.append('purchase_saved_name', payload.purchaseSavedName);
  form.append('sales_saved_name', payload.salesSavedName);
  form.append('sales_original_name', payload.salesOriginalName);
  form.append('mapping', JSON.stringify(payload.config.mapping));
  form.append('policy', JSON.stringify(payload.config.policy));
  if (payload.openingFile) form.append('opening_file', payload.openingFile);
  return parseJsonResponse<{ analysis_job_id: string }>(await fetch('/api/inventory-allocation/analyze-job', { method: 'POST', body: form }));
}

export async function getInventoryAllocationJob(jobId: string, includeReport = false): Promise<InventoryAllocationJob> {
  const query = includeReport ? '?include_report=true' : '';
  return parseJsonResponse<InventoryAllocationJob>(await fetch(`/api/inventory-allocation/analyze-job/${encodeURIComponent(jobId)}${query}`));
}


export async function createSonPhuongProcessedSales(jobId: string, operationId = ''): Promise<{ processed_sales_saved_name: string; result?: InventoryAllocationJob['result']; rows?: number; columns?: number }> {
  const query = operationId ? `?operation_id=${encodeURIComponent(operationId)}` : '';
  return parseJsonResponse<{ processed_sales_saved_name: string; result?: InventoryAllocationJob['result']; rows?: number; columns?: number }>(
    await fetchWithTimeout(`/api/inventory-allocation/analyze-job/${encodeURIComponent(jobId)}/create-sales-fdi${query}`, { method: 'POST' }, 300000, workbookTimeoutMessage),
  );
}

export async function downloadInventoryAllocationReport(jobId: string, operationId = ''): Promise<Blob> {
  const query = operationId ? `?operation_id=${encodeURIComponent(operationId)}` : '';
  const response = await fetchWithTimeout(
    `/api/inventory-allocation/download/${encodeURIComponent(jobId)}${query}`,
    {},
    300000,
    'Tạo báo cáo Phân kho quá 5 phút chưa xong. Hãy kiểm tra trạng thái lỗi và thử lại.',
  );
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(String(payload.detail || payload.error || response.statusText));
  }
  return response.blob();
}


export async function checkForUpdate(): Promise<UpdateManifest> {
  return parseJsonResponse<UpdateManifest>(await fetch('/api/update/check'));
}

export async function applyUpdate(manifest: UpdateManifest): Promise<{ scheduled: boolean; version: string }> {
  return parseJsonResponse<{ scheduled: boolean; version: string }>(await fetch('/api/update/apply', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ manifest, server_url: manifest.server_url }),
  }));
}
