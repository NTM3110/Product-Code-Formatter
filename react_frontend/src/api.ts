import type { CompanyRow, EstimateAnalysis, EstimateSheetSelection, EstimateUploadSummary, GenericAnalyzeResult, InventoryAllocationConfig, InventoryAllocationJob, InventoryPair, InventoryRule, InvoiceStatusOption, LicenseStatus, MissingMstCompanyWarning, MatchRow, OperationProgress, ProcessResult, ProcessedFileStats, ReviewProduct, ReviewRow, UploadSummary } from './types';

async function parseJsonResponse<T>(response: Response): Promise<T> {
  if (response.ok) {
    return response.json() as Promise<T>;
  }
  const payload = await response.json().catch(() => ({}));
  throw new Error(String(payload.detail || payload.error || response.statusText));
}

export async function uploadExcel(file: File): Promise<UploadSummary> {
  const form = new FormData();
  form.append('file', file);
  return parseJsonResponse<UploadSummary>(await fetch('/api/files/upload', { method: 'POST', body: form }));
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

export async function previewGenericProductCodes(payload: { profile: string; products: string[]; word_rules?: Record<string, string>; first_word_rules?: Record<string, string>; repeated_phrase_removals?: string[] }) {
  return parseJsonResponse<{ codes: Record<string, string> }>(
    await fetch('/api/product-preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    }),
  );
}

export async function processGenericWorkbook(payload: Record<string, unknown>): Promise<Blob> {
  const response = await fetch('/api/process', {
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
    await fetch('/api/vietmax/processed-file-stats', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, phase }),
    }),
  );
}

export async function validateFastImportProcessedFile(savedName: string, phase: 'purchase' | 'sales'): Promise<{ valid_rows: number; tk_vat_tu_col: string; ma_kho_col: string }> {
  return parseJsonResponse<{ valid_rows: number; tk_vat_tu_col: string; ma_kho_col: string }>(
    await fetch('/api/vietmax/validate-fast-import-processed', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, phase }),
    }),
  );
}

export async function getLicenseStatus(): Promise<LicenseStatus> {
  return parseJsonResponse<LicenseStatus>(await fetch('/api/license/status'));
}

export async function activateLicense(payload: { license_key: string; server_url?: string; account_id?: string }): Promise<LicenseStatus> {
  return parseJsonResponse<LicenseStatus>(
    await fetch('/api/license/activate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    }),
  );
}

export async function reloadLicense(): Promise<LicenseStatus> {
  return parseJsonResponse<LicenseStatus>(await fetch('/api/license/reload', { method: 'POST' }));
}

export async function getOperationProgress(operationId: string): Promise<OperationProgress> {
  return parseJsonResponse<OperationProgress>(await fetch(`/api/progress/${encodeURIComponent(operationId)}`));
}

export async function createPurchaseReview(savedName: string, comparisonScope: string, wordRules: Record<string, string>, repeatedPhraseRemovals: string[], products: ReviewProduct[], operationId = '', phase: 'purchase' | 'sales' = 'purchase', allowSameCodeSplit = false) {
  return parseJsonResponse<{ products: unknown[]; review_rows: unknown[] }>(
    await fetch('/api/vietmax/review', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, phase, comparison_scope: comparisonScope, price_col: 'P', word_rules: wordRules, repeated_phrase_removals: repeatedPhraseRemovals, products, operation_id: operationId, allow_same_code_split: allowSameCodeSplit }),
    }),
  );
}

export async function createGenericReview(savedName: string, profile: string, comparisonScope: string, wordRules: Record<string, string>, firstWordRules: Record<string, string>, repeatedPhraseRemovals: string[], products: ReviewProduct[], operationId = '') {
  return parseJsonResponse<{ products: unknown[]; review_rows: unknown[] }>(
    await fetch('/api/review', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, profile, comparison_scope: comparisonScope, word_rules: wordRules, first_word_rules: firstWordRules, repeated_phrase_removals: repeatedPhraseRemovals, products, operation_id: operationId }),
    }),
  );
}

const vietmaxInvoiceStatusSkipValues = ['Hóa đơn đã bị hủy'];

type VietmaxColumnPayload = Partial<Record<'company_col' | 'mst_col' | 'address_col' | 'product_col' | 'qty_col' | 'price_col' | 'output_col' | 'invoice_status_col' | 'purchase_price_col', string>> & { invoice_status_skip_values?: string[] };

export async function analyzeVietmaxCompanies(savedName: string, phase: 'purchase' | 'sales' = 'purchase', columns: VietmaxColumnPayload = {}) {
  const company_col = columns.company_col || (phase === 'sales' ? 'I' : 'F');
  const mst_col = columns.mst_col || (phase === 'sales' ? 'J' : 'G');
  const address_col = columns.address_col || (phase === 'sales' ? 'K' : 'H');
  const invoiceStatusSkipValues = columns.invoice_status_skip_values || vietmaxInvoiceStatusSkipValues;
  return parseJsonResponse<{ rows_to_process: number; company_count: number; companies: CompanyRow[]; missing_mst_companies?: MissingMstCompanyWarning[]; missing_mst_row_count?: number; manual_code_overrides?: Record<string, string>; word_rules?: Record<string, string>; repeated_phrase_removals?: string[]; inventory_pairs?: InventoryPair[]; use_default_inventory_pair?: boolean; default_inventory_pair_id?: string; inventory_pair_rules?: InventoryRule[]; sales_match_rules?: MatchRow[]; vietmax_mua_vao_internal_merges?: ReviewRow[]; vietmax_ban_ra_sales_internal_merges?: ReviewRow[]; include_company_prefix?: boolean; prefix_strategy?: string; prefix_mst_digits?: number; prefix_name_words?: number; prefix_name_chars?: number; prefix_missing_mst_strategy?: string; prefix_strategy_values?: Record<string, Record<string, string>>; columns?: Record<string, unknown> }>(
    await fetch('/api/vietmax/analyze', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, phase, company_col, mst_col, address_col, product_col: columns.product_col || 'M', qty_col: columns.qty_col || 'O', price_col: columns.price_col || 'P', invoice_status_col: columns.invoice_status_col || 'AJ', invoice_status_skip_values: invoiceStatusSkipValues }),
    }),
  );
}

export async function previewVietmaxProductCodes(products: string[], wordRules: Record<string, string>, repeatedPhraseRemovals: string[], phase: 'purchase' | 'sales' = 'purchase') {
  return parseJsonResponse<{ codes: Record<string, string> }>(
    await fetch('/api/vietmax/product-preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ phase, products, word_rules: wordRules, repeated_phrase_removals: repeatedPhraseRemovals }),
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
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), 300000);
  let response: Response;
  try {
    response = await fetch('/api/vietmax/process', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ saved_name: savedName, original_name: originalName, payload, operation_id: options.operationId || '', cache_only: Boolean(options.cacheOnly) }),
      signal: controller.signal,
    });
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') {
      throw new Error('Tạo file đã xử lý quá 5 phút chưa xong. Vui lòng thử lại hoặc kiểm tra log trên máy này.');
    }
    throw error;
  } finally {
    window.clearTimeout(timeout);
  }
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

export async function createVietmaxFastImportPackage(purchaseSavedName: string, salesSavedName: string, operationId = ''): Promise<Blob> {
  const response = await fetch('/api/vietmax/fast-import-package', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ purchase_saved_name: purchaseSavedName, sales_saved_name: salesSavedName, operation_id: operationId }),
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(String(payload.detail || payload.error || response.statusText));
  }
  return response.blob();
}

export async function createSalesMatches(salesSavedName: string, purchaseSavedName: string, comparisonScope: string, operationId = '', columns: VietmaxColumnPayload = {}) {
  return parseJsonResponse<{ sales_products: unknown[]; purchase_products: unknown[]; exact_matches: MatchRow[]; matches: MatchRow[]; match_rules?: MatchRow[] }>(
    await fetch('/api/vietmax/sales-match', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
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
      }),
    }),
  );
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

export async function getInventoryAllocationJob(jobId: string): Promise<InventoryAllocationJob> {
  return parseJsonResponse<InventoryAllocationJob>(await fetch(`/api/inventory-allocation/analyze-job/${encodeURIComponent(jobId)}`));
}

export async function downloadInventoryAllocationReport(jobId: string): Promise<Blob> {
  const response = await fetch(`/api/inventory-allocation/download/${encodeURIComponent(jobId)}`);
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(String(payload.detail || payload.error || response.statusText));
  }
  return response.blob();
}
