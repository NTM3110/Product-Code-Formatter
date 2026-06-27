export type UploadSummary = {
  original_name: string;
  saved_name: string;
  columns: Array<{ letter: string; label: string }>;
  preview: Array<Record<string, string>>;
  invoice_statuses: Array<{ value: string; count: number; skip: boolean }>;
};

export type ProcessedFileStats = {
  phase: 'purchase' | 'sales' | string;
  company_count: number;
  processed_company_count: number;
  product_row_count: number;
  processed_product_row_count: number;
  code_col?: string;
  company_col?: string;
  mst_col?: string;
  product_col?: string;
};

export type OperationProgress = {
  operation_id: string;
  status: string;
  done: number;
  total: number;
  percent: number;
  label: string;
};

export type ProcessResult = {
  blob?: Blob;
  processedSavedName: string;
};

export type GenericAnalyzeResult = {
  rows_to_process: number;
  company_count: number;
  companies: CompanyRow[];
  original_name?: string;
  saved_name?: string;
  manual_code_overrides?: Record<string, string>;
  word_rules?: Record<string, string>;
  first_word_rules?: Record<string, string>;
  repeated_phrase_removals?: string[];
  inventory_pairs?: InventoryPair[];
  use_default_inventory_pair?: boolean;
  default_inventory_pair_id?: string;
  inventory_pair_rules?: InventoryRule[];
  include_company_prefix?: boolean;
  prefix_strategy?: string;
  prefix_mst_digits?: number;
  prefix_strategy_values?: Record<string, Record<string, string>>;
  product_review_merges?: unknown[];
  price_range_rules?: Record<string, unknown>;
  price_adjust_all_percent?: number;
  columns?: Record<string, unknown>;
};

export type InventoryAllocationMappingSection = {
  sheet: string;
  header_row: number;
  data_start_row: number;
  invoice_col: string;
  date_col: string;
  code_col: string;
  product_col: string;
  qty_col: string;
  price_col: string;
};

export type InventoryAllocationMapping = {
  purchase: InventoryAllocationMappingSection;
  sales: InventoryAllocationMappingSection;
  opening: InventoryAllocationMappingSection;
};

export type InventoryAllocationPolicy = {
  max_loss_percent: number | null;
  max_profit_percent: number | null;
  ignore_sale_suffix: boolean;
  allow_negative_export: boolean;
  company_profile: string;
  allow_future_purchase_reorder: boolean;
  future_purchase_window_days: number;
};

export type InventoryAllocationConfig = {
  mapping: InventoryAllocationMapping;
  policy: InventoryAllocationPolicy;
};

export type InventoryAllocationSummary = {
  opening_quantity?: number;
  purchase_quantity?: number;
  sales_quantity?: number;
  material_quantity?: number;
  finished_quantity?: number;
  sale_only_code_count?: number;
  range_rejected_lines?: number;
};

export type InventoryReportWarehouse = {
  warehouse_code: string;
  warehouse_name?: string;
  account?: string;
};

export type InventorySalesSummaryRow = {
  key: string;
  index?: number;
  warehouse_code: string;
  variant_code: string;
  product_name: string;
  unit_name: string;
  quantity: number;
  cost_amount: number;
  sale_amount: number;
  profit_amount: number;
  margin_percent?: number | null;
  tax_amount: number;
  total_amount: number;
  row_count?: number;
  margin_sale_amount?: number;
  cost_missing_count?: number;
};

export type InventorySalesDetailRow = {
  key: string;
  summary_key: string;
  row_type?: string;
  warehouse_code: string;
  invoice_date?: string;
  invoice_date_iso?: string;
  invoice_no?: string;
  customer?: string;
  tax_code?: string;
  variant_code?: string;
  product_name?: string;
  unit_name?: string;
  quantity?: number;
  sale_price?: number;
  sale_amount?: number;
  cost_price?: number;
  cost_amount?: number;
  profit_amount?: number;
  tax_rate?: number;
  tax_amount?: number;
  total_amount?: number;
  cost_missing?: boolean;
};

export type InventorySummaryRow = {
  key: string;
  index?: number;
  warehouse_code: string;
  warehouse_name?: string;
  account?: string;
  variant_code: string;
  product_name: string;
  unit_name: string;
  opening_qty: number;
  opening_amount: number;
  in_qty: number;
  in_amount: number;
  out_qty: number;
  out_amount: number;
  ending_qty: number;
  ending_amount: number;
  row_count?: number;
};

export type InventoryLedgerDetailRow = {
  key: string;
  summary_key: string;
  row_type?: string;
  warehouse_code: string;
  variant_code?: string;
  product_name?: string;
  unit_name?: string;
  date?: string;
  date_iso?: string;
  doc_no?: string;
  customer?: string;
  description?: string;
  account?: string;
  unit_price?: number | string;
  sale_unit_price?: number | string;
  sale_amount?: number | string;
  tax_rate?: number | string;
  tax_amount?: number | string;
  total_amount?: number | string;
  cost_missing?: boolean;
  qty_in?: number;
  amount_in?: number;
  qty_out?: number;
  amount_out?: number;
  running_qty?: number;
  running_amount?: number;
  logic_note?: string;
};

export type InventoryAllocationReportView = {
  date_range?: { from?: string; to?: string };
  warehouses?: InventoryReportWarehouse[];
  sales_summary_rows?: InventorySalesSummaryRow[];
  sales_detail_rows?: InventorySalesDetailRow[];
  inventory_summary_rows?: InventorySummaryRow[];
  ledger_detail_rows?: InventoryLedgerDetailRow[];
};

export type InventoryAllocationResult = {
  job_id: string;
  filename: string;
  summary?: InventoryAllocationSummary;
  warnings?: string[];
  allocation_count?: number;
  stock_count?: number;
  verification?: Array<{ group: string; check: string; status: string; difference?: number; explanation?: string }>;
  report_view?: InventoryAllocationReportView;
};

export type InventoryAllocationJob = {
  status: 'queued' | 'running' | 'complete' | 'error' | string;
  progress?: number;
  done?: number;
  total?: number;
  label?: string;
  error?: string;
  result?: InventoryAllocationResult;
};

export type CompanyRow = {
  mst: string;
  company: string;
  safe_id: string;
  value: string;
  selected_product_names: string[];
  all_products: Array<{ name: string; count?: number; minPrice?: number | null; maxPrice?: number | null; priceCount?: number; priceRows?: Array<{ price?: number | null; quantity?: number | null; amount?: number | null; excelRow?: string | number; stt?: string; unit?: string; invoiceNo?: string; invoiceDate?: string; name?: string }> }>;
  process?: boolean;
  pending_process?: boolean;
  committed_prefix?: string;
  default_prefix?: string;
  prefix_strategies?: Record<string, string>;
};

export type InventoryPair = {
  id: string;
  ma_kho: string;
  tk_vat_tu: string;
};

export type InventoryRule = {
  source_col: string;
  operator: 'contains' | 'equals' | string;
  value: string;
  pair_id: string;
  enabled?: boolean;
  priority?: number;
};

export type ReviewRow = {
  confirmed?: boolean;
  code_choice?: 'current' | 'similar' | 'split' | string;
  review_group?: 'dimension_diff' | 'similar_form' | string;
  comparison_scope?: 'all_companies' | 'same_company' | string;
  dimension_only?: boolean;
  product?: string;
  code?: string;
  unit?: string;
  company?: string;
  mst?: string;
  company_key?: string;
  similar_product?: string;
  similar_code?: string;
  similar_unit?: string;
  similar_company?: string;
  similar_mst?: string;
  similar_company_key?: string;
  diff_html?: string;
  similarity?: number | string;
  invoice_no?: string;
  invoice_date?: string;
  similar_invoice_no?: string;
  similar_invoice_date?: string;
  product_key?: string;
  similar_product_key?: string;
  split_code?: string;
  similar_split_code?: string;
  review_type?: string;
};

export type ReviewProduct = {
  purchase_product?: string;
  purchase_code?: string;
  purchase_unit?: string;
  sales_product?: string;
  sales_code?: string;
  sales_unit?: string;
  invoice_no?: string;
  invoice_date?: string;
  purchase_company?: string;
  purchase_mst?: string;
  purchase_company_key?: string;
  sales_company?: string;
  sales_mst?: string;
  sales_company_key?: string;
  product_key?: string;
  company_index?: number;
  product_index?: number;
};

export type MatchRow = {
  confirmed?: boolean;
  sales_product?: string;
  sales_unit?: string;
  sales_price?: string;
  invoice_no?: string;
  invoice_date?: string;
  purchase_code?: string;
  purchase_product?: string;
  purchase_unit?: string;
  purchase_price?: string;
  purchase_company?: string;
  purchase_mst?: string;
  sales_company?: string;
  sales_mst?: string;
  warning?: string;
  unit_warning?: string;
  conversion_formula?: string;
  conversion_mode?: string;
  diff_html?: string;
  similarity?: number;
  score?: number;
};

export type StageKey = 'upload' | 'purchase-review' | 'sales-match' | 'finalize';

export type LicenseStatus = {
  activated: boolean;
  status: string;
  allowed_profiles: string[];
  allowed_companies: string[];
  supported_profiles: string[];
  product_code: string;
  application: string;
  vietmax_allowed: boolean;
  server_url: string;
  account_id: string;
};
