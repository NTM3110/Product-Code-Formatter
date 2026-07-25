export type InvoiceStatusOption = { value: string; count: number; skip: boolean };

export type UploadSummary = {
  original_name: string;
  saved_name: string;
  columns: Array<{ letter: string; label: string }>;
  preview: Array<Record<string, string>>;
  invoice_statuses: InvoiceStatusOption[];
  artifact_id?: string;
  session_id?: string;
};

export type WorkflowArtifact = {
  artifact_id: string;
  saved_name: string;
  original_name: string;
  kind: string;
  signature: string;
  sha256: string;
  size: number;
  valid: boolean;
  created_at: number;
  metadata: Record<string, unknown>;
};

export type WorkflowSession = {
  session_id: string;
  created_at: number;
  updated_at: number;
  artifacts: Record<string, WorkflowArtifact>;
};

export type JobFailure = {
  code: string;
  message: string;
  stage: string;
  field: string;
  details: Record<string, unknown>;
  retryable: boolean;
  operation_id: string;
};

export type WorkflowJob = {
  job_id: string;
  operation_id: string;
  kind: string;
  signature: string;
  status: 'queued' | 'running' | 'succeeded' | 'failed';
  progress: Omit<OperationProgress, 'operation_id' | 'status'>;
  context: Record<string, unknown>;
  result: null | {
    artifact?: WorkflowArtifact;
    processed_saved_name?: string;
    reused?: boolean;
    stats?: Record<string, unknown>;
  };
  error: JobFailure | null;
};

export type EstimateAnalysisSummary = {
  path: string;
  bid_rows: number;
  detail_blocks: number;
  identity_mismatches: number;
  calculation_mismatches: number;
  helper_mismatches: number;
  unclassified_rows: number;
  thvt_rows: number;
  generated_thvt_rows: number;
  thvt_mismatches: number;
  thvt_key_mismatches: number;
  thvt_missing_rows: number;
  thvt_extra_rows: number;
  ok: boolean;
};

export type EstimateAnalysis = {
  summary: EstimateAnalysisSummary;
  warnings: Record<string, unknown[]>;
  sheet_names: string[];
};

export type EstimateSheetInfo = {
  index: number;
  name: string;
  rows: number;
  cols: number;
};

export type EstimateColumnConfig = Record<string, string>;

export type EstimateSheetSelection = {
  bid_sheet_index?: number | null;
  detail_sheet_index?: number | null;
  bid_header_row?: number | null;
  detail_header_row?: number | null;
  bid_columns?: EstimateColumnConfig;
  detail_columns?: EstimateColumnConfig;
};

export type EstimateUploadSummary = {
  original_name: string;
  saved_name: string;
  size: number;
  analysis?: EstimateAnalysis;
  sheet_names?: string[];
  sheets?: EstimateSheetInfo[];
  suggested_sheets?: EstimateSheetSelection;
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
  unit?: 'rows' | 'steps' | 'percent';
};

export type ProcessResult = {
  blob?: Blob;
  processedSavedName: string;
};

export type GenericAnalyzeResult = {
  rows_to_process: number;
  company_count: number;
  companies: CompanyRow[];
  missing_mst_companies?: MissingMstCompanyWarning[];
  missing_mst_row_count?: number;
  original_name?: string;
  saved_name?: string;
  manual_code_overrides?: Record<string, string>;
  product_code_replacements?: Record<string, string>;
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
  prefix_name_words?: number;
  prefix_name_chars?: number;
  prefix_missing_mst_strategy?: string;
  prefix_strategy_values?: Record<string, Record<string, string>>;
  processing_groups?: ProcessingGroup[];
  company_group_assignments?: Record<string, string>;
  form_mapping_presets?: FormMappingPreset[];
  product_review_merges?: unknown[];
  price_range_rules?: Record<string, unknown>;
  price_adjust_all_percent?: number;
  columns?: Record<string, unknown>;
};

export type MissingMstCompanyWarning = {
  company: string;
  count: number;
  rows?: Array<number | string>;
  invoice_nos?: string[];
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
  generic_min_take_quantity?: number | null;
  generic_max_take_quantity?: number | null;
  generic_min_type_count?: number;
  barem_tolerance_percent?: number;
  barem_remainder_max_kg?: number | null;
  ignore_sale_suffix: boolean;
  allow_negative_export: boolean;
  company_profile: string;
  sales_inventory_pairs?: InventoryPair[];
  scenario_count?: number;
  allow_future_purchase_reorder: boolean;
  future_purchase_window_days: number;
};

export type InventoryAllocationConfig = {
  mapping: InventoryAllocationMapping;
  policy: InventoryAllocationPolicy;
  sales_inventory_pairs?: InventoryPair[];
  sales_inventory_pair_rules?: InventoryRule[];
  scenario_count?: number;
};

export type InventoryAllocationSummary = {
  opening_quantity?: number;
  purchase_quantity?: number;
  sales_quantity?: number;
  material_quantity?: number;
  unresolved_material_quantity?: number;
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


export type InventoryAllocationRow = {
  row_number: number;
  variant_code: string;
  product_name: string;
  quantity: number;
  invoice_no?: string;
  invoice_date?: string;
  sale_split_codes?: string;
  material_quantity: number;
  unresolved_material_quantity?: number;
  finished_quantity: number;
  allocation_role?: 'materials' | 'finished_goods' | 'fallback' | string;
  warehouse_code?: string;
  warehouse_account?: string;
  remainder_warehouse_code?: string;
  remainder_warehouse_account?: string;
  negative_warning?: boolean;
  generic_plan_note?: string;
  detail?: string;
  inventory_before_detail?: string;
  inventory_after_detail?: string;
};

export type InventoryAllocationResult = {
  job_id: string;
  filename: string;
  processed_sales_saved_name?: string;
  summary?: InventoryAllocationSummary;
  warnings?: string[];
  allocation_count?: number;
  stock_count?: number;
  verification?: Array<{ group: string; check: string; status: string; difference?: number; explanation?: string }>;
  allocations?: InventoryAllocationRow[];
  missing_barem_report?: Array<{ variant_code: string; product_name?: string; quantity?: number; invoice_no?: string; invoice_date?: string; row_number?: number | string; profile_key?: string; steel_kind?: string; steel_coating?: string; steel_dimension?: string; status?: string; reason?: string }>;
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
  company_id?: string;
  missing_mst?: boolean;
  company: string;
  safe_id: string;
  value: string;
  selected_product_names: string[];
  all_products: Array<{ name: string; count?: number; minPrice?: number | null; maxPrice?: number | null; priceCount?: number; priceRows?: Array<{ price?: number | null; quantity?: number | null; amount?: number | null; excelRow?: string | number; stt?: string; unit?: string; invoiceNo?: string; invoiceDate?: string; name?: string }> }>;
  process?: boolean;
  pending_process?: boolean;
  group_id?: string;
  pending_group_id?: string;
  committed_prefix?: string;
  default_prefix?: string;
  prefix_strategies?: Record<string, string>;
};

export type ProcessingForm = {
  id: string;
  label: string;
  type: 'builtin' | 'template_mapping' | string;
  scope?: 'purchase' | 'sales' | 'both' | string;
  enabled?: boolean;
  builtin_exporter?: string;
  group_id?: string;
  input_phase?: 'purchase' | 'sales' | 'both' | string;
  template_saved_name?: string;
  template_original_name?: string;
  sheet?: string;
  header_row?: number;
  data_start_row?: number;
  copy_style_row?: number;
  output_columns?: FormColumn[];
  input_columns?: FormColumn[];
  output_preview?: Array<Record<string, string>>;
  system_generated?: boolean;
  mappings?: FormMappingRule[];
};

export type ProcessingGroup = {
  id: string;
  label: string;
  builtin?: boolean;
  uses_product_code?: boolean;
  forms?: ProcessingForm[];
};

export type FormMappingRule = {
  target_col: string;
  target_label?: string;
  source_type: 'source_column' | 'constant' | 'empty' | 'text_template' | string;
  source_phase?: 'purchase' | 'sales' | 'both' | string;
  source_col?: string;
  source_label?: string;
  condition_source_col?: string;
  fallback_source_col?: string;
  fallback_delimiter?: string;
  transform?: string;
  transform_rules?: Array<{ source_col?: string; match_type?: 'starts_with' | 'equals' | 'not_equals' | 'contains' | 'not_contains' | 'ends_with' | 'regex' | 'blank' | 'not_blank' | 'gt' | 'gte' | 'lt' | 'lte' | 'default' | string; value?: string; result?: string }>;
  value?: string;
};

export type FormMappingPreset = ProcessingForm & {
  id: string;
  label: string;
};

export type FormColumn = {
  letter: string;
  label: string;
  header?: string;
};

export type FormatMappingDefaults = {
  source_columns: {
    purchase: FormColumn[];
    sales: FormColumn[];
  };
  form_mapping_presets: FormMappingPreset[];
};

export type InventoryPair = {
  id: string;
  ma_kho: string;
  tk_vat_tu: string;
  role?: 'materials' | 'finished_goods' | 'fallback' | string;
  label?: string;
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


export type UpdateManifest = {
  current_version: string;
  version: string;
  available: boolean;
  server_url: string;
  filename?: string;
  download_url?: string;
  sha256?: string;
  size?: number;
  notes?: string;
  current_notes?: string;
};
