export type UploadSummary = {
  original_name: string;
  saved_name: string;
  columns: Array<{ letter: string; label: string }>;
  preview: Array<Record<string, string>>;
  invoice_statuses: Array<{ value: string; count: number; skip: boolean }>;
};

export type ProcessResult = {
  blob: Blob;
  processedSavedName: string;
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

export type InventoryAllocationResult = {
  job_id: string;
  filename: string;
  summary?: InventoryAllocationSummary;
  warnings?: string[];
  allocation_count?: number;
  stock_count?: number;
  verification?: Array<{ group: string; check: string; status: string; difference?: number; explanation?: string }>;
};

export type InventoryAllocationJob = {
  status: 'queued' | 'running' | 'complete' | 'error' | string;
  progress?: number;
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
  all_products: Array<{ name: string; count?: number; priceRows?: Array<{ unit?: string; invoiceNo?: string; invoiceDate?: string }> }>;
  process?: boolean;
  pending_process?: boolean;
  committed_prefix?: string;
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
};

export type ReviewRow = {
  confirmed?: boolean;
  code_choice?: 'current' | 'similar' | 'split' | string;
  review_group?: 'dimension_diff' | string;
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
  purchase_product: string;
  purchase_code: string;
  purchase_unit?: string;
  invoice_no?: string;
  invoice_date?: string;
  purchase_company?: string;
  purchase_mst?: string;
  purchase_company_key?: string;
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
  vietmax_allowed: boolean;
  server_url: string;
  account_id: string;
};
