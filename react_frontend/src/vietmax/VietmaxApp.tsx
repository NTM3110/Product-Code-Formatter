import { Fragment, useEffect, useMemo, useState } from 'react';
import { activateLicense, applyUpdate, analyzeGenericWorkbook, checkForUpdate, analyzeVietmaxCompanies, createGenericReview, createPurchaseReview, createSalesMatches, createSonPhuongProcessedSales, createVietmaxFastImportPackage, downloadCachedFile, downloadInventoryAllocationReport, exportMatches, exportPriceReportWorkbook, fetchInvoiceStatuses, getAppConfig, getInventoryAllocationJob, getLicenseStatus, getOperationProgress, getVietmaxFormatMappingDefaults, getWorkflowSession, importProductCodeReplacements, importVietmaxConfig, inspectProcessedVietmaxFile, previewGenericProductCodes, previewVietmaxProductCodes, processGenericWorkbook, saveVietmaxConfig, reloadLicense, startInventoryAllocation, startWorkflowProcessJob, uploadExcel, uploadFormTemplate, validateFastImportProcessedFile, waitForWorkflowJob } from '../api';
import { useRef } from 'react';
import type { CompanyRow, FormatMappingDefaults, FormColumn, FormMappingPreset, InventoryAllocationConfig, InventoryAllocationJob, InventoryAllocationResult, InventoryPair, InventoryRule, InvoiceStatusOption, LicenseStatus, MissingMstCompanyWarning, MatchRow, OperationProgress, ProcessedFileStats, ProcessingForm, ProcessingGroup, ReviewProduct, ReviewRow, UploadSummary, UpdateManifest, WorkflowJob } from '../types';
import { EstimateExtractorWorkflow, type EstimateExtractorWorkflowHandle } from '../estimate/EstimateExtractorWorkflow';
import { InventoryAllocationExportStage, InventoryAllocationOverviewStage, InventoryAllocationReportStage, InventoryAllocationStage, SonPhuongAllocationReviewStage, SonPhuongSalesPairEditor } from './InventoryAllocationStage';
import { StageNavigation } from './StageNavigation';
import { FastImportExportStage, GenericMappingStage, LoadingStage, MappingStage, PlaceholderStage, ProcessStage, SalesEntryStage, UploadStage, type GenericColumns } from './basicStages';
import { hasVietmaxPurchaseMatch, hoGuomFormatterStages, isGenericProfileKey, isStageId, profiles, stagesForProfile, usesTwoPhaseFrame, type PrefixPresetStrategy, type ProfileKey, type StageDefinition, type StageId, type StagePhase } from './workflowStages';

type PrefixStrategyValues = Record<PrefixPresetStrategy, Record<string, string>>;
type InventoryConfigScope = 'purchase' | 'sales' | 'generic';
type FormatScope = 'purchase' | 'sales' | 'both';
type HoGuomMode = 'estimate' | 'formatter';
const MATERIALS_GROUP_ID = 'materials';
const SERVICE_GROUP_ID = 'services';
const IGNORED_GROUP_ID = 'ignored';
const FORMAT_SCOPE_LABELS: Record<FormatScope, string> = {
  purchase: 'Mua vào',
  sales: 'Bán ra',
  both: 'Mua vào + Bán ra',
};
const MAPPING_RULE_OPTIONS = [
  { value: 'source_column', label: 'Cột nguồn' },
  { value: 'if_rules', label: 'IF' },
  { value: 'text', label: 'Giá trị cố định' },
  { value: 'empty', label: 'Để trống' },
];
type MappingTransformRule = NonNullable<NonNullable<FormMappingPreset['mappings']>[number]['transform_rules']>[number];
const DEFAULT_REVENUE_ACCOUNT_RULES: MappingTransformRule[] = [
  { match_type: 'starts_with', value: '152', result: '5111' },
  { match_type: 'starts_with', value: '155', result: '5112' },
  { match_type: 'default', value: '', result: '5112' },
];
const DEFAULT_MATERIAL_TYPE_RULES: MappingTransformRule[] = [
  { match_type: 'starts_with', value: '152', result: '21' },
  { match_type: 'default', value: '', result: '51' },
];
const TRANSFORM_RULE_MATCH_OPTIONS = [
  { value: 'starts_with', label: 'Bắt đầu bằng', needsValue: true },
  { value: 'equals', label: 'Bằng', needsValue: true },
  { value: 'not_equals', label: 'Khác', needsValue: true },
  { value: 'contains', label: 'Có chứa', needsValue: true },
  { value: 'not_contains', label: 'Không chứa', needsValue: true },
  { value: 'ends_with', label: 'Kết thúc bằng', needsValue: true },
  { value: 'regex', label: 'Regex', needsValue: true },
  { value: 'blank', label: 'Rỗng', needsValue: false },
  { value: 'not_blank', label: 'Không rỗng', needsValue: false },
  { value: 'gt', label: 'Số >', needsValue: true },
  { value: 'gte', label: 'Số >=', needsValue: true },
  { value: 'lt', label: 'Số <', needsValue: true },
  { value: 'lte', label: 'Số <=', needsValue: true },
  { value: 'default', label: 'Mặc định', needsValue: false },
];

const fallbackFdiSourceColumns: FormColumn[] = [
  ['A', 'Mẫu HĐ'],
  ['B', 'Số seri'],
  ['C', 'Số chứng từ'],
  ['D', 'Ngày chứng từ'],
  ['F', 'Người bán'],
  ['G', 'MST người bán'],
  ['H', 'Địa chỉ người bán'],
  ['I', 'Người mua'],
  ['J', 'MST người mua'],
  ['K', 'Địa chỉ người mua'],
  ['L', 'Mã VT'],
  ['M', 'Tên hàng'],
  ['N', 'ĐVT'],
  ['O', 'Số lượng'],
  ['P', 'Đơn giá'],
  ['R', 'Thuế'],
  ['V', 'Tiền hàng'],
  ['W', 'Tiền thuế'],
  ['AQ', 'Mã ngoại tệ'],
  ['AR', 'Tỷ giá'],
  ['AS', 'TK vật tư'],
  ['AT', 'Mã kho'],
  ['AU', 'Mã khách hàng'],
].map(([letter, header]) => ({ letter, header, label: `${letter}. ${header}` }));

function outputColumns(labels: string[]): FormColumn[] {
  return labels.map((header, index) => ({ letter: indexToColumnLetter(index), header, label: `${indexToColumnLetter(index)}. ${header}` }));
}

function indexToColumnLetter(index: number) {
  let value = index + 1;
  let result = '';
  while (value > 0) {
    const mod = (value - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    value = Math.floor((value - mod) / 26);
  }
  return result;
}

function columnLetterToIndex(letter: string | undefined) {
  const normalized = String(letter || '').trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) return Number.MAX_SAFE_INTEGER;
  return normalized.split('').reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}

function mappingRule(target_col: string, source_col = '', source_phase: 'purchase' | 'sales' | 'both' = 'purchase', source_type = 'source_column', extra: Partial<NonNullable<FormMappingPreset['mappings']>[number]> = {}) {
  return { target_col, source_col, source_phase, source_type, ...extra };
}

function isTextMappingRule(rule: Partial<NonNullable<FormMappingPreset['mappings']>[number]>) {
  return rule.source_type === 'constant' || rule.source_type === 'text_template' || rule.source_type === 'text';
}

function textSourceTypeForValue(value: string | undefined) {
  return /\{[A-Za-z]{1,3}\}/.test(String(value || '')) ? 'text_template' : 'constant';
}

function mappingRuleControlValue(rule: Partial<NonNullable<FormMappingPreset['mappings']>[number]>) {
  if (isTextMappingRule(rule)) return 'text';
  if (rule.source_type === 'if_rules' || rule.source_type === 'if_empty_prefix_before_dot' || rule.source_type === 'conditional_rules' || supportsCustomTransformRules(rule.transform)) return 'if_rules';
  return rule.source_type || 'source_column';
}

function defaultConditionRules(rule: Partial<NonNullable<FormMappingPreset['mappings']>[number]>, defaultSourceCol = ''): MappingTransformRule[] {
  if (rule.transform_rules?.length) return rule.transform_rules.map((item) => ({ ...item, source_col: item.source_col || defaultSourceCol }));
  const defaults = defaultTransformRules(rule.transform);
  return defaults.length ? defaults.map((item) => ({ source_col: defaultSourceCol, ...item })) : [{ source_col: defaultSourceCol, match_type: 'blank', value: '', result: '' }];
}

function defaultTransformRules(transform: string | undefined): MappingTransformRule[] {
  if (transform === 'revenue_account_from_inventory_account') return DEFAULT_REVENUE_ACCOUNT_RULES.map((rule) => ({ ...rule }));
  if (transform === 'material_type_from_inventory_account') return DEFAULT_MATERIAL_TYPE_RULES.map((rule) => ({ ...rule }));
  return [];
}

function supportsCustomTransformRules(transform: string | undefined) {
  return transform === 'revenue_account_from_inventory_account' || transform === 'material_type_from_inventory_account';
}

function normalizeTransformRules(raw: unknown, transform: string | undefined, defaultSourceCol = ''): MappingTransformRule[] | undefined {
  if (!supportsCustomTransformRules(transform) && (!Array.isArray(raw) || !raw.length)) return undefined;
  const rows = Array.isArray(raw) && raw.length ? raw : defaultTransformRules(transform);
  const normalizedRules: MappingTransformRule[] = [];
  rows.forEach((item) => {
    const rule = (item || {}) as MappingTransformRule;
    const source_col = String(rule.source_col || defaultSourceCol || '').trim().toUpperCase();
    const match_type = TRANSFORM_RULE_MATCH_OPTIONS.some((option) => option.value === rule.match_type) ? String(rule.match_type) : 'starts_with';
    const result = String(rule.result ?? '').trim();
    const needsValue = TRANSFORM_RULE_MATCH_OPTIONS.find((option) => option.value === match_type)?.needsValue !== false;
    const value = needsValue ? String(rule.value ?? '').trim() : '';
    if (match_type !== 'default' && !source_col) return;
    if (!result || (needsValue && !value)) return;
    normalizedRules.push({ source_col, match_type, value, result });
  });
  return normalizedRules;
}

function mappingTargetOrder(rule: NonNullable<FormMappingPreset['mappings']>[number], outputColumns: FormColumn[] = []) {
  const outputIndex = outputColumns.findIndex((column) => column.letter === rule.target_col);
  return outputIndex >= 0 ? outputIndex + 1 : columnLetterToIndex(rule.target_col);
}

function sortedMappingsForDisplay(mappings: NonNullable<FormMappingPreset['mappings']> = [], outputColumns: FormColumn[] = []) {
  return mappings
    .map((rule, originalIndex) => ({ rule, originalIndex }))
    .sort((left, right) => {
      const byTarget = mappingTargetOrder(left.rule, outputColumns) - mappingTargetOrder(right.rule, outputColumns);
      return byTarget || left.originalIndex - right.originalIndex;
    });
}

function normalizeMappingRuleForSave(rule: NonNullable<FormMappingPreset['mappings']>[number]) {
  const target_col = String(rule.target_col || '').trim().toUpperCase();
  if (!target_col) return null;
  const controlValue = mappingRuleControlValue(rule);
  const source_type = String(controlValue || rule.source_type || 'source_column');
  const normalized = {
    ...rule,
    target_col,
    source_phase: rule.source_phase || 'purchase',
    source_col: String(rule.source_col || '').trim().toUpperCase(),
    condition_source_col: String(rule.condition_source_col || rule.fallback_source_col || '').trim().toUpperCase(),
    fallback_source_col: String(rule.fallback_source_col || '').trim().toUpperCase(),
    fallback_delimiter: rule.fallback_delimiter || '.',
  };
  const transform_rules = normalizeTransformRules(rule.transform_rules, supportsCustomTransformRules(normalized.transform) ? normalized.transform : undefined, normalized.condition_source_col || normalized.source_col);
  if (source_type === 'empty') {
    return { ...normalized, transform: undefined, transform_rules: undefined, source_type: 'empty', source_col: '', condition_source_col: '', fallback_source_col: '' };
  }
  if (isTextMappingRule(normalized)) {
    const value = String(normalized.value ?? '');
    if (!value) return null;
    return { ...normalized, transform: undefined, transform_rules: undefined, source_type: textSourceTypeForValue(value), source_col: '', condition_source_col: '', fallback_source_col: '', value };
  }
  if (source_type === 'if_empty_prefix_before_dot') {
    if (!normalized.source_col && !normalized.fallback_source_col) return null;
    return { ...normalized, transform: undefined, transform_rules: undefined, source_type: 'if_empty_prefix_before_dot' };
  }
  if (source_type === 'if_rules') {
    if (!(transform_rules || []).some((item) => item.source_col || item.match_type === 'default')) return null;
    return { ...normalized, transform: undefined, transform_rules: transform_rules || [], source_type: 'if_rules', source_col: '', condition_source_col: '', fallback_source_col: '' };
  }
  if (source_type === 'conditional_rules') {
    if (!normalized.source_col) return null;
    return { ...normalized, transform: undefined, transform_rules: transform_rules || [], source_type: 'conditional_rules', condition_source_col: '', fallback_source_col: '' };
  }
  if (!normalized.source_col) return null;
  return { ...normalized, transform_rules: undefined, source_type: 'source_column', condition_source_col: '', fallback_source_col: '' };
}

function normalizeFormForSave(form: FormMappingPreset): FormMappingPreset {
  const mappings = (form.mappings || [])
    .map(normalizeMappingRuleForSave)
    .filter((rule): rule is NonNullable<ReturnType<typeof normalizeMappingRuleForSave>> => Boolean(rule))
    .sort((left, right) => mappingTargetOrder(left, form.output_columns || []) - mappingTargetOrder(right, form.output_columns || []));
  return { ...form, mappings };
}

function normalizeFormsForSave(forms: FormMappingPreset[] = []) {
  return forms.map(normalizeFormForSave);
}

function fallbackDefaultFormMappingPresets(phase: 'purchase' | 'sales' | 'all' = 'all'): FormMappingPreset[] {
  const purchase: FormMappingPreset = {
    id: 'fast_hoadonmuahang',
    label: 'Hoadonmuahang',
    scope: 'purchase',
    type: 'builtin',
    enabled: true,
    builtin_exporter: 'fast_hoadonmuahang',
    group_id: MATERIALS_GROUP_ID,
    input_phase: 'purchase',
    sheet: 'Hoadonmuahang',
    output_columns: outputColumns(['Mã khách hàng (ma_kh)', 'Người mua hàng (ong_ba)', 'Diễn giải (dien_giai)', 'Quyển sổ (ma_qs)', 'Mã nx (Tk có) (ma_nx)', 'Số chứng từ (so_ct)', 'Ngày chứng từ (ngay_ct)', 'Mã ngoại tệ (ma_nt)', 'Tỷ giá (ty_gia)', 'Mã kho  (ma_kho)', 'Mã vật tư (ma_vt)', 'Số lượng:Q (so_luong)', 'Giá mua chưa thuế:P0 (gia0)', 'Tiền mua:N0 (tien0)', 'Tk vật tư (tk_vt)', 'Mã dự án (ma_vv_i)', 'Mã ĐVCS (ma_dvcs)', 'Số HĐ thuế (so_ct0)', 'Ngày HĐ (ngay_ct0)', 'Mẫu HĐ (kh_mau_hd)', 'Số seri (so_seri0)', 'Mã thuế (ma_thue)']),
    mappings: [
      mappingRule('A', 'G', 'purchase', 'source_column', { transform: 'mst_or_prefix' }),
      mappingRule('C', '', 'purchase', 'text_template', { value: 'Mua hàng nhập kho HD{C}' }),
      mappingRule('E', '', 'purchase', 'constant', { value: '331' }),
      mappingRule('F', 'C', 'purchase'),
      mappingRule('G', 'D', 'purchase'),
      mappingRule('H', 'AQ', 'purchase'),
      mappingRule('I', 'AR', 'purchase'),
      mappingRule('J', 'AT', 'purchase'),
      mappingRule('K', 'L', 'purchase'),
      mappingRule('L', 'O', 'purchase'),
      mappingRule('M', 'P', 'purchase'),
      mappingRule('N', 'V', 'purchase'),
      mappingRule('O', 'AS', 'purchase'),
      mappingRule('Q', '', 'purchase', 'constant', { value: 'CTY' }),
      mappingRule('R', 'C', 'purchase'),
      mappingRule('S', 'D', 'purchase'),
      mappingRule('V', 'R', 'purchase', 'source_column', { transform: 'tax_code' }),
    ],
  };
  const sales: FormMappingPreset = {
    id: 'fast_hoadonbanhang',
    label: 'Hoadonbanhang',
    scope: 'sales',
    type: 'builtin',
    enabled: true,
    builtin_exporter: 'fast_hoadonbanhang',
    group_id: MATERIALS_GROUP_ID,
    input_phase: 'sales',
    sheet: 'Hoadonbanhang',
    output_columns: outputColumns(['Mã khách hàng (ma_kh)', 'Người mua hàng (ong_ba)', 'Quyển sổ (ma_qs)', 'Số seri (so_seri)', 'Số chứng từ (so_ct)', 'Ngày chứng từ (ngay_ct)', 'NVBH (ma_bp)', 'Số lượng:Q (so_luong)', 'Giá bán n.tệ:P1 (gia_nt2)', 'Tiền bán n.tệ:N1 (tien_nt2)', 'Giá vốn n.tệ:P1 (gia_nt)', 'Tiền vốn n.tệ:N1 (tien_nt)', 'Mã n.tệ (ma_nt)', 'Tỷ giá:R (ty_gia)', 'Giá bán:P0 (gia2)', 'Tiền bán:N0 (tien2)', 'Giá vốn:P0 (gia)', 'Tiền vốn:N0 (tien)', 'Tỷ lệ chiết khấu (tl_ck)', 'Tiền chiết khấu n.tệ:N1 (tien_ck_nt)', 'Tiền chiết khấu:N0 (tien_ck)', 'Mã thuế (ma_thue)', 'Thuế suất (thue_suat)', 'Tiền thuế n.tệ:N1 (tien_thue_nt)', 'Tiền thuế:N0 (tien_thue)', 'Mã nx (Tk nợ) (ma_nx)', 'Tk doanh thu (tk_dt)', 'Tk vật tư (tk_vt)', 'Tk giá vốn (tk_gv)', 'Tk chiết khấu (tk_ck)', 'Tài khoản thuế (tk_thue_co)', 'Mã kho  (ma_kho)', 'Mã vật tư (ma_vt)', 'Diễn giải (dien_giai)', 'Hạn thanh toán:N (han_tt)', 'Hình thức tt (ht_tt)', 'Loại hoá đơn (ma_gd)', 'Mã dự án (ma_vv_i)', 'Mã phí (ma_phi_i)', 'Mã bpht (ma_bpht_i)', 'Khuyến mại (km_ck)', 'Tk cp km (tk_km_i)', 'Mã ĐVCS (ma_dvcs)', 'Sử dụng HĐĐT (sd_hddt_yn)', 'Tỷ giá hđ:R (ty_gia_hd)', 'Thêm thông tin (xu_ly_hddt)', 'Thông tin thêm (gc_dc_hddt)', 'Mã loại (ma_loai)']),
    mappings: [
      mappingRule('A', 'AU', 'sales', 'source_column', { transform: 'mst_or_prefix' }),
      mappingRule('E', 'C', 'sales'),
      mappingRule('F', 'D', 'sales'),
      mappingRule('H', 'O', 'sales'),
      mappingRule('O', 'P', 'sales'),
      mappingRule('P', 'V', 'sales'),
      mappingRule('V', 'R', 'sales', 'source_column', { transform: 'tax_code' }),
      mappingRule('Y', 'W', 'sales'),
      mappingRule('Z', '', 'sales', 'constant', { value: '131' }),
      mappingRule('AA', '', 'sales', 'if_rules', { condition_source_col: 'AS', transform_rules: DEFAULT_REVENUE_ACCOUNT_RULES.map((rule) => ({ source_col: 'AS', ...rule })) }),
      mappingRule('AB', 'AS', 'sales'),
      mappingRule('AC', '', 'sales', 'constant', { value: '632' }),
      mappingRule('AE', '', 'sales', 'constant', { value: '33311' }),
      mappingRule('AF', 'AT', 'sales'),
      mappingRule('AG', 'L', 'sales'),
      mappingRule('AH', '', 'sales', 'constant', { value: 'Xuất bán hàng' }),
      mappingRule('AK', '', 'sales', 'constant', { value: '1' }),
      mappingRule('AQ', '', 'sales', 'constant', { value: 'CTY' }),
      mappingRule('AS', 'AR', 'sales'),
    ],
  };
  const material: FormMappingPreset = {
    id: 'fast_dm_vat_tu',
    label: 'Danh sách vật tư',
    scope: 'both',
    type: 'builtin',
    enabled: true,
    builtin_exporter: 'fast_dm_vat_tu',
    group_id: MATERIALS_GROUP_ID,
    input_phase: 'both',
    sheet: 'DMvat_tu',
    output_columns: outputColumns(['Mã vật tư', 'Tên vật tư', 'ĐVT', 'Theo dõi tồn kho', 'TK vật tư', 'TK giá vốn', 'TK doanh thu', 'TK hàng bán bị trả lại', 'TK sp dở dang', 'Loại vật tư', 'Cho sửa tk kho', 'TK NVL', 'TK chiết khấu', 'TK khuyến mại', 'Mã phụ', 'Tên 2', 'Cách tính giá tồn kho', 'Nhóm vt 1', 'Nhóm vt 2', 'Nhóm vt 3', 'Số lượng tồn tối thiểu', 'Số lượng tồn tối đa', 'TK chênh lệch vật tư']),
    mappings: [mappingRule('A', 'L', 'both'), mappingRule('B', 'M', 'both'), mappingRule('C', 'N', 'both'), mappingRule('D', '', 'both', 'constant', { value: '1' }), mappingRule('E', 'AS', 'both'), mappingRule('F', '', 'both', 'constant', { value: '632' }), mappingRule('G', '', 'both', 'if_rules', { condition_source_col: 'AS', transform_rules: DEFAULT_REVENUE_ACCOUNT_RULES.map((rule) => ({ source_col: 'AS', ...rule })) }), mappingRule('J', '', 'both', 'if_rules', { condition_source_col: 'AS', transform_rules: DEFAULT_MATERIAL_TYPE_RULES.map((rule) => ({ source_col: 'AS', ...rule })) }), mappingRule('K', '', 'both', 'constant', { value: '1' }), mappingRule('Q', '', 'both', 'constant', { value: '4' }), mappingRule('W', '', 'both', 'constant', { value: '632' })],
  };
  const customer: FormMappingPreset = {
    id: 'fast_dm_khach_hang',
    label: 'Danh sách khách hàng',
    scope: 'both',
    type: 'builtin',
    enabled: true,
    builtin_exporter: 'fast_dm_khach_hang',
    group_id: MATERIALS_GROUP_ID,
    input_phase: 'both',
    sheet: 'DMkhachhang',
    output_columns: outputColumns(['Mã khách hàng', 'Tên khách hàng', 'Tên 2', 'Mã số thuế', 'Địa chỉ']),
    mappings: [
      mappingRule('A', 'G', 'purchase', 'source_column', { transform: 'mst_or_prefix' }),
      mappingRule('B', 'F', 'purchase'),
      mappingRule('D', 'G', 'purchase'),
      mappingRule('E', 'H', 'purchase'),
      mappingRule('A', 'AU', 'sales', 'source_column', { transform: 'mst_or_prefix' }),
      mappingRule('B', 'I', 'sales'),
      mappingRule('D', 'J', 'sales'),
      mappingRule('E', 'K', 'sales'),
    ],
  };
  const duplicateReport: FormMappingPreset = {
    id: 'fast_duplicate_invoice_report',
    label: 'Báo cáo trùng số chứng từ',
    scope: 'both',
    type: 'builtin',
    enabled: true,
    builtin_exporter: 'fast_duplicate_report',
    group_id: MATERIALS_GROUP_ID,
    input_phase: 'both',
    sheet: 'Bao_cao_trung_so_ct',
    system_generated: true,
    output_columns: outputColumns(['Số chứng từ', 'Mã khách hàng', 'Tên công ty', 'Mã VT', 'Cảnh báo']),
    mappings: [],
  };
  const forms = [purchase, duplicateReport, sales, material, customer];
  if (phase === 'purchase') return forms.filter((form) => form.scope === 'purchase' || form.scope === 'both');
  if (phase === 'sales') return forms.filter((form) => form.scope === 'sales' || form.scope === 'both');
  return forms;
}

function mergeFormPreset(defaultPreset: FormMappingPreset | undefined, rawPreset: FormMappingPreset): FormMappingPreset {
  const preset = { ...(defaultPreset || {}), ...rawPreset };
  return {
    ...preset,
    id: String(preset.id || rawPreset.id || '').trim(),
    label: String(preset.label || preset.id || ''),
    mappings: Array.isArray(preset.mappings) ? preset.mappings : (defaultPreset?.mappings || []),
    output_columns: Array.isArray(preset.output_columns) ? preset.output_columns : (defaultPreset?.output_columns || []),
  };
}

function defaultProcessingGroups(phase: 'purchase' | 'sales' | 'generic' = 'purchase'): ProcessingGroup[] {
  const materialForms: ProcessingForm[] = [
    { id: 'processed_fdi', label: 'FDI đã xử lý', type: 'builtin', builtin_exporter: 'processed_fdi', enabled: true },
    ...fallbackDefaultFormMappingPresets(phase === 'sales' ? 'sales' : 'purchase'),
  ];
  if (phase !== 'sales') {
    materialForms.push({ id: 'nhap_kho', label: 'Nhập kho companion', type: 'builtin', builtin_exporter: 'nhap_kho', enabled: true });
  }
  return [
    { id: MATERIALS_GROUP_ID, label: 'Nhóm vật tư', builtin: true, uses_product_code: true, forms: materialForms },
    { id: SERVICE_GROUP_ID, label: 'Nhóm dịch vụ', builtin: true, uses_product_code: false, forms: [] },
    { id: IGNORED_GROUP_ID, label: 'Không xử lý', builtin: true, uses_product_code: false, forms: [] },
  ];
}

function normalizeProcessingGroups(raw: unknown, phase: 'purchase' | 'sales' | 'generic' = 'purchase'): ProcessingGroup[] {
  const defaults = defaultProcessingGroups(phase);
  const rows = Array.isArray(raw) ? raw : [];
  const byId = new Map(defaults.map((group) => [group.id, group]));
  rows.forEach((item) => {
    if (!item || typeof item !== 'object') return;
    const group = item as ProcessingGroup;
    const id = String(group.id || '').trim();
    if (!id) return;
    const defaultGroup = byId.get(id);
    const label = defaultGroup?.builtin ? defaultGroup.label : String(group.label || id);
    byId.set(id, { ...group, id, label, forms: Array.isArray(group.forms) ? group.forms : [] });
  });
  return Array.from(byId.values());
}

function singlePhaseFormPreset(form: FormMappingPreset, phase: 'purchase' | 'sales'): FormMappingPreset | null {
  const scope = formatScopeOfForm(form, phase);
  if (scope !== phase && scope !== 'both') return null;
  return {
    ...form,
    scope: phase,
    input_phase: phase,
    mappings: (form.mappings || [])
      .filter((rule) => !rule.source_phase || rule.source_phase === phase || rule.source_phase === 'both')
      .map((rule) => ({ ...rule, source_phase: phase })),
  };
}

function normalizeFormMappingPresets(raw: unknown, phase: 'purchase' | 'sales' | 'all' = 'purchase', backendDefaults?: FormatMappingDefaults | null, singleInputPhase = false): FormMappingPreset[] {
  const defaults = backendDefaults?.form_mapping_presets?.length
    ? backendDefaults.form_mapping_presets.filter((form) => phase === 'all' || form.scope === phase || form.scope === 'both')
    : fallbackDefaultFormMappingPresets(phase);
  const preparedDefaults = singleInputPhase && phase !== 'all'
    ? defaults.map((form) => singlePhaseFormPreset(form, phase)).filter((form): form is FormMappingPreset => Boolean(form))
    : defaults;
  const hasSavedRows = Array.isArray(raw);
  const rows = hasSavedRows ? raw : [];
  const defaultById = new Map(preparedDefaults.map((preset) => [preset.id, preset]));
  const byId = new Map(hasSavedRows ? [] : preparedDefaults.map((preset) => [preset.id, preset]));
  rows.forEach((item) => {
    if (!item || typeof item !== 'object') return [];
    const preset = item as FormMappingPreset;
    const id = String(preset.id || '').trim();
    if (!id) return;
    const merged = mergeFormPreset(defaultById.get(id), { ...preset, id });
    const prepared = singleInputPhase && phase !== 'all' ? singlePhaseFormPreset(merged, phase) : merged;
    if (prepared) byId.set(id, prepared);
  });
  return Array.from(byId.values()).filter((preset) => preset.enabled !== false);
}

function groupAssignmentsFromRows(rows: CompanyRow[]): Record<string, string> {
  const assignments: Record<string, string> = {};
  rows.forEach((company) => {
    const groupId = committedCompanyGroup(company);
    [companyConfigKey(company), company.company_id, company.mst, company.safe_id, company.company]
      .map((value) => String(value || '').trim())
      .filter(Boolean)
      .forEach((key) => { assignments[key] = groupId; });
  });
  return assignments;
}

function genericProfileFormPhase(profile: ProfileKey): 'purchase' | 'sales' | 'all' {
  if (profile === 'son_phuong') return 'all';
  return profile === 'cao_thanh' ? 'sales' : 'purchase';
}

function scopedProfileConfig(profilesCfg: Record<string, any>, profile: ProfileKey, phase: 'purchase' | 'sales') {
  const root = profilesCfg[profile] && typeof profilesCfg[profile] === 'object' ? profilesCfg[profile] : {};
  const scopes = root.scopes && typeof root.scopes === 'object' ? root.scopes : {};
  const scoped = scopes[phase] && typeof scopes[phase] === 'object' ? scopes[phase] : null;
  const scopedConfig = scoped && hasMeaningfulProfileConfig(scoped) ? scoped : null;
  if (profile === 'vietmax') {
    const legacyKey = phase === 'sales' ? 'vietmax_ban_ra' : 'vietmax_mua_vao';
    const legacy = profilesCfg[legacyKey] && typeof profilesCfg[legacyKey] === 'object' ? profilesCfg[legacyKey] : null;
    const legacyConfig = legacy && hasMeaningfulProfileConfig(legacy) ? legacy : null;
    return mergeProfileConfig(root, legacyConfig, scopedConfig);
  }
  if (scopedConfig) return mergeProfileConfig(root, scopedConfig);
  return root;
}

function mergeProfileConfig(...configs: Array<Record<string, any> | null | undefined>) {
  const result: Record<string, any> = {};
  configs.forEach((config) => {
    if (!config || typeof config !== 'object') return;
    Object.entries(config).forEach(([key, value]) => {
      if (key === 'scopes') return;
      if (shouldUseMergedProfileValue(key, value, result[key])) {
        result[key] = value;
      }
    });
  });
  return result;
}

function shouldUseMergedProfileValue(key: string, value: unknown, existing: unknown) {
  if (Array.isArray(value)) return value.length > 0 || !hasConfigValue(existing);
  if (value && typeof value === 'object') {
    if (key === 'word_rules' && isDefaultWordRules(value) && hasConfigValue(existing) && !isDefaultWordRules(existing)) return false;
    return Object.keys(value).length > 0 || !hasConfigValue(existing);
  }
  if (typeof value === 'string') return value.trim().length > 0 || !hasConfigValue(existing);
  return value !== undefined && value !== null;
}

function isDefaultWordRules(value: unknown) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const entries = Object.entries(value as Record<string, unknown>).map(([key, val]) => [String(key).toLowerCase(), String(val).toUpperCase()] as const);
  if (entries.length !== 2) return false;
  const asMap = new Map(entries);
  return asMap.get('đen') === 'DEN' && asMap.get('tôn') === 'TON';
}

function hasMeaningfulProfileConfig(config: Record<string, any>) {
  const nonEmptyCollections = [
    'prefixes',
    'selected_products',
    'removed_companies',
    'first_word_rules',
    'repeated_phrase_removals',
    'price_group_rules',
    'price_range_rules',
    'manual_code_overrides',
    'product_code_replacements',
    'product_review_merges',
    'vietmax_mua_vao_internal_merges',
    'vietmax_ban_ra_sales_internal_merges',
    'vietmax_ban_ra_purchase_match_rules',
    'inventory_pairs',
    'inventory_pair_rules',
    'inventory_allocation_config',
    'prefix_strategy_values',
    'processing_groups',
    'company_group_assignments',
    'form_mapping_presets',
    'columns',
  ];
  if (nonEmptyCollections.some((key) => hasConfigValue(config[key]))) return true;
  if (config.include_company_prefix === false) return true;
  if (config.use_default_inventory_pair === true) return true;
  if (String(config.default_inventory_pair_id || '').trim()) return true;
  if (String(config.prefix_strategy || 'last_2_words') !== 'last_2_words') return true;
  if (Number(config.prefix_mst_digits ?? 3) !== 3) return true;
  if (Number(config.prefix_name_words ?? 2) !== 2) return true;
  if (Number(config.prefix_name_chars ?? 1) !== 1) return true;
  if (String(config.prefix_missing_mst_strategy || 'all_name_words') !== 'all_name_words') return true;
  return false;
}

function hasConfigValue(value: unknown): boolean {
  if (Array.isArray(value)) return value.length > 0;
  if (value && typeof value === 'object') return Object.keys(value).length > 0;
  if (typeof value === 'string') return value.trim().length > 0;
  return value !== undefined && value !== null && value !== false;
}

function isTwoPhaseGenericProfile(profile: ProfileKey) {
  return profile !== 'vietmax' && usesTwoPhaseFrame(profile);
}

function applyGroupAssignments(rows: CompanyRow[], assignments: Record<string, string> = {}): CompanyRow[] {
  return rows.map((company) => {
    const key = companyConfigKey(company);
    const oldProcess = company.process ?? true;
    const groupId = String(company.group_id || assignments[key] || (oldProcess ? MATERIALS_GROUP_ID : IGNORED_GROUP_ID));
    const process = groupId === MATERIALS_GROUP_ID;
    return { ...company, group_id: groupId, pending_group_id: company.pending_group_id || groupId, process, pending_process: process };
  });
}

function formatScopeOfForm(form: ProcessingForm, fallback: FormatScope = 'purchase'): FormatScope {
  const scope = String(form.scope || '').trim();
  return scope === 'sales' || scope === 'both' || scope === 'purchase' ? scope : fallback;
}

function formatScopeMatches(form: ProcessingForm, scope: FormatScope) {
  return formatScopeOfForm(form) === scope;
}

function visibleFormatFormsForScope(scope: FormatScope, purchaseForms: FormMappingPreset[], salesForms: FormMappingPreset[]): FormMappingPreset[] {
  const source = scope === 'sales' ? salesForms : scope === 'purchase' ? purchaseForms : [...purchaseForms, ...salesForms];
  const byId = new Map<string, FormMappingPreset>();
  source.forEach((form) => {
    if (form.enabled === false || !formatScopeMatches(form, scope)) return;
    byId.set(form.id, form);
  });
  return Array.from(byId.values()).sort((left, right) => left.label.localeCompare(right.label, 'vi', { numeric: true, sensitivity: 'base' }));
}

function replaceFormatFormInList(list: FormMappingPreset[], scope: FormatScope, formId: string, updater: (form: FormMappingPreset) => FormMappingPreset): FormMappingPreset[] {
  return list.map((form) => (form.id === formId && formatScopeMatches(form, scope) ? updater(form) : form));
}

function removeFormatFormFromList(list: FormMappingPreset[], formId: string): FormMappingPreset[] {
  return list.filter((form) => form.id !== formId);
}

function upsertFormatForm(list: FormMappingPreset[], form: FormMappingPreset): FormMappingPreset[] {
  if (list.some((item) => item.id === form.id)) {
    return list.map((item) => (item.id === form.id ? form : item));
  }
  return [...list, form];
}

function formWithScope(form: FormMappingPreset, scope: FormatScope): FormMappingPreset {
  const sourcePhase = scope === 'sales' ? 'sales' : 'purchase';
  return {
    ...form,
    scope,
    input_phase: scope === 'both' ? 'both' : scope,
    mappings: (form.mappings || []).map((rule) => {
      if (scope === 'both') {
        const phase = rule.source_phase === 'sales' ? 'sales' : 'purchase';
        return { ...rule, source_phase: phase };
      }
      return { ...rule, source_phase: sourcePhase };
    }),
  };
}

function activeFormsRequirePhase(forms: FormMappingPreset[], phase: 'purchase' | 'sales') {
  return forms.some((form) => {
    if (form.enabled === false || !(form.mappings || []).length) return false;
    const scope = formatScopeOfForm(form, phase);
    if (scope !== phase && scope !== 'both') return false;
    return (form.mappings || []).some((rule) => !rule.source_phase || rule.source_phase === phase || rule.source_phase === 'both');
  });
}

function columnsFromUploadSummary(summary: UploadSummary): FormColumn[] {
  return summary.columns.map((column) => ({
    letter: column.letter,
    label: column.label.includes('.') ? column.label : `${column.letter}. ${column.label.replace(`${column.letter} - `, '')}`,
    header: column.label.replace(`${column.letter} - `, '').replace(`${column.letter}. `, ''),
  }));
}

function formatSourceColumns(raw: FormColumn[] | undefined, fallback: FormColumn[] = fallbackFdiSourceColumns): FormColumn[] {
  if (!Array.isArray(raw) || !raw.length) return fallback;
  const columns = raw
    .map((column) => ({
      letter: String(column.letter || '').trim().toUpperCase(),
      label: String(column.label || column.header || '').trim(),
      header: String(column.header || column.label || '').trim(),
    }))
    .filter((column) => column.letter && column.label);
  return columns.length ? columns : fallback;
}

function groupIdFromLabel(label: string, existing: ProcessingGroup[] = []) {
  const base = label
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'nhom';
  const used = new Set(existing.map((group) => group.id));
  let id = base;
  let suffix = 2;
  while (used.has(id)) {
    id = `${base}_${suffix}`;
    suffix += 1;
  }
  return id;
}

type WorkflowState = {
  stage: StageId;
  purchaseFile: UploadSummary | null;
  purchaseFormSourceColumns: FormColumn[];
  salesFormSourceColumns: FormColumn[];
  genericColumns: GenericColumns;
  purchaseColumns: GenericColumns;
  salesColumns: GenericColumns;
  processedPurchaseSavedName: string;
  processedPurchaseStats: ProcessedFileStats | null;
  salesFile: UploadSummary | null;
  processedSalesSavedName: string;
  processedSalesStats: ProcessedFileStats | null;
  openingStockFile: File | null;
  inventoryAllocationConfig: InventoryAllocationConfig;
  inventoryAllocationJob: InventoryAllocationJob | null;
  inventoryAllocationResult: InventoryAllocationResult | null;
  comparisonScope: string;
  companyRows: CompanyRow[];
  selectedCompanyIndex: number;
  purchaseMissingMstCompanies: MissingMstCompanyWarning[];
  purchaseInvoiceStatuses: InvoiceStatusOption[];
  salesCompanyRows: CompanyRow[];
  selectedSalesCompanyIndex: number;
  salesMissingMstCompanies: MissingMstCompanyWarning[];
  salesInvoiceStatuses: InvoiceStatusOption[];
  productPreviewCodes: Record<string, string>;
  salesProductPreviewCodes: Record<string, string>;
  productCodeOverrides: Record<string, string>;
  salesProductCodeOverrides: Record<string, string>;
  productCodeReplacements: Record<string, string>;
  salesProductCodeReplacements: Record<string, string>;
  purchaseWordRules: Record<string, string>;
  salesWordRules: Record<string, string>;
  firstWordRules: Record<string, string>;
  purchaseRepeatedPhraseRemovals: string[];
  salesRepeatedPhraseRemovals: string[];
  wordRules: Record<string, string>;
  repeatedPhraseRemovals: string[];
  purchaseReviewRows: ReviewRow[];
  salesReviewRows: ReviewRow[];
  purchaseReviewRules: ReviewRow[];
  salesReviewRules: ReviewRow[];
  purchaseReviewGenerated: boolean;
  salesReviewGenerated: boolean;
  priceRangeRules: Record<string, any>;
  priceGroups: CaoThanhPriceGroup[];
  priceFilterAllPercent: number;
  priceAdjustAllPercent: number;
  matches: MatchRow[];
  salesMatchGenerated: boolean;
  salesMatchRules: MatchRow[];
  purchaseInventoryPairs: InventoryPair[];
  purchaseUseDefaultInventoryPair: boolean;
  purchaseDefaultInventoryPairId: string;
  purchaseInventoryPairRules: InventoryRule[];
  salesInventoryPairs: InventoryPair[];
  salesUseDefaultInventoryPair: boolean;
  salesDefaultInventoryPairId: string;
  salesInventoryPairRules: InventoryRule[];
  inventoryPairs: InventoryPair[];
  useDefaultInventoryPair: boolean;
  defaultInventoryPairId: string;
  inventoryPairRules: InventoryRule[];
  includeCompanyPrefix: boolean;
  salesIncludeCompanyPrefix: boolean;
  purchasePrefixStrategy: PrefixPresetStrategy;
  salesPrefixStrategy: PrefixPresetStrategy;
  prefixMstDigits: number;
  prefixNameWords: number;
  prefixNameChars: number;
  prefixMissingMstStrategy: PrefixPresetStrategy;
  purchasePrefixStrategyValues: PrefixStrategyValues;
  salesPrefixStrategyValues: PrefixStrategyValues;
  purchaseProcessingGroups: ProcessingGroup[];
  salesProcessingGroups: ProcessingGroup[];
  purchaseFormMappingPresets: FormMappingPreset[];
  salesFormMappingPresets: FormMappingPreset[];
  purchaseReviewScope: 'all' | 'company';
  salesReviewScope: 'all' | 'company';
};

type ProductCodeReplacementDraftRow = {
  id: string;
  from: string;
  to: string;
};

type RelatedProductCodeWarning = {
  id: string;
  companyIndex: number;
  productName: string;
  companyName: string;
  baseCode: string;
  currentCode: string;
  suggestedCode: string;
};

type RelatedProductCodeUpdate = {
  companyIndex: number;
  productName: string;
  code: string;
};

const defaultInvoiceStatusSkipValues = [
  'Hóa đơn đã bị hủy',
];

function defaultVietmaxColumns(phase: 'purchase' | 'sales'): GenericColumns {
  const isSales = phase === 'sales';
  return {
    company_col: isSales ? 'I' : 'F',
    mst_col: isSales ? 'J' : 'G',
    address_col: isSales ? 'K' : 'H',
    product_col: 'M',
    qty_col: 'O',
    price_col: 'P',
    output_col: 'L',
    invoice_status_col: 'AJ',
    invoice_status_skip_values: defaultInvoiceStatusSkipValues,
  };
}

function defaultGenericColumns(): GenericColumns {
  return {
    company_col: 'F',
    mst_col: 'G',
    address_col: 'H',
    product_col: 'M',
    qty_col: 'O',
    price_col: '',
    output_col: 'L',
    invoice_status_col: 'AJ',
    invoice_status_skip_values: defaultInvoiceStatusSkipValues,
  };
}

function normalizeVietmaxColumns(raw: Record<string, unknown>, phase: 'purchase' | 'sales'): GenericColumns {
  const defaults = defaultVietmaxColumns(phase);
  const letter = (key: keyof GenericColumns) => String(raw[key] ?? defaults[key] ?? '').trim().toUpperCase();
  const skipValues = raw.invoice_status_skip_values;
  return {
    company_col: letter('company_col') || defaults.company_col,
    mst_col: letter('mst_col') || defaults.mst_col,
    address_col: letter('address_col'),
    product_col: letter('product_col') || defaults.product_col,
    qty_col: letter('qty_col'),
    price_col: letter('price_col'),
    output_col: letter('output_col') || defaults.output_col,
    invoice_status_col: letter('invoice_status_col'),
    invoice_status_skip_values: Array.isArray(skipValues) ? skipValues.map((item) => String(item)) : defaults.invoice_status_skip_values,
  };
}

function withInvoiceSkipFlags(statuses: InvoiceStatusOption[], skipValues: string[]): InvoiceStatusOption[] {
  const skipped = new Set(skipValues || []);
  return statuses.map((item) => ({ ...item, skip: skipped.has(item.value) }));
}

function initialWorkflowState(): WorkflowState {
  return {
    stage: 1,
    purchaseFile: null,
    purchaseFormSourceColumns: fallbackFdiSourceColumns,
    salesFormSourceColumns: fallbackFdiSourceColumns,
    genericColumns: defaultGenericColumns(),
    purchaseColumns: defaultVietmaxColumns('purchase'),
    salesColumns: defaultVietmaxColumns('sales'),
    processedPurchaseSavedName: '',
    processedPurchaseStats: null,
    salesFile: null,
    processedSalesSavedName: '',
    processedSalesStats: null,
    openingStockFile: null,
    inventoryAllocationConfig: defaultInventoryAllocationConfig(),
    inventoryAllocationJob: null,
    inventoryAllocationResult: null,
    comparisonScope: 'all_companies',
    companyRows: [],
    selectedCompanyIndex: -1,
    purchaseMissingMstCompanies: [],
    purchaseInvoiceStatuses: [],
    salesCompanyRows: [],
    selectedSalesCompanyIndex: -1,
    salesMissingMstCompanies: [],
    salesInvoiceStatuses: [],
    productPreviewCodes: {},
    salesProductPreviewCodes: {},
    productCodeOverrides: {},
    salesProductCodeOverrides: {},
    productCodeReplacements: {},
    salesProductCodeReplacements: {},
    purchaseWordRules: {},
    salesWordRules: {},
    firstWordRules: {},
    purchaseRepeatedPhraseRemovals: [],
    salesRepeatedPhraseRemovals: [],
    wordRules: {},
    repeatedPhraseRemovals: [],
    purchaseReviewRows: [],
    salesReviewRows: [],
    purchaseReviewRules: [],
    salesReviewRules: [],
    purchaseReviewGenerated: false,
    salesReviewGenerated: false,
    priceRangeRules: {},
    priceGroups: [],
    priceFilterAllPercent: 8,
    priceAdjustAllPercent: 0,
    matches: [],
    salesMatchGenerated: false,
    salesMatchRules: [],
    purchaseInventoryPairs: [],
    purchaseUseDefaultInventoryPair: false,
    purchaseDefaultInventoryPairId: '',
    purchaseInventoryPairRules: [],
    salesInventoryPairs: [],
    salesUseDefaultInventoryPair: false,
    salesDefaultInventoryPairId: '',
    salesInventoryPairRules: [],
    inventoryPairs: [],
    useDefaultInventoryPair: false,
    defaultInventoryPairId: '',
    inventoryPairRules: [],
    includeCompanyPrefix: false,
    salesIncludeCompanyPrefix: false,
    purchasePrefixStrategy: 'last_2_words',
    salesPrefixStrategy: 'last_2_words',
    prefixMstDigits: 3,
    prefixNameWords: 2,
    prefixNameChars: 1,
    prefixMissingMstStrategy: 'all_name_words',
    purchasePrefixStrategyValues: emptyPrefixStrategyValues(),
    salesPrefixStrategyValues: emptyPrefixStrategyValues(),
    purchaseProcessingGroups: defaultProcessingGroups('purchase'),
    salesProcessingGroups: defaultProcessingGroups('sales'),
    purchaseFormMappingPresets: fallbackDefaultFormMappingPresets('purchase'),
    salesFormMappingPresets: fallbackDefaultFormMappingPresets('sales'),
    purchaseReviewScope: 'all',
    salesReviewScope: 'company',
  };
}

function initialWorkflowStates(): Record<ProfileKey, WorkflowState> {
  return {
    son_phuong: { ...initialWorkflowState(), stage: 0.5 },
    cao_thanh: { ...initialWorkflowState(), stage: 0.5 },
    quang_thinh: { ...initialWorkflowState(), stage: 0.5 },
    vietmax: { ...initialWorkflowState(), stage: 0.5 },
    ho_guom: initialWorkflowState(),
    viet_hung: { ...initialWorkflowState(), stage: 0.5 },
  };
}

const FIXED_LICENSE_SERVER_URL = 'http://192.168.1.210:8080';
const FIXED_LICENSE_SERVER_FALLBACK_URL = 'http://192.168.101.13:8080';
const FIXED_LICENSE_ACCOUNT_ID = '6f1f56e8-3b6f-4a86-9a31-9e0e7f62c001';
const CLIENT_RELEASE_VERSION = import.meta.env.VITE_APP_VERSION || '0.4.0';

type ConfigTransferScope = 'purchase' | 'sales' | 'all';

function initialLicenseForm() {
  return { server_url: FIXED_LICENSE_SERVER_URL, account_id: FIXED_LICENSE_ACCOUNT_ID, license_key: '' };
}

function normalizeLicenseProfileText(value: unknown) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function licenseProfileAliases(profile: ProfileKey, label: string) {
  const aliases = [profile, label];
  if (profile === 'vietmax') aliases.push('vietmax_mua_vao', 'vietmax_ban_ra', 'Vietmax mua vào', 'Vietmax bán ra');
  return aliases.map(normalizeLicenseProfileText).filter(Boolean);
}

function licenseAllowsSelectedProfile(profile: ProfileKey, label: string, license: LicenseStatus | null) {
  if (!license?.activated) return false;
  const allowedProfiles = license.allowed_profiles || [];
  if (!allowedProfiles.length) return true;
  const allowed = new Set(allowedProfiles.map(normalizeLicenseProfileText).filter(Boolean));
  return licenseProfileAliases(profile, label).some((item) => allowed.has(item));
}

function licenseAllowsDropdownProfile(profile: ProfileKey, label: string, license: LicenseStatus | null) {
  if (!license?.activated) return true;
  const allowedProfiles = license.allowed_profiles || [];
  if (!allowedProfiles.length) return true;
  return licenseAllowsSelectedProfile(profile, label, license);
}

function selectedProfileLicenseText(profile: ProfileKey, label: string, license: LicenseStatus | null, ready: boolean) {
  if (ready) return 'Được phép dùng.';
  if (!license) return 'Đang kiểm tra license...';
  if (!license.activated) return 'Chưa kích hoạt license.';
  const allowed = (license.allowed_profiles || []).filter(Boolean);
  const suffix = allowed.length ? ` License hiện tại: ${allowed.join(', ')}.` : '';
  return `License đã kích hoạt nhưng chưa bao gồm ${label}.${suffix}`;
}

function sonPhuongSalesInventoryPairs(): InventoryPair[] {
  return [
    { id: 'son-phuong-sales-materials', role: 'materials', label: 'H\u00e0ng h\u00f3a v\u1eadt t\u01b0', ma_kho: 'KHHVT', tk_vat_tu: '156' },
    { id: 'son-phuong-sales-finished', role: 'finished_goods', label: 'Th\u00e0nh ph\u1ea9m th\u00e9p', ma_kho: 'KTP', tk_vat_tu: '155' },
    { id: 'son-phuong-sales-fallback', role: 'fallback', label: 'S\u1ea3n ph\u1ea9m c\u00f2n l\u1ea1i', ma_kho: 'KHOCK', tk_vat_tu: '159' },
  ];
}
function mergeRoleInventoryPairs(current: InventoryPair[] | undefined, defaults: InventoryPair[]): InventoryPair[] {
  const rows = Array.isArray(current) ? current.map((item) => ({ ...item })) : [];
  const keys = new Set(rows.map((item) => String(item.role || item.id || '')));
  defaults.forEach((item) => {
    const key = String(item.role || item.id || '');
    if (!keys.has(key)) rows.push({ ...item });
  });
  return rows;
}


function defaultInventoryAllocationConfig(): InventoryAllocationConfig {
  const invoiceDefaults = {
    sheet: '',
    header_row: 2,
    data_start_row: 3,
    invoice_col: 'C',
    date_col: 'D',
    code_col: 'L',
    product_col: 'M',
    qty_col: 'O',
    price_col: 'P',
  };
  return {
    mapping: {
      purchase: { ...invoiceDefaults },
      sales: { ...invoiceDefaults },
      opening: { sheet: '', header_row: 1, data_start_row: 2, invoice_col: '', date_col: '', code_col: 'A', product_col: 'B', qty_col: 'C', price_col: 'D' },
    },
    policy: {
      max_loss_percent: null,
      max_profit_percent: null,
      generic_min_take_quantity: null,
      generic_max_take_quantity: null,
      generic_min_type_count: 2,
      barem_tolerance_percent: 5,
      ignore_sale_suffix: false,
      allow_negative_export: true,
      company_profile: 'yen_thanh',
      allow_future_purchase_reorder: false,
      future_purchase_window_days: 31,
    },
    sales_inventory_pairs: [],
    sales_inventory_pair_rules: [],
    scenario_count: 100,
  };
}

function normalizeInventoryAllocationConfig(raw: unknown, targetProfile?: ProfileKey): InventoryAllocationConfig {
  const defaults = defaultInventoryAllocationConfig();
  const roleDefaults = targetProfile === 'son_phuong' ? sonPhuongSalesInventoryPairs() : [];
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
    return { ...defaults, sales_inventory_pairs: roleDefaults, scenario_count: 100 };
  }
  const value = raw as Partial<InventoryAllocationConfig>;
  const mapping = (value.mapping && typeof value.mapping === 'object' ? value.mapping : {}) as Partial<InventoryAllocationConfig['mapping']>;
  const policy = (value.policy && typeof value.policy === 'object' ? value.policy : {}) as Partial<InventoryAllocationConfig['policy']>;
  const scenarioCount = Math.max(1, Math.min(1000, Number(value.scenario_count || 100) || 100));
  return {
    mapping: {
      purchase: { ...defaults.mapping.purchase, ...(mapping.purchase || {}) },
      sales: { ...defaults.mapping.sales, ...(mapping.sales || {}) },
      opening: { ...defaults.mapping.opening, ...(mapping.opening || {}) },
    },
    policy: { ...defaults.policy, ...policy },
    sales_inventory_pairs: mergeRoleInventoryPairs(value.sales_inventory_pairs, roleDefaults),
    sales_inventory_pair_rules: Array.isArray(value.sales_inventory_pair_rules) ? value.sales_inventory_pair_rules : [],
    scenario_count: scenarioCount,
  };
}

function inventoryAllocationProfileFor(profile: ProfileKey) {
  return profile === 'son_phuong' ? 'son_phuong' : 'yen_thanh';
}

function salesOutputInvalidation(): Partial<WorkflowState> {
  return { processedSalesSavedName: '', processedSalesStats: null, inventoryAllocationJob: null, inventoryAllocationResult: null };
}

function purchaseOutputInvalidation(): Partial<WorkflowState> {
  return { processedPurchaseSavedName: '', processedPurchaseStats: null, matches: [], salesMatchGenerated: false, ...salesOutputInvalidation() };
}

export function VietmaxApp() {
  const [profile, setProfile] = useState<ProfileKey>('vietmax');
  const [workflows, setWorkflows] = useState<Record<ProfileKey, WorkflowState>>(initialWorkflowStates);
  const workflowsRef = useRef(workflows);
  const configSaveQueueRef = useRef<Promise<void>>(Promise.resolve());
  const estimateWorkflowRef = useRef<EstimateExtractorWorkflowHandle>(null);
  const [license, setLicense] = useState<LicenseStatus | null>(null);
  const [licenseChecking, setLicenseChecking] = useState(true);
  const [licenseCheckError, setLicenseCheckError] = useState('');
  const [licenseForm, setLicenseForm] = useState(initialLicenseForm);
  const [updateManifest, setUpdateManifest] = useState<UpdateManifest | null>(null);
  const [updateBusy, setUpdateBusy] = useState(false);
  const [updateProgress, setUpdateProgress] = useState('');
  const [updateModalOpen, setUpdateModalOpen] = useState(false);
  const [configTransferScope, setConfigTransferScope] = useState<ConfigTransferScope>('all');
  const [status, setStatus] = useState('Chọn profile và bắt đầu theo từng stage. Dữ liệu được giữ khi chuyển stage, chỉ xóa khi bấm Làm lại.');
  const [busy, setBusy] = useState(false);
  const [autoSavingConfig, setAutoSavingConfig] = useState(false);
  const [profileConfigLoading, setProfileConfigLoading] = useState(false);
  const [loadingProgress, setLoadingProgress] = useState<OperationProgress | null>(null);
  const [hoGuomMode, setHoGuomMode] = useState<HoGuomMode>('estimate');
  const workflow = workflows[profile];
  const { stage, purchaseFile, purchaseFormSourceColumns, salesFormSourceColumns, genericColumns, purchaseColumns, salesColumns, processedPurchaseSavedName, processedPurchaseStats, salesFile, processedSalesSavedName, processedSalesStats, openingStockFile, inventoryAllocationConfig, inventoryAllocationJob, inventoryAllocationResult, comparisonScope, companyRows, selectedCompanyIndex, purchaseMissingMstCompanies, purchaseInvoiceStatuses, salesCompanyRows, selectedSalesCompanyIndex, salesMissingMstCompanies, salesInvoiceStatuses, productPreviewCodes, salesProductPreviewCodes, productCodeOverrides, salesProductCodeOverrides, productCodeReplacements, salesProductCodeReplacements, purchaseWordRules, salesWordRules, firstWordRules, purchaseRepeatedPhraseRemovals, salesRepeatedPhraseRemovals, wordRules, repeatedPhraseRemovals, purchaseReviewRows, salesReviewRows, purchaseReviewRules, salesReviewRules, purchaseReviewGenerated, salesReviewGenerated, priceRangeRules, priceGroups, priceFilterAllPercent, priceAdjustAllPercent, matches, salesMatchGenerated, salesMatchRules, purchaseInventoryPairs, purchaseUseDefaultInventoryPair, purchaseDefaultInventoryPairId, purchaseInventoryPairRules, salesInventoryPairs, salesUseDefaultInventoryPair, salesDefaultInventoryPairId, salesInventoryPairRules, inventoryPairs, useDefaultInventoryPair, defaultInventoryPairId, inventoryPairRules, includeCompanyPrefix: purchaseIncludeCompanyPrefix, salesIncludeCompanyPrefix, purchasePrefixStrategy, salesPrefixStrategy, prefixMstDigits, prefixNameWords, prefixNameChars, prefixMissingMstStrategy, purchasePrefixStrategyValues, salesPrefixStrategyValues, purchaseProcessingGroups, salesProcessingGroups, purchaseFormMappingPresets, salesFormMappingPresets, purchaseReviewScope, salesReviewScope } = workflow;
  const includeCompanyPrefix = stage >= 6 ? salesIncludeCompanyPrefix : purchaseIncludeCompanyPrefix;
  const selectedProfile = profiles.find((item) => item.key === profile) ?? profiles[0];
  const visibleProfiles = useMemo(
    () => profiles.filter((item) => licenseAllowsDropdownProfile(item.key, item.label, license)),
    [license],
  );
  const licenseReady = licenseAllowsSelectedProfile(profile, selectedProfile.label, license);
  const licenseProfileText = licenseCheckError || selectedProfileLicenseText(profile, selectedProfile.label, license, licenseReady);
  const isGenericProfile = isGenericProfileKey(profile);
  const isHoGuomFormatter = profile === 'ho_guom' && hoGuomMode === 'formatter';
  const isHoGuomEstimate = profile === 'ho_guom' && hoGuomMode === 'estimate';
  const isGenericWorkflowProfile = isGenericProfile || isHoGuomFormatter;
  const usesNativeStageShell = profile === 'vietmax' || isGenericWorkflowProfile || profile === 'ho_guom';
  const visibleStages = useMemo(() => (isHoGuomFormatter ? hoGuomFormatterStages : stagesForProfile(profile)), [isHoGuomFormatter, profile]);
  const currentStage = visibleStages.find((item) => item.id === stage) ?? visibleStages[0];
  const selectedMatches = useMemo(() => matches.filter((match) => match.confirmed !== false), [matches]);
  const showLicenseBar = stage === 0.5 || stage === 1;
  const activeVietmaxSalesConfig = hasVietmaxPurchaseMatch(profile) && stage >= 6;
  const activeTwoPhaseSalesConfig = isTwoPhaseGenericProfile(profile) && stage >= 6;
  const activeUsesScopedPhase = usesTwoPhaseFrame(profile);
  const activeInventoryConfigScope: InventoryConfigScope = activeUsesScopedPhase ? (activeVietmaxSalesConfig || activeTwoPhaseSalesConfig ? 'sales' : 'purchase') : 'generic';
  const activeWordRules = activeUsesScopedPhase ? (activeVietmaxSalesConfig || activeTwoPhaseSalesConfig ? salesWordRules : purchaseWordRules) : wordRules;
  const activeRepeatedPhraseRemovals = activeUsesScopedPhase ? (activeVietmaxSalesConfig || activeTwoPhaseSalesConfig ? salesRepeatedPhraseRemovals : purchaseRepeatedPhraseRemovals) : repeatedPhraseRemovals;
  const fastImportRequirements = useMemo(() => {
    const purchaseForms = normalizeFormsForSave(purchaseFormMappingPresets);
    const salesForms = normalizeFormsForSave(salesFormMappingPresets);
    return {
      purchaseForms,
      salesForms,
      needsPurchase: activeFormsRequirePhase(purchaseForms, 'purchase') || activeFormsRequirePhase(salesForms, 'purchase'),
      needsSales: activeFormsRequirePhase(purchaseForms, 'sales') || activeFormsRequirePhase(salesForms, 'sales'),
    };
  }, [purchaseFormMappingPresets, salesFormMappingPresets]);

  function updateWorkflow(targetProfile: ProfileKey, update: Partial<WorkflowState>) {
    const current = workflowsRef.current;
    const next = { ...current, [targetProfile]: { ...current[targetProfile], ...update } };
    workflowsRef.current = next;
    setWorkflows(next);
  }

  function persistWorkflowConfig(
    targetWorkflow: WorkflowState,
    phase: 'purchase' | 'sales' | 'all',
    targetProfile: ProfileKey,
    replaceFormMappings = false,
  ) {
    const payloads = buildConfigPayloads(targetWorkflow, phase, targetProfile).map((payload) => (
      replaceFormMappings ? { ...payload, replace_form_mapping_presets: true } : payload
    ));
    const queuedSave = configSaveQueueRef.current
      .catch(() => undefined)
      .then(async () => {
        for (const payload of payloads) await saveVietmaxConfig(payload);
      });
    configSaveQueueRef.current = queuedSave;
    return queuedSave;
  }

  async function waitForPendingConfigSave() {
    await configSaveQueueRef.current;
  }

  async function restoreProcessedCachesFromSession(targetProfile: ProfileKey = profile) {
    const currentWorkflow = workflows[targetProfile];
    let purchaseSavedName = currentWorkflow.processedPurchaseSavedName;
    let salesSavedName = currentWorkflow.processedSalesSavedName;
    try {
      const session = await getWorkflowSession();
      const artifacts = Object.values(session.artifacts || {})
        .filter((artifact) => artifact.valid)
        .sort((left, right) => right.created_at - left.created_at);
      const latestSavedName = (kinds: string[]) => artifacts.find((artifact) => kinds.includes(artifact.kind))?.saved_name || '';
      purchaseSavedName = latestSavedName([
        `processed:${targetProfile}:purchase`,
        `source:${targetProfile}-processed-purchase`,
        `source:${targetProfile}-fast-purchase`,
      ]) || purchaseSavedName;
      salesSavedName = latestSavedName([
        `processed:${targetProfile}:sales`,
        `source:${targetProfile}-processed-sales`,
        `source:${targetProfile}-fast-sales`,
      ]) || salesSavedName;
      updateWorkflow(targetProfile, {
        ...(purchaseSavedName ? { processedPurchaseSavedName: purchaseSavedName } : {}),
        ...(salesSavedName ? { processedSalesSavedName: salesSavedName } : {}),
      });
    } catch {
      // Keep the current in-memory cache references when session recovery is unavailable.
    }
    return { purchaseSavedName, salesSavedName };
  }

  useEffect(() => {
    void checkLicenseStatus();
  }, []);

  useEffect(() => {
    getWorkflowSession()
      .then((session) => {
        const artifacts = Object.values(session.artifacts || {})
          .filter((artifact) => artifact.valid)
          .sort((left, right) => left.created_at - right.created_at);
        setWorkflows((current) => {
          const next = { ...current };
          const latest = (kind: string) => [...artifacts].reverse().find((artifact) => artifact.kind === kind);
          for (const profileItem of profiles) {
            const target = { ...next[profileItem.key] };
            const purchaseSource = latest(`source:${profileItem.key}-purchase`);
            const salesSource = latest(`source:${profileItem.key}-sales`);
            const purchaseProcessed = latest(`processed:${profileItem.key}:purchase`)
              || latest(`source:${profileItem.key}-processed-purchase`)
              || latest(`source:${profileItem.key}-fast-purchase`);
            const salesProcessed = latest(`processed:${profileItem.key}:sales`)
              || latest(`source:${profileItem.key}-processed-sales`)
              || latest(`source:${profileItem.key}-fast-sales`);
            const purchaseSummary = purchaseSource?.metadata?.summary as UploadSummary | undefined;
            const salesSummary = salesSource?.metadata?.summary as UploadSummary | undefined;
            if (purchaseSummary) target.purchaseFile = purchaseSummary;
            if (salesSummary) target.salesFile = salesSummary;
            if (purchaseProcessed) target.processedPurchaseSavedName = purchaseProcessed.saved_name;
            if (salesProcessed) target.processedSalesSavedName = salesProcessed.saved_name;
            next[profileItem.key] = target;
          }
          workflowsRef.current = next;
          return next;
        });
      })
      .catch(() => {
        // Session recovery is best-effort; normal upload stages remain available.
      });
  }, []);

  useEffect(() => {
    if (isGenericWorkflowProfile) void loadGenericProfileConfig(profile);
    if (profile === 'vietmax') void loadVietmaxProfileConfig();
  }, [profile, isGenericWorkflowProfile]);

  useEffect(() => {
    if (!license?.activated || !visibleProfiles.length || visibleProfiles.some((item) => item.key === profile)) return;
    const nextProfile = visibleProfiles[0];
    setStatus(`License hiện tại không bao gồm ${selectedProfile.label}; đã chuyển sang ${nextProfile.label}.`);
    changeProfile(nextProfile.key);
  }, [license, profile, selectedProfile.label, visibleProfiles]);

  useEffect(() => {
    if (isGenericWorkflowProfile && stage === 3 && purchaseFile && !companyRows.length && !busy) {
      void loadGenericCompanies();
    }
  }, [isGenericWorkflowProfile, stage, purchaseFile, companyRows.length, busy]);

  useEffect(() => {
    if (profile === 'vietmax' && (stage === 3 || stage === 4) && purchaseFile && !companyRows.length && !busy) {
      void loadCompanies();
    }
  }, [profile, stage, purchaseFile, companyRows.length, busy]);

  useEffect(() => {
    if (profile === 'vietmax' && (stage === 9 || stage === 10) && salesFile && !salesCompanyRows.length && !busy) {
      void loadSalesCompanies();
    }
  }, [profile, stage, salesFile, salesCompanyRows.length, busy]);

  useEffect(() => {
    if (isTwoPhaseGenericProfile(profile) && stage === 9 && salesFile && !salesCompanyRows.length && !busy) {
      void loadGenericSalesCompanies();
    }
  }, [profile, stage, salesFile, salesCompanyRows.length, busy]);

  useEffect(() => {
    if (profile === 'cao_thanh' && stage === 5 && companyRows.length && !priceGroups.length && !busy) {
      updateCaoThanhPriceGroups();
    }
  }, [profile, stage, companyRows.length, priceGroups.length, busy]);

  useEffect(() => {
    if (profile === 'vietmax' && processedPurchaseSavedName && !processedPurchaseStats && !busy) {
      void refreshProcessedFileStats('purchase', processedPurchaseSavedName);
    }
  }, [profile, processedPurchaseSavedName, processedPurchaseStats, busy]);

  useEffect(() => {
    if (profile === 'vietmax' && processedSalesSavedName && !processedSalesStats && !busy) {
      void refreshProcessedFileStats('sales', processedSalesSavedName);
    }
  }, [profile, processedSalesSavedName, processedSalesStats, busy]);

  async function resetWorkflow() {
    setBusy(true);
    setStatus('Đang hoàn tất lưu cấu hình trước khi làm lại...');
    try {
      await waitForPendingConfigSave();
      const resetState: Partial<WorkflowState> = profile === 'vietmax' || isGenericWorkflowProfile ? { ...initialWorkflowState(), stage: 0.5 } : initialWorkflowState();
      updateWorkflow(profile, resetState);
      setStatus('Đang làm lại và nạp toàn bộ cấu hình đã lưu...');
      if (profile === 'vietmax') {
        await loadVietmaxProfileConfig();
      } else if (isGenericWorkflowProfile) {
        await loadGenericProfileConfig(profile);
      }
      setStatus(profile === 'vietmax' ? 'Đã làm lại và nạp toàn bộ cấu hình Vietmax từ file config.' : `Đã làm lại và nạp toàn bộ cấu hình ${selectedProfile.label} từ file config.`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Không thể làm lại vì cấu hình chưa lưu thành công: ${message}`);
    } finally {
      setBusy(false);
    }
  }

  function canEnterStage(target: StageId) {
    const targetDefinition = visibleStages.find((item) => item.id === target);
    if (!targetDefinition || targetDefinition.disabled) return false;
    if (isHoGuomEstimate) return target === 1 || licenseReady;
    if (isGenericWorkflowProfile) {
      if (!licenseReady) return target === 0.5 || target === 1;
      if (isHoGuomFormatter) {
        if (target === 0.5 || target === 1) return true;
        if (target === 2 || target === 3 || target === 5) return Boolean(purchaseFile && (target < 5 || companyRows.length));
        return false;
      }
      if (isTwoPhaseGenericProfile(profile)) {
        if (target === 0.5 || target === 1 || target === 6 || target === 15) return true;
        if (profile === 'son_phuong') {
          if (target === 2 || target === 3) return Boolean(purchaseFile);
          if (target === 4 || target === 5) return Boolean(purchaseFile && companyRows.length);
          if (target === 7 || target === 8) return Boolean(salesFile && processedPurchaseSavedName);
          if (target >= 10 && target <= 14) return Boolean(inventoryAllocationResult?.job_id || inventoryAllocationJob?.result?.job_id);
          return false;
        }

        if (target === 2 || target === 3) return Boolean(purchaseFile);
        if (target === 4 || target === 5) return Boolean(purchaseFile && companyRows.length);
        if (target === 7 || target === 9) return Boolean(salesFile);
        if (target === 10 || target === 11) return Boolean(salesFile && salesCompanyRows.length);
        return false;
      }
      if (target === 0.5) return true;
      if (target === 1) return true;
      if (target === 2 || target === 3) return Boolean(purchaseFile);
      if (target === 4) return Boolean(purchaseFile && companyRows.length);
      if (profile === 'cao_thanh' && (target === 5 || target === 6)) return Boolean(purchaseFile && companyRows.length && purchaseReviewGenerated);
      if (target === 5) return Boolean(purchaseFile && companyRows.length && purchaseReviewGenerated);
      return false;
    }
    if (profile !== 'vietmax') return licenseReady || target === 1;
    if (!licenseReady) return target === 0.5 || target === 1;
    if (target === 0.5 || target <= 2) return true;
    if (target === 6) return true;
    if (target === 12 || target === 15) return true;
    if (target <= 5) return Boolean(purchaseFile);
    if (target <= 11) return Boolean(salesFile && (purchaseFile || processedPurchaseSavedName));
    if (target === 13 || target === 14) return Boolean(inventoryAllocationResult?.job_id || inventoryAllocationJob?.result?.job_id);
    return true;
  }

  function cacheDebugMessage(prefix: string, needsPurchase = true, needsSales = true, caches: { purchase?: string; sales?: string } = {}) {
    const purchaseCache = caches.purchase ?? processedPurchaseSavedName;
    const salesCache = caches.sales ?? processedSalesSavedName;
    const missing = [
      needsPurchase && !purchaseCache ? 'mua vào' : '',
      needsSales && !salesCache ? 'bán ra' : '',
    ].filter(Boolean).join(', ') || 'không';
    return `${prefix} Thiếu cache: ${missing}. Debug cache: purchase=${purchaseCache || '-'}; sales=${salesCache || '-'}; purchaseFile=${purchaseFile?.original_name || '-'}; salesFile=${salesFile?.original_name || '-'}; stage=${stage}. Nếu vừa sửa cột/nhóm/review sau khi tạo file, hãy quay lại stage tạo file tương ứng để tạo cache lại.`;
  }

  function autosavePhaseForStage(currentStage: StageId): 'purchase' | 'sales' | 'all' {
    if (currentStage === 0.5 || currentStage >= 12) return 'all';
    return currentStage >= 6 ? 'sales' : 'purchase';
  }

  async function autoSaveCurrentConfigBeforeNavigation() {
    if (isHoGuomEstimate) return true;
    const phase = autosavePhaseForStage(stage);
    setAutoSavingConfig(true);
    setStatus('Đang tự lưu cấu hình trước khi chuyển stage...');
    try {
      await persistWorkflowConfig(workflowsRef.current[profile], phase, profile);
      setStatus('Đã tự lưu cấu hình. Đang chuyển stage...');
      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`Không tự lưu được cấu hình nên chưa chuyển stage. ${message}`);
      return false;
    } finally {
      setAutoSavingConfig(false);
    }
  }

  async function goToStage(target: StageId, options: { autosave?: boolean } = {}) {
    if (!canEnterStage(target)) {
      if (target === 12) {
        setStatus(cacheDebugMessage('Chưa đủ cache để vào phân bổ tồn kho.', true, true));
      } else if (target === 15) {
        setStatus(cacheDebugMessage('Chưa đủ cache để xuất FAST theo form mapping.', fastImportRequirements.needsPurchase, fastImportRequirements.needsSales));
      }
      return;
    }
    if (target === stage) return;
    if (isHoGuomEstimate && target === 3) {
      const ready = await estimateWorkflowRef.current?.prepareAnalysis();
      if (!ready) return;
    }
    if (stage === 3 && companyRows.some(hasCompanyDraftChanges)) {
      await applyCompanyAndProductChoices(target);
      return;
    }
    if (stage === 9 && salesCompanyRows.some(hasCompanyDraftChanges)) {
      await applySalesCompanyAndProductChoices(target);
      return;
    }
    if (options.autosave !== false && !(await autoSaveCurrentConfigBeforeNavigation())) return;
    if (target === 12 || target === 15) {
      await restoreProcessedCachesFromSession();
    }
    if ((target === 13 || target === 14) && !inventoryAllocationResult?.report_view) {
      const allocationJobId = inventoryAllocationResult?.job_id || inventoryAllocationJob?.result?.job_id;
      if (allocationJobId) {
        setBusy(true);
        setStatus('Đang tải dữ liệu báo cáo Phân kho...');
        try {
          const reportJob = await getInventoryAllocationJob(allocationJobId, true);
          if (reportJob.status !== 'complete' || !reportJob.result) {
            throw new Error(reportJob.error || reportJob.label || 'Kết quả Phân kho chưa sẵn sàng.');
          }
          updateWorkflow(profile, {
            inventoryAllocationJob: reportJob,
            inventoryAllocationResult: reportJob.result,
          });
        } catch (error) {
          setStatus(error instanceof Error ? error.message : String(error));
          return;
        } finally {
          setBusy(false);
        }
      }
    }
    updateWorkflow(profile, { stage: target });
  }

  function adjacentEnterableStage(direction: 1 | -1) {
    const currentIndex = visibleStages.findIndex((item) => item.id === stage);
    if (currentIndex < 0) return undefined;
    for (let index = currentIndex + direction; index >= 0 && index < visibleStages.length; index += direction) {
      const candidate = visibleStages[index];
      if (canEnterStage(candidate.id)) return candidate;
    }
    return undefined;
  }

  async function goBack() {
    const previous = adjacentEnterableStage(-1);
    if (previous) await goToStage(previous.id);
  }

  async function goNext() {
    const pendingNext = adjacentEnterableStage(1);
    if (stage === 3 && companyRows.some(hasCompanyDraftChanges)) {
      await applyCompanyAndProductChoices(pendingNext?.id);
      return;
    }
    if (stage === 9 && salesCompanyRows.some(hasCompanyDraftChanges)) {
      await applySalesCompanyAndProductChoices(pendingNext?.id);
      return;
    }
    if (!(await autoSaveCurrentConfigBeforeNavigation())) return;
    if (profile === 'vietmax' && stage === 11 && salesFile && !processedSalesSavedName) {
      if (!processedPurchaseSavedName) {
        setStatus(cacheDebugMessage('Cần có cache mua vào trước khi tạo cache bán ra.', true, false));
        return;
      }
      setStatus('Hãy bấm Xuất file bán ra để tạo cache thành công trước khi tiếp tục.');
      return;
    }
    if (profile === 'vietmax' && stage === 12 && !inventoryAllocationResult?.job_id && !inventoryAllocationJob?.result?.job_id) {
      void runInventoryAllocation(13);
      return;
    }
    const next = pendingNext || adjacentEnterableStage(1);
    if (profile === 'son_phuong' && stage === 8 && !inventoryAllocationResult?.job_id && !inventoryAllocationJob?.result?.job_id) {
      void runInventoryAllocation(13);
      return;
    }
    if (!next) {
      if (profile === 'vietmax' && stage >= 11) setStatus(cacheDebugMessage('Chưa có stage tiếp theo sẵn sàng.', true, true));
      return;
    }
    if (isGenericWorkflowProfile) {
      if (isTwoPhaseGenericProfile(profile)) {
        if (stage === 2 && purchaseFile && !companyRows.length) {
          void loadGenericCompanies(next.id);
          return;
        }
        if (stage === 3 && companyRows.some(hasCompanyDraftChanges)) {
          applyCompanyAndProductChoices(next.id);
          return;
        }
        if (stage === 4 && purchaseFile && !purchaseReviewGenerated) {
          setStatus('Hãy bấm Tạo danh sách review trước khi tiếp tục.');
          return;
        }
        if (stage === 5 && purchaseFile && !processedPurchaseSavedName) {
          setStatus('Hãy bấm Xuất file mua vào để tạo cache thành công trước khi tiếp tục.');
          return;
        }
        if (profile !== 'son_phuong' && stage === 7 && salesFile && !salesCompanyRows.length) {
          void loadGenericSalesCompanies(next.id);
          return;
        }
        if (stage === 9 && salesCompanyRows.some(hasCompanyDraftChanges)) {
          applySalesCompanyAndProductChoices(next.id);
          return;
        }
        if (profile !== 'son_phuong' && stage === 10 && salesFile && !salesReviewGenerated) {
          setStatus('Hãy bấm Tạo danh sách review bán ra trước khi tiếp tục.');
          return;
        }
        if (stage === 11 && salesFile && !processedSalesSavedName) {
          setStatus('Hãy bấm Xuất file bán ra để tạo cache thành công trước khi tiếp tục.');
          return;
        }
        void goToStage(next.id, { autosave: false });
        return;
      }
      if (stage === 2 && purchaseFile && !companyRows.length) {
        void loadGenericCompanies(next.id);
        return;
      }
      if (stage === 3 && companyRows.some(hasCompanyDraftChanges)) {
        applyCompanyAndProductChoices(next.id);
        return;
      }
      if (stage === 4 && purchaseFile && !purchaseReviewGenerated) {
        setStatus('Hãy bấm Tạo danh sách review trước khi tiếp tục.');
        return;
      }
      if (profile === 'cao_thanh' && stage === 5 && !priceGroups.length) {
        updateCaoThanhPriceGroups();
      }
      void goToStage(next.id, { autosave: false });
      return;
    }
    if (profile !== 'vietmax') {
      void goToStage(next.id, { autosave: false });
      return;
    }
    if (stage === 4 && purchaseFile && !purchaseReviewGenerated) {
      setStatus('Hãy bấm Tạo danh sách review mua vào trước khi tiếp tục.');
      return;
    }
    if (stage === 5 && purchaseFile && !processedPurchaseSavedName) {
      setStatus('Hãy bấm Xuất file mua vào để tạo cache thành công trước khi tiếp tục.');
      return;
    }
    if (stage === 6 && salesFile && !purchaseFile && !processedPurchaseSavedName) {
      setStatus('Cần tải file mua vào đã xử lý trước khi đi tiếp các stage bán ra.');
      return;
    }
    if (stage === 8 && salesFile && (processedPurchaseSavedName || purchaseFile) && !matches.length && !salesMatchGenerated) {
      void runSalesMatch();
      return;
    }
    if (stage === 10 && salesFile && !salesReviewGenerated) {
      setStatus('Hãy bấm Tạo danh sách review bán ra trước khi tiếp tục.');
      return;
    }
    void goToStage(next.id, { autosave: false });
  }

  async function submitLicense() {
    setBusy(true);
    setStatus('Đang kích hoạt license...');
    try {
      const nextLicense = await activateLicense({
        license_key: licenseForm.license_key,
      });
      setLicense(nextLicense);
      const nextReady = licenseAllowsSelectedProfile(profile, selectedProfile.label, nextLicense);
      setStatus(nextReady ? `Kích hoạt thành công. ${selectedProfile.label} được phép dùng.` : selectedProfileLicenseText(profile, selectedProfile.label, nextLicense, nextReady));
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function refreshLicense() {
    setBusy(true);
    setStatus('Đang tải lại license...');
    try {
      const nextLicense = await reloadLicense();
      setLicense(nextLicense);
      const nextReady = licenseAllowsSelectedProfile(profile, selectedProfile.label, nextLicense);
      setStatus(nextReady ? `Đã tải lại license. ${selectedProfile.label} được phép dùng.` : selectedProfileLicenseText(profile, selectedProfile.label, nextLicense, nextReady));
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function checkUpdate() {
    setUpdateBusy(true);
    try {
      const result = await checkForUpdate();
      setUpdateManifest(result);
      setStatus(result.available ? `Có bản cập nhật ${result.version} trên server.` : `Đang dùng phiên bản mới nhất ${result.current_version}.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setUpdateBusy(false);
    }
  }

  async function installUpdate() {
    if (!updateManifest?.available) return;
    setUpdateBusy(true);
    setUpdateProgress(`Đang tải bản ${updateManifest.version}...`);
    try {
      await applyUpdate(updateManifest);
      setUpdateProgress('Đã tải và xác minh xong. Velopack sẽ đóng ứng dụng, cài bản mới rồi tự mở lại.');
      setStatus(`Đã tải bản ${updateManifest.version}; đang chuyển sang trình cập nhật Velopack.`);
    } catch (error) {
      setUpdateProgress('Cập nhật thất bại. Ứng dụng vẫn đang chạy để bạn kiểm tra lại.');
      setStatus(error instanceof Error ? error.message : String(error));
      setUpdateBusy(false);
    }
  }
  async function upload(kind: 'purchase' | 'sales', file: File | undefined) {
    if (!file) return;
    const targetProfile = profile;
    setBusy(true);
    setStatus(`Đang tải ${kind === 'purchase' ? 'HD mua vào' : 'HD bán ra'}...`);
    try {
      const summary = await uploadExcel(file, `${profile}-${kind}`);
      if (kind === 'purchase') {
        const nextColumns = normalizeVietmaxColumns(purchaseColumns, 'purchase');
        const statuses = nextColumns.invoice_status_col
          ? await fetchInvoiceStatuses(summary.saved_name, nextColumns.invoice_status_col, nextColumns.invoice_status_skip_values).catch(() => withInvoiceSkipFlags(summary.invoice_statuses || [], nextColumns.invoice_status_skip_values))
          : [];
        updateWorkflow(targetProfile, {
          ...purchaseOutputInvalidation(),
          purchaseFile: summary,
          stage: 2,
          purchaseColumns: nextColumns,
          purchaseInvoiceStatuses: statuses,
          companyRows: [],
          selectedCompanyIndex: -1,
          purchaseMissingMstCompanies: [],
          productPreviewCodes: {},
          productCodeOverrides: {},
          purchaseReviewRows: [],
          purchaseReviewGenerated: false,
          salesCompanyRows: [],
          selectedSalesCompanyIndex: -1,
          salesMissingMstCompanies: [],
          salesProductPreviewCodes: {},
          salesProductCodeOverrides: {},
          salesReviewRows: [],
          salesReviewGenerated: false,
          inventoryAllocationJob: null,
          inventoryAllocationResult: null,
        });
      } else {
        const nextColumns = normalizeVietmaxColumns(salesColumns, 'sales');
        const statuses = nextColumns.invoice_status_col
          ? await fetchInvoiceStatuses(summary.saved_name, nextColumns.invoice_status_col, nextColumns.invoice_status_skip_values).catch(() => withInvoiceSkipFlags(summary.invoice_statuses || [], nextColumns.invoice_status_skip_values))
          : [];
        updateWorkflow(targetProfile, {
          ...salesOutputInvalidation(),
          salesFile: summary,
          stage: 7,
          salesColumns: nextColumns,
          salesInvoiceStatuses: statuses,
          salesCompanyRows: [],
          selectedSalesCompanyIndex: -1,
          salesMissingMstCompanies: [],
          salesProductPreviewCodes: {},
          salesProductCodeOverrides: {},
          salesReviewRows: [],
          salesReviewGenerated: false,
          matches: [],
          salesMatchGenerated: false,
        });
      }
      setStatus(`Đã tải ${summary.original_name}.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function uploadProcessed(kind: 'purchase' | 'sales', file: File | undefined) {
    if (!file) return;
    const targetProfile = profile;
    setBusy(true);
    setStatus(`Đang tải file ${kind === 'purchase' ? 'mua vào' : 'bán ra'} đã xử lý...`);
    try {
      const summary = await uploadExcel(file, `${profile}-processed-${kind}`);
      const stats = await inspectProcessedVietmaxFile(summary.saved_name, kind);
      if (kind === 'purchase') {
        updateWorkflow(targetProfile, {
          processedPurchaseSavedName: summary.saved_name,
          processedPurchaseStats: stats,
          purchaseFile: purchaseFile ?? summary,
          matches: [],
          salesMatchGenerated: false,
          salesCompanyRows: [],
          selectedSalesCompanyIndex: -1,
          salesMissingMstCompanies: [],
          salesInvoiceStatuses: [],
          salesProductPreviewCodes: {},
          salesProductCodeOverrides: {},
          salesReviewRows: [],
          salesReviewGenerated: false,
          ...salesOutputInvalidation(),
        });
      } else {
        updateWorkflow(targetProfile, {
          processedSalesSavedName: summary.saved_name,
          processedSalesStats: stats,
          salesFile: salesFile ?? summary,
          inventoryAllocationJob: null,
          inventoryAllocationResult: null,
        });
      }
      setStatus(`Đã tải ${summary.original_name}. ${processedStatsSentence(stats)}`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function uploadFastImportProcessed(kind: 'purchase' | 'sales', file: File | undefined) {
    if (!file) return;
    const targetProfile = profile;
    const label = kind === 'purchase' ? 'mua vào' : 'bán ra';
    setBusy(true);
    setStatus(`Đang tải FDI ${label} đã xử lý cho Xuất FAST...`);
    try {
      const summary = await uploadExcel(file, `${profile}-fast-${kind}`);
      setStatus(`Đã tải file lên. Đang đọc thống kê FDI ${label}...`);
      const stats = await inspectProcessedVietmaxFile(summary.saved_name, kind);
      if (kind === 'purchase') {
        updateWorkflow(targetProfile, {
          processedPurchaseSavedName: summary.saved_name,
          processedPurchaseStats: stats,
          purchaseFile: purchaseFile ?? summary,
        });
      } else {
        updateWorkflow(targetProfile, {
          processedSalesSavedName: summary.saved_name,
          processedSalesStats: stats,
          salesFile: salesFile ?? summary,
        });
      }
      setStatus(`Đã tải FDI ${label} cho Xuất FAST. ${processedStatsSentence(stats)} Khi xuất workbook, app sẽ chỉ dùng các cột mà form mapping đang cần.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function refreshProcessedFileStats(kind: 'purchase' | 'sales', savedName: string) {
    try {
      const stats = await inspectProcessedVietmaxFile(savedName, kind);
      if (kind === 'purchase') updateWorkflow(profile, { processedPurchaseStats: stats });
      else updateWorkflow(profile, { processedSalesStats: stats });
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    }
  }

  function beginProgress(label: string) {
    const operationId = `op_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    let stopped = false;
    setLoadingProgress({ operation_id: operationId, status: 'running', done: 0, total: 1, percent: 0, label });
    setStatus(formatOperationStatus({ operation_id: operationId, status: 'running', done: 0, total: 1, percent: 0, label }, label));
    const poll = async () => {
      while (!stopped) {
        await sleep(250);
        if (stopped) break;
        try {
          const progress = await getOperationProgress(operationId);
          if (!stopped) {
            setLoadingProgress(progress);
            if (progress.status !== 'missing') setStatus(formatOperationStatus(progress, label));
          }
          if (progress.status === 'complete' || progress.status === 'error') break;
        } catch {
          // Keep the visible loading state even if a single poll misses the backend.
        }
      }
    };
    void poll();
    return {
      operationId,
      stop: () => {
        stopped = true;
      },
    };
  }

  function exportSetupErrorMessage(error: unknown, phaseLabel: string) {
    const message = error instanceof Error ? error.message : String(error);
    return `${message} Hãy quay lại phần cấu hình ${phaseLabel}, sửa mapping/cột/nhóm hoặc mã kho/TK vật tư rồi bấm xuất lại.`;
  }

  function reflectWorkflowJob(job: WorkflowJob) {
    const progress: OperationProgress = {
      operation_id: job.operation_id,
      status: job.status,
      done: job.progress.done,
      total: job.progress.total,
      percent: job.progress.percent,
      label: job.progress.label,
    };
    setLoadingProgress(progress);
    if (job.status === 'queued' || job.status === 'running') {
      setStatus(formatOperationStatus(progress, job.progress.label || 'Đang xử lý'));
    }
  }

  async function runExplicitProcessJob(
    source: UploadSummary,
    processPayload: Record<string, unknown>,
    processor: 'vietmax' | 'generic',
  ) {
    const initial = await startWorkflowProcessJob({
      savedName: source.saved_name,
      originalName: source.original_name,
      processPayload,
      processor,
      retry: true,
    });
    const completed = await waitForWorkflowJob(initial, reflectWorkflowJob);
    if (completed.status === 'failed') {
      const failure = completed.error;
      const detail = failure?.details && Object.keys(failure.details).length
        ? ` Chi tiết: ${JSON.stringify(failure.details)}.`
        : '';
      throw new Error(`${failure?.message || 'Không thể tạo file.'} [${failure?.code || 'WORKFLOW_ERROR'}; ${completed.operation_id}].${detail}`);
    }
    const savedName = completed.result?.processed_saved_name || completed.result?.artifact?.saved_name || '';
    if (!savedName) throw new Error(`Job ${completed.operation_id} hoàn tất nhưng không trả về file cache.`);
    return { savedName, job: completed };
  }

  async function loadGenericProfileConfig(targetProfile: ProfileKey) {
    if (!isGenericProfileKey(targetProfile) && targetProfile !== 'ho_guom') return;
    setProfileConfigLoading(true);
    try {
      const [cfg, formatDefaults] = await Promise.all([
        getAppConfig(),
        getVietmaxFormatMappingDefaults(targetProfile).catch(() => null),
      ]);
      const formPhase = genericProfileFormPhase(targetProfile);
      const profilesCfg = (cfg.profiles && typeof cfg.profiles === 'object' ? cfg.profiles : {}) as Record<string, any>;
      if (isTwoPhaseGenericProfile(targetProfile)) {
        const purchaseCfg = scopedProfileConfig(profilesCfg, targetProfile, 'purchase');
        const salesCfg = scopedProfileConfig(profilesCfg, targetProfile, 'sales');
        const purchaseSavedColumns = purchaseCfg.columns && typeof purchaseCfg.columns === 'object' ? purchaseCfg.columns as Record<string, unknown> : {};
        const salesSavedColumns = salesCfg.columns && typeof salesCfg.columns === 'object' ? salesCfg.columns as Record<string, unknown> : {};
        updateWorkflow(targetProfile, {
          purchaseFormSourceColumns: formatSourceColumns(formatDefaults?.source_columns?.purchase),
          salesFormSourceColumns: formatSourceColumns(formatDefaults?.source_columns?.sales),
          genericColumns: normalizeGenericColumns(purchaseSavedColumns),
          purchaseColumns: normalizeVietmaxColumns(purchaseSavedColumns, 'purchase'),
          salesColumns: normalizeVietmaxColumns(salesSavedColumns, 'sales'),
          wordRules: purchaseCfg.word_rules && typeof purchaseCfg.word_rules === 'object' ? purchaseCfg.word_rules : {},
          firstWordRules: purchaseCfg.first_word_rules && typeof purchaseCfg.first_word_rules === 'object' ? purchaseCfg.first_word_rules : {},
          productCodeReplacements: cleanStringMap(purchaseCfg.product_code_replacements),
          salesProductCodeReplacements: cleanStringMap(salesCfg.product_code_replacements),
          repeatedPhraseRemovals: Array.isArray(purchaseCfg.repeated_phrase_removals) ? purchaseCfg.repeated_phrase_removals : [],
          inventoryPairs: Array.isArray(purchaseCfg.inventory_pairs) ? purchaseCfg.inventory_pairs : [],
          useDefaultInventoryPair: Boolean(purchaseCfg.use_default_inventory_pair),
          defaultInventoryPairId: String(purchaseCfg.default_inventory_pair_id || ''),
          inventoryPairRules: Array.isArray(purchaseCfg.inventory_pair_rules) ? purchaseCfg.inventory_pair_rules : [],
          purchaseWordRules: purchaseCfg.word_rules && typeof purchaseCfg.word_rules === 'object' ? purchaseCfg.word_rules : {},
          salesWordRules: salesCfg.word_rules && typeof salesCfg.word_rules === 'object' ? salesCfg.word_rules : {},
          purchaseRepeatedPhraseRemovals: Array.isArray(purchaseCfg.repeated_phrase_removals) ? purchaseCfg.repeated_phrase_removals : [],
          salesRepeatedPhraseRemovals: Array.isArray(salesCfg.repeated_phrase_removals) ? salesCfg.repeated_phrase_removals : [],
          purchaseInventoryPairs: Array.isArray(purchaseCfg.inventory_pairs) ? purchaseCfg.inventory_pairs : [],
          purchaseUseDefaultInventoryPair: Boolean(purchaseCfg.use_default_inventory_pair),
          purchaseDefaultInventoryPairId: String(purchaseCfg.default_inventory_pair_id || ''),
          purchaseInventoryPairRules: Array.isArray(purchaseCfg.inventory_pair_rules) ? purchaseCfg.inventory_pair_rules : [],
          salesInventoryPairs: Array.isArray(salesCfg.inventory_pairs) ? salesCfg.inventory_pairs : [],
          salesUseDefaultInventoryPair: Boolean(salesCfg.use_default_inventory_pair),
          salesDefaultInventoryPairId: String(salesCfg.default_inventory_pair_id || ''),
          salesInventoryPairRules: Array.isArray(salesCfg.inventory_pair_rules) ? salesCfg.inventory_pair_rules : [],
          includeCompanyPrefix: purchaseCfg.include_company_prefix !== false,
          salesIncludeCompanyPrefix: salesCfg.include_company_prefix !== false,
          purchasePrefixStrategy: normalizedPrefixStrategy(purchaseCfg.prefix_strategy || 'last_2_words'),
          salesPrefixStrategy: normalizedPrefixStrategy(salesCfg.prefix_strategy || 'last_2_words'),
          prefixMstDigits: clampPrefixMstDigits(purchaseCfg.prefix_mst_digits ?? 3),
          prefixNameWords: clampPrefixNameWords(purchaseCfg.prefix_name_words ?? 2),
          prefixNameChars: clampPrefixNameChars(purchaseCfg.prefix_name_chars ?? 1),
          prefixMissingMstStrategy: normalizeMissingMstPrefixStrategy(purchaseCfg.prefix_missing_mst_strategy),
          purchasePrefixStrategyValues: normalizePrefixStrategyValues(purchaseCfg.prefix_strategy_values, emptyPrefixStrategyValues()),
          salesPrefixStrategyValues: normalizePrefixStrategyValues(salesCfg.prefix_strategy_values, emptyPrefixStrategyValues()),
          purchaseProcessingGroups: normalizeProcessingGroups(purchaseCfg.processing_groups, 'purchase'),
          salesProcessingGroups: normalizeProcessingGroups(salesCfg.processing_groups, 'sales'),
          purchaseFormMappingPresets: normalizeFormMappingPresets(purchaseCfg.form_mapping_presets, 'purchase', formatDefaults),
          salesFormMappingPresets: normalizeFormMappingPresets(salesCfg.form_mapping_presets, 'sales', formatDefaults),
          inventoryAllocationConfig: normalizeInventoryAllocationConfig(
            profilesCfg[targetProfile]?.inventory_allocation_config || purchaseCfg.inventory_allocation_config || salesCfg.inventory_allocation_config,
            targetProfile,
          ),
          purchaseReviewRules: Array.isArray(purchaseCfg.product_review_merges) ? purchaseCfg.product_review_merges : [],
          salesReviewRules: Array.isArray(salesCfg.product_review_merges) ? salesCfg.product_review_merges : [],
        });
        return;
      }
      const profileCfg = scopedProfileConfig(profilesCfg, targetProfile, formPhase === 'sales' ? 'sales' : 'purchase');
      const globalColumns = cfg.columns && typeof cfg.columns === 'object' ? cfg.columns as Record<string, unknown> : {};
      const savedColumns = profileCfg.columns && typeof profileCfg.columns === 'object' ? profileCfg.columns as Record<string, unknown> : {};
      updateWorkflow(targetProfile, {
        purchaseFormSourceColumns: formatSourceColumns(formatDefaults?.source_columns?.purchase),
        salesFormSourceColumns: formatSourceColumns(formatDefaults?.source_columns?.sales),
        genericColumns: normalizeGenericColumns({ ...globalColumns, ...savedColumns }),
        wordRules: profileCfg.word_rules && typeof profileCfg.word_rules === 'object' ? profileCfg.word_rules : {},
        firstWordRules: profileCfg.first_word_rules && typeof profileCfg.first_word_rules === 'object' ? profileCfg.first_word_rules : {},
        productCodeReplacements: cleanStringMap(profileCfg.product_code_replacements),
        repeatedPhraseRemovals: Array.isArray(profileCfg.repeated_phrase_removals) ? profileCfg.repeated_phrase_removals : [],
        inventoryPairs: Array.isArray(profileCfg.inventory_pairs) ? profileCfg.inventory_pairs : [],
        useDefaultInventoryPair: Boolean(profileCfg.use_default_inventory_pair),
        defaultInventoryPairId: String(profileCfg.default_inventory_pair_id || ''),
        inventoryPairRules: Array.isArray(profileCfg.inventory_pair_rules) ? profileCfg.inventory_pair_rules : [],
        includeCompanyPrefix: profileCfg.include_company_prefix !== false,
        purchasePrefixStrategy: normalizedPrefixStrategy(profileCfg.prefix_strategy || 'last_2_words'),
        prefixMstDigits: clampPrefixMstDigits(profileCfg.prefix_mst_digits ?? 3),
        prefixNameWords: clampPrefixNameWords(profileCfg.prefix_name_words ?? 2),
        prefixNameChars: clampPrefixNameChars(profileCfg.prefix_name_chars ?? 1),
        prefixMissingMstStrategy: normalizeMissingMstPrefixStrategy(profileCfg.prefix_missing_mst_strategy),
        purchasePrefixStrategyValues: normalizePrefixStrategyValues(profileCfg.prefix_strategy_values, emptyPrefixStrategyValues()),
        purchaseProcessingGroups: normalizeProcessingGroups(profileCfg.processing_groups, 'generic'),
        purchaseFormMappingPresets: normalizeFormMappingPresets(profileCfg.form_mapping_presets, formPhase, formatDefaults, true),
        inventoryAllocationConfig: normalizeInventoryAllocationConfig(profileCfg.inventory_allocation_config, targetProfile),
        purchaseReviewRules: Array.isArray(profileCfg.product_review_merges) ? profileCfg.product_review_merges : [],
        priceRangeRules: profileCfg.price_range_rules && typeof profileCfg.price_range_rules === 'object' ? profileCfg.price_range_rules : {},
        priceAdjustAllPercent: Number(profileCfg.price_adjust_all_percent || 0),
      });
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setProfileConfigLoading(false);
    }
  }

  async function checkLicenseStatus() {
    setLicenseChecking(true);
    setLicenseCheckError('');
    try {
      const nextLicense = await getLicenseStatus();
      setLicense(nextLicense);
      setLicenseForm((current) => ({ ...current, server_url: nextLicense.server_url, account_id: nextLicense.account_id }));
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setLicenseCheckError(message);
      setStatus(message);
    } finally {
      setLicenseChecking(false);
    }
  }

  async function loadVietmaxProfileConfig() {
    try {
      const [cfg, formatDefaults] = await Promise.all([
        getAppConfig(),
        getVietmaxFormatMappingDefaults().catch(() => null),
      ]);
      const profilesCfg = (cfg.profiles && typeof cfg.profiles === 'object' ? cfg.profiles : {}) as Record<string, any>;
      const purchaseCfg = scopedProfileConfig(profilesCfg, 'vietmax', 'purchase');
      const salesCfg = scopedProfileConfig(profilesCfg, 'vietmax', 'sales');
      const purchaseSavedColumns = purchaseCfg.columns && typeof purchaseCfg.columns === 'object' ? purchaseCfg.columns as Record<string, unknown> : {};
      const salesSavedColumns = salesCfg.columns && typeof salesCfg.columns === 'object' ? salesCfg.columns as Record<string, unknown> : {};
      updateWorkflow('vietmax', {
        purchaseFormSourceColumns: formatSourceColumns(formatDefaults?.source_columns?.purchase),
        salesFormSourceColumns: formatSourceColumns(formatDefaults?.source_columns?.sales),
        purchaseColumns: normalizeVietmaxColumns(purchaseSavedColumns, 'purchase'),
        salesColumns: normalizeVietmaxColumns(salesSavedColumns, 'sales'),
        purchaseWordRules: purchaseCfg.word_rules && typeof purchaseCfg.word_rules === 'object' ? purchaseCfg.word_rules : {},
        salesWordRules: salesCfg.word_rules && typeof salesCfg.word_rules === 'object' ? salesCfg.word_rules : {},
        purchaseRepeatedPhraseRemovals: Array.isArray(purchaseCfg.repeated_phrase_removals) ? purchaseCfg.repeated_phrase_removals : [],
        salesRepeatedPhraseRemovals: Array.isArray(salesCfg.repeated_phrase_removals) ? salesCfg.repeated_phrase_removals : [],
        productCodeReplacements: cleanStringMap(purchaseCfg.product_code_replacements),
        salesProductCodeReplacements: cleanStringMap(salesCfg.product_code_replacements),
        purchaseInventoryPairs: Array.isArray(purchaseCfg.inventory_pairs) ? purchaseCfg.inventory_pairs : [],
        purchaseUseDefaultInventoryPair: Boolean(purchaseCfg.use_default_inventory_pair),
        purchaseDefaultInventoryPairId: String(purchaseCfg.default_inventory_pair_id || ''),
        purchaseInventoryPairRules: Array.isArray(purchaseCfg.inventory_pair_rules) ? purchaseCfg.inventory_pair_rules : [],
        salesInventoryPairs: Array.isArray(salesCfg.inventory_pairs) ? salesCfg.inventory_pairs : [],
        salesUseDefaultInventoryPair: Boolean(salesCfg.use_default_inventory_pair),
        salesDefaultInventoryPairId: String(salesCfg.default_inventory_pair_id || ''),
        salesInventoryPairRules: Array.isArray(salesCfg.inventory_pair_rules) ? salesCfg.inventory_pair_rules : [],
        includeCompanyPrefix: purchaseCfg.include_company_prefix !== false,
        salesIncludeCompanyPrefix: salesCfg.include_company_prefix !== false,
        purchasePrefixStrategy: normalizedPrefixStrategy(purchaseCfg.prefix_strategy || 'last_2_words'),
        salesPrefixStrategy: normalizedPrefixStrategy(salesCfg.prefix_strategy || 'last_2_words'),
        prefixMstDigits: clampPrefixMstDigits(purchaseCfg.prefix_mst_digits ?? 3),
        prefixNameWords: clampPrefixNameWords(purchaseCfg.prefix_name_words ?? 2),
        prefixNameChars: clampPrefixNameChars(purchaseCfg.prefix_name_chars ?? 1),
        prefixMissingMstStrategy: normalizeMissingMstPrefixStrategy(purchaseCfg.prefix_missing_mst_strategy),
        purchasePrefixStrategyValues: normalizePrefixStrategyValues(purchaseCfg.prefix_strategy_values, emptyPrefixStrategyValues()),
        salesPrefixStrategyValues: normalizePrefixStrategyValues(salesCfg.prefix_strategy_values, emptyPrefixStrategyValues()),
        purchaseProcessingGroups: normalizeProcessingGroups(purchaseCfg.processing_groups, 'purchase'),
        salesProcessingGroups: normalizeProcessingGroups(salesCfg.processing_groups, 'sales'),
        purchaseFormMappingPresets: normalizeFormMappingPresets(purchaseCfg.form_mapping_presets, 'purchase', formatDefaults),
        salesFormMappingPresets: normalizeFormMappingPresets(salesCfg.form_mapping_presets, 'sales', formatDefaults),
        inventoryAllocationConfig: normalizeInventoryAllocationConfig(purchaseCfg.inventory_allocation_config || salesCfg.inventory_allocation_config, 'vietmax'),
        purchaseReviewRules: Array.isArray(purchaseCfg.vietmax_mua_vao_internal_merges) ? purchaseCfg.vietmax_mua_vao_internal_merges : [],
        salesReviewRules: Array.isArray(salesCfg.vietmax_ban_ra_sales_internal_merges) ? salesCfg.vietmax_ban_ra_sales_internal_merges : [],
        salesMatchRules: Array.isArray(salesCfg.vietmax_ban_ra_purchase_match_rules) ? salesCfg.vietmax_ban_ra_purchase_match_rules : [],
      });
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    }
  }

  async function loadGenericCompanies(nextStage?: StageId) {
    if (!purchaseFile || !isGenericWorkflowProfile) return;
    const targetProfile = profile;
    const targetFile = purchaseFile;
    setBusy(true);
    setStatus(`Đang tải danh sách công ty và hàng hóa ${selectedProfile.label}...`);
    try {
      const twoPhase = isTwoPhaseGenericProfile(targetProfile);
      const sourceColumns = twoPhase ? normalizeVietmaxColumns(purchaseColumns, 'purchase') : normalizeGenericColumns(genericColumns);
      const [result, formatDefaults] = await Promise.all([
        analyzeGenericWorkbook({
          saved_name: targetFile.saved_name,
          original_name: targetFile.original_name,
          profile: targetProfile,
          vietmax_phase: 'purchase',
          ...sourceColumns,
        }),
        getVietmaxFormatMappingDefaults(targetProfile).catch(() => null),
      ]);
      const savedWordRules = result.word_rules ?? (twoPhase ? purchaseWordRules : wordRules);
      const savedFirstWordRules = result.first_word_rules ?? firstWordRules;
      const savedProductCodeReplacements = result.product_code_replacements ?? productCodeReplacements;
      const savedRepeatedPhrases = result.repeated_phrase_removals ?? (twoPhase ? purchaseRepeatedPhraseRemovals : repeatedPhraseRemovals);
      const savedInventoryPairs = result.inventory_pairs ?? (twoPhase ? purchaseInventoryPairs : inventoryPairs);
      const loadedGroups = normalizeProcessingGroups(result.processing_groups, twoPhase ? 'purchase' : 'generic');
      const loadedFormPresets = normalizeFormMappingPresets(result.form_mapping_presets, twoPhase ? 'purchase' : genericProfileFormPhase(targetProfile), formatDefaults, !twoPhase);
      const nextCompanies = applyGroupAssignments(result.companies.map((company) => ({
        ...company,
        process: company.process ?? true,
        pending_process: company.pending_process ?? company.process ?? true,
        committed_prefix: company.committed_prefix ?? company.value ?? '',
        selected_product_names: company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name),
      })), result.company_group_assignments ?? {});
      const loadedPrefixStrategy = normalizedPrefixStrategy(result.prefix_strategy || purchasePrefixStrategy);
      const loadedPrefixMstDigits = clampPrefixMstDigits(result.prefix_mst_digits ?? prefixMstDigits);
      const loadedPrefixNameWords = clampPrefixNameWords(result.prefix_name_words ?? prefixNameWords);
      const loadedMissingMstPrefixStrategy = normalizeMissingMstPrefixStrategy(result.prefix_missing_mst_strategy ?? prefixMissingMstStrategy);
      const loadedPrefixValues = normalizePrefixStrategyValues(result.prefix_strategy_values, purchasePrefixStrategyValues);
      const loadedPrefixNameChars = clampPrefixNameChars(result.prefix_name_chars ?? prefixNameChars);
      const nextPrefixValues = seedLoadedPrefixValues(loadedPrefixValues, loadedPrefixStrategy, nextCompanies, loadedPrefixMstDigits, loadedPrefixNameWords, loadedPrefixNameChars, loadedMissingMstPrefixStrategy);
      const displayCompanies = applyPrefixStrategyRows(nextCompanies, loadedPrefixStrategy, loadedPrefixMstDigits, loadedPrefixNameWords, loadedPrefixNameChars, nextPrefixValues, true, loadedMissingMstPrefixStrategy);
      const previewCodes = await loadGenericProductPreviewCodes(targetProfile, displayCompanies, savedWordRules, savedFirstWordRules, savedRepeatedPhrases, savedProductCodeReplacements);
      updateWorkflow(targetProfile, {
        ...purchaseOutputInvalidation(),
        genericColumns: normalizeGenericColumns(sourceColumns),
        ...(twoPhase ? { purchaseColumns: normalizeVietmaxColumns(sourceColumns, 'purchase') } : {}),
        companyRows: displayCompanies,
        selectedCompanyIndex: firstDisplayedCompanyIndex(displayCompanies, loadedGroups),
        purchaseMissingMstCompanies: result.missing_mst_companies ?? [],
        productPreviewCodes: previewCodes,
        productCodeOverrides: result.manual_code_overrides ?? {},
        productCodeReplacements: savedProductCodeReplacements,
        wordRules: savedWordRules,
        firstWordRules: savedFirstWordRules,
        repeatedPhraseRemovals: savedRepeatedPhrases,
        inventoryPairs: savedInventoryPairs,
        useDefaultInventoryPair: result.use_default_inventory_pair ?? useDefaultInventoryPair,
        defaultInventoryPairId: result.default_inventory_pair_id ?? defaultInventoryPairId,
        inventoryPairRules: result.inventory_pair_rules ?? inventoryPairRules,
        ...(twoPhase ? {
          purchaseWordRules: savedWordRules,
          productCodeReplacements: savedProductCodeReplacements,
          purchaseRepeatedPhraseRemovals: savedRepeatedPhrases,
          purchaseInventoryPairs: savedInventoryPairs,
          purchaseUseDefaultInventoryPair: result.use_default_inventory_pair ?? purchaseUseDefaultInventoryPair,
          purchaseDefaultInventoryPairId: result.default_inventory_pair_id ?? purchaseDefaultInventoryPairId,
          purchaseInventoryPairRules: result.inventory_pair_rules ?? purchaseInventoryPairRules,
        } : {}),
        includeCompanyPrefix: result.include_company_prefix ?? includeCompanyPrefix,
        purchasePrefixStrategy: loadedPrefixStrategy,
        prefixMstDigits: loadedPrefixMstDigits,
        prefixNameWords: loadedPrefixNameWords,
        prefixNameChars: loadedPrefixNameChars,
        prefixMissingMstStrategy: loadedMissingMstPrefixStrategy,
        purchasePrefixStrategyValues: nextPrefixValues,
        purchaseProcessingGroups: loadedGroups,
        purchaseFormMappingPresets: loadedFormPresets,
        purchaseReviewRules: Array.isArray(result.product_review_merges) ? result.product_review_merges as ReviewRow[] : purchaseReviewRules,
        priceRangeRules: result.price_range_rules ?? priceRangeRules,
        priceAdjustAllPercent: Number(result.price_adjust_all_percent ?? priceAdjustAllPercent),
        priceGroups: [],
        purchaseReviewRows: [],
        purchaseReviewGenerated: false,
        ...(nextStage ? { stage: nextStage } : {}),
      });
      setStatus(`Đã tải ${result.company_count} công ty, ${result.rows_to_process} dòng ${selectedProfile.label}. Chọn công ty/hàng hóa rồi áp dụng trước khi xuất file.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function updateGenericColumns(update: Partial<GenericColumns>) {
    updateWorkflow(profile, {
      ...purchaseOutputInvalidation(),
      genericColumns: normalizeGenericColumns({ ...genericColumns, ...update }),
      companyRows: [],
      selectedCompanyIndex: -1,
      purchaseMissingMstCompanies: [],
      productPreviewCodes: {},
      productCodeOverrides: {},
      purchaseReviewRows: [],
      purchaseReviewGenerated: false,
      priceGroups: [],
    });
  }

  function updateVietmaxColumns(phase: 'purchase' | 'sales', update: Partial<GenericColumns>) {
    const currentColumns = phase === 'sales' ? salesColumns : purchaseColumns;
    const nextColumns = normalizeVietmaxColumns({ ...currentColumns, ...update }, phase);
    const currentStatuses = phase === 'sales' ? salesInvoiceStatuses : purchaseInvoiceStatuses;
    const nextStatuses = withInvoiceSkipFlags(currentStatuses, nextColumns.invoice_status_skip_values);
    if (phase === 'sales') {
      updateWorkflow(profile, {
        ...salesOutputInvalidation(),
        salesColumns: nextColumns,
        salesInvoiceStatuses: nextStatuses,
        salesCompanyRows: [],
        selectedSalesCompanyIndex: -1,
        salesMissingMstCompanies: [],
        salesProductPreviewCodes: {},
        salesProductCodeOverrides: {},
        salesReviewRows: [],
        salesReviewGenerated: false,
        matches: [],
        salesMatchGenerated: false,
      });
    } else {
      updateWorkflow(profile, {
        ...purchaseOutputInvalidation(),
        purchaseColumns: nextColumns,
        purchaseInvoiceStatuses: nextStatuses,
        companyRows: [],
        selectedCompanyIndex: -1,
        purchaseMissingMstCompanies: [],
        productPreviewCodes: {},
        productCodeOverrides: {},
        purchaseReviewRows: [],
        purchaseReviewGenerated: false,
        matches: [],
        salesMatchGenerated: false,
        salesCompanyRows: [],
        selectedSalesCompanyIndex: -1,
        salesMissingMstCompanies: [],
      });
    }
    if (Object.prototype.hasOwnProperty.call(update, 'invoice_status_col')) {
      void refreshVietmaxInvoiceStatuses(phase, nextColumns);
    }
  }

  function updateVietmaxInvoiceStatusSkipValues(phase: 'purchase' | 'sales', values: string[]) {
    updateVietmaxColumns(phase, { invoice_status_skip_values: values });
  }

  async function refreshVietmaxInvoiceStatuses(phase: 'purchase' | 'sales', columns: GenericColumns) {
    const file = phase === 'sales' ? salesFile : purchaseFile;
    const statusColumn = String(columns.invoice_status_col || '').trim().toUpperCase();
    if (!file || !statusColumn) {
      updateWorkflow(profile, phase === 'sales' ? { salesInvoiceStatuses: [] } : { purchaseInvoiceStatuses: [] });
      return;
    }
    try {
      const statuses = await fetchInvoiceStatuses(file.saved_name, statusColumn, columns.invoice_status_skip_values);
      updateWorkflow(profile, phase === 'sales' ? { salesInvoiceStatuses: statuses } : { purchaseInvoiceStatuses: statuses });
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    }
  }

  async function downloadGenericProcessedFile() {
    if (!purchaseFile || !isGenericWorkflowProfile) return;
    const dirtyCompanies = companyRows.filter(hasCompanyDraftChanges);
    if (dirtyCompanies.length) {
      setStatus('Đang có thay đổi lọc công ty/prefix chưa áp dụng. Bấm Áp dụng lựa chọn công ty và hàng hóa trước khi xuất file.');
      return;
    }
    setBusy(true);
    const exportFormMappings = !isTwoPhaseGenericProfile(profile);
    setStatus(processedPurchaseSavedName ? `Đang tải file ${selectedProfile.label} từ cache...` : `Đang tạo cache file ${selectedProfile.label}...`);
    try {
      let savedName = processedPurchaseSavedName;
      if (!savedName) {
        const result = await runExplicitProcessJob(purchaseFile, {
          export_form_mappings: exportFormMappings,
          ...buildGenericProcessPayload(workflow, profile),
        }, 'generic');
        savedName = result.savedName;
        updateWorkflow(profile, { processedPurchaseSavedName: savedName });
      }
      const blob = await downloadCachedFile(savedName);
      const filename = `${fileStem(purchaseFile.original_name)}_${profile}_${isTwoPhaseGenericProfile(profile) ? 'mua_vao_fdi' : 'fast'}.xls`;
      const saved = await saveBlob(blob, filename);
      setStatus(saved ? `Đã xuất workbook FDI và form mapping ${selectedProfile.label}.` : 'Đã hủy lưu file kết quả.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function downloadGenericSalesProcessedFile() {
    if (!salesFile || !isTwoPhaseGenericProfile(profile)) return;
    if (salesCompanyRows.some(hasCompanyDraftChanges)) {
      setStatus('Đang có thay đổi lọc công ty/prefix bán ra chưa áp dụng. Bấm Áp dụng lựa chọn trước khi xuất file.');
      return;
    }
    setBusy(true);
    try {
      let savedName = processedSalesSavedName;
      if (!savedName) {
        const result = await runExplicitProcessJob(salesFile, {
          export_form_mappings: false,
          ...buildGenericSalesProcessPayload(workflow, profile),
        }, 'generic');
        savedName = result.savedName;
        updateWorkflow(profile, { processedSalesSavedName: savedName });
      }
      const blob = await downloadCachedFile(savedName);
      const filename = `${fileStem(salesFile.original_name)}_${profile}_ban_ra_fdi.xls`;
      const saved = await saveBlob(blob, filename);
      setStatus(saved ? `Đã lưu file bán ra ${selectedProfile.label}.` : 'Đã hủy lưu file bán ra.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function runPurchaseReview() {
    if (!purchaseFile) return;
    const dirtyCompanies = companyRows.filter(hasCompanyDraftChanges);
    if (dirtyCompanies.length) {
      setStatus('Đang có thay đổi lọc công ty chưa áp dụng. Bấm Áp dụng lựa chọn công ty và hàng hóa trước khi review.');
      return;
    }
    const reviewProducts = buildPurchaseReviewProducts(workflow);
    const targetProfile = profile;
    const targetPurchaseFile = purchaseFile;
    const scope = purchaseReviewScope === 'company' ? 'same_company' : 'all_companies';
    const progress = beginProgress('Đang chuẩn bị review Mã VT mua vào');
    setBusy(true);
    setStatus('Đang tạo review Mã VT mua vào bằng logic Vietmax...');
    try {
      const result = await createPurchaseReview(targetPurchaseFile.saved_name, scope, purchaseWordRules, purchaseRepeatedPhraseRemovals, reviewProducts, progress.operationId);
      updateWorkflow(targetProfile, { ...purchaseOutputInvalidation(), purchaseReviewRows: normalizeReviewRows(result.review_rows as ReviewRow[]), purchaseReviewGenerated: true, stage: 4 });
      setStatus(`Đã tạo ${result.review_rows.length} dòng review Mã VT mua vào. Chỉ dòng được tick ở cột Dùng và lựa chọn Dùng mã sẽ được áp dụng khi xử lý.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function runGenericReview() {
    if (!purchaseFile || !isGenericWorkflowProfile) return;
    const dirtyCompanies = companyRows.filter(hasCompanyDraftChanges);
    if (dirtyCompanies.length) {
      setStatus('Dang co thay doi loc cong ty/prefix chua ap dung. Bam Ap dung lua chon cong ty va hang hoa truoc khi review.');
      return;
    }
    const reviewProducts = buildPurchaseReviewProducts(workflow);
    const targetProfile = profile;
    const targetPurchaseFile = purchaseFile;
    const scope = purchaseReviewScope === 'company' ? 'same_company' : 'all_companies';
    const progress = beginProgress(`Dang chuan bi review Ma VT ${selectedProfile.label}`);
    setBusy(true);
    setStatus(`Dang tao review Ma VT ${selectedProfile.label}...`);
    try {
      const result = await createGenericReview(
        targetPurchaseFile.saved_name,
        targetProfile,
        scope,
        isTwoPhaseGenericProfile(targetProfile) ? purchaseWordRules : wordRules,
        firstWordRules,
        isTwoPhaseGenericProfile(targetProfile) ? purchaseRepeatedPhraseRemovals : repeatedPhraseRemovals,
        reviewProducts,
        progress.operationId,
        'purchase',
      );
      updateWorkflow(targetProfile, { ...purchaseOutputInvalidation(), purchaseReviewRows: normalizeReviewRows(result.review_rows as ReviewRow[]), purchaseReviewGenerated: true, priceGroups: [], stage: 4 });
      setStatus(`Da tao ${result.review_rows.length} dong Review Ma VT ${selectedProfile.label}. Ap dung review truoc khi sang buoc tiep theo.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  function updateCaoThanhPriceGroups() {
    if (profile !== 'cao_thanh') return;
    const nextGroups = buildCaoThanhPriceGroups(
      companyRows,
      productPreviewCodes,
      productCodeOverrides,
      productCodeReplacements,
      includeCompanyPrefix,
      priceRangeRules,
      priceAdjustAllPercent,
    );
    updateWorkflow(profile, { priceGroups: nextGroups });
    setStatus(`Da tao ${nextGroups.length} nhom loc don gia Cao Thanh.`);
  }

  function updateCaoThanhGroupPercent(groupKey: string, value: number) {
    updateWorkflow(profile, {
      priceGroups: priceGroups.map((group) => (
        group.key === groupKey
          ? rebuildCaoThanhPriceGroup({ ...group, filterPercent: clampPercent(value, group.filterPercent || 8) }, priceAdjustAllPercent)
          : group
      )),
    });
  }

  function updateCaoThanhBucketMargin(groupKey: string, bucketKey: string, value: number) {
    updateWorkflow(profile, {
      priceGroups: priceGroups.map((group) => (
        group.key === groupKey
          ? {
              ...group,
              buckets: group.buckets.map((bucket) => (
                bucket.key === bucketKey
                  ? rebuildCaoThanhBucket({ ...bucket, marginPercent: clampPercent(value, bucket.marginPercent || 0) })
                  : bucket
              )),
            }
          : group
      )),
    });
  }

  function updateCaoThanhPriceFilterAllPercent(value: number) {
    updateWorkflow(profile, { priceFilterAllPercent: clampPercent(value, priceFilterAllPercent || 8) || 8 });
  }

  function updateCaoThanhPriceAdjustAllPercent(value: number) {
    updateWorkflow(profile, { priceAdjustAllPercent: clampPercent(value, priceAdjustAllPercent || 0) });
  }

  function applyCaoThanhBulkPriceFilter() {
    updateWorkflow(profile, {
      priceGroups: priceGroups.map((group) => rebuildCaoThanhPriceGroup({ ...group, filterPercent: clampPercent(priceFilterAllPercent, 8) || 8 }, priceAdjustAllPercent)),
    });
  }

  function applyCaoThanhBulkMargin() {
    updateWorkflow(profile, {
      priceGroups: priceGroups.map((group) => ({
        ...group,
        buckets: group.buckets.map((bucket) => rebuildCaoThanhBucket({ ...bucket, marginPercent: clampPercent(priceAdjustAllPercent, 0) })),
      })),
    });
  }

  async function exportCaoThanhPriceReport() {
    if (profile !== 'cao_thanh' || !priceGroups.length) return;
    setBusy(true);
    setStatus('Dang xuat bao cao loc gia Cao Thanh...');
    try {
      const payload = caoThanhPriceReportPayload(purchaseFile?.original_name || 'cao_thanh.xlsx', priceGroups);
      const blob = await exportPriceReportWorkbook(payload);
      const saved = await saveBlob(blob, caoThanhReportFileName(purchaseFile?.original_name || 'cao_thanh.xlsx'));
      setStatus(saved ? 'Da xuat bao cao loc gia Cao Thanh.' : 'Da huy luu bao cao loc gia Cao Thanh.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function applyProcessedPurchaseCache(targetProfile: ProfileKey, processedSavedName: string) {
    if (!processedSavedName) return null;
    const stats = await inspectProcessedVietmaxFile(processedSavedName, 'purchase');
    updateWorkflow(targetProfile, {
      ...salesOutputInvalidation(),
      processedPurchaseSavedName: processedSavedName,
      processedPurchaseStats: stats,
      matches: [],
      salesMatchGenerated: false,
      salesCompanyRows: [],
      selectedSalesCompanyIndex: -1,
      salesProductPreviewCodes: {},
      salesProductCodeOverrides: {},
      salesReviewRows: [],
      salesReviewGenerated: false,
    });
    return stats;
  }

  async function runSalesMatch() {
    if (!salesFile) return;
    if (!processedPurchaseSavedName) {
      setStatus('Chưa có file mua vào đã xử lý. Hãy quay lại stage 5, bấm Tạo file mua vào thành công rồi mới chạy Khớp mua vào / bán ra.');
      return;
    }
    const targetProfile = profile;
    const targetSalesFile = salesFile;
    const targetProcessedPurchase = processedPurchaseSavedName;
    const startedAt = Date.now();
    const debugBase = () => `Debug khớp: sales=${targetSalesFile.original_name || targetSalesFile.saved_name}, purchase=${targetProcessedPurchase}, scope=${comparisonScope}, salesProduct=${salesColumns.product_col || 'M'}, salesQty=${salesColumns.qty_col || 'O'}, salesStatus=${salesColumns.invoice_status_col || 'AJ'}, purchasePrice=${purchaseColumns.price_col || 'P'}.`;
    const progress = beginProgress('Đang chuẩn bị khớp mua vào / bán ra');
    setBusy(true);
    setStatus('Đang khớp bán ra với file mua vào đã xử lý KVT/152...');
    try {
      const result = await createSalesMatches(targetSalesFile.saved_name, targetProcessedPurchase, comparisonScope, progress.operationId, { ...salesColumns, purchase_price_col: purchaseColumns.price_col || 'P' });
      const savedRules = result.match_rules?.length ? result.match_rules : salesMatchRules;
      const nextMatches = applySalesMatchRules(result.matches, savedRules, comparisonScope);
      const mismatchCount = nextMatches.filter(hasUnitMismatch).length;
      const elapsed = ((Date.now() - startedAt) / 1000).toFixed(1);
      const debug = `Debug khớp: đọc ${result.sales_products.length} hàng bán, ${result.purchase_products.length} hàng mua; exact ${result.exact_matches.length}; gợi ý ${result.matches.length}; sau rule ${nextMatches.length}; thời gian ${elapsed}s.`;
      updateWorkflow(targetProfile, { ...salesOutputInvalidation(), matches: nextMatches, salesMatchGenerated: true, salesMatchRules: savedRules, salesCompanyRows: [], selectedSalesCompanyIndex: -1, salesMissingMstCompanies: [], salesProductPreviewCodes: {}, salesProductCodeOverrides: {}, salesReviewRows: [], salesReviewGenerated: false, stage: 8 });
      setStatus(`Đã gợi ý ${result.matches.length} dòng khớp. ${result.exact_matches.length} dòng lấy chính xác từ KVT/152. ${mismatchCount} dòng khác ĐVT. ${debug}`);
    } catch (error) {
      const elapsed = ((Date.now() - startedAt) / 1000).toFixed(1);
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`${message} ${debugBase()} thời gian ${elapsed}s. Bấm Khớp lại nếu muốn thử lại.`);
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function downloadMatches() {
    setBusy(true);
    setStatus('Đang xuất Excel danh sách khớp...');
    try {
      const blob = await exportMatches(selectedMatches);
      const saved = await saveBlob(blob, 'vietmax_khop_mua_ban.xls');
      setStatus(saved ? 'Đã xuất Excel danh sách khớp mua vào/bán ra.' : 'Đã hủy lưu danh sách khớp.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function toggleMatch(index: number, confirmed: boolean) {
    updateWorkflow(profile, {
      matches: matches.map((match, rowIndex) => (rowIndex === index ? { ...match, confirmed } : match)),
      salesCompanyRows: [],
      selectedSalesCompanyIndex: -1,
      salesProductPreviewCodes: {},
      salesProductCodeOverrides: {},
      salesReviewRows: [],
      salesReviewGenerated: false,
      ...salesOutputInvalidation(),
    });
  }

  function bulkToggleMatches(confirmed: boolean) {
    updateWorkflow(profile, {
      matches: matches.map((match) => ({ ...match, confirmed })),
      salesCompanyRows: [],
      selectedSalesCompanyIndex: -1,
      salesProductPreviewCodes: {},
      salesProductCodeOverrides: {},
      salesReviewRows: [],
      salesReviewGenerated: false,
      ...salesOutputInvalidation(),
    });
  }

  function updateMatchConversion(index: number, salesQty: string, purchaseQty: string) {
    updateWorkflow(profile, {
      matches: matches.map((match, rowIndex) => {
        if (rowIndex !== index) return match;
        const conversion_formula = salesQty.trim() && purchaseQty.trim() ? conversionFormula(match, salesQty, purchaseQty) : '';
        return { ...match, conversion_formula, conversion_mode: conversion_formula ? 'qty_and_unit' : 'none' };
      }),
      ...salesOutputInvalidation(),
    });
  }

  function saveMatchChoices() {
    const nextWorkflow = { ...workflow, salesMatchRules: buildSalesMatchRules(workflow) };
    updateWorkflow(profile, nextWorkflow);
    void saveWorkflowConfig(nextWorkflow, 'Đã lưu cấu hình khớp mua/bán theo công ty, hàng hóa và công thức quy đổi ĐVT.', 'sales');
  }

  function updateReviewRow(index: number, update: Partial<ReviewRow>) {
    updateWorkflow(profile, {
      ...purchaseOutputInvalidation(),
      purchaseReviewRows: purchaseReviewRows.map((row, rowIndex) => (rowIndex === index ? { ...row, ...update } : row)),
    });
  }

  function updateSalesReviewRow(index: number, update: Partial<ReviewRow>) {
    updateWorkflow(profile, {
      ...salesOutputInvalidation(),
      salesReviewRows: salesReviewRows.map((row, rowIndex) => (rowIndex === index ? { ...row, ...update } : row)),
    });
  }

  function bulkUpdateReviewRows(indices: number[], confirmed: boolean) {
    const targetIndices = new Set(indices);
    updateWorkflow(profile, {
      ...purchaseOutputInvalidation(),
      purchaseReviewRows: purchaseReviewRows.map((row, rowIndex) => (targetIndices.has(rowIndex) ? { ...row, confirmed } : row)),
    });
  }

  function bulkUpdateSalesReviewRows(indices: number[], confirmed: boolean) {
    const targetIndices = new Set(indices);
    updateWorkflow(profile, {
      ...salesOutputInvalidation(),
      salesReviewRows: salesReviewRows.map((row, rowIndex) => (targetIndices.has(rowIndex) ? { ...row, confirmed } : row)),
    });
  }

  function applyReviewChoices() {
    const scope = reviewScopeValue(purchaseReviewScope);
    const nextWorkflow = {
      ...workflow,
      ...purchaseOutputInvalidation(),
      purchaseReviewRows,
      purchaseReviewRules: buildReviewRules(purchaseReviewRules, purchaseReviewRows, scope),
      priceGroups: [],
    };
    updateWorkflow(profile, nextWorkflow);
    void saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu lựa chọn Review Mã VT vào cấu hình.', 'purchase');
  }

  function applySalesReviewChoices() {
    const scope = reviewScopeValue(salesReviewScope);
    const nextWorkflow = {
      ...workflow,
      ...salesOutputInvalidation(),
      salesReviewRows,
      salesReviewRules: buildReviewRules(salesReviewRules, salesReviewRows, scope),
    };
    updateWorkflow(profile, nextWorkflow);
    void saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu lựa chọn Review bán ra. Các dòng được tick sẽ được tính vào file bán ra khi xuất.', 'sales');
  }

  async function loadCompanies() {
    if (!purchaseFile) return;
    const targetProfile = profile;
    setBusy(true);
    setStatus('Đang tải danh sách công ty và hàng hóa mua vào...');
    try {
      const result = await analyzeVietmaxCompanies(purchaseFile.saved_name, 'purchase', purchaseColumns);
      const savedWordRules = result.word_rules ?? purchaseWordRules;
      const savedProductCodeReplacements = result.product_code_replacements ?? productCodeReplacements;
      const savedRepeatedPhrases = result.repeated_phrase_removals ?? purchaseRepeatedPhraseRemovals;
      const savedInventoryPairs = result.inventory_pairs ?? purchaseInventoryPairs;
      const loadedGroups = normalizeProcessingGroups(result.processing_groups, 'purchase');
      const loadedFormPresets = normalizeFormMappingPresets(result.form_mapping_presets, 'purchase');
      const nextCompanies = applyGroupAssignments(result.companies.map((company) => ({
        ...company,
        process: company.process ?? true,
        pending_process: company.pending_process ?? company.process ?? true,
        committed_prefix: company.committed_prefix ?? company.value ?? '',
        selected_product_names: company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name),
      })), result.company_group_assignments ?? {});
      const loadedPrefixStrategy = normalizedPrefixStrategy(result.prefix_strategy || purchasePrefixStrategy);
      const loadedPrefixMstDigits = clampPrefixMstDigits(result.prefix_mst_digits ?? prefixMstDigits);
      const loadedPrefixNameWords = clampPrefixNameWords(result.prefix_name_words ?? prefixNameWords);
      const loadedPrefixNameChars = clampPrefixNameChars(result.prefix_name_chars ?? prefixNameChars);
      const loadedMissingMstPrefixStrategy = normalizeMissingMstPrefixStrategy(result.prefix_missing_mst_strategy ?? prefixMissingMstStrategy);
      const loadedPrefixValues = normalizePrefixStrategyValues(result.prefix_strategy_values, purchasePrefixStrategyValues);
      const nextPrefixValues = seedLoadedPrefixValues(loadedPrefixValues, loadedPrefixStrategy, nextCompanies, loadedPrefixMstDigits, loadedPrefixNameWords, loadedPrefixNameChars, loadedMissingMstPrefixStrategy);
      const displayCompanies = applyPrefixStrategyRows(nextCompanies, loadedPrefixStrategy, loadedPrefixMstDigits, loadedPrefixNameWords, loadedPrefixNameChars, nextPrefixValues, true, loadedMissingMstPrefixStrategy);
      const previewCodes = await loadProductPreviewCodes(displayCompanies, savedWordRules, savedRepeatedPhrases, 'purchase', savedProductCodeReplacements);
      updateWorkflow(targetProfile, { ...purchaseOutputInvalidation(), companyRows: displayCompanies, selectedCompanyIndex: firstDisplayedCompanyIndex(displayCompanies, loadedGroups), purchaseMissingMstCompanies: result.missing_mst_companies ?? [], productPreviewCodes: previewCodes, productCodeOverrides: result.manual_code_overrides ?? {}, productCodeReplacements: savedProductCodeReplacements, purchaseWordRules: savedWordRules, purchaseRepeatedPhraseRemovals: savedRepeatedPhrases, purchaseInventoryPairs: savedInventoryPairs, purchaseUseDefaultInventoryPair: result.use_default_inventory_pair ?? purchaseUseDefaultInventoryPair, purchaseDefaultInventoryPairId: result.default_inventory_pair_id ?? purchaseDefaultInventoryPairId, purchaseInventoryPairRules: result.inventory_pair_rules ?? purchaseInventoryPairRules, includeCompanyPrefix: result.include_company_prefix ?? includeCompanyPrefix, purchasePrefixStrategy: loadedPrefixStrategy, prefixMstDigits: loadedPrefixMstDigits, prefixNameWords: loadedPrefixNameWords, prefixNameChars: loadedPrefixNameChars, prefixMissingMstStrategy: loadedMissingMstPrefixStrategy, purchasePrefixStrategyValues: nextPrefixValues, purchaseProcessingGroups: loadedGroups, purchaseFormMappingPresets: loadedFormPresets, purchaseReviewRules: result.vietmax_mua_vao_internal_merges ?? purchaseReviewRules, purchaseReviewRows: [], purchaseReviewGenerated: false });
      setStatus(`Đã tải ${result.company_count} công ty, ${result.rows_to_process} dòng mua vào. Chọn một dòng công ty để xem hàng hóa và Mã VT preview.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function loadSalesCompanies() {
    if (!salesFile) return;
    const targetProfile = profile;
    setBusy(true);
    setStatus('Đang tải danh sách công ty và hàng hóa bán ra...');
    try {
      const result = await analyzeVietmaxCompanies(salesFile.saved_name, 'sales', salesColumns);
      const savedWordRules = result.word_rules ?? salesWordRules;
      const savedProductCodeReplacements = result.product_code_replacements ?? salesProductCodeReplacements;
      const savedRepeatedPhrases = result.repeated_phrase_removals ?? salesRepeatedPhraseRemovals;
      const savedInventoryPairs = result.inventory_pairs?.length ? result.inventory_pairs : salesInventoryPairs;
      const khhMatchedKeys = confirmedSalesMatchKeys(matches, comparisonScope);
      const loadedGroups = normalizeProcessingGroups(result.processing_groups, 'sales');
      const loadedFormPresets = normalizeFormMappingPresets(result.form_mapping_presets, 'sales');
      const nextCompanies = applyGroupAssignments(result.companies.map((company) => ({
        ...company,
        all_products: company.all_products.filter((product) => !khhMatchedKeys.has(salesProductMatchKey(product.name, company.company, company.mst, comparisonScope))),
        process: company.process ?? true,
        pending_process: company.pending_process ?? company.process ?? true,
        committed_prefix: company.committed_prefix ?? company.value ?? '',
      })).map((company) => ({
        ...company,
        selected_product_names: (company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name)).filter((name) => company.all_products.some((product) => product.name === name)),
      })).filter((company) => company.all_products.length), result.company_group_assignments ?? {});
      const loadedPrefixStrategy = normalizedPrefixStrategy(result.prefix_strategy || salesPrefixStrategy);
      const loadedPrefixMstDigits = clampPrefixMstDigits(result.prefix_mst_digits ?? prefixMstDigits);
      const loadedPrefixNameWords = clampPrefixNameWords(result.prefix_name_words ?? prefixNameWords);
      const loadedPrefixNameChars = clampPrefixNameChars(result.prefix_name_chars ?? prefixNameChars);
      const loadedMissingMstPrefixStrategy = normalizeMissingMstPrefixStrategy(result.prefix_missing_mst_strategy ?? prefixMissingMstStrategy);
      const loadedPrefixValues = normalizePrefixStrategyValues(result.prefix_strategy_values, salesPrefixStrategyValues);
      const nextPrefixValues = seedLoadedPrefixValues(loadedPrefixValues, loadedPrefixStrategy, nextCompanies, loadedPrefixMstDigits, loadedPrefixNameWords, loadedPrefixNameChars, loadedMissingMstPrefixStrategy);
      const displayCompanies = applyPrefixStrategyRows(nextCompanies, loadedPrefixStrategy, loadedPrefixMstDigits, loadedPrefixNameWords, loadedPrefixNameChars, nextPrefixValues, true, loadedMissingMstPrefixStrategy);
      const previewCodes = await loadProductPreviewCodes(displayCompanies, savedWordRules, savedRepeatedPhrases, 'sales', savedProductCodeReplacements);
      updateWorkflow(targetProfile, { ...salesOutputInvalidation(), salesCompanyRows: displayCompanies, selectedSalesCompanyIndex: firstDisplayedCompanyIndex(displayCompanies, loadedGroups), salesMissingMstCompanies: result.missing_mst_companies ?? [], salesProductPreviewCodes: previewCodes, salesProductCodeOverrides: result.manual_code_overrides ?? {}, salesProductCodeReplacements: savedProductCodeReplacements, salesWordRules: savedWordRules, salesRepeatedPhraseRemovals: savedRepeatedPhrases, salesMatchRules: result.sales_match_rules ?? salesMatchRules, salesInventoryPairs: savedInventoryPairs, salesUseDefaultInventoryPair: result.inventory_pairs?.length ? Boolean(result.use_default_inventory_pair) : salesUseDefaultInventoryPair, salesDefaultInventoryPairId: result.inventory_pairs?.length ? (result.default_inventory_pair_id ?? '') : salesDefaultInventoryPairId, salesInventoryPairRules: result.inventory_pair_rules?.length ? result.inventory_pair_rules : salesInventoryPairRules, salesIncludeCompanyPrefix: result.include_company_prefix ?? salesIncludeCompanyPrefix, salesPrefixStrategy: loadedPrefixStrategy, prefixMstDigits: loadedPrefixMstDigits, prefixNameWords: loadedPrefixNameWords, prefixNameChars: loadedPrefixNameChars, prefixMissingMstStrategy: loadedMissingMstPrefixStrategy, salesPrefixStrategyValues: nextPrefixValues, salesProcessingGroups: loadedGroups, salesFormMappingPresets: loadedFormPresets, salesReviewRules: result.vietmax_ban_ra_sales_internal_merges ?? salesReviewRules, salesReviewRows: [], salesReviewGenerated: false });
      setStatus(`Đã tải ${nextCompanies.length} công ty bán ra còn lại sau KVT/152. Chọn công ty/hàng hóa rồi áp dụng trước khi review bán ra.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function updateCompany(index: number, update: Partial<CompanyRow>) {
    updateWorkflow(profile, { companyRows: companyRows.map((company, rowIndex) => (rowIndex === index ? { ...company, ...update } : company)), selectedCompanyIndex: index });
  }

  function updatePendingCompany(index: number, pending: boolean) {
    const groupId = pending ? MATERIALS_GROUP_ID : IGNORED_GROUP_ID;
    const nextRows = companyRows.map((company, rowIndex) => (rowIndex === index ? { ...company, pending_group_id: groupId, pending_process: pending } : company));
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), companyRows: nextRows, selectedCompanyIndex: selectedCompanyIndexAfterGrouping(nextRows, index, purchaseProcessingGroups) });
  }

  function updateCompanyGroup(index: number, groupId: string) {
    const nextRows = companyRows.map((company, rowIndex) => {
      if (rowIndex !== index) return company;
      const process = groupId === MATERIALS_GROUP_ID;
      return { ...company, pending_group_id: groupId, pending_process: process };
    });
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), companyRows: nextRows, selectedCompanyIndex: selectedCompanyIndexAfterGrouping(nextRows, index, purchaseProcessingGroups) });
  }

  function bulkUpdatePendingCompanies(pending: boolean) {
    const groupId = pending ? MATERIALS_GROUP_ID : IGNORED_GROUP_ID;
    const nextRows = companyRows.map((company) => ({ ...company, pending_group_id: groupId, pending_process: pending }));
    updateWorkflow(profile, {
      ...purchaseOutputInvalidation(),
      companyRows: nextRows,
      selectedCompanyIndex: pending ? firstDisplayedCompanyIndex(nextRows, purchaseProcessingGroups) : -1,
    });
  }

  async function applyCompanyAndProductChoices(nextStage?: unknown) {
    const targetStage = isStageId(nextStage) ? nextStage : undefined;
    const activeStrategy = normalizedPrefixStrategy(purchasePrefixStrategy);
    const nextPrefixValues = rememberManualPrefixValues(purchasePrefixStrategyValues, activeStrategy, companyRows, prefixMstDigits, prefixNameWords, prefixNameChars);
    const nextCompanyRows = sortAppliedCompanyRows(companyRows.map((company) => {
      const groupId = pendingCompanyGroup(company);
      const process = groupId === MATERIALS_GROUP_ID;
      return { ...company, value: normalizePrefixValue(company.value), group_id: groupId, pending_group_id: groupId, process, pending_process: process, committed_prefix: normalizePrefixValue(company.value) };
    }), purchaseProcessingGroups);
    const nextWorkflow = {
      ...workflow,
      ...purchaseOutputInvalidation(),
      companyRows: nextCompanyRows,
      purchasePrefixStrategyValues: nextPrefixValues,
      selectedCompanyIndex: firstDisplayedCompanyIndex(nextCompanyRows, purchaseProcessingGroups),
      purchaseReviewRows: [],
      purchaseReviewGenerated: false,
      ...(targetStage ? { stage: targetStage } : {}),
    };
    updateWorkflow(profile, nextWorkflow);
    scrollStageBodyToTop();
    await saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu lựa chọn công ty, hàng hóa vào cấu hình. Review Mã VT sẽ tạo lại theo lựa chọn mới.', 'purchase');
  }

  async function applySalesCompanyAndProductChoices(nextStage?: unknown) {
    const targetStage = isStageId(nextStage) ? nextStage : undefined;
    const activeStrategy = normalizedPrefixStrategy(salesPrefixStrategy);
    const nextPrefixValues = rememberManualPrefixValues(salesPrefixStrategyValues, activeStrategy, salesCompanyRows, prefixMstDigits, prefixNameWords, prefixNameChars);
    const nextSalesCompanyRows = sortAppliedCompanyRows(salesCompanyRows.map((company) => {
      const groupId = pendingCompanyGroup(company);
      const process = groupId === MATERIALS_GROUP_ID;
      return { ...company, value: normalizePrefixValue(company.value), group_id: groupId, pending_group_id: groupId, process, pending_process: process, committed_prefix: normalizePrefixValue(company.value) };
    }), salesProcessingGroups);
    const nextWorkflow = {
      ...workflow,
      ...salesOutputInvalidation(),
      salesCompanyRows: nextSalesCompanyRows,
      salesPrefixStrategyValues: nextPrefixValues,
      selectedSalesCompanyIndex: firstDisplayedCompanyIndex(nextSalesCompanyRows, salesProcessingGroups),
      salesReviewRows: [],
      salesReviewGenerated: false,
      ...(targetStage ? { stage: targetStage } : {}),
    };
    updateWorkflow(profile, nextWorkflow);
    scrollStageBodyToTop();
    await saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu lựa chọn công ty, hàng hóa bán ra. Review bán ra sẽ tạo lại theo lựa chọn mới.', 'sales');
  }

  function selectCompany(index: number) {
    updateWorkflow(profile, { selectedCompanyIndex: index });
  }

  function selectSalesCompany(index: number) {
    updateWorkflow(profile, { selectedSalesCompanyIndex: index });
  }

  function updateCompanyProduct(companyIndex: number, productName: string, selected: boolean) {
    updateWorkflow(profile, {
      selectedCompanyIndex: companyIndex,
      companyRows: companyRows.map((company, rowIndex) => {
        if (rowIndex !== companyIndex) return company;
        const current = new Set(company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name));
        if (selected) {
          current.add(productName);
        } else {
          current.delete(productName);
        }
        return { ...company, selected_product_names: company.all_products.map((product) => product.name).filter((name) => current.has(name)) };
      }),
      purchaseReviewRows: [],
      purchaseReviewGenerated: false,
      ...purchaseOutputInvalidation(),
    });
  }

  function updateSalesPendingCompany(index: number, pending: boolean) {
    const groupId = pending ? MATERIALS_GROUP_ID : IGNORED_GROUP_ID;
    const nextRows = salesCompanyRows.map((company, rowIndex) => (rowIndex === index ? { ...company, pending_group_id: groupId, pending_process: pending } : company));
    updateWorkflow(profile, { ...salesOutputInvalidation(), salesCompanyRows: nextRows, selectedSalesCompanyIndex: selectedCompanyIndexAfterGrouping(nextRows, index, salesProcessingGroups) });
  }

  function updateSalesCompanyGroup(index: number, groupId: string) {
    const nextRows = salesCompanyRows.map((company, rowIndex) => {
      if (rowIndex !== index) return company;
      const process = groupId === MATERIALS_GROUP_ID;
      return { ...company, pending_group_id: groupId, pending_process: process };
    });
    updateWorkflow(profile, { ...salesOutputInvalidation(), salesCompanyRows: nextRows, selectedSalesCompanyIndex: selectedCompanyIndexAfterGrouping(nextRows, index, salesProcessingGroups) });
  }

  function bulkUpdateSalesPendingCompanies(pending: boolean) {
    const groupId = pending ? MATERIALS_GROUP_ID : IGNORED_GROUP_ID;
    const nextRows = salesCompanyRows.map((company) => ({ ...company, pending_group_id: groupId, pending_process: pending }));
    updateWorkflow(profile, {
      ...salesOutputInvalidation(),
      salesCompanyRows: nextRows,
      selectedSalesCompanyIndex: pending ? firstDisplayedCompanyIndex(nextRows, salesProcessingGroups) : -1,
    });
  }

  function updateSalesCompanyProduct(companyIndex: number, productName: string, selected: boolean) {
    updateWorkflow(profile, {
      selectedSalesCompanyIndex: companyIndex,
      salesCompanyRows: salesCompanyRows.map((company, rowIndex) => {
        if (rowIndex !== companyIndex) return company;
        const current = new Set(company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name));
        if (selected) current.add(productName);
        else current.delete(productName);
        return { ...company, selected_product_names: company.all_products.map((product) => product.name).filter((name) => current.has(name)) };
      }),
      salesReviewRows: [],
      salesReviewGenerated: false,
      ...salesOutputInvalidation(),
    });
  }

  function updateProductCode(companyIndex: number, productName: string, code: string) {
    const company = companyRows[companyIndex];
    if (!company) return;
    const key = productKey(companyConfigKey(company), productName);
    updateWorkflow(profile, {
      selectedCompanyIndex: companyIndex,
      productCodeOverrides: { ...productCodeOverrides, [key]: code.toUpperCase() },
      purchaseReviewRows: [],
      purchaseReviewGenerated: false,
      ...purchaseOutputInvalidation(),
    });
  }

  function updateSalesProductCode(companyIndex: number, productName: string, code: string) {
    const company = salesCompanyRows[companyIndex];
    if (!company) return;
    const key = productKey(companyConfigKey(company), productName);
    updateWorkflow(profile, {
      selectedSalesCompanyIndex: companyIndex,
      salesProductCodeOverrides: { ...salesProductCodeOverrides, [key]: code.toUpperCase() },
      salesReviewRows: [],
      salesReviewGenerated: false,
      ...salesOutputInvalidation(),
    });
  }

  function updateCompanyPrefix(index: number, value: string) {
    const company = companyRows[index];
    if (!company) return;
    const activeStrategy = normalizedPrefixStrategy(purchasePrefixStrategy);
    const nextValue = value;
    updateWorkflow(profile, {
      selectedCompanyIndex: index,
      companyRows: companyRows.map((row, rowIndex) => (rowIndex === index ? { ...row, value: nextValue } : row)),
      purchasePrefixStrategyValues: rememberPrefixEdit(purchasePrefixStrategyValues, activeStrategy, company, nextValue, prefixMstDigits, prefixNameWords, prefixNameChars, prefixMissingMstStrategy),
    });
  }

  function updateSalesCompanyPrefix(index: number, value: string) {
    const company = salesCompanyRows[index];
    if (!company) return;
    const activeStrategy = normalizedPrefixStrategy(salesPrefixStrategy);
    const nextValue = value;
    updateWorkflow(profile, {
      selectedSalesCompanyIndex: index,
      salesCompanyRows: salesCompanyRows.map((row, rowIndex) => (rowIndex === index ? { ...row, value: nextValue } : row)),
      salesPrefixStrategyValues: rememberPrefixEdit(salesPrefixStrategyValues, activeStrategy, company, nextValue, prefixMstDigits, prefixNameWords, prefixNameChars, prefixMissingMstStrategy),
    });
  }

  async function updateIncludeCompanyPrefix(include: boolean) {
    const editingSales = stage === 9;
    const nextWorkflow: WorkflowState = {
      ...workflow,
      ...(editingSales ? salesOutputInvalidation() : purchaseOutputInvalidation()),
      ...(editingSales ? { salesIncludeCompanyPrefix: include } : { includeCompanyPrefix: include }),
      purchaseReviewRows: [],
      purchaseReviewGenerated: false,
      salesReviewRows: [],
      salesReviewGenerated: false,
    };
    updateWorkflow(profile, nextWorkflow);
    await saveWorkflowConfig(nextWorkflow, `Đã ${include ? 'bật' : 'tắt'} và lưu cấu hình prefix công ty.`, editingSales ? 'sales' : 'purchase');
  }

  function updatePrefixMstDigits(digits: number) {
    const nextDigits = clampPrefixMstDigits(digits);
    const activeStrategy = normalizedPrefixStrategy(stage === 9 ? salesPrefixStrategy : purchasePrefixStrategy);
    const nextWorkflow: Partial<WorkflowState> = { prefixMstDigits: nextDigits };
    if (stage === 3) {
      const rememberedValues = rememberManualPrefixValues(purchasePrefixStrategyValues, activeStrategy, companyRows, prefixMstDigits, prefixNameWords, prefixNameChars);
      nextWorkflow.purchasePrefixStrategyValues = rememberedValues;
      nextWorkflow.companyRows = applyPrefixStrategyRows(companyRows, activeStrategy, nextDigits, prefixNameWords, prefixNameChars, rememberedValues, false, prefixMissingMstStrategy);
    }
    if (stage === 9) {
      const rememberedValues = rememberManualPrefixValues(salesPrefixStrategyValues, activeStrategy, salesCompanyRows, prefixMstDigits, prefixNameWords, prefixNameChars);
      nextWorkflow.salesPrefixStrategyValues = rememberedValues;
      nextWorkflow.salesCompanyRows = applyPrefixStrategyRows(salesCompanyRows, activeStrategy, nextDigits, prefixNameWords, prefixNameChars, rememberedValues, false, prefixMissingMstStrategy);
    }
    updateWorkflow(profile, nextWorkflow);
  }

  function updatePrefixNameWords(words: number) {
    const nextWords = clampPrefixNameWords(words);
    const activeStrategy = normalizedPrefixStrategy(stage === 9 ? salesPrefixStrategy : purchasePrefixStrategy);
    const nextWorkflow: Partial<WorkflowState> = { prefixNameWords: nextWords };
    if (stage === 3) {
      const rememberedValues = rememberManualPrefixValues(purchasePrefixStrategyValues, activeStrategy, companyRows, prefixMstDigits, prefixNameWords, prefixNameChars);
      nextWorkflow.purchasePrefixStrategyValues = rememberedValues;
      nextWorkflow.companyRows = applyPrefixStrategyRows(companyRows, activeStrategy, prefixMstDigits, nextWords, prefixNameChars, rememberedValues, false, prefixMissingMstStrategy);
    }
    if (stage === 9) {
      const rememberedValues = rememberManualPrefixValues(salesPrefixStrategyValues, activeStrategy, salesCompanyRows, prefixMstDigits, prefixNameWords, prefixNameChars);
      nextWorkflow.salesPrefixStrategyValues = rememberedValues;
      nextWorkflow.salesCompanyRows = applyPrefixStrategyRows(salesCompanyRows, activeStrategy, prefixMstDigits, nextWords, prefixNameChars, rememberedValues, false, prefixMissingMstStrategy);
    }
    updateWorkflow(profile, nextWorkflow);
  }

  function updatePrefixNameChars(chars: number) {
    const nextChars = clampPrefixNameChars(chars);
    const activeStrategy = normalizedPrefixStrategy(stage === 9 ? salesPrefixStrategy : purchasePrefixStrategy);
    const nextWorkflow: Partial<WorkflowState> = { prefixNameChars: nextChars };
    if (stage === 3) {
      const rememberedValues = rememberManualPrefixValues(purchasePrefixStrategyValues, activeStrategy, companyRows, prefixMstDigits, prefixNameWords, prefixNameChars);
      nextWorkflow.purchasePrefixStrategyValues = rememberedValues;
      nextWorkflow.companyRows = applyPrefixStrategyRows(companyRows, activeStrategy, prefixMstDigits, prefixNameWords, nextChars, rememberedValues, false, prefixMissingMstStrategy);
    }
    if (stage === 9) {
      const rememberedValues = rememberManualPrefixValues(salesPrefixStrategyValues, activeStrategy, salesCompanyRows, prefixMstDigits, prefixNameWords, prefixNameChars);
      nextWorkflow.salesPrefixStrategyValues = rememberedValues;
      nextWorkflow.salesCompanyRows = applyPrefixStrategyRows(salesCompanyRows, activeStrategy, prefixMstDigits, prefixNameWords, nextChars, rememberedValues, false, prefixMissingMstStrategy);
    }
    updateWorkflow(profile, nextWorkflow);
  }
  async function updateMissingMstPrefixStrategy(strategy: PrefixPresetStrategy) {
    const nextStrategy = normalizeMissingMstPrefixStrategy(strategy);
    const editingSales = stage === 9;
    const nextWorkflow: WorkflowState = {
      ...workflow,
      ...(editingSales ? salesOutputInvalidation() : purchaseOutputInvalidation()),
      prefixMissingMstStrategy: nextStrategy,
    };
    if (companyRows.length) {
      nextWorkflow.companyRows = applyPrefixStrategyRows(companyRows, normalizedPrefixStrategy(purchasePrefixStrategy), prefixMstDigits, prefixNameWords, prefixNameChars, purchasePrefixStrategyValues, false, nextStrategy);
    }
    if (salesCompanyRows.length) {
      nextWorkflow.salesCompanyRows = applyPrefixStrategyRows(salesCompanyRows, normalizedPrefixStrategy(salesPrefixStrategy), prefixMstDigits, prefixNameWords, prefixNameChars, salesPrefixStrategyValues, false, nextStrategy);
    }
    updateWorkflow(profile, nextWorkflow);
    await saveWorkflowConfig(nextWorkflow, `Đã lưu kiểu prefix cho công ty không MST: ${nextStrategy === 'all_name_words' ? 'Áp tất cả từ đầu' : 'Áp 2 từ'}.`, editingSales ? 'sales' : 'purchase');
  }

  async function applyPurchasePrefixPreset(strategy: PrefixPresetStrategy) {
    const currentStrategy = normalizedPrefixStrategy(purchasePrefixStrategy);
    const rememberedValues = rememberManualPrefixValues(purchasePrefixStrategyValues, currentStrategy, companyRows, prefixMstDigits, prefixNameWords, prefixNameChars);
    const nextWorkflow: WorkflowState = {
      ...workflow,
      ...purchaseOutputInvalidation(),
      purchasePrefixStrategy: strategy,
      purchasePrefixStrategyValues: rememberedValues,
      companyRows: applyPrefixStrategyRows(companyRows, strategy, prefixMstDigits, prefixNameWords, prefixNameChars, rememberedValues, false, prefixMissingMstStrategy),
    };
    updateWorkflow(profile, nextWorkflow);
    await saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu loại prefix mua vào.', 'purchase');
  }

  async function applySalesPrefixPreset(strategy: PrefixPresetStrategy) {
    const currentStrategy = normalizedPrefixStrategy(salesPrefixStrategy);
    const rememberedValues = rememberManualPrefixValues(salesPrefixStrategyValues, currentStrategy, salesCompanyRows, prefixMstDigits, prefixNameWords, prefixNameChars);
    const nextWorkflow: WorkflowState = {
      ...workflow,
      ...salesOutputInvalidation(),
      salesPrefixStrategy: strategy,
      salesPrefixStrategyValues: rememberedValues,
      salesCompanyRows: applyPrefixStrategyRows(salesCompanyRows, strategy, prefixMstDigits, prefixNameWords, prefixNameChars, rememberedValues, false, prefixMissingMstStrategy),
    };
    updateWorkflow(profile, nextWorkflow);
    await saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu loại prefix bán ra.', 'sales');
  }

  async function refreshSalesProductPreviews() {
    if (!salesCompanyRows.length) return;
    setBusy(true);
    setStatus('Đang cập nhật Mã VT preview bán ra...');
    try {
      const previewCodes = isTwoPhaseGenericProfile(profile)
        ? await loadGenericProductPreviewCodes(profile, salesCompanyRows, salesWordRules, firstWordRules, salesRepeatedPhraseRemovals, salesProductCodeReplacements)
        : await loadProductPreviewCodes(salesCompanyRows, salesWordRules, salesRepeatedPhraseRemovals, 'sales', salesProductCodeReplacements);
      const nextWorkflow = { ...workflow, ...salesOutputInvalidation(), salesProductPreviewCodes: previewCodes, salesReviewRows: [], salesReviewGenerated: false };
      updateWorkflow(profile, nextWorkflow);
      await persistWorkflowConfig(nextWorkflow, 'sales', profile);
      setStatus('Đã cập nhật Mã VT preview bán ra và lưu cấu hình bán ra.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function runSalesReview() {
    if (!salesFile) return;
    const dirtyCompanies = salesCompanyRows.filter(hasCompanyDraftChanges);
    if (dirtyCompanies.length) {
      setStatus('Đang có thay đổi lọc công ty/prefix bán ra chưa áp dụng. Bấm Áp dụng lựa chọn công ty và hàng hóa trước khi review.');
      return;
    }
    const reviewProducts = buildSalesReviewProducts(workflow);
    const scope = workflow.salesReviewScope === 'company' ? 'same_company' : 'all_companies';
    const progress = beginProgress('Đang chuẩn bị review Mã VT bán ra');
    setBusy(true);
    setStatus('Đang tạo review Mã VT bán ra theo công ty/hàng hóa đã chọn...');
    try {
      const result = await createPurchaseReview(salesFile.saved_name, scope, salesWordRules, salesRepeatedPhraseRemovals, reviewProducts as ReviewProduct[], progress.operationId, 'sales', true);
      updateWorkflow(profile, { ...salesOutputInvalidation(), salesReviewRows: normalizeReviewRows(result.review_rows as ReviewRow[]), salesReviewGenerated: true, stage: 10 });
      setStatus(`Đã tạo ${result.review_rows.length} dòng Review bán ra. Chỉ dòng được tick mới gộp khi xuất file bán ra.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function saveWorkflowConfig(targetWorkflow: WorkflowState, successMessage = 'Đã lưu cấu hình hiện tại.', phase: 'purchase' | 'sales' | 'all' = 'all', replaceFormMappings = false) {
    setBusy(true);
    setStatus('Đang lưu cấu hình...');
    try {
      await persistWorkflowConfig(targetWorkflow, phase, profile, replaceFormMappings);
      setStatus(successMessage);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function saveCurrentConfig() {
    const latestWorkflow = workflowsRef.current[profile];
    if (latestWorkflow.stage === 3 && latestWorkflow.companyRows.some(hasCompanyDraftChanges)) {
      void applyCompanyAndProductChoices();
      return;
    }
    if (latestWorkflow.stage === 9 && latestWorkflow.salesCompanyRows.some(hasCompanyDraftChanges)) {
      void applySalesCompanyAndProductChoices();
      return;
    }
    void saveWorkflowConfig(latestWorkflow, 'Đã lưu cấu hình hiện tại.', latestWorkflow.stage === 0.5 ? 'all' : (latestWorkflow.stage >= 6 ? 'sales' : 'purchase'));
  }

  async function reloadCurrentConfig() {
    setBusy(true);
    setStatus(`Đang hoàn tất lưu rồi tải lại toàn bộ cấu hình ${selectedProfile.label} từ file config...`);
    try {
      await waitForPendingConfigSave();
      if (profile === 'vietmax') {
        await loadVietmaxProfileConfig();
      } else if (isGenericWorkflowProfile) {
        await loadGenericProfileConfig(profile);
      }
      setStatus(`Đã tải lại cấu hình ${selectedProfile.label}, bao gồm nhóm, form mapping, prefix, review, khớp và phân kho.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function runGenericSalesReview() {
    if (!salesFile || !isTwoPhaseGenericProfile(profile)) return;
    const dirtyCompanies = salesCompanyRows.filter(hasCompanyDraftChanges);
    if (dirtyCompanies.length) {
      setStatus('Đang có thay đổi lọc công ty/prefix bán ra chưa áp dụng. Bấm Áp dụng lựa chọn công ty và hàng hóa trước khi review.');
      return;
    }
    const reviewProducts = buildSalesReviewProducts(workflow);
    const scope = workflow.salesReviewScope === 'company' ? 'same_company' : 'all_companies';
    const progress = beginProgress('Đang chuẩn bị review Mã VT bán ra');
    setBusy(true);
    setStatus(`Đang tạo review Mã VT bán ra ${selectedProfile.label}...`);
    try {
      const result = await createGenericReview(salesFile.saved_name, profile, scope, salesWordRules, firstWordRules, salesRepeatedPhraseRemovals, reviewProducts as ReviewProduct[], progress.operationId, 'sales');
      updateWorkflow(profile, { ...salesOutputInvalidation(), salesReviewRows: normalizeReviewRows(result.review_rows as ReviewRow[]), salesReviewGenerated: true, stage: 10 });
      setStatus(`Đã tạo ${result.review_rows.length} dòng Review Mã VT bán ra ${selectedProfile.label}.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function loadGenericSalesCompanies(nextStage?: StageId) {
    if (!salesFile || !isTwoPhaseGenericProfile(profile)) return;
    const targetProfile = profile;
    setBusy(true);
    setStatus('Đang tải danh sách công ty và hàng hóa bán ra...');
    try {
      const sourceColumns = normalizeVietmaxColumns(salesColumns, 'sales');
      const [result, formatDefaults] = await Promise.all([
        analyzeGenericWorkbook({
          saved_name: salesFile.saved_name,
          original_name: salesFile.original_name,
          profile: targetProfile,
          vietmax_phase: 'sales',
          ...sourceColumns,
        }),
        getVietmaxFormatMappingDefaults(targetProfile).catch(() => null),
      ]);
      const savedWordRules = result.word_rules ?? salesWordRules;
      const savedProductCodeReplacements = result.product_code_replacements ?? salesProductCodeReplacements;
      const savedRepeatedPhrases = result.repeated_phrase_removals ?? salesRepeatedPhraseRemovals;
      const savedInventoryPairs = result.inventory_pairs ?? salesInventoryPairs;
      const loadedGroups = normalizeProcessingGroups(result.processing_groups, 'sales');
      const loadedFormPresets = normalizeFormMappingPresets(result.form_mapping_presets, 'sales', formatDefaults);
      const nextCompanies = applyGroupAssignments(result.companies.map((company) => ({
        ...company,
        process: company.process ?? true,
        pending_process: company.pending_process ?? company.process ?? true,
        committed_prefix: company.committed_prefix ?? company.value ?? '',
        selected_product_names: company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name),
      })), result.company_group_assignments ?? {});
      const loadedPrefixStrategy = normalizedPrefixStrategy(result.prefix_strategy || salesPrefixStrategy);
      const loadedPrefixMstDigits = clampPrefixMstDigits(result.prefix_mst_digits ?? prefixMstDigits);
      const loadedPrefixNameWords = clampPrefixNameWords(result.prefix_name_words ?? prefixNameWords);
      const loadedMissingMstPrefixStrategy = normalizeMissingMstPrefixStrategy(result.prefix_missing_mst_strategy ?? prefixMissingMstStrategy);
      const loadedPrefixValues = normalizePrefixStrategyValues(result.prefix_strategy_values, salesPrefixStrategyValues);
      const loadedPrefixNameChars = clampPrefixNameChars(result.prefix_name_chars ?? prefixNameChars);
      const nextPrefixValues = seedLoadedPrefixValues(loadedPrefixValues, loadedPrefixStrategy, nextCompanies, loadedPrefixMstDigits, loadedPrefixNameWords, loadedPrefixNameChars, loadedMissingMstPrefixStrategy);
      const displayCompanies = applyPrefixStrategyRows(nextCompanies, loadedPrefixStrategy, loadedPrefixMstDigits, loadedPrefixNameWords, loadedPrefixNameChars, nextPrefixValues, true, loadedMissingMstPrefixStrategy);
      const previewCodes = await loadGenericProductPreviewCodes(targetProfile, displayCompanies, savedWordRules, firstWordRules, savedRepeatedPhrases, savedProductCodeReplacements);
      updateWorkflow(targetProfile, {
        ...salesOutputInvalidation(),
        salesColumns: sourceColumns,
        salesCompanyRows: displayCompanies,
        selectedSalesCompanyIndex: firstDisplayedCompanyIndex(displayCompanies, loadedGroups),
        salesMissingMstCompanies: result.missing_mst_companies ?? [],
        salesProductPreviewCodes: previewCodes,
        salesProductCodeOverrides: result.manual_code_overrides ?? {},
        salesProductCodeReplacements: savedProductCodeReplacements,
        salesWordRules: savedWordRules,
        salesRepeatedPhraseRemovals: savedRepeatedPhrases,
        salesInventoryPairs: savedInventoryPairs,
        salesUseDefaultInventoryPair: result.use_default_inventory_pair ?? salesUseDefaultInventoryPair,
        salesDefaultInventoryPairId: result.default_inventory_pair_id ?? salesDefaultInventoryPairId,
        salesInventoryPairRules: result.inventory_pair_rules ?? salesInventoryPairRules,
        salesIncludeCompanyPrefix: result.include_company_prefix ?? salesIncludeCompanyPrefix,
        salesPrefixStrategy: loadedPrefixStrategy,
        prefixMstDigits: loadedPrefixMstDigits,
        prefixNameWords: loadedPrefixNameWords,
        prefixNameChars: loadedPrefixNameChars,
        prefixMissingMstStrategy: loadedMissingMstPrefixStrategy,
        salesPrefixStrategyValues: nextPrefixValues,
        salesProcessingGroups: loadedGroups,
        salesFormMappingPresets: loadedFormPresets,
        salesReviewRules: Array.isArray(result.product_review_merges) ? result.product_review_merges as ReviewRow[] : salesReviewRules,
        salesReviewRows: [],
        salesReviewGenerated: false,
        ...(nextStage ? { stage: nextStage } : {}),
      });
      setStatus(`Đã tải ${result.company_count} công ty, ${result.rows_to_process} dòng bán ra ${selectedProfile.label}.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function saveFormatMappingConfig() {
    const nextPurchase = normalizeFormsForSave(purchaseFormMappingPresets);
    const nextSales = normalizeFormsForSave(salesFormMappingPresets);
    const nextWorkflow = {
      ...workflow,
      purchaseFormMappingPresets: nextPurchase,
      salesFormMappingPresets: nextSales,
    };
    updateWorkflow(profile, nextWorkflow);
    void saveWorkflowConfig(nextWorkflow, 'Đã lưu cấu hình form mapping.', 'all', true);
  }

  async function restoreDefaultFormMappings() {
    setBusy(true);
    const defaultCount = profile === 'ho_guom' ? 3 : 5;
    setStatus(`Đang khôi phục ${defaultCount} form mapping mặc định...`);
    try {
      const defaults = await getVietmaxFormatMappingDefaults(profile);
      const nextWorkflow = {
        ...workflow,
        purchaseFormSourceColumns: formatSourceColumns(defaults.source_columns?.purchase),
        salesFormSourceColumns: formatSourceColumns(defaults.source_columns?.sales),
        purchaseFormMappingPresets: normalizeFormMappingPresets(defaults.form_mapping_presets, 'purchase', defaults),
        salesFormMappingPresets: normalizeFormMappingPresets(defaults.form_mapping_presets, 'sales', defaults),
      };
      updateWorkflow(profile, nextWorkflow);
      await saveWorkflowConfig(nextWorkflow, `Đã khôi phục ${defaultCount} form mapping mặc định cho profile hiện tại.`, 'all', true);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function updateFormatLists(scope: FormatScope, update: (forms: FormMappingPreset[], phase: 'purchase' | 'sales') => FormMappingPreset[]) {
    const nextPurchase = scope === 'sales' ? purchaseFormMappingPresets : update(purchaseFormMappingPresets, 'purchase');
    const nextSales = scope === 'purchase' ? salesFormMappingPresets : update(salesFormMappingPresets, 'sales');
    updateWorkflow(profile, {
      ...(isGenericWorkflowProfile ? purchaseOutputInvalidation() : {}),
      purchaseFormMappingPresets: nextPurchase,
      salesFormMappingPresets: nextSales,
    });
  }

  function updateFormatGroups(scope: FormatScope, update: (groups: ProcessingGroup[], phase: 'purchase' | 'sales') => ProcessingGroup[]) {
    const nextPurchase = scope === 'sales' ? purchaseProcessingGroups : update(purchaseProcessingGroups, 'purchase');
    const nextSales = scope === 'purchase' ? salesProcessingGroups : update(salesProcessingGroups, 'sales');
    updateWorkflow(profile, {
      ...(isGenericWorkflowProfile ? purchaseOutputInvalidation() : {}),
      purchaseProcessingGroups: nextPurchase,
      salesProcessingGroups: nextSales,
    });
  }

  function addFormatGroup(scope: FormatScope, labelOverride?: string) {
    updateFormatGroups(scope, (groups) => {
      const label = String(labelOverride || '').trim() || 'Nhóm mới';
      return [...groups, { id: groupIdFromLabel(label, groups), label, uses_product_code: false, forms: [] }];
    });
  }

  function updateFormatGroup(scope: FormatScope, groupId: string, update: Partial<ProcessingGroup>) {
    updateFormatGroups(scope, (groups) => groups.map((group) => (group.id === groupId ? { ...group, ...update, id: group.id } : group)));
  }

  function deleteFormatGroup(scope: FormatScope, groupId: string) {
    const group = (scope === 'sales' ? salesProcessingGroups : purchaseProcessingGroups).find((item) => item.id === groupId);
    if (!group || group.builtin) return;
    const resetForms = (forms: FormMappingPreset[]) => forms.map((form) => form.group_id === groupId ? { ...form, group_id: MATERIALS_GROUP_ID } : form);
    updateFormatGroups(scope, (groups) => groups.filter((item) => item.id !== groupId));
    updateWorkflow(profile, {
      ...(isGenericWorkflowProfile ? purchaseOutputInvalidation() : {}),
      purchaseFormMappingPresets: scope === 'sales' ? purchaseFormMappingPresets : resetForms(purchaseFormMappingPresets),
      salesFormMappingPresets: scope === 'purchase' ? salesFormMappingPresets : resetForms(salesFormMappingPresets),
    });
  }

  function addFormatForm(scope: FormatScope, groupId = MATERIALS_GROUP_ID) {
    const sourcePhase = scope === 'sales' ? 'sales' : 'purchase';
    const id = `custom_${scope}_${Date.now().toString(36)}`;
    const form: FormMappingPreset = {
      id,
      label: scope === 'purchase' ? 'Form mua vào mới' : scope === 'sales' ? 'Form bán ra mới' : 'Form dùng chung mới',
      scope,
      type: 'template_mapping',
      enabled: true,
      group_id: groupId,
      input_phase: scope === 'both' ? 'both' : scope,
      sheet: '',
      output_columns: outputColumns(['Cột A', 'Cột B', 'Cột C', 'Cột D']),
      mappings: [mappingRule('A', '', sourcePhase)],
    };
    updateFormatLists(scope, (forms) => [...forms, form]);
  }

  function updateFormatForm(scope: FormatScope, formId: string, update: Partial<FormMappingPreset>) {
    updateFormatLists(scope, (forms) => replaceFormatFormInList(forms, scope, formId, (form) => ({ ...form, ...update, scope: formatScopeOfForm(form, scope) })));
  }

  function updateFormatFormScope(scope: FormatScope, formId: string, nextScope: FormatScope) {
    const form = visibleFormatFormsForScope(scope, purchaseFormMappingPresets, salesFormMappingPresets).find((item) => item.id === formId);
    if (!form) return;
    const nextForm = formWithScope(form, nextScope);
    let nextPurchase = removeFormatFormFromList(purchaseFormMappingPresets, formId);
    let nextSales = removeFormatFormFromList(salesFormMappingPresets, formId);
    if (nextScope === 'purchase' || nextScope === 'both') nextPurchase = upsertFormatForm(nextPurchase, nextForm);
    if (nextScope === 'sales' || nextScope === 'both') nextSales = upsertFormatForm(nextSales, nextForm);
    updateWorkflow(profile, {
      ...(isGenericWorkflowProfile ? purchaseOutputInvalidation() : {}),
      purchaseFormMappingPresets: nextPurchase,
      salesFormMappingPresets: nextSales,
    });
  }

  function deleteFormatForm(scope: FormatScope, formId: string) {
    updateFormatForm(scope, formId, { enabled: false });
  }

  function addFormatMappingRule(scope: FormatScope, formId: string) {
    const sourcePhase = scope === 'sales' ? 'sales' : 'purchase';
    updateFormatForm(scope, formId, {
      mappings: [
        ...(visibleFormatFormsForScope(scope, purchaseFormMappingPresets, salesFormMappingPresets).find((form) => form.id === formId)?.mappings || []),
        mappingRule('A', '', sourcePhase),
      ],
    });
  }

  function updateFormatMappingRule(scope: FormatScope, formId: string, index: number, update: Partial<NonNullable<FormMappingPreset['mappings']>[number]>) {
    const form = visibleFormatFormsForScope(scope, purchaseFormMappingPresets, salesFormMappingPresets).find((item) => item.id === formId);
    if (!form) return;
    const mappings = [...(form.mappings || [])];
    mappings[index] = { ...mappings[index], ...update };
    updateFormatForm(scope, formId, { mappings });
  }

  function removeFormatMappingRule(scope: FormatScope, formId: string, index: number) {
    const form = visibleFormatFormsForScope(scope, purchaseFormMappingPresets, salesFormMappingPresets).find((item) => item.id === formId);
    if (!form) return;
    updateFormatForm(scope, formId, { mappings: (form.mappings || []).filter((_, itemIndex) => itemIndex !== index) });
  }

  async function uploadFormatTemplate(scope: FormatScope, formId: string, file: File | undefined) {
    if (!file) return;
    setBusy(true);
    setStatus('Đang tải file mẫu output...');
    try {
      const summary = await uploadFormTemplate(file);
      updateFormatForm(scope, formId, {
        type: 'template_mapping',
        template_original_name: summary.original_name,
        template_saved_name: summary.saved_name,
        output_columns: columnsFromUploadSummary(summary),
        output_preview: summary.preview,
      });
      setStatus(`Đã tải file mẫu ${summary.original_name}. Hãy chọn cột output trong mapping.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function exportCurrentVietmaxConfig(phase: 'purchase' | 'sales') {
    if (profile !== 'vietmax') {
      setStatus('Chỉ xuất cấu hình Vietmax ở profile Vietmax.');
      return;
    }
    const label = phase === 'purchase' ? 'mua vào' : 'bán ra';
    setBusy(true);
    setStatus(`Đang xuất cấu hình Vietmax ${label}...`);
    try {
      const cfg = await getAppConfig();
      const snapshot = buildVietmaxConfigExportSnapshot(workflow, phase, cfg);
      const blob = new Blob([JSON.stringify(snapshot, null, 2)], { type: 'application/json;charset=utf-8' });
      const saved = await saveBlob(blob, `vietmax_${phase === 'purchase' ? 'mua_vao' : 'ban_ra'}_config_${exportTimestamp()}.json`);
      setStatus(saved ? `Đã xuất cấu hình Vietmax ${label}. File gồm cấu hình đã lưu và snapshot UI hiện tại.` : `Đã hủy lưu cấu hình Vietmax ${label}.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function importCurrentVietmaxConfig(phase: 'purchase' | 'sales') {
    if (profile !== 'vietmax') {
      setStatus('Chỉ nhập cấu hình Vietmax ở profile Vietmax.');
      return;
    }
    const allowedStage = phase === 'purchase' ? stage === 0.5 || stage === 1 : stage === 6;
    if (!allowedStage) {
      setStatus(phase === 'purchase'
        ? 'Chỉ nhập cấu hình mua vào ở stage 1 để tránh conflict dữ liệu đang xử lý.'
        : 'Chỉ nhập cấu hình bán ra ở stage 6 để tránh conflict dữ liệu đang xử lý.');
      return;
    }
    const label = phase === 'purchase' ? 'mua vào' : 'bán ra';
    const file = await chooseJsonConfigFile();
    if (!file) {
      setStatus(`Đã hủy nhập cấu hình Vietmax ${label}.`);
      return;
    }
    let shouldReload = false;
    setBusy(true);
    setStatus(`Đang nhập cấu hình Vietmax ${label}...`);
    try {
      const text = await file.text();
      const payload = JSON.parse(text) as Record<string, unknown>;
      await importVietmaxConfig(phase, payload);
      shouldReload = true;
      setStatus(`Đã nhập cấu hình Vietmax ${label}.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
    if (!shouldReload) return;
    if (phase === 'purchase' && purchaseFile) {
      await loadCompanies();
    } else if (phase === 'sales' && salesFile) {
      await loadSalesCompanies();
    }
  }

  async function exportSelectedVietmaxConfig() {
    if (configTransferScope === 'purchase' || configTransferScope === 'sales') {
      await exportCurrentVietmaxConfig(configTransferScope);
      return;
    }
    if (profile !== 'vietmax') {
      setStatus('Chỉ xuất cấu hình theo scope ở profile Vietmax.');
      return;
    }
    setBusy(true);
    setStatus('Đang xuất cấu hình Vietmax mua vào và bán ra...');
    try {
      const cfg = await getAppConfig();
      const payload = {
        export_version: 2,
        app: 'Product Code Formatter',
        profile: 'vietmax',
        scope: 'all',
        phases: {
          purchase: buildVietmaxConfigExportSnapshot(workflow, 'purchase', cfg),
          sales: buildVietmaxConfigExportSnapshot(workflow, 'sales', cfg),
        },
      };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
      const saved = await saveBlob(blob, 'vietmax_all_config_' + exportTimestamp() + '.json');
      setStatus(saved ? 'Đã xuất một file cấu hình gồm cả mua vào và bán ra.' : 'Đã hủy xuất cấu hình Vietmax.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function importSelectedVietmaxConfig() {
    if (configTransferScope === 'purchase' || configTransferScope === 'sales') {
      await importCurrentVietmaxConfig(configTransferScope);
      return;
    }
    if (profile !== 'vietmax') {
      setStatus('Chỉ nhập cấu hình theo scope ở profile Vietmax.');
      return;
    }
    if (stage !== 0.5 && stage !== 1) {
      setStatus('Hãy nhập cấu hình Tất cả ở stage Form mapping hoặc stage 1.');
      return;
    }
    const file = await chooseJsonConfigFile();
    if (!file) {
      setStatus('Đã hủy nhập cấu hình Vietmax.');
      return;
    }
    setBusy(true);
    setStatus('Đang nhập cấu hình Vietmax mua vào và bán ra...');
    try {
      const payload = JSON.parse(await file.text()) as Record<string, unknown>;
      const phases = payload.phases;
      if (!phases || typeof phases !== 'object') {
        throw new Error('File cấu hình Tất cả không có phần phases mua vào/bán ra.');
      }
      const phasePayloads = phases as Record<string, unknown>;
      if (!phasePayloads.purchase || !phasePayloads.sales) {
        throw new Error('File cấu hình Tất cả phải có đủ purchase và sales.');
      }
      await importVietmaxConfig('purchase', phasePayloads.purchase as Record<string, unknown>);
      await importVietmaxConfig('sales', phasePayloads.sales as Record<string, unknown>);
      if (purchaseFile) await loadCompanies();
      if (salesFile) await loadSalesCompanies();
      setStatus('Đã nhập cấu hình Vietmax cho cả mua vào và bán ra.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }
  async function refreshProductPreviews() {
    if (!companyRows.length) return;
    setBusy(true);
    setStatus('Đang cập nhật Mã VT preview theo từ thay riêng và từ lặp...');
    try {
      const previewCodes = isTwoPhaseGenericProfile(profile)
        ? await loadGenericProductPreviewCodes(profile, companyRows, purchaseWordRules, firstWordRules, purchaseRepeatedPhraseRemovals, productCodeReplacements)
        : isGenericWorkflowProfile
        ? await loadGenericProductPreviewCodes(profile, companyRows, wordRules, firstWordRules, repeatedPhraseRemovals, productCodeReplacements)
        : await loadProductPreviewCodes(companyRows, purchaseWordRules, purchaseRepeatedPhraseRemovals, 'purchase', productCodeReplacements);
      const nextWorkflow = { ...workflow, ...purchaseOutputInvalidation(), productPreviewCodes: previewCodes, purchaseReviewRows: [], purchaseReviewGenerated: false, priceGroups: [] };
      updateWorkflow(profile, nextWorkflow);
      await persistWorkflowConfig(nextWorkflow, 'purchase', profile);
      setStatus('Đã cập nhật Mã VT preview và lưu cấu hình.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function updateWordRule(index: number, field: 'from' | 'to', value: string) {
    const entries = Object.entries(activeWordRules);
    entries[index] = field === 'from' ? [value, entries[index]?.[1] || ''] : [entries[index]?.[0] || '', value];
    const nextRules = Object.fromEntries(entries.filter(([from]) => from.trim()));
    const invalidation = inventoryOutputInvalidation();
    if (activeUsesScopedPhase && (activeVietmaxSalesConfig || activeTwoPhaseSalesConfig)) {
      updateWorkflow(profile, { ...invalidation, salesWordRules: nextRules, salesReviewRows: [], salesReviewGenerated: false });
      return;
    }
    if (activeUsesScopedPhase) {
      updateWorkflow(profile, { ...invalidation, purchaseWordRules: nextRules, purchaseReviewRows: [], purchaseReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, wordRules: nextRules, purchaseReviewRows: [], salesReviewRows: [], purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function addWordRule() {
    const nextRules = { ...activeWordRules, '': '' };
    const invalidation = inventoryOutputInvalidation();
    if (activeUsesScopedPhase && (activeVietmaxSalesConfig || activeTwoPhaseSalesConfig)) {
      updateWorkflow(profile, { ...invalidation, salesWordRules: nextRules, salesReviewGenerated: false });
      return;
    }
    if (activeUsesScopedPhase) {
      updateWorkflow(profile, { ...invalidation, purchaseWordRules: nextRules, purchaseReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, wordRules: nextRules, purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function updateRepeatedPhrase(index: number, value: string) {
    const next = activeRepeatedPhraseRemovals.slice();
    next[index] = value;
    const nextPhrases = next.filter((item, rowIndex) => item.trim() || rowIndex === index);
    const invalidation = inventoryOutputInvalidation();
    if (activeUsesScopedPhase && (activeVietmaxSalesConfig || activeTwoPhaseSalesConfig)) {
      updateWorkflow(profile, { ...invalidation, salesRepeatedPhraseRemovals: nextPhrases, salesReviewRows: [], salesReviewGenerated: false });
      return;
    }
    if (activeUsesScopedPhase) {
      updateWorkflow(profile, { ...invalidation, purchaseRepeatedPhraseRemovals: nextPhrases, purchaseReviewRows: [], purchaseReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, repeatedPhraseRemovals: nextPhrases, purchaseReviewRows: [], salesReviewRows: [], purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function addRepeatedPhrase() {
    const nextPhrases = [...activeRepeatedPhraseRemovals, ''];
    const invalidation = inventoryOutputInvalidation();
    if (activeUsesScopedPhase && (activeVietmaxSalesConfig || activeTwoPhaseSalesConfig)) {
      updateWorkflow(profile, { ...invalidation, salesRepeatedPhraseRemovals: nextPhrases, salesReviewGenerated: false });
      return;
    }
    if (activeUsesScopedPhase) {
      updateWorkflow(profile, { ...invalidation, purchaseRepeatedPhraseRemovals: nextPhrases, purchaseReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, repeatedPhraseRemovals: nextPhrases, purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function removeRepeatedPhrase(index: number) {
    const next = activeRepeatedPhraseRemovals.slice();
    next.splice(index, 1);
    const invalidation = inventoryOutputInvalidation();
    if (activeUsesScopedPhase && (activeVietmaxSalesConfig || activeTwoPhaseSalesConfig)) {
      updateWorkflow(profile, { ...invalidation, salesRepeatedPhraseRemovals: next, salesReviewRows: [], salesReviewGenerated: false });
      return;
    }
    if (activeUsesScopedPhase) {
      updateWorkflow(profile, { ...invalidation, purchaseRepeatedPhraseRemovals: next, purchaseReviewRows: [], purchaseReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, repeatedPhraseRemovals: next, purchaseReviewRows: [], salesReviewRows: [], purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  async function downloadProcessedPurchase() {
    if (!purchaseFile) return;
    setBusy(true);
    setStatus(processedPurchaseSavedName ? 'Đang mở file mua vào đã xử lý từ cache...' : 'Đang tạo file mua vào đã xử lý Mã VT...');
    try {
      let savedName = processedPurchaseSavedName;
      let stats = processedPurchaseStats;
      if (!savedName) {
        const result = await runExplicitProcessJob(purchaseFile, buildPurchaseProcessPayload(workflow), 'vietmax');
        savedName = result.savedName;
        stats = await applyProcessedPurchaseCache(profile, savedName);
      }
      const blob = await downloadCachedFile(savedName);
      const saved = await saveBlob(blob, purchaseFile.original_name.replace(/\.(xls|xlsx|xlsm)$/i, '_fdi.xls'));
      const suffix = processedStatsSentence(stats);
      setStatus(saved ? `Đã xuất file mua vào đã xử lý. ${suffix}` : `Đã hủy lưu file mua vào; cache vẫn được giữ. ${suffix}`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function downloadProcessedSales() {
    if (!salesFile) return;
    setBusy(true);
    setStatus(processedSalesSavedName ? 'Đang mở file bán ra đã xử lý từ cache...' : 'Đang tạo file bán ra đã xử lý Mã VT...');
    try {
      let savedName = processedSalesSavedName;
      let stats = processedSalesStats;
      if (!savedName) {
        const result = await runExplicitProcessJob(salesFile, buildSalesProcessPayload(workflow), 'vietmax');
        savedName = result.savedName;
        stats = await inspectProcessedVietmaxFile(savedName, 'sales');
        updateWorkflow(profile, { processedSalesSavedName: savedName, processedSalesStats: stats, inventoryAllocationJob: null, inventoryAllocationResult: null });
      }
      const blob = await downloadCachedFile(savedName);
      const saved = await saveBlob(blob, salesFile.original_name.replace(/\.(xls|xlsx|xlsm)$/i, '_fdi.xls'));
      const suffix = processedStatsSentence(stats);
      setStatus(saved ? `Đã xuất file bán ra đã xử lý. ${suffix}` : `Đã hủy lưu file bán ra; cache vẫn được giữ. ${suffix}`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function downloadSonPhuongProcessedSales() {
    if (!salesFile) {
      setStatus('Ch\u01b0a c\u00f3 file b\u00e1n ra. H\u00e3y t\u1ea3i file b\u00e1n ra trong section Ph\u00e2n kho tr\u01b0\u1edbc.');
      return;
    }
    const jobId = inventoryAllocationResult?.job_id || inventoryAllocationJob?.result?.job_id;
    if (!processedSalesSavedName && !jobId) {
      setStatus('Ch\u01b0a c\u00f3 k\u1ebft qu\u1ea3 Ph\u00e2n kho. H\u00e3y ch\u1ea1y Ph\u00e2n kho tr\u01b0\u1edbc khi t\u1ea1o FDI b\u00e1n ra.');
      return;
    }
    const progress = beginProgress(processedSalesSavedName ? 'Đang tải FDI bán ra từ cache' : 'Đang tạo FDI bán ra đã Phân kho');
    setBusy(true);
    setStatus(processedSalesSavedName ? '\u0110ang t\u1ea3i FDI b\u00e1n ra \u0111\u00e3 Ph\u00e2n kho t\u1eeb cache...' : '\u0110ang t\u1ea1o FDI b\u00e1n ra \u0111\u00e3 Ph\u00e2n kho...');
    try {
      let savedName = processedSalesSavedName;
      if (!savedName) {
        const created = await createSonPhuongProcessedSales(jobId || '', progress.operationId);
        savedName = created.processed_sales_saved_name;
        if (!savedName) throw new Error('Backend kh\u00f4ng tr\u1ea3 v\u1ec1 cache FDI b\u00e1n ra \u0111\u00e3 Ph\u00e2n kho.');
        const stats = await inspectProcessedVietmaxFile(savedName, 'sales').catch(() => processedSalesStats);
        updateWorkflow(profile, {
          processedSalesSavedName: savedName,
          processedSalesStats: stats,
          inventoryAllocationResult: created.result || inventoryAllocationResult,
        });
      }
      const blob = await downloadCachedFile(savedName);
      const filename = salesFile.original_name.replace(/\.(xls|xlsx|xlsm)$/i, '_fdi_phan_kho.xls');
      const saved = await saveBlob(blob, filename);
      setStatus(saved ? '\u0110\u00e3 l\u01b0u FDI b\u00e1n ra \u0111\u00e3 Ph\u00e2n kho.' : '\u0110\u00e3 h\u1ee7y l\u01b0u file; cache Ph\u00e2n kho v\u1eabn \u0111\u01b0\u1ee3c gi\u1eef.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }
  function changeProfile(nextProfile: ProfileKey) {
    setProfile(nextProfile);
    setStatus(`Đang xem profile ${profiles.find((item) => item.key === nextProfile)?.label ?? nextProfile}. Dữ liệu profile khác vẫn được giữ.`);
  }

  function changeHoGuomMode(nextMode: HoGuomMode) {
    if (nextMode === 'formatter') setProfileConfigLoading(true);
    setHoGuomMode(nextMode);
    updateWorkflow('ho_guom', { stage: nextMode === 'formatter' ? 0.5 : 1 });
    setStatus(nextMode === 'formatter'
      ? 'Hồ Gươm đang dùng chế độ Formatter mã VT với cùng frame mặc định.'
      : 'Hồ Gươm đang dùng chế độ Bóc tách dự toán.');
  }

  function updatePurchaseReviewScope(scope: 'all' | 'company') {
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), purchaseReviewScope: scope, purchaseReviewRows: [], purchaseReviewGenerated: false });
  }

  function updateSalesReviewScope(scope: 'all' | 'company') {
    updateWorkflow(profile, { ...salesOutputInvalidation(), salesReviewScope: scope, salesReviewRows: [], salesReviewGenerated: false });
  }

  function updateComparisonScope(value: string) {
    if (stage >= 6) {
      updateWorkflow(profile, {
        ...salesOutputInvalidation(),
        comparisonScope: value,
        matches: [],
        salesMatchGenerated: false,
        salesCompanyRows: [],
        selectedSalesCompanyIndex: -1,
        salesMissingMstCompanies: [],
        salesProductPreviewCodes: {},
        salesProductCodeOverrides: {},
        salesReviewRows: [],
        salesReviewGenerated: false,
      });
      return;
    }
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), comparisonScope: value, purchaseReviewRows: [], purchaseReviewGenerated: false, salesCompanyRows: [], selectedSalesCompanyIndex: -1, salesMissingMstCompanies: [], salesProductPreviewCodes: {}, salesProductCodeOverrides: {}, salesReviewRows: [], salesReviewGenerated: false });
  }

  function updateInventoryAllocationConfig(config: InventoryAllocationConfig) {
    updateWorkflow(profile, {
      inventoryAllocationConfig: config,
      inventoryAllocationJob: null,
      inventoryAllocationResult: null,
      ...(profile === 'son_phuong' ? { processedSalesSavedName: '' } : {}),
    });
  }

  function updateOpeningStockFile(file: File | null) {
    updateWorkflow(profile, {
      openingStockFile: file,
      inventoryAllocationJob: null,
      inventoryAllocationResult: null,
      ...(profile === 'son_phuong' ? { processedSalesSavedName: '' } : {}),
    });
  }

  async function runInventoryAllocation(nextStage?: StageId) {
    const restoredCaches = await restoreProcessedCachesFromSession();
    const purchaseCache = restoredCaches.purchaseSavedName || processedPurchaseSavedName;
    const salesCache = profile === 'son_phuong'
      ? (salesFile?.saved_name || '')
      : (restoredCaches.salesSavedName || processedSalesSavedName);
    if (!purchaseCache || !salesCache) {
      setStatus(cacheDebugMessage('Cần tải hoặc tạo đủ file mua vào và bán ra trước khi phân kho.', true, true, { purchase: purchaseCache, sales: salesCache }));
      return;
    }
    const targetProfile = profile;
    setBusy(true);
    setStatus('Đang gửi dữ liệu phân bổ tồn kho...');
    updateWorkflow(targetProfile, {
      inventoryAllocationJob: { status: 'queued', progress: 0, done: 0, total: 0, label: 'Đang chuẩn bị phân bổ tồn kho...' },
      inventoryAllocationResult: null,
    });
    try {
      const allocationConfig = {
        ...inventoryAllocationConfig,
        policy: {
          ...inventoryAllocationConfig.policy,
          company_profile: inventoryAllocationProfileFor(targetProfile),
          sales_inventory_pairs: inventoryAllocationConfig.sales_inventory_pairs || [],
          scenario_count: inventoryAllocationConfig.scenario_count || 100,
        },
      };
      const started = await startInventoryAllocation({ purchaseSavedName: purchaseCache, salesSavedName: salesCache, salesOriginalName: salesFile?.original_name || 'ban_ra_da_xu_ly.xlsx', openingFile: openingStockFile, config: allocationConfig });
      let nextJob: InventoryAllocationJob = { status: 'queued', progress: 0, done: 0, total: 0, label: 'Đã gửi dữ liệu. Đang chờ backend xử lý...' };
      updateWorkflow(targetProfile, { inventoryAllocationJob: nextJob, inventoryAllocationResult: null });
      while (nextJob.status === 'queued' || nextJob.status === 'running') {
        await sleep(1000);
        nextJob = await getInventoryAllocationJob(started.analysis_job_id);
        updateWorkflow(targetProfile, { inventoryAllocationJob: nextJob, inventoryAllocationResult: nextJob.result ?? null });
        setStatus(formatInventoryJobStatus(nextJob));
      }
      if (nextJob.status === 'error') throw new Error(nextJob.error || nextJob.label || 'Ph\u00e2n b\u1ed5 kho th\u1ea5t b\u1ea1i.');
      if (nextJob.status !== 'complete' || !nextJob.result) throw new Error(nextJob.label || 'Ph\u00e2n b\u1ed5 kho kh\u00f4ng tr\u1ea3 v\u1ec1 k\u1ebft qu\u1ea3 h\u1ee3p l\u1ec7.');
      if (nextStage === 13 && !nextJob.result.report_view) {
        setStatus('Phân kho đã xong. Đang dựng dữ liệu Báo cáo...');
        const reportJob = await getInventoryAllocationJob(nextJob.result.job_id, true);
        if (reportJob.status !== 'complete' || !reportJob.result?.report_view) {
          throw new Error(reportJob.error || reportJob.label || 'Không dựng được dữ liệu Báo cáo Phân kho.');
        }
        nextJob = reportJob;
      }
      updateWorkflow(targetProfile, {
        inventoryAllocationJob: nextJob,
        inventoryAllocationResult: nextJob.result,
        ...(targetProfile === 'son_phuong' ? { processedSalesSavedName: '' } : {}),
      });
      if (nextStage) updateWorkflow(targetProfile, { stage: nextStage });
      setStatus('Đã hoàn tất phân bổ tồn kho. Kiểm tra báo cáo trước khi xuất file.');
    } catch (error) {
      updateWorkflow(targetProfile, { inventoryAllocationJob: null });
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`${message} ${cacheDebugMessage('Debug phân bổ tồn kho.', true, true, { purchase: purchaseCache, sales: salesCache })}`);
    } finally {
      setBusy(false);
    }
  }

  async function downloadInventoryReport() {
    const jobId = inventoryAllocationResult?.job_id || inventoryAllocationJob?.result?.job_id;
    if (!jobId) return;
    const progress = beginProgress('Đang tạo báo cáo Phân kho');
    setBusy(true);
    setStatus('Đang tạo và chuyển báo cáo Phân kho sang định dạng XLS...');
    try {
      const blob = await downloadInventoryAllocationReport(jobId, progress.operationId);
      const filename = toXlsName(inventoryAllocationResult?.filename || inventoryAllocationJob?.result?.filename || 'phan_bo_ton_kho.xls');
      const saved = await saveBlob(blob, filename);
      setStatus(saved ? 'Đã lưu báo cáo phân bổ tồn kho.' : 'Đã hủy lưu báo cáo phân bổ tồn kho.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function downloadFastImportPackageLegacy() {
    if (!processedPurchaseSavedName || !processedSalesSavedName) {
      setStatus('Cần có cả FDI mua vào và FDI bán ra đã xử lý trước khi tạo workbook FAST.');
      return;
    }
    const progress = beginProgress('Đang tạo workbook FAST 5 sheet');
    setBusy(true);
    setStatus('Đang tạo workbook FAST gồm Hoadonmuahang, Hoadonbanhang, DM vật tư và DM khách hàng từ FDI đã xử lý...');
    try {
      const blob = await createVietmaxFastImportPackage(processedPurchaseSavedName, processedSalesSavedName, progress.operationId, {
        profile,
        purchaseOriginalSavedName: purchaseFile?.saved_name || '',
        salesOriginalSavedName: salesFile?.saved_name || '',
        purchaseFormMappingPresets: normalizeFormsForSave(purchaseFormMappingPresets),
        salesFormMappingPresets: normalizeFormsForSave(salesFormMappingPresets),
        purchaseCompanyGroupAssignments: groupAssignmentsFromRows(companyRows),
        salesCompanyGroupAssignments: groupAssignmentsFromRows(salesCompanyRows),
      });
      const saved = await saveBlob(blob, `${profile}_fast_import.xls`);
      setStatus(saved ? 'Đã lưu workbook FAST 5 sheet.' : 'Đã hủy lưu workbook FAST; dữ liệu đã xử lý vẫn được giữ.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function downloadFastImportPackage() {
    const { purchaseForms, salesForms, needsPurchase, needsSales } = fastImportRequirements;
    if (!needsPurchase && !needsSales) {
      setStatus('Chua co form mapping nao dang bat de xuat FAST.');
      return;
    }
    const restoredCaches = await restoreProcessedCachesFromSession();
    const purchaseCache = restoredCaches.purchaseSavedName || processedPurchaseSavedName;
    const salesCache = restoredCaches.salesSavedName || processedSalesSavedName;
    if ((needsPurchase && !purchaseCache) || (needsSales && !salesCache)) {
      setStatus(cacheDebugMessage('Cần có đủ FDI đã xử lý trước khi tạo workbook FAST theo form mapping.', needsPurchase, needsSales));
      return;
    }
    const progress = beginProgress('Dang tao workbook FAST theo form mapping');
    setBusy(true);
    setStatus('Dang tao workbook FAST theo cac form mapping dang bat...');
    try {
      const blob = await createVietmaxFastImportPackage(needsPurchase ? purchaseCache : '', needsSales ? salesCache : '', progress.operationId, {
        profile,
        purchaseOriginalSavedName: needsPurchase ? (purchaseFile?.saved_name || '') : '',
        salesOriginalSavedName: needsSales ? (salesFile?.saved_name || '') : '',
        purchaseFormMappingPresets: purchaseForms,
        salesFormMappingPresets: salesForms,
        purchaseCompanyGroupAssignments: groupAssignmentsFromRows(companyRows),
        salesCompanyGroupAssignments: groupAssignmentsFromRows(salesCompanyRows),
      });
      const saved = await saveBlob(blob, `${profile}_fast_import.xls`);
      setStatus(saved ? 'Da luu workbook FAST theo form mapping.' : 'Da huy luu workbook FAST; du lieu da xu ly van duoc giu.');
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStatus(`${message} ${cacheDebugMessage('Debug xuất FAST.', needsPurchase, needsSales)}`);
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  function inventoryStateForScope(scope: InventoryConfigScope) {
    const scopedProfile = profile === 'vietmax' || isTwoPhaseGenericProfile(profile);
    if (scopedProfile && scope === 'sales') {
      return {
        pairs: salesInventoryPairs,
        useDefault: salesUseDefaultInventoryPair,
        defaultPairId: salesDefaultInventoryPairId,
        rules: salesInventoryPairRules,
      };
    }
    if (scopedProfile && scope === 'purchase') {
      return {
        pairs: purchaseInventoryPairs,
        useDefault: purchaseUseDefaultInventoryPair,
        defaultPairId: purchaseDefaultInventoryPairId,
        rules: purchaseInventoryPairRules,
      };
    }
    return {
      pairs: inventoryPairs,
      useDefault: useDefaultInventoryPair,
      defaultPairId: defaultInventoryPairId,
      rules: inventoryPairRules,
    };
  }

  function addInventoryPair(scope: InventoryConfigScope = activeInventoryConfigScope) {
    const state = inventoryStateForScope(scope);
    const id = `pair-${scope}-${Date.now()}`;
    const nextPairs = [...state.pairs, { id, ma_kho: '', tk_vat_tu: '' }];
    const nextDefaultId = state.defaultPairId || id;
    updateScopedInventoryConfig({ pairs: nextPairs, defaultPairId: nextDefaultId }, scope);
  }

  function updateInventoryPair(index: number, field: 'ma_kho' | 'tk_vat_tu', value: string, scope: InventoryConfigScope = activeInventoryConfigScope) {
    const state = inventoryStateForScope(scope);
    updateScopedInventoryConfig({
      pairs: state.pairs.map((pair, rowIndex) => (rowIndex === index ? { ...pair, [field]: value.toUpperCase() } : pair)),
    }, scope);
  }

  function removeInventoryPair(index: number, scope: InventoryConfigScope = activeInventoryConfigScope) {
    const state = inventoryStateForScope(scope);
    const removed = state.pairs[index];
    const nextPairs = state.pairs.filter((_, rowIndex) => rowIndex !== index);
    updateScopedInventoryConfig({
      pairs: nextPairs,
      defaultPairId: removed?.id === state.defaultPairId ? (nextPairs[0]?.id || '') : state.defaultPairId,
      rules: state.rules.filter((rule) => rule.pair_id !== removed?.id),
    }, scope);
  }

  function updateInventoryDefaults(update: Partial<Pick<WorkflowState, 'useDefaultInventoryPair' | 'defaultInventoryPairId'>>, scope: InventoryConfigScope = activeInventoryConfigScope) {
    updateScopedInventoryConfig({
      useDefault: 'useDefaultInventoryPair' in update ? Boolean(update.useDefaultInventoryPair) : undefined,
      defaultPairId: 'defaultInventoryPairId' in update ? String(update.defaultInventoryPairId || '') : undefined,
    }, scope);
  }

  function addInventoryRule(scope: InventoryConfigScope = activeInventoryConfigScope) {
    const state = inventoryStateForScope(scope);
    updateScopedInventoryConfig({
      rules: [...state.rules, { source_col: 'M', operator: 'contains', value: '', pair_id: state.defaultPairId || state.pairs[0]?.id || '', enabled: true, priority: 1 }],
    }, scope);
  }

  function updateInventoryRule(index: number, update: Partial<InventoryRule>, scope: InventoryConfigScope = activeInventoryConfigScope) {
    const state = inventoryStateForScope(scope);
    updateScopedInventoryConfig({
      rules: state.rules.map((rule, rowIndex) => (rowIndex === index ? { ...rule, ...update } : rule)),
    }, scope);
  }

  function removeInventoryRule(index: number, scope: InventoryConfigScope = activeInventoryConfigScope) {
    const state = inventoryStateForScope(scope);
    updateScopedInventoryConfig({ rules: state.rules.filter((_, rowIndex) => rowIndex !== index) }, scope);
  }

  function updateScopedInventoryConfig(update: { pairs?: InventoryPair[]; useDefault?: boolean; defaultPairId?: string; rules?: InventoryRule[] }, scope: InventoryConfigScope = activeInventoryConfigScope) {
    const invalidation = inventoryOutputInvalidation(scope);
    const scopedProfile = profile === 'vietmax' || isTwoPhaseGenericProfile(profile);
    if (scopedProfile && scope === 'sales') {
      updateWorkflow(profile, {
        ...invalidation,
        ...(update.pairs ? { salesInventoryPairs: update.pairs } : {}),
        ...(update.useDefault !== undefined ? { salesUseDefaultInventoryPair: update.useDefault } : {}),
        ...(update.defaultPairId !== undefined ? { salesDefaultInventoryPairId: update.defaultPairId } : {}),
        ...(update.rules ? { salesInventoryPairRules: update.rules } : {}),
      });
      return;
    }
    if (scopedProfile && scope === 'purchase') {
      updateWorkflow(profile, {
        ...invalidation,
        ...(update.pairs ? { purchaseInventoryPairs: update.pairs } : {}),
        ...(update.useDefault !== undefined ? { purchaseUseDefaultInventoryPair: update.useDefault } : {}),
        ...(update.defaultPairId !== undefined ? { purchaseDefaultInventoryPairId: update.defaultPairId } : {}),
        ...(update.rules ? { purchaseInventoryPairRules: update.rules } : {}),
      });
      return;
    }
    updateWorkflow(profile, {
      ...invalidation,
      ...(update.pairs ? { inventoryPairs: update.pairs } : {}),
      ...(update.useDefault !== undefined ? { useDefaultInventoryPair: update.useDefault } : {}),
      ...(update.defaultPairId !== undefined ? { defaultInventoryPairId: update.defaultPairId } : {}),
      ...(update.rules ? { inventoryPairRules: update.rules } : {}),
    });
  }

  function productCodeReplacementsForScope(scope: InventoryConfigScope = activeInventoryConfigScope) {
    const scopedProfile = profile === 'vietmax' || isTwoPhaseGenericProfile(profile);
    if (scopedProfile && scope === 'sales') return salesProductCodeReplacements;
    return productCodeReplacements;
  }

  function updateScopedProductCodeReplacements(nextReplacements: Record<string, string>, scope: InventoryConfigScope = activeInventoryConfigScope) {
    const normalized = normalizeProductCodeReplacements(nextReplacements);
    const invalidation = inventoryOutputInvalidation(scope);
    const scopedProfile = profile === 'vietmax' || isTwoPhaseGenericProfile(profile);
    if (scopedProfile && scope === 'sales') {
      updateWorkflow(profile, { ...invalidation, salesProductCodeReplacements: normalized, salesReviewRows: [], salesReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, productCodeReplacements: normalized, purchaseReviewRows: [], purchaseReviewGenerated: false, priceGroups: [] });
  }

  function addProductCodeReplacement(scope: InventoryConfigScope = activeInventoryConfigScope) {
    const entries = Object.entries(productCodeReplacementsForScope(scope));
    let key = '';
    let index = entries.length + 1;
    while (!key || entries.some(([existing]) => existing === key)) {
      key = `MA_CU_${index}`;
      index += 1;
    }
    updateScopedProductCodeReplacements({ ...productCodeReplacementsForScope(scope), [key]: `MA_MOI_${entries.length + 1}` }, scope);
  }

  function updateProductCodeReplacement(index: number, field: 'from' | 'to', value: string, scope: InventoryConfigScope = activeInventoryConfigScope) {
    const entries = Object.entries(productCodeReplacementsForScope(scope));
    const current = entries[index] || ['', ''];
    entries[index] = field === 'from'
      ? [sanitizeDisplayProductCode(value), current[1] || '']
      : [current[0] || '', sanitizeDisplayProductCode(value)];
    updateScopedProductCodeReplacements(Object.fromEntries(entries), scope);
  }

  function commitProductCodeReplacements(nextReplacements: Record<string, string>, scope: InventoryConfigScope = activeInventoryConfigScope) {
    updateScopedProductCodeReplacements(nextReplacements, scope);
  }

  function removeProductCodeReplacement(index: number, scope: InventoryConfigScope = activeInventoryConfigScope) {
    const entries = Object.entries(productCodeReplacementsForScope(scope)).filter((_, rowIndex) => rowIndex !== index);
    updateScopedProductCodeReplacements(Object.fromEntries(entries), scope);
  }

  function applyRelatedProductCodeUpdates(updates: RelatedProductCodeUpdate[], scope: InventoryConfigScope = activeInventoryConfigScope) {
    if (!updates.length) return;
    const scopedProfile = profile === 'vietmax' || isTwoPhaseGenericProfile(profile);
    const isSales = scopedProfile && scope === 'sales';
    const sourceCompanies = isSales ? salesCompanyRows : companyRows;
    const sourceOverrides = isSales ? salesProductCodeOverrides : productCodeOverrides;
    const nextOverrides = { ...sourceOverrides };
    updates.forEach((item) => {
      const company = sourceCompanies[item.companyIndex];
      const code = sanitizeDisplayProductCode(item.code).toUpperCase();
      if (!company || !item.productName || !code) return;
      nextOverrides[productKey(companyConfigKey(company), item.productName)] = code;
    });
    const invalidation = inventoryOutputInvalidation(scope);
    if (isSales) {
      updateWorkflow(profile, { ...invalidation, salesProductCodeOverrides: nextOverrides, salesReviewRows: [], salesReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, productCodeOverrides: nextOverrides, purchaseReviewRows: [], purchaseReviewGenerated: false, priceGroups: [] });
  }

  async function importProductCodeReplacementFile(file: File | undefined, scope: InventoryConfigScope = activeInventoryConfigScope) {
    if (!file) return;
    setBusy(true);
    setStatus('Đang nhập danh sách đổi mã VT...');
    try {
      const result = await importProductCodeReplacements(file);
      updateScopedProductCodeReplacements({ ...productCodeReplacementsForScope(scope), ...result.product_code_replacements }, scope);
      setStatus(`Đã nhập ${result.count} dòng đổi mã VT. Bấm Áp dụng đổi mã để cập nhật preview và lưu cấu hình.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function inventoryOutputInvalidation(scope: InventoryConfigScope = activeInventoryConfigScope) {
    if (profile === 'vietmax') {
      return scope === 'sales' ? salesOutputInvalidation() : purchaseOutputInvalidation();
    }
    return stage >= 6 ? salesOutputInvalidation() : purchaseOutputInvalidation();
  }

  const statusTone = statusToneFromMessage(status, busy);
  const nextStage = adjacentEnterableStage(1);
  const canCreateMissingSalesCache = profile === 'vietmax' && stage === 11 && Boolean(salesFile && processedPurchaseSavedName && !processedSalesSavedName);
  const nextDisabled = busy || autoSavingConfig || (!canCreateMissingSalesCache && (!nextStage || (stage === 12 ? !(processedPurchaseSavedName && processedSalesSavedName) : !canEnterStage(nextStage.id))));

  return (
    <main className="desktop-shell">
      <section className={`app-card ${showLicenseBar ? '' : 'compact-flow'} ${profile === 'cao_thanh' ? 'legacy-flow' : ''}`}>
        <header className="app-header">
          <div className="profile-toolbar" aria-label="Company profile controls">
            <label className="profile-dropdown"><span>Công ty áp dụng</span><select value={visibleProfiles.some((item) => item.key === profile) ? profile : ''} disabled={busy || autoSavingConfig || !visibleProfiles.length} onChange={(event) => changeProfile(event.currentTarget.value as ProfileKey)}>{visibleProfiles.map((item) => <option key={item.key} value={item.key}>{item.label}</option>)}</select></label>
            {profile === 'ho_guom' && (
              <label className="profile-dropdown">
                <span>Chế độ Hồ Gươm</span>
                <select value={hoGuomMode} disabled={busy || autoSavingConfig} onChange={(event) => changeHoGuomMode(event.currentTarget.value as HoGuomMode)}>
                  <option value="estimate">Bóc tách dự toán</option>
                  <option value="formatter">Formatter mã VT</option>
                </select>
              </label>
            )}
            <button type="button" className="btn-secondary" disabled={busy || autoSavingConfig} onClick={saveCurrentConfig}>Lưu cấu hình</button>
            <button type="button" className="btn-secondary" disabled={busy || autoSavingConfig || updateBusy} onClick={() => { setUpdateModalOpen(true); void checkUpdate(); }}>{updateBusy ? 'Đang kiểm tra...' : 'Kiểm tra cập nhật'}</button>
            <button type="button" className="btn-secondary" disabled={busy || autoSavingConfig || updateBusy || !updateManifest?.available} onClick={() => void installUpdate()}>{updateManifest?.available ? 'Cập nhật ' + updateManifest.version : 'Cập nhật'}</button>
            {profile === 'vietmax' && (
              <div className="config-transfer-actions" aria-label="Xuất nhập cấu hình Vietmax">
                <select className="config-transfer-scope" value={configTransferScope} disabled={busy || autoSavingConfig} onChange={(event) => setConfigTransferScope(event.currentTarget.value as ConfigTransferScope)}>
                  <option value="purchase">Mua vào</option>
                  <option value="sales">Bán ra</option>
                  <option value="all">Tất cả</option>
                </select>
                <button type="button" className="btn-secondary" disabled={busy || autoSavingConfig} onClick={() => void exportSelectedVietmaxConfig()}>Xuất cấu hình</button>
                <button type="button" className="btn-secondary" disabled={busy || autoSavingConfig || (configTransferScope === 'purchase' ? (stage !== 0.5 && stage !== 1) : configTransferScope === 'sales' ? stage !== 6 : (stage !== 0.5 && stage !== 1))} onClick={() => void importSelectedVietmaxConfig()}>Nhập cấu hình</button>
              </div>
            )}
            <span className="app-version-badge" title="Phiên bản ứng dụng">v{updateManifest?.current_version || CLIENT_RELEASE_VERSION}</span>
          </div>
          {usesNativeStageShell && <StageNavigation stages={visibleStages} stage={stage} busy={busy || autoSavingConfig} canEnterStage={canEnterStage} goToStage={goToStage} />}
        </header>

        <div className={`status-strip ${statusTone === 'warning' ? 'status-warning' : ''} ${statusTone === 'error' ? 'status-error' : ''}`}><strong>Trạng thái</strong>{statusTone !== 'normal' && <span className="status-icon" aria-hidden="true">⚠</span>}<span>{busy ? 'Đang xử lý... ' : ''}{status}</span></div>

        {showLicenseBar && (
          <section className="license-bar">
            <div>
              <strong>License</strong>
              <span>{licenseChecking ? 'Đang kiểm tra license...' : licenseCheckError ? 'Không kiểm tra được license' : license?.status || 'Chưa có trạng thái license'}</span>
              <span className={licenseReady ? 'ok-text' : 'warning-text'}>{licenseProfileText}</span>
            </div>
            {!licenseReady && (
              <div className="license-form compact-form">
                <span className="license-server-fixed">License server: {FIXED_LICENSE_SERVER_URL} (dự phòng: {FIXED_LICENSE_SERVER_FALLBACK_URL})</span>
                <input placeholder="LICENSE_KEY" type="password" value={licenseForm.license_key} onChange={(event) => setLicenseForm({ ...licenseForm, license_key: event.currentTarget.value })} />
                <button type="button" disabled={busy || autoSavingConfig} onClick={submitLicense}>Kích hoạt</button>
              </div>
            )}
            {licenseCheckError && <button type="button" className="btn-secondary" disabled={licenseChecking || busy || autoSavingConfig} onClick={() => void checkLicenseStatus()}>{licenseChecking ? 'Đang kiểm tra...' : 'Kiểm tra lại'}</button>}
            {license?.activated && <button type="button" className="btn-secondary" disabled={busy || autoSavingConfig} onClick={refreshLicense}>Tải lại license</button>}
          </section>
        )}

        <section className="stage-frame">
          <div className={`stage-body ${profile === 'cao_thanh' ? 'legacy-stage-body' : ''}`}>
            {usesTwoPhaseFrame(profile) ? renderTwoPhaseStage() : renderProfileStage()}
          </div>
        </section>

        {usesNativeStageShell && (
          <footer className="action-bar">
            <button type="button" className="btn-secondary" disabled={visibleStages.findIndex((item) => item.id === stage) <= 0 || busy || autoSavingConfig} onClick={() => void goBack()}>Quay lại</button>
            <button type="button" className="btn-danger" disabled={busy || autoSavingConfig} onClick={resetWorkflow}>Làm lại</button>
            <div className="action-spacer" />
            <button type="button" disabled={nextDisabled} onClick={() => void goNext()}>Tiếp tục</button>
          </footer>
        )}
        {updateModalOpen && (
          <div className="modal-overlay" role="presentation" onMouseDown={() => setUpdateModalOpen(false)}>
            <section className="modal-content update-modal" role="dialog" aria-modal="true" aria-labelledby="update-modal-title" onMouseDown={(event) => event.stopPropagation()}>
              <div className="modal-header">
                <h2 id="update-modal-title">Cập nhật ứng dụng</h2>
                <button type="button" className="modal-close" onClick={() => setUpdateModalOpen(false)}>×</button>
              </div>
              <div className="modal-body update-modal-body">
                <div className="update-summary-grid">
                  <div><span>Phiên bản hiện tại</span><strong>{updateManifest?.current_version || 'Đang kiểm tra...'}</strong></div>
                  <div><span>Phiên bản mới</span><strong>{updateManifest?.available ? updateManifest.version : 'Không có bản mới'}</strong></div>
                </div>                <div className="update-release-notes">
                  <div className="update-notes"><strong>Bản đang chạy</strong><p>{updateManifest?.current_notes || 'Chưa có mô tả bản đang chạy.'}</p></div>
                  <div className="update-notes"><strong>Bản mới nhất trên server</strong><p>{updateManifest?.notes || 'Chưa có mô tả bản mới.'}</p></div>
                </div>
                <p className="muted">{updateBusy ? (updateProgress || 'Đang xử lý cập nhật...') : updateManifest?.available ? 'Đã có bản mới. Bấm Cập nhật để tải và khởi động lại ứng dụng.' : updateManifest ? 'Đang dùng phiên bản mới nhất.' : 'Bấm Kiểm tra lại để kiểm tra phiên bản trên server.'}</p>
              </div>
              <div className="modal-footer">
                <button type="button" className="btn-secondary" disabled={updateBusy} onClick={() => void checkUpdate()}>Kiểm tra lại</button>
                <button type="button" disabled={updateBusy || !updateManifest?.available} onClick={() => void installUpdate()}>{updateBusy ? 'Đang tải...' : 'Cập nhật'}</button>
              </div>
            </section>
          </div>
        )}
      </section>
    </main>
  );

  function renderTwoPhaseStage() {
    switch (stage) {
      case 0.5:
        return <FormatMappingStage purchaseFile={purchaseFile} salesFile={salesFile} purchaseGroups={purchaseProcessingGroups} salesGroups={salesProcessingGroups} purchasePresets={purchaseFormMappingPresets} salesPresets={salesFormMappingPresets} allowedScopes={['purchase', 'sales', 'both']} busy={busy} onAddGroup={addFormatGroup} onGroupChange={updateFormatGroup} onDeleteGroup={deleteFormatGroup} onAddForm={addFormatForm} onFormChange={updateFormatForm} onFormScopeChange={updateFormatFormScope} onDeleteForm={deleteFormatForm} onAddMapping={addFormatMappingRule} onMappingChange={updateFormatMappingRule} onRemoveMapping={removeFormatMappingRule} onUploadTemplate={uploadFormatTemplate} onRestoreDefaults={restoreDefaultFormMappings} onSave={saveFormatMappingConfig} />;
      case 1:
        return <UploadStage title={profile === 'vietmax' ? 'HD mua vào' : `HD mua vào ${selectedProfile.label}`} summary={purchaseFile} disabled={busy || !licenseReady} onUpload={(file) => upload('purchase', file)} />;
      case 2:
        return <MappingStage summary={purchaseFile} phase="purchase" scope={comparisonScope} setScope={updateComparisonScope} columns={purchaseColumns} invoiceStatuses={purchaseInvoiceStatuses} onColumnsChange={(update) => updateVietmaxColumns('purchase', update)} onInvoiceStatusSkipValuesChange={(values) => updateVietmaxInvoiceStatusSkipValues('purchase', values)} />;
      case 3:
        if (busy && !companyRows.length) return <LoadingStage title="Đang tải danh sách công ty" detail="Đang đọc workbook và gom công ty/MST/hàng hóa..." />;
        return <CompanyRulesStage companies={companyRows} selectedCompanyIndex={selectedCompanyIndex} processingGroups={purchaseProcessingGroups} productPreviewCodes={productPreviewCodes} productCodeOverrides={productCodeOverrides} productCodeReplacements={productCodeReplacements} wordRules={purchaseWordRules} repeatedPhrases={purchaseRepeatedPhraseRemovals} inventoryPairs={purchaseInventoryPairs} useDefaultInventoryPair={purchaseUseDefaultInventoryPair} defaultInventoryPairId={purchaseDefaultInventoryPairId} inventoryPairRules={purchaseInventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy={purchasePrefixStrategy} prefixMstDigits={prefixMstDigits} prefixNameWords={prefixNameWords} prefixNameChars={prefixNameChars} missingMstPrefixStrategy={prefixMissingMstStrategy} missingMstCompanies={purchaseMissingMstCompanies} onIncludeCompanyPrefixChange={updateIncludeCompanyPrefix} onCompanyPrefixChange={updateCompanyPrefix} onPrefixMstDigitsChange={updatePrefixMstDigits} onPrefixNameWordsChange={updatePrefixNameWords} onPrefixNameCharsChange={updatePrefixNameChars} onMissingMstPrefixStrategyChange={updateMissingMstPrefixStrategy} onApplyPrefixPresetToAll={applyPurchasePrefixPreset} onCompanySelect={selectCompany} onCompanyChange={updatePendingCompany} onCompanyGroupChange={updateCompanyGroup} onBulkCompanyChange={bulkUpdatePendingCompanies} onProductChange={updateCompanyProduct} onProductCodeChange={updateProductCode} onApplyChoices={applyCompanyAndProductChoices} onRefreshPreviews={refreshProductPreviews} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeatedPhrase} onAddRepeated={addRepeatedPhrase} onRemoveRepeated={removeRepeatedPhrase} onProductCodeReplacementChange={(index, field, value) => updateProductCodeReplacement(index, field, value, 'purchase')} onAddProductCodeReplacement={() => addProductCodeReplacement('purchase')} onRemoveProductCodeReplacement={(index) => removeProductCodeReplacement(index, 'purchase')} onProductCodeReplacementsCommit={(replacements) => commitProductCodeReplacements(replacements, 'purchase')} onApplyRelatedProductCodes={(updates) => applyRelatedProductCodeUpdates(updates, 'purchase')} onImportProductCodeReplacements={(file) => importProductCodeReplacementFile(file, 'purchase')} onAddInventoryPair={() => addInventoryPair('purchase')} onInventoryPairChange={(index, field, value) => updateInventoryPair(index, field, value, 'purchase')} onRemoveInventoryPair={(index) => removeInventoryPair(index, 'purchase')} onInventoryDefaultsChange={(update) => updateInventoryDefaults(update, 'purchase')} onAddInventoryRule={() => addInventoryRule('purchase')} onInventoryRuleChange={(index, update) => updateInventoryRule(index, update, 'purchase')} onRemoveInventoryRule={(index) => removeInventoryRule(index, 'purchase')} sonPhuongAllocationConfig={profile === 'son_phuong' ? inventoryAllocationConfig : undefined} onSonPhuongAllocationConfigChange={profile === 'son_phuong' ? updateInventoryAllocationConfig : undefined} />;
      case 4:
        if (busy) return <LoadingStage title="Đang tạo Review Mã VT mua vào" detail="Đang so sánh tên hàng và dựng danh sách mã cần kiểm tra..." progress={loadingProgress} />;
        if (!purchaseReviewGenerated) return <ProcessStage title="Review Mã VT mua vào" detail="Tạo danh sách các mã cần kiểm tra từ cấu hình công ty và hàng hóa đã áp dụng." buttonLabel="Tạo danh sách review" disabled={!purchaseFile || !companyRows.length} onProcess={profile === 'vietmax' ? runPurchaseReview : runGenericReview} />;
        return <ReviewStage rows={purchaseReviewRows} onApply={applyReviewChoices} disabled={!purchaseFile || busy} onRowChange={updateReviewRow} onBulkChange={bulkUpdateReviewRows} title="Review Mã VT mua vào" empty="Không có dòng Mã VT cần review." reviewScope={purchaseReviewScope} onReviewScopeChange={updatePurchaseReviewScope} />;
      case 5:
        if (busy) return <LoadingStage title="Đang tạo file mua vào" detail="Đang xử lý workbook và tạo cache file mua vào để dùng cho các stage bán ra..." progress={loadingProgress} />;
        return <ProcessStage title="Tạo file mua vào" detail="Xuất file FDI mua vào đã xử lý. File này sẽ được cache để dùng cho các stage bán ra và xuất FAST ở stage 15." buttonLabel="Xuất file mua vào" disabled={busy || !purchaseFile || (profile !== 'vietmax' && !companyRows.length)} onProcess={profile === 'vietmax' ? downloadProcessedPurchase : downloadGenericProcessedFile} />;
      case 6:
        return hasVietmaxPurchaseMatch(profile)
          ? <SalesEntryStage salesFile={salesFile} processedPurchaseReady={Boolean(processedPurchaseSavedName)} processedPurchaseStats={processedPurchaseStats} disabled={busy || !licenseReady} onSalesUpload={(file) => upload('sales', file)} onProcessedPurchaseUpload={(file) => uploadProcessed('purchase', file)} />
          : <UploadStage title={`HD bán ra ${selectedProfile.label}`} summary={salesFile} disabled={busy || !licenseReady} onUpload={(file) => upload('sales', file)} />;
      case 7:
        return <MappingStage summary={salesFile} phase="sales" scope={comparisonScope} setScope={updateComparisonScope} columns={salesColumns} invoiceStatuses={salesInvoiceStatuses} onColumnsChange={(update) => updateVietmaxColumns('sales', update)} onInvoiceStatusSkipValuesChange={(values) => updateVietmaxInvoiceStatusSkipValues('sales', values)} />;
      case 8:
        if (profile === 'son_phuong' && isInventoryAllocationRunning(inventoryAllocationJob)) return <LoadingStage title={'\u0110ang Ph\u00e2n kho S\u01a1n Ph\u01b0\u01a1ng'} detail={inventoryAllocationJob?.label || '\u0110ang t\u1ea1o ledger KHHVT v\u00e0 ph\u00e2n lo\u1ea1i KTP/KHOCK...'} progress={inventoryJobProgress(inventoryAllocationJob)} />;
        if (profile === 'son_phuong') return <InventoryAllocationStage profile="son_phuong" purchaseFile={purchaseFile} salesFile={salesFile} processedPurchaseSavedName={processedPurchaseSavedName} processedSalesSavedName={salesFile?.saved_name || ''} processedPurchaseStats={processedPurchaseStats} processedSalesStats={processedSalesStats} openingStockFile={openingStockFile} config={inventoryAllocationConfig} busy={busy} onProcessedPurchaseFileChange={(file) => uploadProcessed('purchase', file)} onProcessedSalesFileChange={(file) => uploadProcessed('sales', file)} onOpeningStockFileChange={updateOpeningStockFile} onConfigChange={updateInventoryAllocationConfig} />;
        if (!hasVietmaxPurchaseMatch(profile)) return <PlaceholderStage title="Khớp mua vào chưa bật cho profile này" detail="Profile này vẫn dùng cùng frame, nhưng chưa có logic khớp mua vào riêng nên sẽ đi thẳng sang công ty bán ra." />;
        if (busy) return <LoadingStage title="Đang khớp mua vào / bán ra" detail={processedPurchaseSavedName ? 'Đang so sánh hàng bán ra với file mua vào đã xử lý và áp dụng cấu hình khớp đã lưu...' : 'Đang tạo cache mua vào rồi khớp với file bán ra...'} progress={loadingProgress} />;
        return <MatchStage rows={matches} disabled={!salesFile || busy || (!purchaseFile && !processedPurchaseSavedName)} onRun={runSalesMatch} onSave={saveMatchChoices} onToggle={toggleMatch} onBulkToggle={bulkToggleMatches} onConversionChange={updateMatchConversion} emptyMessage={processedPurchaseSavedName || purchaseFile ? undefined : 'Cần tải file mua vào đã xử lý trước khi khớp mua/bán.'} />;
      case 9:
        if (profile === 'son_phuong') return <PlaceholderStage title={'Kh\u00f4ng \u00e1p d\u1ee5ng'} detail={'S\u01a1n Ph\u01b0\u01a1ng x\u1eed l\u00fd to\u00e0n b\u1ed9 c\u00f4ng ty b\u00e1n ra trong stage Ph\u00e2n kho.'} />;
        if (busy && !salesCompanyRows.length) return <LoadingStage title="Đang tải danh sách công ty bán ra" detail="Đang lọc các hàng hóa chưa khớp KVT/152 và gom theo công ty..." />;
        return <CompanyRulesStage companies={salesCompanyRows} selectedCompanyIndex={selectedSalesCompanyIndex} processingGroups={salesProcessingGroups} productPreviewCodes={salesProductPreviewCodes} productCodeOverrides={salesProductCodeOverrides} productCodeReplacements={salesProductCodeReplacements} wordRules={salesWordRules} repeatedPhrases={salesRepeatedPhraseRemovals} inventoryPairs={salesInventoryPairs} useDefaultInventoryPair={salesUseDefaultInventoryPair} defaultInventoryPairId={salesDefaultInventoryPairId} inventoryPairRules={salesInventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy={salesPrefixStrategy} prefixMstDigits={prefixMstDigits} prefixNameWords={prefixNameWords} prefixNameChars={prefixNameChars} missingMstPrefixStrategy={prefixMissingMstStrategy} missingMstCompanies={salesMissingMstCompanies} onIncludeCompanyPrefixChange={updateIncludeCompanyPrefix} onCompanyPrefixChange={updateSalesCompanyPrefix} onPrefixMstDigitsChange={updatePrefixMstDigits} onPrefixNameWordsChange={updatePrefixNameWords} onPrefixNameCharsChange={updatePrefixNameChars} onMissingMstPrefixStrategyChange={updateMissingMstPrefixStrategy} onApplyPrefixPresetToAll={applySalesPrefixPreset} onCompanySelect={selectSalesCompany} onCompanyChange={updateSalesPendingCompany} onCompanyGroupChange={updateSalesCompanyGroup} onBulkCompanyChange={bulkUpdateSalesPendingCompanies} onProductChange={updateSalesCompanyProduct} onProductCodeChange={updateSalesProductCode} onApplyChoices={applySalesCompanyAndProductChoices} onRefreshPreviews={refreshSalesProductPreviews} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeatedPhrase} onAddRepeated={addRepeatedPhrase} onRemoveRepeated={removeRepeatedPhrase} onProductCodeReplacementChange={(index, field, value) => updateProductCodeReplacement(index, field, value, 'sales')} onAddProductCodeReplacement={() => addProductCodeReplacement('sales')} onRemoveProductCodeReplacement={(index) => removeProductCodeReplacement(index, 'sales')} onProductCodeReplacementsCommit={(replacements) => commitProductCodeReplacements(replacements, 'sales')} onApplyRelatedProductCodes={(updates) => applyRelatedProductCodeUpdates(updates, 'sales')} onImportProductCodeReplacements={(file) => importProductCodeReplacementFile(file, 'sales')} onAddInventoryPair={() => addInventoryPair('sales')} onInventoryPairChange={(index, field, value) => updateInventoryPair(index, field, value, 'sales')} onRemoveInventoryPair={(index) => removeInventoryPair(index, 'sales')} onInventoryDefaultsChange={(update) => updateInventoryDefaults(update, 'sales')} onAddInventoryRule={() => addInventoryRule('sales')} onInventoryRuleChange={(index, update) => updateInventoryRule(index, update, 'sales')} onRemoveInventoryRule={(index) => removeInventoryRule(index, 'sales')} />;
      case 10:
        if (profile === 'son_phuong') return <SonPhuongAllocationReviewStage result={inventoryAllocationResult ?? inventoryAllocationJob?.result ?? null} />;
        if (busy) return <LoadingStage title="Đang tạo Review Mã VT bán ra" detail="Đang tạo danh sách review theo công ty/hàng hóa bán ra đã áp dụng..." progress={loadingProgress} />;
        if (!salesReviewGenerated) return <ProcessStage title="Review Mã VT bán ra" detail="Tạo danh sách các mã bán ra cần kiểm tra từ cấu hình đã áp dụng." buttonLabel="Tạo danh sách review" disabled={!salesFile || !salesCompanyRows.length} onProcess={profile === 'vietmax' ? runSalesReview : runGenericSalesReview} />;
        return <ReviewStage rows={salesReviewRows} onApply={applySalesReviewChoices} disabled={!salesFile || busy} onRowChange={updateSalesReviewRow} onBulkChange={bulkUpdateSalesReviewRows} title="Review Mã VT bán ra" empty="Không có dòng Mã VT bán ra cần review." reviewScope={salesReviewScope} onReviewScopeChange={updateSalesReviewScope} />;
      case 11:
        if (profile === 'son_phuong' && busy) return <LoadingStage title="Đang tạo FDI bán ra đã Phân kho" detail="Đang tạo và lưu cache FDI bán ra để dùng cho Form Mapping và Xuất FAST..." progress={loadingProgress} />;
        if (profile === 'son_phuong') return <ProcessStage title="Tạo FDI bán ra đã Phân kho" detail="Tải file FDI bán ra đã tách dòng, có Mã VT, Mã kho và TK vật tư để dùng cho 5 form FAST." buttonLabel="Xuất FDI bán ra" disabled={busy || !(inventoryAllocationResult?.job_id || inventoryAllocationJob?.result?.job_id || processedSalesSavedName)} onProcess={downloadSonPhuongProcessedSales} />;
        if (busy) return <LoadingStage title="Đang tạo file bán ra" detail="Đang xử lý workbook bán ra, áp dụng khớp mua vào và lưu cache cho phân bổ tồn kho..." progress={loadingProgress} />;
        return <ProcessStage title="Tạo file bán ra" detail="Xuất file FDI bán ra đã xử lý. File này sẽ được cache để dùng cho xuất FAST ở stage 15." buttonLabel="Xuất file bán ra" disabled={busy || !salesFile || (profile !== 'vietmax' && !salesCompanyRows.length)} onProcess={profile === 'vietmax' ? downloadProcessedSales : downloadGenericSalesProcessedFile} />;
      case 12:
        if (profile === 'son_phuong') return <InventoryAllocationOverviewStage result={inventoryAllocationResult ?? inventoryAllocationJob?.result ?? null} />;
        if (isInventoryAllocationRunning(inventoryAllocationJob)) return <LoadingStage title="Đang phân bổ tồn kho" detail={inventoryAllocationJob?.label || 'Đang chạy phân bổ từ file mua vào và bán ra đã xử lý...'} progress={inventoryJobProgress(inventoryAllocationJob)} />;
        return <InventoryAllocationStage purchaseFile={purchaseFile} salesFile={salesFile} processedPurchaseSavedName={processedPurchaseSavedName} processedSalesSavedName={processedSalesSavedName} processedPurchaseStats={processedPurchaseStats} processedSalesStats={processedSalesStats} openingStockFile={openingStockFile} config={inventoryAllocationConfig} busy={busy} onProcessedPurchaseFileChange={(file) => uploadProcessed('purchase', file)} onProcessedSalesFileChange={(file) => uploadProcessed('sales', file)} onOpeningStockFileChange={updateOpeningStockFile} onConfigChange={updateInventoryAllocationConfig} />;
      case 13:
        return <InventoryAllocationReportStage result={inventoryAllocationResult ?? inventoryAllocationJob?.result ?? null} busy={busy} />;
      case 14:
        if (busy) return <LoadingStage title="Đang tạo báo cáo Phân kho" detail="Đang tạo các sheet báo cáo và chuyển workbook sang định dạng XLS..." progress={loadingProgress} />;
        return <InventoryAllocationExportStage result={inventoryAllocationResult ?? inventoryAllocationJob?.result ?? null} busy={busy} onDownload={downloadInventoryReport} />;
      case 15: {
        const purchaseForms = normalizeFormsForSave(purchaseFormMappingPresets);
        const salesForms = normalizeFormsForSave(salesFormMappingPresets);
        const needsPurchase = activeFormsRequirePhase(purchaseForms, 'purchase') || activeFormsRequirePhase(salesForms, 'purchase');
        const needsSales = activeFormsRequirePhase(purchaseForms, 'sales') || activeFormsRequirePhase(salesForms, 'sales');
        return <FastImportExportStage processedPurchaseSavedName={processedPurchaseSavedName} processedSalesSavedName={processedSalesSavedName} processedPurchaseStats={processedPurchaseStats} processedSalesStats={processedSalesStats} needsPurchase={needsPurchase} needsSales={needsSales} busy={busy} onProcessedPurchaseUpload={(file) => uploadFastImportProcessed('purchase', file)} onProcessedSalesUpload={(file) => uploadFastImportProcessed('sales', file)} onDownload={downloadFastImportPackage} />;
      }
      default:
        return null;
    }
  }

  function renderProfileStage() {
    if (isHoGuomEstimate) {
      return <EstimateExtractorWorkflow ref={estimateWorkflowRef} licenseReady={licenseReady} onStatus={setStatus} stage={stage} onStageChange={goToStage} />;
    }
    if (isHoGuomFormatter && profileConfigLoading) {
      return <LoadingStage title="Đang nạp cấu hình Hồ Gươm" detail="Đang nạp schema FDI mua vào mới và ba form IMEXPNG, IMEXPN1, IMEXPC1..." />;
    }
    if (isGenericWorkflowProfile) {
      if (isTwoPhaseGenericProfile(profile)) return renderTwoPhaseStage();
      if (isTwoPhaseGenericProfile(profile)) {
        switch (stage) {
          case 0.5:
            return <FormatMappingStage purchaseFile={purchaseFile} salesFile={salesFile} purchaseGroups={purchaseProcessingGroups} salesGroups={salesProcessingGroups} purchasePresets={purchaseFormMappingPresets} salesPresets={salesFormMappingPresets} allowedScopes={['purchase', 'sales', 'both']} busy={busy} onAddGroup={addFormatGroup} onGroupChange={updateFormatGroup} onDeleteGroup={deleteFormatGroup} onAddForm={addFormatForm} onFormChange={updateFormatForm} onFormScopeChange={updateFormatFormScope} onDeleteForm={deleteFormatForm} onAddMapping={addFormatMappingRule} onMappingChange={updateFormatMappingRule} onRemoveMapping={removeFormatMappingRule} onUploadTemplate={uploadFormatTemplate} onRestoreDefaults={restoreDefaultFormMappings} onSave={saveFormatMappingConfig} />;
          case 1:
            return <UploadStage title={`HD mua vào ${selectedProfile.label}`} summary={purchaseFile} disabled={busy || !licenseReady} onUpload={(file) => upload('purchase', file)} />;
          case 2:
            return <MappingStage summary={purchaseFile} phase="purchase" scope={comparisonScope} setScope={updateComparisonScope} columns={purchaseColumns} invoiceStatuses={purchaseInvoiceStatuses} onColumnsChange={(update) => updateVietmaxColumns('purchase', update)} onInvoiceStatusSkipValuesChange={(values) => updateVietmaxInvoiceStatusSkipValues('purchase', values)} />;
          case 3:
            if (busy && !companyRows.length) return <LoadingStage title="Đang tải danh sách công ty mua vào" detail="Đang đọc workbook và gom công ty/MST/hàng hóa..." />;
            return <CompanyRulesStage companies={companyRows} selectedCompanyIndex={selectedCompanyIndex} processingGroups={purchaseProcessingGroups} productPreviewCodes={productPreviewCodes} productCodeOverrides={productCodeOverrides} productCodeReplacements={productCodeReplacements} wordRules={purchaseWordRules} repeatedPhrases={purchaseRepeatedPhraseRemovals} inventoryPairs={purchaseInventoryPairs} useDefaultInventoryPair={purchaseUseDefaultInventoryPair} defaultInventoryPairId={purchaseDefaultInventoryPairId} inventoryPairRules={purchaseInventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy={purchasePrefixStrategy} prefixMstDigits={prefixMstDigits} prefixNameWords={prefixNameWords} prefixNameChars={prefixNameChars} missingMstPrefixStrategy={prefixMissingMstStrategy} missingMstCompanies={purchaseMissingMstCompanies} onIncludeCompanyPrefixChange={updateIncludeCompanyPrefix} onCompanyPrefixChange={updateCompanyPrefix} onPrefixMstDigitsChange={updatePrefixMstDigits} onPrefixNameWordsChange={updatePrefixNameWords} onPrefixNameCharsChange={updatePrefixNameChars} onMissingMstPrefixStrategyChange={updateMissingMstPrefixStrategy} onApplyPrefixPresetToAll={applyPurchasePrefixPreset} onCompanySelect={selectCompany} onCompanyChange={updatePendingCompany} onCompanyGroupChange={updateCompanyGroup} onBulkCompanyChange={bulkUpdatePendingCompanies} onProductChange={updateCompanyProduct} onProductCodeChange={updateProductCode} onApplyChoices={applyCompanyAndProductChoices} onRefreshPreviews={refreshProductPreviews} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeatedPhrase} onAddRepeated={addRepeatedPhrase} onRemoveRepeated={removeRepeatedPhrase} onProductCodeReplacementChange={(index, field, value) => updateProductCodeReplacement(index, field, value, 'purchase')} onAddProductCodeReplacement={() => addProductCodeReplacement('purchase')} onRemoveProductCodeReplacement={(index) => removeProductCodeReplacement(index, 'purchase')} onProductCodeReplacementsCommit={(replacements) => commitProductCodeReplacements(replacements, 'purchase')} onApplyRelatedProductCodes={(updates) => applyRelatedProductCodeUpdates(updates, 'purchase')} onImportProductCodeReplacements={(file) => importProductCodeReplacementFile(file, 'purchase')} onAddInventoryPair={() => addInventoryPair('purchase')} onInventoryPairChange={(index, field, value) => updateInventoryPair(index, field, value, 'purchase')} onRemoveInventoryPair={(index) => removeInventoryPair(index, 'purchase')} onInventoryDefaultsChange={(update) => updateInventoryDefaults(update, 'purchase')} onAddInventoryRule={() => addInventoryRule('purchase')} onInventoryRuleChange={(index, update) => updateInventoryRule(index, update, 'purchase')} onRemoveInventoryRule={(index) => removeInventoryRule(index, 'purchase')} />;
          case 4:
            if (busy || !purchaseReviewGenerated) return <LoadingStage title="Đang tạo Review Mã VT mua vào" detail="Đang so sánh tên hàng và dựng danh sách mã cần kiểm tra..." progress={loadingProgress} />;
            return <ReviewStage rows={purchaseReviewRows} onApply={applyReviewChoices} disabled={!purchaseFile || busy} onRowChange={updateReviewRow} onBulkChange={bulkUpdateReviewRows} title="Review Mã VT mua vào" empty="Không có dòng Mã VT cần review." reviewScope={purchaseReviewScope} onReviewScopeChange={updatePurchaseReviewScope} />;
          case 5:
            if (busy) return <LoadingStage title="Đang tạo file mua vào" detail="Đang xử lý workbook và tạo cache file mua vào..." progress={loadingProgress} />;
            return <ProcessStage title="Tạo file mua vào" detail="Xuất file FDI mua vào đã xử lý. Workbook form mapping sẽ được tạo ở stage Xuất FAST." buttonLabel="Xuất file mua vào" disabled={busy || !purchaseFile || !companyRows.length} onProcess={downloadGenericProcessedFile} />;
          case 6:
            return <UploadStage title={`HD bán ra ${selectedProfile.label}`} summary={salesFile} disabled={busy || !licenseReady} onUpload={(file) => upload('sales', file)} />;
          case 7:
            return <MappingStage summary={salesFile} phase="sales" scope={comparisonScope} setScope={updateComparisonScope} columns={salesColumns} invoiceStatuses={salesInvoiceStatuses} onColumnsChange={(update) => updateVietmaxColumns('sales', update)} onInvoiceStatusSkipValuesChange={(values) => updateVietmaxInvoiceStatusSkipValues('sales', values)} />;
          case 8:
            if (busy && !salesCompanyRows.length) return <LoadingStage title="Đang tải danh sách công ty bán ra" detail="Đang đọc workbook bán ra và gom công ty/MST/hàng hóa..." />;
            return <CompanyRulesStage companies={salesCompanyRows} selectedCompanyIndex={selectedSalesCompanyIndex} processingGroups={salesProcessingGroups} productPreviewCodes={salesProductPreviewCodes} productCodeOverrides={salesProductCodeOverrides} productCodeReplacements={salesProductCodeReplacements} wordRules={salesWordRules} repeatedPhrases={salesRepeatedPhraseRemovals} inventoryPairs={salesInventoryPairs} useDefaultInventoryPair={salesUseDefaultInventoryPair} defaultInventoryPairId={salesDefaultInventoryPairId} inventoryPairRules={salesInventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy={salesPrefixStrategy} prefixMstDigits={prefixMstDigits} prefixNameWords={prefixNameWords} prefixNameChars={prefixNameChars} missingMstPrefixStrategy={prefixMissingMstStrategy} missingMstCompanies={salesMissingMstCompanies} onIncludeCompanyPrefixChange={updateIncludeCompanyPrefix} onCompanyPrefixChange={updateSalesCompanyPrefix} onPrefixMstDigitsChange={updatePrefixMstDigits} onPrefixNameWordsChange={updatePrefixNameWords} onPrefixNameCharsChange={updatePrefixNameChars} onMissingMstPrefixStrategyChange={updateMissingMstPrefixStrategy} onApplyPrefixPresetToAll={applySalesPrefixPreset} onCompanySelect={selectSalesCompany} onCompanyChange={updateSalesPendingCompany} onCompanyGroupChange={updateSalesCompanyGroup} onBulkCompanyChange={bulkUpdateSalesPendingCompanies} onProductChange={updateSalesCompanyProduct} onProductCodeChange={updateSalesProductCode} onApplyChoices={applySalesCompanyAndProductChoices} onRefreshPreviews={refreshSalesProductPreviews} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeatedPhrase} onAddRepeated={addRepeatedPhrase} onRemoveRepeated={removeRepeatedPhrase} onProductCodeReplacementChange={(index, field, value) => updateProductCodeReplacement(index, field, value, 'sales')} onAddProductCodeReplacement={() => addProductCodeReplacement('sales')} onRemoveProductCodeReplacement={(index) => removeProductCodeReplacement(index, 'sales')} onProductCodeReplacementsCommit={(replacements) => commitProductCodeReplacements(replacements, 'sales')} onApplyRelatedProductCodes={(updates) => applyRelatedProductCodeUpdates(updates, 'sales')} onImportProductCodeReplacements={(file) => importProductCodeReplacementFile(file, 'sales')} onAddInventoryPair={() => addInventoryPair('sales')} onInventoryPairChange={(index, field, value) => updateInventoryPair(index, field, value, 'sales')} onRemoveInventoryPair={(index) => removeInventoryPair(index, 'sales')} onInventoryDefaultsChange={(update) => updateInventoryDefaults(update, 'sales')} onAddInventoryRule={() => addInventoryRule('sales')} onInventoryRuleChange={(index, update) => updateInventoryRule(index, update, 'sales')} onRemoveInventoryRule={(index) => removeInventoryRule(index, 'sales')} />;
          case 9:
            if (busy || !salesReviewGenerated) return <LoadingStage title="Đang tạo Review Mã VT bán ra" detail="Đang tạo danh sách review theo công ty/hàng hóa bán ra đã áp dụng..." progress={loadingProgress} />;
            return <ReviewStage rows={salesReviewRows} onApply={applySalesReviewChoices} disabled={!salesFile || busy} onRowChange={updateSalesReviewRow} onBulkChange={bulkUpdateSalesReviewRows} title="Review Mã VT bán ra" empty="Không có dòng Mã VT bán ra cần review." reviewScope={salesReviewScope} onReviewScopeChange={updateSalesReviewScope} />;
          case 10:
            if (busy) return <LoadingStage title="Đang tạo file bán ra" detail="Đang xử lý workbook bán ra và tạo cache file bán ra..." progress={loadingProgress} />;
            return <ProcessStage title="Tạo file bán ra" detail="Xuất file FDI bán ra đã xử lý. Workbook form mapping sẽ được tạo ở stage Xuất FAST." buttonLabel="Xuất file bán ra" disabled={busy || !salesFile || !salesCompanyRows.length} onProcess={downloadGenericSalesProcessedFile} />;
          case 15: {
            const purchaseForms = normalizeFormsForSave(purchaseFormMappingPresets);
            const salesForms = normalizeFormsForSave(salesFormMappingPresets);
            const needsPurchase = activeFormsRequirePhase(purchaseForms, 'purchase') || activeFormsRequirePhase(salesForms, 'purchase');
            const needsSales = activeFormsRequirePhase(purchaseForms, 'sales') || activeFormsRequirePhase(salesForms, 'sales');
            return <FastImportExportStage processedPurchaseSavedName={processedPurchaseSavedName} processedSalesSavedName={processedSalesSavedName} processedPurchaseStats={processedPurchaseStats} processedSalesStats={processedSalesStats} needsPurchase={needsPurchase} needsSales={needsSales} busy={busy} onProcessedPurchaseUpload={(file) => uploadFastImportProcessed('purchase', file)} onProcessedSalesUpload={(file) => uploadFastImportProcessed('sales', file)} onDownload={downloadFastImportPackage} />;
          }
          default:
            return null;
        }
      }
      if (Number(stage) === 0.5) {
        return <FormatMappingStage purchaseFile={purchaseFile} salesFile={salesFile} defaultPurchaseColumns={purchaseFormSourceColumns} defaultSalesColumns={salesFormSourceColumns} defaultFormCount={profile === 'ho_guom' ? 3 : 5} purchaseGroups={purchaseProcessingGroups} salesGroups={salesProcessingGroups} purchasePresets={purchaseFormMappingPresets} salesPresets={salesFormMappingPresets} allowedScopes={['purchase']} busy={busy} onAddGroup={addFormatGroup} onGroupChange={updateFormatGroup} onDeleteGroup={deleteFormatGroup} onAddForm={addFormatForm} onFormChange={updateFormatForm} onFormScopeChange={updateFormatFormScope} onDeleteForm={deleteFormatForm} onAddMapping={addFormatMappingRule} onMappingChange={updateFormatMappingRule} onRemoveMapping={removeFormatMappingRule} onUploadTemplate={uploadFormatTemplate} onRestoreDefaults={restoreDefaultFormMappings} onSave={saveFormatMappingConfig} />;
      }
      if (Number(stage) === 4) {
        if (busy) return <LoadingStage title={`Đang tạo Review Mã VT ${selectedProfile.label}`} detail="Đang so sánh các tên hàng gần giống nhau theo danh sách công ty/hàng hóa đã áp dụng..." progress={loadingProgress} />;
        if (!purchaseReviewGenerated) return <ProcessStage title={`Review Mã VT ${selectedProfile.label}`} detail="Tạo danh sách các mã cần kiểm tra từ cấu hình đã áp dụng." buttonLabel="Tạo danh sách review" disabled={!purchaseFile || !companyRows.length} onProcess={runGenericReview} />;
        return <ReviewStage rows={purchaseReviewRows} onApply={applyReviewChoices} disabled={!purchaseFile || busy} onRowChange={updateReviewRow} onBulkChange={bulkUpdateReviewRows} title={`Review Ma VT ${selectedProfile.label}`} empty="Khong co dong Ma VT can review." reviewScope={purchaseReviewScope} onReviewScopeChange={updatePurchaseReviewScope} />;
      }
      if (Number(stage) === 5 && profile === 'cao_thanh') {
        return <CaoThanhPriceStage groups={priceGroups} filterPercent={priceFilterAllPercent} marginPercent={priceAdjustAllPercent} busy={busy} onRefresh={updateCaoThanhPriceGroups} onGroupPercentChange={updateCaoThanhGroupPercent} onBucketMarginChange={updateCaoThanhBucketMargin} onFilterPercentChange={updateCaoThanhPriceFilterAllPercent} onMarginPercentChange={updateCaoThanhPriceAdjustAllPercent} onApplyFilter={applyCaoThanhBulkPriceFilter} onApplyMargin={applyCaoThanhBulkMargin} onExportReport={exportCaoThanhPriceReport} />;
      }
      if ((Number(stage) === 5 && profile !== 'cao_thanh') || Number(stage) === 6) {
        if (busy) return <LoadingStage title={`Đang tạo cache file ${selectedProfile.label}`} detail="Đang xử lý workbook một lần và lưu cache để các lần xuất sau tải nhanh hơn..." progress={loadingProgress} />;
        return <ProcessStage title={`Xuất file ${selectedProfile.label}`} detail="Xuất một workbook .xls gồm FDI đã xử lý và các sheet form mapping theo cấu hình hiện tại." buttonLabel="Xuất file kết quả" disabled={busy || !purchaseFile || !companyRows.length} onProcess={downloadGenericProcessedFile} />;
      }
      switch (stage) {
        case 1:
          return <UploadStage title={selectedProfile.label} summary={purchaseFile} disabled={busy || !licenseReady} onUpload={(file) => upload('purchase', file)} />;
        case 2:
          return <GenericMappingStage summary={purchaseFile} columns={genericColumns} onColumnsChange={updateGenericColumns} />;
        case 3:
          if (busy && !companyRows.length) return <LoadingStage title={`Đang tải danh sách công ty ${selectedProfile.label}`} detail="Đang đọc workbook và gom công ty/MST/hàng hóa..." />;
          return <CompanyRulesStage companies={companyRows} selectedCompanyIndex={selectedCompanyIndex} processingGroups={purchaseProcessingGroups} productPreviewCodes={productPreviewCodes} productCodeOverrides={productCodeOverrides} productCodeReplacements={productCodeReplacements} wordRules={wordRules} repeatedPhrases={repeatedPhraseRemovals} inventoryPairs={inventoryPairs} useDefaultInventoryPair={useDefaultInventoryPair} defaultInventoryPairId={defaultInventoryPairId} inventoryPairRules={inventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy={purchasePrefixStrategy} prefixMstDigits={prefixMstDigits} prefixNameWords={prefixNameWords} prefixNameChars={prefixNameChars} missingMstPrefixStrategy={prefixMissingMstStrategy} missingMstCompanies={purchaseMissingMstCompanies} onIncludeCompanyPrefixChange={updateIncludeCompanyPrefix} onCompanyPrefixChange={updateCompanyPrefix} onPrefixMstDigitsChange={updatePrefixMstDigits} onPrefixNameWordsChange={updatePrefixNameWords} onPrefixNameCharsChange={updatePrefixNameChars} onMissingMstPrefixStrategyChange={updateMissingMstPrefixStrategy} onApplyPrefixPresetToAll={applyPurchasePrefixPreset} onCompanySelect={selectCompany} onCompanyChange={updatePendingCompany} onCompanyGroupChange={updateCompanyGroup} onBulkCompanyChange={bulkUpdatePendingCompanies} onProductChange={updateCompanyProduct} onProductCodeChange={updateProductCode} onApplyChoices={applyCompanyAndProductChoices} onRefreshPreviews={refreshProductPreviews} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeatedPhrase} onAddRepeated={addRepeatedPhrase} onRemoveRepeated={removeRepeatedPhrase} onProductCodeReplacementChange={(index, field, value) => updateProductCodeReplacement(index, field, value, 'generic')} onAddProductCodeReplacement={() => addProductCodeReplacement('generic')} onRemoveProductCodeReplacement={(index) => removeProductCodeReplacement(index, 'generic')} onProductCodeReplacementsCommit={(replacements) => commitProductCodeReplacements(replacements, 'generic')} onApplyRelatedProductCodes={(updates) => applyRelatedProductCodeUpdates(updates, 'generic')} onImportProductCodeReplacements={(file) => importProductCodeReplacementFile(file, 'generic')} onAddInventoryPair={() => addInventoryPair('generic')} onInventoryPairChange={(index, field, value) => updateInventoryPair(index, field, value, 'generic')} onRemoveInventoryPair={(index) => removeInventoryPair(index, 'generic')} onInventoryDefaultsChange={(update) => updateInventoryDefaults(update, 'generic')} onAddInventoryRule={() => addInventoryRule('generic')} onInventoryRuleChange={(index, update) => updateInventoryRule(index, update, 'generic')} onRemoveInventoryRule={(index) => removeInventoryRule(index, 'generic')} />;
        case 4:
          if (busy) return <LoadingStage title={`Đang tạo cache file ${selectedProfile.label}`} detail="Đang xử lý workbook một lần và lưu cache để các lần xuất sau tải nhanh hơn..." progress={loadingProgress} />;
          return <ProcessStage title={`Xuất file ${selectedProfile.label}`} detail="Xuất file đã xử lý Mã VT và file UP đi kèm theo cấu hình công ty/hàng hóa hiện tại." buttonLabel="Xuất file kết quả" disabled={busy || !purchaseFile || !companyRows.length} onProcess={downloadGenericProcessedFile} />;
        default:
          return null;
      }
    }
    return <LegacyProfileWorkspace profile={profile} label={selectedProfile.label} licenseReady={licenseReady} setShellStatus={setStatus} />;
  }
}

function phaseLabel(phase: StagePhase) {
  if (phase === 'purchase') return 'Mua vào';
  if (phase === 'sales') return 'Bán ra';
  if (phase === 'inventory') return 'Tồn kho';
  if (phase === 'allocation') return 'Ph\u00e2n kho';
  if (phase === 'fast') return 'Xuất FAST';
  if (phase === 'price') return 'Lọc đơn giá';
  if (phase === 'estimate') return 'Bóc tách';
  return 'Profile';
}

function statusToneFromMessage(message: string, busy = false): 'normal' | 'warning' | 'error' {
  if (busy) return 'normal';
  const raw = String(message || '').trim();
  if (!raw) return 'normal';
  const normalized = raw
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
  if (/\[[A-Z][A-Z0-9_]+;\s*[a-f0-9]+\]/.test(raw)) return 'error';
  const warningPatterns = [
    'can ',
    'can co',
    'chua ',
    'khong tao duoc',
    'khong doc duoc',
    'khong co',
    'thieu',
    'loi',
    'that bai',
    'canh bao',
    'hay ',
    'dang co thay doi',
    'tk vat tu',
    'ma kho',
  ];
  return warningPatterns.some((pattern) => normalized.includes(pattern)) ? 'warning' : 'normal';
}

function ReviewStage({ rows, onApply, onRowChange, onBulkChange, disabled = false, title = 'Review Mã VT mua vào', empty = 'Chưa có dòng review.', reviewScope = 'all', onReviewScopeChange }: { rows: ReviewRow[]; onApply?: () => void; onRowChange?: (index: number, update: Partial<ReviewRow>) => void; onBulkChange?: (indices: number[], confirmed: boolean) => void; disabled?: boolean; title?: string; empty?: string; reviewScope?: 'all' | 'company'; onReviewScopeChange?: (scope: 'all' | 'company') => void }) {
  const [collapsedGroups, setCollapsedGroups] = useState<Record<string, boolean>>({});
  const [bulkTarget, setBulkTarget] = useState('all');
  const groups = reviewDisplayGroups(rows, reviewScope);
  const allIndices = rows.map((_, index) => index);
  const bulkTargets = [
    { value: 'all', label: `Tất cả (${rows.length})`, indices: allIndices },
    ...groups
      .map((group, index) => ({
        value: `group:${index}`,
        label: `${group.title} (${reviewDisplayGroupCount(group)})`,
        indices: reviewDisplayGroupIndices(group),
      }))
      .filter((target) => target.indices.length),
  ];
  const selectedBulkTarget = bulkTargets.find((target) => target.value === bulkTarget) ?? bulkTargets[0];
  const canBulkUpdate = Boolean(onBulkChange && rows.length && !disabled);
  const bulkUpdate = (confirmed: boolean) => {
    if (!selectedBulkTarget?.indices.length) return;
    onBulkChange?.(selectedBulkTarget.indices, confirmed);
  };
  const toggleGroup = (groupTitle: string) => setCollapsedGroups((current) => ({ ...current, [groupTitle]: !current[groupTitle] }));

  useEffect(() => {
    if (!bulkTargets.some((target) => target.value === bulkTarget)) {
      setBulkTarget('all');
    }
  }, [bulkTarget, bulkTargets]);

  return (
    <div className="list-stage">
      <div className="stage-toolbar review-stage-toolbar">
        <p>{title}</p>
        <div className="toolbar-actions review-toolbar-actions">
          <div className="review-scope-panel" aria-label="Phạm vi review">
            <label className="inline-check">
              <input type="radio" name={`review-scope-${title}`} checked={reviewScope === 'all'} onChange={() => onReviewScopeChange?.('all')} disabled={disabled} />
              So sánh tất cả sản phẩm
            </label>
            <label className="inline-check">
              <input type="radio" name={`review-scope-${title}`} checked={reviewScope === 'company'} onChange={() => onReviewScopeChange?.('company')} disabled={disabled} />
              Chỉ trong cùng công ty
            </label>
          </div>
          <div className="review-bulk-panel" aria-label="Chọn nhanh review">
            <label className="review-bulk-select">
              <span>Phạm vi chọn</span>
              <select value={bulkTarget} disabled={!canBulkUpdate} onChange={(event) => setBulkTarget(event.currentTarget.value)}>
                {bulkTargets.map((target) => <option key={target.value} value={target.value}>{target.label}</option>)}
              </select>
            </label>
            <button type="button" className="btn-secondary" disabled={!canBulkUpdate || !selectedBulkTarget?.indices.length} onClick={() => bulkUpdate(false)}>Bỏ chọn</button>
            <button type="button" className="btn-secondary" disabled={!canBulkUpdate || !selectedBulkTarget?.indices.length} onClick={() => bulkUpdate(true)}>Chọn</button>
          </div>
          <div className="review-primary-actions">
            {onApply && <button type="button" disabled={disabled || !rows.length} onClick={onApply}>Áp dụng Review Mã VT</button>}
          </div>
        </div>
      </div>
      {!rows.length ? <p className="muted">{empty}</p> : <div className="inner-scroll"><ReviewTable groups={groups} collapsedGroups={collapsedGroups} onToggleGroup={toggleGroup} onRowChange={onRowChange} reviewScope={reviewScope} /></div>}
    </div>
  );
}

function MatchStage({ rows, disabled, onRun, onAutoRun, onSave, onToggle, onBulkToggle, onConversionChange, autoRun, emptyMessage = 'Chưa có dòng khớp. Dữ liệu sẽ được giữ khi quay lại stage 7 hoặc sang stage 9.' }: { rows: MatchRow[]; disabled: boolean; onRun: () => void; onAutoRun?: () => void; onSave?: () => void; onToggle: (index: number, confirmed: boolean) => void; onBulkToggle: (confirmed: boolean) => void; onConversionChange: (index: number, salesQty: string, purchaseQty: string) => void; autoRun?: boolean; emptyMessage?: string }) {
  const [autoRunStarted, setAutoRunStarted] = useState(false);

  useEffect(() => {
    if (autoRun && !disabled && rows.length === 0 && !autoRunStarted) {
      setAutoRunStarted(true);
      (onAutoRun || onRun)();
    }
  }, [autoRun, autoRunStarted, disabled, rows.length, onAutoRun, onRun]);

  useEffect(() => {
    if (!autoRun || rows.length > 0) setAutoRunStarted(false);
  }, [autoRun, rows.length]);

  return (
    <div className="list-stage">
      <div className="stage-toolbar">
        <p>Khớp HD mua vào đã xử lý với HD bán ra KVT/152.</p>
        <div className="match-toolbar-actions">
          <button type="button" className="btn-secondary" disabled={disabled} onClick={onRun}>Khớp lại</button>
          {onSave && <button type="button" className="btn-secondary" disabled={disabled || !rows.length} onClick={onSave}>Lưu cấu hình khớp</button>}
          <button type="button" className="btn-secondary" disabled={disabled || !rows.length} onClick={() => onBulkToggle(false)}>Bỏ chọn tất cả</button>
          <button type="button" className="btn-secondary" disabled={disabled || !rows.length} onClick={() => onBulkToggle(true)}>Chọn tất cả</button>
        </div>
      </div>
      {!rows.length ? <p className="muted">{emptyMessage}</p> : <div className="inner-scroll"><MatchTable rows={rows} onToggle={onToggle} onConversionChange={onConversionChange} /></div>}
    </div>
  );
}

function replacementDraftRowsFromMap(replacements: Record<string, string>): ProductCodeReplacementDraftRow[] {
  const rows = Object.entries(replacements).map(([from, to], index) => ({ id: `saved-${index}-${from}`, from, to }));
  return rows.length ? rows : [{ id: `empty-${Date.now()}`, from: '', to: '' }];
}

function productCodeReplacementMapFromDraftRows(rows: ProductCodeReplacementDraftRow[]) {
  return normalizeProductCodeReplacements(Object.fromEntries(rows.map((row) => [row.from, row.to])));
}

function stripCompanyPrefixFromCode(code: string, company: CompanyRow) {
  const cleanCode = sanitizeDisplayProductCode(code).toUpperCase();
  const prefix = sanitizeDisplayProductCode(committedCompanyPrefix(company)).toUpperCase();
  if (prefix && cleanCode.startsWith(`${prefix}.`)) return cleanCode.slice(prefix.length + 1);
  return cleanCode;
}

function codeWithCompanyPrefix(bodyCode: string, company: CompanyRow, includePrefix: boolean) {
  const body = sanitizeDisplayProductCode(bodyCode).toUpperCase();
  const prefix = sanitizeDisplayProductCode(committedCompanyPrefix(company)).toUpperCase();
  return includePrefix && prefix ? `${prefix}.${body}` : body;
}

function relatedProductCodeWarnings(
  companies: CompanyRow[],
  previewCodes: Record<string, string>,
  overrides: Record<string, string>,
  replacements: Record<string, string>,
  includePrefix: boolean,
  previousRows: RelatedProductCodeWarning[] = [],
): RelatedProductCodeWarning[] {
  const normalizedReplacements = normalizeProductCodeReplacements(replacements);
  const replacementEntries = Object.entries(normalizedReplacements)
    .map(([from, to]) => [sanitizeDisplayProductCode(from).toUpperCase(), sanitizeDisplayProductCode(to).toUpperCase()] as const)
    .filter(([from, to]) => from && to);
  if (!replacementEntries.length) return [];
  const previousById = new Map(previousRows.map((row) => [row.id, row]));
  const rows: RelatedProductCodeWarning[] = [];
  companies.forEach((company, companyIndex) => {
    if (isIgnoredCompany(company)) return;
    const selected = new Set(selectedProductNames(company));
    company.all_products.forEach((product) => {
      if (!selected.has(product.name)) return;
      const currentRawCode = productDisplayCode(company, product.name, previewCodes, overrides, {}, includePrefix);
      const bodyCode = stripCompanyPrefixFromCode(currentRawCode, company);
      if (!bodyCode) return;
      replacementEntries.forEach(([oldBase, newBase]) => {
        if (bodyCode === oldBase || !bodyCode.startsWith(oldBase)) return;
        const suffix = bodyCode.slice(oldBase.length);
        if (!/^X[A-Z0-9]/i.test(suffix)) return;
        const suggestedCode = codeWithCompanyPrefix(`${newBase}${suffix}`, company, includePrefix);
        const currentWithReplacement = productDisplayCode(company, product.name, previewCodes, overrides, normalizedReplacements, includePrefix).toUpperCase();
        if (sanitizeDisplayProductCode(currentWithReplacement).toUpperCase() === suggestedCode) return;
        const id = `${companyConfigKey(company)}|||${product.name}|||${oldBase}`;
        rows.push({
          id,
          companyIndex,
          productName: product.name,
          companyName: company.company || company.mst || `Công ty ${companyIndex + 1}`,
          baseCode: oldBase,
          currentCode: sanitizeDisplayProductCode(currentWithReplacement || currentRawCode).toUpperCase(),
          suggestedCode: previousById.get(id)?.suggestedCode || suggestedCode,
        });
      });
    });
  });
  return rows;
}

function ConfigModal({
  isOpen,
  onClose,
  companies,
  productPreviewCodes,
  productCodeOverrides,
  wordRules,
  repeatedPhrases,
  productCodeReplacements,
  includeCompanyPrefix,
  inventoryPairs,
  useDefaultInventoryPair,
  defaultInventoryPairId,
  inventoryPairRules,
  sonPhuongAllocationConfig,
  onSonPhuongAllocationConfigChange,
  onWordRuleChange,
  onAddWordRule,
  onRepeatedChange,
  onAddRepeated,
  onRemoveRepeated,
  onProductCodeReplacementChange,
  onAddProductCodeReplacement,
  onRemoveProductCodeReplacement,
  onProductCodeReplacementsCommit,
  onApplyRelatedProductCodes,
  onImportProductCodeReplacements,
  onAddInventoryPair,
  onInventoryPairChange,
  onRemoveInventoryPair,
  onInventoryDefaultsChange,
  onAddInventoryRule,
  onInventoryRuleChange,
  onRemoveInventoryRule,
  onRefreshPreviews,
  busy,
}: {
  isOpen: boolean;
  onClose: () => void;
  companies: CompanyRow[];
  productPreviewCodes: Record<string, string>;
  productCodeOverrides: Record<string, string>;
  wordRules: Record<string, string>;
  repeatedPhrases: string[];
  productCodeReplacements: Record<string, string>;
  includeCompanyPrefix: boolean;
  inventoryPairs: InventoryPair[];
  useDefaultInventoryPair: boolean;
  defaultInventoryPairId: string;
  inventoryPairRules: InventoryRule[];
  sonPhuongAllocationConfig?: InventoryAllocationConfig;
  onSonPhuongAllocationConfigChange?: (config: InventoryAllocationConfig) => void;
  onWordRuleChange: (index: number, field: 'from' | 'to', value: string) => void;
  onAddWordRule: () => void;
  onRepeatedChange: (index: number, value: string) => void;
  onAddRepeated: () => void;
  onRemoveRepeated: (index: number) => void;
  onProductCodeReplacementChange: (index: number, field: 'from' | 'to', value: string) => void;
  onAddProductCodeReplacement: () => void;
  onRemoveProductCodeReplacement: (index: number) => void;
  onProductCodeReplacementsCommit: (replacements: Record<string, string>) => void;
  onApplyRelatedProductCodes: (updates: RelatedProductCodeUpdate[]) => void;
  onImportProductCodeReplacements: (file: File | undefined) => void;
  onAddInventoryPair: () => void;
  onInventoryPairChange: (index: number, field: 'ma_kho' | 'tk_vat_tu', value: string) => void;
  onRemoveInventoryPair: (index: number) => void;
  onInventoryDefaultsChange: (update: Partial<Pick<WorkflowState, 'useDefaultInventoryPair' | 'defaultInventoryPairId'>>) => void;
  onAddInventoryRule: () => void;
  onInventoryRuleChange: (index: number, update: Partial<InventoryRule>) => void;
  onRemoveInventoryRule: (index: number) => void;
  onRefreshPreviews?: () => void;
  busy?: boolean;
}) {
  const [activeTab, setActiveTab] = useState<'words' | 'repeat' | 'replace' | 'related' | 'inventory'>('words');
  const [inventoryScope, setInventoryScope] = useState<'purchase' | 'sales'>('purchase');
  const [replacementDraftRows, setReplacementDraftRows] = useState<ProductCodeReplacementDraftRow[]>(() => replacementDraftRowsFromMap(productCodeReplacements));
  const [relatedRows, setRelatedRows] = useState<RelatedProductCodeWarning[]>([]);
  const wordEntries = Object.entries(wordRules);
  const replacementSnapshot = JSON.stringify(productCodeReplacements);
  useEffect(() => {
    if (!isOpen) return;
    setReplacementDraftRows(replacementDraftRowsFromMap(productCodeReplacements));
    setRelatedRows([]);
  }, [isOpen, replacementSnapshot]);
  const draftReplacementMap = productCodeReplacementMapFromDraftRows(replacementDraftRows);
  const detectedRelatedRows = relatedProductCodeWarnings(companies, productPreviewCodes, productCodeOverrides, draftReplacementMap, includeCompanyPrefix, relatedRows);
  const commitDraftReplacements = () => {
    const nextMap = productCodeReplacementMapFromDraftRows(replacementDraftRows);
    onProductCodeReplacementsCommit(nextMap);
    return nextMap;
  };
  const closeWithDraftCommit = () => {
    commitDraftReplacements();
    onClose();
  };
  const checkRelatedRows = () => {
    setRelatedRows(detectedRelatedRows);
    if (detectedRelatedRows.length) setActiveTab('related');
  };
  const applyRelatedRows = () => {
    const nextMap = commitDraftReplacements();
    const updates = relatedRows
      .map((row) => ({ companyIndex: row.companyIndex, productName: row.productName, code: row.suggestedCode }))
      .filter((row) => sanitizeDisplayProductCode(row.code));
    onApplyRelatedProductCodes(updates);
    const nextOverrides = { ...productCodeOverrides };
    updates.forEach((item) => {
      const company = companies[item.companyIndex];
      if (!company) return;
      nextOverrides[productKey(companyConfigKey(company), item.productName)] = sanitizeDisplayProductCode(item.code).toUpperCase();
    });
    setRelatedRows(relatedProductCodeWarnings(companies, productPreviewCodes, nextOverrides, nextMap, includeCompanyPrefix));
    onRefreshPreviews?.();
  };
  if (!isOpen) return null;
  return (
    <div className="modal-overlay" onClick={closeWithDraftCommit}>
      <div className="modal-content advanced-config-modal" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Cấu hình nâng cao</h2>
          <button className="modal-close" onClick={closeWithDraftCommit}>×</button>
        </div>
        <div className="modal-body">
          <div className="tab-list">
            <button className={`tab-button ${activeTab === 'replace' ? 'active' : ''}`} onClick={() => setActiveTab('replace')}>{'\u0110\u1ed5i m\u00e3 VT'}</button>
            <button className={`tab-button ${activeTab === 'words' ? 'active' : ''}`} onClick={() => setActiveTab('words')}>Từ riêng</button>
            <button className={`tab-button ${activeTab === 'repeat' ? 'active' : ''}`} onClick={() => setActiveTab('repeat')}>Từ lặp</button>
            <button className={`tab-button ${activeTab === 'related' ? 'active' : ''}`} onClick={() => { checkRelatedRows(); setActiveTab('related'); }}>Mã liên quan{detectedRelatedRows.length ? ` (${detectedRelatedRows.length})` : ''}</button>
            <button className={`tab-button ${activeTab === 'inventory' ? 'active' : ''}`} onClick={() => setActiveTab('inventory')}>Phân kho</button>
          </div>
          <div className={`tab-panel ${activeTab === 'words' ? 'active' : ''}`}>
            <div className="stage-toolbar compact-toolbar"><p>Quy tắc thay từ</p><button type="button" className="btn-secondary" onClick={onAddWordRule}>Thêm từ</button></div>
            <div className="compact-rule-list">{(wordEntries.length ? wordEntries : [['', '']]).map(([from, to], index) => <div className="rule-row" key={`word-rule-${index}`}><input placeholder="Từ gốc" value={from} onChange={(event) => onWordRuleChange(index, 'from', event.currentTarget.value)} /><input placeholder="Mã thay" value={to} onChange={(event) => onWordRuleChange(index, 'to', event.currentTarget.value)} /></div>)}</div>
            <div className="tab-apply-bar"><button type="button" disabled={busy} onClick={() => { onRefreshPreviews?.(); }}>Áp dụng từ riêng</button></div>
          </div>
          <div className={`tab-panel ${activeTab === 'repeat' ? 'active' : ''}`}>
            <div className="stage-toolbar compact-toolbar"><p>Cụm lặp chỉ giữ một lần</p><button type="button" className="btn-secondary" onClick={onAddRepeated}>Thêm cụm</button></div>
            <div className="repeated-rule-list">
              {(repeatedPhrases.length ? repeatedPhrases : ['']).map((phrase, index) => (
                <div className="repeated-rule-item" key={`repeated-${index}`}>
                  <input placeholder="Cụm lặp cần bỏ" value={phrase} onChange={(event) => onRepeatedChange(index, event.currentTarget.value)} />
                  <button type="button" className="btn-secondary compact-table-button" onClick={() => onRemoveRepeated(index)}>Xóa</button>
                </div>
              ))}
            </div>
            <div className="tab-apply-bar"><button type="button" disabled={busy} onClick={() => { onRefreshPreviews?.(); }}>Áp dụng từ lặp</button></div>
          </div>
          <div className={`tab-panel ${activeTab === 'replace' ? 'active' : ''}`}>
            <div className="stage-toolbar compact-toolbar">
              <p>{'\u0110\u1ed5i m\u00e3 VT sau khi t\u1ea1o m\u00e3'}</p>
              <div className="compact-toolbar-actions">
                <label className="btn-secondary file-import-button">
                  {'Nh\u1eadp Excel'}
                  <input type="file" accept=".xls,.xlsx,.xlsm" onChange={(event) => { onImportProductCodeReplacements(event.currentTarget.files?.[0]); event.currentTarget.value = ''; }} />
                </label>
                <button type="button" className="btn-secondary" onClick={() => setReplacementDraftRows((rows) => [...rows, { id: `draft-${Date.now()}-${rows.length}`, from: '', to: '' }])}>Thêm dòng</button>
              </div>
            </div>
            <div className="compact-rule-list">
              {replacementDraftRows.map((row, index) => (
                <div className="rule-row product-code-replacement-row" key={row.id}>
                  <input placeholder={'M\u00e3 VT g\u1ed1c, v\u00ed d\u1ee5 TU100'} value={row.from} onChange={(event) => setReplacementDraftRows((rows) => rows.map((item, rowIndex) => rowIndex === index ? { ...item, from: sanitizeDisplayProductCode(event.currentTarget.value) } : item))} />
                  <input placeholder={'M\u00e3 VT m\u1edbi, v\u00ed d\u1ee5 TU100.001'} value={row.to} onChange={(event) => setReplacementDraftRows((rows) => rows.map((item, rowIndex) => rowIndex === index ? { ...item, to: sanitizeDisplayProductCode(event.currentTarget.value) } : item))} />
                  <button type="button" className="btn-secondary compact-table-button" onClick={() => setReplacementDraftRows((rows) => rows.filter((_, rowIndex) => rowIndex !== index))}>Xóa</button>
                </div>
              ))}
            </div>
            <p className="muted">{'Quy t\u1eafc b\u1ecf qua prefix c\u00f4ng ty: TU100 c\u0169ng \u00e1p d\u1ee5ng cho TA.TU100, TP.TU100...'}</p>
            <div className="tab-apply-bar">
              <button type="button" className="btn-secondary" disabled={busy} onClick={checkRelatedRows}>Kiểm tra mã liên quan</button>
              <button type="button" disabled={busy} onClick={() => { commitDraftReplacements(); onRefreshPreviews?.(); }}>{'\u00c1p d\u1ee5ng \u0111\u1ed5i m\u00e3'}</button>
            </div>
          </div>
          <div className={`tab-panel ${activeTab === 'related' ? 'active' : ''}`}>
            <div className="stage-toolbar compact-toolbar">
              <p>Mã VT liên quan từ quy tắc đổi mã</p>
              <button type="button" className="btn-secondary" disabled={busy} onClick={checkRelatedRows}>Kiểm tra lại</button>
            </div>
            {!relatedRows.length ? <p className="muted">Không có mã liên quan cần sửa. Bấm Kiểm tra lại sau khi đổi mã như TU100 -&gt; TU100.001.</p> : (
              <div className="compact-rule-list related-code-list">
                {relatedRows.map((row, index) => (
                  <div className="rule-row related-code-row" key={row.id}>
                    <span className="related-code-source" title={`${row.companyName} | ${row.productName}`}>{row.baseCode} | {row.currentCode}</span>
                    <input value={row.suggestedCode} onChange={(event) => setRelatedRows((rows) => rows.map((item, rowIndex) => rowIndex === index ? { ...item, suggestedCode: sanitizeDisplayProductCode(event.currentTarget.value).toUpperCase() } : item))} />
                  </div>
                ))}
              </div>
            )}
            <p className="muted">Ví dụ TU100 -&gt; TU100.001 sẽ dò các mã như TU100X45X4 và đề xuất TU100.001X45X4. Khi áp dụng, app ghi vào mã sửa tay của hàng hóa tương ứng.</p>
            <div className="tab-apply-bar"><button type="button" disabled={busy || !relatedRows.length} onClick={applyRelatedRows}>Áp dụng mã liên quan & kiểm tra lại</button></div>
          </div>
          <div className={`tab-panel ${activeTab === 'inventory' ? 'active' : ''}`}>
            {sonPhuongAllocationConfig && <div className="segmented-control">
              <button type="button" className={inventoryScope === 'purchase' ? 'active' : ''} onClick={() => setInventoryScope('purchase')}>{'Mua v\u00e0o'}</button>
              <button type="button" className={inventoryScope === 'sales' ? 'active' : ''} onClick={() => setInventoryScope('sales')}>{'B\u00e1n ra'}</button>
            </div>}
            {(!sonPhuongAllocationConfig || inventoryScope === 'purchase')
              ? <InventoryPairEditor pairs={inventoryPairs} useDefault={useDefaultInventoryPair} defaultPairId={defaultInventoryPairId} rules={inventoryPairRules} busy={busy ?? false} onAddPair={onAddInventoryPair} onPairChange={onInventoryPairChange} onRemovePair={onRemoveInventoryPair} onDefaultsChange={onInventoryDefaultsChange} onAddRule={onAddInventoryRule} onRuleChange={onInventoryRuleChange} onRemoveRule={onRemoveInventoryRule} fixed={Boolean(sonPhuongAllocationConfig)} />
              : <SonPhuongSalesPairEditor config={sonPhuongAllocationConfig} busy={busy ?? false} onChange={(next) => onSonPhuongAllocationConfigChange?.(next)} />}
            <div className="tab-apply-bar"><button type="button" disabled={busy} onClick={() => { onRefreshPreviews?.(); }}>Áp dụng phân kho</button></div>
          </div>
        </div>
        <div className="modal-footer">
          <button type="button" className="btn-secondary" onClick={closeWithDraftCommit}>Đóng</button>
        </div>
      </div>
    </div>
  );
}

const maxCodeLength = 50;

type CompanyDisplayGroup = {
  title: string;
  rows: Array<{ company: CompanyRow; index: number }>;
  className: string;
};

type ReviewDisplayGroup = {
  title: string;
  displayTitle?: string;
  key?: string;
  rows: Array<{ row: ReviewRow; index: number }>;
  className: string;
  children?: ReviewDisplayGroup[];
};

function FormatMappingStage({
  purchaseFile,
  salesFile,
  defaultPurchaseColumns,
  defaultSalesColumns,
  defaultFormCount = 5,
  purchaseGroups,
  salesGroups,
  purchasePresets,
  salesPresets,
  allowedScopes = ['purchase', 'sales', 'both'],
  busy,
  onAddGroup,
  onGroupChange,
  onDeleteGroup,
  onAddForm,
  onFormChange,
  onFormScopeChange,
  onDeleteForm,
  onAddMapping,
  onMappingChange,
  onRemoveMapping,
  onUploadTemplate,
  onRestoreDefaults,
  onSave,
}: {
  purchaseFile: UploadSummary | null;
  salesFile: UploadSummary | null;
  defaultPurchaseColumns?: FormColumn[];
  defaultSalesColumns?: FormColumn[];
  defaultFormCount?: number;
  purchaseGroups: ProcessingGroup[];
  salesGroups: ProcessingGroup[];
  purchasePresets: FormMappingPreset[];
  salesPresets: FormMappingPreset[];
  allowedScopes?: FormatScope[];
  busy: boolean;
  onAddGroup: (scope: FormatScope, label?: string) => void;
  onGroupChange: (scope: FormatScope, groupId: string, update: Partial<ProcessingGroup>) => void;
  onDeleteGroup: (scope: FormatScope, groupId: string) => void;
  onAddForm: (scope: FormatScope, groupId?: string) => void;
  onFormChange: (scope: FormatScope, formId: string, update: Partial<FormMappingPreset>) => void;
  onFormScopeChange: (scope: FormatScope, formId: string, nextScope: FormatScope) => void;
  onDeleteForm: (scope: FormatScope, formId: string) => void;
  onAddMapping: (scope: FormatScope, formId: string) => void;
  onMappingChange: (scope: FormatScope, formId: string, index: number, update: Partial<NonNullable<FormMappingPreset['mappings']>[number]>) => void;
  onRemoveMapping: (scope: FormatScope, formId: string, index: number) => void;
  onUploadTemplate: (scope: FormatScope, formId: string, file: File | undefined) => void;
  onRestoreDefaults: () => void;
  onSave: () => void;
}) {
  const scopes = allowedScopes.length ? allowedScopes : (Object.keys(FORMAT_SCOPE_LABELS) as FormatScope[]);
  const [activeScope, setActiveScope] = useState<FormatScope>(scopes[0] || 'purchase');
  useEffect(() => {
    if (!scopes.includes(activeScope)) {
      setActiveScope(scopes[0] || 'purchase');
    }
  }, [activeScope, scopes]);
  const groups = activeScope === 'sales' ? salesGroups : purchaseGroups;
  const [activeGroupId, setActiveGroupId] = useState(MATERIALS_GROUP_ID);
  useEffect(() => {
    if (!groups.some((group) => group.id === activeGroupId)) {
      setActiveGroupId(groups[0]?.id || MATERIALS_GROUP_ID);
    }
  }, [activeGroupId, groups]);
  const forms = visibleFormatFormsForScope(activeScope, purchasePresets, salesPresets).filter((form) => (form.group_id || MATERIALS_GROUP_ID) === activeGroupId);
  const purchaseColumns = purchaseFile ? columnsFromUploadSummary(purchaseFile) : formatSourceColumns(defaultPurchaseColumns);
  const salesColumns = salesFile ? columnsFromUploadSummary(salesFile) : formatSourceColumns(defaultSalesColumns);
  const sourceColumns = (phase: string | undefined) => phase === 'sales' ? salesColumns : purchaseColumns;
  const sourcePhaseOptions = Array.from(new Set(scopes.flatMap((scope) => (scope === 'sales' ? ['sales'] : scope === 'purchase' ? ['purchase'] : ['purchase', 'sales'])))) as Array<'purchase' | 'sales'>;
  const [showGroupModal, setShowGroupModal] = useState(false);
  const [scopeEditForm, setScopeEditForm] = useState<FormMappingPreset | null>(null);

  return (
    <div className="format-mapping-stage">
      <div className="stage-toolbar format-toolbar">
        <div>
          <p>Mapping form xuất</p>
          <span className="muted">File FDI được nạp sẵn theo mặc định. Tải mẫu output trong từng form để thay cột đích và xem preview.</span>
        </div>
        <div className="format-toolbar-actions">
          <button type="button" className="btn-secondary" disabled={busy} onClick={() => onAddForm(activeScope, activeGroupId)}>Thêm form</button>
          <button type="button" className="btn-secondary" disabled={busy} onClick={onRestoreDefaults}>Khôi phục {defaultFormCount} form mặc định</button>
          <button type="button" className="btn-secondary" disabled={busy} onClick={onSave}>Lưu cấu hình</button>
          <button type="button" disabled={busy} onClick={onSave}>Lưu form mapping</button>
        </div>
      </div>
      <div className="format-sticky-controls">
        <section className="format-control-panel format-scope-panel">
          <div className="format-scope-group-row">
            <div>
              <strong>Phạm vi form</strong>
              <div className="format-scope-tabs">
                {scopes.map((scope) => (
                  <button key={scope} type="button" className={activeScope === scope ? 'active' : ''} onClick={() => setActiveScope(scope)}>{FORMAT_SCOPE_LABELS[scope]}</button>
                ))}
              </div>
            </div>
            <label className="format-group-picker">
              <span>Nhóm</span>
              <select value={activeGroupId} onChange={(event) => setActiveGroupId(event.currentTarget.value)}>
                {groups.map((group) => <option key={group.id} value={group.id}>{group.label}</option>)}
              </select>
            </label>
          </div>
        </section>
        <FormatGroupSummary groups={groups} onOpen={() => setShowGroupModal(true)} />
      </div>
      {showGroupModal && <FormatGroupModal scope={activeScope} groups={groups} busy={busy} onAddGroup={onAddGroup} onGroupChange={onGroupChange} onDeleteGroup={onDeleteGroup} onClose={() => setShowGroupModal(false)} />}
      {scopeEditForm && (
        <FormatScopeModal
          form={scopeEditForm}
          activeScope={activeScope}
          allowedScopes={allowedScopes}
          busy={busy}
          onSave={(nextScope) => {
            onFormScopeChange(activeScope, scopeEditForm.id, nextScope);
            if (scopes.includes(nextScope)) setActiveScope(nextScope);
            setScopeEditForm(null);
          }}
          onClose={() => setScopeEditForm(null)}
        />
      )}
      <div className="format-sample-grid">
        <FormatSamplePanel title="FDI mua vào input" summary={purchaseFile} columns={purchaseColumns} />
        {scopes.some((scope) => scope !== 'purchase') && <FormatSamplePanel title="FDI bán ra input" summary={salesFile} columns={salesColumns} />}
      </div>
      <div className="format-form-list">
        {forms.map((form) => {
          const outputCols = form.output_columns?.length ? form.output_columns : outputColumns(['Cột A', 'Cột B', 'Cột C']);
          const mappings = form.mappings || [];
          const orderedMappings = sortedMappingsForDisplay(mappings, outputCols);
          return (
            <section className="format-form-card" key={form.id}>
              <header className="format-form-header">
                <div className="format-form-title">
                  <input value={form.label} onChange={(event) => onFormChange(activeScope, form.id, { label: event.currentTarget.value })} />
                  <span>{form.template_original_name || form.sheet || form.builtin_exporter || 'Form tùy chỉnh'}</span>
                </div>
                <div className="format-form-actions">
                  <button type="button" className="btn-secondary" disabled={busy} onClick={() => setScopeEditForm(form)}>Phạm vi: {FORMAT_SCOPE_LABELS[formatScopeOfForm(form, activeScope)]}</button>
                  <select value={form.group_id || MATERIALS_GROUP_ID} onChange={(event) => onFormChange(activeScope, form.id, { group_id: event.currentTarget.value })}>
                    {groups.map((group) => <option key={group.id} value={group.id}>{group.label}</option>)}
                  </select>
                  <label className="template-upload-button">
                    <input type="file" accept=".xls,.xlsx,.xlsm" disabled={busy} onChange={(event) => onUploadTemplate(activeScope, form.id, event.currentTarget.files?.[0])} />
                    Tải mẫu output
                  </label>
                  <button type="button" className="btn-secondary" disabled={busy} onClick={() => onAddMapping(activeScope, form.id)}>Thêm mapping</button>
                  <button type="button" className="btn-secondary compact-table-button" disabled={busy} onClick={() => onDeleteForm(activeScope, form.id)}>Xóa form</button>
                </div>
              </header>
              <div className="format-output-preview">
                <FormatSamplePanel title="Cột output" columns={outputCols} preview={form.output_preview} />
              </div>
              <div className="format-mapping-table-wrap">
                <table className="company-table format-mapping-table">
                  <thead>
                    <tr><th>Mapping</th><th>Input</th><th>Rule / Logic</th><th>Output</th><th /></tr>
                  </thead>
                  <tbody>
                    {orderedMappings.map(({ rule, originalIndex }) => {
                      const phase = rule.source_phase === 'sales' ? 'sales' : 'purchase';
                      const sourceOptions = sourceColumns(phase);
                      const ruleControlValue = mappingRuleControlValue(rule);
                      return (
                        <tr key={`${form.id}-${originalIndex}`}>
                          <td className="mapping-expression">{mappingExpression(rule, sourceOptions, outputCols)}</td>
                          <td className="mapping-source-cell">
                            <div className="mapping-input-pair">
                              <select value={phase} onChange={(event) => onMappingChange(activeScope, form.id, originalIndex, { source_phase: event.currentTarget.value })}>
                                {sourcePhaseOptions.map((option) => <option key={option} value={option}>{option === 'sales' ? 'Bán ra' : 'Mua vào'}</option>)}
                              </select>
                              {ruleControlValue === 'if_rules' ? (
                                <span className="mapping-input-note">Cột nằm trong IF</span>
                              ) : (
                                <select value={rule.source_col || ''} onChange={(event) => onMappingChange(activeScope, form.id, originalIndex, { source_col: event.currentTarget.value })}>
                                  <option value="">Chọn cột</option>
                                  {sourceOptions.map((column) => <option key={column.letter} value={column.letter}>{column.label}</option>)}
                                </select>
                              )}
                            </div>
                          </td>
                          <td className="mapping-rule-cell">
                            <div className="mapping-rule-stack">
                              <select value={ruleControlValue} onChange={(event) => {
                                const nextType = event.currentTarget.value;
                                if (nextType === 'if_rules') {
                                  onMappingChange(activeScope, form.id, originalIndex, {
                                    source_type: nextType,
                                    transform: undefined,
                                    transform_rules: defaultConditionRules(rule, rule.condition_source_col || rule.fallback_source_col || rule.source_col || ''),
                                    condition_source_col: '',
                                    source_col: '',
                                    fallback_source_col: '',
                                  });
                                } else {
                                  onMappingChange(activeScope, form.id, originalIndex, {
                                    source_type: nextType === 'text' ? textSourceTypeForValue(rule.value) : nextType,
                                    transform: supportsCustomTransformRules(rule.transform) ? undefined : rule.transform,
                                    transform_rules: undefined,
                                    condition_source_col: '',
                                  });
                                }
                              }}>
                                {MAPPING_RULE_OPTIONS.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                              </select>
                              {isTextMappingRule(rule) && <input value={rule.value || ''} placeholder="Nhập giá trị hoặc {C}" onChange={(event) => onMappingChange(activeScope, form.id, originalIndex, { value: event.currentTarget.value, source_type: textSourceTypeForValue(event.currentTarget.value) })} />}
                              {ruleControlValue === 'if_rules' && (
                                <TransformRuleEditor
                                  rules={defaultConditionRules(rule, rule.condition_source_col || rule.fallback_source_col || rule.source_col || '')}
                                  sourceColumns={sourceOptions}
                                  onChange={(rules) => onMappingChange(activeScope, form.id, originalIndex, { transform_rules: rules })}
                                />
                              )}
                            </div>
                          </td>
                          <td className="mapping-output-cell">
                            <select value={rule.target_col || ''} onChange={(event) => onMappingChange(activeScope, form.id, originalIndex, { target_col: event.currentTarget.value })}>
                              <option value="">Chọn output</option>
                              {outputCols.map((column) => <option key={column.letter} value={column.letter}>{column.label}</option>)}
                            </select>
                          </td>
                          <td><button type="button" className="btn-secondary compact-table-button" disabled={busy} onClick={() => onRemoveMapping(activeScope, form.id, originalIndex)}>Xóa</button></td>
                        </tr>
                      );
                    })}
                    {!mappings.length && <tr><td colSpan={5} className="muted">Chưa có mapping. Bấm Thêm mapping để tạo dòng mới.</td></tr>}
                  </tbody>
                </table>
              </div>
            </section>
          );
        })}
        {!forms.length && <p className="muted">Chưa có form trong nhóm này. Bấm Thêm form để tạo format mới.</p>}
      </div>
    </div>
  );
}

function FormatGroupSummary({ groups, onOpen }: { groups: ProcessingGroup[]; onOpen: () => void }) {
  return (
    <section className="format-control-panel format-group-summary">
      <div className="format-control-heading">
        <strong>Nhóm công ty</strong>
        <button type="button" className="btn-secondary compact-table-button" onClick={onOpen}>Quản lý nhóm</button>
      </div>
      <div className="format-group-chip-list">
        {groups.map((group) => <span key={group.id}>{group.label}</span>)}
      </div>
    </section>
  );
}

function FormatScopeModal({ form, activeScope, allowedScopes, busy, onSave, onClose }: { form: FormMappingPreset; activeScope: FormatScope; allowedScopes: FormatScope[]; busy: boolean; onSave: (scope: FormatScope) => void; onClose: () => void }) {
  const [selectedScope, setSelectedScope] = useState<FormatScope>(formatScopeOfForm(form, activeScope));
  useEffect(() => {
    const current = formatScopeOfForm(form, activeScope);
    setSelectedScope(allowedScopes.includes(current) ? current : allowedScopes[0] || 'purchase');
  }, [form, activeScope, allowedScopes]);
  const safeScopes: FormatScope[] = allowedScopes.length ? allowedScopes : ['purchase'];
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-content format-scope-modal" onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2>Sửa phạm vi form</h2>
          <button className="modal-close" type="button" onClick={onClose}>×</button>
        </div>
        <div className="modal-body format-scope-modal-body">
          <div className="format-scope-form-name">
            <strong>{form.label || form.sheet || form.id}</strong>
            <span className="muted">{form.template_original_name || form.sheet || form.builtin_exporter || 'Form tùy chỉnh'}</span>
          </div>
          <div className="format-scope-choice-list">
            {safeScopes.map((scope) => (
              <label key={scope} className="format-scope-choice">
                <input type="radio" name={`form-scope-${form.id}`} value={scope} checked={selectedScope === scope} disabled={busy} onChange={() => setSelectedScope(scope)} />
                <span>{FORMAT_SCOPE_LABELS[scope]}</span>
              </label>
            ))}
          </div>
        </div>
        <div className="modal-footer">
          <button type="button" className="btn-secondary" onClick={onClose}>Hủy</button>
          <button type="button" disabled={busy || !safeScopes.includes(selectedScope)} onClick={() => onSave(selectedScope)}>Lưu phạm vi</button>
        </div>
      </div>
    </div>
  );
}

function FormatGroupModal({ scope, groups, busy, onAddGroup, onGroupChange, onDeleteGroup, onClose }: { scope: FormatScope; groups: ProcessingGroup[]; busy: boolean; onAddGroup: (scope: FormatScope, label?: string) => void; onGroupChange: (scope: FormatScope, groupId: string, update: Partial<ProcessingGroup>) => void; onDeleteGroup: (scope: FormatScope, groupId: string) => void; onClose: () => void }) {
  const [newGroupLabel, setNewGroupLabel] = useState('');
  const addGroup = () => {
    const label = newGroupLabel.trim();
    onAddGroup(scope, label || undefined);
    setNewGroupLabel('');
  };
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-content format-group-modal" onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <h2>Quản lý nhóm công ty</h2>
          <button className="modal-close" type="button" onClick={onClose}>×</button>
        </div>
        <div className="modal-body format-group-modal-body">
          <div className="format-group-add-row">
            <input value={newGroupLabel} placeholder="Tên nhóm mới" disabled={busy} onChange={(event) => setNewGroupLabel(event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter') addGroup(); }} />
            <button type="button" disabled={busy} onClick={addGroup}>Thêm nhóm</button>
          </div>
      <div className="format-group-list">
        {groups.map((group) => (
          <div className="format-group-row" key={group.id}>
            <input value={group.label} disabled={busy} onChange={(event) => onGroupChange(scope, group.id, { label: event.currentTarget.value })} />
            <span>{group.id}</span>
            <button type="button" className="btn-secondary compact-table-button" disabled={busy || group.builtin} onClick={() => onDeleteGroup(scope, group.id)}>Xóa</button>
          </div>
        ))}
      </div>
        </div>
        <div className="modal-footer">
          <button type="button" onClick={onClose}>Đóng</button>
        </div>
      </div>
    </div>
  );
}

function TransformRuleEditor({ rules, sourceColumns, onChange }: { rules: MappingTransformRule[]; sourceColumns: FormColumn[]; onChange: (rules: MappingTransformRule[]) => void }) {
  const updateRule = (index: number, update: Partial<MappingTransformRule>) => {
    onChange(rules.map((rule, ruleIndex) => (ruleIndex === index ? { ...rule, ...update } : rule)));
  };
  const addRule = () => onChange([...rules, { source_col: '', match_type: 'starts_with', value: '', result: '' }]);
  const removeRule = (index: number) => onChange(rules.filter((_, ruleIndex) => ruleIndex !== index));
  return (
    <div className="transform-rule-editor">
      <div className="transform-rule-editor-header">
        <span>IF</span>
        <button type="button" className="btn-secondary compact-table-button" onClick={addRule}>Thêm</button>
      </div>
      {rules.map((rule, index) => {
        const matchType = TRANSFORM_RULE_MATCH_OPTIONS.some((option) => option.value === rule.match_type) ? String(rule.match_type) : 'starts_with';
        const needsValue = TRANSFORM_RULE_MATCH_OPTIONS.find((option) => option.value === matchType)?.needsValue !== false;
        return (
          <div className="transform-rule-row" key={index}>
            <select value={rule.source_col || ''} onChange={(event) => updateRule(index, { source_col: event.currentTarget.value })}>
              <option value="">Cột điều kiện</option>
              {sourceColumns.map((column) => <option key={column.letter} value={column.letter}>{column.label}</option>)}
            </select>
            <select value={matchType} onChange={(event) => updateRule(index, { match_type: event.currentTarget.value, value: TRANSFORM_RULE_MATCH_OPTIONS.find((option) => option.value === event.currentTarget.value)?.needsValue === false ? '' : rule.value || '' })}>
              {TRANSFORM_RULE_MATCH_OPTIONS.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
            </select>
            <input value={rule.value || ''} disabled={!needsValue} placeholder={needsValue ? 'Giá trị so sánh' : '-'} onChange={(event) => updateRule(index, { value: event.currentTarget.value })} />
            <input value={rule.result || ''} placeholder="Giá trị gán cho output" onChange={(event) => updateRule(index, { result: event.currentTarget.value })} />
            <button type="button" className="btn-secondary compact-table-button" onClick={() => removeRule(index)}>Xóa</button>
          </div>
        );
      })}
    </div>
  );
}

function FormatSamplePanel({ title, summary, columns, preview }: { title: string; summary?: UploadSummary | null; columns: FormColumn[]; preview?: Array<Record<string, string>> }) {
  const rows = preview || summary?.preview || [];
  return (
    <div className="format-sample-panel">
      <strong>{title}</strong>
      {summary?.original_name && <span>{summary.original_name}</span>}
      <div className="format-chip-list">{columns.slice(0, 18).map((column) => <span key={column.letter}>{column.label}</span>)}</div>
      {rows.length > 0 && <SmallPreviewTable rows={rows.slice(0, 3)} />}
    </div>
  );
}

function SmallPreviewTable({ rows }: { rows: Array<Record<string, string>> }) {
  const columns = Object.keys(rows[0] || {}).slice(0, 8);
  if (!rows.length || !columns.length) return null;
  return (
    <div className="small-preview-table-wrap">
      <table className="small-preview-table">
        <thead><tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr></thead>
        <tbody>{rows.map((row, index) => <tr key={index}>{columns.map((column) => <td key={column}>{row[column]}</td>)}</tr>)}</tbody>
      </table>
    </div>
  );
}

function mappingExpression(rule: NonNullable<FormMappingPreset['mappings']>[number], sourceColumns: FormColumn[], outputColumns: FormColumn[]) {
  const source = sourceColumns.find((column) => column.letter === rule.source_col);
  const firstRuleSourceCol = rule.transform_rules?.find((item) => item.source_col)?.source_col;
  const conditionSource = sourceColumns.find((column) => column.letter === (rule.condition_source_col || rule.fallback_source_col || firstRuleSourceCol));
  const target = outputColumns.find((column) => column.letter === rule.target_col);
  const left = isTextMappingRule(rule)
    ? (rule.value || 'Giá trị cố định')
    : source?.label || 'Chọn cột input';
  const suffix = mappingRuleControlValue(rule) === 'if_rules'
    ? ` | IF ${conditionSource?.label || 'chọn cột'}`
    : '';
  return `${left}${suffix} -> ${target?.label || 'Chọn output'}`;
}

function CompanyRulesStage({ companies, selectedCompanyIndex, processingGroups = defaultProcessingGroups('purchase'), productPreviewCodes, productCodeOverrides, productCodeReplacements, wordRules, repeatedPhrases, inventoryPairs, useDefaultInventoryPair, defaultInventoryPairId, inventoryPairRules, busy, showCompanyPrefixControls = false, includeCompanyPrefix = false, prefixStrategy = 'last_2_words', prefixMstDigits = 3, prefixNameWords = 2, prefixNameChars = 1, missingMstPrefixStrategy = 'all_name_words', missingMstCompanies = [], onIncludeCompanyPrefixChange, onCompanyPrefixChange, onPrefixMstDigitsChange, onPrefixNameWordsChange, onPrefixNameCharsChange, onMissingMstPrefixStrategyChange, onApplyPrefixPresetToAll, onCompanySelect, onCompanyChange, onCompanyGroupChange, onBulkCompanyChange, onProductChange, onProductCodeChange, onApplyChoices, onRefreshPreviews, onWordRuleChange, onAddWordRule, onRepeatedChange, onAddRepeated, onRemoveRepeated, onProductCodeReplacementChange, onAddProductCodeReplacement, onRemoveProductCodeReplacement, onProductCodeReplacementsCommit, onApplyRelatedProductCodes, onImportProductCodeReplacements, onAddInventoryPair, onInventoryPairChange, onRemoveInventoryPair, onInventoryDefaultsChange, onAddInventoryRule, onInventoryRuleChange, onRemoveInventoryRule, sonPhuongAllocationConfig, onSonPhuongAllocationConfigChange }: { companies: CompanyRow[]; selectedCompanyIndex: number; processingGroups?: ProcessingGroup[]; productPreviewCodes: Record<string, string>; productCodeOverrides: Record<string, string>; productCodeReplacements: Record<string, string>; wordRules: Record<string, string>; repeatedPhrases: string[]; inventoryPairs: InventoryPair[]; useDefaultInventoryPair: boolean; defaultInventoryPairId: string; inventoryPairRules: InventoryRule[]; busy: boolean; showCompanyPrefixControls?: boolean; includeCompanyPrefix?: boolean; prefixStrategy?: string; prefixMstDigits?: number; prefixNameWords?: number; prefixNameChars?: number; missingMstPrefixStrategy?: PrefixPresetStrategy; missingMstCompanies?: MissingMstCompanyWarning[]; onIncludeCompanyPrefixChange?: (include: boolean) => void; onCompanyPrefixChange?: (index: number, value: string) => void; onPrefixMstDigitsChange?: (digits: number) => void; onPrefixNameWordsChange?: (words: number) => void; onPrefixNameCharsChange?: (chars: number) => void; onMissingMstPrefixStrategyChange?: (strategy: PrefixPresetStrategy) => void; onApplyPrefixPresetToAll?: (strategy: PrefixPresetStrategy) => void; onCompanySelect: (index: number) => void; onCompanyChange: (index: number, pending: boolean) => void; onCompanyGroupChange?: (index: number, groupId: string) => void; onBulkCompanyChange?: (pending: boolean) => void; onProductChange: (companyIndex: number, productName: string, selected: boolean) => void; onProductCodeChange: (companyIndex: number, productName: string, code: string) => void; onApplyChoices: () => void; onRefreshPreviews: () => void; onWordRuleChange: (index: number, field: 'from' | 'to', value: string) => void; onAddWordRule: () => void; onRepeatedChange: (index: number, value: string) => void; onAddRepeated: () => void; onRemoveRepeated: (index: number) => void; onProductCodeReplacementChange: (index: number, field: 'from' | 'to', value: string) => void; onAddProductCodeReplacement: () => void; onRemoveProductCodeReplacement: (index: number) => void; onProductCodeReplacementsCommit: (replacements: Record<string, string>) => void; onApplyRelatedProductCodes: (updates: RelatedProductCodeUpdate[]) => void; onImportProductCodeReplacements: (file: File | undefined) => void; onAddInventoryPair: () => void; onInventoryPairChange: (index: number, field: 'ma_kho' | 'tk_vat_tu', value: string) => void; onRemoveInventoryPair: (index: number) => void; onInventoryDefaultsChange: (update: Partial<Pick<WorkflowState, 'useDefaultInventoryPair' | 'defaultInventoryPairId'>>) => void; onAddInventoryRule: () => void; onInventoryRuleChange: (index: number, update: Partial<InventoryRule>) => void; onRemoveInventoryRule: (index: number) => void; sonPhuongAllocationConfig?: InventoryAllocationConfig; onSonPhuongAllocationConfigChange?: (config: InventoryAllocationConfig) => void }) {
  const [showConfigModal, setShowConfigModal] = useState(false);
  const [collapsedSections, setCollapsedSections] = useState({ wordRules: true, repeatedPhrases: true });
  const wordEntries = Object.entries(wordRules);
  const safeSelectedIndex = selectedCompanyIndex >= 0 && selectedCompanyIndex < companies.length ? selectedCompanyIndex : -1;
  const selectedCompany = safeSelectedIndex >= 0 ? companies[safeSelectedIndex] : undefined;
  const selectedProducts = new Set(selectedCompany?.selected_product_names.length ? selectedCompany.selected_product_names : selectedCompany?.all_products.map((product) => product.name));
  const productRows = selectedCompany?.all_products ?? [];
  const lineSequenceProductCodes = processingGroups.some((group) => group.id === 'payment_voucher');
  const longProducts = lineSequenceProductCodes || !selectedCompany ? [] : productRows.filter((product) => selectedProducts.has(product.name) && productDisplayCode(selectedCompany, product.name, productPreviewCodes, productCodeOverrides, productCodeReplacements, includeCompanyPrefix).length > maxCodeLength);
  const normalProducts = selectedCompany ? productRows.filter((product) => !longProducts.includes(product)) : [];
  const companyCodeLong = !lineSequenceProductCodes && Boolean(selectedCompany?.value && selectedCompany.value.length > maxCodeLength);
  const normalizedGroups = normalizeProcessingGroups(processingGroups, 'purchase');
  const groupIds = new Set(normalizedGroups.map((group) => group.id));
  const duplicatePrefixSet = new Set<string>();
  const prefixCounts = new Map<string, number>();
  companies.forEach(c => {
    const normalizedPrefix = committedCompanyPrefix(c);
    const groupId = displayCompanyGroupId(c, groupIds, 'pending');
    if (groupId !== IGNORED_GROUP_ID && normalizedPrefix) {
      const key = duplicatePrefixKey(groupId, normalizedPrefix);
      prefixCounts.set(key, (prefixCounts.get(key) || 0) + 1);
    }
  });
  Array.from(prefixCounts.entries())
    .filter(([, count]) => count > 1)
    .forEach(([key]) => duplicatePrefixSet.add(key));
  const groupLabelById = new Map(normalizedGroups.map((group) => [group.id, group.label || group.id]));
  const duplicatePrefixWarnings = Array.from(prefixCounts.entries())
    .filter(([, count]) => count > 1)
    .map(([key, count]) => {
      const [groupId, prefix] = duplicatePrefixParts(key);
      return { key, groupId, groupLabel: groupLabelById.get(groupId) || groupId, prefix, count };
    })
    .sort((left, right) => left.groupLabel.localeCompare(right.groupLabel, 'vi', { numeric: true, sensitivity: 'base' }) || left.prefix.localeCompare(right.prefix, 'vi', { numeric: true, sensitivity: 'base' }));
  const groups = companyDisplayGroups(companies, normalizedGroups, duplicatePrefixSet);
  const activePrefixStrategy = normalizedPrefixStrategy(prefixStrategy);
  const activeMissingMstPrefixStrategy = normalizeMissingMstPrefixStrategy(missingMstPrefixStrategy);
  const hasMissingMstCompanies = companies.some(isMissingMstCompany);
  const allProductNames = selectedCompany?.all_products.map(p => p.name) || [];
  const allSelected = allProductNames.length > 0 && allProductNames.every(name => selectedProducts.has(name));
  
  const renderProductRow = (product: { name: string; count?: number }, forceWarning = false) => {
    if (!selectedCompany) return null;
    const code = lineSequenceProductCodes ? `VT.${committedCompanyPrefix(selectedCompany) || '<PREFIX>'}.<SỐ HĐ>.<STT>` : productDisplayCode(selectedCompany, product.name, productPreviewCodes, productCodeOverrides, productCodeReplacements, includeCompanyPrefix);
    const selected = selectedProducts.has(product.name);
    const longCode = selected && code.length > maxCodeLength;
    return <tr key={product.name} className={longCode || forceWarning ? 'danger-row big-select-row' : 'big-select-row'}><td><input type="checkbox" checked={selected} onChange={(event) => onProductChange(safeSelectedIndex, product.name, event.currentTarget.checked)} /></td><td className="product-name-cell">{product.name}</td><td>{product.count ?? ''}</td><td>{lineSequenceProductCodes ? <code>{code}</code> : <input className="code-edit" value={code} onChange={(event) => onProductCodeChange(safeSelectedIndex, product.name, event.currentTarget.value)} />}</td></tr>;
  };

  return (
    <div className="company-workspace">
      <div className={`compact-rules-panel ${showCompanyPrefixControls ? 'prefix-enabled' : ''}`}>
        {!lineSequenceProductCodes && <button type="button" className="btn-secondary" onClick={() => setShowConfigModal(true)}>Cấu hình nâng cao</button>}
        {showCompanyPrefixControls && <section className="company-prefix-card compact-rule-card">
          <label className="inline-check"><input type="checkbox" checked={includeCompanyPrefix} onChange={(event) => onIncludeCompanyPrefixChange?.(event.currentTarget.checked)} /> Dùng prefix công ty</label>
          {includeCompanyPrefix && <>
            <div className="prefix-strategy-row">
              <label>Số ký tự MST:</label>
              <input type="number" min={1} max={10} value={prefixMstDigits} onChange={(event) => onPrefixMstDigitsChange?.(parseInt(event.currentTarget.value) || 3)} />
            </div>
            <div className="prefix-strategy-row">
              <label>Số từ lấy chữ đầu:</label>
              <input type="number" min={1} max={20} value={prefixNameWords} onChange={(event) => onPrefixNameWordsChange?.(parseInt(event.currentTarget.value) || 2)} />
            </div>
            <div className="prefix-strategy-row">
              <label>Số ký tự mỗi từ:</label>
              <input type="number" min={1} max={10} value={prefixNameChars} onChange={(event) => onPrefixNameCharsChange?.(parseInt(event.currentTarget.value) || 1)} />
            </div>
            <div className="prefix-quick-actions">
              <button type="button" className={`prefix-apply-all-button ${activePrefixStrategy === 'last_2_words' ? 'active' : ''}`} disabled={busy || !companies.length} onClick={() => onApplyPrefixPresetToAll?.('last_2_words')}>Áp {prefixNameWords} từ</button>
              <button type="button" className={`prefix-apply-all-button ${activePrefixStrategy === 'all_name_words' ? 'active' : ''}`} disabled={busy || !companies.length} onClick={() => onApplyPrefixPresetToAll?.('all_name_words')}>Áp tất cả từ đầu</button>
              <button type="button" className={`prefix-apply-all-button ${activePrefixStrategy === 'last_3_mst' ? 'active' : ''}`} disabled={busy || !companies.length} onClick={() => onApplyPrefixPresetToAll?.('last_3_mst')}>Áp MST</button>
              <button type="button" className={`prefix-apply-all-button ${activePrefixStrategy === '2_words_mst' ? 'active' : ''}`} disabled={busy || !companies.length} onClick={() => onApplyPrefixPresetToAll?.('2_words_mst')}>Áp {prefixNameWords} từ + MST</button>
            </div>
            {hasMissingMstCompanies && <div className="prefix-missing-mst-rule-row">
              <span>Công ty không MST:</span>
              <button type="button" className={`prefix-apply-all-button ${activeMissingMstPrefixStrategy === 'last_2_words' ? 'active' : ''}`} disabled={busy} onClick={() => onMissingMstPrefixStrategyChange?.('last_2_words')}>Áp {prefixNameWords} từ</button>
              <button type="button" className={`prefix-apply-all-button ${activeMissingMstPrefixStrategy === 'all_name_words' ? 'active' : ''}`} disabled={busy} onClick={() => onMissingMstPrefixStrategyChange?.('all_name_words')}>Áp tất cả từ đầu</button>
            </div>}
          </>}
        </section>}
      </div>

      <ConfigModal
        isOpen={!lineSequenceProductCodes && showConfigModal}
        onClose={() => setShowConfigModal(false)}
        companies={companies}
        productPreviewCodes={productPreviewCodes}
        productCodeOverrides={productCodeOverrides}
        includeCompanyPrefix={includeCompanyPrefix}
        wordRules={wordRules}
        repeatedPhrases={repeatedPhrases}
        productCodeReplacements={productCodeReplacements}
        inventoryPairs={inventoryPairs}
        useDefaultInventoryPair={useDefaultInventoryPair}
        defaultInventoryPairId={defaultInventoryPairId}
        inventoryPairRules={inventoryPairRules}
        sonPhuongAllocationConfig={sonPhuongAllocationConfig}
        onSonPhuongAllocationConfigChange={onSonPhuongAllocationConfigChange}
        onWordRuleChange={onWordRuleChange}
        onAddWordRule={onAddWordRule}
        onRepeatedChange={onRepeatedChange}
        onAddRepeated={onAddRepeated}
        onRemoveRepeated={onRemoveRepeated}
        onProductCodeReplacementChange={onProductCodeReplacementChange}
        onAddProductCodeReplacement={onAddProductCodeReplacement}
        onRemoveProductCodeReplacement={onRemoveProductCodeReplacement}
        onProductCodeReplacementsCommit={onProductCodeReplacementsCommit}
        onApplyRelatedProductCodes={onApplyRelatedProductCodes}
        onImportProductCodeReplacements={onImportProductCodeReplacements}
        onAddInventoryPair={onAddInventoryPair}
        onInventoryPairChange={onInventoryPairChange}
        onRemoveInventoryPair={onRemoveInventoryPair}
        onInventoryDefaultsChange={onInventoryDefaultsChange}
        onAddInventoryRule={onAddInventoryRule}
        onInventoryRuleChange={onInventoryRuleChange}
        onRemoveInventoryRule={onRemoveInventoryRule}
        onRefreshPreviews={onRefreshPreviews}
        busy={busy}
      />

      <div className="company-product-grid">
        <div className="list-stage company-list-card">
          <div className="stage-toolbar">
            <p>Danh sách công ty</p>
            <div className="company-list-actions">
              <button type="button" className="btn-secondary" disabled={busy || !companies.length} onClick={() => onBulkCompanyChange?.(true)}>Chọn tất cả</button>
              <button type="button" className="btn-secondary" disabled={busy || !companies.length} onClick={() => onBulkCompanyChange?.(false)}>Bỏ chọn tất cả</button>
              <button type="button" disabled={busy || !companies.length} onClick={() => onApplyChoices()}>Áp dụng lựa chọn công ty và hàng hóa</button>
            </div>
          </div>
          {duplicatePrefixWarnings.length > 0 && (
            <div className="duplicate-prefix-warning">
              <p className="warning-text">Cảnh báo: Prefix trùng lặp trong cùng nhóm</p>
              <ul>
                {duplicatePrefixWarnings.map(({ key, groupId, groupLabel, prefix, count }) => {
                  const dupCompanies = companies.filter((c) => displayCompanyGroupId(c, groupIds, 'pending') === groupId && committedCompanyPrefix(c) === prefix);
                  return (
                    <li key={key}>
                      <strong>{groupLabel} / {prefix}</strong> ({count} công ty): {dupCompanies.map((c) => c.company).join(', ')}
                    </li>
                  );
                })}
              </ul>
            </div>
          )}
          {!companies.length ? <p className="muted">Đang chờ dữ liệu công ty. Danh sách sẽ tự tải khi vào stage này sau khi đã chọn file mua vào.</p> : <div className="inner-scroll company-table-scroll"><table className={`company-table grouped-company-table ${showCompanyPrefixControls ? 'has-prefix-column' : ''}`}><thead><tr><th>Nhóm</th><th>Công ty</th><th>MST</th>{showCompanyPrefixControls && <th>Prefix</th>}<th>Số hàng</th></tr></thead><tbody>{groups.map((group) => <CompanyGroupRows key={group.title} group={group} safeSelectedIndex={safeSelectedIndex} showPrefix={showCompanyPrefixControls} processingGroups={processingGroups} onCompanySelect={onCompanySelect} onCompanyChange={onCompanyChange} onCompanyGroupChange={onCompanyGroupChange} onCompanyPrefixChange={onCompanyPrefixChange} />)}</tbody></table><div className="scroll-bottom-spacer" aria-hidden="true" /></div>}
        </div>

        <div className="list-stage product-list-card">
          <div className="stage-toolbar"><p>{selectedCompany ? `Hàng hóa của ${selectedCompany.company} - MST ${companyDisplayMst(selectedCompany)}` : 'Hàng hóa / mã VT preview'}</p></div>
          {!selectedCompany ? <p className="muted">Chọn một dòng công ty để xem danh sách hàng hóa.</p> : <>
            {(companyCodeLong || longProducts.length > 0) && <div className="long-code-warning"><p className="warning-text compact-warning">Cảnh báo: mã vượt {maxCodeLength} ký tự sẽ bị cắt đuôi khi xuất file.</p><table className="product-table warning-code-table"><thead><tr><th>Loại</th><th>Tên</th><th colSpan={2}>Mã đang vượt giới hạn</th></tr></thead><tbody>{companyCodeLong && <tr className="danger-row"><td>Công ty</td><td>{selectedCompany.company}</td><td colSpan={2}>{selectedCompany.value}</td></tr>}{longProducts.map((product) => renderProductRow(product, true))}</tbody></table></div>}
            <div className="inner-scroll product-table-scroll"><table className="product-table"><thead><tr><th>Xử lý</th><th>Tên hàng hóa</th><th>Dòng</th><th>{lineSequenceProductCodes ? 'Quy tắc Mã VT theo hóa đơn' : 'Mã VT xem trước / sửa tay'}</th></tr></thead><tbody>{normalProducts.map((product) => renderProductRow(product))}</tbody></table><div className="scroll-bottom-spacer" aria-hidden="true" /></div>
          </>}
        </div>
      </div>
    </div>
  );
}

function CompanyGroupRows({ group, safeSelectedIndex, showPrefix = false, processingGroups = defaultProcessingGroups('purchase'), onCompanySelect, onCompanyChange, onCompanyGroupChange, onCompanyPrefixChange }: { group: CompanyDisplayGroup; safeSelectedIndex: number; showPrefix?: boolean; processingGroups?: ProcessingGroup[]; onCompanySelect: (index: number) => void; onCompanyChange: (index: number, pending: boolean) => void; onCompanyGroupChange?: (index: number, groupId: string) => void; onCompanyPrefixChange?: (index: number, value: string) => void }) {
  if (!group.rows.length) return null;
  return <>{<tr className={`company-section-row ${group.className}`}><td colSpan={showPrefix ? 5 : 4}>{group.title} ({group.rows.length})</td></tr>}{group.rows.map(({ company, index }) => {
    const groupId = pendingCompanyGroup(company);
    const selectedCount = company.selected_product_names.length || company.all_products.length;
    const rowClass = group.className === 'duplicate-section' ? 'duplicate-company-row' : group.className === 'missing-mst-section' ? 'missing-mst-company-row' : '';
    return <tr key={companyRowKey(company, index)} className={`big-select-row ${rowClass} ${index === safeSelectedIndex ? 'selected-row' : ''}`} onClick={() => onCompanySelect(index)}><td onClick={(event) => event.stopPropagation()}><select value={groupId} onChange={(event) => { const nextGroupId = event.currentTarget.value; if (onCompanyGroupChange) onCompanyGroupChange(index, nextGroupId); else onCompanyChange(index, nextGroupId === MATERIALS_GROUP_ID); }}>{processingGroups.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}</select></td><td className="company-name-cell">{company.company}</td><td className="company-mst-cell">{companyDisplayMst(company)}</td>{showPrefix && <td className="company-prefix-cell" onClick={(event) => event.stopPropagation()}><input className="company-prefix-input" value={company.value || ''} onChange={(event) => onCompanyPrefixChange?.(index, event.currentTarget.value)} /></td>}<td className="company-count-cell">{selectedCount} / {company.all_products.length}</td></tr>;
  })}</>;
}

function MissingMstCompanyRows({ rows, showPrefix = false, prefixStrategy = 'last_2_words', prefixMstDigits = 3, prefixNameWords = 2, prefixNameChars = 1 }: { rows: MissingMstCompanyWarning[]; showPrefix?: boolean; prefixStrategy?: PrefixPresetStrategy; prefixMstDigits?: number; prefixNameWords?: number; prefixNameChars?: number }) {
  if (!rows.length) return null;
  return <>{<tr className="company-section-row missing-mst-section"><td colSpan={showPrefix ? 5 : 4}>Công ty không có MST - cần kiểm tra ({rows.length})</td></tr>}{rows.map((item, index) => {
    const invoiceNos = item.invoice_nos?.filter(Boolean) || [];
    const prefix = missingMstDisplayPrefix(item.company, prefixStrategy, prefixMstDigits, prefixNameWords, prefixNameChars);
    return <tr key={`${item.company}-${index}`} className="missing-mst-company-row"><td>—</td><td><div className="missing-mst-company-name"><strong>{item.company}</strong>{invoiceNos.length > 0 && <small>Số HĐ: {invoiceNos.slice(0, 8).join(', ')}{invoiceNos.length > 8 ? ', ...' : ''}</small>}</div></td><td>(trống)</td>{showPrefix && <td>{prefix}</td>}<td>{item.count} dòng</td></tr>;
  })}</>;
}

function missingMstDisplayPrefix(companyName: string, strategy: PrefixPresetStrategy, mstDigits: number, nameWords: number, nameChars: number): string {
  const fakeCompany = { company: companyName, mst: '', prefix_strategies: {} } as CompanyRow;
  if (strategy === 'last_3_mst') return 'Không có MST';
  if (strategy === '2_words_mst') {
    const namePrefix = computePresetPrefix(fakeCompany, 'last_2_words', mstDigits, nameWords, nameChars);
    return namePrefix ? `${namePrefix} + thiếu MST` : 'Không có MST';
  }
  return computePresetPrefix(fakeCompany, strategy, mstDigits, nameWords, nameChars) || 'Không có prefix';
}

function InventoryPairEditor({ pairs, useDefault, defaultPairId, rules, busy, onAddPair, onPairChange, onRemovePair, onDefaultsChange, onAddRule, onRuleChange, onRemoveRule, fixed = false }: { pairs: InventoryPair[]; useDefault: boolean; defaultPairId: string; rules: InventoryRule[]; busy: boolean; onAddPair: () => void; onPairChange: (index: number, field: 'ma_kho' | 'tk_vat_tu', value: string) => void; onRemovePair: (index: number) => void; onDefaultsChange: (update: Partial<Pick<WorkflowState, 'useDefaultInventoryPair' | 'defaultInventoryPairId'>>) => void; onAddRule: () => void; onRuleChange: (index: number, update: Partial<InventoryRule>) => void; onRemoveRule: (index: number) => void; fixed?: boolean }) {
  const [collapsed, setCollapsed] = useState(false);
  const overlapWarnings = inventoryRuleOverlapWarnings(rules);

  if (fixed) {
    const pair = pairs[0] || { id: 'son-phuong-purchase-materials', ma_kho: 'KHHVT', tk_vat_tu: '156' };
    const restore = () => {
      onPairChange(0, 'ma_kho', 'KHHVT');
      onPairChange(0, 'tk_vat_tu', '156');
      onDefaultsChange({ useDefaultInventoryPair: true, defaultInventoryPairId: pair.id });
    };
    return (
      <section className="inventory-editor">
        <div className="inventory-editor-header">
          <div>
            <strong>{'C\u1eb7p kho mua v\u00e0o S\u01a1n Ph\u01b0\u01a1ng'}</strong>
            <p>{'Ch\u1ec9 Nh\u00f3m v\u1eadt t\u01b0 mua v\u00e0o d\u00f9ng c\u1eb7p n\u00e0y. Nh\u00f3m d\u1ecbch v\u1ee5 kh\u00f4ng tham gia Ph\u00e2n kho.'}</p>
          </div>
          <button type="button" className="btn-secondary" disabled={busy} onClick={restore}>{'Kh\u00f4i ph\u1ee5c KHHVT / 156'}</button>
        </div>
        <div className="inventory-card">
          <div className="inventory-table-scroll">
            <table className="inventory-table">
              <thead><tr><th>{'Ph\u1ea1m vi'}</th><th>{'M\u00e3 kho'}</th><th>{'TK v\u1eadt t\u01b0'}</th></tr></thead>
              <tbody><tr>
                <td><strong>{'Nh\u00f3m v\u1eadt t\u01b0'}</strong></td>
                <td><input value={pair.ma_kho} disabled={busy} onChange={(event) => onPairChange(0, 'ma_kho', event.currentTarget.value.toUpperCase())} /></td>
                <td><input value={pair.tk_vat_tu} disabled={busy} onChange={(event) => onPairChange(0, 'tk_vat_tu', event.currentTarget.value)} /></td>
              </tr></tbody>
            </table>
          </div>
          {(!pair.ma_kho.trim() || !pair.tk_vat_tu.trim()) && <p className="warning-text">{'Ph\u1ea3i nh\u1eadp \u0111\u1ee7 M\u00e3 kho v\u00e0 TK v\u1eadt t\u01b0.'}</p>}
        </div>
      </section>
    );
  }
  return (
    <section className={`inventory-editor ${collapsed ? 'collapsed' : ''}`}>
      <div className="inventory-editor-header">
        <div>
          <strong>Cặp Mã kho / TK vật tư</strong>
          <p>Cấu hình dùng khi xuất file cho công ty/hàng hóa đang xử lý.</p>
        </div>
        <label className="inline-check"><input type="checkbox" checked={useDefault} disabled={busy || !pairs.length} onChange={(event) => onDefaultsChange({ useDefaultInventoryPair: event.currentTarget.checked })} /> Dùng cặp mặc định</label>
        <label className="inventory-default-select"><span>Cặp mặc định</span><select value={defaultPairId} disabled={busy || !pairs.length} onChange={(event) => onDefaultsChange({ defaultInventoryPairId: event.currentTarget.value })}><option value="">Chọn cặp</option>{pairs.map((pair) => <option key={pair.id} value={pair.id}>{pairLabel(pair)}</option>)}</select></label>
        <button type="button" className="btn-secondary compact-table-button" onClick={() => setCollapsed(!collapsed)}>{collapsed ? 'Mở rộng' : 'Thu gọn'}</button>
      </div>
      {!collapsed && <div className="inventory-editor-grid">
        <div className="inventory-card">
          <div className="inventory-card-title"><span>Danh sách cặp</span><button type="button" className="btn-secondary" disabled={busy} onClick={onAddPair}>Thêm cặp</button></div>
          <div className="inventory-table-scroll"><table className="inventory-table"><thead><tr><th>Mã kho</th><th>TK vật tư</th><th></th></tr></thead><tbody>{pairs.length ? pairs.map((pair, index) => <tr key={pair.id}><td><input value={pair.ma_kho} onChange={(event) => onPairChange(index, 'ma_kho', event.currentTarget.value)} /></td><td><input value={pair.tk_vat_tu} onChange={(event) => onPairChange(index, 'tk_vat_tu', event.currentTarget.value)} /></td><td><button type="button" className="btn-secondary compact-table-button" disabled={busy} onClick={() => onRemovePair(index)}>Xóa</button></td></tr>) : <tr><td colSpan={3} className="muted">Chưa có cặp. Bấm Thêm cặp để nhập Mã kho và TK vật tư.</td></tr>}</tbody></table></div>
        </div>
        <div className="inventory-card">
          <div className="inventory-card-title"><span>Quy tắc gán cặp</span><button type="button" className="btn-secondary" disabled={busy || !pairs.length} onClick={onAddRule}>Thêm quy tắc</button></div>
          {overlapWarnings.length > 0 && <div className="inventory-rule-warning"><strong>Cảnh báo quy tắc chồng chữ</strong>{overlapWarnings.map((warning) => <p key={warning}>{warning}</p>)}</div>}
          <div className="inventory-table-scroll"><table className="inventory-table inventory-rule-table"><thead><tr><th>Bật</th><th>Ưu tiên</th><th>Cột nguồn</th><th>So sánh</th><th>Giá trị</th><th>Cặp gán</th><th></th></tr></thead><tbody>{rules.length ? rules.map((rule, index) => <tr key={`${rule.source_col}-${index}`}><td><input type="checkbox" checked={rule.enabled !== false} onChange={(event) => onRuleChange(index, { enabled: event.currentTarget.checked })} /></td><td><input className="priority-input" type="number" value={rule.priority ?? 0} onChange={(event) => onRuleChange(index, { priority: Number(event.currentTarget.value) || 0 })} /></td><td><input value={rule.source_col} onChange={(event) => onRuleChange(index, { source_col: event.currentTarget.value.toUpperCase() })} /></td><td><select value={rule.operator || 'contains'} onChange={(event) => onRuleChange(index, { operator: event.currentTarget.value })}><option value="contains">Chứa</option><option value="equals">Bằng</option></select></td><td><input value={rule.value} onChange={(event) => onRuleChange(index, { value: event.currentTarget.value })} /></td><td><select value={rule.pair_id} onChange={(event) => onRuleChange(index, { pair_id: event.currentTarget.value })}><option value="">Chọn cặp</option>{pairs.map((pair) => <option key={pair.id} value={pair.id}>{pairLabel(pair)}</option>)}</select></td><td><button type="button" className="btn-secondary compact-table-button" disabled={busy} onClick={() => onRemoveRule(index)}>Xóa</button></td></tr>) : <tr><td colSpan={7} className="muted">Không có quy tắc. Nếu không bật mặc định, dòng không khớp sẽ để trống Mã kho/TK vật tư.</td></tr>}</tbody></table></div>
        </div>
      </div>}
    </section>
  );
}

function pairLabel(pair: InventoryPair) {
  return `${pair.ma_kho || 'Mã kho?'} / ${pair.tk_vat_tu || 'TK?'}`;
}

function normalizedInventoryRuleValue(value: string) {
  return simpleMatchText(value);
}

function inventoryRulePriorityValue(rule: InventoryRule) {
  return Number.isFinite(Number(rule.priority)) ? Number(rule.priority) : 0;
}

function inventoryRuleOverlapWarnings(rules: InventoryRule[]) {
  const warnings: string[] = [];
  const activeRules = rules
    .map((rule, index) => ({ rule, index, value: normalizedInventoryRuleValue(rule.value) }))
    .filter((item) => item.rule.enabled !== false && item.rule.operator !== 'equals' && item.rule.source_col.trim() && item.value);
  for (let i = 0; i < activeRules.length; i += 1) {
    for (let j = i + 1; j < activeRules.length; j += 1) {
      const left = activeRules[i];
      const right = activeRules[j];
      if (left.rule.source_col.trim().toUpperCase() !== right.rule.source_col.trim().toUpperCase()) continue;
      if (left.value !== right.value && !left.value.includes(right.value) && !right.value.includes(left.value)) continue;
      const leftPriority = inventoryRulePriorityValue(left.rule);
      const rightPriority = inventoryRulePriorityValue(right.rule);
      const winner = leftPriority === rightPriority ? `rule #${Math.min(left.index, right.index) + 1} do đứng trước` : `rule ưu tiên ${Math.max(leftPriority, rightPriority)}`;
      warnings.push(`Rule #${left.index + 1} "${left.rule.value}" chồng với rule #${right.index + 1} "${right.rule.value}" trên cột ${left.rule.source_col}. Dòng khớp cả hai sẽ dùng ${winner}.`);
    }
  }
  return warnings.slice(0, 5);
}

function emptyPrefixStrategyValues(): PrefixStrategyValues {
  return {
    last_2_words: {},
    last_3_mst: {},
    '2_words_mst': {},
    all_name_words: {},
  };
}

function companyRowKey(company: CompanyRow, fallbackIndex = 0): string {
  return company.safe_id || `${companyConfigKey(company) || 'company'}-${fallbackIndex}`;
}

function companyConfigKey(company: CompanyRow): string {
  return company.company_id || company.mst || company.safe_id || company.company;
}

function isMissingMstCompany(company: CompanyRow): boolean {
  return company.missing_mst === true || !String(company.mst || '').trim();
}

function companyDisplayMst(company: CompanyRow): string {
  return String(company.mst || '').trim() || '(trống)';
}

function companyReviewIdentityKey(company: CompanyRow): string {
  return String(company.mst || '').trim() || String(company.company || '').trim();
}

function prefixMemoryKey(company: CompanyRow): string {
  return companyConfigKey(company);
}

function clampPrefixMstDigits(value: unknown): number {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return 3;
  return Math.max(1, Math.min(10, Math.trunc(numberValue)));
}

function clampPrefixNameWords(value: unknown): number {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return 2;
  return Math.max(1, Math.min(20, Math.trunc(numberValue)));
}

function clampPrefixNameChars(value: unknown): number {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return 1;
  return Math.max(1, Math.min(10, Math.trunc(numberValue)));
}

function normalizePrefixStrategyValues(raw: Record<string, Record<string, string>> | undefined, fallback?: PrefixStrategyValues): PrefixStrategyValues {
  const next = emptyPrefixStrategyValues();
  const source = raw && typeof raw === 'object' ? raw : fallback;
  (Object.keys(next) as PrefixPresetStrategy[]).forEach((strategy) => {
    const values = source?.[strategy];
    if (!values || typeof values !== 'object') return;
    next[strategy] = Object.fromEntries(
      Object.entries(values)
        .filter(([key, value]) => key && value != null)
        .map(([key, value]) => [key, normalizePrefixValue(value)]),
    );
  });
  return next;
}

function normalizePrefixValue(value: unknown): string {
  return String(value ?? '').trim().toUpperCase();
}

function rememberPrefixEdit(values: PrefixStrategyValues, strategy: PrefixPresetStrategy, company: CompanyRow, value: string, mstDigits: number, nameWords: number, nameChars: number, missingMstStrategy?: PrefixPresetStrategy): PrefixStrategyValues {
  const effectiveStrategy = isMissingMstCompany(company) ? normalizeMissingMstPrefixStrategy(missingMstStrategy ?? strategy) : strategy;
  const next = rememberManualPrefixValues(values, effectiveStrategy, [company], mstDigits, nameWords, nameChars, { [prefixMemoryKey(company)]: value });
  return next;
}

function rememberManualPrefixValues(values: PrefixStrategyValues, strategy: PrefixPresetStrategy, rows: CompanyRow[], mstDigits: number, nameWords: number, nameChars: number, overrides: Record<string, string> = {}): PrefixStrategyValues {
  const next = {
    ...emptyPrefixStrategyValues(),
    ...values,
    [strategy]: { ...(values[strategy] ?? {}) },
  };
  rows.forEach((company) => {
    const key = prefixMemoryKey(company);
    const value = normalizePrefixValue(Object.prototype.hasOwnProperty.call(overrides, key) ? overrides[key] : company.value);
    const computed = normalizePrefixValue(computePresetPrefix(company, strategy, mstDigits, nameWords, nameChars));
    if (value === computed) {
      delete next[strategy][key];
    } else {
      next[strategy][key] = value;
    }
  });
  return next;
}

function seedLoadedPrefixValues(values: PrefixStrategyValues, strategy: PrefixPresetStrategy, rows: CompanyRow[], mstDigits: number, nameWords: number, nameChars: number, missingMstStrategy?: PrefixPresetStrategy): PrefixStrategyValues {
  const next = {
    ...emptyPrefixStrategyValues(),
    ...values,
    [strategy]: { ...(values[strategy] ?? {}) },
  };
  const normalizedMissingMstStrategy = normalizeMissingMstPrefixStrategy(missingMstStrategy ?? strategy);
  next[normalizedMissingMstStrategy] = { ...(next[normalizedMissingMstStrategy] ?? {}) };
  rows.forEach((company) => {
    const effectiveStrategy = isMissingMstCompany(company) ? normalizedMissingMstStrategy : strategy;
    const key = prefixMemoryKey(company);
    if (Object.prototype.hasOwnProperty.call(next[effectiveStrategy], key)) return;
    const loaded = normalizePrefixValue(company.value);
    if (!loaded) return;
    const computed = normalizePrefixValue(computePresetPrefix(company, effectiveStrategy, mstDigits, nameWords, nameChars));
    const defaultPrefix = normalizePrefixValue(company.default_prefix || company.prefix_strategies?.last_2_words || computePresetPrefix(company, effectiveStrategy, mstDigits, nameWords, nameChars));
    if (loaded !== computed && loaded !== defaultPrefix) {
      next[effectiveStrategy][key] = loaded;
    }
  });
  return next;
}

function applyPrefixStrategyRows(rows: CompanyRow[], strategy: PrefixPresetStrategy, mstDigits: number, nameWords: number, nameChars: number, values: PrefixStrategyValues, commit = false, missingMstStrategy?: PrefixPresetStrategy): CompanyRow[] {
  const normalizedMissingMstStrategy = normalizeMissingMstPrefixStrategy(missingMstStrategy ?? strategy);
  return rows.map((company) => {
    const effectiveStrategy = isMissingMstCompany(company) ? normalizedMissingMstStrategy : strategy;
    const strategyValues = values[effectiveStrategy] ?? {};
    const key = prefixMemoryKey(company);
    const savedValue = Object.prototype.hasOwnProperty.call(strategyValues, key) ? strategyValues[key] : undefined;
    const value = savedValue ?? computePresetPrefix(company, effectiveStrategy, mstDigits, nameWords, nameChars);
    return { ...company, value, ...(commit ? { committed_prefix: normalizePrefixValue(value) } : {}) };
  });
}

function scrollStageBodyToTop() {
  window.requestAnimationFrame(() => {
    const body = document.querySelector('.stage-body');
    if (body instanceof HTMLElement) body.scrollTo({ top: 0, left: 0 });
  });
}

function computeCustomPrefix(company: CompanyRow, option: { name?: string; formula: string; words: number; chars: number; mstDigits?: number }): string {
  const locationPhrases = [
    ['VIET', 'NAM'], ['HA', 'NOI'], ['HO', 'CHI', 'MINH'], ['DA', 'NANG'], ['HAI', 'PHONG'], ['CAN', 'THO'],
    ['BAC', 'NINH'], ['BAC', 'GIANG'], ['HA', 'NAM'], ['NAM', 'DINH'], ['THAI', 'BINH'], ['THANH', 'HOA'],
    ['NGHE', 'AN'], ['HA', 'TINH'], ['QUANG', 'NINH'], ['QUANG', 'NAM'], ['QUANG', 'NGAI'], ['BINH', 'DUONG'],
    ['DONG', 'NAI'], ['LONG', 'AN'], ['TIEN', 'GIANG'], ['VINH', 'PHUC'], ['PHU', 'THO'], ['HUNG', 'YEN'],
  ];
  const rawWords = company.company.split(/\s+/).filter((word) => word.length > 0);
  const words = rawWords.map((word) => sanitizeDisplayProductCode(word).toUpperCase()).filter(Boolean);
  const significant: string[] = [];
  let index = 0;
  while (index < words.length) {
    const phrase = locationPhrases.find((items) => items.every((part, offset) => words[index + offset] === part));
    if (phrase) {
      index += phrase.length;
    } else {
      significant.push(words[index]);
      index += 1;
    }
  }
  const selectedWords = option.formula === 'all_initials' ? significant : significant.slice(-Math.max(1, option.words));
  const charCount = Math.max(1, option.chars);
  const initials = selectedWords.map((word) => word.slice(0, charCount).toUpperCase()).join('');
  const mstDigits = String(company.mst || '').replace(/\D/g, '').slice(-Math.max(1, option.mstDigits ?? option.chars));
  if (option.formula === 'initials' || option.formula === 'all_initials') return initials;
  if (option.formula === 'mst') return mstDigits;
  return initials + mstDigits;
}
function computePresetPrefix(company: CompanyRow, strategy: PrefixPresetStrategy, mstDigits = 3, nameWords = 2, nameChars = 1): string {
  const digits = Math.max(1, Math.min(10, mstDigits));
  const wordCount = clampPrefixNameWords(nameWords);
  const charCount = clampPrefixNameChars(nameChars);
  const mstSuffix = String(company.mst || '').replace(/\D/g, '').slice(-digits);
  const wordsPrefix = wordCount === 2 && charCount === 1 && company.prefix_strategies?.last_2_words ? company.prefix_strategies.last_2_words : computeCustomPrefix(company, { name: 'name words', formula: 'initials', words: wordCount, chars: charCount });
  const allWordsPrefix = charCount === 1 && company.prefix_strategies?.all_name_words ? company.prefix_strategies.all_name_words : computeCustomPrefix(company, { name: 'all words', formula: 'all_initials', words: wordCount, chars: charCount });
  if (strategy === 'last_3_mst') return mstSuffix;
  if (strategy === '2_words_mst') return `${wordsPrefix}${mstSuffix}`;
  if (strategy === 'all_name_words') return allWordsPrefix;
  return wordsPrefix;
}

function normalizedPrefixStrategy(strategy: string): PrefixPresetStrategy {
  return strategy === 'last_3_mst' || strategy === '2_words_mst' || strategy === 'all_name_words' ? strategy : 'last_2_words';
}

function normalizeMissingMstPrefixStrategy(strategy: unknown): PrefixPresetStrategy {
  return strategy === 'last_2_words' || strategy === 'all_name_words' ? strategy : 'all_name_words';
}
function companyDisplayGroups(companies: CompanyRow[], processingGroups: ProcessingGroup[] = defaultProcessingGroups('purchase'), duplicatePrefixes: Set<string> = new Set()): CompanyDisplayGroup[] {
  const rows = companies.map((company, index) => ({ company, index }));
  const groups = normalizeProcessingGroups(processingGroups, 'purchase');
  const groupIds = new Set(groups.map((group) => group.id));
  const duplicateGroups = groups
    .filter((group) => group.id !== IGNORED_GROUP_ID)
    .map((group) => ({
      title: `Prefix trùng lặp - ${group.label || group.id} cần kiểm tra`,
      className: 'duplicate-section',
      rows: rows
        .filter(({ company }) => displayCompanyGroupId(company, groupIds, 'pending') === group.id && duplicatePrefixes.has(duplicatePrefixKey(group.id, committedCompanyPrefix(company))))
        .sort((left, right) => {
          const prefixCompare = committedCompanyPrefix(left.company).localeCompare(committedCompanyPrefix(right.company), 'vi', { numeric: true, sensitivity: 'base' });
          return prefixCompare || left.index - right.index;
        }),
    }))
    .filter((group) => group.rows.length);
  const displayGroups = groups
    .filter((group) => group.id !== IGNORED_GROUP_ID)
    .map((group) => ({
      title: group.label || group.id,
      className: group.id === MATERIALS_GROUP_ID ? 'active-section' : 'company-group-section',
      rows: rows.filter(({ company }) => {
        const groupId = displayCompanyGroupId(company, groupIds, 'pending');
        if (groupId === group.id && duplicatePrefixes.has(duplicatePrefixKey(group.id, committedCompanyPrefix(company)))) return false;
        return groupId === group.id;
      }),
    }))
    .filter((group) => group.rows.length);
  const skippedRows = rows.filter(({ company }) => pendingCompanyGroup(company) === IGNORED_GROUP_ID);
  return [
    ...duplicateGroups,
    ...displayGroups,
    { title: 'Các công ty đã bỏ qua', className: 'skipped-section', rows: skippedRows },
  ];
}
function compareCompanyRowsByPrefix(left: { company: CompanyRow; index: number }, right: { company: CompanyRow; index: number }): number {
  const prefixCompare = committedCompanyPrefix(left.company).localeCompare(committedCompanyPrefix(right.company), 'vi', { numeric: true, sensitivity: 'base' });
  return prefixCompare || left.index - right.index;
}

function firstDisplayedCompanyIndex(companies: CompanyRow[], processingGroups: ProcessingGroup[] = defaultProcessingGroups('purchase')): number {
  return selectableCompanyIndexes(companies, processingGroups)[0] ?? -1;
}

function selectableCompanyIndexes(companies: CompanyRow[], processingGroups: ProcessingGroup[] = defaultProcessingGroups('purchase')): number[] {
  return companyDisplayGroups(companies, processingGroups)
    .filter((group) => group.className !== 'skipped-section')
    .flatMap((group) => group.rows)
    .filter(({ company }) => pendingCompanyGroup(company) !== IGNORED_GROUP_ID)
    .map(({ index }) => index);
}

function selectedCompanyIndexAfterGrouping(companies: CompanyRow[], currentIndex: number, processingGroups: ProcessingGroup[] = defaultProcessingGroups('purchase')): number {
  const indexes = selectableCompanyIndexes(companies, processingGroups);
  if (!indexes.length) return -1;
  if (indexes.includes(currentIndex)) return currentIndex;
  return indexes.find((index) => index > currentIndex) ?? [...indexes].reverse().find((index) => index < currentIndex) ?? indexes[0];
}

function normalizedCompanyPrefix(company: CompanyRow): string {
  return (company.value || '').trim().toUpperCase();
}

function committedCompanyPrefix(company: CompanyRow): string {
  return (company.committed_prefix ?? company.value ?? '').trim().toUpperCase();
}

function committedCompanyGroup(company: CompanyRow): string {
  return String(company.group_id || (company.process === false ? IGNORED_GROUP_ID : MATERIALS_GROUP_ID));
}

function pendingCompanyGroup(company: CompanyRow): string {
  return String(company.pending_group_id || company.group_id || (company.pending_process === false || company.process === false ? IGNORED_GROUP_ID : MATERIALS_GROUP_ID));
}

function displayCompanyGroupId(company: CompanyRow, groupIds: Set<string>, mode: 'pending' | 'committed' = 'pending'): string {
  const groupId = mode === 'committed' ? committedCompanyGroup(company) : pendingCompanyGroup(company);
  if (groupId === IGNORED_GROUP_ID) return IGNORED_GROUP_ID;
  return groupIds.has(groupId) ? groupId : MATERIALS_GROUP_ID;
}

function duplicatePrefixKey(groupId: string, prefix: string): string {
  return `${groupId}\u0000${prefix}`;
}

function duplicatePrefixParts(key: string): [string, string] {
  const separatorIndex = key.indexOf('\u0000');
  if (separatorIndex < 0) return [MATERIALS_GROUP_ID, key];
  return [key.slice(0, separatorIndex), key.slice(separatorIndex + 1)];
}

function isMaterialCompany(company: CompanyRow): boolean {
  return committedCompanyGroup(company) === MATERIALS_GROUP_ID;
}

function isIgnoredCompany(company: CompanyRow): boolean {
  return committedCompanyGroup(company) === IGNORED_GROUP_ID;
}

function hasCompanyDraftChanges(company: CompanyRow): boolean {
  const pendingProcess = company.pending_process ?? company.process ?? true;
  const appliedProcess = company.process ?? true;
  return pendingProcess !== appliedProcess || pendingCompanyGroup(company) !== committedCompanyGroup(company) || normalizedCompanyPrefix(company) !== committedCompanyPrefix(company);
}

function duplicatePrefixSetForRows(companies: CompanyRow[], processingGroups: ProcessingGroup[] = defaultProcessingGroups('purchase')): Set<string> {
  const counts = new Map<string, number>();
  const groupIds = new Set(normalizeProcessingGroups(processingGroups, 'purchase').map((group) => group.id));
  companies.forEach((company) => {
    const groupId = displayCompanyGroupId(company, groupIds, 'committed');
    if (groupId === IGNORED_GROUP_ID) return;
    const prefix = committedCompanyPrefix(company);
    if (!prefix) return;
    const key = duplicatePrefixKey(groupId, prefix);
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  return new Set(Array.from(counts.entries()).filter(([, count]) => count > 1).map(([key]) => key));
}

function sortAppliedCompanyRows(companies: CompanyRow[], processingGroups: ProcessingGroup[] = defaultProcessingGroups('purchase')): CompanyRow[] {
  const duplicatePrefixSet = duplicatePrefixSetForRows(companies, processingGroups);
  const groupOrder = new Map(normalizeProcessingGroups(processingGroups, 'purchase').map((group, index) => [group.id, index]));
  const groupIds = new Set(groupOrder.keys());
  const ignoredOrder = Number.MAX_SAFE_INTEGER;
  return companies
    .map((company, index) => ({ company, index }))
    .sort((left, right) => {
      const leftGroupId = committedCompanyGroup(left.company);
      const rightGroupId = committedCompanyGroup(right.company);
      const leftGroup = leftGroupId === IGNORED_GROUP_ID ? ignoredOrder : (groupOrder.get(leftGroupId) ?? groupOrder.get(MATERIALS_GROUP_ID) ?? 0);
      const rightGroup = rightGroupId === IGNORED_GROUP_ID ? ignoredOrder : (groupOrder.get(rightGroupId) ?? groupOrder.get(MATERIALS_GROUP_ID) ?? 0);
      const groupCompare = leftGroup - rightGroup;
      if (groupCompare) return groupCompare;
      const leftDuplicateKey = duplicatePrefixKey(displayCompanyGroupId(left.company, groupIds, 'committed'), committedCompanyPrefix(left.company));
      const rightDuplicateKey = duplicatePrefixKey(displayCompanyGroupId(right.company, groupIds, 'committed'), committedCompanyPrefix(right.company));
      if (duplicatePrefixSet.has(leftDuplicateKey) || duplicatePrefixSet.has(rightDuplicateKey)) {
        return compareCompanyRowsByPrefix(left, right);
      }
      return left.index - right.index;
    })
    .map(({ company }) => company);
}

function productDisplayCode(company: CompanyRow, productName: string, previewCodes: Record<string, string>, overrides: Record<string, string>, replacements: Record<string, string> = {}, includePrefix: boolean = false) {
  const override = overrides[productKey(companyConfigKey(company), productName)];
  if (override) return applyDisplayProductCodeReplacement(override, replacements);
  const preview = applyDisplayProductCodeReplacement(previewCodes[productName] || '', replacements);
  const appliedPrefix = sanitizeDisplayProductCode(committedCompanyPrefix(company));
  if (includePrefix && appliedPrefix) {
    return applyDisplayProductCodeReplacement(`${appliedPrefix}.${preview}`, replacements);
  }
  return preview;
}

function sanitizeDisplayProductCode(value: string) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '')
    .replace(/[^A-Za-z0-9.]+/g, '');
}


function isInventoryAllocationRunning(job: InventoryAllocationJob | null): boolean {
  const status = String(job?.status || '').toLowerCase();
  return status === 'queued' || status === 'running';
}

function inventoryJobProgress(job: InventoryAllocationJob | null): OperationProgress | null {
  if (!job) return null;
  const done = Math.max(0, Number(job.done ?? 0));
  const total = Math.max(0, Number(job.total ?? 0));
  const percent = total > 0 ? Math.round((Math.min(done, total) / total) * 100) : Math.max(0, Math.min(100, Number(job.progress ?? 0)));
  return {
    operation_id: 'inventory-allocation',
    status: job.status || 'running',
    done,
    total,
    percent,
    label: job.label || 'Đang phân bổ tồn kho...',
  };
}

function formatOperationStatus(progress: OperationProgress | null, fallback: string) {
  if (!progress) return fallback;
  const label = progress.label || fallback;
  const total = Number(progress.total || 0);
  const done = Number(progress.done || 0);
  const percent = Math.max(0, Math.min(100, Number(progress.percent || 0)));
  if (total > 1) return `${label} ${percent}% (${formatCount(done)}/${formatCount(total)} dòng)`;
  if (percent > 0 || progress.status === 'queued' || progress.status === 'running') return `${label} ${percent}%`;
  return label;
}

function formatInventoryJobStatus(job: InventoryAllocationJob | null) {
  return formatOperationStatus(inventoryJobProgress(job), job?.label || 'Đang phân bổ tồn kho...');
}

function CaoThanhPriceStage({
  groups,
  filterPercent,
  marginPercent,
  busy,
  onRefresh,
  onGroupPercentChange,
  onBucketMarginChange,
  onFilterPercentChange,
  onMarginPercentChange,
  onApplyFilter,
  onApplyMargin,
  onExportReport,
}: {
  groups: CaoThanhPriceGroup[];
  filterPercent: number;
  marginPercent: number;
  busy: boolean;
  onRefresh: () => void;
  onGroupPercentChange: (groupKey: string, value: number) => void;
  onBucketMarginChange: (groupKey: string, bucketKey: string, value: number) => void;
  onFilterPercentChange: (value: number) => void;
  onMarginPercentChange: (value: number) => void;
  onApplyFilter: () => void;
  onApplyMargin: () => void;
  onExportReport: () => void;
}) {
  const totals = caoThanhPriceTotals(groups);
  return (
    <div className="cao-price-step">
      <div className="stage-toolbar">
        <div>
          <h3>Loc don gia Cao Thanh</h3>
          <p className="muted">Tam thoi hien bang loc gia co ban. Phan view/report chi tiet se lam rieng o buoc sau.</p>
        </div>
        <div className="toolbar-actions">
          <button type="button" className="btn-secondary" disabled={busy} onClick={onRefresh}>{groups.length ? 'Tao lai nhom gia' : 'Tao nhom gia'}</button>
          <button type="button" className="btn-secondary" disabled={busy || !groups.length} onClick={onExportReport}>Xuat bao cao gia</button>
        </div>
      </div>

      <div className="cao-price-stats">
        <span><strong>{groups.length}</strong> Ma VT</span>
        <span><strong>{formatCount(totals.rows)}</strong> dong gia</span>
        <span><strong>{formatDecimal(totals.quantity)}</strong> so luong</span>
        <span><strong>{formatMoney(totals.amount)}</strong> doanh thu</span>
        <span><strong>{formatMoney(totals.cost)}</strong> gia von du tinh</span>
      </div>

      <div className="cao-price-controls">
        <label>Ap % loc<input type="number" min={0.1} step={0.1} value={filterPercent} onChange={(event) => onFilterPercentChange(Number(event.currentTarget.value || 8))} /></label>
        <button type="button" className="btn-secondary" disabled={busy || !groups.length} onClick={onApplyFilter}>Ap % loc</button>
        <label>Ap % lai<input type="number" min={0} step={0.1} value={marginPercent} onChange={(event) => onMarginPercentChange(Number(event.currentTarget.value || 0))} /></label>
        <button type="button" className="btn-secondary" disabled={busy || !groups.length} onClick={onApplyMargin}>Ap % lai</button>
      </div>

      {!groups.length ? (
        <PlaceholderStage title="Chưa có nhóm giá" detail="Bấm Tạo nhóm giá để tạo danh sách lọc đơn giá từ các công ty/hàng hóa đã áp dụng." />
      ) : (
        <div className="inner-scroll cao-price-scroll">
          <table className="cao-price-table">
            <thead>
              <tr><th>Mã VT</th><th>Hàng hóa</th><th>Dòng giá</th><th>Giá min/max</th><th>% lọc</th><th>Nhóm sau lọc</th></tr>
            </thead>
            <tbody>
              {groups.map((group) => (
                <tr key={group.key}>
                  <td><code>{group.code}</code><div className="muted">{group.companies}</div></td>
                  <td>{group.productNames}</td>
                  <td>{group.sourceRows.length} dong / {formatDecimal(group.quantity)}</td>
                  <td>{formatMoney(group.min)} - {formatMoney(group.max)}</td>
                  <td><input className="price-percent-input" type="number" min={0.1} step={0.1} value={group.filterPercent} onChange={(event) => onGroupPercentChange(group.key, Number(event.currentTarget.value || 8))} /></td>
                  <td>
                    <div className="cao-bucket-list">
                      {group.buckets.map((bucket) => (
                        <div className="cao-bucket" key={bucket.key}>
                          <strong>{bucket.finalCode}</strong>
                          <span>{formatMoney(bucket.min)} - {formatMoney(bucket.max)}</span>
                          <span>{bucket.count} dong, TB {formatMoney(bucket.averagePrice)}</span>
                          <label>% lai <input className="price-percent-input" type="number" min={0} step={0.1} value={bucket.marginPercent} onChange={(event) => onBucketMarginChange(group.key, bucket.key, Number(event.currentTarget.value || 0))} /></label>
                          <span>Gia von {formatMoney(bucket.costUnitPrice)}</span>
                        </div>
                      ))}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function LegacyProfileWorkspace({ profile, label, licenseReady, setShellStatus }: { profile: ProfileKey; label: string; licenseReady: boolean; setShellStatus: (message: string) => void }) {
  if (!licenseReady) {
    return <PlaceholderStage title={`${label}: cần license`} detail="Kích hoạt license trước khi mở workflow profile này." />;
  }
  if (profile === 'cao_thanh') {
    return <CaoThanhWorkflow label={label} setShellStatus={setShellStatus} />;
  }
  return (
    <div className="legacy-profile-workspace native-profile-workspace">
      <section className="profile-migration-card compact-profile-card">
        <span className="upload-step-badge">{label}</span>
        <h3>{label}</h3>
        <p>Profile này vẫn đang chờ migrate workflow riêng sang React. Cao Thành đang được chuyển trước.</p>
      </section>
    </div>
  );
}

type CaoThanhStep = 'upload' | 'company' | 'price' | 'export';

type CaoThanhColumns = {
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

type CaoThanhPriceSourceRow = {
  key: string;
  code: string;
  company: string;
  mst: string;
  productName: string;
  unit: string;
  price: number;
  quantity: number;
  amount: number;
  excelRow: string;
  invoiceNo: string;
  invoiceDate: string;
};

type CaoThanhPriceBucket = {
  key: string;
  label: string;
  finalCode: string;
  min: number;
  max: number;
  count: number;
  quantity: number;
  amount: number;
  averagePrice: number;
  marginPercent: number;
  costUnitPrice: number;
  costAmount: number;
  profitAmount: number;
  profitPercent: number;
  rows: CaoThanhPriceSourceRow[];
};

type CaoThanhPriceGroup = {
  key: string;
  code: string;
  companies: string;
  productNames: string;
  unit: string;
  sourceRows: CaoThanhPriceSourceRow[];
  min: number;
  max: number;
  priceCount: number;
  quantity: number;
  amount: number;
  averagePrice: number;
  filterPercent: number;
  buckets: CaoThanhPriceBucket[];
};

const defaultCaoThanhColumns: CaoThanhColumns = {
  company_col: 'F',
  mst_col: 'G',
  address_col: 'H',
  product_col: 'M',
  qty_col: 'O',
  price_col: 'P',
  output_col: 'L',
  invoice_status_col: 'AJ',
  invoice_status_skip_values: ['Hóa đơn đã bị điều chỉnh', 'Hóa đơn bị thay thế', 'Hóa đơn đã bị thay thế'],
};

function CaoThanhWorkflow({ label, setShellStatus }: { label: string; setShellStatus: (message: string) => void }) {
  const [step, setStep] = useState<CaoThanhStep>('upload');
  const [summary, setSummary] = useState<UploadSummary | null>(null);
  const [columns, setColumns] = useState<CaoThanhColumns>(defaultCaoThanhColumns);
  const [companies, setCompanies] = useState<CompanyRow[]>([]);
  const [selectedCompanyIndex, setSelectedCompanyIndex] = useState(-1);
  const [previewCodes, setPreviewCodes] = useState<Record<string, string>>({});
  const [manualOverrides, setManualOverrides] = useState<Record<string, string>>({});
  const [productCodeReplacements, setProductCodeReplacements] = useState<Record<string, string>>({});
  const [wordRules, setWordRules] = useState<Record<string, string>>({});
  const [firstWordRules, setFirstWordRules] = useState<Record<string, string>>({});
  const [repeatedPhrases, setRepeatedPhrases] = useState<string[]>(['inox']);
  const [includeCompanyPrefix, setIncludeCompanyPrefix] = useState(true);
  const [priceRangeRules, setPriceRangeRules] = useState<Record<string, any>>({});
  const [priceGroups, setPriceGroups] = useState<CaoThanhPriceGroup[]>([]);
  const [priceFilterAllPercent, setPriceFilterAllPercent] = useState(8);
  const [priceAdjustAllPercent, setPriceAdjustAllPercent] = useState(0);
  const [inventoryPairs, setInventoryPairs] = useState<InventoryPair[]>([]);
  const [inventoryPairRules, setInventoryPairRules] = useState<InventoryRule[]>([]);
  const [useDefaultInventoryPair, setUseDefaultInventoryPair] = useState(false);
  const [defaultInventoryPairId, setDefaultInventoryPairId] = useState('');
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('Tải file bán ra Cao Thành để bắt đầu.');

  useEffect(() => {
    let active = true;
    void (async () => {
      try {
        const cfg = await getAppConfig();
        const profilesCfg = (cfg.profiles && typeof cfg.profiles === 'object' ? cfg.profiles : {}) as Record<string, any>;
        const profileCfg = profilesCfg.cao_thanh || {};
        const globalColumns = cfg.selected_profile === 'cao_thanh' && cfg.columns && typeof cfg.columns === 'object' ? cfg.columns as Record<string, unknown> : {};
        const savedColumns = profileCfg.columns && typeof profileCfg.columns === 'object' ? profileCfg.columns as Record<string, unknown> : {};
        if (!active) return;
        setColumns(normalizeCaoThanhColumns({ ...defaultCaoThanhColumns, ...globalColumns, ...savedColumns }));
        setWordRules(cleanStringMap(profileCfg.word_rules));
        setFirstWordRules(cleanStringMap(profileCfg.first_word_rules));
        setRepeatedPhrases(Array.isArray(profileCfg.repeated_phrase_removals) ? profileCfg.repeated_phrase_removals.map(String) : ['inox']);
        setIncludeCompanyPrefix(profileCfg.include_company_prefix !== false);
        setPriceRangeRules(profileCfg.price_range_rules && typeof profileCfg.price_range_rules === 'object' ? profileCfg.price_range_rules : {});
        setPriceAdjustAllPercent(Number(profileCfg.price_adjust_all_percent || 0));
        setManualOverrides(cleanStringMap(profileCfg.manual_code_overrides));
        setProductCodeReplacements(cleanStringMap(profileCfg.product_code_replacements));
        setMessage('Đã tải cấu hình Cao Thành.');
      } catch (error) {
        if (!active) return;
        setMessage(error instanceof Error ? error.message : String(error));
      }
    })();
    return () => { active = false; };
  }, []);

  function updateMessage(nextMessage: string) {
    setMessage(nextMessage);
    setShellStatus(nextMessage);
  }

  async function uploadCaoThanh(file: File | undefined) {
    if (!file) return;
    setBusy(true);
    updateMessage('Đang tải file bán ra Cao Thành...');
    try {
      const nextSummary = await uploadExcel(file);
      setSummary(nextSummary);
      setCompanies([]);
      setPreviewCodes({});
      setPriceGroups([]);
      setStep('company');
      updateMessage(`Đã tải ${nextSummary.original_name}. Chọn cột rồi tải danh sách công ty.`);
    } catch (error) {
      updateMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function analyzeCompanies() {
    if (!summary) return;
    setBusy(true);
    updateMessage('Đang đọc công ty và hàng hóa Cao Thành...');
    try {
      const result = await analyzeGenericWorkbook({ saved_name: summary.saved_name, original_name: summary.original_name, profile: 'cao_thanh', ...columns });
      const rows = sortAppliedCompanyRows(result.companies.map((company) => {
        const process = company.process ?? true;
        const value = normalizePrefixValue(company.value || company.default_prefix || '');
        return { ...company, value, process, pending_process: process, committed_prefix: value };
      }));
      const codes = await loadCaoThanhPreviewCodes(rows, wordRules, firstWordRules, repeatedPhrases, productCodeReplacements);
      setCompanies(rows);
      setSelectedCompanyIndex(firstDisplayedCompanyIndex(rows));
      setPreviewCodes(codes);
      setPriceGroups([]);
      updateMessage(`Đã tải ${result.company_count} công ty, ${result.rows_to_process} dòng xử lý Cao Thành.`);
    } catch (error) {
      updateMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function updateCompanyPending(index: number, pending: boolean) {
    setCompanies((rows) => rows.map((company, rowIndex) => rowIndex === index ? { ...company, pending_process: pending } : company));
    setSelectedCompanyIndex(index);
  }

  function bulkUpdateCompanies(pending: boolean) {
    setCompanies((rows) => rows.map((company) => ({ ...company, pending_process: pending })));
  }

  function updateProduct(companyIndex: number, productName: string, selected: boolean) {
    setCompanies((rows) => rows.map((company, rowIndex) => {
      if (rowIndex !== companyIndex) return company;
      const current = new Set(selectedProductNames(company));
      if (selected) current.add(productName);
      else current.delete(productName);
      return { ...company, selected_product_names: company.all_products.map((product) => product.name).filter((name) => current.has(name)) };
    }));
    setSelectedCompanyIndex(companyIndex);
    setPriceGroups([]);
  }

  function updateProductCode(companyIndex: number, productName: string, code: string) {
    const company = companies[companyIndex];
    if (!company) return;
    setManualOverrides((current) => ({ ...current, [productKey(companyConfigKey(company), productName)]: code.toUpperCase() }));
    setSelectedCompanyIndex(companyIndex);
    setPriceGroups([]);
  }

  function updateCompanyPrefix(index: number, value: string) {
    setCompanies((rows) => rows.map((company, rowIndex) => rowIndex === index ? { ...company, value } : company));
    setSelectedCompanyIndex(index);
    setPriceGroups([]);
  }

  async function refreshPreviewCodes() {
    if (!companies.length) return;
    setBusy(true);
    updateMessage('Đang cập nhật mã VT preview Cao Thành...');
    try {
      const codes = await loadCaoThanhPreviewCodes(companies, wordRules, firstWordRules, repeatedPhrases, productCodeReplacements);
      setPreviewCodes(codes);
      setPriceGroups([]);
      updateMessage('Đã cập nhật mã VT preview Cao Thành.');
    } catch (error) {
      updateMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function applyCompanyChoices() {
    const committed = sortAppliedCompanyRows(companies.map((company) => {
      const process = company.pending_process ?? company.process ?? true;
      const value = normalizePrefixValue(company.value);
      return { ...company, process, pending_process: process, value, committed_prefix: value };
    }));
    const nextGroups = buildCaoThanhPriceGroups(committed, previewCodes, manualOverrides, productCodeReplacements, includeCompanyPrefix, priceRangeRules, priceAdjustAllPercent);
    setCompanies(committed);
    setSelectedCompanyIndex(firstDisplayedCompanyIndex(committed));
    setPriceGroups(nextGroups);
    setStep('price');
    updateMessage(`Đã áp dụng ${committed.filter((company) => company.process !== false).length} công ty. Có ${nextGroups.length} mã VT để lọc giá.`);
  }

  function updateGroupPercent(groupKey: string, value: number) {
    setPriceGroups((groups) => groups.map((group) => group.key === groupKey ? rebuildCaoThanhPriceGroup({ ...group, filterPercent: clampPercent(value, group.filterPercent || 8) }, priceAdjustAllPercent) : group));
  }

  function updateBucketMargin(groupKey: string, bucketKey: string, value: number) {
    setPriceGroups((groups) => groups.map((group) => group.key === groupKey ? { ...group, buckets: group.buckets.map((bucket) => bucket.key === bucketKey ? rebuildCaoThanhBucket({ ...bucket, marginPercent: clampPercent(value, bucket.marginPercent || 0) }) : bucket) } : group));
  }

  function applyBulkPriceFilter() {
    setPriceGroups((groups) => groups.map((group) => rebuildCaoThanhPriceGroup({ ...group, filterPercent: clampPercent(priceFilterAllPercent, 8) }, priceAdjustAllPercent)));
  }

  function applyBulkMargin() {
    setPriceGroups((groups) => groups.map((group) => ({ ...group, buckets: group.buckets.map((bucket) => rebuildCaoThanhBucket({ ...bucket, marginPercent: clampPercent(priceAdjustAllPercent, 0) })) })));
  }

  async function exportPriceReport() {
    if (!priceGroups.length) return;
    setBusy(true);
    updateMessage('Đang xuất báo cáo lọc giá Cao Thành...');
    try {
      const payload = caoThanhPriceReportPayload(summary?.original_name || 'cao_thanh.xlsx', priceGroups);
      const blob = await exportPriceReportWorkbook(payload);
      const saved = await saveBlob(blob, caoThanhReportFileName(summary?.original_name || 'cao_thanh.xlsx'));
      updateMessage(saved ? 'Đã xuất báo cáo lọc giá Cao Thành.' : 'Đã hủy lưu báo cáo lọc giá Cao Thành.');
    } catch (error) {
      updateMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function processCaoThanh() {
    if (!summary) return;
    const committedGroups = priceGroups.length ? priceGroups : buildCaoThanhPriceGroups(companies, previewCodes, manualOverrides, productCodeReplacements, includeCompanyPrefix, priceRangeRules, priceAdjustAllPercent);
    setBusy(true);
    updateMessage('Đang xử lý file Cao Thành...');
    try {
      const rangeRules = caoThanhRangeRules(committedGroups);
      setPriceRangeRules((current) => ({ ...current, ...rangeRules }));
      const payload = buildCaoThanhProcessPayload(summary, columns, companies, manualOverrides, productCodeReplacements, includeCompanyPrefix, wordRules, firstWordRules, repeatedPhrases, rangeRules, priceAdjustAllPercent);
      const blob = await processGenericWorkbook(payload);
      const saved = await saveBlob(blob, `${fileStem(summary.original_name)}_fdi.xls`);
      setStep('export');
      updateMessage(saved ? 'Đã xuất file Cao Thành và lưu cấu hình lọc giá.' : 'Đã xử lý Cao Thành; người dùng đã hủy lưu file.');
    } catch (error) {
      updateMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function updateWordRule(index: number, field: 'from' | 'to', value: string) {
    const entries = Object.entries(wordRules);
    entries[index] = field === 'from' ? [value, entries[index]?.[1] || ''] : [entries[index]?.[0] || '', value];
    setWordRules(Object.fromEntries(entries.filter(([from]) => from.trim())));
    setPriceGroups([]);
  }

  function addWordRule() {
    setWordRules((current) => ({ ...current, '': '' }));
  }

  function updateRepeated(index: number, value: string) {
    setRepeatedPhrases((current) => current.map((item, rowIndex) => rowIndex === index ? value : item));
    setPriceGroups([]);
  }

  function addRepeated() {
    setRepeatedPhrases((current) => [...current, '']);
  }

  function removeRepeated(index: number) {
    setRepeatedPhrases((current) => current.filter((_, rowIndex) => rowIndex !== index));
    setPriceGroups([]);
  }

  function commitProductCodeReplacements(nextReplacements: Record<string, string>) {
    setProductCodeReplacements(normalizeProductCodeReplacements(nextReplacements));
    setPriceGroups([]);
  }

  async function importProductCodeReplacementFile(file: File | undefined) {
    if (!file) return;
    setBusy(true);
    updateMessage('Dang nhap danh sach doi ma VT Cao Thanh...');
    try {
      const result = await importProductCodeReplacements(file);
      setProductCodeReplacements((current) => normalizeProductCodeReplacements({ ...current, ...result.product_code_replacements }));
      setPriceGroups([]);
      updateMessage(`Da nhap ${result.count} dong doi ma VT.`);
    } catch (error) {
      updateMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function applyRelatedProductCodeUpdates(updates: RelatedProductCodeUpdate[]) {
    if (!updates.length) return;
    setManualOverrides((current) => {
      const next = { ...current };
      updates.forEach((item) => {
        const company = companies[item.companyIndex];
        const code = sanitizeDisplayProductCode(item.code).toUpperCase();
        if (!company || !item.productName || !code) return;
        next[productKey(companyConfigKey(company), item.productName)] = code;
      });
      return next;
    });
    setPriceGroups([]);
  }

  const priceTotals = caoThanhPriceTotals(priceGroups);

  return (
    <div className="cao-thanh-workspace">
      <div className="cao-stage-pills" role="tablist" aria-label="Cao Thành workflow">
        <button type="button" className={step === 'upload' ? 'active' : ''} disabled={busy} onClick={() => setStep('upload')}>1. Tải file bán ra</button>
        <button type="button" className={step === 'company' ? 'active' : ''} disabled={busy || !summary} onClick={() => setStep('company')}>2. Cột & công ty</button>
        <button type="button" className={step === 'price' ? 'active' : ''} disabled={busy || !priceGroups.length} onClick={() => setStep('price')}>3. Lọc đơn giá</button>
        <button type="button" className={step === 'export' ? 'active' : ''} disabled={busy || !summary} onClick={() => setStep('export')}>4. Xuất file</button>
      </div>
      <div className="status-bar"><strong>{label}</strong><span>{busy ? 'Đang xử lý... ' : ''}{message}</span></div>

      {step === 'upload' && <UploadStage title="HD bán ra Cao Thành" summary={summary} disabled={busy} onUpload={uploadCaoThanh} />}

      {step === 'company' && (
        <div className="cao-company-step">
          <div className="cao-column-card">
            <div className="stage-toolbar"><h3>Cột xử lý Cao Thành</h3><div className="toolbar-actions"><button type="button" disabled={busy || !summary} onClick={analyzeCompanies}>{companies.length ? 'Tải lại danh sách' : 'Tải danh sách công ty'}</button></div></div>
            <div className="column-grid">{caoThanhColumnFields.map((field) => <label key={field.key}><span>{field.label}</span><select value={columns[field.key]} onChange={(event) => setColumns({ ...columns, [field.key]: event.currentTarget.value })}>{summary?.columns.map((column) => <option key={`${field.key}-${column.letter}`} value={column.letter}>{column.label}</option>) ?? <option value={columns[field.key]}>{columns[field.key]}</option>}</select></label>)}</div>
          </div>
          {companies.length ? <CompanyRulesStage companies={companies} selectedCompanyIndex={selectedCompanyIndex} productPreviewCodes={previewCodes} productCodeOverrides={manualOverrides} productCodeReplacements={productCodeReplacements} wordRules={wordRules} repeatedPhrases={repeatedPhrases} inventoryPairs={inventoryPairs} useDefaultInventoryPair={useDefaultInventoryPair} defaultInventoryPairId={defaultInventoryPairId} inventoryPairRules={inventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy="last_2_words" prefixMstDigits={3} prefixNameWords={2} onIncludeCompanyPrefixChange={(include) => { setIncludeCompanyPrefix(include); setPriceGroups([]); }} onCompanyPrefixChange={updateCompanyPrefix} onPrefixMstDigitsChange={() => {}} onPrefixNameWordsChange={() => {}} onApplyPrefixPresetToAll={() => {}} onCompanySelect={setSelectedCompanyIndex} onCompanyChange={updateCompanyPending} onBulkCompanyChange={bulkUpdateCompanies} onProductChange={updateProduct} onProductCodeChange={updateProductCode} onApplyChoices={applyCompanyChoices} onRefreshPreviews={refreshPreviewCodes} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeated} onAddRepeated={addRepeated} onRemoveRepeated={removeRepeated} onProductCodeReplacementChange={() => {}} onAddProductCodeReplacement={() => {}} onRemoveProductCodeReplacement={() => {}} onProductCodeReplacementsCommit={commitProductCodeReplacements} onApplyRelatedProductCodes={applyRelatedProductCodeUpdates} onImportProductCodeReplacements={importProductCodeReplacementFile} onAddInventoryPair={() => { const id = `pair-${Date.now()}`; setInventoryPairs([...inventoryPairs, { id, ma_kho: '', tk_vat_tu: '' }]); setDefaultInventoryPairId(defaultInventoryPairId || id); }} onInventoryPairChange={(index, field, value) => setInventoryPairs((rows) => rows.map((pair, rowIndex) => rowIndex === index ? { ...pair, [field]: value.toUpperCase() } : pair))} onRemoveInventoryPair={(index) => setInventoryPairs((rows) => rows.filter((_, rowIndex) => rowIndex !== index))} onInventoryDefaultsChange={(update) => { if ('useDefaultInventoryPair' in update) setUseDefaultInventoryPair(Boolean(update.useDefaultInventoryPair)); if ('defaultInventoryPairId' in update) setDefaultInventoryPairId(String(update.defaultInventoryPairId || '')); }} onAddInventoryRule={() => setInventoryPairRules((rows) => [...rows, { source_col: 'M', operator: 'contains', value: '', pair_id: defaultInventoryPairId || inventoryPairs[0]?.id || '', enabled: true, priority: 1 }])} onInventoryRuleChange={(index, update) => setInventoryPairRules((rows) => rows.map((rule, rowIndex) => rowIndex === index ? { ...rule, ...update } : rule))} onRemoveInventoryRule={(index) => setInventoryPairRules((rows) => rows.filter((_, rowIndex) => rowIndex !== index))} /> : <PreviewPanel summary={summary} />}
        </div>
      )}

      {step === 'price' && (
        <div className="cao-price-step">
          <div className="stage-toolbar">
            <div><h3>Lọc đơn giá Cao Thành</h3><p className="muted">Gộp theo Mã VT cuối cùng, chia nhóm đơn giá theo % lọc, rồi xuất file với hậu tố .001/.002 khi cần.</p></div>
            <div className="toolbar-actions"><button type="button" className="btn-secondary" disabled={busy} onClick={() => setStep('company')}>Quay lại công ty</button><button type="button" className="btn-secondary" disabled={busy || !priceGroups.length} onClick={exportPriceReport}>Xuất báo cáo giá</button><button type="button" disabled={busy || !priceGroups.length} onClick={processCaoThanh}>Xử lý & xuất file</button></div>
          </div>
          <div className="cao-price-stats"><span><strong>{priceGroups.length}</strong>Mã VT</span><span><strong>{formatCount(priceTotals.rows)}</strong>Dòng giá</span><span><strong>{formatDecimal(priceTotals.quantity)}</strong>Số lượng</span><span><strong>{formatMoney(priceTotals.amount)}</strong>Doanh thu</span><span><strong>{formatMoney(priceTotals.cost)}</strong>Giá vốn dự tính</span></div>
          <div className="cao-price-controls"><label>Áp % lọc<input type="number" min={0.1} step={0.1} value={priceFilterAllPercent} onChange={(event) => setPriceFilterAllPercent(Number(event.currentTarget.value || 8))} /></label><button type="button" className="btn-secondary" disabled={busy || !priceGroups.length} onClick={applyBulkPriceFilter}>Áp % lọc</button><label>Áp % lãi<input type="number" min={0} step={0.1} value={priceAdjustAllPercent} onChange={(event) => setPriceAdjustAllPercent(Number(event.currentTarget.value || 0))} /></label><button type="button" className="btn-secondary" disabled={busy || !priceGroups.length} onClick={applyBulkMargin}>Áp % lãi</button></div>
          <div className="inner-scroll cao-price-scroll"><table className="cao-price-table"><thead><tr><th>Mã VT</th><th>Hàng hóa</th><th>Dòng giá</th><th>Giá min/max</th><th>% lọc</th><th>Nhóm sau lọc</th></tr></thead><tbody>{priceGroups.map((group) => <tr key={group.key}><td><code>{group.code}</code><div className="muted">{group.companies}</div></td><td>{group.productNames}</td><td>{group.sourceRows.length} dòng / {formatDecimal(group.quantity)}</td><td>{formatMoney(group.min)} → {formatMoney(group.max)}</td><td><input className="price-percent-input" type="number" min={0.1} step={0.1} value={group.filterPercent} onChange={(event) => updateGroupPercent(group.key, Number(event.currentTarget.value || 8))} /></td><td><div className="cao-bucket-list">{group.buckets.map((bucket) => <div className="cao-bucket" key={bucket.key}><strong>{bucket.finalCode}</strong><span>{formatMoney(bucket.min)} → {formatMoney(bucket.max)}</span><span>{bucket.count} dòng, TB {formatMoney(bucket.averagePrice)}</span><label>% lãi <input className="price-percent-input" type="number" min={0} step={0.1} value={bucket.marginPercent} onChange={(event) => updateBucketMargin(group.key, bucket.key, Number(event.currentTarget.value || 0))} /></label><span>Giá vốn {formatMoney(bucket.costUnitPrice)}</span></div>)}</div></td></tr>)}</tbody></table></div>
        </div>
      )}

      {step === 'export' && <ProcessStage title="Xuất file Cao Thành" detail="File đã được xử lý theo lựa chọn công ty, hàng hóa và lọc đơn giá. Có thể xuất lại file hoặc báo cáo giá từ cache cấu hình hiện tại." buttonLabel="Xử lý & xuất lại" disabled={busy || !summary} onProcess={processCaoThanh} />}
    </div>
  );
}

const caoThanhColumnFields: Array<{ key: keyof CaoThanhColumns; label: string }> = [
  { key: 'company_col', label: 'Công ty' },
  { key: 'mst_col', label: 'MST' },
  { key: 'address_col', label: 'Địa chỉ' },
  { key: 'product_col', label: 'Tên hàng' },
  { key: 'qty_col', label: 'Số lượng' },
  { key: 'price_col', label: 'Đơn giá' },
  { key: 'output_col', label: 'Mã VT xuất' },
  { key: 'invoice_status_col', label: 'Trạng thái HĐ' },
];

function normalizeCaoThanhColumns(raw: Record<string, unknown>): CaoThanhColumns {
  const columnValue = (key: keyof CaoThanhColumns) => {
    const value = String(raw[key] ?? defaultCaoThanhColumns[key] ?? '').trim().toUpperCase();
    return value || String(defaultCaoThanhColumns[key] || '').toUpperCase();
  };
  const skipValues = Array.isArray(raw.invoice_status_skip_values)
    ? raw.invoice_status_skip_values.map(String).filter((item) => item.trim())
    : defaultCaoThanhColumns.invoice_status_skip_values;
  return {
    company_col: columnValue('company_col'),
    mst_col: columnValue('mst_col'),
    address_col: columnValue('address_col'),
    product_col: columnValue('product_col'),
    qty_col: columnValue('qty_col'),
    price_col: columnValue('price_col'),
    output_col: columnValue('output_col'),
    invoice_status_col: columnValue('invoice_status_col'),
    invoice_status_skip_values: skipValues,
  };
}

function cleanStringMap(raw: unknown): Record<string, string> {
  if (!raw || typeof raw !== 'object') return {};
  return Object.fromEntries(
    Object.entries(raw as Record<string, unknown>)
      .map(([key, value]) => [key.trim(), String(value ?? '').trim()])
      .filter(([key, value]) => key && value),
  );
}

async function loadCaoThanhPreviewCodes(
  companies: CompanyRow[],
  wordRules: Record<string, string>,
  firstWordRules: Record<string, string>,
  repeatedPhraseRemovals: string[],
  productCodeReplacements: Record<string, string> = {},
) {
  const products = Array.from(new Set(companies.flatMap((company) => company.all_products.map((product) => product.name)).filter(Boolean)));
  if (!products.length) return {};
  const result = await previewGenericProductCodes({
    profile: 'cao_thanh',
    products,
    word_rules: wordRules,
    first_word_rules: firstWordRules,
    repeated_phrase_removals: repeatedPhraseRemovals.filter((phrase) => phrase.trim()),
    product_code_replacements: productCodeReplacements,
  });
  return result.codes;
}

function clampPercent(value: number, fallback: number) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(0, Math.min(1000, parsed));
}

function numericValue(value: unknown): number {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const text = String(value ?? '').trim();
  if (!text) return 0;
  const normalized = text.replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function fileStem(filename: string) {
  return String(filename || 'output').replace(/\.[^.]+$/, '') || 'output';
}

function toXlsName(filename: string) {
  return `${fileStem(filename)}.xls`;
}

function formatDecimal(value: number | undefined) {
  return Number(value || 0).toLocaleString('en-US', { maximumFractionDigits: 3 });
}

function formatMoney(value: number | undefined) {
  return Number(value || 0).toLocaleString('en-US', { maximumFractionDigits: 2 });
}

function buildCaoThanhPriceGroups(
  companies: CompanyRow[],
  previewCodes: Record<string, string>,
  manualOverrides: Record<string, string>,
  productCodeReplacements: Record<string, string>,
  includePrefix: boolean,
  savedRules: Record<string, any>,
  defaultMarginPercent: number,
): CaoThanhPriceGroup[] {
  const grouped = new Map<string, CaoThanhPriceGroup>();
  companies.forEach((company) => {
    if (company.process === false) return;
    const selected = new Set(selectedProductNames(company));
    company.all_products.forEach((product) => {
      if (!selected.has(product.name)) return;
      const code = productDisplayCode(company, product.name, previewCodes, manualOverrides, productCodeReplacements, includePrefix);
      if (!code) return;
      const rows = product.priceRows || [];
      rows.forEach((row, rowIndex) => {
        const price = numericValue(row.price);
        const quantity = numericValue(row.quantity);
        const amount = numericValue(row.amount) || price * quantity;
        if (price <= 0) return;
        const sourceRow: CaoThanhPriceSourceRow = {
          key: `${companyConfigKey(company)}|${product.name}|${row.excelRow ?? rowIndex}|${rowIndex}`,
          code,
          company: company.company,
          mst: company.mst,
          productName: product.name,
          unit: String(row.unit ?? ''),
          price,
          quantity,
          amount,
          excelRow: String(row.excelRow ?? ''),
          invoiceNo: String(row.invoiceNo ?? ''),
          invoiceDate: String(row.invoiceDate ?? ''),
        };
        const current = grouped.get(code);
        if (current) {
          current.sourceRows.push(sourceRow);
          current.companies = mergeLabelList(current.companies, company.company);
          current.productNames = mergeLabelList(current.productNames, product.name);
          current.unit = mergeLabelList(current.unit, sourceRow.unit);
        } else {
          const savedRule = savedRules && typeof savedRules === 'object' ? savedRules[code] : undefined;
          grouped.set(code, {
            key: code,
            code,
            companies: company.company,
            productNames: product.name,
            unit: sourceRow.unit,
            sourceRows: [sourceRow],
            min: 0,
            max: 0,
            priceCount: 0,
            quantity: 0,
            amount: 0,
            averagePrice: 0,
            filterPercent: Number(savedRule?.percent || 8),
            buckets: [],
          });
        }
      });
    });
  });
  return Array.from(grouped.values())
    .map((group) => rebuildCaoThanhPriceGroup(group, defaultMarginPercent))
    .sort((left, right) => left.code.localeCompare(right.code, 'vi'));
}

function mergeLabelList(current: string, next: string) {
  const values = new Set(current.split(' / ').filter(Boolean));
  if (next) values.add(next);
  return Array.from(values).slice(0, 4).join(' / ');
}

function rebuildCaoThanhPriceGroup(group: CaoThanhPriceGroup, defaultMarginPercent: number): CaoThanhPriceGroup {
  const rows = [...group.sourceRows].sort((left, right) => left.price - right.price);
  const prices = rows.map((row) => row.price).filter((price) => price > 0);
  const min = prices.length ? Math.min(...prices) : 0;
  const max = prices.length ? Math.max(...prices) : 0;
  const percent = clampPercent(group.filterPercent, 8) || 8;
  const step = min > 0 ? min * percent / 100 : 0;
  const rawBuckets = new Map<number, CaoThanhPriceSourceRow[]>();
  rows.forEach((row) => {
    const rawIndex = step > 0 ? Math.max(1, Math.floor((row.price - min) / step) + 1) : 1;
    rawBuckets.set(rawIndex, [...(rawBuckets.get(rawIndex) || []), row]);
  });
  const previousMargins = new Map(group.buckets.map((bucket) => [bucket.key, bucket.marginPercent || defaultMarginPercent || 0]));
  const buckets = Array.from(rawBuckets.entries())
    .sort(([left], [right]) => left - right)
    .map(([rawIndex, bucketRows], index) => {
      const suffix = `${index + 1}`.padStart(3, '0');
      const key = `${group.code}|${rawIndex}`;
      return rebuildCaoThanhBucket({
        key,
        label: `Nhóm ${index + 1}`,
        finalCode: `${group.code}.${suffix}`,
        min: Math.min(...bucketRows.map((row) => row.price)),
        max: Math.max(...bucketRows.map((row) => row.price)),
        count: bucketRows.length,
        quantity: bucketRows.reduce((sum, row) => sum + row.quantity, 0),
        amount: bucketRows.reduce((sum, row) => sum + row.amount, 0),
        averagePrice: 0,
        marginPercent: previousMargins.get(key) ?? defaultMarginPercent ?? 0,
        costUnitPrice: 0,
        costAmount: 0,
        profitAmount: 0,
        profitPercent: 0,
        rows: bucketRows,
      });
    });
  const quantity = rows.reduce((sum, row) => sum + row.quantity, 0);
  const amount = rows.reduce((sum, row) => sum + row.amount, 0);
  return {
    ...group,
    sourceRows: rows,
    min,
    max,
    priceCount: new Set(prices).size,
    quantity,
    amount,
    averagePrice: quantity ? amount / quantity : 0,
    filterPercent: percent,
    buckets,
  };
}

function rebuildCaoThanhBucket(bucket: CaoThanhPriceBucket): CaoThanhPriceBucket {
  const quantity = bucket.rows.reduce((sum, row) => sum + row.quantity, 0);
  const amount = bucket.rows.reduce((sum, row) => sum + row.amount, 0);
  const averagePrice = quantity ? amount / quantity : 0;
  const margin = clampPercent(bucket.marginPercent, 0);
  const costUnitPrice = averagePrice / (1 + margin / 100);
  const costAmount = costUnitPrice * quantity;
  const profitAmount = amount - costAmount;
  return {
    ...bucket,
    quantity,
    amount,
    averagePrice,
    marginPercent: margin,
    costUnitPrice,
    costAmount,
    profitAmount,
    profitPercent: amount ? profitAmount / amount * 100 : 0,
  };
}

function caoThanhRangeRules(groups: CaoThanhPriceGroup[]) {
  return Object.fromEntries(groups.filter((group) => group.min > 0 && group.max >= group.min).map((group) => [group.code, {
    min_price: group.min,
    max_price: group.max,
    percent: group.filterPercent,
    groups: group.buckets.map((bucket, index) => ({
      index: index + 1,
      label: bucket.label,
      min_price: bucket.min,
      max_price: bucket.max,
      average_price: bucket.averagePrice,
      adjust_percent: bucket.marginPercent,
    })),
  }]));
}

function caoThanhPriceTotals(groups: CaoThanhPriceGroup[]) {
  return {
    rows: groups.reduce((sum, group) => sum + group.sourceRows.length, 0),
    quantity: groups.reduce((sum, group) => sum + group.quantity, 0),
    amount: groups.reduce((sum, group) => sum + group.amount, 0),
    cost: groups.reduce((sum, group) => sum + group.buckets.reduce((bucketSum, bucket) => bucketSum + bucket.costAmount, 0), 0),
  };
}

function caoThanhPriceReportPayload(originalName: string, groups: CaoThanhPriceGroup[]) {
  const summaryHeaders = ['Mã VT', 'Tên hàng', 'Công ty', 'ĐVT', 'Số dòng', 'SL', 'Giá min', 'Giá max', '% lọc', 'Nhóm', 'Mã VT xuất', 'Giá TB', '% lãi', 'Giá vốn dự tính', 'Tiền hàng', 'Tiền vốn'];
  const detailHeaders = ['Mã VT xuất', 'Mã VT gốc', 'Ngày HĐ', 'Số HĐ', 'Công ty', 'MST', 'Tên hàng', 'ĐVT', 'SL', 'Đơn giá', 'Thành tiền', 'Dòng Excel'];
  const summaryRows = groups.flatMap((group) => group.buckets.map((bucket) => ({
    'Mã VT': group.code,
    'Tên hàng': group.productNames,
    'Công ty': group.companies,
    'ĐVT': group.unit,
    'Số dòng': bucket.count,
    'SL': bucket.quantity,
    'Giá min': bucket.min,
    'Giá max': bucket.max,
    '% lọc': group.filterPercent,
    'Nhóm': bucket.label,
    'Mã VT xuất': bucket.finalCode,
    'Giá TB': bucket.averagePrice,
    '% lãi': bucket.marginPercent,
    'Giá vốn dự tính': bucket.costUnitPrice,
    'Tiền hàng': bucket.amount,
    'Tiền vốn': bucket.costAmount,
  })));
  const detailRows = groups.flatMap((group) => group.buckets.flatMap((bucket) => bucket.rows.map((row) => ({
    'Mã VT xuất': bucket.finalCode,
    'Mã VT gốc': group.code,
    'Ngày HĐ': row.invoiceDate,
    'Số HĐ': row.invoiceNo,
    'Công ty': row.company,
    'MST': row.mst,
    'Tên hàng': row.productName,
    'ĐVT': row.unit,
    'SL': row.quantity,
    'Đơn giá': row.price,
    'Thành tiền': row.amount,
    'Dòng Excel': row.excelRow,
  }))));
  return {
    filename: `${fileStem(originalName)}_bao_cao_loc_gia.xls`,
    sheets: [
      { name: 'Tong hop loc gia', headers: summaryHeaders, rows: summaryRows },
      { name: 'Chi tiet gia', headers: detailHeaders, rows: detailRows },
    ],
  };
}

function caoThanhReportFileName(originalName: string) {
  return `${fileStem(originalName)}_bao_cao_loc_gia.xls`;
}

function buildCaoThanhProcessPayload(
  summary: UploadSummary,
  columns: CaoThanhColumns,
  companies: CompanyRow[],
  manualOverrides: Record<string, string>,
  productCodeReplacements: Record<string, string>,
  includeCompanyPrefix: boolean,
  wordRules: Record<string, string>,
  firstWordRules: Record<string, string>,
  repeatedPhrases: string[],
  priceRangeRules: Record<string, any>,
  priceAdjustAllPercent: number,
) {
  const activeCompanies = companies.filter((company) => company.process !== false);
  return {
    saved_name: summary.saved_name,
    original_name: summary.original_name,
    profile: 'cao_thanh',
    ...columns,
    include_company_prefix: includeCompanyPrefix,
    prefix_strategy: 'last_2_words',
    prefix_mst_digits: 3,
    prefix_strategy_values: emptyPrefixStrategyValues(),
    word_rules: wordRules,
    first_word_rules: firstWordRules,
    repeated_phrase_removals: repeatedPhrases.filter((phrase) => phrase.trim()),
    manual_code_overrides: manualOverrides,
    product_code_replacements: productCodeReplacements,
    price_range_rules: priceRangeRules,
    price_adjust_all_percent: priceAdjustAllPercent,
    prefixes: companyPrefixes(companies),
    removed_companies: Object.fromEntries(companies.filter((company) => company.process === false).map((company) => [companyConfigKey(company), true])),
    skipped_products_map: Object.fromEntries(companies.map((company) => {
      const selected = new Set(selectedProductNames(company));
      const skipped = company.all_products.map((product) => product.name).filter((name) => !selected.has(name));
      return [companyConfigKey(company), skipped];
    }).filter(([, skipped]) => Array.isArray(skipped) && skipped.length)),
    all_mst: companies.map((company) => companyConfigKey(company)),
    process_mst: activeCompanies.map((company) => companyConfigKey(company)),
    mst_safe_id: companies.map((company, index) => `${companyConfigKey(company)}|||${index}`),
    ...companyPrefixFields(companies),
    ...Object.fromEntries(companies.flatMap((company, index) => (company.process === false ? [] : [[`selected_products_${index}`, selectedProductNames(company)]]))),
  };
}
function PreviewPanel({ summary }: { summary: UploadSummary | null }) {
  if (!summary) return <PlaceholderStage title="Chưa có file" detail="Tải file bán ra trước." />;
  const previewKeys = summary.preview.length ? Object.keys(summary.preview[0]) : [];
  return <div className="preview-panel"><h3>Xem trước dữ liệu</h3><div className="preview-scroll"><table><thead><tr>{previewKeys.map((key) => <th key={key}>{key}</th>)}</tr></thead><tbody>{summary.preview.map((row, index) => <tr key={index}>{previewKeys.map((key) => <td key={key}>{row[key]}</td>)}</tr>)}</tbody></table></div></div>;
}

function ReviewTable({ groups, collapsedGroups, onToggleGroup, onRowChange, reviewScope }: { groups: ReviewDisplayGroup[]; collapsedGroups: Record<string, boolean>; onToggleGroup: (groupTitle: string) => void; onRowChange?: (index: number, update: Partial<ReviewRow>) => void; reviewScope?: 'all' | 'company' }) {
  return (
    <table className="review-table grouped-review-table">
      <thead>
        <tr>
          <th className="review-use-column">Dùng</th>
          <th className="review-choice-column">Dùng mã</th>
          <th className="review-product-column">Tên hàng hóa 1</th>
          <th className="review-code-column">Mã VT 1</th>
          <th>Số HD 1</th>
          <th>Ngày HD 1</th>
          <th>ĐVT 1</th>
          <th className="review-product-column">Tên hàng hóa 2</th>
          <th className="review-code-column">Mã VT 2</th>
          <th>Số HD 2</th>
          <th>Ngày HD 2</th>
          <th>ĐVT 2</th>
          <th className="review-code-diff-column">Khác biệt mã</th>
          <th>Công ty</th>
          <th>Độ giống</th>
        </tr>
      </thead>
      <tbody>{groups.map((group) => <ReviewGroupRows key={group.key || group.title} group={group} collapsedGroups={collapsedGroups} onToggleGroup={onToggleGroup} onRowChange={onRowChange} />)}</tbody>
    </table>
  );
}

function reviewDisplayGroupCount(group: ReviewDisplayGroup): number {
  return group.rows.length + (group.children || []).reduce((total, child) => total + reviewDisplayGroupCount(child), 0);
}

function reviewDisplayGroupIndices(group: ReviewDisplayGroup): number[] {
  return [
    ...group.rows.map(({ index }) => index),
    ...(group.children || []).flatMap((child) => reviewDisplayGroupIndices(child)),
  ];
}

function ReviewGroupRows({ group, collapsedGroups, onToggleGroup, onRowChange }: { group: ReviewDisplayGroup; collapsedGroups: Record<string, boolean>; onToggleGroup: (groupTitle: string) => void; onRowChange?: (index: number, update: Partial<ReviewRow>) => void }) {
  const totalRows = reviewDisplayGroupCount(group);
  const groupKey = group.key || group.title;
  const collapsed = Boolean(collapsedGroups[groupKey]);
  const handleCodeChoiceChange = (index: number, row: ReviewRow, choice: string) => {
    const update: Partial<ReviewRow> = { code_choice: choice };
    if (choice === 'split') {
      update.split_code = sanitizeDisplayProductCode(row.split_code || row.code || '');
      update.similar_split_code = sanitizeDisplayProductCode(row.similar_split_code || row.similar_code || '');
    }
    onRowChange?.(index, update);
  };
  return <>{<tr className={`review-section-row ${group.className}`}><td colSpan={15}><button type="button" className="review-section-toggle" aria-expanded={!collapsed} onClick={() => onToggleGroup(groupKey)}><span className="review-section-caret" aria-hidden="true">{collapsed ? '+' : '-'}</span><span className="review-section-label">{group.displayTitle || group.title}</span><span className="review-section-count">{totalRows}</span></button></td></tr>}{!collapsed && group.rows.map(({ row, index }) => (
        <tr key={`${row.product}-${row.similar_product}-${index}`}>
          <td className="review-use-cell"><input className="table-checkbox" type="checkbox" checked={row.confirmed === true} disabled={!onRowChange} onChange={(event) => onRowChange?.(index, { confirmed: event.currentTarget.checked })} /></td>
          <td className="review-choice-cell"><select className="code-choice" value={row.code_choice || 'current'} disabled={!onRowChange} onChange={(event) => handleCodeChoiceChange(index, row, event.currentTarget.value)}><option value="current">Mã VT 1</option><option value="similar">Mã VT 2</option>{row.review_type === 'same_code_split' && <option value="split">Tách mã</option>}</select></td>
          <td className="review-product-cell">{row.product}</td>
          <td className="code-cell review-code-cell"><CodeValue label="Mã 1" value={row.code} />{row.review_type === 'same_code_split' && row.code_choice === 'split' && <SplitCodeInput label="Mã tách 1" value={row.split_code || row.code || ''} disabled={!onRowChange} onChange={(value) => onRowChange?.(index, { split_code: sanitizeDisplayProductCode(value) })} />}</td>
          <td>{row.invoice_no}</td>
          <td>{row.invoice_date}</td>
          <td>{row.unit}</td>
          <td className="review-product-cell">{row.similar_product}</td>
          <td className="code-cell review-code-cell"><CodeValue label="Mã 2" value={row.similar_code} />{row.review_type === 'same_code_split' && row.code_choice === 'split' && <SplitCodeInput label="Mã tách 2" value={row.similar_split_code || row.similar_code || ''} disabled={!onRowChange} onChange={(value) => onRowChange?.(index, { similar_split_code: sanitizeDisplayProductCode(value) })} />}</td>
          <td>{row.similar_invoice_no}</td>
          <td>{row.similar_invoice_date}</td>
          <td>{row.similar_unit}</td>
          <td className="review-code-diff-cell"><CodeDiff current={row.code} target={row.similar_code} /></td>
          <td className="review-company-cell">{row.company || row.similar_company}</td>
          <td>{formatSimilarity(row.similarity)}</td>
        </tr>
      ))}{!collapsed && group.children?.map((child) => <ReviewGroupRows key={child.key || child.title} group={child} collapsedGroups={collapsedGroups} onToggleGroup={onToggleGroup} onRowChange={onRowChange} />)}</>;
}

function SplitCodeInput({ label, value, disabled, onChange }: { label: string; value: string; disabled: boolean; onChange: (value: string) => void }) {
  return (
    <label className="split-code-input">
      <span>{label}</span>
      <input value={value} disabled={disabled} onChange={(event) => onChange(event.currentTarget.value)} />
    </label>
  );
}

function CodeValue({ label, value }: { label: string; value?: string }) {
  return <span className="review-code-stack"><span className="review-code-label">{label}</span><span className="review-code-value">{value || 'Chưa có mã'}</span></span>;
}

function CodeDiff({ current, target }: { current?: string; target?: string }) {
  const currentCode = String(current || '').trim();
  const targetCode = String(target || '').trim();
  const changed = currentCode !== targetCode;
  const pieces = changed ? characterDiffPieces(currentCode, targetCode) : [{ text: 'Hai mã giống nhau', kind: 'neutral' as const }];
  return (
    <div className="code-diff-stack" aria-label={`So sánh Mã VT 1 ${currentCode || 'trống'} với Mã VT 2 ${targetCode || 'trống'}`}>
      <div className="code-compare-pair">
        <span className="code-compare-chip"><strong>Mã 1</strong>{currentCode || 'Trống'}</span>
        <span className="code-compare-chip"><strong>Mã 2</strong>{targetCode || 'Trống'}</span>
      </div>
      <div className={`diff-box code-diff-box ${changed ? '' : 'is-same'}`}>{pieces.map((piece, index) => <span key={index} className={`diff-${piece.kind}`}>{piece.text}</span>)}</div>
    </div>
  );
}

function reviewDisplayGroups(rows: ReviewRow[], reviewScope: 'all' | 'company' = 'all'): ReviewDisplayGroup[] {
  const indexedRows = rows.map((row, index) => ({ row, index }));
  const sectionDefinitions = [
    {
      key: 'same_code_split',
      title: 'M\u00e3 VT gi\u1ed1ng nhau - c\u00f3 th\u1ec3 t\u00e1ch m\u00e3',
      className: 'same-code-split-section',
      filter: ({ row }: { row: ReviewRow }) => row.review_group === 'same_code_split' || row.review_type === 'same_code_split',
    },
    {
      key: 'similar_form',
      title: 'C\u00f9ng form t\u00ean h\u00e0ng',
      className: 'similar-form-section',
      filter: ({ row }: { row: ReviewRow }) => row.review_group === 'similar_form',
    },
    {
      key: 'unit_spelling_diff',
      title: 'Kh\u00e1c c\u00e1ch vi\u1ebft \u0111\u01a1n v\u1ecb',
      className: 'unit-spelling-section',
      filter: ({ row }: { row: ReviewRow }) => row.review_group === 'unit_spelling_diff' || row.review_type === 'unit_spelling_diff',
    },
    {
      key: 'other',
      title: 'Kh\u00e1c',
      className: 'other-section',
      filter: ({ row }: { row: ReviewRow }) => row.review_group !== 'dimension_diff' && row.review_group !== 'similar_form' && row.review_group !== 'same_code_split' && row.review_group !== 'unit_spelling_diff' && row.review_type !== 'same_code_split' && row.review_type !== 'unit_spelling_diff' && row.dimension_only !== true,
    },
    {
      key: 'dimension_diff',
      title: 'G\u1ea7n gi\u1ed1ng k\u00edch th\u01b0\u1edbc',
      className: 'dimension-section',
      filter: ({ row }: { row: ReviewRow }) => row.review_group === 'dimension_diff' || row.dimension_only === true,
    },
  ];
  
  if (reviewScope === 'company') {
    const result: ReviewDisplayGroup[] = [];
    for (const section of sectionDefinitions) {
      const sectionRows = indexedRows.filter(section.filter);
      const companyGroups = new Map<string, typeof indexedRows>();
      sectionRows.forEach(({ row, index }) => {
        const companyName = row.company || row.similar_company || 'Không xác định';
        const companyKey = row.company_key || row.similar_company_key || row.mst || row.similar_mst || companyName;
        const groupKey = `${companyKey}|||${companyName}`;
        if (!companyGroups.has(groupKey)) companyGroups.set(groupKey, []);
        companyGroups.get(groupKey)!.push({ row, index });
      });
      result.push({
        title: section.title,
        key: `section|||${section.key}`,
        className: section.className,
        rows: [],
        children: Array.from(companyGroups.entries()).map(([companyKey, companyRows]) => {
          const companyName = companyRows[0]?.row.company || companyRows[0]?.row.similar_company || 'Không xác định';
          return {
            title: companyName,
            displayTitle: companyName,
            key: `section|||${section.key}|||company|||${companyKey}`,
            className: 'review-company-subsection',
            rows: companyRows,
          };
        }),
      });
    }
    
    return result;
  }
  
  return sectionDefinitions.map((section) => ({ title: section.title, key: `section|||${section.key}`, className: section.className, rows: indexedRows.filter(section.filter) }));
}

function MatchTable({ rows, onToggle, onConversionChange }: { rows: MatchRow[]; onToggle: (index: number, confirmed: boolean) => void; onConversionChange: (index: number, salesQty: string, purchaseQty: string) => void }) {
  const indexedRows = rows.map((row, index) => ({ row, index, mismatch: hasUnitMismatch(row) }));
  const mismatchRows = indexedRows.filter((item) => item.mismatch);
  const normalRows = indexedRows.filter((item) => !item.mismatch);
  const groups = [
    { title: `Khác ĐVT cần kiểm tra (${mismatchRows.length})`, className: 'unit-mismatch-section', rows: mismatchRows },
    { title: `Cùng ĐVT / đã khớp (${normalRows.length})`, className: 'match-normal-section', rows: normalRows },
  ].filter((group) => group.rows.length);
  return (
    <table className="match-table">
      <thead>
        <tr><th>Dùng</th><th>Hàng bán ra</th><th>ĐVT bán</th><th>Đơn giá bán</th><th>Mã VT mua vào</th><th>Hàng mua vào</th><th>ĐVT mua</th><th>Đơn giá mua</th><th>Cảnh báo</th><th>Công thức quy đổi</th><th>Khác biệt</th><th>Độ giống</th></tr>
      </thead>
      <tbody>{groups.map((group) => (
        <Fragment key={group.title}>
          <tr className={`match-section-row ${group.className}`}><td colSpan={12}>{group.title}</td></tr>
          {group.rows.map(({ row, index, mismatch }) => {
        const quantities = conversionQuantities(row);
        return (
          <tr key={`${row.sales_product}-${row.purchase_product}-${index}`} className={mismatch ? 'unit-mismatch-row' : ''}>
            <td><input className="table-checkbox" type="checkbox" checked={row.confirmed !== false} onChange={(event) => onToggle(index, event.currentTarget.checked)} /></td>
            <td>{row.sales_product}</td>
            <td className={mismatch ? 'unit-mismatch-cell' : ''}>{row.sales_unit}</td>
            <td>{row.sales_price}</td>
            <td>{row.purchase_code}</td>
            <td>{row.purchase_product}</td>
            <td className={mismatch ? 'unit-mismatch-cell' : ''}>{row.purchase_unit}</td>
            <td>{row.purchase_price}</td>
            <td className={mismatch ? 'unit-mismatch-cell' : ''}>{row.unit_warning || row.warning}</td>
            <td>{mismatch ? <ConversionFormulaInputs row={row} index={index} salesQty={quantities.salesQty} purchaseQty={quantities.purchaseQty} onChange={onConversionChange} /> : <span className="muted">Cùng ĐVT</span>}</td>
            <td><NameDiff current={row.sales_product} target={row.purchase_product} /></td>
            <td>{formatSimilarity(row.similarity ?? row.score)}</td>
          </tr>
        );
      })}
        </Fragment>
      ))}</tbody>
    </table>
  );
}

function ConversionFormulaInputs({ row, index, salesQty, purchaseQty, onChange }: { row: MatchRow; index: number; salesQty: string; purchaseQty: string; onChange: (index: number, salesQty: string, purchaseQty: string) => void }) {
  const [draft, setDraft] = useState({ salesQty, purchaseQty });
  useEffect(() => setDraft({ salesQty, purchaseQty }), [salesQty, purchaseQty]);
  const updateDraft = (side: 'sales' | 'purchase', value: string) => {
    const nextDraft = side === 'sales' ? { ...draft, salesQty: value } : { ...draft, purchaseQty: value };
    setDraft(nextDraft);
    onChange(index, nextDraft.salesQty, nextDraft.purchaseQty);
  };
  return (
    <div className="conversion-stack">
      <label><input value={draft.salesQty} inputMode="decimal" onChange={(event) => updateDraft('sales', event.currentTarget.value)} /><span>{row.sales_unit || 'ĐVT bán'}</span></label>
      <label><input value={draft.purchaseQty} inputMode="decimal" onChange={(event) => updateDraft('purchase', event.currentTarget.value)} /><span>{row.purchase_unit || 'ĐVT mua'}</span></label>
    </div>
  );
}

function hasUnitMismatch(row: MatchRow) {
  const salesUnit = String(row.sales_unit || '').trim().toLocaleLowerCase('vi-VN');
  const purchaseUnit = String(row.purchase_unit || '').trim().toLocaleLowerCase('vi-VN');
  return Boolean(salesUnit && purchaseUnit && salesUnit !== purchaseUnit);
}

function conversionQuantities(row: MatchRow) {
  const [salesSide = '', purchaseSide = ''] = String(row.conversion_formula || '').split('=');
  return {
    salesQty: quantityWithoutUnit(salesSide, row.sales_unit),
    purchaseQty: quantityWithoutUnit(purchaseSide, row.purchase_unit),
  };
}

function quantityWithoutUnit(value: string, unit: string | undefined) {
  const text = value.trim();
  const unitText = String(unit || '').trim();
  if (unitText && text.toLocaleLowerCase('vi-VN').endsWith(unitText.toLocaleLowerCase('vi-VN'))) {
    return text.slice(0, Math.max(0, text.length - unitText.length)).trim();
  }
  return text.split(/\s+/)[0] || '';
}

function conversionFormula(row: MatchRow, salesQty: string, purchaseQty: string) {
  return `${salesQty.trim()} ${row.sales_unit || ''}`.trim() + ' = ' + `${purchaseQty.trim()} ${row.purchase_unit || ''}`.trim();
}

function buildSalesMatchRules(workflow: WorkflowState): MatchRow[] {
  const rules = new Map<string, MatchRow>();
  for (const rule of workflow.salesMatchRules || []) {
    const key = salesMatchRuleKey(rule, workflow.comparisonScope);
    if (key) rules.set(key, rule);
  }
  for (const match of workflow.matches || []) {
    const key = salesMatchRuleKey(match, workflow.comparisonScope);
    if (!key) continue;
    rules.set(key, {
      confirmed: match.confirmed !== false,
      sales_product: match.sales_product,
      sales_unit: match.sales_unit,
      sales_company: match.sales_company,
      sales_mst: match.sales_mst,
      purchase_code: match.purchase_code,
      purchase_product: match.purchase_product,
      purchase_unit: match.purchase_unit,
      purchase_company: match.purchase_company,
      purchase_mst: match.purchase_mst,
      conversion_mode: match.conversion_mode || 'none',
      conversion_formula: match.conversion_formula || '',
    });
  }
  return Array.from(rules.values());
}

function applySalesMatchRules(matches: MatchRow[], rules: MatchRow[], comparisonScope: string): MatchRow[] {
  if (!rules.length) return matches;
  const ruleMap = new Map(rules.map((rule) => [salesMatchRuleKey(rule, comparisonScope), rule]).filter(([key]) => Boolean(key)) as Array<[string, MatchRow]>);
  return matches.map((match) => {
    const rule = ruleMap.get(salesMatchRuleKey(match, comparisonScope));
    if (!rule) return match;
    const sameUnitPair = simpleMatchText(rule.sales_unit) === simpleMatchText(match.sales_unit) && simpleMatchText(rule.purchase_unit) === simpleMatchText(match.purchase_unit);
    return {
      ...match,
      confirmed: rule.confirmed !== false,
      conversion_mode: sameUnitPair ? (rule.conversion_mode || match.conversion_mode || 'none') : match.conversion_mode,
      conversion_formula: sameUnitPair ? (rule.conversion_formula || '') : match.conversion_formula,
    };
  });
}

function salesMatchRuleKey(row: MatchRow, _comparisonScope: string) {
  const companyKey = simpleMatchText(row.sales_mst || row.sales_company);
  const salesProduct = simpleMatchText(row.sales_product);
  const purchaseProduct = simpleMatchText(row.purchase_product);
  if (!companyKey || !salesProduct || !purchaseProduct) return '';
  return `${companyKey}|||${salesProduct}|||${purchaseProduct}`;
}

function NameDiff({ current, target }: { current?: string; target?: string }) {
  const pieces = nameDiffPieces(current, target);
  if (!pieces.length) return null;
  const changed = pieces.some((piece) => piece.kind !== 'neutral');
  if (!changed) return <span className="match-same-chip">Trùng khớp</span>;
  return <div className="diff-box">{pieces.map((piece, index) => <span key={index} className={`diff-${piece.kind}`}>{piece.text}</span>)}</div>;
}

type DiffPiece = { text: string; kind: 'neutral' | 'removed-word' | 'added-word' | 'removed-char' | 'added-char' | 'arrow' };

function nameDiffPieces(current: string | undefined, target: string | undefined): DiffPiece[] {
  const currentWords = splitWords(current);
  const targetWords = splitWords(target);
  const pieces: DiffPiece[] = [];
  for (const opcode of sequenceOpcodes(currentWords, targetWords, (left, right) => left === right)) {
    const currentSlice = currentWords.slice(opcode.i1, opcode.i2);
    const targetSlice = targetWords.slice(opcode.j1, opcode.j2);
    if (opcode.tag === 'equal') pieces.push({ text: currentSlice.join(' '), kind: 'neutral' });
    else if (opcode.tag === 'delete') pieces.push(...currentSlice.map((word) => ({ text: `- ${word}`, kind: 'removed-word' as const })));
    else if (opcode.tag === 'insert') pieces.push(...targetSlice.map((word) => ({ text: `+ ${word}`, kind: 'added-word' as const })));
    else pieces.push(...replacementDiffPieces(currentSlice, targetSlice));
  }
  return pieces.filter((piece) => piece.text);
}

function replacementDiffPieces(currentWords: string[], targetWords: string[]): DiffPiece[] {
  if (currentWords.length === targetWords.length) {
    return currentWords.flatMap((word, index) => wordsAreSmallCharacterDiff(word, targetWords[index]) ? characterDiffPieces(word, targetWords[index]) : [{ text: `- ${word}`, kind: 'removed-word' as const }, { text: `+ ${targetWords[index]}`, kind: 'added-word' as const }]);
  }
  return [...currentWords.map((word) => ({ text: `- ${word}`, kind: 'removed-word' as const })), ...targetWords.map((word) => ({ text: `+ ${word}`, kind: 'added-word' as const }))];
}

function characterDiffPieces(current: string, target: string): DiffPiece[] {
  const currentPieces: DiffPiece[] = [];
  const targetPieces: DiffPiece[] = [];
  for (const opcode of sequenceOpcodes([...current], [...target], (left, right) => left === right)) {
    const currentText = current.slice(opcode.i1, opcode.i2);
    const targetText = target.slice(opcode.j1, opcode.j2);
    if (opcode.tag === 'equal') {
      currentPieces.push({ text: currentText, kind: 'neutral' });
      targetPieces.push({ text: targetText, kind: 'neutral' });
    } else if (opcode.tag === 'delete') currentPieces.push({ text: currentText, kind: 'removed-char' });
    else if (opcode.tag === 'insert') targetPieces.push({ text: targetText, kind: 'added-char' });
    else {
      currentPieces.push({ text: currentText, kind: 'removed-char' });
      targetPieces.push({ text: targetText, kind: 'added-char' });
    }
  }
  return [...currentPieces, { text: '→', kind: 'arrow' }, ...targetPieces];
}

function normalizeReviewRows(rows: ReviewRow[]) {
  return rows.map((row) => ({ ...row, confirmed: row.confirmed === true, code_choice: row.code_choice || 'current' }));
}

async function loadProductPreviewCodes(companies: CompanyRow[], wordRules: Record<string, string>, repeatedPhraseRemovals: string[], phase: 'purchase' | 'sales' = 'purchase', productCodeReplacements: Record<string, string> = {}) {
  const products = Array.from(new Set(companies.flatMap((company) => company.all_products.map((product) => product.name)).filter(Boolean)));
  if (!products.length) return {};
  const result = await previewVietmaxProductCodes(products, wordRules, repeatedPhraseRemovals, phase, productCodeReplacements);
  return result.codes;
}

function normalizeProductCodeReplacements(raw: Record<string, string> = {}): Record<string, string> {
  return Object.fromEntries(
    Object.entries(raw)
      .map(([oldCode, newCode]) => [sanitizeDisplayProductCode(oldCode), sanitizeDisplayProductCode(newCode)])
      .filter(([oldCode, newCode]) => oldCode && newCode),
  );
}

function applyDisplayProductCodeReplacement(code: string, replacements: Record<string, string> = {}) {
  const cleanCode = sanitizeDisplayProductCode(code);
  const normalized = normalizeProductCodeReplacements(replacements);
  if (!cleanCode || !Object.keys(normalized).length) return cleanCode;
  if (normalized[cleanCode]) return normalized[cleanCode];
  const lastDot = cleanCode.lastIndexOf('.');
  if (lastDot >= 0) {
    const prefix = cleanCode.slice(0, lastDot);
    const body = cleanCode.slice(lastDot + 1);
    if (normalized[body]) return sanitizeDisplayProductCode(`${prefix}.${normalized[body]}`);
  }
  return cleanCode;
}

async function loadGenericProductPreviewCodes(profile: ProfileKey, companies: CompanyRow[], wordRules: Record<string, string>, firstWordRules: Record<string, string>, repeatedPhraseRemovals: string[], productCodeReplacements: Record<string, string> = {}) {
  const products = Array.from(new Set(companies.flatMap((company) => company.all_products.map((product) => product.name)).filter(Boolean)));
  if (!products.length) return {};
  const result = await previewGenericProductCodes({ profile, products, word_rules: wordRules, first_word_rules: firstWordRules, repeated_phrase_removals: repeatedPhraseRemovals, product_code_replacements: productCodeReplacements });
  return result.codes;
}

function buildPurchaseProcessPayload(workflow: WorkflowState) {
  const companies = workflow.companyRows;
  const activeCompanies = companies.filter(isMaterialCompany);
  const activePrefixStrategy = normalizedPrefixStrategy(workflow.purchasePrefixStrategy);
  const purchasePrefixStrategyValues = rememberManualPrefixValues(workflow.purchasePrefixStrategyValues, activePrefixStrategy, companies, workflow.prefixMstDigits, workflow.prefixNameWords, workflow.prefixNameChars);
  const reviewScope = reviewScopeValue(workflow.purchaseReviewScope);
  const columns = normalizeVietmaxColumns(workflow.purchaseColumns, 'purchase');
  return {
    profile: 'vietmax',
    vietmax_phase: 'purchase',
    ...columns,
    purchase_price_col: columns.price_col || 'P',
    include_company_prefix: workflow.includeCompanyPrefix,
    prefix_strategy: activePrefixStrategy,
    prefix_mst_digits: workflow.prefixMstDigits,
    prefix_name_words: workflow.prefixNameWords,
    prefix_name_chars: workflow.prefixNameChars,
    prefix_missing_mst_strategy: normalizeMissingMstPrefixStrategy(workflow.prefixMissingMstStrategy),
    prefix_strategy_values: purchasePrefixStrategyValues,
    comparison_scope: workflow.comparisonScope,
    word_rules: workflow.purchaseWordRules,
    repeated_phrase_removals: workflow.purchaseRepeatedPhraseRemovals.filter((phrase) => phrase.trim()),
    manual_code_overrides: workflow.productCodeOverrides,
    product_code_replacements: workflow.productCodeReplacements,
    vietmax_mua_vao_internal_merges: buildReviewRules(workflow.purchaseReviewRules, workflow.purchaseReviewRows, reviewScope),
    inventory_pairs: workflow.purchaseInventoryPairs.filter((pair) => pair.ma_kho.trim() || pair.tk_vat_tu.trim()),
    use_default_inventory_pair: workflow.purchaseUseDefaultInventoryPair,
    default_inventory_pair_id: workflow.purchaseDefaultInventoryPairId,
    inventory_pair_rules: inventoryRulesForPayload(workflow.purchaseInventoryPairRules, workflow.purchaseInventoryPairs),
    inventory_allocation_config: workflow.inventoryAllocationConfig,
    company_group_assignments: groupAssignmentsFromRows(companies),
    prefixes: companyPrefixes(companies),
    all_mst: companies.map((company) => companyConfigKey(company)),
    process_mst: activeCompanies.map((company) => companyConfigKey(company)),
    mst_safe_id: companies.map((company, index) => `${companyConfigKey(company)}|||${index}`),
    ...companyPrefixFields(companies),
    ...Object.fromEntries(companies.flatMap((company, index) => (!isMaterialCompany(company) ? [] : [[`selected_products_${index}`, selectedProductNames(company)]]))),
  };
}

function buildSalesProcessPayload(workflow: WorkflowState) {
  const companies = workflow.salesCompanyRows;
  const activeCompanies = companies.filter(isMaterialCompany);
  const activePrefixStrategy = normalizedPrefixStrategy(workflow.salesPrefixStrategy);
  const salesPrefixStrategyValues = rememberManualPrefixValues(workflow.salesPrefixStrategyValues, activePrefixStrategy, companies, workflow.prefixMstDigits, workflow.prefixNameWords, workflow.prefixNameChars);
  const reviewScope = reviewScopeValue(workflow.salesReviewScope);
  const columns = normalizeVietmaxColumns(workflow.salesColumns, 'sales');
  return {
    profile: 'vietmax',
    vietmax_phase: 'sales',
    ...columns,
    purchase_price_col: workflow.purchaseColumns.price_col || 'P',
    include_company_prefix: workflow.salesIncludeCompanyPrefix,
    prefix_strategy: activePrefixStrategy,
    prefix_mst_digits: workflow.prefixMstDigits,
    prefix_name_words: workflow.prefixNameWords,
    prefix_name_chars: workflow.prefixNameChars,
    prefix_missing_mst_strategy: normalizeMissingMstPrefixStrategy(workflow.prefixMissingMstStrategy),
    prefix_strategy_values: salesPrefixStrategyValues,
    comparison_scope: workflow.comparisonScope,
    word_rules: workflow.salesWordRules,
    repeated_phrase_removals: workflow.salesRepeatedPhraseRemovals.filter((phrase) => phrase.trim()),
    manual_code_overrides: workflow.salesProductCodeOverrides,
    product_code_replacements: workflow.salesProductCodeReplacements,
    inventory_pairs: workflow.salesInventoryPairs.filter((pair) => pair.ma_kho.trim() || pair.tk_vat_tu.trim()),
    use_default_inventory_pair: workflow.salesUseDefaultInventoryPair,
    default_inventory_pair_id: workflow.salesDefaultInventoryPairId,
    inventory_pair_rules: inventoryRulesForPayload(workflow.salesInventoryPairRules, workflow.salesInventoryPairs),
    inventory_allocation_config: workflow.inventoryAllocationConfig,
    vietmax_processed_purchase_saved_name: workflow.processedPurchaseSavedName,
    vietmax_ban_ra_purchase_matches: workflow.matches.filter((match) => match.confirmed !== false),
    vietmax_ban_ra_purchase_match_rules: buildSalesMatchRules(workflow),
    vietmax_ban_ra_sales_internal_merges: buildReviewRules(workflow.salesReviewRules, workflow.salesReviewRows, reviewScope),
    company_group_assignments: groupAssignmentsFromRows(companies),
    prefixes: companyPrefixes(companies),
    all_mst: companies.map((company) => companyConfigKey(company)),
    process_mst: activeCompanies.map((company) => companyConfigKey(company)),
    mst_safe_id: companies.map((company, index) => `${companyConfigKey(company)}|||${index}`),
    ...companyPrefixFields(companies),
    ...Object.fromEntries(companies.flatMap((company, index) => (!isMaterialCompany(company) ? [] : [[`selected_products_${index}`, selectedProductNames(company)]]))),
  };
}

function buildPurchaseReviewProducts(workflow: WorkflowState): ReviewProduct[] {
  return buildReviewProducts(workflow.companyRows, 'purchase', workflow.productCodeOverrides);
}

function buildSalesReviewProducts(workflow: WorkflowState): ReviewProduct[] {
  return buildReviewProducts(workflow.salesCompanyRows, 'sales', workflow.salesProductCodeOverrides);
}

function buildGenericProcessPayload(workflow: WorkflowState, profile: ProfileKey) {
  const companies = workflow.companyRows;
  const activeCompanies = profile === 'ho_guom'
    ? companies.filter((company) => !isIgnoredCompany(company))
    : companies.filter(isMaterialCompany);
  const twoPhase = isTwoPhaseGenericProfile(profile);
  const activePrefixStrategy = normalizedPrefixStrategy(workflow.purchasePrefixStrategy);
  const prefixStrategyValues = rememberManualPrefixValues(workflow.purchasePrefixStrategyValues, activePrefixStrategy, companies, workflow.prefixMstDigits, workflow.prefixNameWords, workflow.prefixNameChars);
  const columns = twoPhase ? normalizeVietmaxColumns(workflow.purchaseColumns, 'purchase') : normalizeGenericColumns(workflow.genericColumns);
  const reviewScope = reviewScopeValue(workflow.purchaseReviewScope);
  const priceRangeRules = profile === 'cao_thanh'
    ? { ...workflow.priceRangeRules, ...caoThanhRangeRules(workflow.priceGroups) }
    : {};
  const wordRules = twoPhase ? workflow.purchaseWordRules : workflow.wordRules;
  const repeatedPhrases = twoPhase ? workflow.purchaseRepeatedPhraseRemovals : workflow.repeatedPhraseRemovals;
  const manualOverrides = twoPhase ? workflow.productCodeOverrides : workflow.productCodeOverrides;
  const productReplacements = twoPhase ? workflow.productCodeReplacements : workflow.productCodeReplacements;
  const inventoryPairs = twoPhase ? workflow.purchaseInventoryPairs : workflow.inventoryPairs;
  const inventoryRules = twoPhase ? workflow.purchaseInventoryPairRules : workflow.inventoryPairRules;
  return {
    profile,
    vietmax_phase: 'purchase',
    ...columns,
    include_company_prefix: workflow.includeCompanyPrefix,
    prefix_strategy: activePrefixStrategy,
    prefix_mst_digits: workflow.prefixMstDigits,
    prefix_name_words: workflow.prefixNameWords,
    prefix_name_chars: workflow.prefixNameChars,
    prefix_missing_mst_strategy: normalizeMissingMstPrefixStrategy(workflow.prefixMissingMstStrategy),
    prefix_strategy_values: prefixStrategyValues,
    word_rules: wordRules,
    first_word_rules: workflow.firstWordRules,
    repeated_phrase_removals: repeatedPhrases.filter((phrase) => phrase.trim()),
    manual_code_overrides: manualOverrides,
    product_code_replacements: productReplacements,
    product_review_merges: buildReviewRules(workflow.purchaseReviewRules, workflow.purchaseReviewRows, reviewScope),
    price_range_rules: priceRangeRules,
    price_adjust_all_percent: profile === 'cao_thanh' ? workflow.priceAdjustAllPercent : 0,
    inventory_pairs: inventoryPairs.filter((pair) => pair.ma_kho.trim() || pair.tk_vat_tu.trim()),
    use_default_inventory_pair: twoPhase ? workflow.purchaseUseDefaultInventoryPair : workflow.useDefaultInventoryPair,
    default_inventory_pair_id: twoPhase ? workflow.purchaseDefaultInventoryPairId : workflow.defaultInventoryPairId,
    inventory_pair_rules: inventoryRulesForPayload(inventoryRules, inventoryPairs),
    inventory_allocation_config: workflow.inventoryAllocationConfig,
    processing_groups: workflow.purchaseProcessingGroups,
    company_group_assignments: groupAssignmentsFromRows(companies),
    form_mapping_presets: normalizeFormsForSave(workflow.purchaseFormMappingPresets),
    prefixes: companyPrefixes(companies),
    all_mst: companies.map((company) => companyConfigKey(company)),
    process_mst: activeCompanies.map((company) => companyConfigKey(company)),
    mst_safe_id: companies.map((company, index) => `${companyConfigKey(company)}|||${index}`),
    ...companyPrefixFields(companies),
    ...Object.fromEntries(companies.flatMap((company, index) => {
      const usesProductCodes = profile === 'ho_guom' ? !isIgnoredCompany(company) : isMaterialCompany(company);
      return usesProductCodes ? [[`selected_products_${index}`, selectedProductNames(company)]] : [];
    })),
    columns,
    removed_companies: Object.fromEntries(companies.filter(isIgnoredCompany).map((company) => [companyConfigKey(company), true])),
    skipped_products_map: Object.fromEntries(companies.map((company) => {
      const selected = new Set(selectedProductNames(company));
      const skipped = company.all_products.map((product) => product.name).filter((name) => !selected.has(name));
      return [companyConfigKey(company), skipped];
    }).filter(([, skipped]) => Array.isArray(skipped) && skipped.length)),
  };
}

function buildGenericSalesProcessPayload(workflow: WorkflowState, profile: ProfileKey) {
  const companies = workflow.salesCompanyRows;
  const activeCompanies = companies.filter(isMaterialCompany);
  const activePrefixStrategy = normalizedPrefixStrategy(workflow.salesPrefixStrategy);
  const prefixStrategyValues = rememberManualPrefixValues(workflow.salesPrefixStrategyValues, activePrefixStrategy, companies, workflow.prefixMstDigits, workflow.prefixNameWords, workflow.prefixNameChars);
  const columns = normalizeVietmaxColumns(workflow.salesColumns, 'sales');
  const reviewScope = reviewScopeValue(workflow.salesReviewScope);
  return {
    profile,
    vietmax_phase: 'sales',
    ...columns,
    include_company_prefix: workflow.salesIncludeCompanyPrefix,
    prefix_strategy: activePrefixStrategy,
    prefix_mst_digits: workflow.prefixMstDigits,
    prefix_name_words: workflow.prefixNameWords,
    prefix_name_chars: workflow.prefixNameChars,
    prefix_missing_mst_strategy: normalizeMissingMstPrefixStrategy(workflow.prefixMissingMstStrategy),
    prefix_strategy_values: prefixStrategyValues,
    word_rules: workflow.salesWordRules,
    first_word_rules: workflow.firstWordRules,
    repeated_phrase_removals: workflow.salesRepeatedPhraseRemovals.filter((phrase) => phrase.trim()),
    manual_code_overrides: workflow.salesProductCodeOverrides,
    product_code_replacements: workflow.salesProductCodeReplacements,
    product_review_merges: buildReviewRules(workflow.salesReviewRules, workflow.salesReviewRows, reviewScope),
    inventory_pairs: workflow.salesInventoryPairs.filter((pair) => pair.ma_kho.trim() || pair.tk_vat_tu.trim()),
    use_default_inventory_pair: workflow.salesUseDefaultInventoryPair,
    default_inventory_pair_id: workflow.salesDefaultInventoryPairId,
    inventory_pair_rules: inventoryRulesForPayload(workflow.salesInventoryPairRules, workflow.salesInventoryPairs),
    inventory_allocation_config: workflow.inventoryAllocationConfig,
    processing_groups: workflow.salesProcessingGroups,
    company_group_assignments: groupAssignmentsFromRows(companies),
    form_mapping_presets: normalizeFormsForSave(workflow.salesFormMappingPresets),
    prefixes: companyPrefixes(companies),
    all_mst: companies.map((company) => companyConfigKey(company)),
    process_mst: activeCompanies.map((company) => companyConfigKey(company)),
    mst_safe_id: companies.map((company, index) => `${companyConfigKey(company)}|||${index}`),
    ...companyPrefixFields(companies),
    ...Object.fromEntries(companies.flatMap((company, index) => (!isMaterialCompany(company) ? [] : [[`selected_products_${index}`, selectedProductNames(company)]]))),
    columns,
    removed_companies: Object.fromEntries(companies.filter(isIgnoredCompany).map((company) => [companyConfigKey(company), true])),
    skipped_products_map: Object.fromEntries(companies.map((company) => {
      const selected = new Set(selectedProductNames(company));
      const skipped = company.all_products.map((product) => product.name).filter((name) => !selected.has(name));
      return [companyConfigKey(company), skipped];
    }).filter(([, skipped]) => Array.isArray(skipped) && skipped.length)),
  };
}

function inventoryRulesForPayload(rules: InventoryRule[], pairs: InventoryPair[] = []) {
  const validPairIds = new Set(pairs.map((pair) => pair.id).filter(Boolean));
  return rules
    .filter((rule) => rule.source_col.trim() && rule.pair_id.trim() && (!validPairIds.size || validPairIds.has(rule.pair_id)))
    .map((rule) => ({
      ...rule,
      source_col: rule.source_col.trim().toUpperCase(),
      value: rule.value ?? '',
      priority: Number.isFinite(Number(rule.priority)) ? Math.trunc(Number(rule.priority)) : 0,
      enabled: rule.enabled !== false,
    }));
}

function normalizeGenericColumns(raw: Record<string, unknown>): GenericColumns {
  const defaults = defaultGenericColumns();
  const letter = (key: keyof GenericColumns) => String(raw[key] ?? defaults[key] ?? '').trim().toUpperCase();
  const skipValues = raw.invoice_status_skip_values;
  return {
    company_col: letter('company_col') || defaults.company_col,
    mst_col: letter('mst_col') || defaults.mst_col,
    address_col: letter('address_col'),
    product_col: letter('product_col') || defaults.product_col,
    qty_col: letter('qty_col'),
    price_col: letter('price_col'),
    output_col: letter('output_col') || defaults.output_col,
    invoice_status_col: letter('invoice_status_col'),
    invoice_status_skip_values: Array.isArray(skipValues) ? skipValues.map((item) => String(item)) : defaults.invoice_status_skip_values,
  };
}

function buildReviewProducts(companies: CompanyRow[], phase: 'purchase' | 'sales', manualCodeOverrides: Record<string, string> = {}): ReviewProduct[] {
  return companies.flatMap((company, companyIndex) => {
    if (!isMaterialCompany(company)) return [];
    const selected = new Set(selectedProductNames(company));
    return company.all_products.flatMap((product, productIndex) => {
      if (!selected.has(product.name)) return [];
      const firstPriceRow = product.priceRows?.[0];
      const key = productKey(companyConfigKey(company), product.name);
      const manualCode = manualCodeOverrides[key] || '';
      const base = {
        invoice_no: firstPriceRow?.invoiceNo ?? '',
        invoice_date: firstPriceRow?.invoiceDate ?? '',
        product_key: key,
        company_index: companyIndex,
        product_index: productIndex,
        phase,
      };
      if (phase === 'sales') {
        return [{
          ...base,
          sales_product: product.name,
          sales_code: manualCode,
          sales_unit: firstPriceRow?.unit ?? '',
          sales_company: company.company,
          sales_mst: company.mst,
          sales_company_key: companyReviewIdentityKey(company),
        } as ReviewProduct];
      }
      return [{
        ...base,
        purchase_product: product.name,
        purchase_code: manualCode,
        purchase_unit: firstPriceRow?.unit ?? '',
        purchase_company: company.company,
        purchase_mst: company.mst,
        purchase_company_key: companyReviewIdentityKey(company),
      }];
    });
  });
}

function reviewScopeValue(scope: 'all' | 'company') {
  return scope === 'company' ? 'same_company' : 'all_companies';
}

function normalizedReviewComparisonScope(scope: string | undefined, fallback = 'all_companies') {
  return scope === 'same_company' || scope === 'all_companies' ? scope : fallback;
}

function buildReviewRules(existingRules: ReviewRow[], rows: ReviewRow[], comparisonScope: string) {
  const rules = new Map<string, ReviewRow>();
  for (const rule of existingRules || []) {
    if (rule.confirmed !== true) continue;
    const ruleScope = normalizedReviewComparisonScope(rule.comparison_scope, 'all_companies');
    const key = reviewRuleKey(rule, ruleScope);
    const compactRule = compactReviewRule(rule, ruleScope);
    if (key && compactRule) rules.set(key, compactRule);
  }
  for (const row of rows || []) {
    const rule = reviewRuleFromRow(row, comparisonScope);
    const key = reviewRuleKey(rule, comparisonScope);
    if (!key) continue;
    if (row.confirmed === true) {
      rules.set(key, rule);
    } else {
      rules.delete(key);
    }
  }
  return Array.from(rules.values());
}

function textForReviewRule(value: unknown) {
  return String(value ?? '').trim();
}

function reviewCodeChoice(value: unknown) {
  const choice = textForReviewRule(value);
  return choice === 'similar' || choice === 'split' ? choice : 'current';
}

function setReviewRuleText(rule: ReviewRow, key: keyof ReviewRow, value: unknown) {
  const text = textForReviewRule(value);
  if (text) (rule as Record<string, unknown>)[key] = text;
}

function compactReviewRule(row: ReviewRow, comparisonScope: string): ReviewRow | null {
  const product = textForReviewRule(row.product);
  const similarProduct = textForReviewRule(row.similar_product);
  if (!product || !similarProduct) return null;
  const codeChoice = reviewCodeChoice(row.code_choice);
  const rule: ReviewRow = {
    product,
    similar_product: similarProduct,
    confirmed: row.confirmed === true,
    code_choice: codeChoice,
    comparison_scope: normalizedReviewComparisonScope(row.comparison_scope, comparisonScope),
  };
  setReviewRuleText(rule, 'unit', row.unit);
  setReviewRuleText(rule, 'similar_unit', row.similar_unit);
  setReviewRuleText(rule, 'company', row.company);
  setReviewRuleText(rule, 'mst', row.mst);
  setReviewRuleText(rule, 'company_key', row.company_key);
  setReviewRuleText(rule, 'similar_company', row.similar_company);
  setReviewRuleText(rule, 'similar_mst', row.similar_mst);
  setReviewRuleText(rule, 'similar_company_key', row.similar_company_key);
  setReviewRuleText(rule, 'product_key', row.product_key);
  setReviewRuleText(rule, 'similar_product_key', row.similar_product_key);
  setReviewRuleText(rule, 'review_group', row.review_group);
  setReviewRuleText(rule, 'review_type', row.review_type);
  if (row.dimension_only) rule.dimension_only = true;
  if (codeChoice === 'split') {
    setReviewRuleText(rule, 'split_code', row.split_code);
    setReviewRuleText(rule, 'similar_split_code', row.similar_split_code);
  }
  return rule;
}

function reviewRuleFromRow(row: ReviewRow, comparisonScope: string): ReviewRow {
  const codeChoice = row.code_choice || 'current';
  const base = { ...row, confirmed: row.confirmed === true, code_choice: codeChoice, comparison_scope: comparisonScope } as ReviewRow & { comparison_scope: string };
  if (codeChoice === 'current' && row.product && row.similar_product) {
    const swapped = {
      ...base,
      product: row.similar_product,
      similar_product: row.product,
      unit: row.similar_unit,
      similar_unit: row.unit,
      company: row.similar_company,
      mst: row.similar_mst,
      company_key: row.similar_company_key,
      similar_company: row.company,
      similar_mst: row.mst,
      similar_company_key: row.company_key,
      product_key: row.similar_product_key,
      similar_product_key: row.product_key,
    };
    return compactReviewRule(swapped, comparisonScope) || swapped;
  }
  return compactReviewRule(base, comparisonScope) || base;
}

function reviewRuleKey(row: ReviewRow | null | undefined, comparisonScope: string) {
  if (!row?.product || !row?.similar_product) return '';
  const left = reviewRulePart(row.product, row.company, row.mst, row.company_key, comparisonScope);
  const right = reviewRulePart(row.similar_product, row.similar_company, row.similar_mst, row.similar_company_key, comparisonScope);
  if (!left || !right) return '';
  const group = row.review_group || (row.dimension_only ? 'dimension_diff' : 'other');
  return `${comparisonScope}|||${group}|||${[left, right].sort().join('<<<>>>')}`;
}

function reviewRulePart(product: string | undefined, company: string | undefined, mst: string | undefined, companyKey: string | undefined, comparisonScope: string) {
  const productPart = simpleMatchText(product);
  if (!productPart) return '';
  if (comparisonScope === 'same_company') {
    const companyPart = simpleMatchText(companyKey || mst || company);
    return companyPart ? `${companyPart}|||${productPart}` : '';
  }
  return productPart;
}

function buildConfigPayloads(workflow: WorkflowState, phase: 'purchase' | 'sales' | 'all' = 'all', profile: ProfileKey = 'vietmax') {
  if (isGenericProfileKey(profile) || profile === 'ho_guom') {
    if (isTwoPhaseGenericProfile(profile)) {
      const payloads = [];
      if (phase === 'purchase' || phase === 'all') payloads.push(buildGenericProcessPayload(workflow, profile));
      if (phase === 'sales' || phase === 'all') payloads.push(buildGenericSalesProcessPayload(workflow, profile));
      return payloads.length ? payloads : [buildGenericProcessPayload(workflow, profile)];
    }
    return [buildGenericProcessPayload(workflow, profile)];
  }
  const payloads = [];
  if ((phase === 'purchase' || phase === 'all') && workflow.companyRows.length) payloads.push(buildConfigPayload(workflow));
  if ((phase === 'sales' || phase === 'all') && (workflow.salesCompanyRows.length || workflow.salesFile || workflow.matches.length || workflow.salesMatchRules.length)) payloads.push(buildSalesConfigPayload(workflow));
  if (!payloads.length && phase === 'sales') return [buildSalesConfigPayload(workflow)];
  if (!payloads.length && phase === 'purchase') return [buildConfigPayload(workflow)];
  if (!payloads.length && phase === 'all') return [buildConfigPayload(workflow), buildSalesConfigPayload(workflow)];
  return payloads.length ? payloads : [buildConfigPayload(workflow)];
}

function buildConfigPayload(workflow: WorkflowState) {
  const companies = workflow.companyRows;
  return {
    ...buildPurchaseProcessPayload(workflow),
    columns: { ...normalizeVietmaxColumns(workflow.purchaseColumns, 'purchase'), purchase_price_col: normalizeVietmaxColumns(workflow.purchaseColumns, 'purchase').price_col || 'P' },
    prefixes: companyPrefixes(companies),
    removed_companies: Object.fromEntries(companies.filter(isIgnoredCompany).map((company) => [companyConfigKey(company), true])),
    processing_groups: workflow.purchaseProcessingGroups,
    company_group_assignments: groupAssignmentsFromRows(companies),
    form_mapping_presets: normalizeFormsForSave(workflow.purchaseFormMappingPresets),
    skipped_products_map: Object.fromEntries(companies.map((company) => {
      const selected = new Set(selectedProductNames(company));
      const skipped = company.all_products.map((product) => product.name).filter((name) => !selected.has(name));
      return [companyConfigKey(company), skipped];
    }).filter(([, skipped]) => Array.isArray(skipped) && skipped.length)),
    manual_code_overrides: workflow.productCodeOverrides,
    product_code_replacements: workflow.productCodeReplacements,
  };
}

function buildSalesConfigPayload(workflow: WorkflowState) {
  const companies = workflow.salesCompanyRows;
  return {
    ...buildSalesProcessPayload(workflow),
    columns: { ...normalizeVietmaxColumns(workflow.salesColumns, 'sales'), purchase_price_col: normalizeVietmaxColumns(workflow.purchaseColumns, 'purchase').price_col || 'P' },
    removed_companies: Object.fromEntries(companies.filter(isIgnoredCompany).map((company) => [companyConfigKey(company), true])),
    processing_groups: workflow.salesProcessingGroups,
    company_group_assignments: groupAssignmentsFromRows(companies),
    form_mapping_presets: normalizeFormsForSave(workflow.salesFormMappingPresets),
    skipped_products_map: Object.fromEntries(companies.map((company) => {
      const selected = new Set(selectedProductNames(company));
      const skipped = company.all_products.map((product) => product.name).filter((name) => !selected.has(name));
      return [companyConfigKey(company), skipped];
    }).filter(([, skipped]) => Array.isArray(skipped) && skipped.length)),
    manual_code_overrides: workflow.salesProductCodeOverrides,
    product_code_replacements: workflow.salesProductCodeReplacements,
    vietmax_ban_ra_purchase_match_rules: buildSalesMatchRules(workflow),
  };
}

function buildVietmaxConfigExportSnapshot(workflow: WorkflowState, phase: 'purchase' | 'sales', appConfig: Record<string, unknown>) {
  const profileKeyName = phase === 'sales' ? 'vietmax_ban_ra' : 'vietmax_mua_vao';
  const otherProfileKeyName = phase === 'sales' ? 'vietmax_mua_vao' : 'vietmax_ban_ra';
  const profilesConfig = appConfig.profiles && typeof appConfig.profiles === 'object'
    ? appConfig.profiles as Record<string, unknown>
    : {};
  const processPayload = phase === 'sales' ? buildSalesConfigPayload(workflow) : buildConfigPayload(workflow);
  const companyRows = phase === 'sales' ? workflow.salesCompanyRows : workflow.companyRows;
  return {
    export_version: 1,
    exported_at: new Date().toISOString(),
    app: 'Product Code Formatter',
    profile: 'vietmax',
    phase,
    storage_profile: profileKeyName,
    description: phase === 'sales' ? 'Vietmax bán ra' : 'Vietmax mua vào',
    notes: [
      'process_payload là cấu hình xử lý được sinh từ UI hiện tại để tái tạo file với cùng file nguồn.',
      'saved_profile_config là phần đang lưu trong product_code_config.json cho phase này.',
      'saved_config_file_snapshot chứa app_version, selected_profile, columns và profiles đã chuẩn hóa; không xuất license.',
      'current_ui_snapshot dùng để kiểm tra danh sách công ty, prefix và hàng hóa đang được chọn/bỏ qua.',
      'selected_products trong saved_profile_config là tên legacy, thực tế đang lưu danh sách hàng hóa bị bỏ qua theo MST.',
    ],
    source_files: {
      purchase_original_name: workflow.purchaseFile?.original_name || '',
      purchase_saved_name: workflow.purchaseFile?.saved_name || '',
      processed_purchase_saved_name: workflow.processedPurchaseSavedName || '',
      sales_original_name: workflow.salesFile?.original_name || '',
      sales_saved_name: workflow.salesFile?.saved_name || '',
      processed_sales_saved_name: workflow.processedSalesSavedName || '',
    },
    saved_config_file_snapshot: buildSavedConfigFileSnapshot(appConfig),
    process_payload: processPayload,
    saved_profile_config: profilesConfig[profileKeyName] || {},
    paired_saved_profile_config: profilesConfig[otherProfileKeyName] || {},
    all_saved_vietmax_configs: {
      vietmax_mua_vao: profilesConfig.vietmax_mua_vao || {},
      vietmax_ban_ra: profilesConfig.vietmax_ban_ra || {},
    },
    current_ui_snapshot: buildVietmaxUiConfigSnapshot(workflow, phase, companyRows),
  };
}

function buildSavedConfigFileSnapshot(appConfig: Record<string, unknown>) {
  const columns = appConfig.columns && typeof appConfig.columns === 'object' && !Array.isArray(appConfig.columns)
    ? appConfig.columns as Record<string, unknown>
    : {};
  const profiles = appConfig.profiles && typeof appConfig.profiles === 'object' && !Array.isArray(appConfig.profiles)
    ? appConfig.profiles as Record<string, unknown>
    : {};
  return {
    app_version: appConfig.app_version || '',
    selected_profile: appConfig.selected_profile || '',
    columns,
    profiles,
  };
}

function buildVietmaxUiConfigSnapshot(workflow: WorkflowState, phase: 'purchase' | 'sales', companies: CompanyRow[]) {
  const isSales = phase === 'sales';
  const reviewScope = isSales ? workflow.salesReviewScope : workflow.purchaseReviewScope;
  const reviewRows = isSales ? workflow.salesReviewRows : workflow.purchaseReviewRows;
  const reviewRules = isSales ? workflow.salesReviewRules : workflow.purchaseReviewRules;
  const inventoryPairs = isSales ? workflow.salesInventoryPairs : workflow.purchaseInventoryPairs;
  const inventoryRules = isSales ? workflow.salesInventoryPairRules : workflow.purchaseInventoryPairRules;
  const prefixStrategy = isSales ? workflow.salesPrefixStrategy : workflow.purchasePrefixStrategy;
  const prefixStrategyValues = isSales ? workflow.salesPrefixStrategyValues : workflow.purchasePrefixStrategyValues;
  return {
    phase,
    columns: normalizeVietmaxColumns(isSales ? workflow.salesColumns : workflow.purchaseColumns, phase),
    company_count: companies.length,
    processed_company_count: companies.filter((company) => company.process !== false).length,
    product_count: companies.reduce((total, company) => total + company.all_products.length, 0),
    selected_product_count: companies.reduce((total, company) => total + selectedProductNames(company).length, 0),
    include_company_prefix: isSales ? workflow.salesIncludeCompanyPrefix : workflow.includeCompanyPrefix,
    prefix_strategy: prefixStrategy,
    prefix_mst_digits: workflow.prefixMstDigits,
    prefix_name_words: workflow.prefixNameWords,
    prefix_name_chars: workflow.prefixNameChars,
    prefix_missing_mst_strategy: normalizeMissingMstPrefixStrategy(workflow.prefixMissingMstStrategy),
    prefix_strategy_values: prefixStrategyValues,
    word_rules: isSales ? workflow.salesWordRules : workflow.purchaseWordRules,
    repeated_phrase_removals: isSales ? workflow.salesRepeatedPhraseRemovals : workflow.purchaseRepeatedPhraseRemovals,
    manual_code_overrides: isSales ? workflow.salesProductCodeOverrides : workflow.productCodeOverrides,
    product_code_replacements: isSales ? workflow.salesProductCodeReplacements : workflow.productCodeReplacements,
    review_scope: reviewScope,
    review_rows: reviewRows,
    saved_review_rules: buildReviewRules(reviewRules, reviewRows, reviewScopeValue(reviewScope)),
    inventory_pairs: inventoryPairs,
    use_default_inventory_pair: isSales ? workflow.salesUseDefaultInventoryPair : workflow.purchaseUseDefaultInventoryPair,
    default_inventory_pair_id: isSales ? workflow.salesDefaultInventoryPairId : workflow.purchaseDefaultInventoryPairId,
    inventory_pair_rules: inventoryRulesForPayload(inventoryRules, inventoryPairs),
    companies: companies.map(configExportCompanyRow),
    sales_purchase_matches: isSales ? workflow.matches : [],
    saved_sales_purchase_match_rules: isSales ? buildSalesMatchRules(workflow) : [],
  };
}

function configExportCompanyRow(company: CompanyRow, index: number) {
  const selectedNames = selectedProductNames(company);
  const selectedSet = new Set(selectedNames);
  return {
    index,
    mst: company.mst,
    company: company.company,
    safe_id: company.safe_id,
    process: company.process !== false,
    pending_process: company.pending_process ?? company.process ?? true,
    prefix: normalizePrefixValue(company.value),
    committed_prefix: normalizePrefixValue(company.committed_prefix),
    default_prefix: normalizePrefixValue(company.default_prefix),
    prefix_strategies: company.prefix_strategies || {},
    selected_product_count: selectedNames.length,
    skipped_product_count: Math.max(0, company.all_products.length - selectedNames.length),
    selected_product_names: selectedNames,
    skipped_product_names: company.all_products.map((product) => product.name).filter((name) => !selectedSet.has(name)),
    all_products: company.all_products.map((product) => ({
      name: product.name,
      count: product.count ?? 0,
      minPrice: product.minPrice ?? null,
      maxPrice: product.maxPrice ?? null,
      priceCount: product.priceCount ?? 0,
      sampleRows: (product.priceRows || []).slice(0, 5),
    })),
  };
}

function exportTimestamp(date = new Date()) {
  const pad = (value: number) => String(value).padStart(2, '0');
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
    '_',
    pad(date.getHours()),
    pad(date.getMinutes()),
    pad(date.getSeconds()),
  ].join('');
}

function chooseJsonConfigFile(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'application/json,.json';
    input.onchange = () => resolve(input.files?.[0] ?? null);
    input.oncancel = () => resolve(null);
    input.click();
  });
}

function companyPrefixes(companies: CompanyRow[]) {
  return Object.fromEntries(companies.filter((company) => normalizePrefixValue(company.value)).map((company) => [companyConfigKey(company), normalizePrefixValue(company.value)]));
}

function companyPrefixFields(companies: CompanyRow[]) {
  return Object.fromEntries(companies.map((company, index) => [`prefix_${index}`, normalizePrefixValue(company.value)]));
}

function selectedProductNames(company: CompanyRow) {
  return company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name);
}

function processedStatsSentence(stats: ProcessedFileStats | null) {
  if (!stats) return '';
  return `Công ty ${formatCount(stats.processed_company_count)}/${formatCount(stats.company_count)}, dòng hàng ${formatCount(stats.processed_product_row_count)}/${formatCount(stats.product_row_count)}.`;
}

function formatCount(value: number | undefined) {
  return Number(value || 0).toLocaleString('en-US');
}


function productKey(mst: string | undefined, product: string | undefined) {
  return product ? `${mst || ''}|||${product}` : '';
}

function confirmedSalesMatchKeys(rows: MatchRow[], comparisonScope: string) {
  return new Set(rows.filter((row) => row.confirmed !== false && row.sales_product).map((row) => salesProductMatchKey(row.sales_product, row.sales_company, row.sales_mst, comparisonScope)));
}

function salesProductMatchKey(product: string | undefined, company: string | undefined, mst: string | undefined, comparisonScope: string) {
  const productPart = simpleMatchText(product);
  if (comparisonScope === 'same_company') {
    return `${simpleMatchText(mst || company)}|||${productPart}`;
  }
  return productPart;
}

function simpleMatchText(value: string | undefined) {
  return String(value || '').trim().toLocaleLowerCase('vi-VN');
}

function splitWords(value: string | undefined) {
  return String(value || '').trim().split(/\s+/).filter(Boolean);
}

type Opcode = { tag: 'equal' | 'delete' | 'insert' | 'replace'; i1: number; i2: number; j1: number; j2: number };

function sequenceOpcodes<T>(left: T[], right: T[], equal: (left: T, right: T) => boolean): Opcode[] {
  const lengths = Array.from({ length: left.length + 1 }, () => Array(right.length + 1).fill(0));
  for (let i = left.length - 1; i >= 0; i -= 1) {
    for (let j = right.length - 1; j >= 0; j -= 1) {
      lengths[i][j] = equal(left[i], right[j]) ? lengths[i + 1][j + 1] + 1 : Math.max(lengths[i + 1][j], lengths[i][j + 1]);
    }
  }
  const opcodes: Opcode[] = [];
  let i = 0;
  let j = 0;
  while (i < left.length || j < right.length) {
    if (i < left.length && j < right.length && equal(left[i], right[j])) {
      const i1 = i;
      const j1 = j;
      while (i < left.length && j < right.length && equal(left[i], right[j])) {
        i += 1;
        j += 1;
      }
      opcodes.push({ tag: 'equal', i1, i2: i, j1, j2: j });
    } else {
      const i1 = i;
      const j1 = j;
      while (i < left.length && j < right.length && !equal(left[i], right[j])) {
        if (lengths[i + 1][j] >= lengths[i][j + 1]) i += 1;
        else j += 1;
      }
      if (i === i1 && j < right.length) j += 1;
      if (j === j1 && i < left.length) i += 1;
      const tag = i > i1 && j > j1 ? 'replace' : i > i1 ? 'delete' : 'insert';
      opcodes.push({ tag, i1, i2: i, j1, j2: j });
    }
  }
  return opcodes;
}

function wordsAreSmallCharacterDiff(current: string, target: string) {
  if (!current || !target || current === target) return false;
  const changed = sequenceOpcodes([...current], [...target], (left, right) => left === right).reduce((total, opcode) => {
    if (opcode.tag === 'equal') return total;
    return total + Math.max(opcode.i2 - opcode.i1, opcode.j2 - opcode.j1);
  }, 0);
  return changed <= 3 && similarityRatio(current.toLocaleLowerCase(), target.toLocaleLowerCase()) >= 0.66;
}

function similarityRatio(left: string, right: string) {
  if (!left.length && !right.length) return 1;
  const matches = sequenceOpcodes([...left], [...right], (a, b) => a === b).reduce((total, opcode) => total + (opcode.tag === 'equal' ? opcode.i2 - opcode.i1 : 0), 0);
  return (2 * matches) / (left.length + right.length);
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
  const userActivation = (navigator as Navigator & { userActivation?: { isActive: boolean } }).userActivation;
  if (window.showSaveFilePicker && userActivation?.isActive !== false) {
    try {
      const lowerName = filename.toLowerCase();
      const pickerType: { description: string; accept: Record<string, string[]> } = lowerName.endsWith('.json')
        ? { description: 'JSON config', accept: { 'application/json': ['.json'] } }
        : lowerName.endsWith('.zip')
          ? { description: 'ZIP package', accept: { 'application/zip': ['.zip'] } }
        : lowerName.endsWith('.xls')
          ? { description: 'Excel workbook', accept: { 'application/vnd.ms-excel': ['.xls'] } }
          : lowerName.endsWith('.xlsm')
            ? { description: 'Excel macro workbook', accept: { 'application/vnd.ms-excel.sheet.macroEnabled.12': ['.xlsm'] } }
            : { description: 'Excel workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } };
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [pickerType],
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return true;
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return false;
      if (error instanceof DOMException && (error.name === 'NotAllowedError' || error.name === 'SecurityError')) {
        downloadBlob(blob, filename);
        return true;
      }
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
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function sleep(milliseconds: number) {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

function formatSimilarity(value: number | string | undefined) {
  if (typeof value === 'string') return value;
  return typeof value === 'number' ? `${Math.round(value * 100)}%` : '';
}
