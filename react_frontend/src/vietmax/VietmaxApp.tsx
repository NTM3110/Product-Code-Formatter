import { Fragment, useEffect, useMemo, useState } from 'react';
import { activateLicense, analyzeVietmaxCompanies, createPurchaseReview, createSalesMatches, downloadInventoryAllocationReport, exportMatches, getInventoryAllocationJob, getLicenseStatus, previewVietmaxProductCodes, processVietmaxPurchase, saveVietmaxConfig, reloadLicense, startInventoryAllocation, uploadExcel } from '../api';
import type { CompanyRow, InventoryAllocationConfig, InventoryAllocationJob, InventoryAllocationResult, InventoryPair, InventoryRule, LicenseStatus, MatchRow, ReviewProduct, ReviewRow, UploadSummary } from '../types';
import { InventoryAllocationStage } from './InventoryAllocationStage';

type ProfileKey = 'son_phuong' | 'cao_thanh' | 'quang_thinh' | 'vietmax';
type StageId = 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12;
type StagePhase = 'purchase' | 'sales' | 'generic' | 'price' | 'inventory';

type StageDefinition = {
  id: StageId;
  label: string;
  phase: StagePhase;
  short: string;
};

type CustomPrefixOption = {
  name: string;
  formula: string;
  chars: number;
  mstDigits?: number;
};

type WorkflowState = {
  stage: StageId;
  purchaseFile: UploadSummary | null;
  processedPurchaseSavedName: string;
  salesFile: UploadSummary | null;
  processedSalesSavedName: string;
  openingStockFile: File | null;
  inventoryAllocationConfig: InventoryAllocationConfig;
  inventoryAllocationJob: InventoryAllocationJob | null;
  inventoryAllocationResult: InventoryAllocationResult | null;
  comparisonScope: string;
  companyRows: CompanyRow[];
  selectedCompanyIndex: number;
  salesCompanyRows: CompanyRow[];
  selectedSalesCompanyIndex: number;
  productPreviewCodes: Record<string, string>;
  salesProductPreviewCodes: Record<string, string>;
  productCodeOverrides: Record<string, string>;
  salesProductCodeOverrides: Record<string, string>;
  wordRules: Record<string, string>;
  repeatedPhraseRemovals: string[];
  purchaseReviewRows: ReviewRow[];
  salesReviewRows: ReviewRow[];
  purchaseReviewGenerated: boolean;
  salesReviewGenerated: boolean;
  matches: MatchRow[];
  salesMatchGenerated: boolean;
  salesMatchRules: MatchRow[];
  inventoryPairs: InventoryPair[];
  useDefaultInventoryPair: boolean;
  defaultInventoryPairId: string;
  inventoryPairRules: InventoryRule[];
  includeCompanyPrefix: boolean;
  prefixStrategy: string;
  prefixMstDigits: number;
  customPrefixOptions: CustomPrefixOption[];
  purchaseReviewScope: 'all' | 'company';
  salesReviewScope: 'all' | 'company';
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

const profiles: Array<{ key: ProfileKey; label: string; note: string }> = [
  { key: 'son_phuong', label: 'Sơn Phương', note: 'Sẽ migrate sau Vietmax.' },
  { key: 'cao_thanh', label: 'Cao Thành', note: 'Sẽ migrate sau Vietmax, gồm stage lọc đơn giá.' },
  { key: 'quang_thinh', label: 'Quang Thịnh', note: 'Sẽ migrate sau Vietmax.' },
  { key: 'vietmax', label: 'Vietmax', note: 'Đang migrate trước: mua vào rồi bán ra, stage 1-11.' },
];

const vietmaxStages: StageDefinition[] = [
  { id: 1, label: 'Tải file mua vào', phase: 'purchase', short: 'Tải mua vào' },
  { id: 2, label: 'Chọn cột / preview / trạng thái', phase: 'purchase', short: 'Chọn cột' },
  { id: 3, label: 'Công ty & prefix', phase: 'purchase', short: 'Công ty' },
  { id: 4, label: 'Review Mã VT', phase: 'purchase', short: 'Review Mã VT' },
  { id: 5, label: 'Tạo file mua vào', phase: 'purchase', short: 'Tạo mua vào' },
  { id: 6, label: 'Tải file bán ra', phase: 'sales', short: 'Tải bán ra' },
  { id: 7, label: 'Chọn cột / preview / trạng thái', phase: 'sales', short: 'Cột bán ra' },
  { id: 8, label: 'Khớp HD mua vào', phase: 'sales', short: 'Khớp mua vào' },
  { id: 9, label: 'Công ty & prefix', phase: 'sales', short: 'Công ty' },
  { id: 10, label: 'Review Mã VT', phase: 'sales', short: 'Review bán ra' },
  { id: 11, label: 'Tạo file bán ra', phase: 'sales', short: 'Tạo bán ra' },
  { id: 12, label: 'Phân bổ tồn kho', phase: 'inventory', short: 'Phân bổ' },
];

const commonProfileStages: StageDefinition[] = [
  { id: 1, label: 'Tải file', phase: 'generic', short: 'Tải file' },
  { id: 2, label: 'Khớp HD mua vào', phase: 'generic', short: 'Khớp HD' },
  { id: 3, label: 'Chọn cột / preview / trạng thái', phase: 'generic', short: 'Chọn cột' },
  { id: 4, label: 'Công ty & prefix', phase: 'generic', short: 'Công ty' },
  { id: 5, label: 'Review Mã VT', phase: 'generic', short: 'Review Mã VT' },
];

const priceProfileStage: StageDefinition = { id: 6, label: 'Lọc đơn giá', phase: 'price', short: 'Lọc đơn giá' };

function stagesForProfile(profile: ProfileKey): StageDefinition[] {
  if (profile === 'vietmax') return vietmaxStages;
  if (profile === 'cao_thanh') return [...commonProfileStages, priceProfileStage];
  return commonProfileStages;
}

function initialWorkflowState(): WorkflowState {
  return {
    stage: 1,
    purchaseFile: null,
    processedPurchaseSavedName: '',
    salesFile: null,
    processedSalesSavedName: '',
    openingStockFile: null,
    inventoryAllocationConfig: defaultInventoryAllocationConfig(),
    inventoryAllocationJob: null,
    inventoryAllocationResult: null,
    comparisonScope: 'all_companies',
    companyRows: [],
    selectedCompanyIndex: -1,
    salesCompanyRows: [],
    selectedSalesCompanyIndex: -1,
    productPreviewCodes: {},
    salesProductPreviewCodes: {},
    productCodeOverrides: {},
    salesProductCodeOverrides: {},
    wordRules: {},
    repeatedPhraseRemovals: [],
    purchaseReviewRows: [],
    salesReviewRows: [],
    purchaseReviewGenerated: false,
    salesReviewGenerated: false,
    matches: [],
    salesMatchGenerated: false,
    salesMatchRules: [],
    inventoryPairs: [],
    useDefaultInventoryPair: false,
    defaultInventoryPairId: '',
    inventoryPairRules: [],
    includeCompanyPrefix: false,
    prefixStrategy: 'last_2_words',
    prefixMstDigits: 3,
    customPrefixOptions: [
      { name: 'Custom 1', formula: 'initials', chars: 2, mstDigits: 3 },
      { name: 'Custom 2', formula: 'mst', chars: 3, mstDigits: 3 },
      { name: 'Custom 3', formula: 'initials-mst', chars: 2, mstDigits: 3 },
    ],
    purchaseReviewScope: 'all',
    salesReviewScope: 'company',
  };
}

function initialWorkflowStates(): Record<ProfileKey, WorkflowState> {
  return {
    son_phuong: initialWorkflowState(),
    cao_thanh: initialWorkflowState(),
    quang_thinh: initialWorkflowState(),
    vietmax: initialWorkflowState(),
  };
}

function initialLicenseForm() {
  return { server_url: '', account_id: '', license_key: '' };
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
      ignore_sale_suffix: false,
      allow_negative_export: true,
      company_profile: 'yen_thanh',
      allow_future_purchase_reorder: false,
      future_purchase_window_days: 31,
    },
  };
}

function salesOutputInvalidation(): Partial<WorkflowState> {
  return { processedSalesSavedName: '', inventoryAllocationJob: null, inventoryAllocationResult: null };
}

function purchaseOutputInvalidation(): Partial<WorkflowState> {
  return { processedPurchaseSavedName: '', matches: [], salesMatchGenerated: false, ...salesOutputInvalidation() };
}

export function VietmaxApp() {
  const [profile, setProfile] = useState<ProfileKey>('vietmax');
  const [workflows, setWorkflows] = useState<Record<ProfileKey, WorkflowState>>(initialWorkflowStates);
  const [license, setLicense] = useState<LicenseStatus | null>(null);
  const [licenseForm, setLicenseForm] = useState(initialLicenseForm);
  const [status, setStatus] = useState('Chọn profile và bắt đầu theo từng stage. Dữ liệu được giữ khi chuyển stage, chỉ xóa khi bấm Làm lại.');
  const [busy, setBusy] = useState(false);

  const workflow = workflows[profile];
  const { stage, purchaseFile, processedPurchaseSavedName, salesFile, processedSalesSavedName, openingStockFile, inventoryAllocationConfig, inventoryAllocationJob, inventoryAllocationResult, comparisonScope, companyRows, selectedCompanyIndex, salesCompanyRows, selectedSalesCompanyIndex, productPreviewCodes, salesProductPreviewCodes, productCodeOverrides, salesProductCodeOverrides, wordRules, repeatedPhraseRemovals, purchaseReviewRows, salesReviewRows, purchaseReviewGenerated, salesReviewGenerated, matches, salesMatchGenerated, salesMatchRules, inventoryPairs, useDefaultInventoryPair, defaultInventoryPairId, inventoryPairRules, includeCompanyPrefix, prefixStrategy, prefixMstDigits, customPrefixOptions, purchaseReviewScope, salesReviewScope } = workflow;
  const selectedProfile = profiles.find((item) => item.key === profile) ?? profiles[0];
  const licenseReady = Boolean(license?.activated && (profile !== 'vietmax' || license.vietmax_allowed));
  const visibleStages = useMemo(() => stagesForProfile(profile), [profile]);
  const currentStage = visibleStages.find((item) => item.id === stage) ?? visibleStages[0];
  const selectedMatches = useMemo(() => matches.filter((match) => match.confirmed !== false), [matches]);
  const showLicenseBar = stage === 1;

  function updateWorkflow(targetProfile: ProfileKey, update: Partial<WorkflowState>) {
    setWorkflows((current) => ({ ...current, [targetProfile]: { ...current[targetProfile], ...update } }));
  }

  useEffect(() => {
    getLicenseStatus()
      .then((nextLicense) => {
        setLicense(nextLicense);
        setLicenseForm((current) => ({ ...current, server_url: nextLicense.server_url, account_id: nextLicense.account_id }));
      })
      .catch((error) => setStatus(error instanceof Error ? error.message : String(error)));
  }, []);

  useEffect(() => {
    if (profile === 'vietmax' && stage === 3 && purchaseFile && !companyRows.length && !busy) {
      void loadCompanies();
    }
  }, [profile, stage, purchaseFile, companyRows.length, busy]);

  useEffect(() => {
    if (profile === 'vietmax' && stage === 9 && salesFile && !salesCompanyRows.length && !busy) {
      void loadSalesCompanies();
    }
  }, [profile, stage, salesFile, salesCompanyRows.length, busy]);

  useEffect(() => {
    if (profile === 'vietmax' && stage === 4 && purchaseFile && companyRows.length && !purchaseReviewGenerated && !busy) {
      void runPurchaseReview();
    }
  }, [profile, stage, purchaseFile, companyRows.length, purchaseReviewGenerated, busy]);

  useEffect(() => {
    if (profile === 'vietmax' && stage === 10 && salesFile && salesCompanyRows.length && !salesReviewGenerated && !busy) {
      void runSalesReview();
    }
  }, [profile, stage, salesFile, salesCompanyRows.length, salesReviewGenerated, busy]);

  function updateCustomPrefix(index: number, option: CustomPrefixOption) {
    const nextOptions = [...customPrefixOptions];
    nextOptions[index] = option;
    updateWorkflow(profile, { customPrefixOptions: nextOptions });
  }

  function resetWorkflow() {
    updateWorkflow(profile, initialWorkflowState());
    setStatus(profile === 'vietmax' ? 'Đã làm lại. Hãy tải file mua vào Vietmax từ stage 1.' : `Đã làm lại shell ${selectedProfile.label}.`);
  }

  function canEnterStage(target: StageId) {
    if (!visibleStages.some((item) => item.id === target)) return false;
    if (profile !== 'vietmax') return licenseReady || target === 1;
    if (!licenseReady) return target === 1;
    if (target <= 2) return true;
    if (target <= 6) return Boolean(purchaseFile);
    if (target <= 11) return Boolean(salesFile);
    return Boolean(processedPurchaseSavedName && processedSalesSavedName);
  }

  function goToStage(target: StageId) {
    if (!canEnterStage(target)) return;
    updateWorkflow(profile, { stage: target });
  }

  function goBack() {
    const currentIndex = visibleStages.findIndex((item) => item.id === stage);
    const previous = visibleStages[Math.max(0, currentIndex - 1)];
    if (previous) goToStage(previous.id);
  }

  function goNext() {
    const currentIndex = visibleStages.findIndex((item) => item.id === stage);
    const next = visibleStages[currentIndex + 1];
    if (!next) return;
    if (profile !== 'vietmax') {
      goToStage(next.id);
      return;
    }
    if (stage === 4 && purchaseFile && !purchaseReviewGenerated) {
      void runPurchaseReview();
      return;
    }
    if (stage === 5 && purchaseFile && !processedPurchaseSavedName) {
      void prepareProcessedPurchaseCache(next.id);
      return;
    }
    if (stage === 8 && salesFile && processedPurchaseSavedName && !matches.length) {
      void runSalesMatch();
      return;
    }
    if (stage === 10 && salesFile && !salesReviewGenerated) {
      void runSalesReview();
      return;
    }
    goToStage(next.id);
  }

  async function submitLicense() {
    setBusy(true);
    setStatus('Đang kích hoạt license...');
    try {
      const nextLicense = await activateLicense(licenseForm);
      setLicense(nextLicense);
      setStatus(nextLicense.vietmax_allowed ? 'Kích hoạt thành công. Vietmax được phép dùng.' : 'License chưa cho phép Vietmax.');
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
      setStatus(nextLicense.vietmax_allowed ? 'Đã tải lại license. Vietmax được phép dùng.' : 'Đã tải lại license, nhưng Vietmax chưa được phép.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function upload(kind: 'purchase' | 'sales', file: File | undefined) {
    if (!file) return;
    const targetProfile = profile;
    setBusy(true);
    setStatus(`Đang tải ${kind === 'purchase' ? 'HD mua vào' : 'HD bán ra'}...`);
    try {
      const summary = await uploadExcel(file);
      if (kind === 'purchase') {
        updateWorkflow(targetProfile, {
          ...purchaseOutputInvalidation(),
          purchaseFile: summary,
          stage: 2,
          companyRows: [],
          selectedCompanyIndex: -1,
          productPreviewCodes: {},
          productCodeOverrides: {},
          purchaseReviewRows: [],
          purchaseReviewGenerated: false,
          salesCompanyRows: [],
          selectedSalesCompanyIndex: -1,
          salesProductPreviewCodes: {},
          salesProductCodeOverrides: {},
          salesReviewRows: [],
          salesReviewGenerated: false,
          inventoryAllocationJob: null,
          inventoryAllocationResult: null,
        });
      } else {
        updateWorkflow(targetProfile, {
          ...salesOutputInvalidation(),
          salesFile: summary,
          stage: 7,
          salesCompanyRows: [],
          selectedSalesCompanyIndex: -1,
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
    setBusy(true);
    setStatus('Đang tạo review Mã VT mua vào bằng logic Vietmax...');
    try {
      const result = await createPurchaseReview(targetPurchaseFile.saved_name, scope, wordRules, repeatedPhraseRemovals, reviewProducts);
      updateWorkflow(targetProfile, { ...purchaseOutputInvalidation(), purchaseReviewRows: normalizeReviewRows(result.review_rows as ReviewRow[]), purchaseReviewGenerated: true, stage: 4 });
      setStatus(`Đã tạo ${result.review_rows.length} dòng review Mã VT mua vào. Chỉ dòng được tick ở cột Dùng và lựa chọn Dùng mã sẽ được áp dụng khi xử lý.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function applyProcessedPurchaseCache(targetProfile: ProfileKey, processedSavedName: string) {
    if (!processedSavedName) return;
    updateWorkflow(targetProfile, {
      ...salesOutputInvalidation(),
      processedPurchaseSavedName: processedSavedName,
      matches: [],
      salesMatchGenerated: false,
      salesCompanyRows: [],
      selectedSalesCompanyIndex: -1,
      salesProductPreviewCodes: {},
      salesProductCodeOverrides: {},
      salesReviewRows: [],
      salesReviewGenerated: false,
    });
  }

  async function prepareProcessedPurchaseCache(nextStage?: StageId) {
    if (!purchaseFile) return;
    if (processedPurchaseSavedName) {
      if (nextStage) goToStage(nextStage);
      return;
    }
    const targetProfile = profile;
    const targetPurchaseFile = purchaseFile;
    setBusy(true);
    setStatus('Đang tạo cache file mua vào đã xử lý để dùng cho khớp mua/bán...');
    try {
      const result = await processVietmaxPurchase(targetPurchaseFile.saved_name, targetPurchaseFile.original_name, buildPurchaseProcessPayload(workflow));
      if (!result.processedSavedName) throw new Error('Không tạo được cache file mua vào đã xử lý.');
      applyProcessedPurchaseCache(targetProfile, result.processedSavedName);
      setStatus('Đã tạo cache file mua vào đã xử lý. Stage khớp mua/bán có thể dùng file này dù bạn chưa xuất Excel.');
      if (nextStage) updateWorkflow(targetProfile, { stage: nextStage });
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function runSalesMatch() {
    if (!salesFile) return;
    const targetProfile = profile;
    const targetSalesFile = salesFile;
    const targetPurchaseFile = purchaseFile;
    let targetProcessedPurchase = processedPurchaseSavedName;
    setBusy(true);
    setStatus(targetProcessedPurchase ? 'Đang khớp bán ra với file mua vào đã xử lý KHH/152...' : 'Đang chuẩn bị cache file mua vào đã xử lý trước khi khớp mua/bán...');
    try {
      if (!targetProcessedPurchase) {
        if (!targetPurchaseFile) throw new Error('Chưa có file mua vào để tạo cache xử lý.');
        const purchaseResult = await processVietmaxPurchase(targetPurchaseFile.saved_name, targetPurchaseFile.original_name, buildPurchaseProcessPayload(workflow));
        targetProcessedPurchase = purchaseResult.processedSavedName;
        if (!targetProcessedPurchase) throw new Error('Không tạo được cache file mua vào đã xử lý.');
        applyProcessedPurchaseCache(targetProfile, targetProcessedPurchase);
        setStatus('Đã tạo cache mua vào. Đang khớp bán ra với file mua vào đã xử lý KHH/152...');
      }
      const result = await createSalesMatches(targetSalesFile.saved_name, targetProcessedPurchase, comparisonScope);
      const savedRules = result.match_rules?.length ? result.match_rules : salesMatchRules;
      const nextMatches = applySalesMatchRules(result.matches, savedRules, comparisonScope);
      const mismatchCount = nextMatches.filter(hasUnitMismatch).length;
      updateWorkflow(targetProfile, { ...salesOutputInvalidation(), matches: nextMatches, salesMatchGenerated: true, salesMatchRules: savedRules, salesCompanyRows: [], selectedSalesCompanyIndex: -1, salesProductPreviewCodes: {}, salesProductCodeOverrides: {}, salesReviewRows: [], salesReviewGenerated: false, stage: 8 });
      setStatus(`Đã gợi ý ${result.matches.length} dòng khớp. ${result.exact_matches.length} dòng lấy chính xác từ KHH/152. ${mismatchCount} dòng khác ĐVT.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function downloadMatches() {
    setBusy(true);
    setStatus('Đang xuất Excel danh sách khớp...');
    try {
      const blob = await exportMatches(selectedMatches);
      const saved = await saveBlob(blob, 'vietmax_khop_mua_ban.xlsx');
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
    void saveWorkflowConfig(nextWorkflow, 'Đã lưu cấu hình khớp mua/bán theo công ty, hàng hóa và công thức quy đổi ĐVT.');
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
    const nextWorkflow = { ...workflow, ...purchaseOutputInvalidation(), purchaseReviewRows };
    updateWorkflow(profile, nextWorkflow);
    void saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu lựa chọn Review Mã VT vào cấu hình.');
  }

  function applySalesReviewChoices() {
    const nextWorkflow = { ...workflow, ...salesOutputInvalidation(), salesReviewRows };
    updateWorkflow(profile, nextWorkflow);
    void saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu lựa chọn Review bán ra. Các dòng được tick sẽ được tính vào file bán ra khi xuất.');
  }

  async function loadCompanies() {
    if (!purchaseFile) return;
    const targetProfile = profile;
    setBusy(true);
    setStatus('Đang tải danh sách công ty và hàng hóa mua vào...');
    try {
      const result = await analyzeVietmaxCompanies(purchaseFile.saved_name, 'purchase');
      const savedWordRules = result.word_rules ?? wordRules;
      const savedRepeatedPhrases = result.repeated_phrase_removals ?? repeatedPhraseRemovals;
      const savedInventoryPairs = result.inventory_pairs ?? inventoryPairs;
      const nextCompanies = result.companies.map((company) => ({
        ...company,
        process: company.process ?? true,
        pending_process: company.pending_process ?? company.process ?? true,
        committed_prefix: company.committed_prefix ?? company.value ?? '',
        selected_product_names: company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name),
      }));
      const previewCodes = await loadProductPreviewCodes(nextCompanies, savedWordRules, savedRepeatedPhrases);
      updateWorkflow(targetProfile, { ...purchaseOutputInvalidation(), companyRows: nextCompanies, selectedCompanyIndex: firstDisplayedCompanyIndex(nextCompanies), productPreviewCodes: previewCodes, productCodeOverrides: {}, wordRules: savedWordRules, repeatedPhraseRemovals: savedRepeatedPhrases, inventoryPairs: savedInventoryPairs, useDefaultInventoryPair: result.use_default_inventory_pair ?? useDefaultInventoryPair, defaultInventoryPairId: result.default_inventory_pair_id ?? defaultInventoryPairId, inventoryPairRules: result.inventory_pair_rules ?? inventoryPairRules, purchaseReviewRows: [], purchaseReviewGenerated: false });
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
      const result = await analyzeVietmaxCompanies(salesFile.saved_name, 'sales');
      const savedInventoryPairs = result.inventory_pairs?.length ? result.inventory_pairs : inventoryPairs;
      const khhMatchedKeys = confirmedSalesMatchKeys(matches, comparisonScope);
      const nextCompanies = result.companies.map((company) => ({
        ...company,
        all_products: company.all_products.filter((product) => !khhMatchedKeys.has(salesProductMatchKey(product.name, company.company, company.mst, comparisonScope))),
        process: company.process ?? true,
        pending_process: company.pending_process ?? company.process ?? true,
        committed_prefix: company.committed_prefix ?? company.value ?? '',
      })).map((company) => ({
        ...company,
        selected_product_names: (company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name)).filter((name) => company.all_products.some((product) => product.name === name)),
      })).filter((company) => company.all_products.length);
      const previewCodes = await loadProductPreviewCodes(nextCompanies, wordRules, repeatedPhraseRemovals, 'sales');
      updateWorkflow(targetProfile, { ...salesOutputInvalidation(), salesCompanyRows: nextCompanies, selectedSalesCompanyIndex: firstDisplayedCompanyIndex(nextCompanies), salesProductPreviewCodes: previewCodes, salesProductCodeOverrides: {}, salesMatchRules: result.sales_match_rules ?? salesMatchRules, inventoryPairs: savedInventoryPairs, useDefaultInventoryPair: result.inventory_pairs?.length ? Boolean(result.use_default_inventory_pair) : useDefaultInventoryPair, defaultInventoryPairId: result.inventory_pairs?.length ? (result.default_inventory_pair_id ?? '') : defaultInventoryPairId, inventoryPairRules: result.inventory_pair_rules?.length ? result.inventory_pair_rules : inventoryPairRules, salesReviewRows: [], salesReviewGenerated: false });
      setStatus(`Đã tải ${nextCompanies.length} công ty bán ra còn lại sau KHH/152. Chọn công ty/hàng hóa rồi áp dụng trước khi review bán ra.`);
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
    const nextRows = companyRows.map((company, rowIndex) => (rowIndex === index ? { ...company, pending_process: pending } : company));
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), companyRows: nextRows, selectedCompanyIndex: index });
  }

  function bulkUpdatePendingCompanies(pending: boolean) {
    updateWorkflow(profile, {
      ...purchaseOutputInvalidation(),
      companyRows: companyRows.map((company) => ({ ...company, pending_process: pending })),
      selectedCompanyIndex: companyRows.length ? Math.max(0, selectedCompanyIndex) : -1,
    });
  }

  function applyCompanyAndProductChoices() {
    const nextCompanyRows = sortAppliedCompanyRows(companyRows.map((company) => {
      const process = company.pending_process ?? company.process ?? true;
      return { ...company, process, pending_process: process, committed_prefix: company.value || '' };
    }));
    const nextWorkflow = {
      ...workflow,
      ...purchaseOutputInvalidation(),
      companyRows: nextCompanyRows,
      selectedCompanyIndex: firstDisplayedCompanyIndex(nextCompanyRows),
      purchaseReviewRows: [],
      purchaseReviewGenerated: false,
    };
    updateWorkflow(profile, nextWorkflow);
    void saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu lựa chọn công ty, hàng hóa vào cấu hình. Review Mã VT sẽ tạo lại theo lựa chọn mới.');
  }

  function applySalesCompanyAndProductChoices() {
    const nextSalesCompanyRows = sortAppliedCompanyRows(salesCompanyRows.map((company) => {
      const process = company.pending_process ?? company.process ?? true;
      return { ...company, process, pending_process: process, committed_prefix: company.value || '' };
    }));
    const nextWorkflow = {
      ...workflow,
      ...salesOutputInvalidation(),
      salesCompanyRows: nextSalesCompanyRows,
      selectedSalesCompanyIndex: firstDisplayedCompanyIndex(nextSalesCompanyRows),
      salesReviewRows: [],
      salesReviewGenerated: false,
    };
    updateWorkflow(profile, nextWorkflow);
    void saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu lựa chọn công ty, hàng hóa bán ra. Review bán ra sẽ tạo lại theo lựa chọn mới.');
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
    const nextRows = salesCompanyRows.map((company, rowIndex) => (rowIndex === index ? { ...company, pending_process: pending } : company));
    updateWorkflow(profile, { ...salesOutputInvalidation(), salesCompanyRows: nextRows, selectedSalesCompanyIndex: index });
  }

  function bulkUpdateSalesPendingCompanies(pending: boolean) {
    updateWorkflow(profile, {
      ...salesOutputInvalidation(),
      salesCompanyRows: salesCompanyRows.map((company) => ({ ...company, pending_process: pending })),
      selectedSalesCompanyIndex: salesCompanyRows.length ? Math.max(0, selectedSalesCompanyIndex) : -1,
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
    const key = productKey(company.mst, productName);
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
    const key = productKey(company.mst, productName);
    updateWorkflow(profile, {
      selectedSalesCompanyIndex: companyIndex,
      salesProductCodeOverrides: { ...salesProductCodeOverrides, [key]: code.toUpperCase() },
      salesReviewRows: [],
      salesReviewGenerated: false,
      ...salesOutputInvalidation(),
    });
  }

  function updateCompanyPrefix(index: number, value: string) {
    updateWorkflow(profile, {
      selectedCompanyIndex: index,
      companyRows: companyRows.map((company, rowIndex) => (rowIndex === index ? { ...company, value: value.toUpperCase() } : company)),
    });
  }

  function updateSalesCompanyPrefix(index: number, value: string) {
    updateWorkflow(profile, {
      selectedSalesCompanyIndex: index,
      salesCompanyRows: salesCompanyRows.map((company, rowIndex) => (rowIndex === index ? { ...company, value: value.toUpperCase() } : company)),
    });
  }

  function updateIncludeCompanyPrefix(include: boolean) {
    updateWorkflow(profile, {
      includeCompanyPrefix: include,
      purchaseReviewRows: [],
      purchaseReviewGenerated: false,
      salesReviewRows: [],
      salesReviewGenerated: false,
      ...purchaseOutputInvalidation(),
    });
  }

  function updatePrefixStrategy(strategy: string) {
    updateWorkflow(profile, {
      prefixStrategy: strategy,
    });
  }

  function updatePrefixMstDigits(digits: number) {
    updateWorkflow(profile, {
      prefixMstDigits: Math.max(1, Math.min(10, digits)),
    });
  }

  async function refreshSalesProductPreviews() {
    if (!salesCompanyRows.length) return;
    setBusy(true);
    setStatus('Đang cập nhật Mã VT preview bán ra...');
    try {
      const previewCodes = await loadProductPreviewCodes(salesCompanyRows, wordRules, repeatedPhraseRemovals, 'sales');
      updateWorkflow(profile, { ...salesOutputInvalidation(), salesProductPreviewCodes: previewCodes, salesReviewRows: [], salesReviewGenerated: false });
      setStatus('Đã cập nhật Mã VT preview bán ra.');
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
    setBusy(true);
    setStatus('Đang tạo review Mã VT bán ra theo công ty/hàng hóa đã chọn...');
    try {
      const result = await createPurchaseReview(salesFile.saved_name, scope, wordRules, repeatedPhraseRemovals, reviewProducts as ReviewProduct[]);
      updateWorkflow(profile, { ...salesOutputInvalidation(), salesReviewRows: normalizeReviewRows(result.review_rows as ReviewRow[]), salesReviewGenerated: true, stage: 10 });
      setStatus(`Đã tạo ${result.review_rows.length} dòng Review bán ra. Chỉ dòng được tick mới gộp khi xuất file bán ra.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function saveWorkflowConfig(targetWorkflow: WorkflowState, successMessage = 'Đã lưu cấu hình hiện tại.') {
    setBusy(true);
    setStatus('Đang lưu cấu hình...');
    try {
      for (const payload of buildConfigPayloads(targetWorkflow)) {
        await saveVietmaxConfig(payload);
      }
      setStatus(successMessage);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function saveCurrentConfig() {
    void saveWorkflowConfig(workflow);
  }

  async function refreshProductPreviews() {
    if (!companyRows.length) return;
    setBusy(true);
    setStatus('Đang cập nhật Mã VT preview theo từ thay riêng và từ lặp...');
    try {
      const previewCodes = await loadProductPreviewCodes(companyRows, wordRules, repeatedPhraseRemovals);
      const nextWorkflow = { ...workflow, ...purchaseOutputInvalidation(), productPreviewCodes: previewCodes, purchaseReviewRows: [], purchaseReviewGenerated: false };
      updateWorkflow(profile, nextWorkflow);
      for (const payload of buildConfigPayloads(nextWorkflow)) {
        await saveVietmaxConfig(payload);
      }
      setStatus('Đã cập nhật Mã VT preview và lưu cấu hình.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function updateWordRule(index: number, field: 'from' | 'to', value: string) {
    const entries = Object.entries(wordRules);
    entries[index] = field === 'from' ? [value, entries[index]?.[1] || ''] : [entries[index]?.[0] || '', value];
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), wordRules: Object.fromEntries(entries.filter(([from]) => from.trim())), purchaseReviewRows: [], salesReviewRows: [], purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function addWordRule() {
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), wordRules: { ...wordRules, '': '' }, purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function updateRepeatedPhrase(index: number, value: string) {
    const next = repeatedPhraseRemovals.slice();
    next[index] = value;
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), repeatedPhraseRemovals: next.filter((item, rowIndex) => item.trim() || rowIndex === index), purchaseReviewRows: [], salesReviewRows: [], purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function addRepeatedPhrase() {
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), repeatedPhraseRemovals: [...repeatedPhraseRemovals, ''], purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function removeRepeatedPhrase(index: number) {
    const next = repeatedPhraseRemovals.slice();
    next.splice(index, 1);
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), repeatedPhraseRemovals: next, purchaseReviewRows: [], salesReviewRows: [], purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  async function downloadProcessedPurchase() {
    if (!purchaseFile) return;
    setBusy(true);
      setStatus('Đang tạo file mua vào đã xử lý Mã VT...');
    try {
      const result = await processVietmaxPurchase(purchaseFile.saved_name, purchaseFile.original_name, buildPurchaseProcessPayload(workflow));
      const saved = await saveBlob(result.blob, purchaseFile.original_name.replace(/\.(xlsx|xlsm)$/i, '_fdi.xlsx'));
      if (result.processedSavedName) applyProcessedPurchaseCache(profile, result.processedSavedName);
      setStatus(saved ? 'Đã xuất và lưu cache file mua vào đã xử lý. Bán ra sẽ dùng file này cho KHH/152.' : 'Đã tạo cache file mua vào đã xử lý, nhưng bạn đã hủy lưu bản xuất. Bán ra vẫn dùng cache này cho KHH/152.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function downloadProcessedSales() {
    if (!salesFile) return;
    setBusy(true);
      setStatus('Đang tạo file bán ra đã xử lý Mã VT...');
    try {
      const result = await processVietmaxPurchase(salesFile.saved_name, salesFile.original_name, buildSalesProcessPayload(workflow));
      const saved = await saveBlob(result.blob, salesFile.original_name.replace(/\.(xlsx|xlsm)$/i, '_fdi.xlsx'));
      if (result.processedSavedName) updateWorkflow(profile, { processedSalesSavedName: result.processedSavedName, inventoryAllocationJob: null, inventoryAllocationResult: null });
      setStatus(saved ? 'Đã xuất và lưu cache file bán ra đã xử lý. Có thể tiếp tục sang Stage 12 phân bổ tồn kho.' : 'Đã tạo cache file bán ra đã xử lý, nhưng bạn đã hủy lưu bản xuất. Stage 12 vẫn dùng cache này.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function changeProfile(nextProfile: ProfileKey) {
    setProfile(nextProfile);
    setStatus(`Đang xem shell ${profiles.find((item) => item.key === nextProfile)?.label ?? nextProfile}. Dữ liệu profile khác vẫn được giữ.`);
  }

  function updatePurchaseReviewScope(scope: 'all' | 'company') {
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), purchaseReviewScope: scope, purchaseReviewRows: [], purchaseReviewGenerated: false });
  }

  function updateSalesReviewScope(scope: 'all' | 'company') {
    updateWorkflow(profile, { ...salesOutputInvalidation(), salesReviewScope: scope, salesReviewRows: [], salesReviewGenerated: false });
  }

  function updateComparisonScope(value: string) {
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), comparisonScope: value, purchaseReviewRows: [], purchaseReviewGenerated: false, salesCompanyRows: [], selectedSalesCompanyIndex: -1, salesProductPreviewCodes: {}, salesProductCodeOverrides: {}, salesReviewRows: [], salesReviewGenerated: false });
  }

  function updateInventoryAllocationConfig(config: InventoryAllocationConfig) {
    updateWorkflow(profile, { inventoryAllocationConfig: config, inventoryAllocationJob: null, inventoryAllocationResult: null });
  }

  function updateOpeningStockFile(file: File | null) {
    updateWorkflow(profile, { openingStockFile: file, inventoryAllocationJob: null, inventoryAllocationResult: null });
  }

  async function runInventoryAllocation() {
    if (!processedPurchaseSavedName || !processedSalesSavedName || !salesFile) {
      setStatus('Cần xuất file mua vào và bán ra trước khi phân bổ tồn kho.');
      return;
    }
    const targetProfile = profile;
    setBusy(true);
    setStatus('Đang gửi dữ liệu phân bổ tồn kho...');
    try {
      const started = await startInventoryAllocation({ purchaseSavedName: processedPurchaseSavedName, salesSavedName: processedSalesSavedName, salesOriginalName: salesFile.original_name, openingFile: openingStockFile, config: inventoryAllocationConfig });
      let nextJob: InventoryAllocationJob = { status: 'queued', progress: 0, label: 'Đã gửi dữ liệu. Đang chờ backend xử lý...' };
      updateWorkflow(targetProfile, { inventoryAllocationJob: nextJob, inventoryAllocationResult: null });
      while (nextJob.status !== 'complete') {
        await sleep(900);
        nextJob = await getInventoryAllocationJob(started.analysis_job_id);
        updateWorkflow(targetProfile, { inventoryAllocationJob: nextJob, inventoryAllocationResult: nextJob.result ?? null });
        setStatus(nextJob.label || 'Đang phân bổ tồn kho...');
        if (nextJob.status === 'error') throw new Error(nextJob.error || nextJob.label || 'Phân bổ tồn kho thất bại.');
      }
      setStatus('Đã hoàn tất phân bổ tồn kho. Có thể tải file báo cáo Stage 12.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function downloadInventoryReport() {
    const jobId = inventoryAllocationResult?.job_id || inventoryAllocationJob?.result?.job_id;
    if (!jobId) return;
    setBusy(true);
    setStatus('Đang tải báo cáo phân bổ tồn kho...');
    try {
      const blob = await downloadInventoryAllocationReport(jobId);
      const filename = inventoryAllocationResult?.filename || inventoryAllocationJob?.result?.filename || 'phan_bo_ton_kho.xlsx';
      const saved = await saveBlob(blob, filename);
      setStatus(saved ? 'Đã lưu báo cáo phân bổ tồn kho.' : 'Đã hủy lưu báo cáo phân bổ tồn kho.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function addInventoryPair() {
    const id = `pair-${Date.now()}`;
    updateWorkflow(profile, {
      ...purchaseOutputInvalidation(),
      inventoryPairs: [...inventoryPairs, { id, ma_kho: '', tk_vat_tu: '' }],
      defaultInventoryPairId: defaultInventoryPairId || id,
    });
  }

  function updateInventoryPair(index: number, field: 'ma_kho' | 'tk_vat_tu', value: string) {
    updateWorkflow(profile, {
      ...purchaseOutputInvalidation(),
      inventoryPairs: inventoryPairs.map((pair, rowIndex) => (rowIndex === index ? { ...pair, [field]: value.toUpperCase() } : pair)),
    });
  }

  function removeInventoryPair(index: number) {
    const removed = inventoryPairs[index];
    const nextPairs = inventoryPairs.filter((_, rowIndex) => rowIndex !== index);
    updateWorkflow(profile, {
      ...purchaseOutputInvalidation(),
      inventoryPairs: nextPairs,
      defaultInventoryPairId: removed?.id === defaultInventoryPairId ? (nextPairs[0]?.id || '') : defaultInventoryPairId,
      inventoryPairRules: inventoryPairRules.filter((rule) => rule.pair_id !== removed?.id),
    });
  }

  function updateInventoryDefaults(update: Partial<Pick<WorkflowState, 'useDefaultInventoryPair' | 'defaultInventoryPairId'>>) {
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), ...update });
  }

  function addInventoryRule() {
    updateWorkflow(profile, {
      ...purchaseOutputInvalidation(),
      inventoryPairRules: [...inventoryPairRules, { source_col: 'M', operator: 'contains', value: '', pair_id: defaultInventoryPairId || inventoryPairs[0]?.id || '', enabled: true }],
    });
  }

  function updateInventoryRule(index: number, update: Partial<InventoryRule>) {
    updateWorkflow(profile, {
      ...purchaseOutputInvalidation(),
      inventoryPairRules: inventoryPairRules.map((rule, rowIndex) => (rowIndex === index ? { ...rule, ...update } : rule)),
    });
  }

  function removeInventoryRule(index: number) {
    updateWorkflow(profile, { ...purchaseOutputInvalidation(), inventoryPairRules: inventoryPairRules.filter((_, rowIndex) => rowIndex !== index) });
  }

  const nextStage = visibleStages[visibleStages.findIndex((item) => item.id === stage) + 1];

  return (
    <main className="desktop-shell">
      <section className={`app-card ${showLicenseBar ? '' : 'compact-flow'}`}>
        <header className="app-header">
          <div className="profile-toolbar" aria-label="Company profile controls">
            <label className="profile-dropdown"><span>Công ty áp dụng</span><select value={profile} disabled={busy} onChange={(event) => changeProfile(event.currentTarget.value as ProfileKey)}>{profiles.map((item) => <option key={item.key} value={item.key}>{item.label}</option>)}</select></label>
            <button type="button" className="btn-secondary" disabled={busy} onClick={saveCurrentConfig}>Lưu cấu hình</button>
          </div>
          <StageNavigation stages={visibleStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />
        </header>

        <div className="status-strip"><strong>Trạng thái</strong><span>{busy ? 'Đang xử lý... ' : ''}{status}</span></div>

        {showLicenseBar && (
          <section className="license-bar">
            <div>
              <strong>License</strong>
              <span>{license?.status || 'Đang kiểm tra license...'}</span>
              <span className={licenseReady ? 'ok-text' : 'warning-text'}>{licenseReady ? 'Được phép dùng.' : 'Cần license cho profile đang chọn.'}</span>
            </div>
            {!licenseReady && (
              <div className="license-form compact-form">
                <input placeholder="License server" value={licenseForm.server_url} onChange={(event) => setLicenseForm({ ...licenseForm, server_url: event.currentTarget.value })} />
                <input placeholder="Account" value={licenseForm.account_id} onChange={(event) => setLicenseForm({ ...licenseForm, account_id: event.currentTarget.value })} />
                <input placeholder="License key" type="password" value={licenseForm.license_key} onChange={(event) => setLicenseForm({ ...licenseForm, license_key: event.currentTarget.value })} />
                <button type="button" disabled={busy} onClick={submitLicense}>Kích hoạt</button>
              </div>
            )}
            {license?.activated && <button type="button" className="btn-secondary" disabled={busy} onClick={refreshLicense}>Tải lại license</button>}
          </section>
        )}

        <section className="stage-frame">
          <div className="stage-body">
            {profile !== 'vietmax' ? renderProfileStage() : renderVietmaxStage()}
          </div>
        </section>

        <footer className="action-bar">
          <button type="button" className="btn-secondary" disabled={stage === 1 || busy} onClick={goBack}>Quay lại</button>
          <button type="button" className="btn-danger" disabled={busy} onClick={resetWorkflow}>Làm lại</button>
          <div className="action-spacer" />
          <button type="button" disabled={busy || !nextStage || !canEnterStage(nextStage.id)} onClick={goNext}>Tiếp tục</button>
        </footer>
      </section>
    </main>
  );

  function renderVietmaxStage() {
    switch (stage) {
      case 1:
        return <UploadStage title="HD mua vào" summary={purchaseFile} disabled={busy || !licenseReady} onUpload={(file) => upload('purchase', file)} />;
      case 2:
        return <MappingStage summary={purchaseFile} phase="purchase" scope={comparisonScope} setScope={updateComparisonScope} />;
      case 3:
        if (busy && !companyRows.length) return <LoadingStage title="Đang tải danh sách công ty" detail="Đang đọc workbook và gom công ty/MST/hàng hóa..." />;
        return <CompanyRulesStage companies={companyRows} selectedCompanyIndex={selectedCompanyIndex} productPreviewCodes={productPreviewCodes} productCodeOverrides={productCodeOverrides} wordRules={wordRules} repeatedPhrases={repeatedPhraseRemovals} inventoryPairs={inventoryPairs} useDefaultInventoryPair={useDefaultInventoryPair} defaultInventoryPairId={defaultInventoryPairId} inventoryPairRules={inventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy={prefixStrategy} prefixMstDigits={prefixMstDigits} onIncludeCompanyPrefixChange={updateIncludeCompanyPrefix} onCompanyPrefixChange={updateCompanyPrefix} onPrefixStrategyChange={updatePrefixStrategy} onPrefixMstDigitsChange={updatePrefixMstDigits} onApplyPrefixPresetToAll={(strategy) => {
            updateWorkflow(profile, {
              companyRows: companyRows.map((company) => ({ ...company, value: computePresetPrefix(company, strategy, prefixMstDigits) })),
            });
          }} onCompanySelect={selectCompany} onCompanyChange={updatePendingCompany} onBulkCompanyChange={bulkUpdatePendingCompanies} onProductChange={updateCompanyProduct} onProductCodeChange={updateProductCode} onApplyChoices={applyCompanyAndProductChoices} onRefreshPreviews={refreshProductPreviews} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeatedPhrase} onAddRepeated={addRepeatedPhrase} onRemoveRepeated={removeRepeatedPhrase} onAddInventoryPair={addInventoryPair} onInventoryPairChange={updateInventoryPair} onRemoveInventoryPair={removeInventoryPair} onInventoryDefaultsChange={updateInventoryDefaults} onAddInventoryRule={addInventoryRule} onInventoryRuleChange={updateInventoryRule} onRemoveInventoryRule={removeInventoryRule} />;
      case 4:
        if (busy) return <LoadingStage title="Đang tạo Review Mã VT mua vào" detail="Đang so sánh tên hàng và dựng danh sách mã cần kiểm tra..." />;
        return <ReviewStage rows={purchaseReviewRows} onRun={runPurchaseReview} onApply={applyReviewChoices} disabled={!purchaseFile || busy} onRowChange={updateReviewRow} onBulkChange={bulkUpdateReviewRows} title="Review Mã VT mua vào" empty="Tạo danh sách review sau khi chọn cột và công ty. Backend dùng logic Vietmax từ mã vietmax.xlsx; khác biệt chỉ do hoa/thường hoặc khoảng trắng sẽ không xuất hiện ở đây." reviewScope={purchaseReviewScope} onReviewScopeChange={updatePurchaseReviewScope} />;
      case 5:
        if (busy) return <LoadingStage title="Đang tạo file mua vào" detail="Đang xử lý workbook và tạo cache file mua vào để dùng cho các stage bán ra..." />;
        return <ProcessStage title="Tạo file mua vào" detail="Xuất file mua vào bằng logic Vietmax. Chỉ các công ty/hàng hóa đã chọn và các dòng Review Mã VT được tick mới tạo manual override. File đã xử lý sẽ được cache để dùng cho các stage bán ra." buttonLabel="Xuất file mua vào" disabled={busy || !purchaseFile} onProcess={downloadProcessedPurchase} />;
      case 6:
        return <UploadStage title="HD bán ra" summary={salesFile} disabled={busy || !licenseReady} onUpload={(file) => upload('sales', file)} />;
      case 7:
        return <MappingStage summary={salesFile} phase="sales" scope={comparisonScope} setScope={updateComparisonScope} />;
      case 8:
        if (busy) return <LoadingStage title="Đang khớp mua vào / bán ra" detail={processedPurchaseSavedName ? 'Đang so sánh hàng bán ra với file mua vào đã xử lý và áp dụng cấu hình khớp đã lưu...' : 'Đang tạo cache mua vào rồi khớp với file bán ra...'} />;
        return <MatchStage rows={matches} disabled={!salesFile || busy} onRun={runSalesMatch} onSave={saveMatchChoices} onToggle={toggleMatch} onBulkToggle={bulkToggleMatches} onConversionChange={updateMatchConversion} autoRun={Boolean(salesFile) && matches.length === 0 && !busy && !salesMatchGenerated} emptyMessage={processedPurchaseSavedName ? undefined : 'Chưa có cache file mua vào đã xử lý. Ứng dụng sẽ tự tạo cache trước khi khớp mua/bán.'} />;
      case 9:
        if (busy && !salesCompanyRows.length) return <LoadingStage title="Đang tải danh sách công ty bán ra" detail="Đang lọc các hàng hóa chưa khớp KHH/152 và gom theo công ty..." />;
        return <CompanyRulesStage companies={salesCompanyRows} selectedCompanyIndex={selectedSalesCompanyIndex} productPreviewCodes={salesProductPreviewCodes} productCodeOverrides={salesProductCodeOverrides} wordRules={wordRules} repeatedPhrases={repeatedPhraseRemovals} inventoryPairs={inventoryPairs} useDefaultInventoryPair={useDefaultInventoryPair} defaultInventoryPairId={defaultInventoryPairId} inventoryPairRules={inventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy={prefixStrategy} prefixMstDigits={prefixMstDigits} onIncludeCompanyPrefixChange={updateIncludeCompanyPrefix} onCompanyPrefixChange={updateSalesCompanyPrefix} onPrefixStrategyChange={updatePrefixStrategy} onPrefixMstDigitsChange={updatePrefixMstDigits} onApplyPrefixPresetToAll={(strategy) => {
            updateWorkflow(profile, {
              salesCompanyRows: salesCompanyRows.map((company) => ({ ...company, value: computePresetPrefix(company, strategy, prefixMstDigits) })),
            });
          }} onCompanySelect={selectSalesCompany} onCompanyChange={updateSalesPendingCompany} onBulkCompanyChange={bulkUpdateSalesPendingCompanies} onProductChange={updateSalesCompanyProduct} onProductCodeChange={updateSalesProductCode} onApplyChoices={applySalesCompanyAndProductChoices} onRefreshPreviews={refreshSalesProductPreviews} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeatedPhrase} onAddRepeated={addRepeatedPhrase} onRemoveRepeated={removeRepeatedPhrase} onAddInventoryPair={addInventoryPair} onInventoryPairChange={updateInventoryPair} onRemoveInventoryPair={removeInventoryPair} onInventoryDefaultsChange={updateInventoryDefaults} onAddInventoryRule={addInventoryRule} onInventoryRuleChange={updateInventoryRule} onRemoveInventoryRule={removeInventoryRule} />;
      case 10:
        if (busy) return <LoadingStage title="Đang tạo Review Mã VT bán ra" detail="Đang tạo danh sách review theo công ty/hàng hóa bán ra đã áp dụng..." />;
        return <ReviewStage rows={salesReviewRows} onRun={runSalesReview} onApply={applySalesReviewChoices} disabled={!salesFile || busy} onRowChange={updateSalesReviewRow} onBulkChange={bulkUpdateSalesReviewRows} title="Review Mã VT bán ra" empty="Tạo danh sách review sau khi chọn công ty bán ra. Chỉ hàng hóa/công ty đã chọn mới được đưa vào review và xuất Excel." reviewScope={salesReviewScope} onReviewScopeChange={updateSalesReviewScope} />;
      case 11:
        if (busy) return <LoadingStage title="Đang tạo file bán ra" detail="Đang xử lý workbook bán ra, áp dụng khớp mua vào và lưu cache cho phân bổ tồn kho..." />;
        return <ProcessStage title="Tạo file bán ra" detail="Xuất file bán ra bằng logic Vietmax. Chỉ công ty/hàng hóa đã chọn và dòng Review bán ra được tick mới áp dụng gộp tên hàng." buttonLabel="Xuất file bán ra" disabled={busy || !salesFile} onProcess={downloadProcessedSales} />;
      case 12:
        if (busy) return <LoadingStage title="Đang phân bổ tồn kho" detail={inventoryAllocationJob?.label || 'Đang chạy phân bổ từ file mua vào và bán ra đã xử lý...'} />;
        return <InventoryAllocationStage purchaseFile={purchaseFile} salesFile={salesFile} processedPurchaseSavedName={processedPurchaseSavedName} processedSalesSavedName={processedSalesSavedName} openingStockFile={openingStockFile} config={inventoryAllocationConfig} job={inventoryAllocationJob} result={inventoryAllocationResult} busy={busy} onOpeningStockFileChange={updateOpeningStockFile} onConfigChange={updateInventoryAllocationConfig} onProcess={runInventoryAllocation} onDownload={downloadInventoryReport} />;
      default:
        return null;
    }
  }

  function renderProfileStage() {
    return <ProfilePlaceholder profile={selectedProfile.label} stage={currentStage} />;
  }
}

function StageNavigation({ stages, stage, busy, canEnterStage, goToStage }: { stages: StageDefinition[]; stage: StageId; busy: boolean; canEnterStage: (target: StageId) => boolean; goToStage: (target: StageId) => void }) {
  const purchaseStages = stages.filter((item) => item.phase === 'purchase');
  const salesStages = stages.filter((item) => item.phase === 'sales');
  const inventoryStages = stages.filter((item) => item.phase === 'inventory');
  const shellStages = stages.filter((item) => item.phase === 'generic' || item.phase === 'price');

  if (purchaseStages.length || salesStages.length || inventoryStages.length) {
    return (
      <div className="stage-groups" aria-label="Stage navigation">
        <StageGroup title="Mua vào" stages={purchaseStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />
        <StageGroup title="Bán ra" stages={salesStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />
        <StageGroup title="Tồn kho" stages={inventoryStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />
      </div>
    );
  }

  return (
    <div className="stage-groups" aria-label="Stage navigation">
      <StageGroup title="PySide shell" stages={shellStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />
    </div>
  );
}

function StageGroup({ title, stages, stage, busy, canEnterStage, goToStage }: { title: string; stages: StageDefinition[]; stage: StageId; busy: boolean; canEnterStage: (target: StageId) => boolean; goToStage: (target: StageId) => void }) {
  return (
    <div className="stage-group">
      <span className="stage-group-label">{title}</span>
      <div className="stage-group-pills">
        {stages.map((item) => (
          <button
            key={item.id}
            className={`step-pill ${item.id === stage ? 'active' : ''} ${item.phase}`}
            disabled={!canEnterStage(item.id) || busy}
            onClick={() => goToStage(item.id)}
            type="button"
            title={`${item.id}. ${item.label}`}
          >
            <span>{item.id}.</span> {item.short}
          </button>
        ))}
      </div>
    </div>
  );
}

function phaseLabel(phase: StagePhase) {
  if (phase === 'purchase') return 'Mua vào';
  if (phase === 'sales') return 'Bán ra';
  if (phase === 'inventory') return 'Tồn kho';
  if (phase === 'price') return 'Lọc đơn giá';
  return 'PySide shell';
}

function UploadStage({ title, summary, disabled, onUpload }: { title: string; summary: UploadSummary | null; disabled: boolean; onUpload: (file: File | undefined) => void }) {
  const stepNumber = title.includes('bán') ? 6 : 1;
  const uploadLabel = `Chọn ${title}`;
  return (
    <div className="center-stage">
      <span className="upload-step-badge">BƯỚC {stepNumber}</span>
      <p className="description">Tải lên file Excel <strong>.xlsx</strong> hoặc <strong>.xlsm</strong> cho <strong>{title}</strong>.</p>
      <label className={`upload-label ${summary ? 'has-file' : ''}`}>
        <input className="upload-input" type="file" accept=".xlsx,.xlsm" disabled={disabled} onChange={(event) => onUpload(event.currentTarget.files?.[0])} />
        <span className="upload-mark">XLSX</span>
        <strong>{summary?.original_name || 'Kéo thả hoặc chọn file Excel'}</strong>
        <small>{summary ? `${summary.columns.length} cột đã đọc` : 'Hệ thống sẽ đọc bảng, nhận diện cột và giữ nguyên file gốc.'}</small>
        <span className="upload-button">{summary ? `Đổi ${title}` : uploadLabel}</span>
      </label>
    </div>
  );
}

function MappingStage({ summary, phase, scope, setScope }: { summary: UploadSummary | null; phase: 'purchase' | 'sales'; scope: string; setScope: (value: string) => void }) {
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

function ReviewStage({ rows, onRun, onApply, onRowChange, onBulkChange, disabled = false, title = 'Review Mã VT mua vào', empty = 'Chưa có dòng review.', reviewScope = 'all', onReviewScopeChange }: { rows: ReviewRow[]; onRun?: () => void; onApply?: () => void; onRowChange?: (index: number, update: Partial<ReviewRow>) => void; onBulkChange?: (indices: number[], confirmed: boolean) => void; disabled?: boolean; title?: string; empty?: string; reviewScope?: 'all' | 'company'; onReviewScopeChange?: (scope: 'all' | 'company') => void }) {
  const [collapsedGroups, setCollapsedGroups] = useState<Record<string, boolean>>({});
  const groups = reviewDisplayGroups(rows, reviewScope);
  const dimensionIndices = groups[0]?.rows.map(({ index }) => index) ?? [];
  const otherIndices = groups[1]?.rows.map(({ index }) => index) ?? [];
  const allIndices = rows.map((_, index) => index);
  const canBulkUpdate = Boolean(onBulkChange && rows.length && !disabled);
  const bulkUpdate = (indices: number[], confirmed: boolean) => onBulkChange?.(indices, confirmed);
  const toggleGroup = (groupTitle: string) => setCollapsedGroups((current) => ({ ...current, [groupTitle]: !current[groupTitle] }));

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
            <div className="review-bulk-group">
              <button type="button" className="btn-secondary" disabled={!canBulkUpdate || !dimensionIndices.length} onClick={() => bulkUpdate(dimensionIndices, false)}>Bỏ chọn tất cả (gần giống kích thước)</button>
              <button type="button" className="btn-secondary" disabled={!canBulkUpdate || !dimensionIndices.length} onClick={() => bulkUpdate(dimensionIndices, true)}>Chọn (gần giống kích thước)</button>
            </div>
            <div className="review-bulk-group">
              <button type="button" className="btn-secondary" disabled={!canBulkUpdate || !otherIndices.length} onClick={() => bulkUpdate(otherIndices, false)}>Bỏ chọn tất cả (khác)</button>
              <button type="button" className="btn-secondary" disabled={!canBulkUpdate || !otherIndices.length} onClick={() => bulkUpdate(otherIndices, true)}>Chọn tất cả (khác)</button>
            </div>
            <div className="review-bulk-group all-review-bulk-group">
              <button type="button" className="btn-secondary" disabled={!canBulkUpdate} onClick={() => bulkUpdate(allIndices, false)}>Bỏ chọn tất cả</button>
              <button type="button" className="btn-secondary" disabled={!canBulkUpdate} onClick={() => bulkUpdate(allIndices, true)}>Chọn tất cả</button>
            </div>
          </div>
          <div className="review-primary-actions">
            {onRun && <button type="button" disabled={disabled} onClick={onRun}>Tạo danh sách review</button>}
            {onApply && <button type="button" disabled={disabled || !rows.length} onClick={onApply}>Áp dụng Review Mã VT</button>}
          </div>
        </div>
      </div>
      {!rows.length ? <p className="muted">{empty}</p> : <div className="inner-scroll"><ReviewTable groups={groups} collapsedGroups={collapsedGroups} onToggleGroup={toggleGroup} onRowChange={onRowChange} reviewScope={reviewScope} /></div>}
    </div>
  );
}

function MatchStage({ rows, disabled, onRun, onSave, onToggle, onBulkToggle, onConversionChange, autoRun, emptyMessage = 'Chưa có dòng khớp. Dữ liệu sẽ được giữ khi quay lại stage 7 hoặc sang stage 9.' }: { rows: MatchRow[]; disabled: boolean; onRun: () => void; onSave?: () => void; onToggle: (index: number, confirmed: boolean) => void; onBulkToggle: (confirmed: boolean) => void; onConversionChange: (index: number, salesQty: string, purchaseQty: string) => void; autoRun?: boolean; emptyMessage?: string }) {
  const [autoRunStarted, setAutoRunStarted] = useState(false);

  useEffect(() => {
    if (autoRun && !disabled && rows.length === 0 && !autoRunStarted) {
      setAutoRunStarted(true);
      onRun();
    }
  }, [autoRun, autoRunStarted, disabled, rows.length, onRun]);

  useEffect(() => {
    if (!autoRun || rows.length > 0) setAutoRunStarted(false);
  }, [autoRun, rows.length]);

  return (
    <div className="list-stage">
      <div className="stage-toolbar">
        <p>Khớp HD mua vào đã xử lý với HD bán ra KHH/152.</p>
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

function ConfigModal({
  isOpen,
  onClose,
  prefixStrategy,
  prefixMstDigits,
  onPrefixStrategyChange,
  onPrefixMstDigitsChange,
  onApplyPrefixPreset,
  wordRules,
  repeatedPhrases,
  inventoryPairs,
  useDefaultInventoryPair,
  defaultInventoryPairId,
  inventoryPairRules,
  onWordRuleChange,
  onAddWordRule,
  onRepeatedChange,
  onAddRepeated,
  onRemoveRepeated,
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
  prefixStrategy: string;
  prefixMstDigits: number;
  onPrefixStrategyChange?: (strategy: string) => void;
  onPrefixMstDigitsChange?: (digits: number) => void;
  onApplyPrefixPreset?: (strategy: 'last_2_words' | 'last_3_mst' | '2_words_mst') => void;
  wordRules: Record<string, string>;
  repeatedPhrases: string[];
  inventoryPairs: InventoryPair[];
  useDefaultInventoryPair: boolean;
  defaultInventoryPairId: string;
  inventoryPairRules: InventoryRule[];
  onWordRuleChange: (index: number, field: 'from' | 'to', value: string) => void;
  onAddWordRule: () => void;
  onRepeatedChange: (index: number, value: string) => void;
  onAddRepeated: () => void;
  onRemoveRepeated: (index: number) => void;
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
  const [activeTab, setActiveTab] = useState<'prefix' | 'words' | 'repeat' | 'inventory'>('prefix');
  const wordEntries = Object.entries(wordRules);
  const activePrefixStrategy = normalizedPrefixStrategy(prefixStrategy);
  if (!isOpen) return null;
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-content" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Cấu hình nâng cao</h2>
          <button className="modal-close" onClick={onClose}>×</button>
        </div>
        <div className="modal-body">
          <div className="tab-list">
            <button className={`tab-button ${activeTab === 'prefix' ? 'active' : ''}`} onClick={() => setActiveTab('prefix')}>Prefix</button>
            <button className={`tab-button ${activeTab === 'words' ? 'active' : ''}`} onClick={() => setActiveTab('words')}>Từ riêng</button>
            <button className={`tab-button ${activeTab === 'repeat' ? 'active' : ''}`} onClick={() => setActiveTab('repeat')}>Từ lặp</button>
            <button className={`tab-button ${activeTab === 'inventory' ? 'active' : ''}`} onClick={() => setActiveTab('inventory')}>Phân kho</button>
          </div>
          <div className={`tab-panel ${activeTab === 'prefix' ? 'active' : ''}`}>
            <div className="stage-toolbar compact-toolbar"><p>Cấu hình prefix công ty</p></div>
            <div className="prefix-config-panel">
              <label className="prefix-choice">
                <input type="radio" name="prefix-mode" checked={activePrefixStrategy === 'last_2_words'} onChange={() => onPrefixStrategyChange?.('last_2_words')} />
                <span>2 từ</span>
              </label>
              <label className="prefix-choice">
                <input type="radio" name="prefix-mode" checked={activePrefixStrategy === 'last_3_mst'} onChange={() => onPrefixStrategyChange?.('last_3_mst')} />
                <span>MST</span>
              </label>
              <label className="prefix-choice">
                <input type="radio" name="prefix-mode" checked={activePrefixStrategy === '2_words_mst'} onChange={() => onPrefixStrategyChange?.('2_words_mst')} />
                <span>2 từ + MST</span>
              </label>
              <label className="prefix-mst-length">
                <span>Số ký tự MST</span>
                <input type="number" min={1} max={10} value={prefixMstDigits} disabled={activePrefixStrategy === 'last_2_words'} onChange={(event) => onPrefixMstDigitsChange?.(parseInt(event.currentTarget.value) || 3)} />
              </label>
            </div>
            <div className="tab-apply-bar"><button type="button" disabled={busy} onClick={() => { onApplyPrefixPreset?.(activePrefixStrategy); }}>Áp dụng prefix</button></div>
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
          <div className={`tab-panel ${activeTab === 'inventory' ? 'active' : ''}`}>
            <InventoryPairEditor pairs={inventoryPairs} useDefault={useDefaultInventoryPair} defaultPairId={defaultInventoryPairId} rules={inventoryPairRules} busy={busy ?? false} onAddPair={onAddInventoryPair} onPairChange={onInventoryPairChange} onRemovePair={onRemoveInventoryPair} onDefaultsChange={onInventoryDefaultsChange} onAddRule={onAddInventoryRule} onRuleChange={onInventoryRuleChange} onRemoveRule={onRemoveInventoryRule} />
            <div className="tab-apply-bar"><button type="button" disabled={busy} onClick={() => { onRefreshPreviews?.(); }}>Áp dụng phân kho</button></div>
          </div>
        </div>
        <div className="modal-footer">
          <button type="button" className="btn-secondary" onClick={onClose}>Đóng</button>
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
  rows: Array<{ row: ReviewRow; index: number }>;
  className: string;
};

function CompanyRulesStage({ companies, selectedCompanyIndex, productPreviewCodes, productCodeOverrides, wordRules, repeatedPhrases, inventoryPairs, useDefaultInventoryPair, defaultInventoryPairId, inventoryPairRules, busy, showCompanyPrefixControls = false, includeCompanyPrefix = false, prefixStrategy = 'last_2_words', prefixMstDigits = 3, onIncludeCompanyPrefixChange, onCompanyPrefixChange, onPrefixStrategyChange, onPrefixMstDigitsChange, onApplyPrefixPresetToAll, onCompanySelect, onCompanyChange, onBulkCompanyChange, onProductChange, onProductCodeChange, onApplyChoices, onRefreshPreviews, onWordRuleChange, onAddWordRule, onRepeatedChange, onAddRepeated, onRemoveRepeated, onAddInventoryPair, onInventoryPairChange, onRemoveInventoryPair, onInventoryDefaultsChange, onAddInventoryRule, onInventoryRuleChange, onRemoveInventoryRule }: { companies: CompanyRow[]; selectedCompanyIndex: number; productPreviewCodes: Record<string, string>; productCodeOverrides: Record<string, string>; wordRules: Record<string, string>; repeatedPhrases: string[]; inventoryPairs: InventoryPair[]; useDefaultInventoryPair: boolean; defaultInventoryPairId: string; inventoryPairRules: InventoryRule[]; busy: boolean; showCompanyPrefixControls?: boolean; includeCompanyPrefix?: boolean; prefixStrategy?: string; prefixMstDigits?: number; onIncludeCompanyPrefixChange?: (include: boolean) => void; onCompanyPrefixChange?: (index: number, value: string) => void; onPrefixStrategyChange?: (strategy: string) => void; onPrefixMstDigitsChange?: (digits: number) => void; onApplyPrefixPresetToAll?: (strategy: 'last_2_words' | 'last_3_mst' | '2_words_mst') => void; onCompanySelect: (index: number) => void; onCompanyChange: (index: number, pending: boolean) => void; onBulkCompanyChange?: (pending: boolean) => void; onProductChange: (companyIndex: number, productName: string, selected: boolean) => void; onProductCodeChange: (companyIndex: number, productName: string, code: string) => void; onApplyChoices: () => void; onRefreshPreviews: () => void; onWordRuleChange: (index: number, field: 'from' | 'to', value: string) => void; onAddWordRule: () => void; onRepeatedChange: (index: number, value: string) => void; onAddRepeated: () => void; onRemoveRepeated: (index: number) => void; onAddInventoryPair: () => void; onInventoryPairChange: (index: number, field: 'ma_kho' | 'tk_vat_tu', value: string) => void; onRemoveInventoryPair: (index: number) => void; onInventoryDefaultsChange: (update: Partial<Pick<WorkflowState, 'useDefaultInventoryPair' | 'defaultInventoryPairId'>>) => void; onAddInventoryRule: () => void; onInventoryRuleChange: (index: number, update: Partial<InventoryRule>) => void; onRemoveInventoryRule: (index: number) => void }) {
  const [showConfigModal, setShowConfigModal] = useState(false);
  const [collapsedSections, setCollapsedSections] = useState({ wordRules: true, repeatedPhrases: true });
  const wordEntries = Object.entries(wordRules);
  const safeSelectedIndex = selectedCompanyIndex >= 0 && selectedCompanyIndex < companies.length ? selectedCompanyIndex : -1;
  const selectedCompany = safeSelectedIndex >= 0 ? companies[safeSelectedIndex] : undefined;
  const selectedProducts = new Set(selectedCompany?.selected_product_names.length ? selectedCompany.selected_product_names : selectedCompany?.all_products.map((product) => product.name));
  const productRows = selectedCompany?.all_products ?? [];
  const longProducts = selectedCompany ? productRows.filter((product) => selectedProducts.has(product.name) && productDisplayCode(selectedCompany, product.name, productPreviewCodes, productCodeOverrides, includeCompanyPrefix).length > maxCodeLength) : [];
  const normalProducts = selectedCompany ? productRows.filter((product) => !longProducts.includes(product)) : [];
  const companyCodeLong = Boolean(selectedCompany?.value && selectedCompany.value.length > maxCodeLength);
  const duplicatePrefixSet = new Set<string>();
  
  // Detect duplicates from committed state only; drafts apply after the main Apply button.
  const prefixCounts = new Map<string, number>();
  companies.forEach(c => {
    const isActive = c.process ?? true;
    const normalizedPrefix = committedCompanyPrefix(c);
    if (isActive && normalizedPrefix) {
      prefixCounts.set(normalizedPrefix, (prefixCounts.get(normalizedPrefix) || 0) + 1);
    }
  });
  const duplicatePrefixes = Array.from(prefixCounts.entries()).filter(([, count]) => count > 1);
  duplicatePrefixes.forEach(([prefix]) => duplicatePrefixSet.add(prefix));
  const groups = companyDisplayGroups(companies, duplicatePrefixSet);
  const activePrefixStrategy = normalizedPrefixStrategy(prefixStrategy);
  const allProductNames = selectedCompany?.all_products.map(p => p.name) || [];
  const allSelected = allProductNames.length > 0 && allProductNames.every(name => selectedProducts.has(name));
  
  const renderProductRow = (product: { name: string; count?: number }, forceWarning = false) => {
    if (!selectedCompany) return null;
    const code = productDisplayCode(selectedCompany, product.name, productPreviewCodes, productCodeOverrides, includeCompanyPrefix);
    const selected = selectedProducts.has(product.name);
    const longCode = selected && code.length > maxCodeLength;
    return <tr key={product.name} className={longCode || forceWarning ? 'danger-row big-select-row' : 'big-select-row'}><td><input type="checkbox" checked={selected} onChange={(event) => onProductChange(safeSelectedIndex, product.name, event.currentTarget.checked)} /></td><td>{product.name}</td><td>{product.count ?? ''}</td><td><input className="code-edit" value={code} onChange={(event) => onProductCodeChange(safeSelectedIndex, product.name, event.currentTarget.value)} /></td></tr>;
  };

  return (
    <div className="company-workspace">
      <div className={`compact-rules-panel ${showCompanyPrefixControls ? 'prefix-enabled' : ''}`}>
        <button type="button" className="btn-secondary" onClick={() => setShowConfigModal(true)}>Cấu hình nâng cao</button>
        {showCompanyPrefixControls && <section className="company-prefix-card compact-rule-card">
          <label className="inline-check"><input type="checkbox" checked={includeCompanyPrefix} onChange={(event) => onIncludeCompanyPrefixChange?.(event.currentTarget.checked)} /> Dùng prefix công ty</label>
          {includeCompanyPrefix && <>
            <div className="prefix-strategy-row">
              <label>Chiến lược prefix:</label>
              <select value={activePrefixStrategy} onChange={(event) => onPrefixStrategyChange?.(event.currentTarget.value)}>
                <option value="last_2_words">2 chữ cuối (bỏ tỉnh/VN)</option>
                <option value="last_3_mst">{prefixMstDigits} số cuối MST</option>
                <option value="2_words_mst">2 chữ + {prefixMstDigits} số MST</option>
              </select>
              {(activePrefixStrategy === 'last_3_mst' || activePrefixStrategy === '2_words_mst') && <>
                <label>Số ký tự MST:</label>
                <input type="number" min={1} max={10} value={prefixMstDigits} onChange={(event) => onPrefixMstDigitsChange?.(parseInt(event.currentTarget.value) || 3)} />
              </>}
            </div>
            <div className="prefix-quick-actions">
              <button type="button" className="prefix-apply-all-button" disabled={busy || !companies.length} onClick={() => onApplyPrefixPresetToAll?.('last_2_words')}>Áp 2 từ</button>
              <button type="button" className="prefix-apply-all-button" disabled={busy || !companies.length} onClick={() => onApplyPrefixPresetToAll?.('last_3_mst')}>Áp MST</button>
              <button type="button" className="prefix-apply-all-button" disabled={busy || !companies.length} onClick={() => onApplyPrefixPresetToAll?.('2_words_mst')}>Áp 2 từ + MST</button>
            </div>
            <p className="muted">Sửa 3 prefix trong modal Cấu hình nâng cao. Danh sách chỉ sắp lại khi bấm Áp dụng lựa chọn công ty và hàng hóa.</p>
          </>}
        </section>}
      </div>

      <ConfigModal 
        isOpen={showConfigModal} 
        onClose={() => setShowConfigModal(false)} 
        prefixStrategy={activePrefixStrategy}
        prefixMstDigits={prefixMstDigits}
        onPrefixStrategyChange={onPrefixStrategyChange}
        onPrefixMstDigitsChange={onPrefixMstDigitsChange}
        onApplyPrefixPreset={onApplyPrefixPresetToAll}
        wordRules={wordRules}
        repeatedPhrases={repeatedPhrases}
        inventoryPairs={inventoryPairs}
        useDefaultInventoryPair={useDefaultInventoryPair}
        defaultInventoryPairId={defaultInventoryPairId}
        inventoryPairRules={inventoryPairRules}
        onWordRuleChange={onWordRuleChange}
        onAddWordRule={onAddWordRule}
        onRepeatedChange={onRepeatedChange}
        onAddRepeated={onAddRepeated}
        onRemoveRepeated={onRemoveRepeated}
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
              <button type="button" disabled={busy || !companies.length} onClick={onApplyChoices}>Áp dụng lựa chọn công ty và hàng hóa</button>
            </div>
          </div>
          {duplicatePrefixes.length > 0 && (
            <div className="duplicate-prefix-warning">
              <p className="warning-text">Cảnh báo: Prefix trùng lặp</p>
              <ul>
                {duplicatePrefixes.map(([prefix, count]) => {
                  const dupCompanies = companies.filter((c) => (c.process ?? true) && committedCompanyPrefix(c) === prefix);
                  return (
                    <li key={prefix}>
                      <strong>{prefix}</strong> ({count} công ty): {dupCompanies.map((c) => c.company).join(', ')}
                    </li>
                  );
                })}
              </ul>
            </div>
          )}
          {!companies.length ? <p className="muted">Đang chờ dữ liệu công ty. Danh sách sẽ tự tải khi vào stage này sau khi đã chọn file mua vào.</p> : <div className="inner-scroll company-table-scroll"><table className="company-table grouped-company-table"><thead><tr><th>Dùng</th><th>Công ty</th><th>MST</th>{showCompanyPrefixControls && <th>Prefix</th>}<th>Số hàng</th></tr></thead><tbody>{groups.map((group) => <CompanyGroupRows key={group.title} group={group} safeSelectedIndex={safeSelectedIndex} showPrefix={showCompanyPrefixControls} onCompanySelect={onCompanySelect} onCompanyChange={onCompanyChange} onCompanyPrefixChange={onCompanyPrefixChange} />)}</tbody></table><div className="scroll-bottom-spacer" aria-hidden="true" /></div>}
        </div>

        <div className="list-stage product-list-card">
          <div className="stage-toolbar"><p>{selectedCompany ? `Hàng hóa của ${selectedCompany.company} - MST ${selectedCompany.mst}` : 'Hàng hóa / mã VT preview'}</p></div>
          {!selectedCompany ? <p className="muted">Chọn một dòng công ty để xem danh sách hàng hóa.</p> : <>
            {(companyCodeLong || longProducts.length > 0) && <div className="long-code-warning"><p className="warning-text compact-warning">Cảnh báo: mã vượt {maxCodeLength} ký tự sẽ bị cắt đuôi khi xuất file.</p><table className="product-table warning-code-table"><thead><tr><th>Loại</th><th>Tên</th><th colSpan={2}>Mã đang vượt giới hạn</th></tr></thead><tbody>{companyCodeLong && <tr className="danger-row"><td>Công ty</td><td>{selectedCompany.company}</td><td colSpan={2}>{selectedCompany.value}</td></tr>}{longProducts.map((product) => renderProductRow(product, true))}</tbody></table></div>}
            <div className="inner-scroll product-table-scroll"><table className="product-table"><thead><tr><th>Xử lý</th><th>Tên hàng hóa</th><th>Dòng</th><th>Mã VT xem trước / sửa tay</th></tr></thead><tbody>{normalProducts.map((product) => renderProductRow(product))}</tbody></table><div className="scroll-bottom-spacer" aria-hidden="true" /></div>
          </>}
        </div>
      </div>
    </div>
  );
}

function CompanyGroupRows({ group, safeSelectedIndex, showPrefix = false, onCompanySelect, onCompanyChange, onCompanyPrefixChange }: { group: CompanyDisplayGroup; safeSelectedIndex: number; showPrefix?: boolean; onCompanySelect: (index: number) => void; onCompanyChange: (index: number, pending: boolean) => void; onCompanyPrefixChange?: (index: number, value: string) => void }) {
  if (!group.rows.length) return null;
  return <>{<tr className={`company-section-row ${group.className}`}><td colSpan={showPrefix ? 5 : 4}>{group.title} ({group.rows.length})</td></tr>}{group.rows.map(({ company, index }) => {
    const pending = company.pending_process ?? company.process ?? true;
    const selectedCount = company.selected_product_names.length || company.all_products.length;
    return <tr key={`${company.mst}-${index}`} className={`big-select-row ${group.className === 'duplicate-section' ? 'duplicate-company-row' : ''} ${index === safeSelectedIndex ? 'selected-row' : ''}`} onClick={() => onCompanySelect(index)}><td onClick={(event) => event.stopPropagation()}><input type="checkbox" checked={pending} onChange={(event) => onCompanyChange(index, event.currentTarget.checked)} /></td><td>{company.company}</td><td>{company.mst}</td>{showPrefix && <td onClick={(event) => event.stopPropagation()}><input className="company-prefix-input" value={company.value || ''} onChange={(event) => onCompanyPrefixChange?.(index, event.currentTarget.value)} /></td>}<td>{selectedCount} / {company.all_products.length}</td></tr>;
  })}</>;
}

function InventoryPairEditor({ pairs, useDefault, defaultPairId, rules, busy, onAddPair, onPairChange, onRemovePair, onDefaultsChange, onAddRule, onRuleChange, onRemoveRule }: { pairs: InventoryPair[]; useDefault: boolean; defaultPairId: string; rules: InventoryRule[]; busy: boolean; onAddPair: () => void; onPairChange: (index: number, field: 'ma_kho' | 'tk_vat_tu', value: string) => void; onRemovePair: (index: number) => void; onDefaultsChange: (update: Partial<Pick<WorkflowState, 'useDefaultInventoryPair' | 'defaultInventoryPairId'>>) => void; onAddRule: () => void; onRuleChange: (index: number, update: Partial<InventoryRule>) => void; onRemoveRule: (index: number) => void }) {
  const [collapsed, setCollapsed] = useState(false);
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
          <div className="inventory-table-scroll"><table className="inventory-table inventory-rule-table"><thead><tr><th>Bật</th><th>Cột nguồn</th><th>So sánh</th><th>Giá trị</th><th>Cặp gán</th><th></th></tr></thead><tbody>{rules.length ? rules.map((rule, index) => <tr key={`${rule.source_col}-${index}`}><td><input type="checkbox" checked={rule.enabled !== false} onChange={(event) => onRuleChange(index, { enabled: event.currentTarget.checked })} /></td><td><input value={rule.source_col} onChange={(event) => onRuleChange(index, { source_col: event.currentTarget.value.toUpperCase() })} /></td><td><select value={rule.operator || 'contains'} onChange={(event) => onRuleChange(index, { operator: event.currentTarget.value })}><option value="contains">Chứa</option><option value="equals">Bằng</option></select></td><td><input value={rule.value} onChange={(event) => onRuleChange(index, { value: event.currentTarget.value })} /></td><td><select value={rule.pair_id} onChange={(event) => onRuleChange(index, { pair_id: event.currentTarget.value })}><option value="">Chọn cặp</option>{pairs.map((pair) => <option key={pair.id} value={pair.id}>{pairLabel(pair)}</option>)}</select></td><td><button type="button" className="btn-secondary compact-table-button" disabled={busy} onClick={() => onRemoveRule(index)}>Xóa</button></td></tr>) : <tr><td colSpan={6} className="muted">Không có quy tắc. Nếu không bật mặc định, dòng không khớp sẽ để trống Mã kho/TK vật tư.</td></tr>}</tbody></table></div>
        </div>
      </div>}
    </section>
  );
}

function pairLabel(pair: InventoryPair) {
  return `${pair.ma_kho || 'Mã kho?'} / ${pair.tk_vat_tu || 'TK?'}`;
}

function computeCustomPrefix(company: CompanyRow, option: CustomPrefixOption): string {
  const provinceNames = new Set(['HÀ NỘI', 'HỒ CHÍ MINH', 'ĐÀ NẴNG', 'HẢI PHÒNG', 'CẦN THƠ', 'VIỆT NAM', 'Viet Nam', 'VIET NAM']);
  const words = company.company.split(/\s+/).filter((word) => word.length > 0);
  const filtered = words.filter((word) => !provinceNames.has(word));
  const significant = filtered.length ? filtered : words;
  const initials = significant.slice(-Math.max(1, option.chars)).map((word) => word[0]?.toUpperCase() || '').join('');
  const mstDigits = company.mst.slice(-Math.max(1, option.mstDigits ?? option.chars));
  if (option.formula === 'initials') return initials;
  if (option.formula === 'mst') return mstDigits;
  return initials + mstDigits;
}

function computePresetPrefix(company: CompanyRow, strategy: 'last_2_words' | 'last_3_mst' | '2_words_mst', mstDigits = 3): string {
  const digits = Math.max(1, Math.min(10, mstDigits));
  const mstSuffix = (company.mst || '').slice(-digits);
  const wordsPrefix = company.prefix_strategies?.last_2_words || computeCustomPrefix(company, { name: '2 words', formula: 'initials', chars: 2 });
  if (strategy === 'last_3_mst') return mstSuffix;
  if (strategy === '2_words_mst') return `${wordsPrefix}${mstSuffix}`;
  return wordsPrefix;
}

function normalizedPrefixStrategy(strategy: string): 'last_2_words' | 'last_3_mst' | '2_words_mst' {
  return strategy === 'last_3_mst' || strategy === '2_words_mst' ? strategy : 'last_2_words';
}

function companyDisplayGroups(companies: CompanyRow[], duplicatePrefixSet = duplicatePrefixSetForRows(companies)): CompanyDisplayGroup[] {
  const rows = companies.map((company, index) => ({ company, index }));
  const isActive = (company: CompanyRow) => company.process ?? true;
  const hasDuplicatePrefix = (company: CompanyRow) => isActive(company) && duplicatePrefixSet.has(committedCompanyPrefix(company));
  return [
    { title: 'Prefix trùng nhau - cần kiểm tra', className: 'duplicate-section', rows: rows.filter(({ company }) => hasDuplicatePrefix(company)) },
    { title: 'Các công ty đang xử lý', className: 'active-section', rows: rows.filter(({ company }) => isActive(company) && !hasDuplicatePrefix(company)) },
    { title: 'Các công ty đã bỏ qua', className: 'skipped-section', rows: rows.filter(({ company }) => !isActive(company)) },
  ];
}

function firstDisplayedCompanyIndex(companies: CompanyRow[]): number {
  for (const group of companyDisplayGroups(companies)) {
    if (group.rows.length) return group.rows[0].index;
  }
  return -1;
}

function normalizedCompanyPrefix(company: CompanyRow): string {
  return (company.value || '').trim().toUpperCase();
}

function committedCompanyPrefix(company: CompanyRow): string {
  return (company.committed_prefix ?? company.value ?? '').trim().toUpperCase();
}

function hasCompanyDraftChanges(company: CompanyRow): boolean {
  const pendingProcess = company.pending_process ?? company.process ?? true;
  const appliedProcess = company.process ?? true;
  return pendingProcess !== appliedProcess || normalizedCompanyPrefix(company) !== committedCompanyPrefix(company);
}

function duplicatePrefixSetForRows(companies: CompanyRow[]): Set<string> {
  const counts = new Map<string, number>();
  companies.forEach((company) => {
    if (!(company.process ?? true)) return;
    const prefix = committedCompanyPrefix(company);
    if (!prefix) return;
    counts.set(prefix, (counts.get(prefix) || 0) + 1);
  });
  return new Set(Array.from(counts.entries()).filter(([, count]) => count > 1).map(([prefix]) => prefix));
}

function sortAppliedCompanyRows(companies: CompanyRow[]): CompanyRow[] {
  const duplicatePrefixSet = duplicatePrefixSetForRows(companies);
  return companies
    .map((company, index) => ({ company, index }))
    .sort((left, right) => {
      const leftProcess = left.company.process ?? true;
      const rightProcess = right.company.process ?? true;
      const leftGroup = leftProcess ? (duplicatePrefixSet.has(committedCompanyPrefix(left.company)) ? 0 : 1) : 2;
      const rightGroup = rightProcess ? (duplicatePrefixSet.has(committedCompanyPrefix(right.company)) ? 0 : 1) : 2;
      return leftGroup - rightGroup || left.index - right.index;
    })
    .map(({ company }) => company);
}

function productDisplayCode(company: CompanyRow, productName: string, previewCodes: Record<string, string>, overrides: Record<string, string>, includePrefix: boolean = false) {
  const override = overrides[productKey(company.mst, productName)];
  if (override) return override;
  const preview = previewCodes[productName] || '';
  const appliedPrefix = committedCompanyPrefix(company);
  if (includePrefix && appliedPrefix) {
    return `${appliedPrefix}.${preview}`;
  }
  return preview;
}


function ProcessStage({ title, detail, buttonLabel = 'Xuất file', disabled, onProcess }: { title: string; detail: string; buttonLabel?: string; disabled: boolean; onProcess: () => void }) {
  return <div className="placeholder-stage"><h3>{title}</h3><p>{detail}</p><button type="button" disabled={disabled} onClick={onProcess}>{buttonLabel}</button></div>;
}

function PlaceholderStage({ title, detail }: { title: string; detail: string }) {
  return <div className="placeholder-stage"><h3>{title}</h3><p>{detail}</p></div>;
}

function LoadingStage({ title, detail }: { title: string; detail: string }) {
  return <div className="loading-stage"><div className="loading-spinner" aria-hidden="true" /><h3>{title}</h3><p>{detail}</p></div>;
}

function ProfilePlaceholder({ profile, stage }: { profile: string; stage: StageDefinition }) {
  const detail = stage.phase === 'price'
    ? 'Cao Thành có stage lọc đơn giá trong PySide. React chỉ hiển thị shell đúng luồng, chưa gọi backend cho stage này.'
    : 'Shell này phản ánh luồng PySide hiện tại. API React chưa migrate profile này, nên nội dung là placeholder an toàn và không giả lập xử lý.';
  return <PlaceholderStage title={`${profile}: ${stage.label}`} detail={detail} />;
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
      <tbody>{groups.map((group) => <ReviewGroupRows key={group.title} group={group} collapsed={Boolean(collapsedGroups[group.title])} onToggleGroup={onToggleGroup} onRowChange={onRowChange} />)}</tbody>
    </table>
  );
}

function ReviewGroupRows({ group, collapsed, onToggleGroup, onRowChange }: { group: ReviewDisplayGroup; collapsed: boolean; onToggleGroup: (groupTitle: string) => void; onRowChange?: (index: number, update: Partial<ReviewRow>) => void }) {
  if (!group.rows.length) return null;
  return <>{<tr className={`review-section-row ${group.className}`}><td colSpan={15}><button type="button" className="review-section-toggle" aria-expanded={!collapsed} onClick={() => onToggleGroup(group.title)}><span className="review-section-caret" aria-hidden="true">{collapsed ? '+' : '-'}</span><span className="review-section-label">{group.title}</span><span className="review-section-count">{group.rows.length}</span></button></td></tr>}{!collapsed && group.rows.map(({ row, index }) => (
        <tr key={`${row.product}-${row.similar_product}-${index}`}>
          <td className="review-use-cell"><input className="table-checkbox" type="checkbox" checked={row.confirmed !== false} disabled={!onRowChange} onChange={(event) => onRowChange?.(index, { confirmed: event.currentTarget.checked })} /></td>
          <td className="review-choice-cell"><select className="code-choice" value={row.code_choice || 'current'} disabled={!onRowChange} onChange={(event) => onRowChange?.(index, { code_choice: event.currentTarget.value })}><option value="current">Mã VT 1</option><option value="similar">Mã VT 2</option>{row.review_type === 'same_code_split' && <option value="split">Tách mã</option>}</select></td>
          <td className="review-product-cell">{row.product}</td>
          <td className="code-cell review-code-cell"><CodeValue label="Mã 1" value={row.code} /></td>
          <td>{row.invoice_no}</td>
          <td>{row.invoice_date}</td>
          <td>{row.unit}</td>
          <td className="review-product-cell">{row.similar_product}</td>
          <td className="code-cell review-code-cell"><CodeValue label="Mã 2" value={row.similar_code} /></td>
          <td>{row.similar_invoice_no}</td>
          <td>{row.similar_invoice_date}</td>
          <td>{row.similar_unit}</td>
          <td className="review-code-diff-cell"><CodeDiff current={row.code} target={row.similar_code} /></td>
          <td className="review-company-cell">{row.company || row.similar_company}</td>
          <td>{formatSimilarity(row.similarity)}</td>
        </tr>
      ))}</>;
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
  
  if (reviewScope === 'company') {
    // Group by company first, then by dimension/other within each company
    const companyGroups = new Map<string, typeof indexedRows>();
    indexedRows.forEach(({ row, index }) => {
      const companyKey = row.company || 'Không xác định';
      if (!companyGroups.has(companyKey)) {
        companyGroups.set(companyKey, []);
      }
      companyGroups.get(companyKey)!.push({ row, index });
    });
    
    const result: ReviewDisplayGroup[] = [];
    companyGroups.forEach((companyRows, companyName) => {
      const dimensionRows = companyRows.filter(({ row }) => row.review_group === 'dimension_diff' || row.dimension_only === true);
      const otherRows = companyRows.filter(({ row }) => row.review_group !== 'dimension_diff' && row.dimension_only !== true);
      
      if (dimensionRows.length) {
        result.push({ title: `${companyName} - Gần giống kích thước`, className: 'dimension-section', rows: dimensionRows });
      }
      if (otherRows.length) {
        result.push({ title: `${companyName} - Khác`, className: 'other-section', rows: otherRows });
      }
    });
    
    return result.length ? result : [
      { title: 'Gần giống kích thước', className: 'dimension-section', rows: [] },
      { title: 'Khác', className: 'other-section', rows: [] }
    ];
  }
  
  return [
    { title: 'Gần giống kích thước', className: 'dimension-section', rows: indexedRows.filter(({ row }) => row.review_group === 'dimension_diff' || row.dimension_only === true) },
    { title: 'Khác', className: 'other-section', rows: indexedRows.filter(({ row }) => row.review_group !== 'dimension_diff' && row.dimension_only !== true) },
  ];
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
  return rows.map((row) => ({ ...row, confirmed: row.confirmed !== false, code_choice: row.code_choice || 'current' }));
}

async function loadProductPreviewCodes(companies: CompanyRow[], wordRules: Record<string, string>, repeatedPhraseRemovals: string[], phase: 'purchase' | 'sales' = 'purchase') {
  const products = Array.from(new Set(companies.flatMap((company) => company.all_products.map((product) => product.name)).filter(Boolean)));
  if (!products.length) return {};
  const result = await previewVietmaxProductCodes(products, wordRules, repeatedPhraseRemovals, phase);
  return result.codes;
}

function buildPurchaseProcessPayload(workflow: WorkflowState) {
  const companies = workflow.companyRows;
  const activeCompanies = companies.filter((company) => company.process !== false);
  return {
    profile: 'vietmax',
    vietmax_phase: 'purchase',
    company_col: 'F',
    mst_col: 'G',
    product_col: 'M',
    qty_col: 'O',
    output_col: 'L',
    price_col: 'P',
    purchase_price_col: 'P',
    invoice_status_col: 'AJ',
    include_company_prefix: workflow.includeCompanyPrefix,
    comparison_scope: workflow.comparisonScope,
    word_rules: workflow.wordRules,
    repeated_phrase_removals: workflow.repeatedPhraseRemovals.filter((phrase) => phrase.trim()),
    manual_code_overrides: { ...workflow.productCodeOverrides, ...reviewManualCodeOverrides(workflow.purchaseReviewRows) },
    vietmax_mua_vao_internal_merges: confirmedPurchaseReviewMerges(workflow.purchaseReviewRows, workflow.comparisonScope),
    inventory_pairs: workflow.inventoryPairs.filter((pair) => pair.ma_kho.trim() || pair.tk_vat_tu.trim()),
    use_default_inventory_pair: workflow.useDefaultInventoryPair,
    default_inventory_pair_id: workflow.defaultInventoryPairId,
    inventory_pair_rules: workflow.inventoryPairRules.filter((rule) => rule.source_col.trim() && rule.pair_id.trim()),
    prefixes: companyPrefixes(companies),
    all_mst: companies.map((company) => company.mst),
    process_mst: activeCompanies.map((company) => company.mst),
    mst_safe_id: companies.map((company, index) => `${company.mst}|||${index}`),
    ...companyPrefixFields(companies),
    ...Object.fromEntries(companies.flatMap((company, index) => (company.process === false ? [] : [[`selected_products_${index}`, selectedProductNames(company)]]))),
  };
}

function buildSalesProcessPayload(workflow: WorkflowState) {
  const companies = workflow.salesCompanyRows;
  const activeCompanies = companies.filter((company) => company.process !== false);
  return {
    profile: 'vietmax',
    vietmax_phase: 'sales',
    company_col: 'I',
    mst_col: 'J',
    product_col: 'M',
    qty_col: 'O',
    output_col: 'L',
    price_col: 'P',
    purchase_price_col: 'P',
    invoice_status_col: 'AJ',
    include_company_prefix: workflow.includeCompanyPrefix,
    comparison_scope: workflow.comparisonScope,
    word_rules: workflow.wordRules,
    repeated_phrase_removals: workflow.repeatedPhraseRemovals.filter((phrase) => phrase.trim()),
    manual_code_overrides: workflow.salesProductCodeOverrides,
    inventory_pairs: workflow.inventoryPairs.filter((pair) => pair.ma_kho.trim() || pair.tk_vat_tu.trim()),
    use_default_inventory_pair: workflow.useDefaultInventoryPair,
    default_inventory_pair_id: workflow.defaultInventoryPairId,
    inventory_pair_rules: workflow.inventoryPairRules.filter((rule) => rule.source_col.trim() && rule.pair_id.trim()),
    vietmax_processed_purchase_saved_name: workflow.processedPurchaseSavedName,
    vietmax_ban_ra_purchase_matches: workflow.matches.filter((match) => match.confirmed !== false),
    vietmax_ban_ra_purchase_match_rules: buildSalesMatchRules(workflow),
    vietmax_ban_ra_sales_internal_merges: confirmedSalesReviewMerges(workflow.salesReviewRows, workflow.comparisonScope),
    prefixes: companyPrefixes(companies),
    all_mst: companies.map((company) => company.mst),
    process_mst: activeCompanies.map((company) => company.mst),
    mst_safe_id: companies.map((company, index) => `${company.mst}|||${index}`),
    ...companyPrefixFields(companies),
    ...Object.fromEntries(companies.flatMap((company, index) => (company.process === false ? [] : [[`selected_products_${index}`, selectedProductNames(company)]]))),
  };
}

function buildPurchaseReviewProducts(workflow: WorkflowState): ReviewProduct[] {
  return buildReviewProducts(workflow.companyRows, 'purchase');
}

function buildSalesReviewProducts(workflow: WorkflowState): ReviewProduct[] {
  return buildReviewProducts(workflow.salesCompanyRows, 'sales');
}

function buildReviewProducts(companies: CompanyRow[], phase: 'purchase' | 'sales'): ReviewProduct[] {
  return companies.flatMap((company, companyIndex) => {
    if (company.process === false) return [];
    const selected = new Set(selectedProductNames(company));
    return company.all_products.flatMap((product, productIndex) => {
      if (!selected.has(product.name)) return [];
      const firstPriceRow = product.priceRows?.[0];
      return [{
        purchase_product: product.name,
        purchase_code: '',
        purchase_unit: firstPriceRow?.unit ?? '',
        invoice_no: firstPriceRow?.invoiceNo ?? '',
        invoice_date: firstPriceRow?.invoiceDate ?? '',
        purchase_company: company.company,
        purchase_mst: company.mst,
        purchase_company_key: company.mst,
        product_key: productKey(company.mst, product.name),
        company_index: companyIndex,
        product_index: productIndex,
        phase,
      }];
    });
  });
}

function confirmedSalesReviewMerges(rows: ReviewRow[], comparisonScope: string) {
  return rows.flatMap((row) => {
    if (row.confirmed === false || !row.product || !row.similar_product) return [];
    const merge = { ...row, comparison_scope: comparisonScope };
    if (row.code_choice === 'current') {
      return [{ ...merge, product: row.similar_product, similar_product: row.product }];
    }
    return [{ ...merge, product: row.product, similar_product: row.similar_product }];
  });
}

function confirmedPurchaseReviewMerges(rows: ReviewRow[], comparisonScope: string) {
  return rows.flatMap((row) => {
    if (row.confirmed === false || row.code_choice === 'split' || !row.product || !row.similar_product) return [];
    const merge = { ...row, comparison_scope: comparisonScope };
    if (row.code_choice === 'similar') {
      return [{ ...merge, product: row.product, similar_product: row.similar_product, similar_code: row.similar_code }];
    }
    return [{ ...merge, product: row.similar_product, similar_product: row.product, similar_code: row.code }];
  });
}

function buildConfigPayloads(workflow: WorkflowState) {
  const payloads = [];
  if (workflow.companyRows.length) payloads.push(buildConfigPayload(workflow));
  if (workflow.salesCompanyRows.length || workflow.salesFile || workflow.matches.length || workflow.salesMatchRules.length) payloads.push(buildSalesConfigPayload(workflow));
  return payloads.length ? payloads : [buildConfigPayload(workflow)];
}

function buildConfigPayload(workflow: WorkflowState) {
  const companies = workflow.companyRows;
  return {
    ...buildPurchaseProcessPayload(workflow),
    columns: {
      company_col: 'F',
      mst_col: 'G',
      address_col: 'H',
      product_col: 'M',
      qty_col: 'O',
      price_col: 'P',
      purchase_price_col: 'P',
      output_col: 'L',
      invoice_status_col: 'AJ',
    },
    prefixes: companyPrefixes(companies),
    removed_companies: Object.fromEntries(companies.filter((company) => company.process === false).map((company) => [company.mst, true])),
    skipped_products_map: Object.fromEntries(companies.map((company) => {
      const selected = new Set(selectedProductNames(company));
      const skipped = company.all_products.map((product) => product.name).filter((name) => !selected.has(name));
      return [company.mst, skipped];
    }).filter(([, skipped]) => Array.isArray(skipped) && skipped.length)),
    manual_code_overrides: { ...workflow.productCodeOverrides, ...reviewManualCodeOverrides(workflow.purchaseReviewRows) },
  };
}

function buildSalesConfigPayload(workflow: WorkflowState) {
  const companies = workflow.salesCompanyRows;
  return {
    ...buildSalesProcessPayload(workflow),
    columns: {
      company_col: 'I',
      mst_col: 'J',
      address_col: 'K',
      product_col: 'M',
      qty_col: 'O',
      price_col: 'P',
      purchase_price_col: 'P',
      output_col: 'L',
      invoice_status_col: 'AJ',
    },
    removed_companies: Object.fromEntries(companies.filter((company) => company.process === false).map((company) => [company.mst, true])),
    skipped_products_map: Object.fromEntries(companies.map((company) => {
      const selected = new Set(selectedProductNames(company));
      const skipped = company.all_products.map((product) => product.name).filter((name) => !selected.has(name));
      return [company.mst, skipped];
    }).filter(([, skipped]) => Array.isArray(skipped) && skipped.length)),
    manual_code_overrides: workflow.salesProductCodeOverrides,
    vietmax_ban_ra_purchase_match_rules: buildSalesMatchRules(workflow),
  };
}

function companyPrefixes(companies: CompanyRow[]) {
  return Object.fromEntries(companies.filter((company) => company.value.trim()).map((company) => [company.mst, company.value.trim()]));
}

function companyPrefixFields(companies: CompanyRow[]) {
  return Object.fromEntries(companies.map((company, index) => [`prefix_${index}`, company.value.trim()]));
}

function selectedProductNames(company: CompanyRow) {
  return company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name);
}


function reviewManualCodeOverrides(rows: ReviewRow[]) {
  const overrides: Record<string, string> = {};
  for (const row of rows) {
    if (row.confirmed === false) continue;
    const currentKey = row.product_key || productKey(row.mst, row.product);
    const similarKey = row.similar_product_key || productKey(row.similar_mst, row.similar_product);
    if (row.code_choice === 'split') {
      if (currentKey && row.split_code) overrides[currentKey] = row.split_code;
      if (similarKey && row.similar_split_code) overrides[similarKey] = row.similar_split_code;
    } else if (row.code_choice === 'similar') {
      if (currentKey && row.similar_code) overrides[currentKey] = row.similar_code;
    } else if (similarKey && row.code) {
      overrides[similarKey] = row.code;
    }
  }
  return overrides;
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
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function sleep(milliseconds: number) {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

function formatSimilarity(value: number | string | undefined) {
  if (typeof value === 'string') return value;
  return typeof value === 'number' ? `${Math.round(value * 100)}%` : '';
}
