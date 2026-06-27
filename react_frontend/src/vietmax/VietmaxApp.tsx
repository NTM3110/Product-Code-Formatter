import { Fragment, useEffect, useMemo, useState } from 'react';
import { activateLicense, analyzeGenericWorkbook, analyzeVietmaxCompanies, createGenericReview, createPurchaseReview, createSalesMatches, createVietmaxFastImportPackage, downloadCachedFile, downloadInventoryAllocationReport, exportMatches, exportPriceReportWorkbook, getAppConfig, getInventoryAllocationJob, getLicenseStatus, getOperationProgress, importVietmaxConfig, inspectProcessedVietmaxFile, previewGenericProductCodes, previewVietmaxProductCodes, processGenericWorkbook, processVietmaxPurchase, saveVietmaxConfig, reloadLicense, startInventoryAllocation, uploadExcel, validateFastImportProcessedFile } from '../api';
import type { CompanyRow, InventoryAllocationConfig, InventoryAllocationJob, InventoryAllocationResult, InventoryPair, InventoryRule, LicenseStatus, MatchRow, OperationProgress, ProcessedFileStats, ReviewProduct, ReviewRow, UploadSummary } from '../types';
import { InventoryAllocationExportStage, InventoryAllocationReportStage, InventoryAllocationStage } from './InventoryAllocationStage';
import { StageNavigation } from './StageNavigation';
import { isGenericProfileKey, isStageId, profiles, stagesForProfile, type PrefixPresetStrategy, type ProfileKey, type StageDefinition, type StageId, type StagePhase } from './workflowStages';

type PrefixStrategyValues = Record<PrefixPresetStrategy, Record<string, string>>;
type InventoryConfigScope = 'purchase' | 'sales' | 'generic';

type GenericColumns = {
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

type WorkflowState = {
  stage: StageId;
  purchaseFile: UploadSummary | null;
  genericColumns: GenericColumns;
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
  salesCompanyRows: CompanyRow[];
  selectedSalesCompanyIndex: number;
  productPreviewCodes: Record<string, string>;
  salesProductPreviewCodes: Record<string, string>;
  productCodeOverrides: Record<string, string>;
  salesProductCodeOverrides: Record<string, string>;
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
  purchasePrefixStrategy: PrefixPresetStrategy;
  salesPrefixStrategy: PrefixPresetStrategy;
  prefixMstDigits: number;
  purchasePrefixStrategyValues: PrefixStrategyValues;
  salesPrefixStrategyValues: PrefixStrategyValues;
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

const defaultInvoiceStatusSkipValues = [
  'Hóa đơn đã bị điều chỉnh',
  'Hóa đơn bị thay thế',
  'Hóa đơn đã bị thay thế',
  'Hóa đơn đã bị hủy',
];

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

function initialWorkflowState(): WorkflowState {
  return {
    stage: 1,
    purchaseFile: null,
    genericColumns: defaultGenericColumns(),
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
    salesCompanyRows: [],
    selectedSalesCompanyIndex: -1,
    productPreviewCodes: {},
    salesProductPreviewCodes: {},
    productCodeOverrides: {},
    salesProductCodeOverrides: {},
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
    purchasePrefixStrategy: 'last_2_words',
    salesPrefixStrategy: 'last_2_words',
    prefixMstDigits: 3,
    purchasePrefixStrategyValues: emptyPrefixStrategyValues(),
    salesPrefixStrategyValues: emptyPrefixStrategyValues(),
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
  const [license, setLicense] = useState<LicenseStatus | null>(null);
  const [licenseForm, setLicenseForm] = useState(initialLicenseForm);
  const [status, setStatus] = useState('Chọn profile và bắt đầu theo từng stage. Dữ liệu được giữ khi chuyển stage, chỉ xóa khi bấm Làm lại.');
  const [busy, setBusy] = useState(false);
  const [loadingProgress, setLoadingProgress] = useState<OperationProgress | null>(null);

  const workflow = workflows[profile];
  const { stage, purchaseFile, genericColumns, processedPurchaseSavedName, processedPurchaseStats, salesFile, processedSalesSavedName, processedSalesStats, openingStockFile, inventoryAllocationConfig, inventoryAllocationJob, inventoryAllocationResult, comparisonScope, companyRows, selectedCompanyIndex, salesCompanyRows, selectedSalesCompanyIndex, productPreviewCodes, salesProductPreviewCodes, productCodeOverrides, salesProductCodeOverrides, purchaseWordRules, salesWordRules, firstWordRules, purchaseRepeatedPhraseRemovals, salesRepeatedPhraseRemovals, wordRules, repeatedPhraseRemovals, purchaseReviewRows, salesReviewRows, purchaseReviewRules, salesReviewRules, purchaseReviewGenerated, salesReviewGenerated, priceRangeRules, priceGroups, priceFilterAllPercent, priceAdjustAllPercent, matches, salesMatchGenerated, salesMatchRules, purchaseInventoryPairs, purchaseUseDefaultInventoryPair, purchaseDefaultInventoryPairId, purchaseInventoryPairRules, salesInventoryPairs, salesUseDefaultInventoryPair, salesDefaultInventoryPairId, salesInventoryPairRules, inventoryPairs, useDefaultInventoryPair, defaultInventoryPairId, inventoryPairRules, includeCompanyPrefix, purchasePrefixStrategy, salesPrefixStrategy, prefixMstDigits, purchasePrefixStrategyValues, salesPrefixStrategyValues, purchaseReviewScope, salesReviewScope } = workflow;
  const selectedProfile = profiles.find((item) => item.key === profile) ?? profiles[0];
  const licenseReady = Boolean(license?.activated && (profile !== 'vietmax' || license.vietmax_allowed));
  const isGenericProfile = isGenericProfileKey(profile);
  const usesNativeStageShell = profile === 'vietmax' || isGenericProfile;
  const visibleStages = useMemo(() => stagesForProfile(profile), [profile]);
  const currentStage = visibleStages.find((item) => item.id === stage) ?? visibleStages[0];
  const selectedMatches = useMemo(() => matches.filter((match) => match.confirmed !== false), [matches]);
  const showLicenseBar = stage === 1;
  const activeVietmaxSalesConfig = profile === 'vietmax' && stage >= 6;
  const activeInventoryConfigScope: InventoryConfigScope = profile === 'vietmax' ? (activeVietmaxSalesConfig ? 'sales' : 'purchase') : 'generic';
  const activeWordRules = profile === 'vietmax' ? (activeVietmaxSalesConfig ? salesWordRules : purchaseWordRules) : wordRules;
  const activeRepeatedPhraseRemovals = profile === 'vietmax' ? (activeVietmaxSalesConfig ? salesRepeatedPhraseRemovals : purchaseRepeatedPhraseRemovals) : repeatedPhraseRemovals;

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
    if (isGenericProfile) void loadGenericProfileConfig(profile);
  }, [profile]);

  useEffect(() => {
    if (isGenericProfile && stage === 3 && purchaseFile && !companyRows.length && !busy) {
      void loadGenericCompanies();
    }
  }, [isGenericProfile, stage, purchaseFile, companyRows.length, busy]);

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
    if (profile === 'vietmax' && stage === 4 && purchaseFile && companyRows.length && !purchaseReviewGenerated && !busy) {
      if (companyRows.some(hasCompanyDraftChanges)) {
        applyCompanyAndProductChoices(4);
        return;
      }
      void runPurchaseReview();
    }
  }, [profile, stage, purchaseFile, companyRows.length, purchaseReviewGenerated, busy]);

  useEffect(() => {
    if (profile === 'vietmax' && stage === 10 && salesFile && salesCompanyRows.length && !salesReviewGenerated && !busy) {
      if (salesCompanyRows.some(hasCompanyDraftChanges)) {
        applySalesCompanyAndProductChoices(10);
        return;
      }
      void runSalesReview();
    }
  }, [profile, stage, salesFile, salesCompanyRows.length, salesReviewGenerated, busy]);

  useEffect(() => {
    if (isGenericProfile && stage === 4 && purchaseFile && companyRows.length && !purchaseReviewGenerated && !busy) {
      if (companyRows.some(hasCompanyDraftChanges)) {
        applyCompanyAndProductChoices(4);
        return;
      }
      void runGenericReview();
    }
  }, [isGenericProfile, stage, purchaseFile, companyRows.length, purchaseReviewGenerated, busy]);

  useEffect(() => {
    if (profile === 'cao_thanh' && stage === 5 && companyRows.length && !priceGroups.length && !busy) {
      updateCaoThanhPriceGroups();
    }
  }, [profile, stage, companyRows.length, priceGroups.length, busy]);

  useEffect(() => {
    if (profile === 'vietmax' && stage === 5 && purchaseFile && !processedPurchaseSavedName && !busy) {
      void prepareProcessedPurchaseCache();
    }
  }, [profile, stage, purchaseFile, processedPurchaseSavedName, busy]);

  useEffect(() => {
    if (profile === 'vietmax' && stage === 11 && salesFile && !processedSalesSavedName && !busy) {
      void prepareProcessedSalesCache();
    }
  }, [profile, stage, salesFile, processedSalesSavedName, busy]);

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

  function resetWorkflow() {
    updateWorkflow(profile, initialWorkflowState());
    setStatus(profile === 'vietmax' ? 'Đã làm lại. Hãy tải file mua vào Vietmax từ stage 1.' : `Đã làm lại profile ${selectedProfile.label}.`);
  }

  function canEnterStage(target: StageId) {
    if (!visibleStages.some((item) => item.id === target)) return false;
    if (isGenericProfile) {
      if (!licenseReady) return target === 1;
      if (target === 1) return true;
      if (target === 2 || target === 3) return Boolean(purchaseFile);
      if (target === 4) return Boolean(purchaseFile && companyRows.length);
      if (profile === 'cao_thanh' && (target === 5 || target === 6)) return Boolean(purchaseFile && companyRows.length && purchaseReviewGenerated);
      if (target === 5) return Boolean(purchaseFile && companyRows.length && purchaseReviewGenerated);
      return false;
    }
    if (profile !== 'vietmax') return licenseReady || target === 1;
    if (!licenseReady) return target === 1;
    if (target <= 2) return true;
    if (target === 6 || target === 12) return true;
    if (target <= 5) return Boolean(purchaseFile);
    if (target <= 11) return Boolean(salesFile && (purchaseFile || processedPurchaseSavedName));
    if (target === 13 || target === 14) return Boolean(inventoryAllocationResult?.job_id || inventoryAllocationJob?.result?.job_id);
    return true;
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
    if (isGenericProfile) {
      if (stage === 2 && purchaseFile && !companyRows.length) {
        void loadGenericCompanies(next.id);
        return;
      }
      if (stage === 3 && companyRows.some(hasCompanyDraftChanges)) {
        applyCompanyAndProductChoices(next.id);
        return;
      }
      if (stage === 4 && purchaseFile && !purchaseReviewGenerated) {
        void runGenericReview();
        return;
      }
      if (profile === 'cao_thanh' && stage === 5 && !priceGroups.length) {
        updateCaoThanhPriceGroups();
      }
      goToStage(next.id);
      return;
    }
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
    if (stage === 6 && salesFile && !purchaseFile && !processedPurchaseSavedName) {
      setStatus('Cần tải file mua vào đã xử lý trước khi đi tiếp các stage bán ra.');
      return;
    }
    if (stage === 8 && salesFile && (processedPurchaseSavedName || purchaseFile) && !matches.length && !salesMatchGenerated) {
      void runSalesMatch();
      return;
    }
    if (stage === 10 && salesFile && !salesReviewGenerated) {
      void runSalesReview();
      return;
    }
    if (stage === 11 && salesFile && !processedSalesSavedName) {
      void prepareProcessedSalesCache(next.id);
      return;
    }
    if (stage === 12 && !inventoryAllocationResult?.job_id && !inventoryAllocationJob?.result?.job_id) {
      void runInventoryAllocation(next.id);
      return;
    }
    goToStage(next.id);
  }

  async function submitLicense() {
    setBusy(true);
    setStatus('Đang kích hoạt license...');
    try {
      const nextLicense = await activateLicense({
        license_key: licenseForm.license_key,
        server_url: licenseForm.server_url || undefined,
        account_id: licenseForm.account_id || undefined,
      });
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

  async function uploadProcessed(kind: 'purchase' | 'sales', file: File | undefined) {
    if (!file) return;
    const targetProfile = profile;
    setBusy(true);
    setStatus(`Đang tải file ${kind === 'purchase' ? 'mua vào' : 'bán ra'} đã xử lý...`);
    try {
      const summary = await uploadExcel(file);
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
      const summary = await uploadExcel(file);
      const stats = await inspectProcessedVietmaxFile(summary.saved_name, kind);
      const validation = await validateFastImportProcessedFile(summary.saved_name, kind);
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
      setStatus(`Đã tải FDI ${label} cho Xuất FAST. ${processedStatsSentence(stats)} Đã kiểm tra ${validation.valid_rows} dòng có đủ TK vật tư (${validation.tk_vat_tu_col}) và Mã kho (${validation.ma_kho_col}).`);
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

  async function loadGenericProfileConfig(targetProfile: ProfileKey) {
    if (!isGenericProfileKey(targetProfile)) return;
    try {
      const cfg = await getAppConfig();
      const profilesCfg = (cfg.profiles && typeof cfg.profiles === 'object' ? cfg.profiles : {}) as Record<string, any>;
      const profileCfg = profilesCfg[targetProfile] || {};
      const globalColumns = cfg.columns && typeof cfg.columns === 'object' ? cfg.columns as Record<string, unknown> : {};
      const savedColumns = profileCfg.columns && typeof profileCfg.columns === 'object' ? profileCfg.columns as Record<string, unknown> : {};
      updateWorkflow(targetProfile, {
        genericColumns: normalizeGenericColumns({ ...globalColumns, ...savedColumns }),
        wordRules: profileCfg.word_rules && typeof profileCfg.word_rules === 'object' ? profileCfg.word_rules : {},
        firstWordRules: profileCfg.first_word_rules && typeof profileCfg.first_word_rules === 'object' ? profileCfg.first_word_rules : {},
        repeatedPhraseRemovals: Array.isArray(profileCfg.repeated_phrase_removals) ? profileCfg.repeated_phrase_removals : [],
        inventoryPairs: Array.isArray(profileCfg.inventory_pairs) ? profileCfg.inventory_pairs : [],
        useDefaultInventoryPair: Boolean(profileCfg.use_default_inventory_pair),
        defaultInventoryPairId: String(profileCfg.default_inventory_pair_id || ''),
        inventoryPairRules: Array.isArray(profileCfg.inventory_pair_rules) ? profileCfg.inventory_pair_rules : [],
        includeCompanyPrefix: profileCfg.include_company_prefix !== false,
        purchasePrefixStrategy: normalizedPrefixStrategy(profileCfg.prefix_strategy || 'last_2_words'),
        prefixMstDigits: clampPrefixMstDigits(profileCfg.prefix_mst_digits ?? 3),
        purchasePrefixStrategyValues: normalizePrefixStrategyValues(profileCfg.prefix_strategy_values, emptyPrefixStrategyValues()),
        purchaseReviewRules: Array.isArray(profileCfg.product_review_merges) ? profileCfg.product_review_merges : [],
        priceRangeRules: profileCfg.price_range_rules && typeof profileCfg.price_range_rules === 'object' ? profileCfg.price_range_rules : {},
        priceAdjustAllPercent: Number(profileCfg.price_adjust_all_percent || 0),
      });
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    }
  }

  async function loadGenericCompanies(nextStage?: StageId) {
    if (!purchaseFile || !isGenericProfile) return;
    const targetProfile = profile;
    const targetFile = purchaseFile;
    setBusy(true);
    setStatus(`Đang tải danh sách công ty và hàng hóa ${selectedProfile.label}...`);
    try {
      const result = await analyzeGenericWorkbook({
        saved_name: targetFile.saved_name,
        original_name: targetFile.original_name,
        profile: targetProfile,
        ...genericColumns,
      });
      const savedWordRules = result.word_rules ?? wordRules;
      const savedFirstWordRules = result.first_word_rules ?? firstWordRules;
      const savedRepeatedPhrases = result.repeated_phrase_removals ?? repeatedPhraseRemovals;
      const savedInventoryPairs = result.inventory_pairs ?? inventoryPairs;
      const nextCompanies = result.companies.map((company) => ({
        ...company,
        process: company.process ?? true,
        pending_process: company.pending_process ?? company.process ?? true,
        committed_prefix: company.committed_prefix ?? company.value ?? '',
        selected_product_names: company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name),
      }));
      const loadedPrefixStrategy = normalizedPrefixStrategy(result.prefix_strategy || purchasePrefixStrategy);
      const loadedPrefixMstDigits = clampPrefixMstDigits(result.prefix_mst_digits ?? prefixMstDigits);
      const loadedPrefixValues = normalizePrefixStrategyValues(result.prefix_strategy_values, purchasePrefixStrategyValues);
      const nextPrefixValues = seedLoadedPrefixValues(loadedPrefixValues, loadedPrefixStrategy, nextCompanies, loadedPrefixMstDigits);
      const displayCompanies = applyPrefixStrategyRows(nextCompanies, loadedPrefixStrategy, loadedPrefixMstDigits, nextPrefixValues, true);
      const previewCodes = await loadGenericProductPreviewCodes(targetProfile, displayCompanies, savedWordRules, savedFirstWordRules, savedRepeatedPhrases);
      updateWorkflow(targetProfile, {
        ...purchaseOutputInvalidation(),
        genericColumns: normalizeGenericColumns(genericColumns),
        companyRows: displayCompanies,
        selectedCompanyIndex: firstDisplayedCompanyIndex(displayCompanies),
        productPreviewCodes: previewCodes,
        productCodeOverrides: result.manual_code_overrides ?? {},
        wordRules: savedWordRules,
        firstWordRules: savedFirstWordRules,
        repeatedPhraseRemovals: savedRepeatedPhrases,
        inventoryPairs: savedInventoryPairs,
        useDefaultInventoryPair: result.use_default_inventory_pair ?? useDefaultInventoryPair,
        defaultInventoryPairId: result.default_inventory_pair_id ?? defaultInventoryPairId,
        inventoryPairRules: result.inventory_pair_rules ?? inventoryPairRules,
        includeCompanyPrefix: result.include_company_prefix ?? includeCompanyPrefix,
        purchasePrefixStrategy: loadedPrefixStrategy,
        prefixMstDigits: loadedPrefixMstDigits,
        purchasePrefixStrategyValues: nextPrefixValues,
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
      productPreviewCodes: {},
      productCodeOverrides: {},
      purchaseReviewRows: [],
      purchaseReviewGenerated: false,
      priceGroups: [],
    });
  }

  async function downloadGenericProcessedFile() {
    if (!purchaseFile || !isGenericProfile) return;
    const dirtyCompanies = companyRows.filter(hasCompanyDraftChanges);
    if (dirtyCompanies.length) {
      setStatus('Đang có thay đổi lọc công ty/prefix chưa áp dụng. Bấm Áp dụng lựa chọn công ty và hàng hóa trước khi xuất file.');
      return;
    }
    setBusy(true);
    setStatus(`Đang tạo file ${selectedProfile.label}...`);
    try {
      const blob = await processGenericWorkbook({
        saved_name: purchaseFile.saved_name,
        original_name: purchaseFile.original_name,
        ...buildGenericProcessPayload(workflow, profile),
      });
      const saved = await saveBlob(blob, `${fileStem(purchaseFile.original_name)}_fdi.xls`);
      setStatus(saved ? `Đã xuất file kết quả ${selectedProfile.label}.` : 'Đã hủy lưu file kết quả.');
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
    if (!purchaseFile || !isGenericProfile) return;
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
      const result = await createGenericReview(targetPurchaseFile.saved_name, targetProfile, scope, wordRules, firstWordRules, repeatedPhraseRemovals, reviewProducts, progress.operationId);
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

  async function prepareProcessedPurchaseCache(nextStage?: StageId) {
    if (!purchaseFile) return;
    if (processedPurchaseSavedName) {
      if (nextStage) goToStage(nextStage);
      return;
    }
    const targetProfile = profile;
    const targetPurchaseFile = purchaseFile;
    const progress = beginProgress('Đang tạo cache file mua vào đã xử lý');
    setBusy(true);
    setStatus('Đang tạo cache file mua vào đã xử lý để dùng cho khớp mua/bán...');
    try {
      const result = await processVietmaxPurchase(targetPurchaseFile.saved_name, targetPurchaseFile.original_name, buildPurchaseProcessPayload(workflow), { cacheOnly: true, operationId: progress.operationId });
      if (!result.processedSavedName) throw new Error('Không tạo được cache file mua vào đã xử lý.');
      const stats = await applyProcessedPurchaseCache(targetProfile, result.processedSavedName);
      setStatus(`Đã tạo cache file mua vào đã xử lý. ${processedStatsSentence(stats)}`);
      if (nextStage) updateWorkflow(targetProfile, { stage: nextStage });
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function prepareProcessedSalesCache(nextStage?: StageId) {
    if (!salesFile) return;
    if (processedSalesSavedName) {
      if (nextStage) goToStage(nextStage);
      return;
    }
    const targetProfile = profile;
    const targetSalesFile = salesFile;
    const progress = beginProgress('Đang tạo cache file bán ra đã xử lý');
    setBusy(true);
    setStatus('Đang tạo cache file bán ra đã xử lý để xuất nhanh và dùng cho phân bổ tồn kho...');
    try {
      const result = await processVietmaxPurchase(targetSalesFile.saved_name, targetSalesFile.original_name, buildSalesProcessPayload(workflow), { cacheOnly: true, operationId: progress.operationId });
      if (!result.processedSavedName) throw new Error('Không tạo được cache file bán ra đã xử lý.');
      const stats = await inspectProcessedVietmaxFile(result.processedSavedName, 'sales');
      updateWorkflow(targetProfile, {
        processedSalesSavedName: result.processedSavedName,
        processedSalesStats: stats,
        inventoryAllocationJob: null,
        inventoryAllocationResult: null,
        ...(nextStage ? { stage: nextStage } : {}),
      });
      setStatus(`Đã tạo cache file bán ra đã xử lý. ${processedStatsSentence(stats)}`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function runSalesMatch() {
    if (!salesFile) return;
    const targetProfile = profile;
    const targetSalesFile = salesFile;
    const targetPurchaseFile = purchaseFile;
    let targetProcessedPurchase = processedPurchaseSavedName;
    const progress = beginProgress(targetProcessedPurchase ? 'Đang chuẩn bị khớp mua vào / bán ra' : 'Đang chuẩn bị cache file mua vào đã xử lý');
    setBusy(true);
    setStatus(targetProcessedPurchase ? 'Đang khớp bán ra với file mua vào đã xử lý KVT/152...' : 'Đang chuẩn bị cache file mua vào đã xử lý trước khi khớp mua/bán...');
    try {
      if (!targetProcessedPurchase) {
        if (!targetPurchaseFile) throw new Error('Chưa có file mua vào để tạo cache xử lý.');
        const purchaseResult = await processVietmaxPurchase(targetPurchaseFile.saved_name, targetPurchaseFile.original_name, buildPurchaseProcessPayload(workflow), { cacheOnly: true, operationId: progress.operationId });
        targetProcessedPurchase = purchaseResult.processedSavedName;
        if (!targetProcessedPurchase) throw new Error('Không tạo được cache file mua vào đã xử lý.');
        await applyProcessedPurchaseCache(targetProfile, targetProcessedPurchase);
        setStatus('Đã tạo cache mua vào. Đang khớp bán ra với file mua vào đã xử lý KVT/152...');
      }
      const result = await createSalesMatches(targetSalesFile.saved_name, targetProcessedPurchase, comparisonScope, progress.operationId);
      const savedRules = result.match_rules?.length ? result.match_rules : salesMatchRules;
      const nextMatches = applySalesMatchRules(result.matches, savedRules, comparisonScope);
      const mismatchCount = nextMatches.filter(hasUnitMismatch).length;
      updateWorkflow(targetProfile, { ...salesOutputInvalidation(), matches: nextMatches, salesMatchGenerated: true, salesMatchRules: savedRules, salesCompanyRows: [], selectedSalesCompanyIndex: -1, salesProductPreviewCodes: {}, salesProductCodeOverrides: {}, salesReviewRows: [], salesReviewGenerated: false, stage: 8 });
      setStatus(`Đã gợi ý ${result.matches.length} dòng khớp. ${result.exact_matches.length} dòng lấy chính xác từ KVT/152. ${mismatchCount} dòng khác ĐVT.`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
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
      const result = await analyzeVietmaxCompanies(purchaseFile.saved_name, 'purchase');
      const savedWordRules = result.word_rules ?? purchaseWordRules;
      const savedRepeatedPhrases = result.repeated_phrase_removals ?? purchaseRepeatedPhraseRemovals;
      const savedInventoryPairs = result.inventory_pairs ?? purchaseInventoryPairs;
      const nextCompanies = result.companies.map((company) => ({
        ...company,
        process: company.process ?? true,
        pending_process: company.pending_process ?? company.process ?? true,
        committed_prefix: company.committed_prefix ?? company.value ?? '',
        selected_product_names: company.selected_product_names.length ? company.selected_product_names : company.all_products.map((product) => product.name),
      }));
      const loadedPrefixStrategy = normalizedPrefixStrategy(result.prefix_strategy || purchasePrefixStrategy);
      const loadedPrefixMstDigits = clampPrefixMstDigits(result.prefix_mst_digits ?? prefixMstDigits);
      const loadedPrefixValues = normalizePrefixStrategyValues(result.prefix_strategy_values, purchasePrefixStrategyValues);
      const nextPrefixValues = seedLoadedPrefixValues(loadedPrefixValues, loadedPrefixStrategy, nextCompanies, loadedPrefixMstDigits);
      const displayCompanies = applyPrefixStrategyRows(nextCompanies, loadedPrefixStrategy, loadedPrefixMstDigits, nextPrefixValues, true);
      const previewCodes = await loadProductPreviewCodes(displayCompanies, savedWordRules, savedRepeatedPhrases);
      updateWorkflow(targetProfile, { ...purchaseOutputInvalidation(), companyRows: displayCompanies, selectedCompanyIndex: firstDisplayedCompanyIndex(displayCompanies), productPreviewCodes: previewCodes, productCodeOverrides: result.manual_code_overrides ?? {}, purchaseWordRules: savedWordRules, purchaseRepeatedPhraseRemovals: savedRepeatedPhrases, purchaseInventoryPairs: savedInventoryPairs, purchaseUseDefaultInventoryPair: result.use_default_inventory_pair ?? purchaseUseDefaultInventoryPair, purchaseDefaultInventoryPairId: result.default_inventory_pair_id ?? purchaseDefaultInventoryPairId, purchaseInventoryPairRules: result.inventory_pair_rules ?? purchaseInventoryPairRules, includeCompanyPrefix: result.include_company_prefix ?? includeCompanyPrefix, purchasePrefixStrategy: loadedPrefixStrategy, prefixMstDigits: loadedPrefixMstDigits, purchasePrefixStrategyValues: nextPrefixValues, purchaseReviewRules: result.vietmax_mua_vao_internal_merges ?? purchaseReviewRules, purchaseReviewRows: [], purchaseReviewGenerated: false });
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
      const savedWordRules = result.word_rules ?? salesWordRules;
      const savedRepeatedPhrases = result.repeated_phrase_removals ?? salesRepeatedPhraseRemovals;
      const savedInventoryPairs = result.inventory_pairs?.length ? result.inventory_pairs : salesInventoryPairs;
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
      const loadedPrefixStrategy = normalizedPrefixStrategy(result.prefix_strategy || salesPrefixStrategy);
      const loadedPrefixMstDigits = clampPrefixMstDigits(result.prefix_mst_digits ?? prefixMstDigits);
      const loadedPrefixValues = normalizePrefixStrategyValues(result.prefix_strategy_values, salesPrefixStrategyValues);
      const nextPrefixValues = seedLoadedPrefixValues(loadedPrefixValues, loadedPrefixStrategy, nextCompanies, loadedPrefixMstDigits);
      const displayCompanies = applyPrefixStrategyRows(nextCompanies, loadedPrefixStrategy, loadedPrefixMstDigits, nextPrefixValues, true);
      const previewCodes = await loadProductPreviewCodes(displayCompanies, savedWordRules, savedRepeatedPhrases, 'sales');
      updateWorkflow(targetProfile, { ...salesOutputInvalidation(), salesCompanyRows: displayCompanies, selectedSalesCompanyIndex: firstDisplayedCompanyIndex(displayCompanies), salesProductPreviewCodes: previewCodes, salesProductCodeOverrides: result.manual_code_overrides ?? {}, salesWordRules: savedWordRules, salesRepeatedPhraseRemovals: savedRepeatedPhrases, salesMatchRules: result.sales_match_rules ?? salesMatchRules, salesInventoryPairs: savedInventoryPairs, salesUseDefaultInventoryPair: result.inventory_pairs?.length ? Boolean(result.use_default_inventory_pair) : salesUseDefaultInventoryPair, salesDefaultInventoryPairId: result.inventory_pairs?.length ? (result.default_inventory_pair_id ?? '') : salesDefaultInventoryPairId, salesInventoryPairRules: result.inventory_pair_rules?.length ? result.inventory_pair_rules : salesInventoryPairRules, includeCompanyPrefix: result.include_company_prefix ?? includeCompanyPrefix, salesPrefixStrategy: loadedPrefixStrategy, prefixMstDigits: loadedPrefixMstDigits, salesPrefixStrategyValues: nextPrefixValues, salesReviewRules: result.vietmax_ban_ra_sales_internal_merges ?? salesReviewRules, salesReviewRows: [], salesReviewGenerated: false });
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

  function applyCompanyAndProductChoices(nextStage?: unknown) {
    const targetStage = isStageId(nextStage) ? nextStage : undefined;
    const activeStrategy = normalizedPrefixStrategy(purchasePrefixStrategy);
    const nextPrefixValues = rememberManualPrefixValues(purchasePrefixStrategyValues, activeStrategy, companyRows, prefixMstDigits);
    const nextCompanyRows = sortAppliedCompanyRows(companyRows.map((company) => {
      const process = company.pending_process ?? company.process ?? true;
      return { ...company, value: normalizePrefixValue(company.value), process, pending_process: process, committed_prefix: normalizePrefixValue(company.value) };
    }));
    const nextWorkflow = {
      ...workflow,
      ...purchaseOutputInvalidation(),
      companyRows: nextCompanyRows,
      purchasePrefixStrategyValues: nextPrefixValues,
      selectedCompanyIndex: firstDisplayedCompanyIndex(nextCompanyRows),
      purchaseReviewRows: [],
      purchaseReviewGenerated: false,
      ...(targetStage ? { stage: targetStage } : {}),
    };
    updateWorkflow(profile, nextWorkflow);
    scrollStageBodyToTop();
    void saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu lựa chọn công ty, hàng hóa vào cấu hình. Review Mã VT sẽ tạo lại theo lựa chọn mới.', 'purchase');
  }

  function applySalesCompanyAndProductChoices(nextStage?: unknown) {
    const targetStage = isStageId(nextStage) ? nextStage : undefined;
    const activeStrategy = normalizedPrefixStrategy(salesPrefixStrategy);
    const nextPrefixValues = rememberManualPrefixValues(salesPrefixStrategyValues, activeStrategy, salesCompanyRows, prefixMstDigits);
    const nextSalesCompanyRows = sortAppliedCompanyRows(salesCompanyRows.map((company) => {
      const process = company.pending_process ?? company.process ?? true;
      return { ...company, value: normalizePrefixValue(company.value), process, pending_process: process, committed_prefix: normalizePrefixValue(company.value) };
    }));
    const nextWorkflow = {
      ...workflow,
      ...salesOutputInvalidation(),
      salesCompanyRows: nextSalesCompanyRows,
      salesPrefixStrategyValues: nextPrefixValues,
      selectedSalesCompanyIndex: firstDisplayedCompanyIndex(nextSalesCompanyRows),
      salesReviewRows: [],
      salesReviewGenerated: false,
      ...(targetStage ? { stage: targetStage } : {}),
    };
    updateWorkflow(profile, nextWorkflow);
    scrollStageBodyToTop();
    void saveWorkflowConfig(nextWorkflow, 'Đã áp dụng và lưu lựa chọn công ty, hàng hóa bán ra. Review bán ra sẽ tạo lại theo lựa chọn mới.', 'sales');
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
    const company = companyRows[index];
    if (!company) return;
    const activeStrategy = normalizedPrefixStrategy(purchasePrefixStrategy);
    const nextValue = value;
    updateWorkflow(profile, {
      selectedCompanyIndex: index,
      companyRows: companyRows.map((row, rowIndex) => (rowIndex === index ? { ...row, value: nextValue } : row)),
      purchasePrefixStrategyValues: rememberPrefixEdit(purchasePrefixStrategyValues, activeStrategy, company, nextValue, prefixMstDigits),
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
      salesPrefixStrategyValues: rememberPrefixEdit(salesPrefixStrategyValues, activeStrategy, company, nextValue, prefixMstDigits),
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

  function updatePrefixMstDigits(digits: number) {
    const nextDigits = clampPrefixMstDigits(digits);
    const activeStrategy = normalizedPrefixStrategy(stage === 9 ? salesPrefixStrategy : purchasePrefixStrategy);
    const nextWorkflow: Partial<WorkflowState> = { prefixMstDigits: nextDigits };
    if (activeStrategy !== 'last_2_words') {
      if (stage === 3) {
        const rememberedValues = rememberManualPrefixValues(purchasePrefixStrategyValues, activeStrategy, companyRows, prefixMstDigits);
        nextWorkflow.purchasePrefixStrategyValues = rememberedValues;
        nextWorkflow.companyRows = applyPrefixStrategyRows(companyRows, activeStrategy, nextDigits, rememberedValues);
      }
      if (stage === 9) {
        const rememberedValues = rememberManualPrefixValues(salesPrefixStrategyValues, activeStrategy, salesCompanyRows, prefixMstDigits);
        nextWorkflow.salesPrefixStrategyValues = rememberedValues;
        nextWorkflow.salesCompanyRows = applyPrefixStrategyRows(salesCompanyRows, activeStrategy, nextDigits, rememberedValues);
      }
    }
    updateWorkflow(profile, nextWorkflow);
  }

  function applyPurchasePrefixPreset(strategy: PrefixPresetStrategy) {
    const currentStrategy = normalizedPrefixStrategy(purchasePrefixStrategy);
    const rememberedValues = rememberManualPrefixValues(purchasePrefixStrategyValues, currentStrategy, companyRows, prefixMstDigits);
    updateWorkflow(profile, {
      purchasePrefixStrategy: strategy,
      purchasePrefixStrategyValues: rememberedValues,
      companyRows: applyPrefixStrategyRows(companyRows, strategy, prefixMstDigits, rememberedValues),
    });
  }

  function applySalesPrefixPreset(strategy: PrefixPresetStrategy) {
    const currentStrategy = normalizedPrefixStrategy(salesPrefixStrategy);
    const rememberedValues = rememberManualPrefixValues(salesPrefixStrategyValues, currentStrategy, salesCompanyRows, prefixMstDigits);
    updateWorkflow(profile, {
      salesPrefixStrategy: strategy,
      salesPrefixStrategyValues: rememberedValues,
      salesCompanyRows: applyPrefixStrategyRows(salesCompanyRows, strategy, prefixMstDigits, rememberedValues),
    });
  }

  async function refreshSalesProductPreviews() {
    if (!salesCompanyRows.length) return;
    setBusy(true);
    setStatus('Đang cập nhật Mã VT preview bán ra...');
    try {
      const previewCodes = await loadProductPreviewCodes(salesCompanyRows, salesWordRules, salesRepeatedPhraseRemovals, 'sales');
      const nextWorkflow = { ...workflow, ...salesOutputInvalidation(), salesProductPreviewCodes: previewCodes, salesReviewRows: [], salesReviewGenerated: false };
      updateWorkflow(profile, nextWorkflow);
      for (const payload of buildConfigPayloads(nextWorkflow, 'sales')) {
        await saveVietmaxConfig(payload);
      }
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

  async function saveWorkflowConfig(targetWorkflow: WorkflowState, successMessage = 'Đã lưu cấu hình hiện tại.', phase: 'purchase' | 'sales' | 'all' = 'all') {
    setBusy(true);
    setStatus('Đang lưu cấu hình...');
    try {
      for (const payload of buildConfigPayloads(targetWorkflow, phase, profile)) {
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
    void saveWorkflowConfig(workflow, 'Đã lưu cấu hình hiện tại.', stage >= 6 ? 'sales' : 'purchase');
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
    const requiredStage = phase === 'purchase' ? 1 : 6;
    if (stage !== requiredStage) {
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

  async function refreshProductPreviews() {
    if (!companyRows.length) return;
    setBusy(true);
    setStatus('Đang cập nhật Mã VT preview theo từ thay riêng và từ lặp...');
    try {
      const previewCodes = isGenericProfile
        ? await loadGenericProductPreviewCodes(profile, companyRows, wordRules, firstWordRules, repeatedPhraseRemovals)
        : await loadProductPreviewCodes(companyRows, purchaseWordRules, purchaseRepeatedPhraseRemovals);
      const nextWorkflow = { ...workflow, ...purchaseOutputInvalidation(), productPreviewCodes: previewCodes, purchaseReviewRows: [], purchaseReviewGenerated: false, priceGroups: [] };
      updateWorkflow(profile, nextWorkflow);
      for (const payload of buildConfigPayloads(nextWorkflow, 'purchase', profile)) {
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
    const entries = Object.entries(activeWordRules);
    entries[index] = field === 'from' ? [value, entries[index]?.[1] || ''] : [entries[index]?.[0] || '', value];
    const nextRules = Object.fromEntries(entries.filter(([from]) => from.trim()));
    const invalidation = inventoryOutputInvalidation();
    if (profile === 'vietmax' && activeVietmaxSalesConfig) {
      updateWorkflow(profile, { ...invalidation, salesWordRules: nextRules, salesReviewRows: [], salesReviewGenerated: false });
      return;
    }
    if (profile === 'vietmax') {
      updateWorkflow(profile, { ...invalidation, purchaseWordRules: nextRules, purchaseReviewRows: [], purchaseReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, wordRules: nextRules, purchaseReviewRows: [], salesReviewRows: [], purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function addWordRule() {
    const nextRules = { ...activeWordRules, '': '' };
    const invalidation = inventoryOutputInvalidation();
    if (profile === 'vietmax' && activeVietmaxSalesConfig) {
      updateWorkflow(profile, { ...invalidation, salesWordRules: nextRules, salesReviewGenerated: false });
      return;
    }
    if (profile === 'vietmax') {
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
    if (profile === 'vietmax' && activeVietmaxSalesConfig) {
      updateWorkflow(profile, { ...invalidation, salesRepeatedPhraseRemovals: nextPhrases, salesReviewRows: [], salesReviewGenerated: false });
      return;
    }
    if (profile === 'vietmax') {
      updateWorkflow(profile, { ...invalidation, purchaseRepeatedPhraseRemovals: nextPhrases, purchaseReviewRows: [], purchaseReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, repeatedPhraseRemovals: nextPhrases, purchaseReviewRows: [], salesReviewRows: [], purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function addRepeatedPhrase() {
    const nextPhrases = [...activeRepeatedPhraseRemovals, ''];
    const invalidation = inventoryOutputInvalidation();
    if (profile === 'vietmax' && activeVietmaxSalesConfig) {
      updateWorkflow(profile, { ...invalidation, salesRepeatedPhraseRemovals: nextPhrases, salesReviewGenerated: false });
      return;
    }
    if (profile === 'vietmax') {
      updateWorkflow(profile, { ...invalidation, purchaseRepeatedPhraseRemovals: nextPhrases, purchaseReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, repeatedPhraseRemovals: nextPhrases, purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  function removeRepeatedPhrase(index: number) {
    const next = activeRepeatedPhraseRemovals.slice();
    next.splice(index, 1);
    const invalidation = inventoryOutputInvalidation();
    if (profile === 'vietmax' && activeVietmaxSalesConfig) {
      updateWorkflow(profile, { ...invalidation, salesRepeatedPhraseRemovals: next, salesReviewRows: [], salesReviewGenerated: false });
      return;
    }
    if (profile === 'vietmax') {
      updateWorkflow(profile, { ...invalidation, purchaseRepeatedPhraseRemovals: next, purchaseReviewRows: [], purchaseReviewGenerated: false });
      return;
    }
    updateWorkflow(profile, { ...invalidation, repeatedPhraseRemovals: next, purchaseReviewRows: [], salesReviewRows: [], purchaseReviewGenerated: false, salesReviewGenerated: false });
  }

  async function downloadProcessedPurchase() {
    if (!purchaseFile) return;
    setBusy(true);
    const progress = processedPurchaseSavedName ? null : beginProgress('Đang tạo file mua vào đã xử lý');
    setStatus(processedPurchaseSavedName ? 'Đang mở file mua vào đã xử lý từ cache...' : 'Đang tạo file mua vào đã xử lý Mã VT...');
    try {
      let savedName = processedPurchaseSavedName;
      let stats = processedPurchaseStats;
      if (!savedName) {
        const result = await processVietmaxPurchase(purchaseFile.saved_name, purchaseFile.original_name, buildPurchaseProcessPayload(workflow), { cacheOnly: true, operationId: progress?.operationId });
        savedName = result.processedSavedName;
        if (!savedName) throw new Error('Không tạo được cache file mua vào đã xử lý.');
        stats = await applyProcessedPurchaseCache(profile, savedName);
      }
      const blob = await downloadCachedFile(savedName);
      const saved = await saveBlob(blob, purchaseFile.original_name.replace(/\.(xls|xlsx|xlsm)$/i, '_fdi.xls'));
      const suffix = processedStatsSentence(stats);
      setStatus(saved ? `Đã xuất file mua vào đã xử lý. ${suffix}` : `Đã hủy lưu file mua vào; cache vẫn được giữ. ${suffix}`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress?.stop();
      if (progress) setLoadingProgress(null);
      setBusy(false);
    }
  }

  async function downloadProcessedSales() {
    if (!salesFile) return;
    setBusy(true);
    const progress = processedSalesSavedName ? null : beginProgress('Đang tạo file bán ra đã xử lý');
    setStatus(processedSalesSavedName ? 'Đang mở file bán ra đã xử lý từ cache...' : 'Đang tạo file bán ra đã xử lý Mã VT...');
    try {
      let savedName = processedSalesSavedName;
      let stats = processedSalesStats;
      if (!savedName) {
        const result = await processVietmaxPurchase(salesFile.saved_name, salesFile.original_name, buildSalesProcessPayload(workflow), { cacheOnly: true, operationId: progress?.operationId });
        savedName = result.processedSavedName;
        if (!savedName) throw new Error('Không tạo được cache file bán ra đã xử lý.');
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
      progress?.stop();
      if (progress) setLoadingProgress(null);
      setBusy(false);
    }
  }

  function changeProfile(nextProfile: ProfileKey) {
    setProfile(nextProfile);
    setStatus(`Đang xem profile ${profiles.find((item) => item.key === nextProfile)?.label ?? nextProfile}. Dữ liệu profile khác vẫn được giữ.`);
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

  async function runInventoryAllocation(nextStage?: StageId) {
    if (!processedPurchaseSavedName || !processedSalesSavedName) {
      setStatus('Cần tải hoặc tạo cả file mua vào và bán ra đã xử lý trước khi phân bổ tồn kho.');
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
      const allocationConfig = { ...inventoryAllocationConfig, policy: { ...inventoryAllocationConfig.policy, company_profile: inventoryAllocationProfileFor(targetProfile) } };
      const started = await startInventoryAllocation({ purchaseSavedName: processedPurchaseSavedName, salesSavedName: processedSalesSavedName, salesOriginalName: salesFile?.original_name || 'ban_ra_da_xu_ly.xlsx', openingFile: openingStockFile, config: allocationConfig });
      let nextJob: InventoryAllocationJob = { status: 'queued', progress: 0, done: 0, total: 0, label: 'Đã gửi dữ liệu. Đang chờ backend xử lý...' };
      updateWorkflow(targetProfile, { inventoryAllocationJob: nextJob, inventoryAllocationResult: null });
      while (nextJob.status !== 'complete') {
        await sleep(500);
        nextJob = await getInventoryAllocationJob(started.analysis_job_id);
        updateWorkflow(targetProfile, { inventoryAllocationJob: nextJob, inventoryAllocationResult: nextJob.result ?? null });
        setStatus(formatInventoryJobStatus(nextJob));
        if (nextJob.status === 'error') throw new Error(nextJob.error || nextJob.label || 'Phân bổ tồn kho thất bại.');
      }
      if (nextStage) updateWorkflow(targetProfile, { stage: nextStage });
      setStatus('Đã hoàn tất phân bổ tồn kho. Kiểm tra báo cáo trước khi xuất file.');
    } catch (error) {
      updateWorkflow(targetProfile, { inventoryAllocationJob: null });
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
      const filename = toXlsName(inventoryAllocationResult?.filename || inventoryAllocationJob?.result?.filename || 'phan_bo_ton_kho.xls');
      const saved = await saveBlob(blob, filename);
      setStatus(saved ? 'Đã lưu báo cáo phân bổ tồn kho.' : 'Đã hủy lưu báo cáo phân bổ tồn kho.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function downloadFastImportPackage() {
    if (!processedPurchaseSavedName || !processedSalesSavedName) {
      setStatus('Cần có cả FDI mua vào và FDI bán ra đã xử lý trước khi tạo workbook FAST.');
      return;
    }
    const progress = beginProgress('Đang tạo workbook FAST 4 sheet');
    setBusy(true);
    setStatus('Đang tạo workbook FAST gồm Hoadonmuahang, Hoadonbanhang, DM vật tư và DM khách hàng từ FDI đã xử lý...');
    try {
      const blob = await createVietmaxFastImportPackage(processedPurchaseSavedName, processedSalesSavedName, progress.operationId);
      const saved = await saveBlob(blob, 'vietmax_fast_import.xls');
      setStatus(saved ? 'Đã lưu workbook FAST 4 sheet.' : 'Đã hủy lưu workbook FAST; dữ liệu đã xử lý vẫn được giữ.');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      progress.stop();
      setLoadingProgress(null);
      setBusy(false);
    }
  }

  function inventoryStateForScope(scope: InventoryConfigScope) {
    if (profile === 'vietmax' && scope === 'sales') {
      return {
        pairs: salesInventoryPairs,
        useDefault: salesUseDefaultInventoryPair,
        defaultPairId: salesDefaultInventoryPairId,
        rules: salesInventoryPairRules,
      };
    }
    if (profile === 'vietmax' && scope === 'purchase') {
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
    if (profile === 'vietmax' && scope === 'sales') {
      updateWorkflow(profile, {
        ...invalidation,
        ...(update.pairs ? { salesInventoryPairs: update.pairs } : {}),
        ...(update.useDefault !== undefined ? { salesUseDefaultInventoryPair: update.useDefault } : {}),
        ...(update.defaultPairId !== undefined ? { salesDefaultInventoryPairId: update.defaultPairId } : {}),
        ...(update.rules ? { salesInventoryPairRules: update.rules } : {}),
      });
      return;
    }
    if (profile === 'vietmax' && scope === 'purchase') {
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

  function inventoryOutputInvalidation(scope: InventoryConfigScope = activeInventoryConfigScope) {
    if (profile === 'vietmax') {
      return scope === 'sales' ? salesOutputInvalidation() : purchaseOutputInvalidation();
    }
    return stage >= 6 ? salesOutputInvalidation() : purchaseOutputInvalidation();
  }

  const nextStage = visibleStages[visibleStages.findIndex((item) => item.id === stage) + 1];
  const nextDisabled = busy || !nextStage || (stage === 12 ? !(processedPurchaseSavedName && processedSalesSavedName) : !canEnterStage(nextStage.id));

  return (
    <main className="desktop-shell">
      <section className={`app-card ${showLicenseBar ? '' : 'compact-flow'} ${profile === 'cao_thanh' ? 'legacy-flow' : ''}`}>
        <header className="app-header">
          <div className="profile-toolbar" aria-label="Company profile controls">
            <label className="profile-dropdown"><span>Công ty áp dụng</span><select value={profile} disabled={busy} onChange={(event) => changeProfile(event.currentTarget.value as ProfileKey)}>{profiles.map((item) => <option key={item.key} value={item.key}>{item.label}</option>)}</select></label>
            <button type="button" className="btn-secondary" disabled={busy} onClick={saveCurrentConfig}>Lưu cấu hình</button>
            {profile === 'vietmax' && (
              <div className="config-export-actions" aria-label="Xuất/nhập cấu hình Vietmax">
                <button type="button" className="btn-secondary" disabled={busy} onClick={() => void exportCurrentVietmaxConfig('purchase')}>Xuất cấu hình mua vào</button>
                <button type="button" className="btn-secondary" disabled={busy} onClick={() => void exportCurrentVietmaxConfig('sales')}>Xuất cấu hình bán ra</button>
                <button type="button" className="btn-secondary" disabled={busy || stage !== 1} title="Chỉ nhập ở stage 1 để tránh conflict dữ liệu đang xử lý" onClick={() => void importCurrentVietmaxConfig('purchase')}>Nhập cấu hình mua vào</button>
                <button type="button" className="btn-secondary" disabled={busy || stage !== 6} title="Chỉ nhập ở stage 6 để tránh conflict dữ liệu đang xử lý" onClick={() => void importCurrentVietmaxConfig('sales')}>Nhập cấu hình bán ra</button>
              </div>
            )}
          </div>
          {usesNativeStageShell && <StageNavigation stages={visibleStages} stage={stage} busy={busy} canEnterStage={canEnterStage} goToStage={goToStage} />}
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
                <input placeholder="License server/IP, vd 192.168.1.10:3000" value={licenseForm.server_url} onChange={(event) => setLicenseForm({ ...licenseForm, server_url: event.currentTarget.value })} />
                <input placeholder="LICENSE_KEY" type="password" value={licenseForm.license_key} onChange={(event) => setLicenseForm({ ...licenseForm, license_key: event.currentTarget.value })} />
                <button type="button" disabled={busy} onClick={submitLicense}>Kích hoạt</button>
              </div>
            )}
            {license?.activated && <button type="button" className="btn-secondary" disabled={busy} onClick={refreshLicense}>Tải lại license</button>}
          </section>
        )}

        <section className="stage-frame">
          <div className={`stage-body ${profile === 'cao_thanh' ? 'legacy-stage-body' : ''}`}>
            {profile !== 'vietmax' ? renderProfileStage() : renderVietmaxStage()}
          </div>
        </section>

        {usesNativeStageShell && (
          <footer className="action-bar">
            <button type="button" className="btn-secondary" disabled={stage === 1 || busy} onClick={goBack}>Quay lại</button>
            <button type="button" className="btn-danger" disabled={busy} onClick={resetWorkflow}>Làm lại</button>
            <div className="action-spacer" />
            <button type="button" disabled={nextDisabled} onClick={goNext}>Tiếp tục</button>
          </footer>
        )}
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
        return <CompanyRulesStage companies={companyRows} selectedCompanyIndex={selectedCompanyIndex} productPreviewCodes={productPreviewCodes} productCodeOverrides={productCodeOverrides} wordRules={purchaseWordRules} repeatedPhrases={purchaseRepeatedPhraseRemovals} inventoryPairs={purchaseInventoryPairs} useDefaultInventoryPair={purchaseUseDefaultInventoryPair} defaultInventoryPairId={purchaseDefaultInventoryPairId} inventoryPairRules={purchaseInventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy={purchasePrefixStrategy} prefixMstDigits={prefixMstDigits} onIncludeCompanyPrefixChange={updateIncludeCompanyPrefix} onCompanyPrefixChange={updateCompanyPrefix} onPrefixMstDigitsChange={updatePrefixMstDigits} onApplyPrefixPresetToAll={applyPurchasePrefixPreset} onCompanySelect={selectCompany} onCompanyChange={updatePendingCompany} onBulkCompanyChange={bulkUpdatePendingCompanies} onProductChange={updateCompanyProduct} onProductCodeChange={updateProductCode} onApplyChoices={applyCompanyAndProductChoices} onRefreshPreviews={refreshProductPreviews} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeatedPhrase} onAddRepeated={addRepeatedPhrase} onRemoveRepeated={removeRepeatedPhrase} onAddInventoryPair={() => addInventoryPair('purchase')} onInventoryPairChange={(index, field, value) => updateInventoryPair(index, field, value, 'purchase')} onRemoveInventoryPair={(index) => removeInventoryPair(index, 'purchase')} onInventoryDefaultsChange={(update) => updateInventoryDefaults(update, 'purchase')} onAddInventoryRule={() => addInventoryRule('purchase')} onInventoryRuleChange={(index, update) => updateInventoryRule(index, update, 'purchase')} onRemoveInventoryRule={(index) => removeInventoryRule(index, 'purchase')} />;
      case 4:
        if (busy || !purchaseReviewGenerated) return <LoadingStage title="Đang tạo Review Mã VT mua vào" detail="Đang so sánh tên hàng và dựng danh sách mã cần kiểm tra..." progress={loadingProgress} />;
        return <ReviewStage rows={purchaseReviewRows} onApply={applyReviewChoices} disabled={!purchaseFile || busy} onRowChange={updateReviewRow} onBulkChange={bulkUpdateReviewRows} title="Review Mã VT mua vào" empty="Không có dòng Mã VT cần review." reviewScope={purchaseReviewScope} onReviewScopeChange={updatePurchaseReviewScope} />;
      case 5:
        if (busy) return <LoadingStage title="Đang tạo file mua vào" detail="Đang xử lý workbook và tạo cache file mua vào để dùng cho các stage bán ra..." progress={loadingProgress} />;
        return <ProcessStage title="Tạo file mua vào" detail="Xuất file FDI mua vào đã xử lý bằng logic Vietmax. File này sẽ được cache để dùng cho khớp mua/bán, phân bổ tồn kho và xuất FAST ở stage 15." buttonLabel="Xuất file mua vào" disabled={busy || !purchaseFile} onProcess={downloadProcessedPurchase} />;
      case 6:
        return <SalesEntryStage salesFile={salesFile} processedPurchaseReady={Boolean(processedPurchaseSavedName)} processedPurchaseStats={processedPurchaseStats} disabled={busy || !licenseReady} onSalesUpload={(file) => upload('sales', file)} onProcessedPurchaseUpload={(file) => uploadProcessed('purchase', file)} />;
      case 7:
        return <MappingStage summary={salesFile} phase="sales" scope={comparisonScope} setScope={updateComparisonScope} />;
      case 8:
        if (busy) return <LoadingStage title="Đang khớp mua vào / bán ra" detail={processedPurchaseSavedName ? 'Đang so sánh hàng bán ra với file mua vào đã xử lý và áp dụng cấu hình khớp đã lưu...' : 'Đang tạo cache mua vào rồi khớp với file bán ra...'} progress={loadingProgress} />;
        return <MatchStage rows={matches} disabled={!salesFile || busy || (!purchaseFile && !processedPurchaseSavedName)} onRun={runSalesMatch} onSave={saveMatchChoices} onToggle={toggleMatch} onBulkToggle={bulkToggleMatches} onConversionChange={updateMatchConversion} autoRun={Boolean(salesFile && (purchaseFile || processedPurchaseSavedName)) && matches.length === 0 && !busy} emptyMessage={processedPurchaseSavedName || purchaseFile ? undefined : 'Cần tải file mua vào đã xử lý trước khi khớp mua/bán.'} />;
      case 9:
        if (busy && !salesCompanyRows.length) return <LoadingStage title="Đang tải danh sách công ty bán ra" detail="Đang lọc các hàng hóa chưa khớp KVT/152 và gom theo công ty..." />;
        return <CompanyRulesStage companies={salesCompanyRows} selectedCompanyIndex={selectedSalesCompanyIndex} productPreviewCodes={salesProductPreviewCodes} productCodeOverrides={salesProductCodeOverrides} wordRules={salesWordRules} repeatedPhrases={salesRepeatedPhraseRemovals} inventoryPairs={salesInventoryPairs} useDefaultInventoryPair={salesUseDefaultInventoryPair} defaultInventoryPairId={salesDefaultInventoryPairId} inventoryPairRules={salesInventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy={salesPrefixStrategy} prefixMstDigits={prefixMstDigits} onIncludeCompanyPrefixChange={updateIncludeCompanyPrefix} onCompanyPrefixChange={updateSalesCompanyPrefix} onPrefixMstDigitsChange={updatePrefixMstDigits} onApplyPrefixPresetToAll={applySalesPrefixPreset} onCompanySelect={selectSalesCompany} onCompanyChange={updateSalesPendingCompany} onBulkCompanyChange={bulkUpdateSalesPendingCompanies} onProductChange={updateSalesCompanyProduct} onProductCodeChange={updateSalesProductCode} onApplyChoices={applySalesCompanyAndProductChoices} onRefreshPreviews={refreshSalesProductPreviews} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeatedPhrase} onAddRepeated={addRepeatedPhrase} onRemoveRepeated={removeRepeatedPhrase} onAddInventoryPair={() => addInventoryPair('sales')} onInventoryPairChange={(index, field, value) => updateInventoryPair(index, field, value, 'sales')} onRemoveInventoryPair={(index) => removeInventoryPair(index, 'sales')} onInventoryDefaultsChange={(update) => updateInventoryDefaults(update, 'sales')} onAddInventoryRule={() => addInventoryRule('sales')} onInventoryRuleChange={(index, update) => updateInventoryRule(index, update, 'sales')} onRemoveInventoryRule={(index) => removeInventoryRule(index, 'sales')} />;
      case 10:
        if (busy || !salesReviewGenerated) return <LoadingStage title="Đang tạo Review Mã VT bán ra" detail="Đang tạo danh sách review theo công ty/hàng hóa bán ra đã áp dụng..." progress={loadingProgress} />;
        return <ReviewStage rows={salesReviewRows} onApply={applySalesReviewChoices} disabled={!salesFile || busy} onRowChange={updateSalesReviewRow} onBulkChange={bulkUpdateSalesReviewRows} title="Review Mã VT bán ra" empty="Không có dòng Mã VT bán ra cần review." reviewScope={salesReviewScope} onReviewScopeChange={updateSalesReviewScope} />;
      case 11:
        if (busy) return <LoadingStage title="Đang tạo file bán ra" detail="Đang xử lý workbook bán ra, áp dụng khớp mua vào và lưu cache cho phân bổ tồn kho..." progress={loadingProgress} />;
        return <ProcessStage title="Tạo file bán ra" detail="Xuất file FDI bán ra đã xử lý bằng logic Vietmax. File này sẽ được cache để dùng cho phân bổ tồn kho và xuất FAST ở stage 15." buttonLabel="Xuất file bán ra" disabled={busy || !salesFile} onProcess={downloadProcessedSales} />;
      case 12:
        if (isInventoryAllocationRunning(inventoryAllocationJob)) return <LoadingStage title="Đang phân bổ tồn kho" detail={inventoryAllocationJob?.label || 'Đang chạy phân bổ từ file mua vào và bán ra đã xử lý...'} progress={inventoryJobProgress(inventoryAllocationJob)} />;
        return <InventoryAllocationStage purchaseFile={purchaseFile} salesFile={salesFile} processedPurchaseSavedName={processedPurchaseSavedName} processedSalesSavedName={processedSalesSavedName} processedPurchaseStats={processedPurchaseStats} processedSalesStats={processedSalesStats} openingStockFile={openingStockFile} config={inventoryAllocationConfig} busy={busy} onProcessedPurchaseFileChange={(file) => uploadProcessed('purchase', file)} onProcessedSalesFileChange={(file) => uploadProcessed('sales', file)} onOpeningStockFileChange={updateOpeningStockFile} onConfigChange={updateInventoryAllocationConfig} />;
      case 13:
        return <InventoryAllocationReportStage result={inventoryAllocationResult ?? inventoryAllocationJob?.result ?? null} busy={busy} />;
      case 14:
        return <InventoryAllocationExportStage result={inventoryAllocationResult ?? inventoryAllocationJob?.result ?? null} busy={busy} onDownload={downloadInventoryReport} />;
      case 15:
        return <FastImportExportStage processedPurchaseSavedName={processedPurchaseSavedName} processedSalesSavedName={processedSalesSavedName} processedPurchaseStats={processedPurchaseStats} processedSalesStats={processedSalesStats} busy={busy} onProcessedPurchaseUpload={(file) => uploadFastImportProcessed('purchase', file)} onProcessedSalesUpload={(file) => uploadFastImportProcessed('sales', file)} onDownload={downloadFastImportPackage} />;
      default:
        return null;
    }
  }

  function renderProfileStage() {
    if (isGenericProfile) {
      if (Number(stage) === 4) {
        if (busy || (!purchaseReviewGenerated && purchaseFile && companyRows.length)) return <LoadingStage title={`Dang tao Review Ma VT ${selectedProfile.label}`} detail="Dang so sanh cac ten hang gan giong nhau theo danh sach cong ty/hang hoa da ap dung..." progress={loadingProgress} />;
        return <ReviewStage rows={purchaseReviewRows} onApply={applyReviewChoices} disabled={!purchaseFile || busy} onRowChange={updateReviewRow} onBulkChange={bulkUpdateReviewRows} title={`Review Ma VT ${selectedProfile.label}`} empty="Khong co dong Ma VT can review." reviewScope={purchaseReviewScope} onReviewScopeChange={updatePurchaseReviewScope} />;
      }
      if (Number(stage) === 5 && profile === 'cao_thanh') {
        return <CaoThanhPriceStage groups={priceGroups} filterPercent={priceFilterAllPercent} marginPercent={priceAdjustAllPercent} busy={busy} onRefresh={updateCaoThanhPriceGroups} onGroupPercentChange={updateCaoThanhGroupPercent} onBucketMarginChange={updateCaoThanhBucketMargin} onFilterPercentChange={updateCaoThanhPriceFilterAllPercent} onMarginPercentChange={updateCaoThanhPriceAdjustAllPercent} onApplyFilter={applyCaoThanhBulkPriceFilter} onApplyMargin={applyCaoThanhBulkMargin} onExportReport={exportCaoThanhPriceReport} />;
      }
      if ((Number(stage) === 5 && profile !== 'cao_thanh') || Number(stage) === 6) {
        if (busy) return <LoadingStage title={`Dang tao file ${selectedProfile.label}`} detail="Dang xu ly workbook va dong goi file ket qua..." />;
        return <ProcessStage title={`Xuat file ${selectedProfile.label}`} detail="Xuat file da xu ly Ma VT va file nhap kho di kem theo cau hinh cong ty/hang hoa hien tai." buttonLabel="Xuat file ket qua" disabled={busy || !purchaseFile || !companyRows.length} onProcess={downloadGenericProcessedFile} />;
      }
      switch (stage) {
        case 1:
          return <UploadStage title={selectedProfile.label} summary={purchaseFile} disabled={busy || !licenseReady} onUpload={(file) => upload('purchase', file)} />;
        case 2:
          return <GenericMappingStage summary={purchaseFile} columns={genericColumns} onColumnsChange={updateGenericColumns} />;
        case 3:
          if (busy && !companyRows.length) return <LoadingStage title={`Đang tải danh sách công ty ${selectedProfile.label}`} detail="Đang đọc workbook và gom công ty/MST/hàng hóa..." />;
          return <CompanyRulesStage companies={companyRows} selectedCompanyIndex={selectedCompanyIndex} productPreviewCodes={productPreviewCodes} productCodeOverrides={productCodeOverrides} wordRules={wordRules} repeatedPhrases={repeatedPhraseRemovals} inventoryPairs={inventoryPairs} useDefaultInventoryPair={useDefaultInventoryPair} defaultInventoryPairId={defaultInventoryPairId} inventoryPairRules={inventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy={purchasePrefixStrategy} prefixMstDigits={prefixMstDigits} onIncludeCompanyPrefixChange={updateIncludeCompanyPrefix} onCompanyPrefixChange={updateCompanyPrefix} onPrefixMstDigitsChange={updatePrefixMstDigits} onApplyPrefixPresetToAll={applyPurchasePrefixPreset} onCompanySelect={selectCompany} onCompanyChange={updatePendingCompany} onBulkCompanyChange={bulkUpdatePendingCompanies} onProductChange={updateCompanyProduct} onProductCodeChange={updateProductCode} onApplyChoices={applyCompanyAndProductChoices} onRefreshPreviews={refreshProductPreviews} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeatedPhrase} onAddRepeated={addRepeatedPhrase} onRemoveRepeated={removeRepeatedPhrase} onAddInventoryPair={() => addInventoryPair('generic')} onInventoryPairChange={(index, field, value) => updateInventoryPair(index, field, value, 'generic')} onRemoveInventoryPair={(index) => removeInventoryPair(index, 'generic')} onInventoryDefaultsChange={(update) => updateInventoryDefaults(update, 'generic')} onAddInventoryRule={() => addInventoryRule('generic')} onInventoryRuleChange={(index, update) => updateInventoryRule(index, update, 'generic')} onRemoveInventoryRule={(index) => removeInventoryRule(index, 'generic')} />;
        case 4:
          if (busy) return <LoadingStage title={`Đang tạo file ${selectedProfile.label}`} detail="Đang xử lý workbook và đóng gói file kết quả..." />;
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
  if (phase === 'fast') return 'Xuất FAST';
  if (phase === 'price') return 'Lọc đơn giá';
  return 'Profile';
}

function UploadStage({ title, summary, disabled, onUpload }: { title: string; summary: UploadSummary | null; disabled: boolean; onUpload: (file: File | undefined) => void }) {
  const stepNumber = title.includes('bán') ? 6 : 1;
  const uploadLabel = `Chọn ${title}`;
  return (
    <div className="center-stage">
      <span className="upload-step-badge">BƯỚC {stepNumber}</span>
      <p className="description">Tải lên file Excel <strong>.xls</strong>, <strong>.xlsx</strong> hoặc <strong>.xlsm</strong> cho <strong>{title}</strong>.</p>
      <label className={`upload-label ${summary ? 'has-file' : ''}`}>
        <input className="upload-input" type="file" accept=".xls,.xlsx,.xlsm" disabled={disabled} onChange={(event) => onUpload(event.currentTarget.files?.[0])} />
        <span className="upload-mark">XLSX</span>
        <strong>{summary?.original_name || 'Kéo thả hoặc chọn file Excel'}</strong>
        <small>{summary ? `${summary.columns.length} cột đã đọc` : 'Hệ thống sẽ đọc bảng, nhận diện cột và giữ nguyên file gốc.'}</small>
        <span className="upload-button">{summary ? `Đổi ${title}` : uploadLabel}</span>
      </label>
    </div>
  );
}

function SalesEntryStage({ salesFile, processedPurchaseReady, processedPurchaseStats, disabled, onSalesUpload, onProcessedPurchaseUpload }: { salesFile: UploadSummary | null; processedPurchaseReady: boolean; processedPurchaseStats: ProcessedFileStats | null; disabled: boolean; onSalesUpload: (file: File | undefined) => void; onProcessedPurchaseUpload: (file: File | undefined) => void }) {
  return (
    <div className="skip-entry-stage">
      <section className="skip-required-panel">
        <div className="stage-toolbar">
          <h3>Bỏ qua mua vào</h3>
          <p>Upload file mua vào đã xử lý trước khi chạy các bước bán ra.</p>
        </div>
        <ProcessedFileUpload label="Mua vào đã xử lý" ready={processedPurchaseReady} stats={processedPurchaseStats} disabled={disabled} onUpload={onProcessedPurchaseUpload} />
      </section>
      <UploadStage title="HD bán ra" summary={salesFile} disabled={disabled} onUpload={onSalesUpload} />
    </div>
  );
}

function ProcessedFileUpload({ label, ready, stats, disabled, onUpload }: { label: string; ready: boolean; stats: ProcessedFileStats | null; disabled: boolean; onUpload: (file: File | undefined) => void }) {
  return (
    <label className={`processed-file-upload ${ready ? 'has-file' : ''}`}>
      <input className="upload-input" type="file" accept=".xls,.xlsx,.xlsm" disabled={disabled} onChange={(event) => onUpload(event.currentTarget.files?.[0])} />
      <span className="upload-mark">FDI</span>
      <strong>{ready ? `${label} đã sẵn sàng` : `Chọn ${label}`}</strong>
      <small>{ready ? 'Có thể dùng file này cho stage sau.' : 'Bắt buộc khi bỏ qua các stage xử lý trước đó.'}</small>
      <ProcessedStatsSummary stats={stats} />
      <span className="upload-button">{ready ? `Đổi ${label}` : `Tải ${label}`}</span>
    </label>
  );
}

function ProcessedStatsSummary({ stats }: { stats: ProcessedFileStats | null }) {
  if (!stats) return <div className="processed-stats muted">Chưa có thống kê file đã xử lý.</div>;
  return (
    <div className="processed-stats">
      <span>Công ty <strong>{formatCount(stats.processed_company_count)} / {formatCount(stats.company_count)}</strong></span>
      <span>Dòng hàng <strong>{formatCount(stats.processed_product_row_count)} / {formatCount(stats.product_row_count)}</strong></span>
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

function GenericMappingStage({ summary, columns, onColumnsChange }: { summary: UploadSummary | null; columns: GenericColumns; onColumnsChange: (update: Partial<GenericColumns>) => void }) {
  if (!summary) return <PlaceholderStage title="Chưa có file" detail="Quay lại stage tải file trước khi chọn cột." />;
  const previewKeys = summary.preview.length ? Object.keys(summary.preview[0]) : [];
  const columnFields: Array<{ key: keyof GenericColumns; label: string; allowBlank?: boolean }> = [
    { key: 'company_col', label: 'Tên công ty' },
    { key: 'mst_col', label: 'MST' },
    { key: 'address_col', label: 'Địa chỉ', allowBlank: true },
    { key: 'product_col', label: 'Tên hàng' },
    { key: 'qty_col', label: 'Số lượng', allowBlank: true },
    { key: 'price_col', label: 'Đơn giá', allowBlank: true },
    { key: 'output_col', label: 'Mã VT' },
    { key: 'invoice_status_col', label: 'Trạng thái HĐ', allowBlank: true },
  ];
  return (
    <div className="stage-grid">
      <div className="form-panel">
        <p className="description left">File: <strong>{summary.original_name}</strong>. Chọn đúng cột trước khi tải danh sách công ty.</p>
        <div className="column-grid">
          {columnFields.map((field) => (
            <label key={field.key}>
              <span>{field.label}</span>
              <select value={String(columns[field.key] || '')} onChange={(event) => onColumnsChange({ [field.key]: event.currentTarget.value } as Partial<GenericColumns>)}>
                {field.allowBlank && <option value="">Không dùng</option>}
                {summary.columns.map((column) => <option key={`${field.key}-${column.letter}`} value={column.letter} title={column.label}>{column.letter}</option>)}
              </select>
            </label>
          ))}
        </div>
      </div>
      <div className="preview-panel">
        <h3>Xem trước dữ liệu</h3>
        <div className="preview-scroll"><table><thead><tr>{previewKeys.map((key) => <th key={key}>{key}</th>)}</tr></thead><tbody>{summary.preview.map((row, index) => <tr key={index}>{previewKeys.map((key) => <td key={key}>{row[key]}</td>)}</tr>)}</tbody></table></div>
      </div>
    </div>
  );
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

function ConfigModal({
  isOpen,
  onClose,
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
  const [activeTab, setActiveTab] = useState<'words' | 'repeat' | 'inventory'>('words');
  const wordEntries = Object.entries(wordRules);
  if (!isOpen) return null;
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-content advanced-config-modal" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h2>Cấu hình nâng cao</h2>
          <button className="modal-close" onClick={onClose}>×</button>
        </div>
        <div className="modal-body">
          <div className="tab-list">
            <button className={`tab-button ${activeTab === 'words' ? 'active' : ''}`} onClick={() => setActiveTab('words')}>Từ riêng</button>
            <button className={`tab-button ${activeTab === 'repeat' ? 'active' : ''}`} onClick={() => setActiveTab('repeat')}>Từ lặp</button>
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
  displayTitle?: string;
  key?: string;
  rows: Array<{ row: ReviewRow; index: number }>;
  className: string;
  children?: ReviewDisplayGroup[];
};

function CompanyRulesStage({ companies, selectedCompanyIndex, productPreviewCodes, productCodeOverrides, wordRules, repeatedPhrases, inventoryPairs, useDefaultInventoryPair, defaultInventoryPairId, inventoryPairRules, busy, showCompanyPrefixControls = false, includeCompanyPrefix = false, prefixStrategy = 'last_2_words', prefixMstDigits = 3, onIncludeCompanyPrefixChange, onCompanyPrefixChange, onPrefixMstDigitsChange, onApplyPrefixPresetToAll, onCompanySelect, onCompanyChange, onBulkCompanyChange, onProductChange, onProductCodeChange, onApplyChoices, onRefreshPreviews, onWordRuleChange, onAddWordRule, onRepeatedChange, onAddRepeated, onRemoveRepeated, onAddInventoryPair, onInventoryPairChange, onRemoveInventoryPair, onInventoryDefaultsChange, onAddInventoryRule, onInventoryRuleChange, onRemoveInventoryRule }: { companies: CompanyRow[]; selectedCompanyIndex: number; productPreviewCodes: Record<string, string>; productCodeOverrides: Record<string, string>; wordRules: Record<string, string>; repeatedPhrases: string[]; inventoryPairs: InventoryPair[]; useDefaultInventoryPair: boolean; defaultInventoryPairId: string; inventoryPairRules: InventoryRule[]; busy: boolean; showCompanyPrefixControls?: boolean; includeCompanyPrefix?: boolean; prefixStrategy?: string; prefixMstDigits?: number; onIncludeCompanyPrefixChange?: (include: boolean) => void; onCompanyPrefixChange?: (index: number, value: string) => void; onPrefixMstDigitsChange?: (digits: number) => void; onApplyPrefixPresetToAll?: (strategy: PrefixPresetStrategy) => void; onCompanySelect: (index: number) => void; onCompanyChange: (index: number, pending: boolean) => void; onBulkCompanyChange?: (pending: boolean) => void; onProductChange: (companyIndex: number, productName: string, selected: boolean) => void; onProductCodeChange: (companyIndex: number, productName: string, code: string) => void; onApplyChoices: () => void; onRefreshPreviews: () => void; onWordRuleChange: (index: number, field: 'from' | 'to', value: string) => void; onAddWordRule: () => void; onRepeatedChange: (index: number, value: string) => void; onAddRepeated: () => void; onRemoveRepeated: (index: number) => void; onAddInventoryPair: () => void; onInventoryPairChange: (index: number, field: 'ma_kho' | 'tk_vat_tu', value: string) => void; onRemoveInventoryPair: (index: number) => void; onInventoryDefaultsChange: (update: Partial<Pick<WorkflowState, 'useDefaultInventoryPair' | 'defaultInventoryPairId'>>) => void; onAddInventoryRule: () => void; onInventoryRuleChange: (index: number, update: Partial<InventoryRule>) => void; onRemoveInventoryRule: (index: number) => void }) {
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
              <label>Số ký tự MST:</label>
              <input type="number" min={1} max={10} value={prefixMstDigits} onChange={(event) => onPrefixMstDigitsChange?.(parseInt(event.currentTarget.value) || 3)} />
            </div>
            <div className="prefix-quick-actions">
              <button type="button" className={`prefix-apply-all-button ${activePrefixStrategy === 'last_2_words' ? 'active' : ''}`} disabled={busy || !companies.length} onClick={() => onApplyPrefixPresetToAll?.('last_2_words')}>Áp 2 từ</button>
              <button type="button" className={`prefix-apply-all-button ${activePrefixStrategy === 'last_3_mst' ? 'active' : ''}`} disabled={busy || !companies.length} onClick={() => onApplyPrefixPresetToAll?.('last_3_mst')}>Áp MST</button>
              <button type="button" className={`prefix-apply-all-button ${activePrefixStrategy === '2_words_mst' ? 'active' : ''}`} disabled={busy || !companies.length} onClick={() => onApplyPrefixPresetToAll?.('2_words_mst')}>Áp 2 từ + MST</button>
            </div>
          </>}
        </section>}
      </div>

      <ConfigModal 
        isOpen={showConfigModal} 
        onClose={() => setShowConfigModal(false)} 
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
              <button type="button" disabled={busy || !companies.length} onClick={() => onApplyChoices()}>Áp dụng lựa chọn công ty và hàng hóa</button>
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
    return <tr key={companyRowKey(company, index)} className={`big-select-row ${group.className === 'duplicate-section' ? 'duplicate-company-row' : ''} ${index === safeSelectedIndex ? 'selected-row' : ''}`} onClick={() => onCompanySelect(index)}><td onClick={(event) => event.stopPropagation()}><input type="checkbox" checked={pending} onChange={(event) => onCompanyChange(index, event.currentTarget.checked)} /></td><td>{company.company}</td><td>{company.mst}</td>{showPrefix && <td onClick={(event) => event.stopPropagation()}><input className="company-prefix-input" value={company.value || ''} onChange={(event) => onCompanyPrefixChange?.(index, event.currentTarget.value)} /></td>}<td>{selectedCount} / {company.all_products.length}</td></tr>;
  })}</>;
}

function InventoryPairEditor({ pairs, useDefault, defaultPairId, rules, busy, onAddPair, onPairChange, onRemovePair, onDefaultsChange, onAddRule, onRuleChange, onRemoveRule }: { pairs: InventoryPair[]; useDefault: boolean; defaultPairId: string; rules: InventoryRule[]; busy: boolean; onAddPair: () => void; onPairChange: (index: number, field: 'ma_kho' | 'tk_vat_tu', value: string) => void; onRemovePair: (index: number) => void; onDefaultsChange: (update: Partial<Pick<WorkflowState, 'useDefaultInventoryPair' | 'defaultInventoryPairId'>>) => void; onAddRule: () => void; onRuleChange: (index: number, update: Partial<InventoryRule>) => void; onRemoveRule: (index: number) => void }) {
  const [collapsed, setCollapsed] = useState(false);
  const overlapWarnings = inventoryRuleOverlapWarnings(rules);
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
  };
}

function companyRowKey(company: CompanyRow, fallbackIndex = 0): string {
  return company.safe_id || `${company.mst || 'no-mst'}-${company.company || 'company'}-${fallbackIndex}`;
}

function prefixMemoryKey(company: CompanyRow): string {
  return company.mst || company.safe_id || company.company;
}

function clampPrefixMstDigits(value: unknown): number {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return 3;
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

function rememberPrefixEdit(values: PrefixStrategyValues, strategy: PrefixPresetStrategy, company: CompanyRow, value: string, mstDigits: number): PrefixStrategyValues {
  const next = rememberManualPrefixValues(values, strategy, [company], mstDigits, { [prefixMemoryKey(company)]: value });
  return next;
}

function rememberManualPrefixValues(values: PrefixStrategyValues, strategy: PrefixPresetStrategy, rows: CompanyRow[], mstDigits: number, overrides: Record<string, string> = {}): PrefixStrategyValues {
  const next = {
    ...emptyPrefixStrategyValues(),
    ...values,
    [strategy]: { ...(values[strategy] ?? {}) },
  };
  rows.forEach((company) => {
    const key = prefixMemoryKey(company);
    const value = normalizePrefixValue(Object.prototype.hasOwnProperty.call(overrides, key) ? overrides[key] : company.value);
    const computed = normalizePrefixValue(computePresetPrefix(company, strategy, mstDigits));
    if (value === computed) {
      delete next[strategy][key];
    } else {
      next[strategy][key] = value;
    }
  });
  return next;
}

function seedLoadedPrefixValues(values: PrefixStrategyValues, strategy: PrefixPresetStrategy, rows: CompanyRow[], mstDigits: number): PrefixStrategyValues {
  const next = {
    ...emptyPrefixStrategyValues(),
    ...values,
    [strategy]: { ...(values[strategy] ?? {}) },
  };
  rows.forEach((company) => {
    const key = prefixMemoryKey(company);
    if (Object.prototype.hasOwnProperty.call(next[strategy], key)) return;
    const loaded = normalizePrefixValue(company.value);
    if (!loaded) return;
    const computed = normalizePrefixValue(computePresetPrefix(company, strategy, mstDigits));
    const defaultPrefix = normalizePrefixValue(company.default_prefix || company.prefix_strategies?.last_2_words || computePresetPrefix(company, 'last_2_words', mstDigits));
    if (loaded !== computed && loaded !== defaultPrefix) {
      next[strategy][key] = loaded;
    }
  });
  return next;
}

function applyPrefixStrategyRows(rows: CompanyRow[], strategy: PrefixPresetStrategy, mstDigits: number, values: PrefixStrategyValues, commit = false): CompanyRow[] {
  const strategyValues = values[strategy] ?? {};
  return rows.map((company) => {
    const key = prefixMemoryKey(company);
    const savedValue = Object.prototype.hasOwnProperty.call(strategyValues, key) ? strategyValues[key] : undefined;
    const value = savedValue ?? computePresetPrefix(company, strategy, mstDigits);
    return { ...company, value, ...(commit ? { committed_prefix: normalizePrefixValue(value) } : {}) };
  });
}

function scrollStageBodyToTop() {
  window.requestAnimationFrame(() => {
    const body = document.querySelector('.stage-body');
    if (body instanceof HTMLElement) body.scrollTo({ top: 0, left: 0 });
  });
}

function computeCustomPrefix(company: CompanyRow, option: { name?: string; formula: string; chars: number; mstDigits?: number }): string {
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

function computePresetPrefix(company: CompanyRow, strategy: PrefixPresetStrategy, mstDigits = 3): string {
  const digits = Math.max(1, Math.min(10, mstDigits));
  const mstSuffix = (company.mst || '').slice(-digits);
  const wordsPrefix = company.prefix_strategies?.last_2_words || computeCustomPrefix(company, { name: '2 words', formula: 'initials', chars: 2 });
  if (strategy === 'last_3_mst') return mstSuffix;
  if (strategy === '2_words_mst') return `${wordsPrefix}${mstSuffix}`;
  return wordsPrefix;
}

function normalizedPrefixStrategy(strategy: string): PrefixPresetStrategy {
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
  if (override) return sanitizeDisplayProductCode(override);
  const preview = sanitizeDisplayProductCode(previewCodes[productName] || '');
  const appliedPrefix = sanitizeDisplayProductCode(committedCompanyPrefix(company));
  if (includePrefix && appliedPrefix) {
    return `${appliedPrefix}.${preview}`;
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


function ProcessStage({
  title,
  detail,
  buttonLabel = 'Xuất file',
  disabled,
  onProcess,
}: {
  title: string;
  detail: string;
  buttonLabel?: string;
  disabled: boolean;
  onProcess: () => void;
}) {
  return (
    <div className="placeholder-stage">
      <h3>{title}</h3>
      <p>{detail}</p>
      <div className="process-actions">
        <button type="button" disabled={disabled} onClick={onProcess}>{buttonLabel}</button>
      </div>
    </div>
  );
}

function FastImportExportStage({
  processedPurchaseSavedName,
  processedSalesSavedName,
  processedPurchaseStats,
  processedSalesStats,
  busy,
  onProcessedPurchaseUpload,
  onProcessedSalesUpload,
  onDownload,
}: {
  processedPurchaseSavedName: string;
  processedSalesSavedName: string;
  processedPurchaseStats: ProcessedFileStats | null;
  processedSalesStats: ProcessedFileStats | null;
  busy: boolean;
  onProcessedPurchaseUpload: (file: File | undefined) => void;
  onProcessedSalesUpload: (file: File | undefined) => void;
  onDownload: () => void;
}) {
  const canExport = Boolean(processedPurchaseSavedName && processedSalesSavedName);
  const renderUpload = (
    label: string,
    ready: boolean,
    stats: ProcessedFileStats | null,
    onUpload: (file: File | undefined) => void,
  ) => (
    <label className={`inventory-file-status ${ready ? 'ready' : ''}`}>
      <input
        type="file"
        accept=".xls,.xlsx,.xlsm"
        disabled={busy}
        onChange={(event) => {
          onUpload(event.currentTarget.files?.[0]);
          event.currentTarget.value = '';
        }}
      />
      <span>{label}</span>
      <strong>{ready ? 'Đã có FDI đã xử lý' : 'Tải file FDI đã xử lý'}</strong>
      <small>{ready ? processedStatsSentence(stats) : 'Dùng khi muốn bỏ qua các stage trước và xuất FAST trực tiếp.'}</small>
    </label>
  );

  return (
    <div className="inventory-allocation-stage export-only-stage">
      <section className="inventory-allocation-card inventory-export-card">
        <div className="report-stage-heading">
          <span>Xuất FAST</span>
          <h3>Xuất workbook FAST 4 sheet</h3>
          <p>Workbook sẽ gồm Hoadonmuahang, Hoadonbanhang, DMvat_tu và DMkhachhang. Chỉ các cột có dữ liệu trong file mẫu FAST mới được điền.</p>
        </div>
        <div className="inventory-file-grid">
          {renderUpload('FDI mua vào đã xử lý', Boolean(processedPurchaseSavedName), processedPurchaseStats, onProcessedPurchaseUpload)}
          {renderUpload('FDI bán ra đã xử lý', Boolean(processedSalesSavedName), processedSalesStats, onProcessedSalesUpload)}
          <div className={`inventory-file-status ${canExport ? 'ready' : ''}`}>
            <span>Workbook xuất</span>
            <strong>{canExport ? 'Sẵn sàng xuất' : 'Cần đủ 2 file FDI'}</strong>
            <small>{canExport ? 'Một file .xls với các sheet FAST.' : 'Tải hoặc tạo cả FDI mua vào và FDI bán ra trước.'}</small>
          </div>
        </div>
        <div className="export-action-panel">
          <button type="button" disabled={busy || !canExport} onClick={onDownload}>Xuất workbook FAST</button>
        </div>
      </section>
    </div>
  );
}

function PlaceholderStage({ title, detail }: { title: string; detail: string }) {
  return <div className="placeholder-stage"><h3>{title}</h3><p>{detail}</p></div>;
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
  return label;
}

function formatInventoryJobStatus(job: InventoryAllocationJob | null) {
  return formatOperationStatus(inventoryJobProgress(job), job?.label || 'Đang phân bổ tồn kho...');
}

function LoadingStage({ title, detail, progress }: { title: string; detail: string; progress?: OperationProgress | null }) {
  const percent = Math.max(0, Math.min(100, Number(progress?.percent ?? 0)));
  const hasRealRowProgress = Boolean(progress && Number(progress.total || 0) > 1);
  const rowProgress = hasRealRowProgress ? progress : null;
  return (
    <div className="loading-stage">
      <div className="loading-spinner" aria-hidden="true" />
      <h3>{title}</h3>
      <p>{progress?.label || detail}</p>
      {rowProgress && (
        <div className="loading-progress-block">
          <div className="loading-progress-bar" aria-label="Tiến trình xử lý"><span style={{ width: `${percent}%` }} /></div>
          <strong>{percent}%</strong>
          <small>{formatCount(rowProgress.done)} / {formatCount(rowProgress.total)} dòng</small>
        </div>
      )}
    </div>
  );
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
        <PlaceholderStage title="Chua co nhom gia" detail="Bam Tao nhom gia de tao danh sach loc don gia tu cac cong ty/hang hoa da ap dung." />
      ) : (
        <div className="inner-scroll cao-price-scroll">
          <table className="cao-price-table">
            <thead>
              <tr><th>Ma VT</th><th>Hang hoa</th><th>Dong gia</th><th>Gia min/max</th><th>% loc</th><th>Nhom sau loc</th></tr>
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
      const codes = await loadCaoThanhPreviewCodes(rows, wordRules, firstWordRules, repeatedPhrases);
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
    setManualOverrides((current) => ({ ...current, [productKey(company.mst, productName)]: code.toUpperCase() }));
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
      const codes = await loadCaoThanhPreviewCodes(companies, wordRules, firstWordRules, repeatedPhrases);
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
    const nextGroups = buildCaoThanhPriceGroups(committed, previewCodes, manualOverrides, includeCompanyPrefix, priceRangeRules, priceAdjustAllPercent);
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
    const committedGroups = priceGroups.length ? priceGroups : buildCaoThanhPriceGroups(companies, previewCodes, manualOverrides, includeCompanyPrefix, priceRangeRules, priceAdjustAllPercent);
    setBusy(true);
    updateMessage('Đang xử lý file Cao Thành...');
    try {
      const rangeRules = caoThanhRangeRules(committedGroups);
      setPriceRangeRules((current) => ({ ...current, ...rangeRules }));
      const payload = buildCaoThanhProcessPayload(summary, columns, companies, manualOverrides, includeCompanyPrefix, wordRules, firstWordRules, repeatedPhrases, rangeRules, priceAdjustAllPercent);
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
          {companies.length ? <CompanyRulesStage companies={companies} selectedCompanyIndex={selectedCompanyIndex} productPreviewCodes={previewCodes} productCodeOverrides={manualOverrides} wordRules={wordRules} repeatedPhrases={repeatedPhrases} inventoryPairs={inventoryPairs} useDefaultInventoryPair={useDefaultInventoryPair} defaultInventoryPairId={defaultInventoryPairId} inventoryPairRules={inventoryPairRules} busy={busy} showCompanyPrefixControls includeCompanyPrefix={includeCompanyPrefix} prefixStrategy="last_2_words" prefixMstDigits={3} onIncludeCompanyPrefixChange={(include) => { setIncludeCompanyPrefix(include); setPriceGroups([]); }} onCompanyPrefixChange={updateCompanyPrefix} onPrefixMstDigitsChange={() => {}} onApplyPrefixPresetToAll={() => {}} onCompanySelect={setSelectedCompanyIndex} onCompanyChange={updateCompanyPending} onBulkCompanyChange={bulkUpdateCompanies} onProductChange={updateProduct} onProductCodeChange={updateProductCode} onApplyChoices={applyCompanyChoices} onRefreshPreviews={refreshPreviewCodes} onWordRuleChange={updateWordRule} onAddWordRule={addWordRule} onRepeatedChange={updateRepeated} onAddRepeated={addRepeated} onRemoveRepeated={removeRepeated} onAddInventoryPair={() => { const id = `pair-${Date.now()}`; setInventoryPairs([...inventoryPairs, { id, ma_kho: '', tk_vat_tu: '' }]); setDefaultInventoryPairId(defaultInventoryPairId || id); }} onInventoryPairChange={(index, field, value) => setInventoryPairs((rows) => rows.map((pair, rowIndex) => rowIndex === index ? { ...pair, [field]: value.toUpperCase() } : pair))} onRemoveInventoryPair={(index) => setInventoryPairs((rows) => rows.filter((_, rowIndex) => rowIndex !== index))} onInventoryDefaultsChange={(update) => { if ('useDefaultInventoryPair' in update) setUseDefaultInventoryPair(Boolean(update.useDefaultInventoryPair)); if ('defaultInventoryPairId' in update) setDefaultInventoryPairId(String(update.defaultInventoryPairId || '')); }} onAddInventoryRule={() => setInventoryPairRules((rows) => [...rows, { source_col: 'M', operator: 'contains', value: '', pair_id: defaultInventoryPairId || inventoryPairs[0]?.id || '', enabled: true, priority: 1 }])} onInventoryRuleChange={(index, update) => setInventoryPairRules((rows) => rows.map((rule, rowIndex) => rowIndex === index ? { ...rule, ...update } : rule))} onRemoveInventoryRule={(index) => setInventoryPairRules((rows) => rows.filter((_, rowIndex) => rowIndex !== index))} /> : <PreviewPanel summary={summary} />}
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
) {
  const products = Array.from(new Set(companies.flatMap((company) => company.all_products.map((product) => product.name)).filter(Boolean)));
  if (!products.length) return {};
  const result = await previewGenericProductCodes({
    profile: 'cao_thanh',
    products,
    word_rules: wordRules,
    first_word_rules: firstWordRules,
    repeated_phrase_removals: repeatedPhraseRemovals.filter((phrase) => phrase.trim()),
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
      const code = productDisplayCode(company, product.name, previewCodes, manualOverrides, includePrefix);
      if (!code) return;
      const rows = product.priceRows || [];
      rows.forEach((row, rowIndex) => {
        const price = numericValue(row.price);
        const quantity = numericValue(row.quantity);
        const amount = numericValue(row.amount) || price * quantity;
        if (price <= 0) return;
        const sourceRow: CaoThanhPriceSourceRow = {
          key: `${company.mst}|${product.name}|${row.excelRow ?? rowIndex}|${rowIndex}`,
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
    price_range_rules: priceRangeRules,
    price_adjust_all_percent: priceAdjustAllPercent,
    prefixes: companyPrefixes(companies),
    removed_companies: Object.fromEntries(companies.filter((company) => company.process === false).map((company) => [company.mst, true])),
    skipped_products_map: Object.fromEntries(companies.map((company) => {
      const selected = new Set(selectedProductNames(company));
      const skipped = company.all_products.map((product) => product.name).filter((name) => !selected.has(name));
      return [company.mst, skipped];
    }).filter(([, skipped]) => Array.isArray(skipped) && skipped.length)),
    all_mst: companies.map((company) => company.mst),
    process_mst: activeCompanies.map((company) => company.mst),
    mst_safe_id: companies.map((company, index) => `${company.mst}|||${index}`),
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

async function loadProductPreviewCodes(companies: CompanyRow[], wordRules: Record<string, string>, repeatedPhraseRemovals: string[], phase: 'purchase' | 'sales' = 'purchase') {
  const products = Array.from(new Set(companies.flatMap((company) => company.all_products.map((product) => product.name)).filter(Boolean)));
  if (!products.length) return {};
  const result = await previewVietmaxProductCodes(products, wordRules, repeatedPhraseRemovals, phase);
  return result.codes;
}

async function loadGenericProductPreviewCodes(profile: ProfileKey, companies: CompanyRow[], wordRules: Record<string, string>, firstWordRules: Record<string, string>, repeatedPhraseRemovals: string[]) {
  const products = Array.from(new Set(companies.flatMap((company) => company.all_products.map((product) => product.name)).filter(Boolean)));
  if (!products.length) return {};
  const result = await previewGenericProductCodes({ profile, products, word_rules: wordRules, first_word_rules: firstWordRules, repeated_phrase_removals: repeatedPhraseRemovals });
  return result.codes;
}

function buildPurchaseProcessPayload(workflow: WorkflowState) {
  const companies = workflow.companyRows;
  const activeCompanies = companies.filter((company) => company.process !== false);
  const activePrefixStrategy = normalizedPrefixStrategy(workflow.purchasePrefixStrategy);
  const purchasePrefixStrategyValues = rememberManualPrefixValues(workflow.purchasePrefixStrategyValues, activePrefixStrategy, companies, workflow.prefixMstDigits);
  const reviewScope = reviewScopeValue(workflow.purchaseReviewScope);
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
    prefix_strategy: activePrefixStrategy,
    prefix_mst_digits: workflow.prefixMstDigits,
    prefix_strategy_values: purchasePrefixStrategyValues,
    comparison_scope: workflow.comparisonScope,
    word_rules: workflow.purchaseWordRules,
    repeated_phrase_removals: workflow.purchaseRepeatedPhraseRemovals.filter((phrase) => phrase.trim()),
    manual_code_overrides: workflow.productCodeOverrides,
    vietmax_mua_vao_internal_merges: buildReviewRules(workflow.purchaseReviewRules, workflow.purchaseReviewRows, reviewScope),
    inventory_pairs: workflow.purchaseInventoryPairs.filter((pair) => pair.ma_kho.trim() || pair.tk_vat_tu.trim()),
    use_default_inventory_pair: workflow.purchaseUseDefaultInventoryPair,
    default_inventory_pair_id: workflow.purchaseDefaultInventoryPairId,
    inventory_pair_rules: inventoryRulesForPayload(workflow.purchaseInventoryPairRules, workflow.purchaseInventoryPairs),
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
  const activePrefixStrategy = normalizedPrefixStrategy(workflow.salesPrefixStrategy);
  const salesPrefixStrategyValues = rememberManualPrefixValues(workflow.salesPrefixStrategyValues, activePrefixStrategy, companies, workflow.prefixMstDigits);
  const reviewScope = reviewScopeValue(workflow.salesReviewScope);
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
    prefix_strategy: activePrefixStrategy,
    prefix_mst_digits: workflow.prefixMstDigits,
    prefix_strategy_values: salesPrefixStrategyValues,
    comparison_scope: workflow.comparisonScope,
    word_rules: workflow.salesWordRules,
    repeated_phrase_removals: workflow.salesRepeatedPhraseRemovals.filter((phrase) => phrase.trim()),
    manual_code_overrides: workflow.salesProductCodeOverrides,
    inventory_pairs: workflow.salesInventoryPairs.filter((pair) => pair.ma_kho.trim() || pair.tk_vat_tu.trim()),
    use_default_inventory_pair: workflow.salesUseDefaultInventoryPair,
    default_inventory_pair_id: workflow.salesDefaultInventoryPairId,
    inventory_pair_rules: inventoryRulesForPayload(workflow.salesInventoryPairRules, workflow.salesInventoryPairs),
    vietmax_processed_purchase_saved_name: workflow.processedPurchaseSavedName,
    vietmax_ban_ra_purchase_matches: workflow.matches.filter((match) => match.confirmed !== false),
    vietmax_ban_ra_purchase_match_rules: buildSalesMatchRules(workflow),
    vietmax_ban_ra_sales_internal_merges: buildReviewRules(workflow.salesReviewRules, workflow.salesReviewRows, reviewScope),
    prefixes: companyPrefixes(companies),
    all_mst: companies.map((company) => company.mst),
    process_mst: activeCompanies.map((company) => company.mst),
    mst_safe_id: companies.map((company, index) => `${company.mst}|||${index}`),
    ...companyPrefixFields(companies),
    ...Object.fromEntries(companies.flatMap((company, index) => (company.process === false ? [] : [[`selected_products_${index}`, selectedProductNames(company)]]))),
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
  const activeCompanies = companies.filter((company) => company.process !== false);
  const activePrefixStrategy = normalizedPrefixStrategy(workflow.purchasePrefixStrategy);
  const prefixStrategyValues = rememberManualPrefixValues(workflow.purchasePrefixStrategyValues, activePrefixStrategy, companies, workflow.prefixMstDigits);
  const columns = normalizeGenericColumns(workflow.genericColumns);
  const reviewScope = reviewScopeValue(workflow.purchaseReviewScope);
  const priceRangeRules = profile === 'cao_thanh'
    ? { ...workflow.priceRangeRules, ...caoThanhRangeRules(workflow.priceGroups) }
    : {};
  return {
    profile,
    ...columns,
    include_company_prefix: workflow.includeCompanyPrefix,
    prefix_strategy: activePrefixStrategy,
    prefix_mst_digits: workflow.prefixMstDigits,
    prefix_strategy_values: prefixStrategyValues,
    word_rules: workflow.wordRules,
    first_word_rules: workflow.firstWordRules,
    repeated_phrase_removals: workflow.repeatedPhraseRemovals.filter((phrase) => phrase.trim()),
    manual_code_overrides: workflow.productCodeOverrides,
    product_review_merges: buildReviewRules(workflow.purchaseReviewRules, workflow.purchaseReviewRows, reviewScope),
    price_range_rules: priceRangeRules,
    price_adjust_all_percent: profile === 'cao_thanh' ? workflow.priceAdjustAllPercent : 0,
    inventory_pairs: workflow.inventoryPairs.filter((pair) => pair.ma_kho.trim() || pair.tk_vat_tu.trim()),
    use_default_inventory_pair: workflow.useDefaultInventoryPair,
    default_inventory_pair_id: workflow.defaultInventoryPairId,
    inventory_pair_rules: inventoryRulesForPayload(workflow.inventoryPairRules, workflow.inventoryPairs),
    prefixes: companyPrefixes(companies),
    all_mst: companies.map((company) => company.mst),
    process_mst: activeCompanies.map((company) => company.mst),
    mst_safe_id: companies.map((company, index) => `${company.mst}|||${index}`),
    ...companyPrefixFields(companies),
    ...Object.fromEntries(companies.flatMap((company, index) => (company.process === false ? [] : [[`selected_products_${index}`, selectedProductNames(company)]]))),
    columns,
    removed_companies: Object.fromEntries(companies.filter((company) => company.process === false).map((company) => [company.mst, true])),
    skipped_products_map: Object.fromEntries(companies.map((company) => {
      const selected = new Set(selectedProductNames(company));
      const skipped = company.all_products.map((product) => product.name).filter((name) => !selected.has(name));
      return [company.mst, skipped];
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
    if (company.process === false) return [];
    const selected = new Set(selectedProductNames(company));
    return company.all_products.flatMap((product, productIndex) => {
      if (!selected.has(product.name)) return [];
      const firstPriceRow = product.priceRows?.[0];
      const key = productKey(company.mst, product.name);
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
          sales_company_key: company.mst,
        } as ReviewProduct];
      }
      return [{
        ...base,
        purchase_product: product.name,
        purchase_code: manualCode,
        purchase_unit: firstPriceRow?.unit ?? '',
        purchase_company: company.company,
        purchase_mst: company.mst,
        purchase_company_key: company.mst,
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
  if (isGenericProfileKey(profile)) return [buildGenericProcessPayload(workflow, profile)];
  const payloads = [];
  if ((phase === 'purchase' || phase === 'all') && workflow.companyRows.length) payloads.push(buildConfigPayload(workflow));
  if ((phase === 'sales' || phase === 'all') && (workflow.salesCompanyRows.length || workflow.salesFile || workflow.matches.length || workflow.salesMatchRules.length)) payloads.push(buildSalesConfigPayload(workflow));
  if (!payloads.length && phase === 'sales') return [buildSalesConfigPayload(workflow)];
  if (!payloads.length && phase === 'purchase') return [buildConfigPayload(workflow)];
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
    manual_code_overrides: workflow.productCodeOverrides,
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
    company_count: companies.length,
    processed_company_count: companies.filter((company) => company.process !== false).length,
    product_count: companies.reduce((total, company) => total + company.all_products.length, 0),
    selected_product_count: companies.reduce((total, company) => total + selectedProductNames(company).length, 0),
    include_company_prefix: workflow.includeCompanyPrefix,
    prefix_strategy: prefixStrategy,
    prefix_mst_digits: workflow.prefixMstDigits,
    prefix_strategy_values: prefixStrategyValues,
    word_rules: isSales ? workflow.salesWordRules : workflow.purchaseWordRules,
    repeated_phrase_removals: isSales ? workflow.salesRepeatedPhraseRemovals : workflow.purchaseRepeatedPhraseRemovals,
    manual_code_overrides: isSales ? workflow.salesProductCodeOverrides : workflow.productCodeOverrides,
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
  return Object.fromEntries(companies.filter((company) => normalizePrefixValue(company.value)).map((company) => [company.mst, normalizePrefixValue(company.value)]));
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
  if (window.showSaveFilePicker) {
    try {
      const lowerName = filename.toLowerCase();
      const pickerType: { description: string; accept: Record<string, string[]> } = lowerName.endsWith('.json')
        ? { description: 'JSON config', accept: { 'application/json': ['.json'] } }
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
