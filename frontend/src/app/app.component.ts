import { Component } from '@angular/core';
import { CommonModule, KeyValuePipe } from '@angular/common';
import { HttpClient, HttpErrorResponse, HttpEventType } from '@angular/common/http';
import { FormsModule } from '@angular/forms';

type ProfileKey = 'son_phuong' | 'cao_thanh' | 'quang_thinh';
type SuspectSectionKey = 'near_phrase' | 'misorder';
type Stage3Phase = 'customize' | 'price';

type ConfigProfile = {
  prefixes: Record<string, string>;
  selected_products: Record<string, string[]>;
  removed_companies: Record<string, boolean>;
  word_rules: Record<string, string>;
  first_word_rules: Record<string, string>;
  repeated_phrase_removals: string[];
  price_group_rules: Record<string, unknown>;
  price_range_rules: Record<string, unknown>;
  manual_code_overrides: Record<string, string>;
  include_company_prefix: boolean;
  output_path: string;
};

type AppConfig = {
  selected_profile: ProfileKey;
  columns: Record<string, string>;
  profiles: Record<ProfileKey, ConfigProfile>;
};

type PriceDetailRow = {
  key: string;
  excelRow: string;
  stt: string;
  invoiceNo: string;
  invoiceDate: string;
  companyName: string;
  productName: string;
  unit: string;
  price: number;
  quantity: number;
  totalAmount: number;
  costAmount: number;
  deltaAmount: number;
  deltaTotal: number;
  deltaPercent: number;
};

type PriceBucket = {
  key: string;
  label: string;
  finalCode?: string;
  count: number;
  min: number;
  max: number;
  averagePrice: number;
  marginPercent: number;
  draftMarginPercent?: number;
  adjustedAverage: number;
  lossCount?: number;
  hasLoss?: boolean;
  quantity?: number;
  totalAmount?: number;
  rows: Array<{ key: string; excelRow: string; stt?: string; invoiceNo?: string; invoiceDate?: string; companyName: string; productName: string; unit?: string; price: number; quantity?: number; totalAmount?: number }>;
  details: PriceDetailRow[] | null;
};

type PrefixOption = {
  key: string;
  label: string;
  value: string;
  warning: string;
};

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule, KeyValuePipe],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  step = 1;
  stage3Phase: Stage3Phase = 'customize';
  error: string | null = null;
  isLoading = false;
  processingProgress: number | null = null;
  processingProgressLabel = '';
  showConfigOperationLoading = false;
  configOperationLabel = '';
  showErrorModal = false;
  errorMessage = '';

  profiles: Array<{ key: ProfileKey; label: string; note: string }> = [
    { key: 'son_phuong', label: 'Sơn Phương', note: 'Lấy chữ đầu của mỗi từ, giữ nguyên số/kích thước và các từ được cấu hình.' },
    { key: 'cao_thanh', label: 'Cao Thành', note: 'Giữ nguyên 2 từ đầu, phần còn lại lấy chữ đầu của mã hàng hóa.' },
    { key: 'quang_thinh', label: 'Quang Thịnh', note: 'Bỏ chữ Sơn và phần sau HÀNG KM, lấy 2 ký tự mỗi từ, ngăn bằng dấu chấm.' }
  ];
  selectedProfile: ProfileKey = 'son_phuong';
  config: AppConfig | null = null;

  fileToUpload: File | null = null;
  savedName = '';
  originalName = '';
  columns: any[] = [];
  preview: any[] = [];
  previewKeys: string[] = [];

  companyCol = 'F';
  mstCol = 'G';
  addressCol = 'H';
  productCol = 'N';
  qtyCol = 'P';
  priceCol = '';
  outputCol = 'M';
  outputPath = '';

  rowsToProcess = 0;
  companyCount = 0;
  companies: any[] = [];
  duplicateCompanies: any[] = [];
  normalCompanies: any[] = [];

  showProductModal = false;
  currentCompany: any = null;
  currentProductList: any[] = [];
  targetProductKey = '';
  blinkingProductKey = '';
  showPrefixModal = false;
  prefixMstSuffixLength = 3;
  productCodeGroups: any[] = [];
  productCodeSections: any[] = [];
  priceCodeSections: any[] = [];
  expandedProductCodeSections: Record<string, boolean> = {};
  expandedProductCodeGroups: Record<string, boolean> = {};
  stageActionLoading = false;
  stageActionLoadingLabel = '';

  wordRules: Record<string, string> = {};
  firstWordRules: Record<string, string> = {};
  repeatedPhraseRemovals: string[] = [];
  wordRuleRows: Array<{ word: string; output: string }> = [];
  firstWordRuleRows: Array<{ word: string; output: string }> = [];
  repeatedPhraseRemovalRows: string[] = [];
  showWordRuleModal = false;
  includeCompanyPrefix = true;

  priceGroupRules: Record<string, any> = {};
  priceRangeRules: Record<string, any> = {};
  priceConflictRows: any[] = [];
  priceConflictGroups: any[] = [];
  priceCodeSummary = { base: 0, single: 0, multi: 0, split: 0 };
  showPriceGroupModal = false;
  showPriceModalRows = false;
  expandedPriceBuckets: Record<string, boolean> = {};
  priceFilterAllPercent = 8;
  priceAdjustAllPercent = 0;
  priceAdjustAllPercentDirty = false;
  manualCodeOverrides: Record<string, string> = {};
  longCodeCounts: Record<string, number> = {};
  misorderGroups: any[] = [];
  misorderCanonicalCodes: Record<string, string> = {};
  nearPhraseGroups: any[] = [];
  nearPhraseChoices: Record<string, string> = {};
  showSuspectModal = false;
  activeSuspectSection: SuspectSectionKey = 'near_phrase';
  showAddressModal = false;
  currentAddressCompany: any = null;
  showSkippedModal = false;
  selectedProfileLabelText = '';
  selectedProfileNoteText = '';
  wordRuleCountValue = 0;
  skippedCompanyList: any[] = [];
  private configOperationTimer: ReturnType<typeof setTimeout> | null = null;
  private configOperationId = 0;
  private codePreviewCache = new Map<string, string>();
  private priceGroupsCacheValid = false;
  private processingProgressTimer: ReturnType<typeof setInterval> | null = null;
  private processingProgressClearTimer: ReturnType<typeof setTimeout> | null = null;
  private productBlinkClearTimer: ReturnType<typeof setTimeout> | null = null;
  private readonly companyPrefixStopWords = new Set([
    'CONG', 'TY', 'TNHH', 'TM', 'DV', 'CP', 'CO', 'LTD', 'MTV', 'XNK', 'XD',
    'TINH', 'PHO', 'TP', 'HUYEN', 'QUAN', 'THI', 'XA', 'PHUONG'
  ]);
  private readonly vietnamLocationPhrases = [
    'An Giang', 'Bà Rịa Vũng Tàu', 'Bắc Giang', 'Bắc Kạn', 'Bạc Liêu', 'Bắc Ninh', 'Bến Tre', 'Bình Dương', 'Bình Định', 'Bình Phước', 'Bình Thuận',
    'Cà Mau', 'Cao Bằng', 'Cần Thơ', 'Đà Nẵng', 'Đắk Lắk', 'Đắk Nông', 'Điện Biên', 'Đồng Nai', 'Đồng Tháp', 'Gia Lai', 'Hà Giang', 'Hà Nam',
    'Hà Nội', 'Hà Tĩnh', 'Hải Dương', 'Hải Phòng', 'Hậu Giang', 'Hòa Bình', 'Hồ Chí Minh', 'Hưng Yên', 'Khánh Hòa', 'Kiên Giang', 'Kon Tum',
    'Lai Châu', 'Lâm Đồng', 'Lạng Sơn', 'Lào Cai', 'Long An', 'Nam Định', 'Nghệ An', 'Ninh Bình', 'Ninh Thuận', 'Phú Thọ', 'Phú Yên', 'Quảng Bình',
    'Quảng Nam', 'Quảng Ngãi', 'Quảng Ninh', 'Quảng Trị', 'Sóc Trăng', 'Sơn La', 'Tây Ninh', 'Thái Bình', 'Thái Nguyên', 'Thanh Hóa',
    'Thành phố Huế', 'Thừa Thiên Huế', 'Huế', 'Tiền Giang', 'Trà Vinh', 'Tuyên Quang', 'Vĩnh Long', 'Vĩnh Phúc', 'Yên Bái'
  ].map(name => this.normalizeCompanyTokens(name));

  constructor(private http: HttpClient) {
    this.refreshUiDerivedState();
    this.loadConfig();
  }

  loadConfig() {
    this.http.get<any>('/api/config').subscribe({
      next: (cfg) => {
        this.config = cfg;
        const savedProfile = cfg.selected_profile || 'son_phuong';
        this.selectedProfile = ((savedProfile === 'quang_thinh_1' || savedProfile === 'quang_thinh_2') ? 'quang_thinh' : savedProfile) as ProfileKey;
        this.companyCol = cfg.columns?.company_col || this.companyCol;
        this.mstCol = cfg.columns?.mst_col || this.mstCol;
        this.addressCol = cfg.columns?.address_col || this.addressCol;
        this.productCol = cfg.columns?.product_col || this.productCol;
        this.qtyCol = cfg.columns?.qty_col || this.qtyCol;
        this.priceCol = cfg.columns?.price_col || this.priceCol;
        this.outputCol = cfg.columns?.output_col || this.outputCol;
        this.applyProfileColumnDefaults();
        this.applyProfileConfig();
      }
    });
  }

  applyProfileConfig() {
    const profile: ConfigProfile = { ...this.emptyProfileState(), ...(this.config?.profiles?.[this.selectedProfile] || {}) };
    this.wordRules = { ...(profile.word_rules || {}) };
    this.firstWordRules = { ...(profile.first_word_rules || {}) };
    this.repeatedPhraseRemovals = this.normalizePhraseList(profile.repeated_phrase_removals || []);
    this.includeCompanyPrefix = profile.include_company_prefix !== false;
    this.priceGroupRules = { ...(profile.price_group_rules || {}) };
    this.priceRangeRules = { ...(profile.price_range_rules || {}) };
    this.manualCodeOverrides = { ...(profile.manual_code_overrides || {}) };
    this.outputPath = profile.output_path || '';
    this.invalidateCodePreviewCache();
    this.clearPriceGroupsCache();
    this.refreshUiDerivedState();
  }

  async onProfileChange() {
    const operationId = this.beginConfigOperation('Đang tải cấu hình...');
    try {
      this.applyProfileColumnDefaults();
      this.applyProfileConfig();
      await this.yieldToBrowser();
      if (this.companies.length) {
        this.applySavedProfileToCompanies();
        this.invalidateCodePreviewCache();
        this.verifyPrefixes(true, false);
        await this.refreshDerivedCodeViewsChunked();
      }
    } finally {
      this.endConfigOperation(operationId);
    }
  }

  applyProfileColumnDefaults() {
    const usingCaoThanhDefaults =
      this.companyCol === 'I' &&
      this.mstCol === 'J' &&
      this.addressCol === 'K' &&
      this.outputCol === 'L' &&
      this.productCol === 'M' &&
      this.qtyCol === 'O' &&
      this.priceCol === 'P';

    if (this.selectedProfile === 'cao_thanh') {
      this.companyCol = 'I';
      this.mstCol = 'J';
      this.addressCol = 'K';
      this.outputCol = 'L';
      this.productCol = 'M';
      this.qtyCol = 'O';
      this.priceCol = 'P';
    } else if (usingCaoThanhDefaults) {
      this.companyCol = 'F';
      this.mstCol = 'G';
      this.addressCol = 'H';
      this.productCol = 'M';
      this.qtyCol = 'O';
      this.priceCol = '';
      this.outputCol = 'L';
    } else {
      this.companyCol = this.companyCol || 'F';
      this.mstCol = this.mstCol || 'G';
      this.addressCol = this.addressCol || 'H';
      this.productCol = 'M';
      this.qtyCol = 'O';
      this.priceCol = '';
      this.outputCol = 'L';
    }
  }

  selectedProfileLabel() {
    return this.selectedProfileLabelText;
  }

  selectedProfileNote() {
    return this.selectedProfileNoteText;
  }

  wordRuleCount() {
    return this.wordRuleCountValue;
  }

  emptyProfileState(profileKey: ProfileKey = this.selectedProfile): ConfigProfile {
    return {
      prefixes: {},
      selected_products: {},
      removed_companies: {},
      word_rules: { 'đen': 'DEN', 'tôn': 'TON' },
      first_word_rules: {},
      repeated_phrase_removals: profileKey === 'cao_thanh' ? ['inox'] : [],
      price_group_rules: {},
      price_range_rules: {},
      manual_code_overrides: {},
      include_company_prefix: true,
      output_path: ''
    };
  }

  currentProfileSnapshot(): ConfigProfile {
    const existing = { ...this.emptyProfileState(), ...(this.config?.profiles?.[this.selectedProfile] || {}) };
    const prefixes: Record<string, string> = { ...existing.prefixes };
    const skippedProducts: Record<string, string[]> = { ...existing.selected_products };
    const removedCompanies: Record<string, boolean> = { ...existing.removed_companies };
    if (this.companies.length) {
      for (const company of this.companies) {
        if (company.process) {
          const prefix = (company.value || '').trim().toUpperCase();
          const defaultPrefix = (company.default_prefix || '').trim().toUpperCase();
          if (prefix && prefix !== defaultPrefix) {
            prefixes[company.mst] = prefix;
          } else {
            delete prefixes[company.mst];
          }
          skippedProducts[company.mst] = this.companySkippedProducts(company);
          delete removedCompanies[company.mst];
        } else {
          removedCompanies[company.mst] = true;
        }
      }
    }
    return {
      ...existing,
      prefixes,
      selected_products: this.deltaSkippedProducts(skippedProducts),
      removed_companies: removedCompanies,
      word_rules: { ...this.wordRules },
      first_word_rules: { ...this.firstWordRules },
      repeated_phrase_removals: this.normalizePhraseList(this.repeatedPhraseRemovals),
      price_group_rules: { ...this.priceGroupRules },
      price_range_rules: { ...this.priceRangeRules },
      manual_code_overrides: { ...this.manualCodeOverrides },
      include_company_prefix: this.includeCompanyPrefix,
      output_path: this.outputPath || ''
    };
  }

  companySkippedProducts(company: any) {
    const allProducts = (company.all_products || []).map((product: any) => product.name);
    const selectedSet = new Set(Array.from(company.selected_products || []).map(String));
    return allProducts.filter((productName: string) => !selectedSet.has(productName));
  }

  deltaSkippedProducts(skippedProducts: Record<string, string[]>) {
    const delta: Record<string, string[]> = {};
    for (const [mst, products] of Object.entries(skippedProducts || {})) {
      const company = this.companies.find(item => item.mst === mst);
      if (!company) {
        if (Array.isArray(products) && products.length) delta[mst] = products;
        continue;
      }
      const allProducts = (company.all_products || []).map((product: any) => product.name);
      const skipped = Array.isArray(products) ? products : [];
      if (skipped.length && skipped.length < allProducts.length) delta[mst] = skipped;
    }
    return delta;
  }

  sameStringSet(left: string[], right: string[]) {
    if (left.length !== right.length) return false;
    const rightSet = new Set(right);
    return left.every(item => rightSet.has(item));
  }

  configSnapshot(): AppConfig {
    const profiles: Record<ProfileKey, ConfigProfile> = {
      son_phuong: { ...this.emptyProfileState('son_phuong'), ...(this.config?.profiles?.son_phuong || {}) },
      cao_thanh: { ...this.emptyProfileState('cao_thanh'), ...(this.config?.profiles?.cao_thanh || {}) },
      quang_thinh: { ...this.emptyProfileState('quang_thinh'), ...(this.config?.profiles?.quang_thinh || {}) }
    };
    profiles[this.selectedProfile] = this.currentProfileSnapshot();
    return {
      selected_profile: this.selectedProfile,
      columns: {
        ...(this.config?.columns || {}),
        company_col: this.companyCol,
        mst_col: this.mstCol,
        address_col: this.addressCol,
        product_col: this.productCol,
        qty_col: this.qtyCol,
        price_col: this.priceCol,
        output_col: this.outputCol
      },
      profiles
    };
  }

  saveProfileConfig(showSuccess = true) {
    const operationId = this.beginConfigOperation('Đang lưu cấu hình...');
    this.isLoading = true;
    this.http.post<AppConfig>('/api/config', this.configSnapshot()).subscribe({
      next: (cfg) => {
        this.config = cfg;
        this.applyProfileConfig();
        this.isLoading = false;
        this.endConfigOperation(operationId);
        if (showSuccess) this.showMessage('Đã lưu cấu hình cho ' + this.selectedProfileLabel() + '.');
      },
      error: (err) => {
        this.endConfigOperation(operationId);
        this.fail(err, 'Không lưu được cấu hình.');
      }
    });
  }

  deleteProfileCache() {
    if (!window.confirm('Xóa cấu hình đã lưu cho ' + this.selectedProfileLabel() + '?')) return;
    const operationId = this.beginConfigOperation('Đang xóa cache cấu hình...');
    this.isLoading = true;
    const next = this.configSnapshot();
    next.profiles[this.selectedProfile] = this.emptyProfileState();
    this.http.post<AppConfig>('/api/config', next).subscribe({
      next: async (cfg) => {
        this.config = cfg;
        this.applyProfileConfig();
        for (const company of this.companies) company.process = true;
        this.verifyPrefixes(true, false);
        await this.refreshDerivedCodeViewsChunked();
        this.isLoading = false;
        this.endConfigOperation(operationId);
        this.showMessage('Đã xóa cache cấu hình cho ' + this.selectedProfileLabel() + '.');
      },
      error: (err) => {
        this.endConfigOperation(operationId);
        this.fail(err, 'Không xóa được cache cấu hình.');
      }
    });
  }

  async exportConfig() {
    const operationId = this.beginConfigOperation('Đang xuất cấu hình...');
    await this.yieldToBrowser();
    const profile = this.currentProfileSnapshot();
    const exportData = {
      version: 6,
      active_profile: this.selectedProfile,
      profiles: {
        [this.selectedProfile]: {
          companies: profile.prefixes,
          products: this.legacyProducts(profile.manual_code_overrides),
          removed_companies: profile.removed_companies,
          excluded_products: {},
          formula_options: this.legacyFormulaOptions(profile.word_rules),
          first_word_rules: profile.first_word_rules,
          repeated_phrase_removals: profile.repeated_phrase_removals,
          include_company_prefix: profile.include_company_prefix,
          price_group_rules: this.legacyPriceRules(profile.price_group_rules),
          output_path: profile.output_path,
          format_rule: this.selectedProfile
        }
      }
    };
    this.downloadJson(exportData, `product-code-${this.selectedProfile}-config.json`);
    this.endConfigOperation(operationId);
  }

  legacyProducts(products: Record<string, string>) {
    const result: Record<string, string> = {};
    for (const [key, code] of Object.entries(products || {})) {
      result[key.replace('|||', '|')] = code;
    }
    return result;
  }

  legacyFormulaOptions(wordRules: Record<string, string>) {
    return {
      son_phuong_keep_words: this.selectedProfile === 'son_phuong' ? this.wordRuleList(wordRules) : [],
      cao_thanh_keep_words: this.selectedProfile === 'cao_thanh' ? this.wordRuleList(wordRules) : []
    };
  }

  wordRuleList(wordRules: Record<string, string>) {
    return Object.entries(wordRules || {}).map(([word, output]) => ({ word, output }));
  }

  legacyPriceRules(priceRules: Record<string, unknown>) {
    const result: Record<string, unknown> = {};
    for (const [key, rule] of Object.entries(priceRules || {})) {
      result[key.replace('|||', '|')] = rule;
    }
    return result;
  }

  downloadJson(data: unknown, filename: string) {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    window.URL.revokeObjectURL(url);
    link.remove();
  }

  onConfigImportChange(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(String(reader.result || '{}'));
        this.importConfig(parsed);
      } catch {
        this.showMessage('File cấu hình không phải JSON hợp lệ.');
      } finally {
        input.value = '';
      }
    };
    reader.readAsText(file, 'utf-8');
  }

  importConfig(data: unknown) {
    const operationId = this.beginConfigOperation('Đang nhập cấu hình...');
    this.isLoading = true;
    this.http.post<AppConfig>(`/api/config/profile/${this.selectedProfile}`, data).subscribe({
      next: async (cfg) => {
        this.config = cfg;
        this.applyProfileColumnDefaults();
        this.applyProfileConfig();
        if (this.companies.length) {
          this.applySavedProfileToCompanies();
          this.verifyPrefixes(true, false);
          await this.refreshDerivedCodeViewsChunked();
        }
        this.isLoading = false;
        this.endConfigOperation(operationId);
        this.showMessage('Đã nhập cấu hình cho ' + this.selectedProfileLabel() + '.');
      },
      error: (err) => {
        this.endConfigOperation(operationId);
        this.fail(err, 'Không nhập được cấu hình.');
      }
    });
  }

  onFileChange(event: any) {
    this.fileToUpload = event.target.files?.[0] || null;
  }

  uploadFile() {
    if (!this.fileToUpload) return;
    this.error = null;
    this.isLoading = true;
    this.clearPriceGroupsCache();
    const formData = new FormData();
    formData.append('file', this.fileToUpload);
    this.http.post<any>('/api/mapping', formData).subscribe({
      next: (res) => {
        this.savedName = res.saved_name;
        this.originalName = res.original_name;
        this.columns = res.columns;
        this.preview = res.preview || [];
        this.previewKeys = this.preview.length ? Object.keys(this.preview[0]) : [];
        this.step = 2;
        this.isLoading = false;
      },
      error: (err) => this.fail(err, 'Không tải được file.')
    });
  }

  checkCompanies() {
    this.error = null;
    this.isLoading = true;
    this.clearPriceGroupsCache();
    const payload = this.basePayload();
    this.http.post<any>('/api/check', payload).subscribe({
      next: (res) => {
        this.rowsToProcess = res.rows_to_process;
        this.companyCount = res.company_count;
        this.companies = (res.companies || []).map((c: any) => {
          const selected = new Set<string>(c.selected_product_names || (c.all_products || []).map((p: any) => p.name));
          return { ...c, process: true, selected_products: selected };
        });
        this.applySavedProfileToCompanies();
        this.invalidateCodePreviewCache();
        this.verifyPrefixes();
        this.step = 3;
        this.stage3Phase = 'customize';
        this.isLoading = false;
      },
      error: (err) => this.fail(err, 'Không kiểm tra được đơn vị.')
    });
  }

  applySavedProfileToCompanies() {
    const profile: ConfigProfile = { ...this.emptyProfileState(), ...(this.config?.profiles?.[this.selectedProfile] || {}) };
    const prefixes = profile.prefixes || {};
    const skippedMap = profile.selected_products || {};
    const removedCompanies = profile.removed_companies || {};
    for (const c of this.companies) {
      c.value = c.default_prefix || c.value || '';
      c.process = true;
      c.selected_products = new Set<string>((c.all_products || []).map((product: any) => product.name));
      if (prefixes[c.mst]) c.value = prefixes[c.mst];
      if (Array.isArray(skippedMap[c.mst])) {
        const skippedSet = new Set<string>(skippedMap[c.mst]);
        c.selected_products = new Set<string>((c.all_products || []).map((product: any) => product.name).filter((name: string) => !skippedSet.has(name)));
      }
      c.process = !removedCompanies[c.mst];
    }
    this.refreshUiDerivedState();
  }

  basePayload() {
    return {
      saved_name: this.savedName,
      original_name: this.originalName,
      profile: this.selectedProfile,
      company_col: this.companyCol,
      mst_col: this.mstCol,
      address_col: this.addressCol,
      product_col: this.productCol,
      qty_col: this.qtyCol,
      price_col: this.priceCol,
      output_col: this.outputCol
    };
  }

  sortCompanies() {
    const prefixCount = new Map<string, number>();
    if (this.includeCompanyPrefix) {
      for (const item of this.companies) {
        if (!item.process) continue;
        const p = (item.value || '').trim().toUpperCase();
        if (p) prefixCount.set(p, (prefixCount.get(p) || 0) + 1);
      }
    }
    const lastSix = (mst: string) => ((mst || '').replace(/\D/g, '').slice(-6)).padStart(6, '0');
    const sorter = (a: any, b: any) => {
      const pa = (a.value || '').trim().toUpperCase();
      const pb = (b.value || '').trim().toUpperCase();
      return pa.localeCompare(pb) || lastSix(a.mst).localeCompare(lastSix(b.mst)) || a.company.localeCompare(b.company);
    };
    const normalSorter = (a: any, b: any) => lastSix(a.mst).localeCompare(lastSix(b.mst)) || a.company.localeCompare(b.company);
    const processedCompanies = this.companies.filter(c => c.process);
    this.duplicateCompanies = processedCompanies.filter(c => {
      const p = (c.value || '').trim().toUpperCase();
      return p && (prefixCount.get(p) || 0) > 1;
    }).sort(sorter);
    this.normalCompanies = this.companies.filter(c => !this.duplicateCompanies.includes(c)).sort(normalSorter);
  }

  async runWithStageLoading(label: string, work: () => void | Promise<void>) {
    this.stageActionLoadingLabel = label;
    this.stageActionLoading = true;
    await this.yieldToBrowser();
    try {
      await work();
    } finally {
      await this.yieldToBrowser();
      this.stageActionLoading = false;
      this.stageActionLoadingLabel = '';
    }
  }

  onCompanyPrefixToggleChange() {
    this.runWithStageLoading('Đang cập nhật nhóm mã VT...', () => this.verifyPrefixes());
  }

  verifyPrefixes(updateList = true, refreshDerived = true) {
    this.invalidateCodePreviewCache();
    const prefixCount = new Map<string, number>();
    if (this.includeCompanyPrefix) {
      for (const item of this.companies) {
        if (!item.process) continue;
        const p = (item.value || '').trim().toUpperCase();
        if (p) prefixCount.set(p, (prefixCount.get(p) || 0) + 1);
      }
    }
    for (const item of this.companies) {
      if (!item.process) {
        item.status = 'Bỏ qua';
        item.needs_manual = false;
        continue;
      }
      if (!this.includeCompanyPrefix) {
        item.status = 'Không dùng tiền tố';
        item.needs_manual = false;
        continue;
      }
      const p = (item.value || '').trim().toUpperCase();
      item.value = p;
      if (!p) {
        item.status = 'Thiếu tiền tố';
        item.needs_manual = true;
      } else if (!/^[A-Z0-9]{1,20}$/.test(p)) {
        item.status = 'Chỉ dùng A-Z hoặc 0-9';
        item.needs_manual = true;
      } else if ((prefixCount.get(p) || 0) > 1) {
        item.status = `Trùng tiền tố ${p}`;
        item.needs_manual = true;
      } else {
        item.status = 'Hợp lệ';
        item.needs_manual = false;
      }
    }
    if (updateList) {
      this.sortCompanies();
      if (refreshDerived) this.refreshDerivedCodeViews();
    }
    this.refreshUiDerivedState();
  }

  openPrefixModal() {
    this.verifyPrefixes(true, false);
    this.refreshUiDerivedState();
    this.showPrefixModal = true;
  }

  closePrefixModal() {
    this.showPrefixModal = false;
  }

  onPrefixSuffixLengthChange() {
    this.prefixMstSuffixLength = Math.max(3, Math.floor(Number(this.prefixMstSuffixLength || 3)));
    this.runWithStageLoading('Đang kiểm tra trùng MST...', () => this.verifyPrefixes(true, false));
  }

  verifyPrefixesWithLoading(label = 'Đang kiểm tra tiền tố...', updateList = true, refreshDerived = true) {
    this.runWithStageLoading(label, () => this.verifyPrefixes(updateList, refreshDerived));
  }

  applyPrefixOption(company: any, value: string) {
    company.value = this.normalizeCodeText(value || '');
    this.runWithStageLoading('Đang áp dụng prefix...', () => this.verifyPrefixes());
  }

  onManualPrefixChange() {
    this.runWithStageLoading('Đang cập nhật prefix...', () => this.verifyPrefixes());
  }

  prefixOptions(company: any): PrefixOption[] {
    const current = this.normalizeCodeText(company?.value || '');
    const mstSuffix = this.companyMstSuffix(company);
    const initials = this.companyNameInitials(company?.company || '');
    const initialsMst = initials && mstSuffix ? `${initials}${mstSuffix}` : '';
    const options = [
      { key: 'current', label: 'Đang chọn', value: current },
      { key: 'mst', label: `${this.prefixMstSuffixLength} số cuối MST`, value: mstSuffix },
      { key: 'initials', label: '2 chữ từ tên công ty', value: initials },
      { key: 'initials-mst', label: '2 chữ + MST', value: initialsMst }
    ];
    return options
      .filter(option => option.value)
      .map(option => ({ ...option, warning: this.prefixOptionWarning(company, option.value, option.key) }));
  }

  prefixOptionWarning(company: any, value: string, source: string) {
    const candidate = this.normalizeCodeText(value || '');
    if (!candidate) return '';
    const duplicatesSelected = this.companies.some(item => item !== company && item.process && this.normalizeCodeText(item.value || '') === candidate);
    if (duplicatesSelected) return `Sẽ trùng ${candidate} với công ty khác`;
    if (source === 'mst') {
      const suffixDuplicated = this.companies.some(item => item !== company && item.process && this.companyMstSuffix(item) === candidate);
      if (suffixDuplicated) return `Trùng ${this.prefixMstSuffixLength} số cuối MST, hãy thử ${this.prefixMstSuffixLength + 1} số, 2 chữ, 2 chữ + MST hoặc nhập tay.`;
    }
    return '';
  }

  companyMstSuffix(company: any) {
    const digits = String(company?.mst || '').replace(/\D/g, '');
    if (!digits) return '';
    return digits.slice(-Math.max(3, Number(this.prefixMstSuffixLength || 3)));
  }

  normalizeCompanyTokens(value: string) {
    return (value || '')
      .replace(/\u0110/g, 'D')
      .replace(/\u0111/g, 'd')
      .replace(/[ĐÄ]/g, 'D')
      .replace(/[đÄ‘]/g, 'd')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toUpperCase()
      .replace(/[^A-Z0-9]+/g, ' ')
      .trim()
      .split(/\s+/)
      .filter(Boolean);
  }

  companyNameInitials(companyName: string) {
    const tokens = this.removeLocationPhrases(this.normalizeCompanyTokens(companyName))
      .filter(token => !this.companyPrefixStopWords.has(token));
    const picked = tokens.slice(-2);
    return picked.map(token => token[0] || '').join('');
  }

  removeLocationPhrases(tokens: string[]) {
    const result: string[] = [];
    for (let index = 0; index < tokens.length;) {
      const phrase = this.vietnamLocationPhrases.find(items => items.length && items.every((item, offset) => tokens[index + offset] === item));
      if (phrase) {
        index += phrase.length;
      } else {
        result.push(tokens[index]);
        index++;
      }
    }
    return result;
  }

  goToPriceStage() {
    this.runWithStageLoading('Đang chuẩn bị lọc giá...', () => {
      this.verifyPrefixes();
      if (this.companies.some(company => company.process && company.needs_manual)) {
        this.showMessage('Còn tiền tố bị thiếu hoặc bị trùng. Hãy xử lý ở phần tùy chỉnh tiền tố trước khi lọc giá.');
        return;
      }
      this.stage3Phase = 'price';
      this.refreshPriceGroups();
      if (this.selectedProfile === 'cao_thanh') this.showPriceGroupModal = true;
    });
  }

  backToCustomizeStage() {
    this.stage3Phase = 'customize';
  }

  navigateBackFromStepThree() {
    if (this.stage3Phase === 'price') {
      this.backToCustomizeStage();
    } else {
      this.step--;
    }
  }

  refreshProductCodeGroups() {
    const grouped = new Map<string, any>();
    const sectioned = new Map<string, any>();
    for (const company of this.companies || []) {
      if (!company.process) continue;
      const sectionKey = this.includeCompanyPrefix ? `company|||${company.mst}` : 'all-products';
      if (!sectioned.has(sectionKey)) {
        sectioned.set(sectionKey, {
          key: sectionKey,
          title: this.includeCompanyPrefix ? `${company.mst} - ${company.company}` : 'Tất cả mã VT',
          company: this.includeCompanyPrefix ? company : null,
          productCount: 0,
          rowCount: 0,
          groups: new Map<string, any>()
        });
      }
      const section = sectioned.get(sectionKey);
      for (const product of company.all_products || []) {
        if (!company.selected_products?.has(product.name)) continue;
        const code = this.productCodeFor(company, product.name || '');
        const key = code || '(blank)';
        if (!grouped.has(key)) {
          grouped.set(key, {
            key: `code-group|||${key}`,
            code,
            products: [],
            companies: new Map<string, any>(),
            productCount: 0,
            rowCount: 0,
            hasLongCode: false
          });
        }
        const group = grouped.get(key);
        group.products.push({
          key: this.productKey(company.mst, product.name || ''),
          company,
          product,
          productName: product.name || '(Tên hàng hóa trống)',
          rowSummary: this.productRowSummary(product),
          priceSummary: this.productPriceSummary(product)
        });
        group.companies.set(company.mst, company);
        group.productCount++;
        group.rowCount += Number(product.count || 0);
        if ((this.productCodeFor(company, product.name || '', false) || '').length > 50) group.hasLongCode = true;
        if (!section.groups.has(key)) {
          section.groups.set(key, {
            key: `${sectionKey}|||code|||${key}`,
            code,
            products: [],
            companies: new Map<string, any>(),
            productCount: 0,
            rowCount: 0,
            hasLongCode: false
          });
        }
        const sectionGroup = section.groups.get(key);
        sectionGroup.products.push({
          key: this.productKey(company.mst, product.name || ''),
          company,
          product,
          productName: product.name || '(Tên hàng hóa trống)',
          rowSummary: this.productRowSummary(product),
          priceSummary: this.productPriceSummary(product)
        });
        sectionGroup.companies.set(company.mst, company);
        sectionGroup.productCount++;
        sectionGroup.rowCount += Number(product.count || 0);
        section.productCount++;
        section.rowCount += Number(product.count || 0);
        if ((this.productCodeFor(company, product.name || '', false) || '').length > 50) sectionGroup.hasLongCode = true;
      }
    }
    this.productCodeGroups = Array.from(grouped.values()).map(group => ({
      ...group,
      companyList: Array.from(group.companies.values()).sort((a: any, b: any) => String(a.mst).localeCompare(String(b.mst))),
      products: group.products.sort((a: any, b: any) => a.productName.localeCompare(b.productName, 'vi-VN'))
    })).sort((a, b) => (a.code || '').localeCompare(b.code || '', 'vi-VN'));
    this.productCodeSections = Array.from(sectioned.values()).map(section => ({
      ...section,
      codeCount: section.groups.size,
      expanded: this.productSectionExpanded(section.key),
      groups: Array.from(section.groups.values()).map((group: any) => ({
        ...group,
        expanded: this.productGroupExpanded(group.key),
        companyList: Array.from(group.companies.values()).sort((a: any, b: any) => String(a.mst).localeCompare(String(b.mst))),
        products: group.products.sort((a: any, b: any) => a.productName.localeCompare(b.productName, 'vi-VN'))
      })).sort((a: any, b: any) => (a.code || '').localeCompare(b.code || '', 'vi-VN'))
    })).sort((a, b) => a.title.localeCompare(b.title, 'vi-VN'));
  }

  productSectionExpanded(key: string) {
    return this.expandedProductCodeSections[key] !== false;
  }

  productGroupExpanded(key: string) {
    return this.expandedProductCodeGroups[key] !== false;
  }

  toggleProductCodeSection(section: any) {
    this.expandedProductCodeSections[section.key] = !this.productSectionExpanded(section.key);
    this.refreshProductCodeGroups();
    this.refreshPriceCodeSections();
  }

  toggleProductCodeGroup(group: any) {
    this.expandedProductCodeGroups[group.key] = !this.productGroupExpanded(group.key);
    this.refreshProductCodeGroups();
    this.refreshPriceCodeSections();
  }

  openProductModal(company: any, targetProductName = '') {
    if (this.productBlinkClearTimer) {
      clearTimeout(this.productBlinkClearTimer);
      this.productBlinkClearTimer = null;
    }
    this.currentCompany = company;
    this.currentProductList = (company.all_products || []).map((p: any) => ({
      ...p,
      key: this.productKey(company.mst, p.name || ''),
      selected: company.selected_products.has(p.name),
      code: this.productCodeFor(company, p.name),
      rowSummary: this.productRowSummary(p),
      priceSummary: this.productPriceSummary(p)
    }));
    this.targetProductKey = targetProductName ? this.productKey(company.mst, targetProductName) : '';
    this.blinkingProductKey = this.targetProductKey;
    this.showProductModal = true;
    if (this.targetProductKey) {
      setTimeout(() => this.scrollToTargetProduct(), 0);
      this.productBlinkClearTimer = setTimeout(() => {
        this.blinkingProductKey = '';
        this.productBlinkClearTimer = null;
      }, 1600);
    }
  }

  productModalRowId(product: any) {
    return `product-row-${encodeURIComponent(product?.key || '')}`.replace(/%/g, '-');
  }

  scrollToTargetProduct() {
    const product = this.currentProductList.find(item => item.key === this.targetProductKey);
    if (!product) return;
    const element = document.getElementById(this.productModalRowId(product));
    element?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  productBaseCode(company: any, productName: string, trim = true) {
    return this.cachedBuildCodePreview(company, productName || '', trim);
  }

  productCodeFor(company: any, productName: string, trim = true) {
    const key = this.productKey(company.mst, productName || '');
    return this.manualCodeOverrides[key] || this.productBaseCode(company, productName, trim);
  }

  refreshDerivedCodeViews() {
    if (this.selectedProfile === 'cao_thanh') {
      if (this.showPriceGroupModal || this.stage3Phase === 'price') this.refreshPriceGroups();
    } else {
      this.priceConflictRows = [];
      this.priceConflictGroups = [];
      this.priceCodeSections = [];
      this.misorderGroups = [];
      this.nearPhraseGroups = [];
    }
    this.updateLongCodeCounts();
    if (this.showProductModal && this.currentCompany) {
      for (const product of this.currentProductList) {
        product.code = this.productCodeFor(this.currentCompany, product.name || '');
      }
    }
    this.refreshProductCodeGroups();
    if (this.showSuspectModal) this.refreshActiveSuspectSection();
  }

  async refreshDerivedCodeViewsChunked() {
    if (this.selectedProfile === 'cao_thanh') {
      if (this.showPriceGroupModal || this.stage3Phase === 'price') {
        this.refreshPriceGroups();
        await this.yieldToBrowser();
      }
    } else {
      this.priceConflictRows = [];
      this.priceConflictGroups = [];
      this.priceCodeSections = [];
      this.misorderGroups = [];
      this.nearPhraseGroups = [];
    }
    await this.updateLongCodeCountsChunked();
    if (this.showProductModal && this.currentCompany) {
      for (const product of this.currentProductList) {
        product.code = this.productCodeFor(this.currentCompany, product.name || '');
      }
    }
    this.refreshProductCodeGroups();
    if (this.showSuspectModal) this.refreshActiveSuspectSection();
  }

  productCodePreview(product: any) {
    if (!this.currentCompany) return '';
    const productName = product?.name || '';
    return product?.code || this.productCodeFor(this.currentCompany, productName);
  }

  productCodeLength(product: any) {
    return (this.productCodePreview(product) || '').length;
  }

  productCodeTooLong(product: any) {
    return this.productCodeLength(product) > 50;
  }

  productRowSummary(product: any) {
    const count = Number(product?.count || 0);
    if (!count) return '';
    return `${count} dòng`;
  }

  productPriceSummary(product: any) {
    if (product?.minPrice === null || product?.minPrice === undefined) return '';
    const minPrice = Number(product.minPrice);
    const maxPrice = Number(product.maxPrice);
    if (!Number.isFinite(minPrice) || !Number.isFinite(maxPrice)) return '';
    const range = minPrice === maxPrice
      ? this.formatPrice(minPrice)
      : `${this.formatPrice(minPrice)} → ${this.formatPrice(maxPrice)}`;
    const priceCount = Number(product?.priceCount || 0);
    return priceCount > 1 ? `${range} (${priceCount} mức giá)` : range;
  }

  updateLongCodeCounts() {
    const next: Record<string, number> = {};
    for (const company of this.companies || []) {
      next[company.mst] = this.computeLongCodeCount(company);
    }
    this.longCodeCounts = next;
  }

  async updateLongCodeCountsChunked() {
    const next: Record<string, number> = {};
    let processed = 0;
    for (const company of this.companies || []) {
      next[company.mst] = 0;
      for (const product of company?.all_products || []) {
        if (!company.selected_products?.has(product.name)) continue;
        const code = this.productCodeFor(company, product.name || '', false);
        if ((code || '').length > 50) next[company.mst]++;
        processed++;
        if (processed % 150 === 0) await this.yieldToBrowser();
      }
    }
    this.longCodeCounts = next;
  }

  computeLongCodeCount(company: any) {
    let count = 0;
    for (const product of company?.all_products || []) {
      if (!company.selected_products?.has(product.name)) continue;
      const code = this.productCodeFor(company, product.name || '', false);
      if ((code || '').length > 50) count++;
    }
    return count;
  }

  saveProductSelection() {
    this.runWithStageLoading('Đang lưu hàng hóa...', () => {
      if (!this.currentCompany) return;
      this.currentCompany.selected_products = new Set<string>(
        this.currentProductList.filter(p => p.selected).map(p => p.name)
      );
      const nextOverrides = { ...this.manualCodeOverrides };
      for (const product of this.currentProductList) {
        const productName = product.name || '';
        const key = this.productKey(this.currentCompany.mst, productName);
        const baseCode = this.normalizeCodeText(this.productBaseCode(this.currentCompany, productName));
        const editedCode = this.normalizeCodeText(product.code || '');
        if (editedCode && editedCode !== baseCode) {
          nextOverrides[key] = editedCode;
        } else {
          delete nextOverrides[key];
        }
      }
      this.manualCodeOverrides = nextOverrides;
      this.invalidateCodePreviewCache();
      this.closeProductModal();
      this.refreshDerivedCodeViews();
    });
  }

  closeProductModal() {
    if (this.productBlinkClearTimer) {
      clearTimeout(this.productBlinkClearTimer);
      this.productBlinkClearTimer = null;
    }
    this.showProductModal = false;
    this.currentCompany = null;
    this.currentProductList = [];
    this.targetProductKey = '';
    this.blinkingProductKey = '';
  }

  openWordRuleModal() {
    this.wordRuleRows = this.sortedRuleRows(this.wordRules);
    this.firstWordRuleRows = this.sortedRuleRows(this.firstWordRules);
    this.repeatedPhraseRemovalRows = [...this.repeatedPhraseRemovals];
    if (!this.wordRuleRows.length) this.addWordRule();
    this.showWordRuleModal = true;
  }

  sortedRuleRows(rules: Record<string, string>) {
    return Object.keys(rules || {})
      .sort((a, b) => a.localeCompare(b, 'vi-VN'))
      .map(word => ({ word, output: rules[word] }));
  }

  addWordRule(target: 'rest' | 'first' = 'rest') {
    const rows = target === 'first' ? this.firstWordRuleRows : this.wordRuleRows;
    rows.push({ word: '', output: '' });
  }

  removeWordRule(index: number, target: 'rest' | 'first' = 'rest') {
    const rows = target === 'first' ? this.firstWordRuleRows : this.wordRuleRows;
    rows.splice(index, 1);
  }

  addRepeatedPhraseRemoval() {
    this.repeatedPhraseRemovalRows.push('');
  }

  removeRepeatedPhraseRemoval(index: number) {
    this.repeatedPhraseRemovalRows.splice(index, 1);
  }

  normalizePhraseList(values: string[]) {
    const result: string[] = [];
    const seen = new Set<string>();
    for (const value of values || []) {
      const phrase = String(value || '').trim();
      const key = this.normalizeRuleKey(phrase);
      if (!phrase || !key || seen.has(key)) continue;
      seen.add(key);
      result.push(phrase);
    }
    return result.sort((a, b) => a.localeCompare(b, 'vi-VN'));
  }

  rowsToRules(rows: Array<{ word: string; output: string }>, label: string) {
    const rules: Record<string, string> = {};
    const seen = new Map<string, string>();
    for (const row of rows) {
      const word = (row.word || '').trim();
      const output = this.normalizeCodeText((row.output || '').trim() || word);
      if (!word || !output) continue;
      const key = this.normalizeRuleKey(word);
      const existing = seen.get(key);
      if (existing && existing !== word) {
        throw new Error(`${label}: "${existing}" và "${word}" giống nhau sau khi bỏ dấu. Hãy giữ một dòng để tránh trùng quy tắc.`);
      }
      seen.set(key, word);
      rules[word] = output;
    }
    return Object.fromEntries(Object.entries(rules).sort(([a], [b]) => a.localeCompare(b, 'vi-VN')));
  }

  applyWordRules() {
    this.runWithStageLoading('Đang áp dụng quy tắc...', () => {
      try {
        this.wordRules = this.rowsToRules(this.wordRuleRows, 'Từ thứ 3 trở đi');
        this.firstWordRules = this.rowsToRules(this.firstWordRuleRows, 'Hai từ đầu tiên');
        this.repeatedPhraseRemovals = this.normalizePhraseList(this.repeatedPhraseRemovalRows);
      } catch (error) {
        this.showMessage(error instanceof Error ? error.message : 'Quy tắc từ thay riêng bị trùng.');
        return;
      }
      this.showWordRuleModal = false;
      this.invalidateCodePreviewCache();
      this.refreshDerivedCodeViews();
      this.refreshUiDerivedState();
    });
  }

  normalizeRuleKey(value: string) {
    return (value || '')
      .replace(/\u0110/g, 'D')
      .replace(/\u0111/g, 'd')
      .replace(/[ĐÄ]/g, 'D')
      .replace(/[đÄ‘]/g, 'd')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLocaleLowerCase('vi-VN')
      .trim()
      .replace(/\s+/g, ' ');
  }

  productKey(mst: string, product: string) {
    return `${mst}|||${product}`;
  }

  refreshPriceGroups() {
    if (this.priceGroupsCacheValid) {
      this.refreshPriceConflictViews();
      return;
    }
    this.priceConflictRows = this.buildPriceConflictRows();
    this.priceGroupsCacheValid = true;
    this.refreshPriceConflictViews();
  }

  rebuildPriceGroups() {
    this.priceGroupsCacheValid = false;
    this.refreshPriceGroups();
  }

  refreshPriceConflictViews() {
    this.priceConflictGroups = this.groupPriceRows(this.priceConflictRows);
    this.refreshPriceCodeSummary();
    this.refreshPriceCodeSections();
  }

  buildPriceConflictRows() {
    const grouped = new Map<string, any>();
    for (const company of this.companies) {
      if (!company.process) continue;
      for (const product of company.all_products || []) {
        if (!company.selected_products.has(product.name)) continue;
        const code = this.productCodeFor(company, product.name);
        if (!code) continue;
        const priceItems = (product.priceRows || [])
          .map((item: any, index: number) => ({ item, index, price: Number(item.price), quantity: this.validQuantity(item.quantity) }))
          .filter((item: any) => Number.isFinite(item.price));
        if (!priceItems.length) continue;
        const productKey = this.productKey(company.mst, product.name);
        const savedByProduct = this.priceGroupRules[productKey];
        let row = grouped.get(code);
        if (!row) {
          const savedByCode = this.priceRangeRules[code];
          row = {
            key: `price-code|||${code}`,
            code,
            companies: [],
            companyKeys: new Set<string>(),
            products: [],
            productKeySet: new Set<string>(),
            sourceRows: [],
            priceRowCount: 0,
            min: 0,
            max: 0,
            average: 0,
            filterPercent: Number(savedByProduct?.percent || savedByCode?.percent || 8),
            draftFilterPercent: Number(savedByProduct?.percent || savedByCode?.percent || 8),
            bulkAdjustPercent: 0,
            savedRule: savedByProduct || savedByCode || null,
            sectionBuckets: {},
            buckets: [] as PriceBucket[]
          };
          grouped.set(code, row);
        }
        if (!row.companyKeys.has(company.mst)) {
          row.companyKeys.add(company.mst);
          row.companies.push(company);
        }
        if (!row.productKeySet.has(productKey)) {
          row.productKeySet.add(productKey);
          row.products.push({ key: productKey, name: product.name, company });
        }
        for (const priceItem of priceItems) {
          row.sourceRows.push({
            key: `${productKey}|||${priceItem.item.excelRow || priceItem.index}`,
            company,
            product,
            productKey,
            excelRow: priceItem.item.excelRow || '',
            stt: priceItem.item.stt || '',
            invoiceNo: priceItem.item.invoiceNo || '',
            invoiceDate: priceItem.item.invoiceDate || '',
            name: priceItem.item.name || product.name,
            unit: priceItem.item.unit || '',
            price: priceItem.price,
            quantity: priceItem.quantity,
            totalAmount: this.lineTotalAmount(priceItem.price, priceItem.quantity, priceItem.item.amount)
          });
        }
      }
    }
    const rows: any[] = [];
    for (const row of grouped.values()) {
      const prices = row.sourceRows.map((item: any) => item.price);
      const unique = Array.from(new Set(prices));
      row.count = unique.length;
      row.hasMultiplePrices = unique.length > 1;
      row.priceRowCount = prices.length;
      row.min = Math.min(...prices);
      row.max = Math.max(...prices);
      row.quantity = this.totalQuantity(row.sourceRows);
      row.totalAmount = this.totalAmount(row.sourceRows);
      row.average = this.weightedAveragePrice(row.sourceRows);
      row.companyDisplay = row.companies.map((company: any) => `${company.mst} - ${company.company}`).join(' | ');
      row.productDisplay = row.products.map((product: any) => product.name).join(' | ');
      row.productCount = row.products.length;
      row.buckets = this.buildPriceBuckets(row);
      rows.push(row);
    }
    return rows.sort((a, b) => a.code.localeCompare(b.code));
  }

  groupPriceRows(rows: any[]) {
    const multiPriceRows = (rows || []).filter(row => row.hasMultiplePrices);
    return multiPriceRows.length ? [{ key: 'all-price-conflicts', title: 'Các Mã VT có nhiều đơn giá', rows: multiPriceRows }] : [];
  }

  stageProductCodeTotal() {
    this.ensurePriceCodeSummary();
    return this.stage3Phase === 'price' ? this.priceCodeSummary.base : this.productCodeGroups.length;
  }

  singlePriceCodeTotal() {
    this.ensurePriceCodeSummary();
    return this.priceCodeSummary.single;
  }

  multiPriceCodeTotal() {
    this.ensurePriceCodeSummary();
    return this.priceCodeSummary.multi;
  }

  splitPriceCodeTotal() {
    this.ensurePriceCodeSummary();
    return this.priceCodeSummary.split;
  }

  ensurePriceCodeSummary() {
    if (this.priceCodeSummary.base !== (this.priceConflictRows || []).length) {
      this.refreshPriceCodeSummary();
    }
  }

  refreshPriceCodeSummary() {
    const rows = this.priceConflictRows || [];
    this.priceCodeSummary = rows.reduce((summary, row) => {
      summary.base += 1;
      if (row.hasMultiplePrices) {
        summary.multi += 1;
      } else {
        summary.single += 1;
      }
      summary.split += Math.max(1, (row.buckets || []).length);
      return summary;
    }, { base: 0, single: 0, multi: 0, split: 0 });
  }

  validQuantity(value: unknown) {
    const quantity = Number(value);
    return Number.isFinite(quantity) && quantity > 0 ? quantity : 1;
  }

  totalQuantity(rows: Array<{ quantity?: number }>) {
    return (rows || []).reduce((sum, item) => sum + this.validQuantity(item.quantity), 0);
  }

  lineTotalAmount(price: unknown, quantity: unknown, totalAmount: unknown) {
    const parsedAmount = Number(totalAmount);
    if (Number.isFinite(parsedAmount) && parsedAmount > 0) return parsedAmount;
    const parsedPrice = Number(price || 0);
    return parsedPrice * this.validQuantity(quantity);
  }

  totalAmount(rows: Array<{ price?: number; quantity?: number; totalAmount?: number }>) {
    return (rows || []).reduce((sum, item) => {
      return sum + this.lineTotalAmount(item.price, item.quantity, item.totalAmount);
    }, 0);
  }

  weightedAveragePrice(rows: Array<{ price?: number; quantity?: number; totalAmount?: number }>) {
    const quantity = this.totalQuantity(rows);
    return quantity > 0 ? this.totalAmount(rows) / quantity : 0;
  }

  refreshPriceCodeSections() {
    const sectioned = new Map<string, any>();
    for (const row of this.priceConflictRows || []) {
      const companies = this.includeCompanyPrefix ? (row.companies || []) : [null];
      for (const company of companies) {
        const sectionKey = this.includeCompanyPrefix ? `company|||${company?.mst || ''}` : 'all-products';
        if (!sectioned.has(sectionKey)) {
          sectioned.set(sectionKey, {
            key: sectionKey,
            title: this.includeCompanyPrefix ? `${company?.mst || ''} - ${company?.company || ''}` : 'Tất cả mã VT',
            company: this.includeCompanyPrefix ? company : null,
            productCount: 0,
            rowCount: 0,
            singlePriceCodeCount: 0,
            multiPriceCodeCount: 0,
            groups: new Map<string, any>()
          });
        }
        const section = sectioned.get(sectionKey);
        const codeKey = row.code || '(blank)';
        const groupKey = `${sectionKey}|||code|||${codeKey}`;
        const sourceRows = this.includeCompanyPrefix && company
          ? (row.sourceRows || []).filter((sourceRow: any) => sourceRow.company?.mst === company.mst)
          : (row.sourceRows || []);
        const products = this.includeCompanyPrefix && company
          ? (row.products || []).filter((product: any) => product.company?.mst === company.mst)
          : (row.products || []);
        const buckets = this.priceSectionBuckets(row, sectionKey, sourceRows);
        const priceRowCount = sourceRows.length || Number(row.priceRowCount || 0);
        const productCount = products.length || Number(row.productCount || 0);
        const lossCount = buckets.reduce((sum: number, bucket: any) => sum + Number(bucket.lossCount || 0), 0);
        section.groups.set(codeKey, {
          key: groupKey,
          code: row.code,
          productCount,
          rowCount: priceRowCount,
          priceRowCount,
          hasMultiplePrices: row.hasMultiplePrices,
          bucketCount: buckets.length,
          average: this.weightedAveragePrice(sourceRows),
          quantity: this.totalQuantity(sourceRows),
          totalAmount: this.totalAmount(sourceRows),
          filterPercent: Number(row.filterPercent || 0),
          draftFilterPercent: Number(row.draftFilterPercent ?? row.filterPercent ?? 0),
          productDisplay: products.map((product: any) => product.name).join(' | ') || row.productDisplay,
          companyDisplay: this.includeCompanyPrefix ? '' : row.companyDisplay,
          buckets,
          hasLoss: lossCount > 0,
          lossCount,
          expanded: this.productGroupExpanded(groupKey)
        });
        section.productCount += productCount;
        section.rowCount += priceRowCount;
        if (row.hasMultiplePrices) {
          section.multiPriceCodeCount++;
        } else {
          section.singlePriceCodeCount++;
        }
      }
    }
    this.priceCodeSections = Array.from(sectioned.values()).map(section => {
      const groups = Array.from(section.groups.values()).sort((a: any, b: any) => (a.code || '').localeCompare(b.code || '', 'vi-VN'));
      return {
        ...section,
        codeCount: section.groups.size,
        expanded: this.productSectionExpanded(section.key),
        groups,
        priceSubsections: this.priceCodeSubsections(section.key, groups)
      };
    }).sort((a, b) => a.title.localeCompare(b.title, 'vi-VN'));
  }

  priceCodeSubsections(sectionKey: string, groups: any[]) {
    const singleGroups = groups.filter(group => !group.hasMultiplePrices);
    const multiGroups = groups.filter(group => group.hasMultiplePrices);
    return [
      this.priceCodeSubsection(sectionKey, 'multi', 'Mã VT nhiều đơn giá', 'Các mã cần kiểm tra/tách nhóm giá', multiGroups),
      this.priceCodeSubsection(sectionKey, 'single', 'Mã VT 1 đơn giá', 'Các mã chỉ có một đơn giá', singleGroups)
    ].filter(subsection => subsection.codeCount > 0);
  }

  priceCodeSubsection(sectionKey: string, type: string, title: string, description: string, groups: any[]) {
    const key = `${sectionKey}|||price-kind|||${type}`;
    return {
      key,
      type,
      title,
      description,
      codeCount: groups.length,
      rowCount: groups.reduce((sum, group) => sum + Number(group.priceRowCount || 0), 0),
      bucketCount: groups.reduce((sum, group) => sum + Number(group.bucketCount || 0), 0),
      groups,
      expanded: this.productSectionExpanded(key)
    };
  }

  priceSectionBuckets(row: any, sectionKey: string, sourceRows: any[]) {
    const allRows = row.sourceRows || [];
    if (sourceRows.length === allRows.length) return row.buckets || [];
    row.sectionBuckets = row.sectionBuckets || {};
    if (!row.sectionBuckets[sectionKey]) {
      const sectionRow = {
        ...row,
        key: `${row.key}|||${sectionKey}`,
        sourceRows
      };
      row.sectionBuckets[sectionKey] = this.buildPriceBuckets(sectionRow)
        .map((bucket: PriceBucket, index: number) => this.preserveBucketDraft(bucket, row.buckets || [], index));
    }
    return row.sectionBuckets[sectionKey];
  }

  clearPriceRowSectionBuckets(row: any) {
    row.sectionBuckets = {};
  }

  priceBaseline(average: number, percent: number) {
    if (!Number.isFinite(average) || !Number.isFinite(percent)) return 0;
    return average * (1 - percent / 100);
  }

  formatSignedPrice(value: number) {
    if (!Number.isFinite(value)) return '';
    if (Math.abs(value) < 0.000001) return this.formatPrice(0);
    const sign = value > 0 ? '+' : '-';
    return `${sign}${this.formatPrice(Math.abs(value))}`;
  }

  buildPriceBuckets(row: any) {
    const filterPercent = Number(row.filterPercent || 8);
    const sorted = [...(row.sourceRows || [])].sort((left: any, right: any) => left.price - right.price);
    const groupedRows: any[][] = [];
    for (const item of sorted) {
      const current = groupedRows[groupedRows.length - 1];
      if (!current?.length) {
        groupedRows.push([item]);
        continue;
      }
      const average = this.weightedAveragePrice(current);
      const deviation = average > 0 ? Math.abs((item.price - average) / average) * 100 : 0;
      if (deviation <= filterPercent) {
        current.push(item);
      } else {
        groupedRows.push([item]);
      }
    }
    return groupedRows.map((items, index) => {
      const key = `${row.key}|||bucket|||${index + 1}`;
      const averagePrice = this.weightedAveragePrice(items);
      const quantity = this.totalQuantity(items);
      const totalAmount = this.totalAmount(items);
      const savedGroup = this.savedPriceBucketRule(row, index, items);
      const marginPercent = Number(savedGroup?.adjust_percent || 0);
      const bucket: PriceBucket = {
        key,
        label: `Nhóm ${index + 1}`,
        finalCode: this.finalPriceBucketCode(row.code, index, groupedRows.length),
        count: items.length,
        min: Math.min(...items.map(item => item.price)),
        max: Math.max(...items.map(item => item.price)),
        averagePrice,
        marginPercent,
        draftMarginPercent: marginPercent,
        adjustedAverage: this.priceBaseline(averagePrice, marginPercent),
        quantity,
        totalAmount,
        rows: items.map(item => ({
          key: item.key,
          excelRow: String(item.excelRow || ''),
          stt: String(item.stt || ''),
          invoiceNo: String(item.invoiceNo || ''),
          invoiceDate: String(item.invoiceDate || ''),
          companyName: String(item.company.company || ''),
          productName: String(item.product.name || item.name || ''),
          unit: String(item.unit || ''),
          price: Number(item.price || 0),
          quantity: this.validQuantity(item.quantity),
          totalAmount: this.lineTotalAmount(item.price, item.quantity, item.totalAmount)
        })),
        details: null
      };
      this.refreshPriceBucketSummary(bucket, false);
      return bucket;
    });
  }

  refreshPriceBucketSummary(bucket: PriceBucket, clearDetails = true) {
    const marginPercent = Number(bucket.marginPercent || 0);
    bucket.adjustedAverage = this.priceBaseline(bucket.averagePrice, marginPercent);
    bucket.lossCount = this.computeBucketLossCount(bucket);
    bucket.hasLoss = bucket.lossCount > 0;
    if (clearDetails) bucket.details = null;
    return bucket;
  }

  preserveBucketDraft(bucket: PriceBucket, previousBuckets: PriceBucket[], index: number) {
    const previous = this.matchPreviousBucket(bucket, previousBuckets, index);
    if (!previous) return bucket;
    const marginPercent = Number(previous.draftMarginPercent ?? previous.marginPercent ?? bucket.marginPercent ?? 0);
    bucket.marginPercent = marginPercent;
    bucket.draftMarginPercent = marginPercent;
    return this.refreshPriceBucketSummary(bucket, false);
  }

  matchPreviousBucket(bucket: PriceBucket, previousBuckets: PriceBucket[], index: number) {
    return previousBuckets.find(previous => previous.min === bucket.min && previous.max === bucket.max)
      || previousBuckets[index]
      || null;
  }

  finalPriceBucketCode(code: string, index: number, bucketCount: number) {
    if (!code || bucketCount <= 1) return code || '';
    return `${code}.${String(index + 1).padStart(3, '0')}`;
  }

  savedPriceBucketRule(row: any, index: number, items: Array<{ price: number }>) {
    const groups = Array.isArray(row?.savedRule?.groups) ? row.savedRule.groups : [];
    const min = Math.min(...items.map(item => item.price));
    const max = Math.max(...items.map(item => item.price));
    return groups.find((group: any) => Number(group?.min_price) === min && Number(group?.max_price) === max)
      || groups.find((group: any) => Number(group?.index) === index + 1)
      || null;
  }

  priceBucketDetails(bucket: PriceBucket) {
    if (bucket.details) return bucket.details;
    bucket.details = bucket.rows
      .slice()
      .sort((left, right) => Number(left.excelRow || 0) - Number(right.excelRow || 0) || left.productName.localeCompare(right.productName, 'vi-VN'))
      .map(item => {
        const deltaAmount = item.price - bucket.adjustedAverage;
        const quantity = this.validQuantity(item.quantity);
        const costAmount = bucket.adjustedAverage * quantity;
        const deltaTotal = deltaAmount * quantity;
        const deltaPercent = item.price > 0 ? (deltaAmount / item.price) * 100 : 0;
        return {
          key: item.key,
          excelRow: String(item.excelRow || ''),
          stt: String(item.stt || ''),
          invoiceNo: String(item.invoiceNo || ''),
          invoiceDate: String(item.invoiceDate || ''),
          companyName: String(item.companyName || ''),
          productName: String(item.productName || ''),
          unit: String(item.unit || ''),
          price: Number(item.price || 0),
          quantity,
          totalAmount: this.lineTotalAmount(item.price, item.quantity, item.totalAmount),
          costAmount,
          deltaAmount,
          deltaTotal,
          deltaPercent
        };
      });
    return bucket.details;
  }

  togglePriceBucket(bucket: PriceBucket) {
    this.expandedPriceBuckets[bucket.key] = !this.expandedPriceBuckets[bucket.key];
    if (this.priceBucketExpanded(bucket)) this.priceBucketDetails(bucket);
  }

  priceBucketExpanded(bucket: PriceBucket) {
    return Boolean(this.expandedPriceBuckets[bucket.key]);
  }

  expandedPriceBucketDetails(bucket: PriceBucket) {
    return this.priceBucketExpanded(bucket) ? this.priceBucketDetails(bucket) : [];
  }

  bucketDeltaLabel(detail: PriceDetailRow) {
    if (Math.abs(detail.deltaPercent) < 0.000001) return 'Hòa vốn';
    return detail.deltaPercent > 0 ? 'Lãi' : 'Lỗ';
  }

  computeBucketLossCount(bucket: PriceBucket) {
    return (bucket.rows || []).filter(item => item.price < bucket.adjustedAverage).length;
  }

  bucketLossCount(bucket: PriceBucket) {
    return Number(bucket.lossCount ?? this.computeBucketLossCount(bucket));
  }

  bucketHasLoss(bucket: PriceBucket) {
    return Boolean(bucket.hasLoss ?? this.bucketLossCount(bucket) > 0);
  }

  groupLossCount(group: any) {
    return Number(group?.lossCount ?? (group?.buckets || []).reduce((sum: number, bucket: PriceBucket) => sum + this.bucketLossCount(bucket), 0));
  }

  groupHasLoss(group: any) {
    return Boolean(group?.hasLoss ?? this.groupLossCount(group) > 0);
  }

  onPriceGroupPercentChange(_row: any) {
  }

  onPriceBucketMarginChange(_bucket: PriceBucket) {
  }

  onPriceAdjustAllPercentChange() {
    this.priceAdjustAllPercentDirty = true;
  }

  applyPriceAdjustPercentToBuckets(buckets: PriceBucket[], percent: number) {
    const nextPercent = Number(percent || 0);
    for (const bucket of buckets || []) {
      bucket.marginPercent = nextPercent;
      bucket.draftMarginPercent = nextPercent;
      this.refreshPriceBucketSummary(bucket);
    }
  }

  applyPriceAdjustPercentToAll() {
    return this.runWithStageLoading('Đang áp % lãi...', () => {
      for (const row of this.priceConflictRows) {
        this.applyPriceAdjustPercentToBuckets(row.buckets || [], this.priceAdjustAllPercent);
        this.clearPriceRowSectionBuckets(row);
      }
      this.priceAdjustAllPercentDirty = false;
      this.refreshPriceConflictViews();
    });
  }

  applyPriceFilterPercentToAll() {
    return this.runWithStageLoading('Đang áp % lọc...', () => {
      const nextPercent = this.percentValue(this.priceFilterAllPercent, 8);
      for (const row of this.priceConflictRows) {
        row.draftFilterPercent = nextPercent;
        this.applyPriceRowFilterDraft(row);
      }
      this.refreshPriceConflictViews();
    });
  }

  applyPriceAdjustPercentToRow(row: any) {
    return this.runWithStageLoading('Đang áp % lãi trong mã VT...', () => {
      this.applyPriceAdjustPercentToBuckets(row.buckets || [], this.percentValue(row.bulkAdjustPercent, 0));
      this.clearPriceRowSectionBuckets(row);
      this.refreshPriceConflictViews();
    });
  }

  percentValue(value: unknown, fallback: number) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : fallback;
  }

  applyPriceRowFilterDraft(row: any) {
    const previousBuckets = row.buckets || [];
    row.filterPercent = this.percentValue(row.draftFilterPercent, row.filterPercent || 8);
    row.draftFilterPercent = row.filterPercent;
    row.buckets = this.buildPriceBuckets(row)
      .map((bucket: PriceBucket, index: number) => this.preserveBucketDraft(bucket, previousBuckets, index));
    this.clearPriceRowSectionBuckets(row);
  }

  applyPriceRowBucketDrafts(row: any) {
    for (const bucket of row.buckets || []) {
      bucket.marginPercent = this.percentValue(bucket.draftMarginPercent, bucket.marginPercent || 0);
      bucket.draftMarginPercent = bucket.marginPercent;
      this.refreshPriceBucketSummary(bucket);
    }
    this.clearPriceRowSectionBuckets(row);
  }

  applyPriceDraftsToRows() {
    if (this.priceAdjustAllPercentDirty) {
      for (const row of this.priceConflictRows) {
        this.applyPriceAdjustPercentToBuckets(row.buckets || [], this.priceAdjustAllPercent);
        this.clearPriceRowSectionBuckets(row);
      }
      this.priceAdjustAllPercentDirty = false;
    }
    for (const row of this.priceConflictRows) {
      const draftFilterPercent = this.percentValue(row.draftFilterPercent, row.filterPercent || 8);
      if (draftFilterPercent !== Number(row.filterPercent || 0)) {
        row.draftFilterPercent = draftFilterPercent;
        this.applyPriceRowFilterDraft(row);
      } else {
        this.applyPriceRowBucketDrafts(row);
      }
    }
    this.refreshPriceConflictViews();
  }

  openPriceGroupModal() {
    this.refreshPriceGroups();
    this.showPriceModalRows = false;
    this.showPriceGroupModal = true;
  }

  togglePriceModalRows() {
    this.showPriceModalRows = !this.showPriceModalRows;
  }

  applyPriceGroupRules() {
    return this.runWithStageLoading('Đang áp dụng lọc giá...', () => {
      this.applyPriceDraftsToRows();
      const next: Record<string, any> = {};
      const ranges = { ...this.priceRangeRules };
      for (const row of this.priceConflictRows) {
        const percent = Number(row.filterPercent || 8);
        const groups = (row.buckets || []).map((bucket: PriceBucket, index: number) => ({
          index: index + 1,
          label: bucket.label,
          min_price: bucket.min,
          max_price: bucket.max,
          average_price: bucket.averagePrice,
          adjust_percent: Number(bucket.marginPercent || 0)
        }));
        if (row.hasMultiplePrices !== false) {
          for (const product of row.products || []) {
            next[product.key] = {
              base_code: row.code,
              min_price: row.min,
              max_price: row.max,
              percent,
              groups
            };
          }
        }
        ranges[row.code] = {
          min_price: row.min,
          max_price: row.max,
          percent,
          groups
        };
      }
      this.priceGroupRules = next;
      this.priceRangeRules = ranges;
      this.priceGroupsCacheValid = true;
      this.refreshPriceConflictViews();
      this.showPriceModalRows = false;
      this.showPriceGroupModal = false;
    });
  }
  openAddressModal(company: any) {
    this.currentAddressCompany = company;
    this.showAddressModal = true;
  }

  closeAddressModal() {
    this.showAddressModal = false;
    this.currentAddressCompany = null;
  }

  formatPrice(value: number) {
    if (!Number.isFinite(value)) return '';
    return Math.abs(value - Math.round(value)) < 0.000001
      ? Math.round(value).toLocaleString('en-US')
      : value.toLocaleString('en-US', { maximumFractionDigits: 2 });
  }

  formatQuantity(value: number) {
    if (!Number.isFinite(value)) return '';
    return Math.abs(value - Math.round(value)) < 0.000001
      ? Math.round(value).toLocaleString('en-US')
      : value.toLocaleString('en-US', { maximumFractionDigits: 3 });
  }

  normalizeProductWords(name: string) {
    return (name || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toUpperCase()
      .replace(/[^A-Z0-9.]+/g, ' ')
      .trim()
      .split(/\s+/)
      .filter(Boolean);
  }

  misorderKey(name: string) {
    return this.normalizeProductWords(name).slice().sort().join('|');
  }

  productOrderKey(name: string) {
    return this.normalizeProductWords(name).join('|');
  }

  tokenPart(token: string, keepNumeric: boolean, defaultLen = 1, preserveUpperCode = false, keepSlash = false, rules = this.wordRules) {
    const key = this.normalizeRuleKey(token);
    const ruleKey = Object.keys(rules || {}).find(k => this.normalizeRuleKey(k) === key);
    if (ruleKey) return this.normalizeCodeText(rules[ruleKey]);
    const compact = this.normalizeCodeText(token, keepSlash);
    if (preserveUpperCode && this.isUpperCodeToken(token)) return compact;
    if (keepNumeric && /\d/.test(token)) return compact;
    return compact.slice(0, defaultLen);
  }

  phraseRulePart(words: string[], start: number, rules: Record<string, string>) {
    const ruleEntries = Object.keys(rules || {}).map(rule => ({ rule, words: this.codeWords(rule) }));
    const maxLength = Math.min(
      words.length - start,
      ruleEntries.reduce((max, entry) => Math.max(max, entry.words.length), 0)
    );
    for (let length = maxLength; length > 0; length--) {
      const phrase = words.slice(start, start + length).join(' ');
      const entry = ruleEntries.find(item => this.normalizeRuleKey(item.rule) === this.normalizeRuleKey(phrase));
      if (entry) return { part: this.normalizeCodeText(rules[entry.rule]), length };
    }
    return null;
  }

  tokenParts(words: string[], rules: Record<string, string>, fallback: (word: string) => string) {
    const parts: string[] = [];
    for (let i = 0; i < words.length;) {
      const matched = this.phraseRulePart(words, i, rules);
      if (matched) {
        parts.push(matched.part);
        i += matched.length;
      } else {
        parts.push(fallback(words[i]));
        i++;
      }
    }
    return parts;
  }

  removeRepeatedPhrases(words: string[], repeatedPhrases: string[]) {
    const phrases = this.normalizePhraseList(repeatedPhrases)
      .map(phrase => this.codeWords(phrase))
      .filter(items => items.length)
      .sort((a, b) => b.length - a.length);
    if (!phrases.length) return words;
    const seen = new Set<string>();
    const result: string[] = [];
    for (let i = 0; i < words.length;) {
      let match: string[] | null = null;
      for (const phrase of phrases) {
        if (phrase.length > words.length - i) continue;
        const current = words.slice(i, i + phrase.length).join(' ');
        if (this.normalizeRuleKey(current) === this.normalizeRuleKey(phrase.join(' '))) {
          match = phrase;
          break;
        }
      }
      if (match) {
        const currentWords = words.slice(i, i + match.length);
        const key = this.normalizeRuleKey(currentWords.join(' '));
        if (!seen.has(key)) {
          result.push(...currentWords);
          seen.add(key);
        }
        i += match.length;
      } else {
        result.push(words[i]);
        i++;
      }
    }
    return result;
  }

  normalizeCodeText(value: string, keepSlash = false) {
    const allowedPattern = keepSlash ? /[^A-Z0-9./]+/g : /[^A-Z0-9.]+/g;
    return (value || '')
      .replace(/\u0110/g, 'D')
      .replace(/\u0111/g, 'd')
      .replace(/[ĐÄ]/g, 'D')
      .replace(/[đÄ‘]/g, 'd')
      .replace(/[\u03a6\u03c6\u03d5\u00d8\u00f8\u2205\u2300\u0424\u0444\uff06]/g, 'F')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toUpperCase()
      .replace(allowedPattern, '');
  }

  isUpperCodeToken(value: string) {
    const raw = (value || '')
      .replace(/\u0110/g, 'D')
      .replace(/\u0111/g, 'd')
      .replace(/[ĐÄ]/g, 'D')
      .replace(/[đÄ‘]/g, 'd')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim();
    return /^[A-Z0-9./]+$/.test(raw) && /[A-Z]/.test(raw);
  }

  normalizeSep(value: string) {
    return (value || '').replace(/(?<=\d)\s*([xX*])\s*(?=\d)/g, 'x').trim();
  }

  codeWords(value: string) {
    return this.normalizeSep(value).split(/\s+/).filter(Boolean);
  }

  reducerDimensionToken(words: string[]) {
    const last = words[words.length - 1] || '';
    const normalized = this.normalizeCodeText(last, true);
    return /\d+(?:[.,]\d+)?\/\d+/.test(normalized) ? normalized : '';
  }

  isConReducerProduct(words: string[]) {
    return this.normalizeCodeText(words[0] || '') === 'CON' && Boolean(this.reducerDimensionToken(words));
  }

  normalizeCaoThanhConReducerWords(words: string[]) {
    if (!this.isConReducerProduct(words)) return words;
    const hasThu = words.some(word => this.normalizeCodeText(word) === 'THU');
    return hasThu ? words : [words[0], 'thu', ...words.slice(1)];
  }

  normalizeCaoThanhInoxGradeWords(words: string[]) {
    const normalized: string[] = [];
    for (let index = 0; index < words.length; index++) {
      const current = words[index];
      const currentToken = this.normalizeCodeText(current);
      const previousToken = this.normalizeCodeText(normalized[normalized.length - 1] || '');
      if (currentToken === '304' && previousToken === 'INOX') continue;
      normalized.push(current);
    }
    return normalized;
  }

  trimCode(value: string) {
    return (value || '').replace(/\.+/g, '.').replace(/^\.+|\.+$/g, '').slice(0, 50).replace(/\.+$/g, '');
  }

  isVolumeToken(word: string) {
    return /^\d+(?:[,.]\d+)?[lL]$/.test(word || '');
  }

  buildCodePreview(company: any, productName: string, trim = true) {
    const prefix = (company.value || '').trim().toUpperCase();
    const sourceName = this.selectedProfile === 'cao_thanh'
      ? (productName || '').replace(/\([^)]*\)/g, ' ')
      : (productName || '');
    let words = this.removeRepeatedPhrases(this.codeWords(sourceName), this.repeatedPhraseRemovals);
    let tail = '';
    if (this.selectedProfile === 'cao_thanh') {
      words = this.normalizeCaoThanhInoxGradeWords(words);
      words = this.normalizeCaoThanhConReducerWords(words);
      const first = this.tokenParts(
        words.slice(0, 2),
        this.firstWordRules,
        w => this.tokenPart(w, true, w.length, true, true, {})
      );
      const rest = this.tokenParts(
        words.slice(2),
        this.wordRules,
        w => this.tokenPart(w, true, 1, true, true, {})
      );
      tail = first.join('') + rest.join('');
    } else if (this.selectedProfile === 'quang_thinh') {
      const filtered: string[] = [];
      for (const w of words) {
        const n = this.normalizeCodeText(w);
        if (n === 'HANG') break;
        if (n === 'SON') continue;
        filtered.push(w);
      }
      tail = this.tokenParts(
        filtered,
        this.wordRules,
        w => this.isVolumeToken(w) ? this.normalizeCodeText(w) : this.tokenPart(w, true, 2, false, false, {})
      ).filter(Boolean).join('.');
    } else {
      const parts: string[] = [];
      for (let i = 0; i < words.length; i++) {
        const cur = this.normalizeCodeText(words[i]);
        const next = words[i + 1] ? this.normalizeCodeText(words[i + 1]) : '';
        if (cur === 'C' && /^\d+(?:MM)?$/.test(next)) {
          parts.push('Cx' + next.replace(/MM$/, ''));
          i++;
        } else if (cur === 'X') {
          parts.push('x');
        } else if (/\d/.test(words[i])) {
          parts.push(this.normalizeCodeText(words[i]));
        } else {
          const matched = this.phraseRulePart(words, i, this.wordRules);
          if (matched) {
            parts.push(matched.part);
            i += matched.length - 1;
          } else {
            parts.push(this.tokenPart(words[i], false, 1, false, false, {}));
          }
        }
      }
      tail = parts.join('');
    }
    const code = this.includeCompanyPrefix ? `${prefix}.${tail}` : tail;
    return trim ? this.trimCode(code) : code.replace(/\.+/g, '.').replace(/^\.+|\.+$/g, '');
  }

  ruleMatchScore(productName: string) {
    const sourceName = this.selectedProfile === 'cao_thanh'
      ? (productName || '').replace(/\([^)]*\)/g, ' ')
      : (productName || '');
    const words = this.removeRepeatedPhrases(this.codeWords(sourceName), this.repeatedPhraseRemovals);
    let score = 0;
    const firstMatch = this.phraseRulePart(words.slice(0, 2), 0, this.firstWordRules);
    if (firstMatch) score += 100 + firstMatch.length;
    for (let i = 2; i < words.length;) {
      const matched = this.phraseRulePart(words, i, this.wordRules);
      if (matched) {
        score += matched.length;
        i += matched.length;
      } else {
        i++;
      }
    }
    return score;
  }

  preferredMisorderCode(items: Array<{ code: string; ruleScore: number }>) {
    const sorted = [...items].sort((left, right) => {
      if (right.ruleScore !== left.ruleScore) return right.ruleScore - left.ruleScore;
      if (left.code.length !== right.code.length) return left.code.length - right.code.length;
      return left.code.localeCompare(right.code, 'vi-VN');
    });
    return sorted[0]?.code || '';
  }

  refreshMisorderGroups() {
    if (this.selectedProfile !== 'cao_thanh') {
      this.misorderGroups = [];
      return;
    }
    const groups: any[] = [];
    for (const company of this.companies) {
      if (!company.process) continue;
      const map = new Map<string, any[]>();
      for (const product of company.all_products || []) {
        if (!company.selected_products.has(product.name)) continue;
        const code = this.productCodeFor(company, product.name);
        const words = this.normalizeProductWords(product.name);
        if (words.length < 2) continue;
        const key = this.misorderKey(product.name);
        if (!map.has(key)) map.set(key, []);
        map.get(key)!.push({
          key: this.productKey(company.mst, product.name),
          product,
          code,
          ruleScore: this.ruleMatchScore(product.name || ''),
          orderKey: this.productOrderKey(product.name)
        });
      }
      for (const [wordKey, items] of map.entries()) {
        const orderCount = new Set(items.map(item => item.orderKey)).size;
        if (items.length > 1 && orderCount > 1) {
          const groupKey = `${company.mst}|||${wordKey}`;
          groups.push({ key: groupKey, company, items });
          const codes = new Set(items.map(item => item.code));
          if (!codes.has(this.misorderCanonicalCodes[groupKey])) this.misorderCanonicalCodes[groupKey] = this.preferredMisorderCode(items);
        }
      }
    }
    this.misorderGroups = groups.sort((a, b) => a.company.mst.localeCompare(b.company.mst));
  }

  openSuspectModal(section: SuspectSectionKey = 'near_phrase') {
    this.showSuspectModal = true;
    this.selectSuspectSection(section);
  }

  selectSuspectSection(section: SuspectSectionKey) {
    this.activeSuspectSection = section;
    this.refreshActiveSuspectSection();
  }

  refreshActiveSuspectSection() {
    if (this.activeSuspectSection === 'near_phrase') {
      this.refreshNearPhraseGroups();
    } else {
      this.refreshMisorderGroups();
    }
  }

  closeSuspectModal() {
    this.showSuspectModal = false;
  }

  openMisorderModal() {
    this.openSuspectModal('misorder');
  }

  applyMisorderChoices(groups = this.misorderGroups) {
    const next = { ...this.manualCodeOverrides };
    for (const group of groups) {
      const selected = this.misorderCanonicalCodes[group.key];
      if (!selected) continue;
      for (const item of group.items) {
        next[item.key] = selected;
      }
    }
    this.applyManualCodeOverrides(next);
  }

  applyManualCodeOverrides(next: Record<string, string>) {
    this.runWithStageLoading('Đang gộp mã tình nghi...', () => {
      this.manualCodeOverrides = next;
      this.closeSuspectModal();
      this.invalidateCodePreviewCache();
      this.refreshDerivedCodeViews();
      this.refreshUiDerivedState();
    });
  }
  phraseDistance(left: string, right: string) {
    if (Math.abs(left.length - right.length) > 1) return 2;
    const dp = Array.from({ length: left.length + 1 }, () => Array(right.length + 1).fill(0));
    for (let i = 0; i <= left.length; i++) dp[i][0] = i;
    for (let j = 0; j <= right.length; j++) dp[0][j] = j;
    for (let i = 1; i <= left.length; i++) {
      for (let j = 1; j <= right.length; j++) {
        const cost = left[i - 1] === right[j - 1] ? 0 : 1;
        dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
      }
    }
    return dp[left.length][right.length];
  }

  phraseCode(words: string[]) {
    return this.tokenParts(
      words,
      this.firstWordRules,
      word => this.tokenPart(word, true, word.length, true, true, {})
    ).join('');
  }

  refreshNearPhraseGroups() {
    const candidates = new Map<string, any>();
    if (this.selectedProfile !== 'cao_thanh') {
      this.nearPhraseGroups = [];
      return;
    }
    for (const company of this.companies) {
      if (!company.process) continue;
      for (const product of company.all_products || []) {
        if (!company.selected_products.has(product.name)) continue;
        const words = this.removeRepeatedPhrases(this.codeWords((product.name || '').replace(/\([^)]*\)/g, ' ')), this.repeatedPhraseRemovals);
        if (words.length < 2) continue;
        const phraseWords = words.slice(0, 2);
        const phrase = phraseWords.join(' ');
        const norm = this.normalizeCodeText(phrase);
        if (!norm || /^\d+$/.test(norm)) continue;
        if (!candidates.has(norm)) {
          candidates.set(norm, { phrase, norm, code: this.phraseCode(phraseWords), examples: [] });
        }
        const item = candidates.get(norm);
        if (item.examples.length < 3) item.examples.push(product.name);
      }
    }
    const items = Array.from(candidates.values()).sort((a, b) => a.norm.length - b.norm.length || a.phrase.localeCompare(b.phrase, 'vi-VN'));
    const groups: any[] = [];
    for (let i = 0; i < items.length; i++) {
      for (let j = i + 1; j < items.length; j++) {
        if (items[j].norm.length - items[i].norm.length > 1) break;
        if (this.phraseDistance(items[i].norm, items[j].norm) !== 1) continue;
        const key = `${items[i].norm}|||${items[j].norm}`;
        groups.push({ key, left: items[i], right: items[j] });
        const validChoices = new Set(['', items[i].code, items[j].code]);
        if (!validChoices.has(this.nearPhraseChoices[key])) this.nearPhraseChoices[key] = '';
      }
    }
    this.nearPhraseGroups = groups;
  }

  openNearPhraseModal() {
    this.openSuspectModal('near_phrase');
  }

  applyNearPhraseChoices() {
    const nextFirstRules = { ...this.firstWordRules };
    for (const group of this.nearPhraseGroups) {
      const code = this.nearPhraseChoices[group.key];
      if (!code) continue;
      nextFirstRules[group.left.phrase] = code;
      nextFirstRules[group.right.phrase] = code;
    }
    this.firstWordRules = this.rowsToRules(this.sortedRuleRows(nextFirstRules), 'Hai từ đầu tiên');
    this.closeSuspectModal();
    this.invalidateCodePreviewCache();
    this.refreshDerivedCodeViews();
    this.refreshUiDerivedState();
  }

  processFile() {
    this.error = null;
    this.isLoading = true;
    this.startProcessingProgress();
    const payload: any = {
      ...this.basePayload(),
      output_path: this.outputPath,
      word_rules: this.wordRules,
      first_word_rules: this.firstWordRules,
      repeated_phrase_removals: this.repeatedPhraseRemovals,
      include_company_prefix: this.includeCompanyPrefix,
      price_group_rules: this.priceGroupRules,
      price_range_rules: this.priceRangeRules,
      manual_code_overrides: this.manualCodeOverrides,
      all_mst: [],
      mst_safe_id: [],
      process_mst: [],
      removed_companies: {},
      prefixes: {},
      selected_products_map: {},
      skipped_products_map: {}
    };
    const validationError = this.fillCompanyPayload(payload);
    if (validationError) {
      this.errorMessage = validationError;
      this.showErrorModal = true;
      this.isLoading = false;
      this.clearProcessingProgress();
      return;
    }
    this.http.post('/api/process', payload, { responseType: 'blob', observe: 'events', reportProgress: true }).subscribe({
      next: (event) => {
        if (event.type === HttpEventType.Sent) {
          this.setProcessingProgress(5, 'Đang gửi dữ liệu...');
          return;
        }
        if (event.type === HttpEventType.UploadProgress) {
          this.setProcessingProgress(this.eventProgress(event.loaded, event.total, 5, 20), 'Đang gửi dữ liệu...');
          return;
        }
        if (event.type === HttpEventType.ResponseHeader) {
          this.setProcessingProgress(90, 'Đang nhận file kết quả...');
          return;
        }
        if (event.type === HttpEventType.DownloadProgress) {
          this.setProcessingProgress(this.eventProgress(event.loaded, event.total, 90, 99), 'Đang tải file kết quả...');
          return;
        }
        if (event.type !== HttpEventType.Response) return;

        const blob = new Blob([event.body as Blob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = this.originalName.replace(/\.[^/.]+$/, '') + '_formatted.xlsx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
        this.finishProcessingProgress();
        this.isLoading = false;
        this.loadConfig();
      },
      error: (err) => {
        this.clearProcessingProgress();
        this.fail(err, 'Không xử lý được file.');
      }
    });
  }

  startProcessingProgress() {
    this.clearProcessingProgress();
    this.setProcessingProgress(1, 'Đang chuẩn bị xử lý...');
    this.processingProgressTimer = setInterval(() => {
      const current = this.processingProgress ?? 1;
      if (current < 90) {
        this.setProcessingProgress(current + 1, 'Đang xử lý file...');
      } else if (current < 95) {
        this.setProcessingProgress(current + 0.25, 'Đang hoàn tất file...');
      }
    }, 500);
  }

  setProcessingProgress(value: number, label: string) {
    const next = Math.max(0, Math.min(100, Math.round(value)));
    this.processingProgress = Math.max(this.processingProgress ?? 0, next);
    this.processingProgressLabel = label;
  }

  eventProgress(loaded: number, total: number | undefined, start: number, end: number) {
    if (!total) return start;
    return start + (loaded / total) * (end - start);
  }

  finishProcessingProgress() {
    if (this.processingProgressTimer) {
      clearInterval(this.processingProgressTimer);
      this.processingProgressTimer = null;
    }
    this.setProcessingProgress(100, 'Hoàn tất 100%');
    if (this.processingProgressClearTimer) clearTimeout(this.processingProgressClearTimer);
    this.processingProgressClearTimer = setTimeout(() => this.clearProcessingProgress(), 900);
  }

  clearProcessingProgress() {
    if (this.processingProgressTimer) {
      clearInterval(this.processingProgressTimer);
      this.processingProgressTimer = null;
    }
    if (this.processingProgressClearTimer) {
      clearTimeout(this.processingProgressClearTimer);
      this.processingProgressClearTimer = null;
    }
    this.processingProgress = null;
    this.processingProgressLabel = '';
  }

  fillCompanyPayload(payload: any) {
    const used = new Map<string, string>();
    for (const item of this.companies) {
      payload.all_mst.push(item.mst);
      payload.mst_safe_id.push(`${item.mst}|||${item.safe_id}`);
      payload[`prefix_${item.safe_id}`] = item.value;
      if (!item.process) {
        payload.removed_companies[item.mst] = true;
        continue;
      }
      payload.process_mst.push(item.mst);
      const prefix = (item.value || '').trim().toUpperCase();
      if (this.includeCompanyPrefix) {
        if (!prefix) return `MST ${item.mst}: cần nhập tiền tố.`;
        if (!/^[A-Z0-9]{1,20}$/.test(prefix)) return `MST ${item.mst}: tiền tố chỉ gồm A-Z hoặc 0-9.`;
        if (used.has(prefix) && used.get(prefix) !== item.mst) return `MST ${item.mst}: tiền tố ${prefix} bị trùng.`;
        used.set(prefix, item.mst);
        const defaultPrefix = (item.default_prefix || '').trim().toUpperCase();
        if (prefix && prefix !== defaultPrefix) payload.prefixes[item.mst] = prefix;
      }
      const selectedProducts = Array.from(item.selected_products).map(String);
      const skippedProducts = this.companySkippedProducts(item);
      const allProducts = (item.all_products || []).map((product: any) => product.name);
      if (!this.sameStringSet(selectedProducts, allProducts)) {
        payload.selected_products_map[item.mst] = selectedProducts;
      }
      if (skippedProducts.length) {
        payload.skipped_products_map[item.mst] = skippedProducts;
      }
      payload[`selected_products_${item.safe_id}`] = selectedProducts;
    }
    return null;
  }

  skippedCompanies() {
    return this.skippedCompanyList;
  }

  openSkippedModal() {
    this.showSkippedModal = true;
  }

  restoreSkippedCompany(item: any) {
    this.runWithStageLoading('Đang khôi phục mục đã bỏ qua...', () => {
      if (item.kind === 'product') {
        item.company.selected_products.add(item.productName);
      } else {
        item.company.process = true;
      }
      this.verifyPrefixes();
    });
  }

  refreshUiDerivedState() {
    const profile = this.profiles.find(p => p.key === this.selectedProfile);
    this.selectedProfileLabelText = profile?.label || '';
    this.selectedProfileNoteText = profile?.note || '';
    this.wordRuleCountValue = Object.keys(this.wordRules || {}).length + Object.keys(this.firstWordRules || {}).length;
    const skippedItems: any[] = [];
    for (const company of this.companies) {
      if (!company.process) {
        skippedItems.push({
          key: `company|||${company.mst}`,
          kind: 'company',
          company,
          mst: company.mst,
          companyName: company.company,
          label: 'Đơn vị bị bỏ qua'
        });
        continue;
      }
      for (const product of company.all_products || []) {
        if (company.selected_products?.has(product.name)) continue;
        skippedItems.push({
          key: `product|||${company.mst}|||${product.name}`,
          kind: 'product',
          company,
          mst: company.mst,
          companyName: company.company,
          productName: product.name,
          label: 'Hàng hóa bị bỏ qua'
        });
      }
    }
    this.skippedCompanyList = skippedItems;
    this.refreshProductCodeGroups();
  }

  cachedBuildCodePreview(company: any, productName: string, trim = true) {
    const key = `${this.selectedProfile}|${this.includeCompanyPrefix ? '1' : '0'}|${trim ? '1' : '0'}|${company?.mst || ''}|${company?.value || ''}|${productName || ''}`;
    const cached = this.codePreviewCache.get(key);
    if (cached !== undefined) return cached;
    const code = this.buildCodePreview(company, productName || '', trim);
    this.codePreviewCache.set(key, code);
    return code;
  }

  invalidateCodePreviewCache() {
    this.codePreviewCache.clear();
    this.invalidatePriceGroupsCache();
  }

  invalidatePriceGroupsCache() {
    this.priceGroupsCacheValid = false;
  }

  clearPriceGroupsCache() {
    this.priceGroupsCacheValid = false;
    this.priceConflictRows = [];
    this.priceConflictGroups = [];
    this.priceCodeSections = [];
    this.priceCodeSummary = { base: 0, single: 0, multi: 0, split: 0 };
    this.expandedPriceBuckets = {};
  }

  beginConfigOperation(label: string) {
    const operationId = ++this.configOperationId;
    this.configOperationLabel = label;
    this.showConfigOperationLoading = false;
    if (this.configOperationTimer) clearTimeout(this.configOperationTimer);
    this.configOperationTimer = setTimeout(() => {
      if (this.configOperationId === operationId) this.showConfigOperationLoading = true;
    }, 2000);
    return operationId;
  }

  endConfigOperation(operationId: number) {
    if (operationId !== this.configOperationId) return;
    if (this.configOperationTimer) {
      clearTimeout(this.configOperationTimer);
      this.configOperationTimer = null;
    }
    this.showConfigOperationLoading = false;
    this.configOperationLabel = '';
  }

  yieldToBrowser() {
    return new Promise<void>(resolve => setTimeout(resolve, 0));
  }

  showMessage(message: string) {
    this.errorMessage = message;
    this.showErrorModal = true;
  }

  fail(err: unknown, fallback: string) {
    this.isLoading = false;
    if (this.isBackendConnectionFailure(err)) {
      this.errorMessage = this.backendUnavailableMessage();
      this.showErrorModal = true;
      return;
    }

    if (err instanceof HttpErrorResponse && err.error instanceof Blob) {
      err.error.text().then((text: string) => {
        try {
          this.errorMessage = this.extractBackendError(JSON.parse(text)) || fallback;
        } catch {
          this.errorMessage = fallback;
        }
        this.showErrorModal = true;
      });
      return;
    }

    this.errorMessage = this.extractHttpBackendError(err) || fallback;
    this.showErrorModal = true;
  }

  isBackendConnectionFailure(err: unknown) {
    if (!(err instanceof HttpErrorResponse)) return false;
    if (err.status === 0 || err.error instanceof ProgressEvent) return true;
    return [502, 503, 504].includes(err.status) && !this.extractHttpBackendError(err);
  }

  extractHttpBackendError(err: unknown) {
    if (!(err instanceof HttpErrorResponse)) return '';
    return this.extractBackendError(err.error);
  }

  extractBackendError(value: unknown) {
    if (!this.isRecord(value)) return '';
    const message = value['error'];
    return typeof message === 'string' ? message : '';
  }

  isRecord(value: unknown): value is Record<string, unknown> {
    return typeof value === 'object' && value !== null;
  }

  backendUnavailableMessage() {
    return 'Không kết nối được backend của ứng dụng. Hãy đóng cửa sổ này, mở lại ProductCodeFormatter.exe và đợi vài giây để ứng dụng khởi động xong.';
  }

  closeModal() {
    this.showErrorModal = false;
  }

  goBack() {
    if (this.step === 3) {
      this.navigateBackFromStepThree();
    } else if (this.step > 1) {
      this.step--;
    }
    this.error = null;
  }

  trackByKey(_: number, item: any) {
    return item.key || item.mst || item.name || item.word;
  }

  trackByIndex(index: number) {
    return index;
  }
}
