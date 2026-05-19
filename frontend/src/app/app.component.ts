import { Component } from '@angular/core';
import { CommonModule, KeyValuePipe } from '@angular/common';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { FormsModule } from '@angular/forms';

type ProfileKey = 'son_phuong' | 'cao_thanh' | 'quang_thinh';

type ConfigProfile = {
  prefixes: Record<string, string>;
  selected_products: Record<string, string[]>;
  removed_companies: Record<string, boolean>;
  word_rules: Record<string, string>;
  first_word_rules: Record<string, string>;
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

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule, KeyValuePipe],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  step = 1;
  error: string | null = null;
  isLoading = false;
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

  wordRules: Record<string, string> = {};
  firstWordRules: Record<string, string> = {};
  wordRuleRows: Array<{ word: string; output: string }> = [];
  firstWordRuleRows: Array<{ word: string; output: string }> = [];
  showWordRuleModal = false;
  includeCompanyPrefix = true;

  priceGroupRules: Record<string, any> = {};
  priceRangeRules: Record<string, any> = {};
  priceConflictRows: any[] = [];
  priceConflictGroups: any[] = [];
  showPriceGroupModal = false;
  manualCodeOverrides: Record<string, string> = {};
  misorderGroups: any[] = [];
  misorderCanonicalCodes: Record<string, string> = {};
  showMisorderModal = false;
  showAddressModal = false;
  currentAddressCompany: any = null;
  showSkippedModal = false;

  constructor(private http: HttpClient) {
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
    this.includeCompanyPrefix = profile.include_company_prefix !== false;
    this.priceGroupRules = { ...(profile.price_group_rules || {}) };
    this.priceRangeRules = { ...(profile.price_range_rules || {}) };
    this.manualCodeOverrides = { ...(profile.manual_code_overrides || {}) };
    this.outputPath = profile.output_path || '';
  }

  onProfileChange() {
    this.applyProfileColumnDefaults();
    this.applyProfileConfig();
    if (this.companies.length) {
      this.applySavedProfileToCompanies();
      this.verifyPrefixes();
      this.refreshPriceGroups();
      this.refreshMisorderGroups();
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
    return this.profiles.find(p => p.key === this.selectedProfile)?.label || '';
  }

  selectedProfileNote() {
    return this.profiles.find(p => p.key === this.selectedProfile)?.note || '';
  }

  wordRuleCount() {
    return Object.keys(this.wordRules || {}).length + Object.keys(this.firstWordRules || {}).length;
  }

  emptyProfileState(): ConfigProfile {
    return {
      prefixes: {},
      selected_products: {},
      removed_companies: {},
      word_rules: { 'đen': 'DEN', 'tôn': 'TON' },
      first_word_rules: {},
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
    const selectedProducts: Record<string, string[]> = { ...existing.selected_products };
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
          selectedProducts[company.mst] = Array.from(company.selected_products || []).map(String);
          delete removedCompanies[company.mst];
        } else {
          removedCompanies[company.mst] = true;
        }
      }
    }
    return {
      ...existing,
      prefixes,
      selected_products: this.deltaSelectedProducts(selectedProducts),
      removed_companies: removedCompanies,
      word_rules: { ...this.wordRules },
      first_word_rules: { ...this.firstWordRules },
      price_group_rules: { ...this.priceGroupRules },
      price_range_rules: { ...this.priceRangeRules },
      manual_code_overrides: { ...this.manualCodeOverrides },
      include_company_prefix: this.includeCompanyPrefix,
      output_path: this.outputPath || ''
    };
  }

  deltaSelectedProducts(selectedProducts: Record<string, string[]>) {
    const delta: Record<string, string[]> = {};
    for (const [mst, products] of Object.entries(selectedProducts || {})) {
      const company = this.companies.find(item => item.mst === mst);
      if (!company) {
        if (Array.isArray(products) && products.length) delta[mst] = products;
        continue;
      }
      const allProducts = (company.all_products || []).map((product: any) => product.name);
      const selected = Array.isArray(products) ? products : [];
      if (!this.sameStringSet(selected, allProducts)) delta[mst] = selected;
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
      son_phuong: { ...this.emptyProfileState(), ...(this.config?.profiles?.son_phuong || {}) },
      cao_thanh: { ...this.emptyProfileState(), ...(this.config?.profiles?.cao_thanh || {}) },
      quang_thinh: { ...this.emptyProfileState(), ...(this.config?.profiles?.quang_thinh || {}) }
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
    this.isLoading = true;
    this.http.post<AppConfig>('/api/config', this.configSnapshot()).subscribe({
      next: (cfg) => {
        this.config = cfg;
        this.applyProfileConfig();
        this.isLoading = false;
        if (showSuccess) this.showMessage('Đã lưu cấu hình cho ' + this.selectedProfileLabel() + '.');
      },
      error: (err) => this.fail(err, 'Không lưu được cấu hình.')
    });
  }

  deleteProfileCache() {
    if (!window.confirm('Xóa cấu hình đã lưu cho ' + this.selectedProfileLabel() + '?')) return;
    const next = this.configSnapshot();
    next.profiles[this.selectedProfile] = this.emptyProfileState();
    this.http.post<AppConfig>('/api/config', next).subscribe({
      next: (cfg) => {
        this.config = cfg;
        this.applyProfileConfig();
        for (const company of this.companies) company.process = true;
        this.verifyPrefixes();
        this.showMessage('Đã xóa cache cấu hình cho ' + this.selectedProfileLabel() + '.');
      },
      error: (err) => this.fail(err, 'Không xóa được cache cấu hình.')
    });
  }

  exportConfig() {
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
          include_company_prefix: profile.include_company_prefix,
          price_group_rules: this.legacyPriceRules(profile.price_group_rules),
          output_path: profile.output_path,
          format_rule: this.selectedProfile
        }
      }
    };
    this.downloadJson(exportData, `product-code-${this.selectedProfile}-config.json`);
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
    this.isLoading = true;
    this.http.post<AppConfig>(`/api/config/profile/${this.selectedProfile}`, data).subscribe({
      next: (cfg) => {
        this.config = cfg;
        this.applyProfileColumnDefaults();
        this.applyProfileConfig();
        if (this.companies.length) {
          this.applySavedProfileToCompanies();
          this.verifyPrefixes();
        }
        this.isLoading = false;
        this.showMessage('Đã nhập cấu hình cho ' + this.selectedProfileLabel() + '.');
      },
      error: (err) => this.fail(err, 'Không nhập được cấu hình.')
    });
  }

  onFileChange(event: any) {
    this.fileToUpload = event.target.files?.[0] || null;
  }

  uploadFile() {
    if (!this.fileToUpload) return;
    this.error = null;
    this.isLoading = true;
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
        this.verifyPrefixes();
        this.refreshPriceGroups();
        this.step = 3;
        this.isLoading = false;
      },
      error: (err) => this.fail(err, 'Không kiểm tra được đơn vị.')
    });
  }

  applySavedProfileToCompanies() {
    const profile: ConfigProfile = { ...this.emptyProfileState(), ...(this.config?.profiles?.[this.selectedProfile] || {}) };
    const prefixes = profile.prefixes || {};
    const selectedMap = profile.selected_products || {};
    const removedCompanies = profile.removed_companies || {};
    for (const c of this.companies) {
      if (prefixes[c.mst]) c.value = prefixes[c.mst];
      if (Array.isArray(selectedMap[c.mst])) c.selected_products = new Set<string>(selectedMap[c.mst]);
      c.process = !removedCompanies[c.mst];
    }
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
    const visibleCompanies = this.companies.filter(c => c.process);
    this.duplicateCompanies = visibleCompanies.filter(c => {
      const p = (c.value || '').trim().toUpperCase();
      return p && (prefixCount.get(p) || 0) > 1;
    }).sort(sorter);
    this.normalCompanies = visibleCompanies.filter(c => !this.duplicateCompanies.includes(c)).sort(normalSorter);
  }

  verifyPrefixes(updateList = true) {
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
    if (updateList) this.sortCompanies();
    this.refreshPriceGroups();
    this.refreshMisorderGroups();
  }

  openProductModal(company: any) {
    this.currentCompany = company;
    this.currentProductList = (company.all_products || []).map((p: any) => ({
      ...p,
      selected: company.selected_products.has(p.name),
      code: this.productCodeFor(company, p.name)
    }));
    this.showProductModal = true;
  }

  productBaseCode(company: any, productName: string) {
    return this.buildCodePreview(company, productName || '', true);
  }

  productCodeFor(company: any, productName: string) {
    const key = this.productKey(company.mst, productName || '');
    return this.manualCodeOverrides[key] || this.productBaseCode(company, productName);
  }

  productCodePreview(product: any) {
    if (!this.currentCompany) return '';
    const productName = product?.name || '';
    return product?.code || this.productCodeFor(this.currentCompany, productName);
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

  longCodeCount(company: any) {
    let count = 0;
    for (const product of company?.all_products || []) {
      if (!company.selected_products?.has(product.name)) continue;
      const key = this.productKey(company.mst, product.name || '');
      const code = this.manualCodeOverrides[key] || this.buildCodePreview(company, product.name || '', false);
      if ((code || '').length > 50) count++;
    }
    return count;
  }

  saveProductSelection() {
    if (this.currentCompany) {
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
    }
    this.closeProductModal();
    this.refreshPriceGroups();
    this.refreshMisorderGroups();
  }

  closeProductModal() {
    this.showProductModal = false;
    this.currentCompany = null;
    this.currentProductList = [];
  }

  openWordRuleModal() {
    this.wordRuleRows = this.sortedRuleRows(this.wordRules);
    this.firstWordRuleRows = this.sortedRuleRows(this.firstWordRules);
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
    try {
      this.wordRules = this.rowsToRules(this.wordRuleRows, 'Từ thứ 3 trở đi');
      this.firstWordRules = this.rowsToRules(this.firstWordRuleRows, 'Hai từ đầu tiên');
    } catch (error) {
      this.showMessage(error instanceof Error ? error.message : 'Quy tắc từ thay riêng bị trùng.');
      return;
    }
    this.showWordRuleModal = false;
    this.refreshPriceGroups();
    this.refreshMisorderGroups();
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
    this.priceConflictRows = this.buildPriceConflictRows();
    this.priceConflictGroups = this.groupPriceRows(this.priceConflictRows);
  }

  buildPriceConflictRows() {
    const rows: any[] = [];
    for (const company of this.companies) {
      if (!company.process) continue;
      for (const product of company.all_products || []) {
        if (!company.selected_products.has(product.name)) continue;
        const prices = (product.priceRows || []).map((r: any) => Number(r.price)).filter((n: number) => Number.isFinite(n));
        const unique = Array.from(new Set(prices));
        if (unique.length < 2) continue;
        const key = this.productKey(company.mst, product.name);
        const saved = this.priceGroupRules[key];
      const row: any = {
          key,
          company,
          product,
          code: this.buildCodePreview(company, product.name),
          count: unique.length,
          min: Math.min(...prices),
          max: Math.max(...prices),
          percent: Number(saved?.percent || 8),
          savedRule: saved || null
        };
        row.priceGroups = this.buildPriceGroupPreview(row);
        rows.push(row);
      }
    }
    return rows.sort((a, b) => a.company.company.localeCompare(b.company.company) || a.product.name.localeCompare(b.product.name));
  }

  groupPriceRows(rows: any[]) {
    const map = new Map<string, any>();
    for (const row of rows) {
      const key = row.company.mst;
      if (!map.has(key)) map.set(key, { company: row.company, rows: [] });
      map.get(key).rows.push(row);
    }
    return Array.from(map.values());
  }

  buildPriceGroupPreview(row: any) {
    const min = Number(row.savedRule?.min_price || row.min);
    const percent = Number(row.percent || 8);
    const step = min * percent / 100;
    if (!min || !step) return [];
    const rawMap = new Map<number, any[]>();
    for (const item of row.product.priceRows || []) {
      const price = Number(item.price);
      if (!Number.isFinite(price)) continue;
      const raw = Math.floor((price - min) / step) + 1;
      const group = Math.max(1, raw);
      if (!rawMap.has(group)) rawMap.set(group, []);
      rawMap.get(group)!.push(item);
    }
    const occupied = Array.from(rawMap.keys()).sort((a, b) => a - b);
    return occupied.map((raw, index) => {
      const items = rawMap.get(raw) || [];
      const example = items[0] || {};
      const from = min + (raw - 1) * step;
      const to = min + raw * step;
      return {
        suffix: `.${String(index + 1).padStart(3, '0')}`,
        rangeDisplay: `${this.formatPrice(from)} - < ${this.formatPrice(to)}`,
        exampleRowDisplay: example.excelRow || '',
        exampleNameDisplay: example.name || row.product.name,
        examplePriceDisplay: this.formatPrice(Number(example.price)),
        count: items.length
      };
    });
  }

  onPriceGroupPercentChange(row: any) {
    row.priceGroups = this.buildPriceGroupPreview(row);
  }

  openPriceGroupModal() {
    this.refreshPriceGroups();
    this.showPriceGroupModal = true;
  }

  applyPriceGroupRules() {
    const next: Record<string, any> = {};
    const ranges = { ...this.priceRangeRules };
    for (const row of this.priceConflictRows) {
      next[row.key] = {
        base_code: row.code,
        min_price: row.min,
        max_price: row.max,
        percent: Number(row.percent || 8)
      };
      const old = ranges[row.code];
      ranges[row.code] = {
        min_price: old ? Math.min(Number(old.min_price), row.min) : row.min,
        max_price: old ? Math.max(Number(old.max_price), row.max) : row.max,
        percent: Number(row.percent || 8)
      };
    }
    this.priceGroupRules = next;
    this.priceRangeRules = ranges;
    this.showPriceGroupModal = false;
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
    const words = this.codeWords(sourceName);
    let tail = '';
    if (this.selectedProfile === 'cao_thanh') {
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

  refreshMisorderGroups() {
    const groups: any[] = [];
    for (const company of this.companies) {
      if (!company.process) continue;
      const map = new Map<string, any[]>();
      for (const product of company.all_products || []) {
        if (!company.selected_products.has(product.name)) continue;
        const words = this.normalizeProductWords(product.name);
        if (words.length < 2) continue;
        const key = this.misorderKey(product.name);
        if (!map.has(key)) map.set(key, []);
        map.get(key)!.push({
          key: this.productKey(company.mst, product.name),
          product,
          code: this.manualCodeOverrides[this.productKey(company.mst, product.name)] || this.buildCodePreview(company, product.name),
          orderKey: this.productOrderKey(product.name)
        });
      }
      for (const [wordKey, items] of map.entries()) {
        const orderCount = new Set(items.map(item => item.orderKey)).size;
        if (items.length > 1 && orderCount > 1) {
          const groupKey = `${company.mst}|||${wordKey}`;
          groups.push({ key: groupKey, company, items });
          if (!this.misorderCanonicalCodes[groupKey]) this.misorderCanonicalCodes[groupKey] = items[0].code;
        }
      }
    }
    this.misorderGroups = groups.sort((a, b) => a.company.mst.localeCompare(b.company.mst));
  }

  openMisorderModal() {
    this.refreshMisorderGroups();
    this.showMisorderModal = true;
  }

  applyMisorderChoices() {
    const next = { ...this.manualCodeOverrides };
    for (const group of this.misorderGroups) {
      const selected = this.misorderCanonicalCodes[group.key];
      if (!selected) continue;
      for (const item of group.items) {
        next[item.key] = selected;
      }
    }
    this.manualCodeOverrides = next;
    this.showMisorderModal = false;
  }

  processFile() {
    this.error = null;
    this.isLoading = true;
    const payload: any = {
      ...this.basePayload(),
      output_path: this.outputPath,
      word_rules: this.wordRules,
      first_word_rules: this.firstWordRules,
      include_company_prefix: this.includeCompanyPrefix,
      price_group_rules: this.priceGroupRules,
      price_range_rules: this.priceRangeRules,
      manual_code_overrides: this.manualCodeOverrides,
      all_mst: [],
      mst_safe_id: [],
      process_mst: [],
      removed_companies: {},
      prefixes: {},
      selected_products_map: {}
    };
    const validationError = this.fillCompanyPayload(payload);
    if (validationError) {
      this.errorMessage = validationError;
      this.showErrorModal = true;
      this.isLoading = false;
      return;
    }
    this.http.post('/api/process', payload, { responseType: 'blob', observe: 'response' }).subscribe({
      next: (response) => {
        const blob = new Blob([response.body as Blob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = this.originalName.replace(/\.[^/.]+$/, '') + '_formatted.xlsx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
        this.isLoading = false;
        this.loadConfig();
      },
      error: (err) => this.fail(err, 'Không xử lý được file.')
    });
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
      const allProducts = (item.all_products || []).map((product: any) => product.name);
      if (!this.sameStringSet(selectedProducts, allProducts)) {
        payload.selected_products_map[item.mst] = selectedProducts;
      }
      payload[`selected_products_${item.safe_id}`] = selectedProducts;
    }
    return null;
  }

  skippedCompanies() {
    return this.companies.filter(company => !company.process);
  }

  openSkippedModal() {
    this.showSkippedModal = true;
  }

  restoreSkippedCompany(company: any) {
    company.process = true;
    this.verifyPrefixes();
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
    if (this.step > 1) this.step--;
    this.error = null;
  }

  trackByKey(_: number, item: any) {
    return item.key || item.mst || item.name || item.word;
  }
}
