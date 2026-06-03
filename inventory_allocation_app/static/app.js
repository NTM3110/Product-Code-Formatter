const defaults = {
  purchase: { header_row: 2, data_start_row: 3, invoice_col: 'C', date_col: 'D', code_col: 'L', product_col: 'M', qty_col: 'O', price_col: 'P' },
  sales: { header_row: 2, data_start_row: 3, invoice_col: 'C', date_col: 'D', code_col: 'L', product_col: 'M', qty_col: 'O', price_col: 'P' },
  opening: { header_row: 1, data_start_row: 2, invoice_col: '', date_col: '', code_col: 'A', product_col: 'B', qty_col: 'C', price_col: 'D' }
};

const mappingVersion = 'invoice-l-m-o-p-date-v3';
const policyVersion = 'company-profile-son-phuong-barem-v3';
const defaultPolicy = {
  max_loss_percent: '10',
  max_profit_percent: '25',
  ignore_sale_suffix: false,
  allow_negative_export: true,
  company_profile: 'yen_thanh',
  son_phuong_split_counts: { pipe_box: 2, box: 2, pipe: 2 },
  generic_split_variance_percent: '',
  barem_tolerance_percent: '5',
  generic_min_take_quantity: '',
  generic_max_take_quantity: '',
  allow_future_purchase_reorder: false,
  future_purchase_window_days: '31'
};

const fieldLabels = {
  header_row: 'Dòng header',
  data_start_row: 'Dòng dữ liệu',
  invoice_col: 'Cột số HĐ',
  date_col: 'Cột ngày HĐ',
  code_col: 'Cột mã VT',
  product_col: 'Cột tên hàng',
  qty_col: 'Cột số lượng',
  price_col: 'Cột đơn giá'
};

let mapping = loadMapping();
let policy = loadPolicy();
const previewData = {};
let currentResult = null;
let activeLedgerWarehouse = 'KHH';
let activeInventorySummaryWarehouse = 'KHH';
let activeSalesSummaryWarehouse = 'KHH';
let activeSalesInvoiceWarehouse = 'KHH';
let progressTimer = null;
let progressValue = 0;
let isAnalyzing = false;
let currentWizardStep = 1;
const maxWizardStep = 4;
let defaultBaremRows = [];
let activeBaremTab = 'box|black';
let steelDetectionRows = [];
let steelDetectionUnknownRows = [];
let steelDetectionUnknownOnly = false;
let steelDetectionCompanyRules = [];
const baremTabs = [
  { key: 'box|black', kind: 'box', coating: 'black', label: 'Hộp đen' },
  { key: 'pipe|black', kind: 'pipe', coating: 'black', label: 'Ống đen' },
  { key: 'pipe|galvanized', kind: 'pipe', coating: 'galvanized', label: 'Ống mạ kẽm' },
  { key: 'box|galvanized', kind: 'box', coating: 'galvanized', label: 'Hộp mạ kẽm' }
];

function loadMapping() {
  try {
    if (localStorage.getItem('inventory_mapping_version') !== mappingVersion) {
      localStorage.setItem('inventory_mapping_version', mappingVersion);
      localStorage.removeItem('inventory_mapping');
      return structuredClone(defaults);
    }
    const stored = JSON.parse(localStorage.getItem('inventory_mapping') || '{}');
    return Object.fromEntries(
      Object.keys(defaults).map(kind => [kind, { ...defaults[kind], ...(stored[kind] || {}) }])
    );
  } catch (_) {
    return structuredClone(defaults);
  }
}

function loadPolicy() {
  try {
    if (localStorage.getItem('inventory_policy_version') !== policyVersion) {
      localStorage.setItem('inventory_policy_version', policyVersion);
      localStorage.setItem('inventory_policy', JSON.stringify(defaultPolicy));
      return { ...defaultPolicy };
    }
    return { ...defaultPolicy, ...JSON.parse(localStorage.getItem('inventory_policy') || '{}') };
  } catch (_) {
    return { ...defaultPolicy };
  }
}

function escapeHtml(value) {
  return String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
}

function numberLabel(value) {
  return Number(value || 0).toLocaleString('vi-VN', { maximumFractionDigits: 4 });
}

function signedNumberLabel(value) {
  const number = Number(value || 0);
  const label = Math.abs(number).toLocaleString('vi-VN', { maximumFractionDigits: 4 });
  if (number > 0) return `+${label}`;
  if (number < 0) return `-${label}`;
  return label;
}

function renderDefaultBaremModal() {
  const tab = baremTabs.find(item => item.key === activeBaremTab) || baremTabs[0];
  const rows = defaultBaremRows
    .filter(row => row.kind === tab.kind && row.coating === tab.coating)
    .sort((a, b) => String(a.dimension).localeCompare(String(b.dimension), 'vi', { numeric: true })
      || Number(a.thickness || 0) - Number(b.thickness || 0));
  document.getElementById('barem_modal_tabs').innerHTML = baremTabs.map(item => `
    <button class="warehouse-tab ${item.key === activeBaremTab ? 'active' : ''}" type="button" data-barem-tab="${escapeHtml(item.key)}">
      ${escapeHtml(item.label)}
    </button>
  `).join('');
  document.getElementById('barem_modal_rows').innerHTML = rows.map(row => `
    <tr>
      <td>${escapeHtml(row.dimension)}</td>
      <td class="num">${numberLabel(row.thickness)}</td>
      <td class="num">${numberLabel(row.weight)}</td>
    </tr>
  `).join('');
  document.getElementById('barem_modal_count').textContent = `${numberLabel(rows.length)} dòng trong bảng ${tab.label}.`;
}

async function showDefaultBaremModal() {
  const modal = document.getElementById('barem_modal');
  modal.classList.remove('hidden');
  if (!defaultBaremRows.length) {
    document.getElementById('barem_modal_rows').innerHTML = '<tr><td colspan="3">Đang tải bảng barem...</td></tr>';
    const response = await fetch(`/api/default-barem?t=${Date.now()}`, { cache: 'no-store' });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Không đọc được bảng barem mặc định.');
    defaultBaremRows = data.rows || [];
  }
  renderDefaultBaremModal();
}

function hideDefaultBaremModal() {
  document.getElementById('barem_modal').classList.add('hidden');
}

function marginPercentFromAmounts(saleAmount, costAmount) {
  const sale = Number(saleAmount || 0);
  if (!sale) return null;
  return (sale - Number(costAmount || 0)) / sale * 100;
}

function marginPercentLabelFromAmounts(saleAmount, costAmount, warehouseCode = '') {
  if (String(warehouseCode || '').toUpperCase() === 'KTP') return '';
  const percent = marginPercentFromAmounts(saleAmount, costAmount);
  return percent === null ? '' : `${signedNumberLabel(percent)}%`;
}

function moneyLabel(value) {
  return Number(value || 0).toLocaleString('vi-VN', { maximumFractionDigits: 0 });
}

function signedMoneyLabel(value) {
  const number = Number(value || 0);
  const label = Math.abs(number).toLocaleString('vi-VN', { maximumFractionDigits: 0 });
  if (number > 0) return `+${label}`;
  if (number < 0) return `-${label}`;
  return label;
}

function dateLabelFromIso(value) {
  if (!value) return '--';
  const [year, month, day] = String(value).split('-');
  if (!year || !month || !day) return escapeHtml(value);
  return `${day}/${month}/${year}`;
}

function compareIsoDate(value, boundary, operator) {
  if (!boundary || !value) return true;
  return operator === 'from' ? value >= boundary : value <= boundary;
}

function isoInRange(value, fromDate, toDate) {
  return compareIsoDate(value, fromDate, 'from') && compareIsoDate(value, toDate, 'to');
}

function isColumnField(field) {
  return field.endsWith('_col');
}

function columnOptions(kind, field) {
  const selected = String(mapping[kind][field] || '').toUpperCase();
  const columns = previewData[kind]?.columns || [];
  const allowEmpty = field === 'invoice_col' || field === 'date_col' || field === 'product_col' || field === 'price_col';
  const known = columns.some(column => column.letter === selected);
  const options = [];
  if (allowEmpty) options.push(`<option value="" ${selected === '' ? 'selected' : ''}>-- Không dùng --</option>`);
  if (selected && !known) {
    options.push(`<option value="${escapeHtml(selected)}" selected>${escapeHtml(selected)} - Chưa đọc tên cột</option>`);
  }
  if (!columns.length && selected && known === false) return options.join('');
  return options.concat(columns.map(column => {
    const label = column.header ? `${column.letter} - ${column.header}` : `${column.letter} - (trống)`;
    return `<option value="${escapeHtml(column.letter)}" ${column.letter === selected ? 'selected' : ''}>${escapeHtml(label)}</option>`;
  })).join('');
}

function fieldControl(kind, field) {
  if (isColumnField(field)) {
    return `<select id="${kind}_${field}" data-kind="${kind}" data-field="${field}">${columnOptions(kind, field)}</select>`;
  }
  return `<input id="${kind}_${field}" data-kind="${kind}" data-field="${field}"
    value="${escapeHtml(mapping[kind][field])}" type="number" min="1">`;
}

function renderMapping() {
  document.querySelectorAll('.mapping').forEach(panel => {
    const kind = panel.dataset.kind;
    const fields = panel.querySelector('.fields');
    fields.innerHTML = Object.keys(fieldLabels).map(field => `
      <div class="field">
        <label for="${kind}_${field}">${fieldLabels[field]}</label>
        ${fieldControl(kind, field)}
      </div>
    `).join('');
  });
  document.querySelectorAll('.field input, .field select').forEach(control => control.addEventListener('change', () => {
    saveMapping();
    if (control.dataset.field === 'header_row') previewFile(control.dataset.kind);
    if (control.dataset.kind === 'purchase') renderPurchaseClassificationPreview();
  }));
}

function saveMapping() {
  document.querySelectorAll('.field input, .field select').forEach(control => {
    const value = control.type === 'number' ? Number(control.value || 1) : control.value.trim().toUpperCase();
    mapping[control.dataset.kind][control.dataset.field] = value;
  });
  localStorage.setItem('inventory_mapping', JSON.stringify(mapping));
  localStorage.setItem('inventory_mapping_version', mappingVersion);
}

function setStatus(message, isError = false) {
  const status = document.getElementById('status');
  status.textContent = message;
  status.style.color = isError ? '#ad2c2c' : '';
}

function setUpdateStatus(message, isError = false) {
  const status = document.getElementById('update_status');
  status.textContent = message;
  status.style.color = isError ? '#ad2c2c' : '';
}

function setProcessingLocked(locked) {
  const companyProfile = document.getElementById('company_profile');
  if (companyProfile) companyProfile.disabled = locked;
  document.querySelector('.company-card')?.classList.toggle('locked', locked);
  ['purchase_file', 'sales_file', 'opening_file', 'barem_file'].forEach(id => {
    const input = document.getElementById(id);
    if (!input) return;
    input.disabled = locked;
    input.closest('.upload')?.classList.toggle('disabled', locked);
  });
  ['apply_policy', 'unlimit_policy', 'reset_mapping', 'show_default_barem'].forEach(id => {
    const button = document.getElementById(id);
    if (button) button.disabled = locked;
  });
  ['wizard_back', 'wizard_next'].forEach(id => {
    const button = document.getElementById(id);
    if (button) button.disabled = locked;
  });
}

function updateWizardStep(scroll = false) {
  document.querySelectorAll('[data-wizard-step]').forEach(panel => {
    panel.classList.toggle('hidden', Number(panel.dataset.wizardStep) !== currentWizardStep);
  });
  const steps = [...document.querySelectorAll('.step-indicator .step')];
  const lines = [...document.querySelectorAll('.step-indicator .line')];
  steps.forEach((step, index) => {
    const number = index + 1;
    step.classList.toggle('active', number === currentWizardStep);
    step.classList.toggle('completed', number < currentWizardStep);
  });
  lines.forEach((line, index) => {
    line.classList.toggle('active', index + 1 < currentWizardStep);
  });
  document.getElementById('wizard_back').classList.toggle('hidden', currentWizardStep === 1);
  document.getElementById('wizard_next').classList.toggle('hidden', currentWizardStep === maxWizardStep);
  document.getElementById('analyze').classList.toggle('hidden', currentWizardStep !== maxWizardStep);
  if (scroll) document.getElementById('workflow').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function canLeaveWizardStep(step) {
  if (step === 1 && !document.getElementById('purchase_file').files.length) {
    setStatus('Chọn file mua vào trước khi sang bước kiểm tra phân loại.', true);
    return false;
  }
  if (step === 4 && !document.getElementById('sales_file').files.length) {
    setStatus('Chọn file bán ra trước khi tính phân bổ.', true);
    return false;
  }
  return true;
}

function goWizardStep(delta) {
  if (delta > 0 && !canLeaveWizardStep(currentWizardStep)) return;
  currentWizardStep = Math.max(1, Math.min(maxWizardStep, currentWizardStep + delta));
  setStatus('');
  updateWizardStep(true);
  if (currentWizardStep === 2) renderPurchaseClassificationPreview();
}

async function loadAppVersion() {
  try {
    const response = await fetch(`/api/version?t=${Date.now()}`, { cache: 'no-store' });
    if (!response.ok) return;
    const data = await response.json();
    document.getElementById('app_version').textContent = data.version || '0.0';
  } catch (error) {
    // The app may be restarting during an update.
  }
}

function setUploadProgress(value, label) {
  const percent = Math.max(0, Math.min(100, Math.round(value)));
  document.getElementById('analyze_progress').classList.remove('hidden');
  document.getElementById('upload_percent').textContent = `${percent}%`;
  document.getElementById('upload_fill').style.width = `${percent}%`;
  if (label) document.getElementById('upload_label').textContent = label;
}

function setCalcProgress(value, label) {
  const percent = Math.max(0, Math.min(100, Math.round(value)));
  document.getElementById('analyze_progress').classList.remove('hidden');
  document.getElementById('calc_percent').textContent = `${percent}%`;
  document.getElementById('calc_fill').style.width = `${percent}%`;
  document.getElementById('calc_fill').classList.remove('indeterminate');
  if (label) document.getElementById('calc_label').textContent = label;
}

function setCalcWaiting(label) {
  document.getElementById('analyze_progress').classList.remove('hidden');
  document.getElementById('calc_percent').textContent = '...';
  document.getElementById('calc_fill').style.width = '100%';
  document.getElementById('calc_fill').classList.add('indeterminate');
  if (label) document.getElementById('calc_label').textContent = label;
}

function resetProgress() {
  clearInterval(progressTimer);
  progressTimer = null;
  progressValue = 0;
  setUploadProgress(0, '1. Tải file lên server');
  setCalcProgress(0, '2. Chờ tính toán');
}

function finishProgress(success) {
  clearInterval(progressTimer);
  progressTimer = null;
  if (success) {
    setUploadProgress(100, '1. Tải file hoàn tất');
    setCalcProgress(100, '2. Hoàn tất Stage 2');
    setTimeout(() => document.getElementById('analyze_progress').classList.add('hidden'), 900);
  } else {
    document.getElementById('analyze_progress').classList.add('hidden');
  }
}

function savePolicy() {
  const companyProfile = document.getElementById('company_profile').value;
  policy = {
    max_loss_percent: companyProfile === 'son_phuong' ? '' : document.getElementById('max_loss_percent').value,
    max_profit_percent: companyProfile === 'son_phuong' ? '' : document.getElementById('max_profit_percent').value,
    ignore_sale_suffix: companyProfile === 'son_phuong' ? false : document.getElementById('ignore_sale_suffix').checked,
    allow_negative_export: companyProfile === 'son_phuong' ? false : document.getElementById('allow_negative_export').checked,
    allow_future_purchase_reorder: companyProfile === 'son_phuong' ? false : document.getElementById('allow_future_purchase_reorder').checked,
    future_purchase_window_days: document.getElementById('future_purchase_window_days').value,
    company_profile: companyProfile,
    son_phuong_split_counts: {
      pipe_box: document.getElementById('sp_split_pipe_box').value,
      box: document.getElementById('sp_split_box').value,
      pipe: document.getElementById('sp_split_pipe').value
    },
    generic_split_variance_percent: document.getElementById('generic_split_variance_percent').value,
    barem_tolerance_percent: document.getElementById('barem_tolerance_percent').value,
    generic_min_take_quantity: document.getElementById('generic_min_take_quantity').value,
    generic_max_take_quantity: document.getElementById('generic_max_take_quantity').value
  };
  localStorage.setItem('inventory_policy', JSON.stringify(policy));
  localStorage.setItem('inventory_policy_version', policyVersion);
  renderPolicyLabelSafe();
}

function draftPolicy() {
  const companyProfile = document.getElementById('company_profile').value;
  return {
    max_loss_percent: companyProfile === 'son_phuong' ? '' : document.getElementById('max_loss_percent').value,
    max_profit_percent: companyProfile === 'son_phuong' ? '' : document.getElementById('max_profit_percent').value,
    ignore_sale_suffix: companyProfile === 'son_phuong' ? false : document.getElementById('ignore_sale_suffix').checked,
    allow_negative_export: companyProfile === 'son_phuong' ? false : document.getElementById('allow_negative_export').checked,
    allow_future_purchase_reorder: companyProfile === 'son_phuong' ? false : document.getElementById('allow_future_purchase_reorder').checked,
    future_purchase_window_days: document.getElementById('future_purchase_window_days').value,
    company_profile: companyProfile,
    son_phuong_split_counts: {
      pipe_box: document.getElementById('sp_split_pipe_box').value,
      box: document.getElementById('sp_split_box').value,
      pipe: document.getElementById('sp_split_pipe').value
    },
    generic_split_variance_percent: document.getElementById('generic_split_variance_percent').value,
    barem_tolerance_percent: document.getElementById('barem_tolerance_percent').value,
    generic_min_take_quantity: document.getElementById('generic_min_take_quantity').value,
    generic_max_take_quantity: document.getElementById('generic_max_take_quantity').value
  };
}

function renderCompanyDetectionRules(profile) {
  const target = document.getElementById('company_detection_rules');
  if (!target) return;
  if (profile === 'son_phuong') {
    target.innerHTML = `
      <div class="detection-card">
        <strong>Sơn Phương - phân biệt theo Tên hàng hóa trong file mua vào</strong>
        <div class="detection-grid">
          <div><b>Thép hộp</b><span>Có "ống hộp", "hộp", "vuông", "CN", hoặc kích thước dạng 4 phần như 30x60x1.4x6.0.</span></div>
          <div><b>Thép ống</b><span>Có "ống", "tròn", "Φ", "Ø", "F" hoặc "phi". Riêng "thép ống hộp các loại" được phép lấy cả ống và hộp.</span></div>
          <div><b>Không lấy vào ống/hộp</b><span>Những tên không nhận diện rõ như thép cuộn cán, dầu, mỡ, phụ kiện khác sẽ không vào pool chọn mã ống/hộp.</span></div>
        </div>
      </div>`;
    return;
  }
  target.innerHTML = `
    <div class="detection-card">
      <strong>Yến Thanh - phân biệt theo mã VT đã xử lý</strong>
      <div class="detection-grid">
        <div><b>Mã gốc và hậu tố</b><span>Mã bán ra có thể gắn .001/.002/.003 theo lô mua vào cùng mã gốc.</span></div>
        <div><b>Chọn lô</b><span>Ưu tiên lô theo ngày nhập, tồn kho và khoảng lãi/lỗ đã cấu hình.</span></div>
        <div><b>Không đủ kho HH</b><span>Phần không lấy được từ kho hàng hóa sẽ chuyển sang kho thành phẩm, trừ khi bật xuất âm.</span></div>
      </div>
    </div>`;
}

function renderSteelDetectionEmpty(message) {
  const status = document.getElementById('steel_detection_status');
  const summary = document.getElementById('steel_detection_summary');
  const table = document.getElementById('steel_detection_table');
  const rows = document.getElementById('steel_detection_rows');
  const filterButton = document.getElementById('filter_unknown_steel');
  if (status) status.textContent = message;
  if (summary) summary.innerHTML = '';
  if (rows) rows.innerHTML = '';
  if (table) table.classList.add('hidden');
  if (filterButton) filterButton.classList.add('hidden');
  steelDetectionRows = [];
  steelDetectionUnknownRows = [];
  steelDetectionUnknownOnly = false;
  steelDetectionCompanyRules = [];
  document.getElementById('company_rule_panel')?.classList.add('hidden');
  const companyRows = document.getElementById('company_rule_rows');
  if (companyRows) companyRows.innerHTML = '';
}

function renderCompanyRuleTable() {
  const panel = document.getElementById('company_rule_panel');
  const rows = document.getElementById('company_rule_rows');
  if (!panel || !rows) return;
  panel.classList.toggle('hidden', steelDetectionCompanyRules.length === 0);
  rows.innerHTML = steelDetectionCompanyRules.map(rule => {
    const topProfiles = (rule.top_profiles || []).slice(0, 5).map(profile =>
      `${profile.profile_code} (${numberLabel(profile.quantity)} kg/${numberLabel(profile.count)} dòng)`
    ).join('; ');
    const examples = (rule.examples || []).slice(0, 4).map(example =>
      `${example.source_variant_code || ''} → ${example.profile_code || 'không rõ'}`
    ).join('; ');
    return `
      <tr>
        <td class="text-cell">${escapeHtml(rule.company || '')}</td>
        <td>${escapeHtml(rule.tax_code || '')}</td>
        <td class="num">${numberLabel(rule.total)}</td>
        <td class="num">${numberLabel(rule.pipe)}</td>
        <td class="num">${numberLabel(rule.box)}</td>
        <td class="num">${numberLabel(rule.unknown)}</td>
        <td class="num">${numberLabel(rule.profile_count)}</td>
        <td class="text-cell">${escapeHtml(topProfiles)}</td>
        <td class="text-cell">${escapeHtml(examples)}</td>
      </tr>`;
  }).join('');
}

function renderSteelDetectionTable() {
  const visibleRows = steelDetectionUnknownOnly ? steelDetectionUnknownRows : steelDetectionRows;
  const table = document.getElementById('steel_detection_table');
  const rows = document.getElementById('steel_detection_rows');
  const filterButton = document.getElementById('filter_unknown_steel');
  if (rows) {
    rows.innerHTML = visibleRows.map(row => `
      <tr class="steel-${escapeHtml(row.kind)}">
        <td>${escapeHtml(row.row_number)}</td>
        <td><code>${escapeHtml(row.variant_code)}</code></td>
        <td><code>${escapeHtml(row.profile_code || '')}</code></td>
        <td>${escapeHtml(row.product_name)}</td>
        <td><b>${escapeHtml(row.kind_label)}</b></td>
        <td>${escapeHtml(row.reason)}</td>
        <td class="num">${numberLabel(row.quantity)}</td>
        <td class="num">${priceLabel(row.unit_price)}</td>
      </tr>`).join('');
  }
  if (table) table.classList.toggle('hidden', visibleRows.length === 0);
  if (filterButton) {
    filterButton.classList.toggle('hidden', steelDetectionRows.length === 0 && steelDetectionUnknownRows.length === 0);
    filterButton.classList.toggle('active', steelDetectionUnknownOnly);
    filterButton.textContent = steelDetectionUnknownOnly
      ? 'Hien tat ca dong'
      : `Chi hien khong phan loai (${numberLabel(steelDetectionUnknownRows.length)})`;
  }
}

async function renderPurchaseClassificationPreview() {
  const input = document.getElementById('purchase_file');
  const status = document.getElementById('steel_detection_status');
  if (!input || !input.files.length) {
    renderSteelDetectionEmpty('Chọn file mua vào để xem bảng phân loại.');
    return;
  }
  saveMapping();
  const form = new FormData();
  form.append('file', input.files[0]);
  form.append('mapping', JSON.stringify(mapping));
  form.append('policy', JSON.stringify(draftPolicy()));
  if (status) status.textContent = 'Đang đọc file mua vào và phân loại ống/hộp...';
  try {
    const response = await fetch('/api/purchase-classification-preview', { method: 'POST', body: form });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Không phân loại được file mua vào.');
    const counts = data.counts || {};
    document.getElementById('steel_detection_summary').innerHTML = [
      summaryCard(data.total || 0, 'Dòng mua vào'),
      summaryCard(counts.pipe || 0, 'Thép ống'),
      summaryCard(counts.box || 0, 'Thép hộp'),
      summaryCard(counts.unknown || 0, 'Không phân loại')
    ].join('');
    document.getElementById('steel_detection_status').textContent = data.profile === 'son_phuong'
      ? `Đã phân loại theo Tên hàng hóa. ${data.limited ? 'Bảng dưới chỉ hiển thị 250 dòng đầu.' : 'Bảng dưới hiển thị toàn bộ dòng đọc được.'}`
      : 'Yến Thanh không dùng phân loại ống/hộp, bảng này chỉ để kiểm tra tên/mã mua vào.';
    steelDetectionRows = data.rows || [];
    steelDetectionUnknownRows = data.unknown_rows || [];
    steelDetectionCompanyRules = data.company_rules || [];
    steelDetectionUnknownOnly = false;
    renderSteelDetectionTable();
    renderCompanyRuleTable();
  } catch (error) {
    renderSteelDetectionEmpty(error.message);
  }
}

function renderPolicyLabel() {
  const current = draftPolicy();
  const loss = current.max_loss_percent === '' ? 'không giới hạn' : `-${numberLabel(current.max_loss_percent)}%`;
  const profit = current.max_profit_percent === '' ? 'không giới hạn' : `+${numberLabel(current.max_profit_percent)}%`;
  document.getElementById('range_label').textContent =
    current.max_loss_percent === '' && current.max_profit_percent === ''
      ? 'Không giới hạn khoảng lãi/lỗ'
      : `Nhận mã từ ${loss} đến ${profit}`;
}

function renderPolicyLabelSafe() {
  const current = draftPolicy();
  const isSonPhuong = current.company_profile === 'son_phuong';
  renderCompanyDetectionRules(current.company_profile);
  document.getElementById('max_loss_percent').disabled = current.ignore_sale_suffix || isSonPhuong;
  document.getElementById('max_profit_percent').disabled = current.ignore_sale_suffix || isSonPhuong;
  document.querySelectorAll('.son-phuong-field').forEach(item => {
    item.classList.toggle('hidden', !isSonPhuong);
  });
  ['sp_split_pipe_box', 'sp_split_box', 'sp_split_pipe', 'generic_split_variance_percent'].forEach(id => {
    document.getElementById(id)?.closest('.range-field')?.classList.toggle('hidden', true);
  });
  document.querySelectorAll('.yen-thanh-field').forEach(item => {
    item.classList.toggle('hidden', isSonPhuong);
  });
  document.querySelectorAll('.shared-policy-field').forEach(item => {
    item.classList.toggle('hidden', isSonPhuong);
  });
  document.getElementById('policy_company_hint').textContent = isSonPhuong
    ? 'Priority chọn mã riêng cho Sơn Phương.'
    : 'Điều kiện riêng cho Yến Thanh.';
  document.getElementById('policy_help').innerHTML = isSonPhuong
    ? 'Sơn Phương <strong>không dùng khoảng lãi/lỗ</strong>. File barem áp trực tiếp theo mã VT; nếu mã khác công ty thì app dò theo loại ống/hộp và dimension trong tên hàng.'
    : 'Tỷ lệ được tính theo <strong>(đơn giá bán - đơn giá vốn) / đơn giá bán</strong>. Yến Thanh có thêm tùy chọn gộp hậu tố .001/.002 và xuất âm.';
  document.getElementById('unlimit_policy').classList.toggle('hidden', isSonPhuong);
  if (isSonPhuong) {
    document.getElementById('range_label').textContent = 'Sơn Phương: chọn mã theo nhóm ống/hộp, lãi thấp nhất, tồn nhiều nhất, rồi làm tròn theo barem chung.';
    return;
  }
  if (current.ignore_sale_suffix) {
    const futureText = current.allow_future_purchase_reorder ? `, keo HD mua vao sau ngay ban trong ${numberLabel(current.future_purchase_window_days || 31)} ngay` : '';
    document.getElementById('range_label').textContent = `Ban ra dung ma goc, khong gioi han khoang lai/lo${current.allow_negative_export ? ', KTP duoc xuat am' : ''}${futureText}`;
    return;
  }
  const loss = current.max_loss_percent === '' ? 'khong gioi han' : `-${numberLabel(current.max_loss_percent)}%`;
  const profit = current.max_profit_percent === '' ? 'khong gioi han' : `+${numberLabel(current.max_profit_percent)}%`;
  const profile = current.company_profile === 'son_phuong' ? 'Son Phuong' : 'Yen Thanh';
  const futureText = current.allow_future_purchase_reorder ? `, keo HD mua vao sau ngay ban trong ${numberLabel(current.future_purchase_window_days || 31)} ngay` : '';
  document.getElementById('range_label').textContent =
    current.max_loss_percent === '' && current.max_profit_percent === ''
      ? `${profile}: Khong gioi han khoang lai/lo${current.allow_negative_export ? ', KTP duoc xuat am' : ''}${futureText}`
      : `${profile}: Nhan ma tu ${loss} den ${profit}${current.allow_negative_export ? ', KTP duoc xuat am' : ''}${futureText}`;
}

async function previewFile(kind) {
  const input = document.getElementById(`${kind}_file`);
  const name = document.getElementById(`${kind}_name`);
  const preview = document.getElementById(`${kind}_preview`);
  if (!input.files.length) {
    name.textContent = kind === 'opening' ? 'Không có tồn đầu kỳ' : 'Chưa chọn file';
    preview.innerHTML = '';
    delete previewData[kind];
    renderMapping();
    if (kind === 'purchase') renderSteelDetectionEmpty('Chọn file mua vào để xem bảng phân loại.');
    return;
  }
  name.textContent = input.files[0].name;
  const form = new FormData();
  form.append('file', input.files[0]);
  form.append('header_row', mapping[kind].header_row);
  preview.textContent = 'Đang đọc file...';
  try {
    const response = await fetch('/api/preview', { method: 'POST', body: form });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Không đọc được file.');
    previewData[kind] = data;
    renderMapping();
    preview.innerHTML = `<p>Sheet: <strong>${escapeHtml(data.active_sheet)}</strong> - tên cột lấy từ dòng <strong>${data.header_row}</strong> - đến cột ${escapeHtml(data.max_column)}</p>
      <table>
        <thead><tr>${data.preview_letters.map(letter => `<th>${escapeHtml(letter)}</th>`).join('')}</tr></thead>
        <tbody>${data.rows.map((row, index) => `<tr class="${index + 1 === data.header_row ? 'selected-header' : ''}">${row.map(value => `<td>${escapeHtml(value)}</td>`).join('')}</tr>`).join('')}</tbody>
      </table>`;
    if (kind === 'purchase') renderPurchaseClassificationPreview();
  } catch (error) {
    preview.textContent = error.message;
    if (kind === 'purchase') renderSteelDetectionEmpty(error.message);
  }
}

function summaryCard(value, label, suffix = '') {
  return `<div class="stat"><strong>${numberLabel(value)}${suffix}</strong><span>${label}</span></div>`;
}

function percentLabel(value) {
  if (value === null || value === undefined) return '-';
  return `${signedNumberLabel(value)}%`;
}

function lotSourceLabel(lot) {
  const invoice = lot.invoice_no ? `HĐ ${lot.invoice_no}` : `dòng ${lot.row_number}`;
  const location = lot.invoice_date ? `${invoice}, ngày ${lot.invoice_date}` : invoice;
  if (lot.summary_count > 1) return `${lot.source || 'Mua vào'} - ${numberLabel(lot.summary_count)} dòng mua, giá vốn TB - ${location}`;
  return `${lot.source || 'Mua vào'} - ${location}`;
}

function saleRowTable(row) {
  return `
    <div class="sale-row-card">
      <div>
        <span class="code-group-label">Dòng bán ra</span>
        <strong>${escapeHtml(row.product_name || '-')}</strong>
      </div>
      <div class="sale-row-meta">
        <span>HĐ ${escapeHtml(row.invoice_no || '-')}</span>
        <span>Ngày ${escapeHtml(row.invoice_date || '-')}</span>
        <span>Mã bán: <code>${escapeHtml(row.variant_code)}</code></span>
        <span>SL bán: <b>${numberLabel(row.quantity)}</b></span>
        <span class="from-stock">Mua vào/tồn: ${numberLabel(row.material_quantity)}</span>
        <span class="from-finished">Thành phẩm: ${numberLabel(row.finished_quantity)}</span>
      </div>
    </div>`;
}

function allocationUsageTable(row) {
  if (!row.used.length) {
    return '<p class="empty-result">Không lấy được số lượng nào từ hóa đơn mua vào hoặc tồn đầu kỳ.</p>';
  }
  return `
    <div class="split-list">
      ${row.used.map(lot => `
        <div class="split-card">
          <div class="split-code-block">
            <span>Mã bán được gắn</span>
            <code>${escapeHtml(lot.ledger_variant_code || lot.purchase_variant_code || lot.variant_code)}</code>
          </div>
          <div class="split-code-block">
            <span>Mã mua vào</span>
            <code>${escapeHtml(lot.variant_code)}</code>
          </div>
          <div class="split-copy">
            <strong>${escapeHtml(lotSourceLabel(lot))}</strong>
            <span>HĐ mua vào: ${escapeHtml(lot.invoice_no || '-')} | Ngày: ${escapeHtml(lot.invoice_date || '-')}</span>
          </div>
          <div class="split-numbers">
            <span class="from-stock">SL lấy: ${numberLabel(lot.quantity)}</span>
            <span>Vốn: ${lot.unit_cost == null ? '-' : numberLabel(lot.unit_cost)}</span>
            <span>Lãi/lỗ: ${percentLabel(lot.profit_percent)}</span>
            ${lot.barem_weight ? `<span>Barem: ${numberLabel(lot.barem_weight)} kg</span>` : ''}
            ${lot.barem_remainder ? '<span class="from-finished">Phần kg dư cuối</span>' : ''}
            ${lot.negative_export ? '<span class="from-finished">Chưa ghép đủ KHH</span>' : ''}
          </div>
        </div>
      `).join('')}
    </div>`;
}

function allocationRejectedTable(row) {
  if (!row.rejected.length) return '';
  return `
    <div class="detail-subsection">
      <strong>Lô không đạt khoảng lãi/lỗ</strong>
      <table class="detail-table rejected-table">
        <thead><tr><th>Mã VT</th><th>Số HĐ mua vào</th><th>Ngày HĐ</th><th class="num">SL còn</th><th>Lý do</th></tr></thead>
        <tbody>
          ${row.rejected.map(lot => `
            <tr>
              <td><code>${escapeHtml(lot.variant_code)}</code></td>
              <td>${escapeHtml(lot.invoice_no || '-')}</td>
              <td>${escapeHtml(lot.invoice_date || '-')}</td>
              <td class="num">${numberLabel(lot.quantity)}</td>
              <td>${escapeHtml(lot.reason)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>`;
}

function allocationBalanceTable(row) {
  const balances = new Map();
  const keyFor = lot => `${lot.variant_code}|${lot.source}|${lot.invoice_no || ''}|${lot.row_number}`;
  row.inventory_before.forEach(lot => {
    balances.set(keyFor(lot), { ...lot, before: lot.quantity, after: 0 });
  });
  row.inventory_after.forEach(lot => {
    const key = keyFor(lot);
    const item = balances.get(key) || { ...lot, before: 0, after: 0 };
    item.after = lot.quantity;
    balances.set(key, item);
  });
  if (!balances.size) return '';
  return `
    <div class="detail-subsection balance-subsection">
      <strong>Tồn kho liên quan trước và sau khi bán</strong>
      <table class="detail-table">
        <thead><tr><th>Mã VT</th><th>Nguồn</th><th class="num">Trước khi bán</th><th class="num">Sau khi bán</th></tr></thead>
        <tbody>
          ${Array.from(balances.values()).map(lot => `
            <tr>
              <td><code>${escapeHtml(lot.variant_code)}</code></td>
              <td>${escapeHtml(lotSourceLabel(lot))}</td>
              <td class="num">${numberLabel(lot.before)}</td>
              <td class="num">${numberLabel(lot.after)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>`;
}

function allocationDetail(row) {
  const invoice = row.invoice_no ? `Hóa đơn bán ra ${escapeHtml(row.invoice_no)}` : 'Hóa đơn bán ra';
  return `
    <section class="allocation-card">
      ${saleRowTable(row)}
      <div class="detail-subsection split-subsection"><strong>Tách dòng bán ra theo mã VT mua vào</strong></div>
      ${allocationUsageTable(row)}
      <details class="allocation-more">
        <summary>Tồn kho trước/sau và lô bị loại</summary>
        ${allocationBalanceTable(row)}
        ${allocationRejectedTable(row)}
      </details>
    </section>`;
}

function invoiceGroupKey(row) {
  return `${row.invoice_no || 'khong-so-hd'}|||${row.invoice_date || ''}`;
}

function invoiceGroupTitle(group) {
  const invoice = group.invoice_no ? `Hóa đơn bán ra ${escapeHtml(group.invoice_no)}` : 'Hóa đơn bán ra chưa có số';
  const date = group.invoice_date ? ` - ngày ${escapeHtml(group.invoice_date)}` : '';
  return `${invoice}${date}`;
}

function allocationInvoiceGroup(group) {
  return `
    <section class="invoice-group-card">
      <div class="invoice-group-head">
        <div>
          <span class="code-group-label">Hóa đơn bán ra</span>
          <strong>${invoiceGroupTitle(group)}</strong>
        </div>
        <div class="allocation-card-tags">
          <span>${numberLabel(group.rows.length)} dòng bán ra</span>
          <span>${numberLabel(group.used_count)} mã mua vào/tồn</span>
        </div>
      </div>
      <div class="invoice-row-list">
        ${group.rows.map(row => allocationDetail(row)).join('')}
      </div>
    </section>`;
}

function allocationGroups(rows) {
  const groups = new Map();
  rows.forEach(row => {
    const key = invoiceGroupKey(row);
    if (!groups.has(key)) {
      groups.set(key, {
        invoice_no: row.invoice_no || '',
        invoice_date: row.invoice_date || '',
        rows: [],
        used_count: 0
      });
    }
    const group = groups.get(key);
    group.rows.push(row);
    group.used_count += row.used?.length || 0;
  });
  return Array.from(groups.values());
}

function quantityLabel(value) {
  return Number(value || 0).toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function priceLabel(value) {
  if (value === null || value === undefined || value === '') return '';
  return Number(value || 0).toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function baseCode(value) {
  return String(value || '').replace(/\.\d{3}$/, '');
}

function sumLedger(rows, field) {
  return rows.reduce((total, row) => total + Number(row[field] || 0), 0);
}

function ledgerRowsForRange(group, fromDate, toDate) {
  const rows = group.rows || [];
  const openingRows = rows.filter(row => row.type === 'opening' || (fromDate && row.date_iso && row.date_iso < fromDate));
  const periodRows = rows.filter(row => (
    row.type !== 'opening'
    && !(fromDate && row.date_iso && row.date_iso < fromDate)
    && compareIsoDate(row.date_iso, fromDate, 'from')
    && compareIsoDate(row.date_iso, toDate, 'to')
  ));
  const openingQty = sumLedger(openingRows, 'qty_in') - sumLedger(openingRows, 'qty_out');
  const openingAmount = sumLedger(openingRows, 'amount_in') - sumLedger(openingRows, 'amount_out');
  const inQty = sumLedger(periodRows, 'qty_in');
  const inAmount = sumLedger(periodRows, 'amount_in');
  const outQty = sumLedger(periodRows, 'qty_out');
  const outAmount = sumLedger(periodRows, 'amount_out');
  let runningQty = openingQty;
  let runningAmount = openingAmount;
  const rowsWithBalance = periodRows.map(row => {
    runningQty += Number(row.qty_in || 0) - Number(row.qty_out || 0);
    runningAmount += Number(row.amount_in || 0) - Number(row.amount_out || 0);
    return { ...row, running_qty: runningQty, running_amount: runningAmount };
  });
  return {
    periodRows: rowsWithBalance,
    openingQty,
    openingAmount,
    inQty,
    inAmount,
    outQty,
    outAmount,
    endingQty: openingQty + inQty - outQty,
    endingAmount: openingAmount + inAmount - outAmount
  };
}

function combinedRowSortKey(row) {
  const typeOrder = { opening: 0, purchase: 1, finished_receipt: 1, sale: 2 };
  const warehouseOrder = { KHH: 0, KTP: 1 };
  return [
    row.date_iso || '0000-00-00',
    row.doc_no || '',
    typeOrder[row.type] ?? 9,
    warehouseOrder[row.warehouse_code] ?? 9,
    row.sequence || 0
  ].join('|');
}

function hasLedgerActivity(view) {
  return view.periodRows.length
    || Math.abs(view.openingQty) > 0.0000001
    || Math.abs(view.openingAmount) > 0.0000001
    || Math.abs(view.endingQty) > 0.0000001
    || Math.abs(view.endingAmount) > 0.0000001;
}

function inventorySummaryRows(warehouse, fromDate, toDate) {
  return (warehouse.groups || []).map(group => {
    const view = ledgerRowsForRange(group, fromDate, toDate);
    return {
      group,
      view
    };
  }).filter(item => hasLedgerActivity(item.view));
}

function docTypeLabel(row) {
  if (row.type === 'purchase') return 'PN';
  if (row.type === 'opening') return 'TD';
  return 'PX';
}

function accountName(row) {
  if (row.account === '331') return 'Phải trả cho người bán';
  if (row.account === '6321') return 'Xuất kho giá vốn';
  if (row.account === '154') return 'Chi phí sản xuất kinh doanh dở dang';
  return '';
}

function summaryForRange(fromDate, toDate) {
  const sourceRows = currentResult?.allocations || [];
  const rows = sourceRows.filter(row => isoInRange(row.invoice_date_iso, fromDate, toDate));
  return rows.reduce((summary, row) => {
    summary.sales_quantity += Number(row.quantity || 0);
    summary.material_quantity += Number(row.material_quantity || 0);
    summary.finished_quantity += Number(row.finished_quantity || 0);
    summary.negative_export_quantity += (row.used || []).reduce(
      (total, item) => total + (item.negative_export ? Number(item.quantity || 0) : 0),
      0
    );
    if (Number(row.rejected_count || 0) > 0 || (row.rejected || []).length) summary.range_rejected_lines += 1;
    return summary;
  }, {
    sales_quantity: 0,
    material_quantity: 0,
    finished_quantity: 0,
    negative_export_quantity: 0,
    range_rejected_lines: 0,
    missing_barem_count: currentResult?.summary?.missing_barem_count || 0
  });
}

function renderSummaryCards(summary) {
  const isSonPhuong = currentResult?.policy?.company_profile === 'son_phuong';
  document.getElementById('summary').innerHTML = [
    summaryCard(summary.sales_quantity, 'Số lượng bán ra'),
    summaryCard(summary.material_quantity, 'Từ kho hàng hóa'),
    summaryCard(isSonPhuong ? summary.negative_export_quantity : summary.finished_quantity, isSonPhuong ? 'Chưa ghép đủ KHH' : 'Từ kho thành phẩm'),
    summaryCard(
      isSonPhuong ? summary.missing_barem_count : summary.range_rejected_lines,
      isSonPhuong ? 'Mã thiếu barem' : 'Dòng bị loại do lãi/lỗ'
    )
  ].join('');
}

function ledgerSummaryRow(label, inQty, inAmount, outQty = '', outAmount = '', balanceQty = '', balanceAmount = '') {
  return `
    <tr class="ledger-summary-row">
      <td></td><td></td><td></td><td><strong>${escapeHtml(label)}</strong></td><td></td><td></td>
      <td class="num"><strong>${inQty === '' ? '' : quantityLabel(inQty)}</strong></td>
      <td class="num"><strong>${inAmount === '' ? '' : moneyLabel(inAmount)}</strong></td>
      <td class="num"><strong>${outQty === '' ? '' : quantityLabel(outQty)}</strong></td>
      <td class="num"><strong>${outAmount === '' ? '' : moneyLabel(outAmount)}</strong></td>
      <td class="num"><strong>${balanceQty === '' ? '' : quantityLabel(balanceQty)}</strong></td>
      <td class="num"><strong>${balanceAmount === '' ? '' : moneyLabel(balanceAmount)}</strong></td>
    </tr>`;
}

function ledgerRowClasses(row, warehouseCode = '') {
  return [
    row.qty_in ? 'ledger-row-in' : '',
    row.qty_out ? 'ledger-row-out' : '',
    row.future_purchase_reordered ? 'ledger-row-future-reorder' : '',
    `ledger-warehouse-${String(warehouseCode || row.warehouse_code || '').toLowerCase()}`
  ].filter(Boolean).join(' ');
}

function ledgerTransactionRow(row, warehouseCode = '') {
  return `
    <tr class="${ledgerRowClasses(row, warehouseCode)}">
      <td>${escapeHtml(row.date || dateLabelFromIso(row.date_iso))}</td>
      <td>${escapeHtml(row.doc_no || '')}</td>
      <td class="text-cell">${escapeHtml(row.customer || '')}</td>
      <td class="text-cell">${escapeHtml(row.description || '')}</td>
      <td>${escapeHtml(row.account || '')}</td>
      <td class="num">${priceLabel(row.unit_price)}</td>
      <td class="num">${quantityLabel(row.qty_in)}</td>
      <td class="num">${moneyLabel(row.amount_in)}</td>
      <td class="num">${quantityLabel(row.qty_out)}</td>
      <td class="num">${moneyLabel(row.amount_out)}</td>
      <td class="num">${quantityLabel(row.running_qty)}</td>
      <td class="num">${moneyLabel(row.running_amount)}</td>
    </tr>`;
}

function ledgerGroupHtml(group, warehouse, fromDate, toDate) {
  const view = ledgerRowsForRange(group, fromDate, toDate);
  const hasActivity = view.periodRows.length || Math.abs(view.openingQty) > 0.0000001 || Math.abs(view.endingQty) > 0.0000001;
  if (!hasActivity) return '';
  return `
    <tr class="ledger-group-row">
      <td colspan="12">Vật tư: ${escapeHtml(group.variant_code)} - ${escapeHtml(group.product_name || '')}, Đvt: ${escapeHtml(group.unit_name || '-')}, TK: ${escapeHtml(group.account || warehouse.account || '156')}</td>
    </tr>
    ${ledgerSummaryRow('Tồn đầu kỳ', view.openingQty, view.openingAmount, '', '', view.openingQty, view.openingAmount)}
    ${ledgerSummaryRow('Nhập trong kỳ', view.inQty, view.inAmount)}
    ${ledgerSummaryRow('Xuất trong kỳ', '', '', view.outQty, view.outAmount)}
    ${ledgerSummaryRow('Tồn cuối kỳ', view.endingQty, view.endingAmount, '', '', view.endingQty, view.endingAmount)}
    ${view.periodRows.map(row => ledgerTransactionRow(row, warehouse.warehouse_code || '')).join('')}
  `;
}

function currentLedgerContext() {
  if (!currentResult?.ledger) return;
  const ledger = currentResult.ledger;
  const fromDate = document.getElementById('ledger_from').value;
  const toDate = document.getElementById('ledger_to').value;
  const warehouses = ledger.warehouses || [ledger];
  if (warehouses.length === 1 && activeLedgerWarehouse === 'ALL') {
    activeLedgerWarehouse = warehouses[0]?.warehouse_code || 'KHH';
  }
  const activeWarehouse = activeLedgerWarehouse === 'ALL'
    ? null
    : (warehouses.find(item => item.warehouse_code === activeLedgerWarehouse) || warehouses[0] || {});
  if (activeWarehouse) activeLedgerWarehouse = activeWarehouse.warehouse_code || activeLedgerWarehouse;
  return { fromDate, toDate, warehouses, activeWarehouse };
}

function renderLedgerShell(context) {
  const { fromDate, toDate, warehouses, activeWarehouse } = context;
  document.getElementById('ledger_warehouse').textContent =
    activeLedgerWarehouse === 'ALL'
      ? 'KHO: KHH + KTP'
      : `KHO: ${activeWarehouse?.warehouse_code || ''} - ${activeWarehouse?.warehouse_name || ''}`;
  document.getElementById('ledger_range_text').textContent =
    `TỪ NGÀY: ${dateLabelFromIso(fromDate)} ĐẾN NGÀY: ${dateLabelFromIso(toDate)}`;
  renderSummaryCards(summaryForRange(fromDate, toDate));
  document.getElementById('ledger_warehouse_tabs').innerHTML = [
    ...warehouses.map(warehouse => `
    <button class="warehouse-tab ${warehouse.warehouse_code === activeLedgerWarehouse ? 'active' : ''}" type="button" data-warehouse="${escapeHtml(warehouse.warehouse_code || '')}">
      ${escapeHtml(warehouse.warehouse_code || '')} - ${escapeHtml(warehouse.warehouse_name || '')}
    </button>`),
    ...(warehouses.length > 1
      ? [`<button class="warehouse-tab ${activeLedgerWarehouse === 'ALL' ? 'active' : ''}" type="button" data-warehouse="ALL">KHH + KTP</button>`]
      : [])
  ].join('');
}

function combinedLedgerSections(fromDate, toDate) {
  const context = currentLedgerContext();
  if (!context) return [];
  const sections = new Map();
  (context.warehouses || []).forEach(warehouse => {
    (warehouse.groups || []).forEach(group => {
      const view = ledgerRowsForRange(group, fromDate, toDate);
      if (!hasLedgerActivity(view)) return;
      const key = baseCode(group.variant_code);
      if (!sections.has(key)) {
        sections.set(key, {
          warehouse: { warehouse_code: 'KHH+KTP', warehouse_name: 'KHO HANG HOA + KHO THANH PHAM' },
          base_code: baseCode(group.variant_code),
          product_name: group.product_name || '',
          unit_name: group.unit_name || '',
          account: group.account || warehouse.account || '',
          detail_codes: new Set(),
          opening_by_warehouse: {},
          rows: []
        });
      }
      const section = sections.get(key);
      section.detail_codes.add(group.variant_code || '');
      const warehouseKey = warehouse.warehouse_code || '';
      section.opening_by_warehouse[warehouseKey] = section.opening_by_warehouse[warehouseKey] || { qty: 0, amount: 0 };
      section.opening_by_warehouse[warehouseKey].qty += view.openingQty || 0;
      section.opening_by_warehouse[warehouseKey].amount += view.openingAmount || 0;
      view.periodRows.forEach(row => {
        section.rows.push({ ...row, warehouse_code: warehouse.warehouse_code || '', variant_code: row.variant_code || group.variant_code || '' });
        if (row.variant_code) section.detail_codes.add(row.variant_code);
      });
      section.openingQty = (section.openingQty || 0) + view.openingQty;
      section.openingAmount = (section.openingAmount || 0) + view.openingAmount;
      section.inQty = (section.inQty || 0) + view.inQty;
      section.inAmount = (section.inAmount || 0) + view.inAmount;
      section.outQty = (section.outQty || 0) + view.outQty;
      section.outAmount = (section.outAmount || 0) + view.outAmount;
      section.endingQty = (section.endingQty || 0) + view.endingQty;
      section.endingAmount = (section.endingAmount || 0) + view.endingAmount;
    });
  });
  return Array.from(sections.values()).map(section => {
    section.rows.sort((a, b) => combinedRowSortKey(a).localeCompare(combinedRowSortKey(b)));
    const runningBalances = new Map(Object.entries(section.opening_by_warehouse || {}).map(([key, value]) => [
      key,
      { qty: Number(value.qty || 0), amount: Number(value.amount || 0) }
    ]));
    section.rows = section.rows.map(row => {
      const key = row.warehouse_code || '';
      const current = runningBalances.get(key) || { qty: 0, amount: 0 };
      current.qty += Number(row.qty_in || 0) - Number(row.qty_out || 0);
      current.amount += Number(row.amount_in || 0) - Number(row.amount_out || 0);
      runningBalances.set(key, current);
      return { ...row, running_qty: current.qty, running_amount: current.amount };
    });
    return section;
  }).sort((a, b) => a.base_code.localeCompare(b.base_code));
}

function combinedSummaryRow(label, inQty, inAmount, outQty = '', outAmount = '', balanceQty = '', balanceAmount = '') {
  return `
    <tr class="ledger-summary-row">
      <td></td><td></td><td></td><td><strong>${escapeHtml(label)}</strong></td><td></td><td></td><td></td><td></td>
      <td class="num"><strong>${inQty === '' ? '' : quantityLabel(inQty)}</strong></td>
      <td class="num"><strong>${inAmount === '' ? '' : moneyLabel(inAmount)}</strong></td>
      <td class="num"><strong>${outQty === '' ? '' : quantityLabel(outQty)}</strong></td>
      <td class="num"><strong>${outAmount === '' ? '' : moneyLabel(outAmount)}</strong></td>
      <td></td><td></td><td></td>
      <td class="num"><strong>${balanceQty === '' ? '' : quantityLabel(balanceQty)}</strong></td>
      <td class="num"><strong>${balanceAmount === '' ? '' : moneyLabel(balanceAmount)}</strong></td>
      <td></td>
    </tr>`;
}

function combinedTransactionRow(row, warehouse) {
  const saleAmount = row.sale_amount === '' || row.sale_amount === undefined ? '' : Number(row.sale_amount || 0);
  const costAmount = Number(row.amount_out || 0);
  const margin = row.margin_percent !== undefined && row.margin_percent !== null && row.margin_percent !== ''
    ? Number(row.margin_percent)
    : (row.type === 'sale' && row.warehouse_code !== 'KTP' && saleAmount ? ((saleAmount - costAmount) / saleAmount * 100) : '');
  return `
    <tr class="${ledgerRowClasses(row, row.warehouse_code || warehouse.warehouse_code || '')}">
      <td>${escapeHtml(row.date || dateLabelFromIso(row.date_iso))}</td>
      <td>${escapeHtml(row.doc_no || '')}</td>
      <td class="text-cell">${escapeHtml(row.customer || '')}</td>
      <td class="text-cell">${escapeHtml(row.description || '')}</td>
      <td>${escapeHtml(row.account || '')}</td>
      <td>${escapeHtml(row.warehouse_code || warehouse.warehouse_code || '')}</td>
      <td>${escapeHtml(row.variant_code || '')}</td>
      <td class="num">${priceLabel(row.unit_price)}</td>
      <td class="num">${quantityLabel(row.qty_in)}</td>
      <td class="num">${moneyLabel(row.amount_in)}</td>
      <td class="num">${quantityLabel(row.qty_out)}</td>
      <td class="num">${moneyLabel(row.amount_out)}</td>
      <td class="num">${priceLabel(row.sale_unit_price)}</td>
      <td class="num">${saleAmount === '' ? '' : moneyLabel(saleAmount)}</td>
      <td class="num">${margin === '' ? '' : signedNumberLabel(margin)}</td>
      <td class="num">${quantityLabel(row.running_qty)}</td>
      <td class="num">${moneyLabel(row.running_amount)}</td>
      <td class="text-cell logic-cell">${escapeHtml(row.logic_note || '')}</td>
    </tr>`;
}

function setLedgerTableMode(isCombined) {
  document.getElementById('single_ledger_scroll')?.classList.toggle('hidden', isCombined);
  document.getElementById('combined_ledger_scroll')?.classList.toggle('hidden', !isCombined);
}

function renderCombinedLedger(fromDate, toDate) {
  const rows = combinedLedgerSections(fromDate, toDate).map(section => `
    <tr class="ledger-warehouse-row"><td colspan="18">KHO: ${escapeHtml(section.warehouse.warehouse_code || '')} - ${escapeHtml(section.warehouse.warehouse_name || '')}</td></tr>
    <tr class="ledger-group-row"><td colspan="6">Vật tư: ${escapeHtml(section.base_code)} - ${escapeHtml(section.product_name || '')}, Đvt: ${escapeHtml(section.unit_name || '-')}, TK: ${escapeHtml(section.account || '')}</td><td colspan="12">${escapeHtml(Array.from(section.detail_codes).sort().join(', '))}</td></tr>
    ${combinedSummaryRow('Tồn đầu kỳ', section.openingQty || 0, section.openingAmount || 0, '', '', section.openingQty || 0, section.openingAmount || 0)}
    ${combinedSummaryRow('Nhập trong kỳ', section.inQty || 0, section.inAmount || 0)}
    ${combinedSummaryRow('Xuất trong kỳ', '', '', section.outQty || 0, section.outAmount || 0)}
    ${combinedSummaryRow('Tồn cuối kỳ', section.endingQty || 0, section.endingAmount || 0, '', '', section.endingQty || 0, section.endingAmount || 0)}
    ${section.rows.map(row => combinedTransactionRow(row, section.warehouse)).join('')}
  `).join('');
  document.getElementById('combined_ledger_rows').innerHTML = rows;
  document.getElementById('ledger_empty').classList.toggle('hidden', rows.length > 0);
}

function warehouseByCode(code) {
  return (currentResult?.ledger?.warehouses || []).find(warehouse => warehouse.warehouse_code === code)
    || (currentResult?.ledger?.warehouses || [])[0]
    || {};
}

function warehouseTabsHtml(activeCode) {
  return (currentResult?.ledger?.warehouses || []).map(warehouse => `
    <button class="warehouse-tab ${warehouse.warehouse_code === activeCode ? 'active' : ''}" type="button" data-warehouse="${escapeHtml(warehouse.warehouse_code || '')}">
      ${escapeHtml(warehouse.warehouse_code || '')} - ${escapeHtml(warehouse.warehouse_name || '')}
    </button>
  `).join('');
}

function renderInventorySummary() {
  if (!currentResult?.ledger) return;
  const fromDate = document.getElementById('ledger_from').value;
  const toDate = document.getElementById('ledger_to').value;
  const activeWarehouse = warehouseByCode(activeInventorySummaryWarehouse);
  activeInventorySummaryWarehouse = activeWarehouse.warehouse_code || activeInventorySummaryWarehouse;
  const rows = inventorySummaryRows(activeWarehouse, fromDate, toDate);
  document.getElementById('summary_warehouse_tabs').innerHTML = warehouseTabsHtml(activeInventorySummaryWarehouse);
  document.getElementById('inventory_summary_warehouse').textContent =
    `KHO: ${activeWarehouse.warehouse_code || ''} - ${activeWarehouse.warehouse_name || ''}`;
  document.getElementById('inventory_summary_range').textContent =
    `TỪ NGÀY: ${dateLabelFromIso(fromDate)} ĐẾN NGÀY: ${dateLabelFromIso(toDate)}`;
  document.getElementById('inventory_summary_empty').classList.toggle('hidden', rows.length > 0);
  document.getElementById('inventory_summary_rows').innerHTML = rows.map((item, index) => {
    const { group, view } = item;
    return `
      <tr data-summary-code="${escapeHtml(group.variant_code || '')}">
        <td class="num">${index + 1}</td>
        <td class="summary-code">${escapeHtml(group.variant_code || '')}</td>
        <td class="text-cell">${escapeHtml(group.product_name || '')}</td>
        <td>${escapeHtml(group.unit_name || '')}</td>
        <td class="num">${quantityLabel(view.openingQty)}</td>
        <td class="num">${moneyLabel(view.openingAmount)}</td>
        <td class="num">${quantityLabel(view.inQty)}</td>
        <td class="num">${moneyLabel(view.inAmount)}</td>
        <td class="num">${quantityLabel(view.outQty)}</td>
        <td class="num">${moneyLabel(view.outAmount)}</td>
        <td class="num">${quantityLabel(view.endingQty)}</td>
        <td class="num">${moneyLabel(view.endingAmount)}</td>
      </tr>`;
  }).join('');
  const totals = rows.reduce((sum, item) => {
    const view = item.view || {};
    sum.openingQty += view.openingQty || 0;
    sum.openingAmount += view.openingAmount || 0;
    sum.inQty += view.inQty || 0;
    sum.inAmount += view.inAmount || 0;
    sum.outQty += view.outQty || 0;
    sum.outAmount += view.outAmount || 0;
    sum.endingQty += view.endingQty || 0;
    sum.endingAmount += view.endingAmount || 0;
    return sum;
  }, { openingQty: 0, openingAmount: 0, inQty: 0, inAmount: 0, outQty: 0, outAmount: 0, endingQty: 0, endingAmount: 0 });
  document.getElementById('inventory_summary_footer').innerHTML = rows.length ? `
    <tr>
      <td colspan="4">Tổng cộng:</td>
      <td class="num">${quantityLabel(totals.openingQty)}</td>
      <td class="num">${moneyLabel(totals.openingAmount)}</td>
      <td class="num">${quantityLabel(totals.inQty)}</td>
      <td class="num">${moneyLabel(totals.inAmount)}</td>
      <td class="num">${quantityLabel(totals.outQty)}</td>
      <td class="num">${moneyLabel(totals.outAmount)}</td>
      <td class="num">${quantityLabel(totals.endingQty)}</td>
      <td class="num">${moneyLabel(totals.endingAmount)}</td>
    </tr>
  ` : '';
}

function showInventoryDetail(variantCode) {
  const warehouse = warehouseByCode(activeInventorySummaryWarehouse);
  const group = (warehouse.groups || []).find(item => item.variant_code === variantCode);
  if (!group) return;
  const fromDate = document.getElementById('ledger_from').value;
  const toDate = document.getElementById('ledger_to').value;
  const view = ledgerRowsForRange(group, fromDate, toDate);
  const rows = view.periodRows;
  document.getElementById('inventory_detail_title').textContent =
    `${group.variant_code || ''} - ${group.product_name || ''}`;
  document.getElementById('inventory_detail_subtitle').textContent =
    `KHO: ${warehouse.warehouse_code || ''} - ${warehouse.warehouse_name || ''} | Từ ngày ${dateLabelFromIso(fromDate)} đến ngày ${dateLabelFromIso(toDate)}`;
  document.getElementById('inventory_detail_rows').innerHTML = rows.map(row => `
    <tr class="${ledgerRowClasses(row, warehouse.warehouse_code || '')}">
      <td>${escapeHtml(row.date || dateLabelFromIso(row.date_iso))}</td>
      <td>${docTypeLabel(row)}</td>
      <td>${escapeHtml(row.doc_no || '')}</td>
      <td class="text-cell">${escapeHtml(row.customer || '')}</td>
      <td class="text-cell">${escapeHtml(row.description || '')}</td>
      <td>${escapeHtml(row.account || '')}</td>
      <td>${escapeHtml(warehouse.warehouse_code || '')}</td>
      <td class="num">${priceLabel(row.unit_price)}</td>
      <td class="num">${quantityLabel(row.qty_in)}</td>
      <td class="num">${moneyLabel(row.amount_in)}</td>
      <td class="num">${quantityLabel(row.qty_out)}</td>
      <td class="num">${moneyLabel(row.amount_out)}</td>
      <td>${escapeHtml(accountName(row))}</td>
      <td>${escapeHtml(row.customer_tax_code || '')}</td>
      <td>${docTypeLabel(row)}</td>
      <td>CTY</td>
    </tr>
  `).join('');
  document.getElementById('inventory_detail_footer').innerHTML = `
    <tr><td colspan="8" class="detail-total-label">Tồn đầu kỳ:</td><td class="num">${quantityLabel(view.openingQty)}</td><td class="num">${moneyLabel(view.openingAmount)}</td><td></td><td></td><td colspan="4"></td></tr>
    <tr><td colspan="8" class="detail-total-label">Nhập trong kỳ:</td><td class="num">${quantityLabel(view.inQty)}</td><td class="num">${moneyLabel(view.inAmount)}</td><td></td><td></td><td colspan="4"></td></tr>
    <tr><td colspan="8" class="detail-total-label">Xuất trong kỳ:</td><td></td><td></td><td class="num">${quantityLabel(view.outQty)}</td><td class="num">${moneyLabel(view.outAmount)}</td><td colspan="4"></td></tr>
    <tr><td colspan="8" class="detail-total-label">Tồn cuối kỳ:</td><td class="num">${quantityLabel(view.endingQty)}</td><td class="num">${moneyLabel(view.endingAmount)}</td><td></td><td></td><td colspan="4"></td></tr>
  `;
  document.getElementById('inventory_detail_modal').classList.remove('hidden');
}

function hideInventoryDetail() {
  document.getElementById('inventory_detail_modal').classList.add('hidden');
}

function salesRowsForWarehouse(warehouseCode) {
  const fromDate = document.getElementById('ledger_from').value;
  const toDate = document.getElementById('ledger_to').value;
  return (currentResult?.sales_report_rows || [])
    .filter(row => row.warehouse_code === warehouseCode && isoInRange(row.invoice_date_iso, fromDate, toDate));
}

function renderSalesSummary() {
  const warehouse = warehouseByCode(activeSalesSummaryWarehouse);
  activeSalesSummaryWarehouse = warehouse.warehouse_code || activeSalesSummaryWarehouse;
  const rows = salesRowsForWarehouse(activeSalesSummaryWarehouse);
  const grouped = new Map();
  rows.forEach(row => {
    const key = row.variant_code || '';
    if (!grouped.has(key)) {
      grouped.set(key, {
        variant_code: key,
        product_name: row.product_name || '',
        unit_name: row.unit_name || '',
        quantity: 0,
        cost_amount: 0,
        sale_amount: 0,
        profit_amount: 0,
        tax_amount: 0,
        total_amount: 0
      });
    }
    const item = grouped.get(key);
    item.quantity += Number(row.quantity || 0);
    item.cost_amount += Number(row.cost_amount || 0);
    item.sale_amount += Number(row.sale_amount || 0);
    item.profit_amount += Number(row.profit_amount ?? (Number(row.sale_amount || 0) - Number(row.cost_amount || 0)));
    item.tax_amount += Number(row.tax_amount || 0);
    item.total_amount += Number(row.total_amount || 0);
  });
  const fromDate = document.getElementById('ledger_from').value;
  const toDate = document.getElementById('ledger_to').value;
  document.getElementById('sales_summary_warehouse_tabs').innerHTML = warehouseTabsHtml(activeSalesSummaryWarehouse);
  document.getElementById('sales_summary_warehouse').textContent =
    `KHO: ${warehouse.warehouse_code || ''} - ${warehouse.warehouse_name || ''}`;
  document.getElementById('sales_summary_range').textContent =
    `TỪ NGÀY: ${dateLabelFromIso(fromDate)} ĐẾN NGÀY: ${dateLabelFromIso(toDate)}`;
  const summaryRows = Array.from(grouped.values()).sort((a, b) => a.variant_code.localeCompare(b.variant_code));
  const totals = summaryRows.reduce((sum, row) => {
    sum.quantity += Number(row.quantity || 0);
    sum.cost_amount += Number(row.cost_amount || 0);
    sum.sale_amount += Number(row.sale_amount || 0);
    sum.profit_amount += Number(row.profit_amount || 0);
    sum.tax_amount += Number(row.tax_amount || 0);
    sum.total_amount += Number(row.total_amount || 0);
    return sum;
  }, { quantity: 0, cost_amount: 0, sale_amount: 0, profit_amount: 0, tax_amount: 0, total_amount: 0 });
  document.getElementById('sales_summary_rows').innerHTML = summaryRows
    .map((row, index) => `
      <tr data-sales-code="${escapeHtml(row.variant_code)}">
        <td class="num">${index + 1}</td>
        <td class="summary-code">${escapeHtml(row.variant_code)}</td>
        <td class="text-cell">${escapeHtml(row.product_name)}</td>
        <td>${escapeHtml(row.unit_name)}</td>
        <td class="num">${quantityLabel(row.quantity)}</td>
        <td class="num">${moneyLabel(row.cost_amount)}</td>
        <td class="num">${moneyLabel(row.sale_amount)}</td>
        <td class="num">${signedMoneyLabel(row.profit_amount)}</td>
        <td class="num">${marginPercentLabelFromAmounts(row.sale_amount, row.cost_amount, activeSalesSummaryWarehouse)}</td>
        <td class="num">${moneyLabel(row.tax_amount)}</td>
        <td class="num">${moneyLabel(row.total_amount)}</td>
      </tr>
    `).join('');
  document.getElementById('sales_summary_footer').innerHTML = summaryRows.length ? `
    <tr>
      <td colspan="4">Tổng cộng:</td>
      <td class="num">${quantityLabel(totals.quantity)}</td>
      <td class="num">${moneyLabel(totals.cost_amount)}</td>
      <td class="num">${moneyLabel(totals.sale_amount)}</td>
      <td class="num">${signedMoneyLabel(totals.profit_amount)}</td>
      <td class="num">${marginPercentLabelFromAmounts(totals.sale_amount, totals.cost_amount, activeSalesSummaryWarehouse)}</td>
      <td class="num">${moneyLabel(totals.tax_amount)}</td>
      <td class="num">${moneyLabel(totals.total_amount)}</td>
    </tr>
  ` : '';
}

function showSalesDetail(variantCode) {
  const warehouse = warehouseByCode(activeSalesSummaryWarehouse);
  const fromDate = document.getElementById('ledger_from').value;
  const toDate = document.getElementById('ledger_to').value;
  const rows = salesRowsForWarehouse(activeSalesSummaryWarehouse)
    .filter(row => (row.variant_code || '') === variantCode)
    .sort((a, b) => `${a.invoice_date_iso || ''}|${a.invoice_no || ''}|${a.row_number || 0}`.localeCompare(`${b.invoice_date_iso || ''}|${b.invoice_no || ''}|${b.row_number || 0}`));
  const first = rows[0] || {};
  const totals = rows.reduce((sum, row) => {
    sum.quantity += Number(row.quantity || 0);
    sum.cost_amount += Number(row.cost_amount || 0);
    sum.sale_amount += Number(row.sale_amount || 0);
    sum.profit_amount += Number(row.profit_amount ?? (Number(row.sale_amount || 0) - Number(row.cost_amount || 0)));
    sum.tax_amount += Number(row.tax_amount || 0);
    sum.total_amount += Number(row.total_amount || 0);
    return sum;
  }, { quantity: 0, cost_amount: 0, sale_amount: 0, profit_amount: 0, tax_amount: 0, total_amount: 0 });
  document.getElementById('sales_detail_title').textContent =
    `${variantCode || ''} - ${first.product_name || ''}`;
  document.getElementById('sales_detail_subtitle').textContent =
    `KHO: ${warehouse.warehouse_code || ''} - ${warehouse.warehouse_name || ''} | Từ ngày ${dateLabelFromIso(fromDate)} đến ngày ${dateLabelFromIso(toDate)} | ${rows.length} dòng`;
  document.getElementById('sales_detail_rows').innerHTML = rows.map(row => `
    <tr class="ledger-row-out ledger-warehouse-${escapeHtml((row.warehouse_code || '').toLowerCase())}">
      <td>${escapeHtml(row.invoice_date || dateLabelFromIso(row.invoice_date_iso))}</td>
      <td>${escapeHtml(row.invoice_no || '')}</td>
      <td class="text-cell">${escapeHtml(row.customer || '')}</td>
      <td>${escapeHtml(row.warehouse_code || '')}</td>
      <td class="summary-code">${escapeHtml(row.variant_code || '')}</td>
      <td class="text-cell">${escapeHtml(row.product_name || '')}</td>
      <td>${escapeHtml(row.unit_name || '')}</td>
      <td class="num">${quantityLabel(row.quantity)}</td>
      <td class="num">${priceLabel(row.cost_price)}</td>
      <td class="num">${moneyLabel(row.cost_amount)}</td>
      <td class="num">${priceLabel(row.sale_price)}</td>
      <td class="num">${moneyLabel(row.sale_amount)}</td>
      <td class="num">${signedMoneyLabel(row.profit_amount ?? (Number(row.sale_amount || 0) - Number(row.cost_amount || 0)))}</td>
      <td class="num">${marginPercentLabelFromAmounts(row.sale_amount, row.cost_amount, row.warehouse_code)}</td>
      <td class="num">${numberLabel(row.tax_rate)}%</td>
      <td class="num">${moneyLabel(row.tax_amount)}</td>
      <td class="num">${moneyLabel(row.total_amount)}</td>
    </tr>
  `).join('');
  document.getElementById('sales_detail_footer').innerHTML = rows.length ? `
    <tr>
      <td colspan="7">Tổng cộng:</td>
      <td class="num">${quantityLabel(totals.quantity)}</td>
      <td></td>
      <td class="num">${moneyLabel(totals.cost_amount)}</td>
      <td></td>
      <td class="num">${moneyLabel(totals.sale_amount)}</td>
      <td class="num">${signedMoneyLabel(totals.profit_amount)}</td>
      <td class="num">${marginPercentLabelFromAmounts(totals.sale_amount, totals.cost_amount, activeSalesSummaryWarehouse)}</td>
      <td></td>
      <td class="num">${moneyLabel(totals.tax_amount)}</td>
      <td class="num">${moneyLabel(totals.total_amount)}</td>
    </tr>
  ` : '';
  document.getElementById('sales_detail_modal').classList.remove('hidden');
}

function hideSalesDetail() {
  document.getElementById('sales_detail_modal').classList.add('hidden');
}

function renderSalesInvoice() {
  const warehouse = warehouseByCode(activeSalesInvoiceWarehouse);
  activeSalesInvoiceWarehouse = warehouse.warehouse_code || activeSalesInvoiceWarehouse;
  const rows = salesRowsForWarehouse(activeSalesInvoiceWarehouse);
  const groups = new Map();
  rows.forEach(row => {
    const key = `${row.invoice_date_iso || ''}|${row.invoice_no || ''}|${row.customer || ''}|${row.tax_rate || 0}`;
    if (!groups.has(key)) {
      groups.set(key, {
        invoice_date: row.invoice_date || dateLabelFromIso(row.invoice_date_iso),
        invoice_date_iso: row.invoice_date_iso || '',
        invoice_no: row.invoice_no || '',
        customer: row.customer || '',
        tax_rate: Number(row.tax_rate || 0),
        rows: [],
        sale_amount: 0,
        tax_amount: 0,
        total_amount: 0
      });
    }
    const group = groups.get(key);
    group.rows.push(row);
    group.sale_amount += Number(row.sale_amount || 0);
    group.tax_amount += Number(row.tax_amount || 0);
    group.total_amount += Number(row.total_amount || 0);
  });
  const fromDate = document.getElementById('ledger_from').value;
  const toDate = document.getElementById('ledger_to').value;
  document.getElementById('sales_invoice_warehouse_tabs').innerHTML = warehouseTabsHtml(activeSalesInvoiceWarehouse);
  document.getElementById('sales_invoice_warehouse').textContent =
    `KHO: ${warehouse.warehouse_code || ''} - ${warehouse.warehouse_name || ''}`;
  document.getElementById('sales_invoice_range').textContent =
    `TỪ NGÀY: ${dateLabelFromIso(fromDate)} ĐẾN NGÀY: ${dateLabelFromIso(toDate)}`;
  document.getElementById('sales_invoice_rows').innerHTML = Array.from(groups.values())
    .sort((a, b) => `${a.invoice_date_iso}|${a.invoice_no}`.localeCompare(`${b.invoice_date_iso}|${b.invoice_no}`))
    .map(group => `
      <tr class="invoice-group-row">
        <td>${escapeHtml(group.invoice_date || '')}</td>
        <td>${escapeHtml(group.invoice_no || '')}</td>
        <td class="text-cell">${escapeHtml(group.customer || '')}<br><span>Xuất bán cho khách</span></td>
        <td></td><td>${escapeHtml(activeSalesInvoiceWarehouse)}<br>6321</td>
        <td></td><td></td><td></td><td></td><td></td>
      </tr>
      ${group.rows.map(row => `
        <tr>
          <td></td><td></td>
          <td class="text-cell">${escapeHtml(row.variant_code || '')} - ${escapeHtml(row.product_name || '')}</td>
          <td>${escapeHtml(row.unit_name || '')}</td>
          <td></td>
          <td class="num">${quantityLabel(row.quantity)}</td>
          <td class="num">${priceLabel(row.cost_price)}</td>
          <td class="num">${moneyLabel(row.cost_amount)}</td>
          <td class="num">${priceLabel(row.sale_price)}</td>
          <td class="num">${moneyLabel(row.sale_amount)}</td>
        </tr>
      `).join('')}
      <tr class="invoice-total-row"><td colspan="9">Tiền hàng:</td><td class="num">${moneyLabel(group.sale_amount)}</td></tr>
      <tr class="invoice-total-row"><td colspan="9">Tiền thuế:</td><td class="num">${moneyLabel(group.tax_amount)}</td></tr>
      <tr class="invoice-total-row"><td colspan="9">Tổng tiền thanh toán:</td><td class="num">${moneyLabel(group.total_amount)}</td></tr>
    `).join('');
}

function renderVerification() {
  document.getElementById('verification_rows').innerHTML = (currentResult?.verification || []).map(row => `
    <tr class="${row.status === 'OK' ? 'ok' : 'bad'}">
      <td>${escapeHtml(row.group)}</td>
      <td class="text-cell">${escapeHtml(row.check)}</td>
      <td class="num">${numberLabel(row.original_value)}</td>
      <td class="num">${numberLabel(row.processed_value)}</td>
      <td class="num">${numberLabel(row.difference)}</td>
      <td class="num">${numberLabel(row.tolerance)}</td>
      <td><strong>${escapeHtml(row.status)}</strong></td>
      <td class="text-cell">${escapeHtml(row.explanation)}</td>
    </tr>
  `).join('');
}

function renderLedger() {
  const context = currentLedgerContext();
  if (!context) return;
  const { fromDate, toDate, activeWarehouse } = context;
  renderLedgerShell(context);
  if (activeLedgerWarehouse === 'ALL') {
    setLedgerTableMode(true);
    renderCombinedLedger(fromDate, toDate);
    renderInventorySummary();
    renderSalesSummary();
    renderSalesInvoice();
    renderVerification();
    return;
  }
  setLedgerTableMode(false);
  const rows = [];
  const warehouseRows = [];
  (activeWarehouse.groups || []).forEach(group => {
    const html = ledgerGroupHtml(group, activeWarehouse, fromDate, toDate);
    if (html) warehouseRows.push(html);
  });
  if (warehouseRows.length) {
    rows.push(`
      <tr class="ledger-warehouse-row">
        <td colspan="10">KHO: ${escapeHtml(activeWarehouse.warehouse_code || '')} - ${escapeHtml(activeWarehouse.warehouse_name || '')}</td>
      </tr>
      ${warehouseRows.join('')}
    `);
  }
  document.getElementById('ledger_rows').innerHTML = rows.join('');
  document.getElementById('ledger_empty').classList.toggle('hidden', rows.length > 0);
  renderInventorySummary();
  renderSalesSummary();
  renderSalesInvoice();
  renderVerification();
}

async function renderLedgerWithProgress(label = 'Đang áp dụng khoảng ngày...') {
  const context = currentLedgerContext();
  if (!context) return;
  const { fromDate, toDate, activeWarehouse } = context;
  renderLedgerShell(context);
  if (activeLedgerWarehouse === 'ALL') {
    setLedgerTableMode(true);
    setLedgerLoading(true, `${label}...`);
    await sleep(0);
    renderCombinedLedger(fromDate, toDate);
    renderInventorySummary();
    renderSalesSummary();
    renderSalesInvoice();
    renderVerification();
    return;
  }
  setLedgerTableMode(false);
  const groups = activeWarehouse.groups || [];
  const chunks = [];
  const chunkSize = 25;
  let matchedGroups = 0;
  document.getElementById('ledger_rows').innerHTML = `
    <tr class="ledger-warehouse-row">
      <td colspan="10">KHO: ${escapeHtml(activeWarehouse.warehouse_code || '')} - ${escapeHtml(activeWarehouse.warehouse_name || '')}</td>
    </tr>`;
  document.getElementById('ledger_empty').classList.add('hidden');
  for (let index = 0; index < groups.length; index += chunkSize) {
    const end = Math.min(index + chunkSize, groups.length);
    setLedgerLoading(true, `${label} ${end}/${groups.length} nhóm mã VT...`);
    await sleep(0);
    const html = groups.slice(index, end)
      .map(group => ledgerGroupHtml(group, activeWarehouse, fromDate, toDate))
      .filter(Boolean)
      .join('');
    if (html) {
      matchedGroups += 1;
      chunks.push(html);
      document.getElementById('ledger_rows').insertAdjacentHTML('beforeend', html);
    }
  }
  if (!chunks.length) {
    document.getElementById('ledger_rows').innerHTML = '';
    document.getElementById('ledger_empty').classList.remove('hidden');
  }
  renderInventorySummary();
  renderSalesSummary();
  renderSalesInvoice();
  renderVerification();
}

function setLedgerLoading(isLoading, label = 'Đang áp dụng khoảng ngày...') {
  const overlay = document.getElementById('ledger_loading');
  if (!overlay) return;
  overlay.querySelector('span').textContent = label;
  overlay.classList.toggle('hidden', !isLoading);
}

function setDefaultLedgerRange(data) {
  const range = data.ledger?.date_range || {};
  document.getElementById('ledger_from').value = range.from || '';
  document.getElementById('ledger_to').value = range.to || '';
}

function activatePanel(targetId) {
  document.querySelectorAll('.tab').forEach(tab => tab.classList.toggle('active', tab.dataset.target === targetId));
  document.querySelectorAll('.panel').forEach(panel => panel.classList.toggle('hidden', panel.id !== targetId));
}

async function renderResult(data) {
  currentResult = data;
  activeLedgerWarehouse = 'KHH';
  activeInventorySummaryWarehouse = 'KHH';
  activeSalesSummaryWarehouse = 'KHH';
  activeSalesInvoiceWarehouse = 'KHH';
  document.getElementById('result').classList.remove('hidden');
  document.getElementById('download').href = `/api/download/${data.job_id}`;
  const warningPanel = document.getElementById('warnings');
  if (data.warnings.length) {
    const warningTitle = data.policy?.company_profile === 'son_phuong' ? 'Cảnh báo chọn mã / barem' : 'Cảnh báo lô ngoài khoảng';
    warningPanel.classList.remove('hidden');
    warningPanel.innerHTML = `
      <details class="warnings-tab">
        <summary>${warningTitle} <span>${numberLabel(data.warnings.length)} dòng</span></summary>
        <div class="warnings-list">
          ${data.warnings.map(item => `<div>${escapeHtml(item)}</div>`).join('')}
        </div>
      </details>`;
  } else {
    warningPanel.classList.add('hidden');
    warningPanel.innerHTML = '';
  }
  const saleOnlyCodes = data.sale_only_codes || [];
  document.getElementById('sale_only_count').textContent = `${numberLabel(saleOnlyCodes.length)} mã`;
  document.getElementById('sale_only_empty').classList.toggle('hidden', saleOnlyCodes.length > 0);
  document.getElementById('sale_only_table').classList.toggle('hidden', saleOnlyCodes.length === 0);
  document.getElementById('sale_only_rows').innerHTML = saleOnlyCodes.map(row => `
    <tr>
      <td>${escapeHtml(row.variant_code)}</td>
      <td>${escapeHtml(row.base_code)}</td>
      <td>${escapeHtml(row.product_name)}</td>
      <td class="num">${numberLabel(row.opening_quantity)}</td>
      <td class="num">${numberLabel(row.purchase_quantity)}</td>
      <td class="num">${numberLabel(row.row_count)}</td>
      <td class="num from-finished">${numberLabel(row.quantity)}</td>
      <td>${escapeHtml(row.rows.join(', '))}</td>
    </tr>
  `).join('');
  const missingBaremRows = data.missing_barem_report || [];
  document.getElementById('barem_missing_count').textContent = `${numberLabel(missingBaremRows.length)} mã`;
  document.getElementById('barem_missing_empty').classList.toggle('hidden', missingBaremRows.length > 0);
  document.getElementById('barem_missing_table').classList.toggle('hidden', missingBaremRows.length === 0);
  document.getElementById('barem_missing_rows').innerHTML = missingBaremRows.map(row => `
    <tr>
      <td>${escapeHtml(row.variant_code)}</td>
      <td>${escapeHtml(row.product_name)}</td>
      <td class="num">${numberLabel(row.quantity)}</td>
      <td>${escapeHtml(row.invoice_no)}</td>
      <td>${escapeHtml(row.invoice_date)}</td>
      <td>${escapeHtml(row.row_number)}</td>
      <td>${escapeHtml(row.reason)}</td>
    </tr>
  `).join('');
  const futureReorderRows = data.future_purchase_reorder_report || [];
  document.getElementById('future_reorder_count').textContent = `${numberLabel(futureReorderRows.length)} dÃ²ng`;
  document.getElementById('future_reorder_empty').classList.toggle('hidden', futureReorderRows.length > 0);
  document.getElementById('future_reorder_table').classList.toggle('hidden', futureReorderRows.length === 0);
  document.getElementById('future_reorder_rows').innerHTML = futureReorderRows.map(row => `
    <tr>
      <td>${escapeHtml(row.purchase_variant_code)}</td>
      <td>${escapeHtml(row.sale_variant_code)}</td>
      <td>${escapeHtml(row.product_name)}</td>
      <td>${escapeHtml(row.purchase_invoice_no)}</td>
      <td>${escapeHtml(row.purchase_original_date)}</td>
      <td>${escapeHtml(row.effective_date)}</td>
      <td>${escapeHtml(row.sale_invoice_no)}</td>
      <td>${escapeHtml(row.sale_date)}</td>
      <td class="num">${quantityLabel(row.quantity)}</td>
      <td class="num">${priceLabel(row.unit_cost)}</td>
      <td class="num">${numberLabel(row.future_reorder_days)}</td>
      <td>${escapeHtml(row.logic_note)}</td>
    </tr>
  `).join('');
  const ambiguousSteelRows = data.ambiguous_steel_rows || [];
  document.getElementById('ambiguous_steel_count').textContent = `${numberLabel(ambiguousSteelRows.length)} dòng`;
  document.getElementById('ambiguous_steel_empty').classList.toggle('hidden', ambiguousSteelRows.length > 0);
  document.getElementById('ambiguous_steel_table').classList.toggle('hidden', ambiguousSteelRows.length === 0);
  let lastAmbiguousCompany = null;
  document.getElementById('ambiguous_steel_rows').innerHTML = ambiguousSteelRows.map(row => {
    const company = row.company || '(Không có tên công ty)';
    const groupRow = company !== lastAmbiguousCompany
      ? `<tr class="group-row"><td colspan="12">${escapeHtml(company)}</td></tr>`
      : '';
    lastAmbiguousCompany = company;
    return `${groupRow}
      <tr>
        <td>${escapeHtml(company)}</td>
        <td>${escapeHtml(row.tax_code)}</td>
        <td>${escapeHtml(row.variant_code)}</td>
        <td>${escapeHtml(row.product_name)}</td>
        <td class="num">${numberLabel(row.quantity)}</td>
        <td class="num">${priceLabel(row.unit_price)}</td>
        <td>${escapeHtml(row.invoice_no)}</td>
        <td>${escapeHtml(row.invoice_date)}</td>
        <td>${escapeHtml(row.row_number)}</td>
        <td>${escapeHtml(row.detected_kind)}</td>
        <td>${escapeHtml(row.dimension)}</td>
        <td>${escapeHtml(row.reason)}</td>
      </tr>`;
  }).join('');
  setDefaultLedgerRange(data);
  activatePanel('ledger_panel');
  await sleep(0);
  await renderLedgerWithProgress('Đang dựng bảng kết quả');
}

function uploadAnalyzeJob(form) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/api/analyze-job');
    xhr.upload.addEventListener('progress', event => {
      if (!event.lengthComputable) {
        setUploadProgress(8, '1. Đang tải file lên server...');
        return;
      }
      const percent = Math.min(100, (event.loaded / event.total) * 100);
      setUploadProgress(percent, '1. Đang tải file lên server...');
    });
    xhr.addEventListener('load', () => {
      let payload = {};
      try {
        payload = JSON.parse(xhr.responseText || '{}');
      } catch (_) {
        reject(new Error('Server trả về dữ liệu không hợp lệ.'));
        return;
      }
      if (xhr.status < 200 || xhr.status >= 300) {
        reject(new Error(payload.error || 'Không tạo được tiến trình xử lý.'));
        return;
      }
      setUploadProgress(100, '1. Tải file hoàn tất');
      resolve(payload.analysis_job_id);
    });
    xhr.addEventListener('error', () => reject(new Error('Không tải được file lên server.')));
    xhr.addEventListener('abort', () => reject(new Error('Đã hủy tải file.')));
    xhr.send(form);
  });
}

async function pollAnalyzeJob(jobId) {
  while (true) {
    const response = await fetch(`/api/analyze-job/${jobId}`);
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || 'Không đọc được trạng thái xử lý.');
    if (data.status === 'queued' || Number(data.progress || 0) <= 1) {
      setCalcWaiting(`2. ${data.label || 'Đang tính toán...'}`);
    } else {
      setCalcProgress(data.progress || 0, `2. ${data.label || 'Đang tính toán...'}`);
    }
    if (data.status === 'complete') return data.result;
    if (data.status === 'error') throw new Error(data.error || data.label || 'Không tính được phân bổ.');
    await new Promise(resolve => setTimeout(resolve, 800));
  }
}

async function analyze() {
  if (isAnalyzing) {
    setStatus('Đang xử lý file. Chờ hoàn tất trước khi đổi công ty hoặc chạy lại.', true);
    return;
  }
  const purchase = document.getElementById('purchase_file');
  const sales = document.getElementById('sales_file');
  const opening = document.getElementById('opening_file');
  const barem = document.getElementById('barem_file');
  if (!purchase.files.length || !sales.files.length) {
    setStatus('Cần chọn cả hóa đơn mua vào và hóa đơn bán ra.', true);
    return;
  }
  saveMapping();
  savePolicy();
  isAnalyzing = true;
  setProcessingLocked(true);
  currentResult = null;
  document.getElementById('result').classList.add('hidden');
  document.getElementById('workflow').classList.remove('collapsed');
  document.getElementById('workflow').classList.add('processing');
  const button = document.getElementById('analyze');
  button.disabled = true;
  button.textContent = 'Đang tính phân bổ...';
  setStatus('Đang tải file...');
  resetProgress();
  const form = new FormData();
  form.append('purchase_file', purchase.files[0]);
  form.append('sales_file', sales.files[0]);
  if (opening.files.length) form.append('opening_file', opening.files[0]);
  if (barem.files.length) form.append('barem_file', barem.files[0]);
  form.append('mapping', JSON.stringify(mapping));
  form.append('policy', JSON.stringify(policy));
  try {
    const analysisJobId = await uploadAnalyzeJob(form);
    setStatus('Đã tải file. Đang tính toán...');
    setCalcProgress(1, '2. Đã nhận file. Đang bắt đầu tính...');
    const data = await pollAnalyzeJob(analysisJobId);
    await renderResult(data);
    document.getElementById('workflow').classList.add('collapsed');
    currentWizardStep = maxWizardStep;
    updateWizardStep(false);
    setStatus(`Đã tính ${data.allocation_count} dòng bán ra và ${data.stock_count} mã tồn kho.`);
    finishProgress(true);
    document.getElementById('result').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    setStatus(error.message, true);
    finishProgress(false);
  } finally {
    isAnalyzing = false;
    setProcessingLocked(false);
    document.getElementById('workflow').classList.remove('processing');
    button.disabled = false;
    button.textContent = 'Tính phân bổ kho';
  }
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function waitForUpdatedServer() {
  let sawOffline = false;
  const started = Date.now();
  while (Date.now() - started < 90000) {
    await sleep(1200);
    try {
      const response = await fetch(`/api/version?t=${Date.now()}`, { cache: 'no-store' });
      if (response.ok) {
        if (sawOffline || Date.now() - started > 8000) {
          window.location.reload();
          return;
        }
      }
    } catch (error) {
      sawOffline = true;
    }
  }
  window.location.reload();
}

async function updateExe() {
  const input = document.getElementById('update_exe_file');
  const button = document.getElementById('update_exe');
  if (!input.files.length) {
    setUpdateStatus('Chọn file .exe mới trước.', true);
    return;
  }
  const file = input.files[0];
  if (!file.name.toLowerCase().endsWith('.exe')) {
    setUpdateStatus('File cập nhật phải là .exe.', true);
    return;
  }
  const form = new FormData();
  form.append('exe_file', file);
  button.disabled = true;
  button.textContent = 'Đang cập nhật...';
  setUpdateStatus('Đang gửi file .exe mới. App sẽ tự dừng và mở lại.');
  try {
    const response = await fetch('/api/update-exe', { method: 'POST', body: form });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || 'Không cập nhật được exe.');
    setUpdateStatus(data.message || 'Đã nhận file cập nhật. Đang chờ app mở lại...');
    waitForUpdatedServer();
  } catch (error) {
    button.disabled = false;
    button.textContent = 'Cập nhật .exe';
    setUpdateStatus(error.message, true);
  }
}

loadAppVersion();
renderMapping();
document.getElementById('max_loss_percent').value = policy.max_loss_percent;
document.getElementById('max_profit_percent').value = policy.max_profit_percent;
document.getElementById('company_profile').value = policy.company_profile || 'yen_thanh';
document.getElementById('ignore_sale_suffix').checked = !!policy.ignore_sale_suffix;
document.getElementById('allow_negative_export').checked = !!policy.allow_negative_export;
document.getElementById('allow_future_purchase_reorder').checked = !!policy.allow_future_purchase_reorder;
document.getElementById('future_purchase_window_days').value = policy.future_purchase_window_days ?? '31';
document.getElementById('sp_split_pipe_box').value = policy.son_phuong_split_counts?.pipe_box || 2;
document.getElementById('sp_split_box').value = policy.son_phuong_split_counts?.box || 2;
document.getElementById('sp_split_pipe').value = policy.son_phuong_split_counts?.pipe || 2;
document.getElementById('generic_split_variance_percent').value = policy.generic_split_variance_percent ?? '';
document.getElementById('barem_tolerance_percent').value = policy.barem_tolerance_percent ?? '5';
document.getElementById('generic_min_take_quantity').value = policy.generic_min_take_quantity ?? '';
document.getElementById('generic_max_take_quantity').value = policy.generic_max_take_quantity ?? '';
document.getElementById('max_loss_percent').addEventListener('input', renderPolicyLabelSafe);
document.getElementById('max_profit_percent').addEventListener('input', renderPolicyLabelSafe);
document.getElementById('company_profile').addEventListener('change', () => {
  renderPolicyLabelSafe();
  savePolicy();
  renderPurchaseClassificationPreview();
  setStatus('Đã chọn công ty/logic xử lý. Tiếp theo chọn file và bấm Tính phân bổ kho.');
});
document.getElementById('ignore_sale_suffix').addEventListener('change', renderPolicyLabelSafe);
document.getElementById('allow_negative_export').addEventListener('change', renderPolicyLabelSafe);
document.getElementById('allow_future_purchase_reorder').addEventListener('change', renderPolicyLabelSafe);
['sp_split_pipe_box', 'sp_split_box', 'sp_split_pipe', 'generic_split_variance_percent', 'barem_tolerance_percent', 'generic_min_take_quantity', 'generic_max_take_quantity', 'future_purchase_window_days'].forEach(id => {
  document.getElementById(id).addEventListener('input', renderPolicyLabelSafe);
});
renderPolicyLabelSafe();
updateWizardStep(false);
['purchase', 'sales', 'opening'].forEach(kind => {
  document.getElementById(`${kind}_file`).addEventListener('change', () => previewFile(kind));
});
document.getElementById('barem_file').addEventListener('change', event => {
  const file = event.target.files?.[0];
  document.getElementById('barem_name').textContent = file ? file.name : 'Chưa chọn file';
});
document.getElementById('show_default_barem').addEventListener('click', async () => {
  try {
    await showDefaultBaremModal();
  } catch (error) {
    hideDefaultBaremModal();
    setStatus(error.message, true);
  }
});
document.getElementById('barem_modal_close').addEventListener('click', hideDefaultBaremModal);
document.getElementById('barem_modal').addEventListener('click', event => {
  if (event.target.id === 'barem_modal') hideDefaultBaremModal();
});
document.getElementById('barem_modal_tabs').addEventListener('click', event => {
  const button = event.target.closest('[data-barem-tab]');
  if (!button) return;
  activeBaremTab = button.dataset.baremTab;
  renderDefaultBaremModal();
});
document.getElementById('filter_unknown_steel').addEventListener('click', () => {
  steelDetectionUnknownOnly = !steelDetectionUnknownOnly;
  renderSteelDetectionTable();
});
document.getElementById('reset_mapping').addEventListener('click', () => {
  mapping = structuredClone(defaults);
  localStorage.setItem('inventory_mapping', JSON.stringify(mapping));
  localStorage.setItem('inventory_mapping_version', mappingVersion);
  renderMapping();
  ['purchase', 'sales', 'opening'].forEach(kind => {
    if (document.getElementById(`${kind}_file`).files.length) previewFile(kind);
  });
  renderPurchaseClassificationPreview();
});
document.getElementById('analyze').addEventListener('click', analyze);
document.getElementById('wizard_back').addEventListener('click', () => goWizardStep(-1));
document.getElementById('wizard_next').addEventListener('click', () => goWizardStep(1));
document.getElementById('update_exe').addEventListener('click', updateExe);
document.getElementById('apply_policy').addEventListener('click', () => {
  savePolicy();
  if (!document.getElementById('result').classList.contains('hidden')) {
    analyze();
    return;
  }
  setStatus('Đã lưu điều kiện nhận kho. Bấm Tính phân bổ kho để chạy dữ liệu.');
});
document.getElementById('unlimit_policy').addEventListener('click', () => {
  document.getElementById('max_loss_percent').value = '';
  document.getElementById('max_profit_percent').value = '';
  document.getElementById('ignore_sale_suffix').checked = false;
  document.getElementById('allow_negative_export').checked = true;
  document.getElementById('generic_split_variance_percent').value = '';
  document.getElementById('generic_min_take_quantity').value = '';
  document.getElementById('generic_max_take_quantity').value = '';
  savePolicy();
  if (!document.getElementById('result').classList.contains('hidden')) {
    analyze();
    return;
  }
  setStatus('Da luu che do khong gioi han. Bam Tinh phan bo kho de chay du lieu.');
});
document.getElementById('edit_stage').addEventListener('click', () => {
  document.getElementById('workflow').classList.remove('collapsed');
  document.getElementById('result').classList.add('hidden');
  currentWizardStep = 1;
  updateWizardStep(false);
  document.getElementById('workflow').scrollIntoView({ behavior: 'smooth', block: 'start' });
});
document.getElementById('apply_ledger_range').addEventListener('click', () => {
  const fromDate = document.getElementById('ledger_from').value;
  const toDate = document.getElementById('ledger_to').value;
  if (fromDate && toDate && fromDate > toDate) {
    setStatus('Khoảng ngày không hợp lệ: Từ ngày phải nhỏ hơn hoặc bằng Đến ngày.', true);
    return;
  }
  const button = document.getElementById('apply_ledger_range');
  button.disabled = true;
  setLedgerLoading(true);
  setStatus('Đang áp dụng khoảng ngày cho sổ chi tiết hàng hóa...');
  setTimeout(async () => {
    try {
      await renderLedgerWithProgress('Đang lọc dữ liệu theo khoảng ngày');
      setStatus('Đã áp dụng khoảng ngày cho sổ chi tiết hàng hóa.');
    } finally {
      setLedgerLoading(false);
      button.disabled = false;
    }
  }, 30);
});
document.getElementById('ledger_warehouse_tabs').addEventListener('click', event => {
  const button = event.target.closest('[data-warehouse]');
  if (!button) return;
  activeLedgerWarehouse = button.dataset.warehouse;
  setLedgerLoading(true, 'Đang đổi kho hiển thị...');
  setTimeout(async () => {
    await renderLedgerWithProgress('Đang đổi kho hiển thị');
    setLedgerLoading(false);
  }, 30);
});
document.getElementById('summary_warehouse_tabs').addEventListener('click', event => {
  const button = event.target.closest('[data-warehouse]');
  if (!button) return;
  activeInventorySummaryWarehouse = button.dataset.warehouse;
  renderInventorySummary();
});
document.getElementById('inventory_summary_rows').addEventListener('dblclick', event => {
  const row = event.target.closest('tr[data-summary-code]');
  if (!row) return;
  showInventoryDetail(row.dataset.summaryCode);
});
document.getElementById('inventory_detail_close').addEventListener('click', hideInventoryDetail);
document.getElementById('inventory_detail_modal').addEventListener('click', event => {
  if (event.target.id === 'inventory_detail_modal') hideInventoryDetail();
});
document.getElementById('sales_summary_warehouse_tabs').addEventListener('click', event => {
  const button = event.target.closest('[data-warehouse]');
  if (!button) return;
  activeSalesSummaryWarehouse = button.dataset.warehouse;
  renderSalesSummary();
});
document.getElementById('sales_summary_rows').addEventListener('dblclick', event => {
  const row = event.target.closest('tr[data-sales-code]');
  if (!row) return;
  showSalesDetail(row.dataset.salesCode);
});
document.getElementById('sales_detail_close').addEventListener('click', hideSalesDetail);
document.getElementById('sales_detail_modal').addEventListener('click', event => {
  if (event.target.id === 'sales_detail_modal') hideSalesDetail();
});
document.getElementById('sales_invoice_warehouse_tabs').addEventListener('click', event => {
  const button = event.target.closest('[data-warehouse]');
  if (!button) return;
  activeSalesInvoiceWarehouse = button.dataset.warehouse;
  renderSalesInvoice();
});
document.querySelectorAll('.tab').forEach(tab => tab.addEventListener('click', () => {
  activatePanel(tab.dataset.target);
  if (tab.dataset.target === 'inventory_summary_panel') renderInventorySummary();
  if (tab.dataset.target === 'sales_summary_panel') renderSalesSummary();
  if (tab.dataset.target === 'sales_invoice_panel') renderSalesInvoice();
  if (tab.dataset.target === 'verification_panel') renderVerification();
}));
