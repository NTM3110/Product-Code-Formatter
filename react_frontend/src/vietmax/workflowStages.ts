export type ProfileKey = 'son_phuong' | 'cao_thanh' | 'quang_thinh' | 'vietmax' | 'ho_guom' | 'viet_hung';
export type StageId = 0.5 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15;
export type StagePhase = 'format' | 'purchase' | 'sales' | 'generic' | 'price' | 'inventory' | 'fast' | 'estimate';

export type StageDefinition = {
  id: StageId;
  label: string;
  phase: StagePhase;
  short: string;
  disabled?: boolean;
  disabledReason?: string;
};

export type PrefixPresetStrategy = 'last_2_words' | 'last_3_mst' | '2_words_mst' | 'all_name_words';

export const profiles: Array<{ key: ProfileKey; label: string; note: string }> = [
  { key: 'son_phuong', label: 'Sơn Phương', note: 'Dùng cùng khung mua vào/bán ra; xử lý theo cấu hình riêng Sơn Phương.' },
  { key: 'cao_thanh', label: 'Cao Thành', note: 'Quy trình formatter có thêm stage lọc đơn giá.' },
  { key: 'quang_thinh', label: 'Quang Thịnh', note: 'Dùng quy trình formatter cơ bản.' },
  { key: 'vietmax', label: 'Vietmax', note: 'Mua vào, bán ra, khớp mua vào, tồn kho và FAST.' },
  { key: 'ho_guom', label: 'Hồ Gươm', note: 'Bóc tách dự toán từ Dự thầu và Chiết tính.' },
  { key: 'viet_hung', label: 'Việt Hưng', note: 'Dùng khung workflow formatter cơ bản.' },
];

export const vietmaxStages: StageDefinition[] = [
  { id: 0.5, label: 'Cấu hình form mapping', phase: 'format', short: 'Form mapping' },
  { id: 1, label: 'Tải file mua vào', phase: 'purchase', short: 'Tải mua vào' },
  { id: 2, label: 'Chọn cột / preview / trạng thái', phase: 'purchase', short: 'Chọn cột' },
  { id: 3, label: 'Công ty & prefix', phase: 'purchase', short: 'Công ty' },
  { id: 4, label: 'Review Mã VT', phase: 'purchase', short: 'Review Mã VT' },
  { id: 5, label: 'Tạo file mua vào', phase: 'purchase', short: 'Tạo mua vào' },
  { id: 6, label: 'Tải file bán ra', phase: 'sales', short: 'Tải bán ra' },
  { id: 7, label: 'Chọn cột / preview / trạng thái', phase: 'sales', short: 'Cột bán ra' },
  { id: 8, label: 'Khớp HĐ mua vào', phase: 'sales', short: 'Khớp mua vào' },
  { id: 9, label: 'Công ty & prefix', phase: 'sales', short: 'Công ty' },
  { id: 10, label: 'Review Mã VT', phase: 'sales', short: 'Review bán ra' },
  { id: 11, label: 'Tạo file bán ra', phase: 'sales', short: 'Tạo bán ra' },
  { id: 12, label: 'Phân bổ tồn kho', phase: 'inventory', short: 'Phân bổ' },
  { id: 13, label: 'Xem báo cáo tồn kho', phase: 'inventory', short: 'Báo cáo' },
  { id: 14, label: 'Xuất file phân bổ', phase: 'inventory', short: 'Xuất file' },
  { id: 15, label: 'Xuất FAST', phase: 'fast', short: 'Xuất FAST' },
];

export const commonProfileStages: StageDefinition[] = [
  { id: 0.5, label: 'Cấu hình form mapping', phase: 'format', short: 'Form mapping' },
  { id: 1, label: 'Tải file', phase: 'generic', short: 'Tải file' },
  { id: 2, label: 'Chọn cột', phase: 'generic', short: 'Chọn cột' },
  { id: 3, label: 'Công ty & hàng hóa', phase: 'generic', short: 'Công ty' },
  { id: 4, label: 'Xuất file', phase: 'generic', short: 'Xuất file' },
];

export const caoThanhStages: StageDefinition[] = [
  { id: 0.5, label: 'Cấu hình form mapping', phase: 'format', short: 'Form mapping' },
  { id: 1, label: 'Tải file bán ra', phase: 'generic', short: 'Tải file' },
  { id: 2, label: 'Chọn cột', phase: 'generic', short: 'Chọn cột' },
  { id: 3, label: 'Công ty & hàng hóa', phase: 'generic', short: 'Công ty' },
  { id: 4, label: 'Review Mã VT', phase: 'generic', short: 'Review Mã VT' },
  { id: 5, label: 'Lọc đơn giá', phase: 'price', short: 'Lọc giá' },
  { id: 6, label: 'Xuất file', phase: 'generic', short: 'Xuất file' },
];

export const estimateStages: StageDefinition[] = [
  { id: 1, label: 'Tải file dự toán', phase: 'estimate', short: 'Tải file' },
  { id: 2, label: 'Chọn sheet', phase: 'estimate', short: 'Chọn sheet' },
  { id: 3, label: 'Bóc tách / xuất file', phase: 'estimate', short: 'Bóc tách' },
];

export type ProfileCapabilities = {
  twoPhaseFrame?: boolean;
  vietmaxPurchaseMatch?: boolean;
  inventoryAllocation?: boolean;
  priceFilter?: boolean;
  estimateExtractor?: boolean;
};

export const profileCapabilities: Record<ProfileKey, ProfileCapabilities> = {
  son_phuong: { twoPhaseFrame: true },
  cao_thanh: { priceFilter: true },
  quang_thinh: {},
  vietmax: { twoPhaseFrame: true, vietmaxPurchaseMatch: true, inventoryAllocation: true },
  ho_guom: { estimateExtractor: true },
  viet_hung: {},
};

const baseTwoPhaseStageIds = new Set<StageId>([0.5, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 15]);
const vietmaxOnlyStageIds = new Set<StageId>([8, 12, 13, 14]);

export function baseTwoPhaseProfileStages(): StageDefinition[] {
  return vietmaxStages
    .filter((stage) => baseTwoPhaseStageIds.has(stage.id))
    .map((stage) => (
      vietmaxOnlyStageIds.has(stage.id)
        ? { ...stage, disabled: true, disabledReason: 'Stage này chỉ dùng cho Vietmax.' }
        : stage
    ));
}

export function isStageId(value: unknown): value is StageId {
  return typeof value === 'number' && [0.5, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15].includes(value);
}

export function isGenericProfileKey(value: ProfileKey) {
  return value === 'son_phuong' || value === 'quang_thinh' || value === 'cao_thanh' || value === 'viet_hung';
}

export function usesTwoPhaseFrame(profile: ProfileKey) {
  return Boolean(profileCapabilities[profile]?.twoPhaseFrame);
}

export function hasVietmaxPurchaseMatch(profile: ProfileKey) {
  return Boolean(profileCapabilities[profile]?.vietmaxPurchaseMatch);
}

export function hasInventoryAllocation(profile: ProfileKey) {
  return Boolean(profileCapabilities[profile]?.inventoryAllocation);
}

export function hasPriceFilter(profile: ProfileKey) {
  return Boolean(profileCapabilities[profile]?.priceFilter);
}

export function stagesForProfile(profile: ProfileKey): StageDefinition[] {
  if (profile === 'vietmax') return vietmaxStages;
  if (profile === 'ho_guom') return estimateStages;
  if (usesTwoPhaseFrame(profile)) return baseTwoPhaseProfileStages();
  if (profile === 'cao_thanh') return caoThanhStages;
  return [
    ...commonProfileStages.slice(0, 4),
    { id: 4, label: 'Review Mã VT', phase: 'generic', short: 'Review Mã VT' },
    { id: 5, label: 'Xuất file', phase: 'generic', short: 'Xuất file' },
  ];
}
