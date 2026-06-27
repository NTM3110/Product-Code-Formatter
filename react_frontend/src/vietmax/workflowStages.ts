export type ProfileKey = 'son_phuong' | 'cao_thanh' | 'quang_thinh' | 'vietmax';
export type StageId = 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15;
export type StagePhase = 'purchase' | 'sales' | 'generic' | 'price' | 'inventory' | 'fast';

export type StageDefinition = {
  id: StageId;
  label: string;
  phase: StagePhase;
  short: string;
};

export type PrefixPresetStrategy = 'last_2_words' | 'last_3_mst' | '2_words_mst';

export const profiles: Array<{ key: ProfileKey; label: string; note: string }> = [
  { key: 'son_phuong', label: 'Sơn Phương', note: 'Sẽ migrate sau Vietmax.' },
  { key: 'cao_thanh', label: 'Cao Thành', note: 'Sẽ migrate sau Vietmax, gồm stage lọc đơn giá.' },
  { key: 'quang_thinh', label: 'Quang Thịnh', note: 'Sẽ migrate sau Vietmax.' },
  { key: 'vietmax', label: 'Vietmax', note: 'Đang migrate trước: mua vào rồi bán ra, stage 1-11.' },
];

export const vietmaxStages: StageDefinition[] = [
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
  { id: 13, label: 'Xem báo cáo tồn kho', phase: 'inventory', short: 'Báo cáo' },
  { id: 14, label: 'Xuất file phân bổ', phase: 'inventory', short: 'Xuất file' },
  { id: 15, label: 'Xuất FAST', phase: 'fast', short: 'Xuất FAST' },
];

export const commonProfileStages: StageDefinition[] = [
  { id: 1, label: 'Tải file', phase: 'generic', short: 'Tải file' },
  { id: 2, label: 'Chọn cột', phase: 'generic', short: 'Chọn cột' },
  { id: 3, label: 'Công ty & hàng hóa', phase: 'generic', short: 'Công ty' },
  { id: 4, label: 'Xuất file', phase: 'generic', short: 'Xuất file' },
];

export const caoThanhStages: StageDefinition[] = [
  { id: 1, label: 'Tải file bán ra', phase: 'generic', short: 'Tải file' },
  { id: 2, label: 'Chọn cột', phase: 'generic', short: 'Chọn cột' },
  { id: 3, label: 'Công ty & hàng hóa', phase: 'generic', short: 'Công ty' },
  { id: 4, label: 'Review Mã VT', phase: 'generic', short: 'Review Mã VT' },
  { id: 5, label: 'Lọc đơn giá', phase: 'price', short: 'Lọc giá' },
  { id: 6, label: 'Xuất file', phase: 'generic', short: 'Xuất file' },
];

export function isStageId(value: unknown): value is StageId {
  return typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= 15;
}

export function isGenericProfileKey(value: ProfileKey) {
  return value === 'son_phuong' || value === 'quang_thinh' || value === 'cao_thanh';
}

export function stagesForProfile(profile: ProfileKey): StageDefinition[] {
  if (profile === 'vietmax') return vietmaxStages;
  if (profile === 'cao_thanh') return caoThanhStages;
  return [
    ...commonProfileStages.slice(0, 3),
    { id: 4, label: 'Review Mã VT', phase: 'generic', short: 'Review Mã VT' },
    { id: 5, label: 'Xuất file', phase: 'generic', short: 'Xuất file' },
  ];
}
