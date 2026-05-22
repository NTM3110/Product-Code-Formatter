import { HttpClient } from '@angular/common/http';
import { fakeAsync, tick } from '@angular/core/testing';
import { of } from 'rxjs';
import { AppComponent } from './app.component';

const config = {
  selected_profile: 'son_phuong',
  columns: {},
  profiles: {
    son_phuong: {},
    cao_thanh: {},
    quang_thinh: {}
  }
};

function createComponent() {
  const http = {
    get: jasmine.createSpy('get').and.returnValue(of(config)),
    post: jasmine.createSpy('post').and.returnValue(of(config))
  };
  return new AppComponent(http as unknown as HttpClient);
}

describe('AppComponent performance helpers', () => {
  it('does not show config loading before two seconds', fakeAsync(() => {
    const component = createComponent();
    const operationId = component.beginConfigOperation('Đang tải cấu hình...');

    tick(1999);

    expect(component.showConfigOperationLoading).toBeFalse();
    component.endConfigOperation(operationId);
    tick(1);
    expect(component.showConfigOperationLoading).toBeFalse();
  }));

  it('shows and clears delayed config loading after two seconds', fakeAsync(() => {
    const component = createComponent();
    const operationId = component.beginConfigOperation('Đang tải cấu hình...');

    tick(2000);

    expect(component.showConfigOperationLoading).toBeTrue();
    expect(component.configOperationLabel).toBe('Đang tải cấu hình...');
    component.endConfigOperation(operationId);
    expect(component.showConfigOperationLoading).toBeFalse();
    expect(component.configOperationLabel).toBe('');
  }));

  it('ignores stale timers from superseded config operations', fakeAsync(() => {
    const component = createComponent();
    const firstOperation = component.beginConfigOperation('Đang tải cấu hình cũ...');

    tick(1900);
    const secondOperation = component.beginConfigOperation('Đang tải cấu hình mới...');
    component.endConfigOperation(firstOperation);
    tick(199);

    expect(component.showConfigOperationLoading).toBeFalse();
    tick(1801);
    expect(component.showConfigOperationLoading).toBeTrue();
    expect(component.configOperationLabel).toBe('Đang tải cấu hình mới...');
    component.endConfigOperation(secondOperation);
    expect(component.showConfigOperationLoading).toBeFalse();
  }));

  it('skips Cao Thanh-only derived work for other profiles', () => {
    const component = createComponent();
    spyOn(component, 'refreshPriceGroups');
    spyOn(component, 'refreshMisorderGroups');
    spyOn(component, 'refreshNearPhraseGroups');
    component.selectedProfile = 'son_phuong';
    component.companies = [];

    component.refreshDerivedCodeViews();

    expect(component.refreshPriceGroups).not.toHaveBeenCalled();
    expect(component.refreshMisorderGroups).not.toHaveBeenCalled();
    expect(component.refreshNearPhraseGroups).not.toHaveBeenCalled();
    expect(component.priceConflictRows).toEqual([]);
    expect(component.misorderGroups).toEqual([]);
    expect(component.nearPhraseGroups).toEqual([]);
  });

  it('opens the shared suspect modal by computing only the default two-word section', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    const nearPhraseRefresh = spyOn(component, 'refreshNearPhraseGroups');
    const misorderRefresh = spyOn(component, 'refreshMisorderGroups');

    component.openSuspectModal();

    expect(component.showSuspectModal).toBeTrue();
    expect(component.activeSuspectSection).toBe('near_phrase');
    expect(nearPhraseRefresh).toHaveBeenCalledTimes(1);
    expect(misorderRefresh).not.toHaveBeenCalled();
  });

  it('computes only the newly selected suspect section when switching choices', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    const nearPhraseRefresh = spyOn(component, 'refreshNearPhraseGroups');
    const misorderRefresh = spyOn(component, 'refreshMisorderGroups');

    component.openSuspectModal();
    nearPhraseRefresh.calls.reset();
    misorderRefresh.calls.reset();

    component.selectSuspectSection('misorder');
    expect(misorderRefresh).toHaveBeenCalledTimes(1);
    expect(nearPhraseRefresh).not.toHaveBeenCalled();
  });

  it('caches base product code previews until invalidated', () => {
    const component = createComponent();
    const company = { mst: '123', value: 'ABC' };
    spyOn(component, 'buildCodePreview').and.callThrough();

    const first = component.productBaseCode(company, 'Sơn đen 10');
    const second = component.productBaseCode(company, 'Sơn đen 10');
    component.invalidateCodePreviewCache();
    const third = component.productBaseCode(company, 'Sơn đen 10');

    expect(second).toBe(first);
    expect(third).toBe(first);
    expect(component.buildCodePreview).toHaveBeenCalledTimes(2);
  });

  it('refreshes cached profile, word-rule, and skipped-company template values', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.wordRules = { a: 'A' };
    component.firstWordRules = { b: 'B' };
    component.companies = [
      {
        process: true,
        mst: '123',
        company: 'Active company',
        selected_products: new Set(['A']),
        all_products: [{ name: 'A' }, { name: 'B' }]
      },
      { process: false, mst: '456', company: 'Skipped company', selected_products: new Set(), all_products: [] }
    ];

    component.refreshUiDerivedState();

    expect(component.selectedProfileLabelText).toBe('Cao Thành');
    expect(component.wordRuleCountValue).toBe(2);
    expect(component.skippedCompanyList.length).toBe(2);
    expect(component.skippedCompanyList.some(item => item.kind === 'company')).toBeTrue();
    expect(component.skippedCompanyList.some(item => item.kind === 'product' && item.productName === 'B')).toBeTrue();
  });

  it('stores skipped products instead of selected products in profile snapshot', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.companies = [
      {
        mst: '123',
        company: 'Active company',
        process: true,
        value: 'AC',
        default_prefix: 'AC',
        selected_products: new Set(['A']),
        all_products: [{ name: 'A' }, { name: 'B' }]
      }
    ];

    const snapshot = component.currentProfileSnapshot();

    expect(snapshot.selected_products['123']).toEqual(['B']);
  });

  it('restores selected products by excluding saved skipped products', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.config = {
      ...config,
      profiles: {
        ...config.profiles,
        cao_thanh: {
          ...component.emptyProfileState('cao_thanh'),
          selected_products: { '123': ['B'] }
        }
      }
    } as any;
    component.companies = [
      {
        mst: '123',
        company: 'Active company',
        process: true,
        selected_products: new Set<string>(),
        all_products: [{ name: 'A' }, { name: 'B' }]
      }
    ];

    component.applySavedProfileToCompanies();

    expect(Array.from(component.companies[0].selected_products)).toEqual(['A']);
  });

  it('normalizes Cao Thanh Côn reducer products to include implied thu', () => {
    const component = createComponent();
    const company = { mst: '123', value: 'CT' };
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;

    const explicitThu = component.buildCodePreview(company, 'Côn thu ren inox 304 DN32/25');
    const impliedThu = component.buildCodePreview(company, 'Côn ren inox 304 DN32/25');

    expect(explicitThu).toBe('CONTHURI304DN32/25');
    expect(impliedThu).toBe('CONTHURI304DN32/25');
  });

  it('groups price filtering by final customized code', () => {
    const component = createComponent();
    const company = {
      mst: '123',
      company: 'Cao Thành',
      process: true,
      value: 'CT',
      selected_products: new Set(['Ống inox 1', 'Ống inox 2']),
      all_products: [
        {
          name: 'Ống inox 1',
          priceRows: [
            { excelRow: 1, name: 'Ống inox 1', price: 100 },
            { excelRow: 2, name: 'Ống inox 1', price: 120 }
          ]
        },
        {
          name: 'Ống inox 2',
          priceRows: [
            { excelRow: 3, name: 'Ống inox 2', price: 104 },
            { excelRow: 4, name: 'Ống inox 2', price: 126 }
          ]
        }
      ]
    };
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [company];
    component.manualCodeOverrides = {
      '123|||Ống inox 1': 'ONGA',
      '123|||Ống inox 2': 'ONGA'
    };

    const rows = component.buildPriceConflictRows();

    expect(rows.length).toBe(1);
    expect(rows[0].code).toBe('ONGA');
    expect(rows[0].priceRowCount).toBe(4);
  });

  it('prefers misorder canonical code from stronger word-rule matches', () => {
    const component = createComponent();
    const company = {
      mst: '123',
      company: 'Cao Thành',
      process: true,
      value: 'YT',
      selected_products: new Set(['Bulong inox ren suốt 304 M12x70', 'Bulong ren suốt inox 304 M12x70']),
      all_products: [
        { name: 'Bulong inox ren suốt 304 M12x70' },
        { name: 'Bulong ren suốt inox 304 M12x70' }
      ]
    };
    component.selectedProfile = 'cao_thanh';
    component.companies = [company];
    component.firstWordRules = { 'Bulong inox': 'BULONGI' };
    component.wordRules = { 'suốt': 'S', 'ren': 'R', 'inox': 'I' };

    component.refreshMisorderGroups();

    expect(component.misorderGroups.length).toBe(1);
    expect(component.misorderCanonicalCodes[component.misorderGroups[0].key]).toContain('BULONGI');
  });

  it('keeps price bucket details lazy until expanded', () => {
    const component = createComponent();
    const row = {
      key: 'price-code|||ABC',
      filterPercent: 5,
      sourceRows: [
        { key: '1', excelRow: 1, company: { company: 'A' }, product: { name: 'P1' }, name: 'P1', price: 100 },
        { key: '2', excelRow: 2, company: { company: 'A' }, product: { name: 'P2' }, name: 'P2', price: 102 },
        { key: '3', excelRow: 3, company: { company: 'A' }, product: { name: 'P3' }, name: 'P3', price: 130 }
      ]
    };

    const buckets = component.buildPriceBuckets(row);

    expect(buckets[0].details).toBeNull();
    component.togglePriceBucket(buckets[0]);
    expect(component.expandedPriceBuckets[buckets[0].key]).toBeTrue();
    expect(component.priceBucketDetails(buckets[0]).length).toBeGreaterThan(0);
  });

  it('calculates profit loss percent from adjusted bucket average', () => {
    const component = createComponent();
    const bucket = {
      key: 'bucket-1',
      label: 'Nhóm 1',
      count: 2,
      min: 100,
      max: 110,
      averagePrice: 105,
      marginPercent: 10,
      adjustedAverage: component.priceBaseline(105, 10),
      rows: [
        { key: '1', excelRow: '1', company: { company: 'A' }, product: { name: 'P1' }, name: 'P1', price: 100 },
        { key: '2', excelRow: '2', company: { company: 'A' }, product: { name: 'P2' }, name: 'P2', price: 110 }
      ],
      details: null
    };

    const details = component.priceBucketDetails(bucket as any);

    expect(details[0].deltaPercent).toBeGreaterThan(0);
    expect(details[1].deltaPercent).toBeGreaterThan(details[0].deltaPercent);
  });

  it('applies adjustment percent to all buckets across all codes', () => {
    const component = createComponent();
    component.priceAdjustAllPercent = 7;
    component.priceConflictRows = [
      {
        buckets: [
          { averagePrice: 100, marginPercent: 0, adjustedAverage: 100, details: [{ key: 'a' }], rows: [] },
          { averagePrice: 200, marginPercent: 2, adjustedAverage: 196, details: [{ key: 'b' }], rows: [] }
        ]
      },
      {
        buckets: [
          { averagePrice: 150, marginPercent: 3, adjustedAverage: 145.5, details: [{ key: 'c' }], rows: [] }
        ]
      }
    ];

    component.applyPriceAdjustPercentToAll();

    expect(component.priceConflictRows[0].buckets[0].marginPercent).toBe(7);
    expect(component.priceConflictRows[0].buckets[1].marginPercent).toBe(7);
    expect(component.priceConflictRows[1].buckets[0].marginPercent).toBe(7);
    expect(component.priceConflictRows[0].buckets[0].details).toBeNull();
  });

  it('applies filter percent to all price rows', () => {
    const component = createComponent();
    component.priceFilterAllPercent = 6;
    component.priceConflictRows = [
      {
        filterPercent: 8,
        sourceRows: [
          { key: '1', excelRow: 1, company: { company: 'A' }, product: { name: 'P1' }, name: 'P1', price: 100 },
          { key: '2', excelRow: 2, company: { company: 'A' }, product: { name: 'P2' }, name: 'P2', price: 102 },
          { key: '3', excelRow: 3, company: { company: 'A' }, product: { name: 'P3' }, name: 'P3', price: 130 }
        ],
        buckets: []
      },
      {
        filterPercent: 8,
        sourceRows: [
          { key: '4', excelRow: 4, company: { company: 'B' }, product: { name: 'Q1' }, name: 'Q1', price: 200 },
          { key: '5', excelRow: 5, company: { company: 'B' }, product: { name: 'Q2' }, name: 'Q2', price: 201 },
          { key: '6', excelRow: 6, company: { company: 'B' }, product: { name: 'Q3' }, name: 'Q3', price: 260 }
        ],
        buckets: []
      }
    ];

    component.applyPriceFilterPercentToAll();

    expect(component.priceConflictRows[0].filterPercent).toBe(6);
    expect(component.priceConflictRows[1].filterPercent).toBe(6);
  });

  it('applies adjustment percent only within one code and bucket edits still override later', () => {
    const component = createComponent();
    const row = {
      bulkAdjustPercent: 5,
      buckets: [
        { key: 'b1', label: 'Nhóm 1', count: 1, min: 100, max: 100, averagePrice: 100, marginPercent: 0, adjustedAverage: 100, details: null, rows: [] },
        { key: 'b2', label: 'Nhóm 2', count: 1, min: 200, max: 200, averagePrice: 200, marginPercent: 0, adjustedAverage: 200, details: null, rows: [] }
      ]
    };

    component.applyPriceAdjustPercentToRow(row);
    row.buckets[1].marginPercent = 9;
    component.onPriceBucketMarginChange(row.buckets[1]);

    expect(row.buckets[0].marginPercent).toBe(5);
    expect(row.buckets[1].marginPercent).toBe(9);
  });

  it('stores bucket adjustment percents under each price rule object', () => {
    const component = createComponent();
    component.priceConflictRows = [
      {
        code: 'YT.BULONI304M14X80',
        min: 15800,
        max: 18500,
        filterPercent: 8,
        products: [{ key: '0105011506|||Bulon inox 304 M14x80' }],
        buckets: [
          { key: 'b1', label: 'Nhóm 1', min: 15800, max: 17000, averagePrice: 16400, marginPercent: 2.5, adjustedAverage: 15990, count: 2, rows: [], details: null },
          { key: 'b2', label: 'Nhóm 2', min: 18000, max: 18500, averagePrice: 18250, marginPercent: 3, adjustedAverage: 17702.5, count: 2, rows: [], details: null }
        ]
      }
    ];

    component.applyPriceGroupRules();

    expect(component.priceGroupRules['0105011506|||Bulon inox 304 M14x80'].groups.length).toBe(2);
    expect(component.priceGroupRules['0105011506|||Bulon inox 304 M14x80'].groups[0].adjust_percent).toBe(2.5);
    expect(component.priceGroupRules['0105011506|||Bulon inox 304 M14x80'].groups[1].adjust_percent).toBe(3);
  });

  it('flags buckets with loss rows', () => {
    const component = createComponent();
    const bucket = {
      key: 'bucket-1',
      label: 'Nhóm 1',
      count: 2,
      min: 90,
      max: 110,
      averagePrice: 100,
      marginPercent: 5,
      adjustedAverage: 95,
      rows: [
        { key: '1', excelRow: '1', companyName: 'A', productName: 'P1', price: 90 },
        { key: '2', excelRow: '2', companyName: 'A', productName: 'P2', price: 110 }
      ],
      details: null
    } as any;

    expect(component.bucketHasLoss(bucket)).toBeTrue();
    expect(component.bucketLossCount(bucket)).toBe(1);
  });

  it('advances processing progress while keeping it below completion', fakeAsync(() => {
    const component = createComponent();

    component.startProcessingProgress();
    tick(1500);

    expect(component.processingProgress).toBeGreaterThan(1);
    expect(component.processingProgress).toBeLessThan(100);
    expect(component.processingProgressLabel).toBe('Đang xử lý file...');
    component.clearProcessingProgress();
    expect(component.processingProgress).toBeNull();
  }));

  it('finishes processing progress at 100 then clears it', fakeAsync(() => {
    const component = createComponent();

    component.startProcessingProgress();
    component.finishProcessingProgress();

    expect(component.processingProgress).toBe(100);
    expect(component.processingProgressLabel).toBe('Hoàn tất 100%');
    tick(900);
    expect(component.processingProgress).toBeNull();
  }));
});
