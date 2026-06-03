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
    expect(component.activeCompanyList.map(company => company.mst)).toEqual(['123']);
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
    } as unknown as typeof component.config;
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

  it('resets company state before applying saved profile values', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.config = {
      ...config,
      profiles: {
        ...config.profiles,
        cao_thanh: component.emptyProfileState('cao_thanh')
      }
    } as unknown as typeof component.config;
    component.companies = [
      {
        mst: '123',
        company: 'Active company',
        process: false,
        value: 'OLD',
        default_prefix: 'DF',
        selected_products: new Set(['A']),
        all_products: [{ name: 'A' }, { name: 'B' }]
      }
    ];

    component.applySavedProfileToCompanies();

    expect(component.companies[0].process).toBeTrue();
    expect(component.companies[0].value).toBe('DF');
    expect(Array.from(component.companies[0].selected_products)).toEqual(['A', 'B']);
  });

  it('starts step 3 in customization phase after checking companies', () => {
    const http = {
      get: jasmine.createSpy('get').and.returnValue(of(config)),
      post: jasmine.createSpy('post').and.returnValue(of({
        rows_to_process: 1,
        company_count: 1,
        companies: [
          { mst: '0105011506', safe_id: '0105011506', company: 'Tấn Phát Hà Nội', value: 'TP', default_prefix: 'TP', count: 1, all_products: [{ name: 'Bulong inox' }] }
        ]
      }))
    };
    const component = new AppComponent(http as unknown as HttpClient);

    component.checkCompanies();

    expect(component.step).toBe(3);
    expect(component.stage3Phase).toBe('customize');
  });

  it('derives prefix initials after ignoring Vietnam province names', () => {
    const component = createComponent();

    expect(component.companyNameInitials('Tấn Phát Hà Nội')).toBe('TP');
    expect(component.companyNameInitials('Công ty TNHH Minh Long Đà Nẵng')).toBe('ML');
    expect(component.companyNameInitials('Công ty TNHH Tấn Phát Huế')).toBe('TP');
    expect(component.companyNameInitials('Công ty TNHH Thương mại Dịch vụ Kỹ thuật Yến Thanh')).toBe('YT');
  });

  it('keeps business words available for prefix initials', () => {
    const component = createComponent();

    expect(component.companyNameInitials('Công ty TNHH Đầu Tư')).toBe('DT');
    expect(component.companyNameInitials('Công ty TNHH Thương Mại')).toBe('TM');
    expect(component.companyNameInitials('Công ty TNHH Sản Xuất')).toBe('SX');
    expect(component.companyNameInitials('Công ty TNHH Dịch Vụ')).toBe('DV');
  });

  it('keeps the current Stage 3.2 screen intact after saving config', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.step = 3;
    component.stage3Phase = 'price';
    component.priceConflictRows = [{ key: 'price-code|||A', code: 'A' }];
    component.priceCodeSections = [{ key: 'section', title: 'Công ty', groups: [{ key: 'group', code: 'A' }] }];

    component.saveProfileConfig();

    expect(component.step).toBe(3);
    expect(component.stage3Phase).toBe('price');
    expect(component.priceConflictRows.length).toBe(1);
    expect(component.priceCodeSections.length).toBe(1);
    expect(component.showErrorModal).toBeTrue();
  });

  it('shows stage loading when verifying prefixes from the prefix modal', fakeAsync(() => {
    const component = createComponent();

    component.verifyPrefixesWithLoading('Đang cập nhật lựa chọn...');

    expect(component.stageActionLoading).toBeTrue();
    expect(component.stageActionLoadingLabel).toBe('Đang cập nhật lựa chọn...');
    tick(0);
    tick(0);
    expect(component.stageActionLoading).toBeFalse();
  }));

  it('warns when three-digit MST suffix options collide', () => {
    const component = createComponent();
    component.prefixMstSuffixLength = 3;
    component.companies = [
      { process: true, mst: '0105011506', company: 'Tấn Phát Hà Nội', value: 'TP' },
      { process: true, mst: '0205011506', company: 'Minh Long Đà Nẵng', value: 'ML' }
    ];

    const suffixOption = component.prefixOptions(component.companies[0]).find(option => option.key === 'mst');

    expect(suffixOption?.value).toBe('506');
    expect(suffixOption?.warning).toContain('Trùng 3 số cuối MST');
  });

  it('warns before applying a prefix option that matches another selected prefix', () => {
    const component = createComponent();
    component.companies = [
      { process: true, mst: '0105011506', company: 'Tấn Phát Hà Nội', value: 'TP' },
      { process: true, mst: '0205011507', company: 'Tấn Phú Đà Nẵng', value: 'TP' }
    ];

    const currentOption = component.prefixOptions(component.companies[0]).find(option => option.key === 'current');

    expect(currentOption?.warning).toContain('Sẽ trùng TP');
  });

  it('keeps unchecked companies visible in the prefix modal lists', () => {
    const component = createComponent();
    component.companies = [
      { process: false, mst: '0105011506', company: 'Skipped', value: 'SK' },
      { process: true, mst: '0205011507', company: 'Active', value: 'AC' }
    ];

    component.sortCompanies();

    expect(component.normalCompanies.map(company => company.mst)).toContain('0105011506');
  });

  it('groups the stage 3 product list by final customized code', () => {
    const component = createComponent();
    const company = {
      mst: '123',
      company: 'Cao Thành',
      process: true,
      value: 'CT',
      selected_products: new Set(['Ống inox 1', 'Ống inox 2']),
      all_products: [
        { name: 'Ống inox 1', count: 2 },
        { name: 'Ống inox 2', count: 3 }
      ]
    };
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [company];
    component.manualCodeOverrides = {
      '123|||Ống inox 1': 'ONGA',
      '123|||Ống inox 2': 'ONGA'
    };

    component.refreshProductCodeGroups();

    expect(component.productCodeGroups.length).toBe(1);
    expect(component.productCodeGroups[0].code).toBe('ONGA');
    expect(component.productCodeGroups[0].productCount).toBe(2);
    expect(component.productCodeGroups[0].rowCount).toBe(5);
  });

  it('groups stage 3 codes by company when company prefixes are enabled', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = true;
    component.companies = [
      {
        mst: '111',
        company: 'Company One',
        process: true,
        value: 'C1',
        selected_products: new Set(['Ống inox 1']),
        all_products: [{ name: 'Ống inox 1', count: 2 }]
      },
      {
        mst: '222',
        company: 'Company Two',
        process: true,
        value: 'C2',
        selected_products: new Set(['Ống inox 1']),
        all_products: [{ name: 'Ống inox 1', count: 3 }]
      }
    ];

    component.refreshProductCodeGroups();

    expect(component.productCodeSections.length).toBe(2);
    expect(component.productCodeSections.map(section => section.title)).toContain('111 - Company One');
    expect(component.productCodeSections.map(section => section.title)).toContain('222 - Company Two');
    expect(component.productCodeSections.every(section => section.codeCount === 1)).toBeTrue();
  });

  it('keeps a single code section when company prefixes are disabled', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [
      {
        mst: '111',
        company: 'Company One',
        process: true,
        value: 'C1',
        selected_products: new Set(['Ống inox 1']),
        all_products: [{ name: 'Ống inox 1', count: 2 }]
      },
      {
        mst: '222',
        company: 'Company Two',
        process: true,
        value: 'C2',
        selected_products: new Set(['Ống inox 1']),
        all_products: [{ name: 'Ống inox 1', count: 3 }]
      }
    ];

    component.refreshProductCodeGroups();

    expect(component.productCodeSections.length).toBe(1);
    expect(component.productCodeSections[0].title).toBe('Tất cả mã VT');
    expect(component.productCodeSections[0].codeCount).toBe(1);
  });

  it('refreshes Stage 3.1 summaries and Mã VT list after updating company selection', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = true;
    component.companies = [
      {
        mst: '111',
        company: 'Company One',
        process: true,
        value: 'C1',
        selected_products: new Set(['A', 'B']),
        all_products: [{ name: 'A', count: 2 }, { name: 'B', count: 3 }]
      },
      {
        mst: '222',
        company: 'Company Two',
        process: true,
        value: 'C2',
        selected_products: new Set(['C']),
        all_products: [{ name: 'C', count: 4 }]
      }
    ];

    component.verifyPrefixes();
    expect(component.companyCount).toBe(2);
    expect(component.rowsToProcess).toBe(9);
    expect(component.productCodeSections.length).toBe(2);

    component.companies[1].process = false;
    component.companies[0].selected_products.delete('B');
    component.verifyPrefixes();

    expect(component.companyCount).toBe(1);
    expect(component.rowsToProcess).toBe(2);
    expect(component.productCodeSections.length).toBe(1);
    expect(component.productCodeSections[0].rowCount).toBe(2);
  });

  it('skips a company from Stage 3.1 and moves it to skipped items', fakeAsync(() => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = true;
    component.companies = [
      {
        mst: '111',
        company: 'Company One',
        process: true,
        value: 'C1',
        selected_products: new Set(['A']),
        all_products: [{ name: 'A', count: 2 }]
      },
      {
        mst: '222',
        company: 'Company Two',
        process: true,
        value: 'C2',
        selected_products: new Set(['B']),
        all_products: [{ name: 'B', count: 3 }]
      }
    ];

    component.verifyPrefixes();
    component.skipCompany(component.companies[1]);
    tick();
    tick();

    expect(component.companies[1].process).toBeFalse();
    expect(component.companyCount).toBe(1);
    expect(component.rowsToProcess).toBe(2);
    expect(component.productCodeSections.length).toBe(1);
    expect(component.activeCompanyList.map(company => company.mst)).toEqual(['111']);
    expect(component.skippedCompanyList.some(item => item.kind === 'company' && item.mst === '222')).toBeTrue();

    const skippedCompany = component.skippedCompanyList.find(item => item.kind === 'company' && item.mst === '222');
    component.restoreSkippedCompany(skippedCompany);
    tick();
    tick();

    expect(component.activeCompanyList.map(company => company.mst)).toEqual(['111', '222']);
    expect(component.skippedCompanyList.some(item => item.kind === 'company' && item.mst === '222')).toBeFalse();
  }));

  it('stages company customization changes and recomputes only after applying', fakeAsync(() => {
    const component = createComponent();
    component.companies = [
      { mst: '111', company: 'Company One', process: true, value: 'C1', selected_products: new Set(['A']), all_products: [{ name: 'A', count: 2 }] },
      { mst: '222', company: 'Company Two', process: true, value: 'C2', selected_products: new Set(['B']), all_products: [{ name: 'B', count: 3 }] }
    ];
    component.verifyPrefixes();
    spyOn(component, 'verifyPrefixes').and.callThrough();

    component.openSkippedModal();
    component.stageSkipCompany(component.companies[1]);

    expect(component.companies[1].process).toBeTrue();
    expect(component.customizationActiveCompanies.map(company => company.mst)).toEqual(['111']);
    expect(component.customizationSkippedItems.some(item => item.mst === '222')).toBeTrue();
    expect(component.verifyPrefixes).not.toHaveBeenCalled();

    component.applyCompanyCustomization();
    tick();
    tick();

    expect(component.companies[1].process).toBeFalse();
    expect(component.verifyPrefixes).toHaveBeenCalledTimes(1);
    expect(component.showSkippedModal).toBeFalse();
  }));

  it('refreshes Stage 3.1 sections when company prefix mode changes', fakeAsync(() => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = true;
    component.companies = [
      { mst: '111', company: 'Company One', process: true, value: 'C1', selected_products: new Set(['A']), all_products: [{ name: 'A', count: 1 }] },
      { mst: '222', company: 'Company Two', process: true, value: 'C2', selected_products: new Set(['A']), all_products: [{ name: 'A', count: 1 }] }
    ];
    component.verifyPrefixes();
    expect(component.productCodeSections.length).toBe(2);

    component.includeCompanyPrefix = false;
    component.onCompanyPrefixToggleChange();
    tick(0);
    tick(0);

    expect(component.productCodeSections.length).toBe(1);
    expect(component.productCodeSections[0].title).toBe('Tất cả mã VT');
    expect(component.stageProductCodeTotal()).toBe(1);
  }));

  it('toggles product code company sections and code groups', () => {
    const component = createComponent();
    component.includeCompanyPrefix = false;
    component.companies = [
      {
        mst: '111',
        company: 'Company One',
        process: true,
        value: 'C1',
        selected_products: new Set(['A']),
        all_products: [{ name: 'A', count: 1 }]
      }
    ];
    component.refreshProductCodeGroups();

    const section = component.productCodeSections[0];
    const group = section.groups[0];
    component.toggleProductCodeSection(section);
    component.toggleProductCodeGroup(group);

    expect(component.productCodeSections[0].expanded).toBeFalse();
    expect(component.productCodeSections[0].groups[0].expanded).toBeFalse();
  });

  it('opens product modal at the selected product and marks it for blinking', fakeAsync(() => {
    const component = createComponent();
    const company = {
      mst: '123',
      company: 'Company One',
      process: true,
      value: 'C1',
      selected_products: new Set(['A', 'B']),
      all_products: [{ name: 'A', count: 1 }, { name: 'B', count: 1 }]
    };
    spyOn(component, 'scrollToTargetProduct');

    component.openProductModal(company, 'B');

    expect(component.showProductModal).toBeTrue();
    expect(component.targetProductKey).toBe('123|||B');
    expect(component.blinkingProductKey).toBe('123|||B');
    expect(component.currentProductList[1].key).toBe('123|||B');
    tick(0);
    expect(component.scrollToTargetProduct).toHaveBeenCalled();
    tick(1600);
    expect(component.blinkingProductKey).toBe('');
  }));

  it('opens company products from customization and returns to that modal when closed', () => {
    const component = createComponent();
    const company = {
      mst: '123',
      company: 'Company One',
      process: false,
      value: 'C1',
      selected_products: new Set(['A']),
      all_products: [{ name: 'A', count: 1 }, { name: 'B', count: 1 }]
    };
    component.showSkippedModal = true;

    component.openProductsFromCompanyCustomization(company, 'B');

    expect(component.showSkippedModal).toBeFalse();
    expect(component.showProductModal).toBeTrue();
    expect(component.targetProductKey).toBe('123|||B');

    component.closeProductModal();

    expect(component.showProductModal).toBeFalse();
    expect(component.showSkippedModal).toBeTrue();
  });

  it('shows localized loading state while stage work is running', fakeAsync(() => {
    const component = createComponent();
    let done = false;

    component.runWithStageLoading('Đang cập nhật nhóm mã VT...', async () => {
      await component.yieldToBrowser();
      done = true;
    });

    expect(component.stageActionLoading).toBeTrue();
    expect(component.stageActionLoadingLabel).toBe('Đang cập nhật nhóm mã VT...');
    tick(0);
    tick(0);
    expect(done).toBeTrue();
    expect(component.stageActionLoading).toBeFalse();
  }));

  it('flags product groups with generated codes over 50 characters before trimming', () => {
    const component = createComponent();
    const longName = 'SieuDaiMotKhongCat SieuDaiHaiKhongCat kích thước rất dài DN100 PN16';
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [
      {
        mst: '123',
        company: 'Cao Thành',
        process: true,
        value: 'CT',
        selected_products: new Set([longName]),
        all_products: [{ name: longName, count: 1 }]
      }
    ];
    component.manualCodeOverrides = {
      [`123|||${longName}`]: 'CODETHATISDELIBERATELYLONGERTHANFIFTYCHARACTERSFORWARNING'
    };

    component.refreshProductCodeGroups();

    expect(component.productCodeGroups[0].hasLongCode).toBeTrue();
  });

  it('normalizes Cao Thanh Côn reducer products to include implied thu', () => {
    const component = createComponent();
    const company = { mst: '123', value: 'CT' };
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;

    const explicitThu = component.buildCodePreview(company, 'Côn thu ren inox 304 DN32/25');
    const impliedThu = component.buildCodePreview(company, 'Côn ren inox 304 DN32/25');

    expect(explicitThu).toBe('CONTHURIDN32/25');
    expect(impliedThu).toBe('CONTHURIDN32/25');
  });

  it('collapses inox and inox 304 to the same Cao Thanh code', () => {
    const component = createComponent();
    const company = { mst: '123', value: 'CT' };
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;

    const plainInox = component.buildCodePreview(company, 'Bulong inox M12x70');
    const inox304 = component.buildCodePreview(company, 'Bulong inox 304 M12x70');

    expect(plainInox).toBe('BULONGINOXM12X70');
    expect(inox304).toBe('BULONGINOXM12X70');
  });

  it('keeps standalone 304 when it is not an inox grade suffix', () => {
    const component = createComponent();
    const company = { mst: '123', value: 'CT' };
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;

    const code = component.buildCodePreview(company, 'Bulong 304 M12x70');

    expect(code).toBe('BULONG304M12X70');
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

  it('calculates price averages weighted by quantity', () => {
    const component = createComponent();
    const company = {
      mst: '123',
      company: 'Cao Thành',
      process: true,
      value: 'CT',
      selected_products: new Set(['Ống inox 1']),
      all_products: [
        {
          name: 'Ống inox 1',
          priceRows: [
            { excelRow: 1, name: 'Ống inox 1', price: 100, quantity: 1, amount: 100 },
            { excelRow: 2, name: 'Ống inox 1', price: 200, quantity: 3, amount: 600 }
          ]
        }
      ]
    };
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [company];
    component.manualCodeOverrides = { '123|||Ống inox 1': 'ONGA' };

    const rows = component.buildPriceConflictRows();

    expect(rows[0].quantity).toBe(4);
    expect(rows[0].totalAmount).toBe(700);
    expect(rows[0].average).toBe(175);
  });

  it('recalculates zero total amounts from unit price and quantity', () => {
    const component = createComponent();
    const company = {
      mst: '123',
      company: 'Cao Thành',
      process: true,
      value: 'CT',
      selected_products: new Set(['Ống inox 1']),
      all_products: [
        {
          name: 'Ống inox 1',
          priceRows: [
            { excelRow: 1, name: 'Ống inox 1', price: 9000, quantity: 40, amount: 0 },
            { excelRow: 2, name: 'Ống inox 1', price: 31200, quantity: 250, amount: 0 }
          ]
        }
      ]
    };
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [company];
    component.manualCodeOverrides = { '123|||Ống inox 1': 'ONGA' };

    const rows = component.buildPriceConflictRows();
    const details = component.priceBucketDetails(rows[0].buckets[0]);

    expect(rows[0].totalAmount).toBe(8160000);
    expect(rows[0].average).toBeCloseTo(28137.931, 3);
    expect(details[0].totalAmount).toBe(360000);
  });

  it('builds Stage 3.2 price sections by company when company prefixes are enabled', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = true;
    component.companies = [
      {
        mst: '111',
        company: 'Company One',
        process: true,
        value: 'C1',
        selected_products: new Set(['Ống inox 1']),
        all_products: [
          {
            name: 'Ống inox 1',
            priceRows: [
              { excelRow: 1, name: 'Ống inox 1', price: 100 },
              { excelRow: 2, name: 'Ống inox 1', price: 130 }
            ]
          }
        ]
      },
      {
        mst: '222',
        company: 'Company Two',
        process: true,
        value: 'C2',
        selected_products: new Set(['Ống inox 1']),
        all_products: [
          {
            name: 'Ống inox 1',
            priceRows: [
              { excelRow: 3, name: 'Ống inox 1', price: 200 },
              { excelRow: 4, name: 'Ống inox 1', price: 250 }
            ]
          }
        ]
      }
    ];
    component.manualCodeOverrides = {
      '111|||Ống inox 1': 'C1.ONGA',
      '222|||Ống inox 1': 'C2.ONGA'
    };

    component.refreshPriceGroups();

    expect(component.priceCodeSections.length).toBe(2);
    expect(component.priceCodeSections.map(section => section.title)).toContain('111 - Company One');
    expect(component.priceCodeSections.map(section => section.title)).toContain('222 - Company Two');
    expect(component.priceCodeSections.every(section => section.codeCount === 1)).toBeTrue();
    expect(component.priceCodeSections.every(section => section.groups[0].bucketCount === 2)).toBeTrue();
  });

  it('keeps Stage 3.2 price sections under all codes when company prefixes are disabled', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [
      {
        mst: '111',
        company: 'Company One',
        process: true,
        value: 'C1',
        selected_products: new Set(['Ống inox 1']),
        all_products: [
          {
            name: 'Ống inox 1',
            priceRows: [
              { excelRow: 1, name: 'Ống inox 1', price: 100 },
              { excelRow: 2, name: 'Ống inox 1', price: 102 }
            ]
          }
        ]
      },
      {
        mst: '222',
        company: 'Company Two',
        process: true,
        value: 'C2',
        selected_products: new Set(['Ống inox 1']),
        all_products: [
          {
            name: 'Ống inox 1',
            priceRows: [
              { excelRow: 3, name: 'Ống inox 1', price: 130 }
            ]
          }
        ]
      }
    ];
    component.manualCodeOverrides = {
      '111|||Ống inox 1': 'ONGA',
      '222|||Ống inox 1': 'ONGA'
    };

    component.refreshPriceGroups();

    expect(component.priceCodeSections.length).toBe(1);
    expect(component.priceCodeSections[0].title).toBe('Tất cả mã VT');
    expect(component.priceCodeSections[0].codeCount).toBe(1);
    expect(component.priceCodeSections[0].rowCount).toBe(3);
    expect(component.priceCodeSections[0].groups[0].companyDisplay).toContain('111 - Company One');
    expect(component.priceCodeSections[0].groups[0].hasLoss).toBeTrue();
  });

  it('includes single-price Mã VT in Stage 3.2 totals and main sections', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [
      {
        mst: '111',
        company: 'Company One',
        process: true,
        value: 'C1',
        selected_products: new Set(['Ống inox 1', 'Ống inox 2']),
        all_products: [
          { name: 'Ống inox 1', priceRows: [{ excelRow: 1, name: 'Ống inox 1', price: 100 }] },
          { name: 'Ống inox 2', priceRows: [{ excelRow: 2, name: 'Ống inox 2', price: 120 }, { excelRow: 3, name: 'Ống inox 2', price: 130 }] }
        ]
      }
    ];
    component.manualCodeOverrides = {
      '111|||Ống inox 1': 'ONGA',
      '111|||Ống inox 2': 'ONGB'
    };

    component.refreshPriceGroups();
    component.stage3Phase = 'price';

    expect(component.stageProductCodeTotal()).toBe(2);
    expect(component.singlePriceCodeTotal()).toBe(1);
    expect(component.multiPriceCodeTotal()).toBe(1);
    expect(component.splitPriceCodeTotal()).toBe(3);
    expect(component.priceCodeSections[0].codeCount).toBe(2);
    expect(component.priceCodeSections[0].priceSubsections.length).toBe(2);
    expect(component.priceCodeSections[0].priceSubsections[0].title).toBe('Mã VT nhiều đơn giá');
    expect(component.priceCodeSections[0].priceSubsections[0].groups[0].code).toBe('ONGB');
    expect(component.priceCodeSections[0].priceSubsections[1].title).toBe('Mã VT 1 đơn giá');
    expect(component.priceCodeSections[0].priceSubsections[1].groups[0].code).toBe('ONGA');
    expect(component.priceConflictGroups[0].rows.length).toBe(1);
  });

  it('counts Stage 3.2 base Mã VT before and after price splitting', () => {
    const component = createComponent();
    component.priceConflictRows = [
      { code: 'ABC', hasMultiplePrices: false, buckets: [{ finalCode: 'ABC' }] },
      { code: 'XYZ', hasMultiplePrices: true, buckets: [{ finalCode: 'XYZ.001' }, { finalCode: 'XYZ.002' }] }
    ];

    component.refreshPriceConflictViews();
    component.stage3Phase = 'price';

    expect(component.stageProductCodeTotal()).toBe(2);
    expect(component.singlePriceCodeTotal()).toBe(1);
    expect(component.multiPriceCodeTotal()).toBe(1);
    expect(component.splitPriceCodeTotal()).toBe(3);
  });

  it('shows final split Mã VT codes for price buckets', () => {
    const component = createComponent();
    const row: {
      key: string;
      code: string;
      filterPercent: number;
      draftFilterPercent: number;
      sourceRows: Array<Record<string, unknown>>;
      buckets: ReturnType<AppComponent['buildPriceBuckets']>;
    } = {
      key: 'price-code|||ABC',
      code: 'ABC',
      filterPercent: 5,
      draftFilterPercent: 5,
      sourceRows: [
        { key: '1', excelRow: 1, company: { company: 'A' }, product: { name: 'P1' }, name: 'P1', price: 100 },
        { key: '2', excelRow: 2, company: { company: 'A' }, product: { name: 'P2' }, name: 'P2', price: 130 }
      ],
      buckets: []
    };

    const buckets = component.buildPriceBuckets(row);

    expect(buckets.length).toBe(2);
    expect(buckets[0].finalCode).toBe('ABC.001');
    expect(buckets[1].finalCode).toBe('ABC.002');
  });

  it('does not store price group rules for single-price codes', fakeAsync(() => {
    const component = createComponent();
    component.priceConflictRows = [
      { code: 'SINGLE', hasMultiplePrices: false, min: 100, max: 100, filterPercent: 8, products: [{ key: '1|||Single' }], buckets: [{ label: 'Nhóm 1', min: 100, max: 100, averagePrice: 100, marginPercent: 7 }] },
      { code: 'MULTI', hasMultiplePrices: true, min: 100, max: 130, filterPercent: 8, products: [{ key: '1|||Multi' }], buckets: [] }
    ];

    component.applyPriceGroupRules();
    tick(0);
    tick(0);

    expect(component.priceGroupRules['1|||Single']).toBeUndefined();
    expect(component.priceGroupRules['1|||Multi']).toBeDefined();
    expect(component.priceRangeRules['SINGLE'].groups[0].adjust_percent).toBe(7);
  }));

  it('keeps single-price margin on the main Stage 3.2 page after applying filters', fakeAsync(() => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [
      {
        mst: '123',
        company: 'Cao Thành',
        process: true,
        value: 'CT',
        selected_products: new Set(['Ống inox 1']),
        all_products: [
          { name: 'Ống inox 1', priceRows: [{ excelRow: 1, name: 'Ống inox 1', price: 41000, quantity: 125, amount: 5125000 }] }
        ]
      }
    ];
    component.manualCodeOverrides = { '123|||Ống inox 1': 'ONGA' };
    component.refreshPriceGroups();
    component.priceConflictRows[0].buckets[0].draftMarginPercent = 7;

    component.applyPriceGroupRules();
    tick(0);
    tick(0);

    const bucket = component.priceCodeSections[0].groups[0].buckets[0];
    expect(component.priceGroupRules['123|||Ống inox 1']).toBeUndefined();
    expect(bucket.marginPercent).toBe(7);
    expect(bucket.adjustedAverage).toBe(38130);
  }));

  it('keeps multi-price margins on the main Stage 3.2 page after applying filters', fakeAsync(() => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [
      {
        mst: '123',
        company: 'Cao Thành',
        process: true,
        value: 'CT',
        selected_products: new Set(['Ống inox 1']),
        all_products: [
          {
            name: 'Ống inox 1',
            priceRows: [
              { excelRow: 1, name: 'Ống inox 1', price: 100, quantity: 1, amount: 100 },
              { excelRow: 2, name: 'Ống inox 1', price: 130, quantity: 1, amount: 130 }
            ]
          }
        ]
      }
    ];
    component.manualCodeOverrides = { '123|||Ống inox 1': 'ONGA' };
    component.refreshPriceGroups();
    component.priceConflictRows[0].buckets[0].draftMarginPercent = 7;
    component.priceConflictRows[0].buckets[1].draftMarginPercent = 7;

    component.applyPriceGroupRules();
    tick(0);
    tick(0);

    const buckets = component.priceCodeSections[0].groups[0].buckets;
    expect(buckets.length).toBe(2);
    expect(buckets[0].marginPercent).toBe(7);
    expect(buckets[1].marginPercent).toBe(7);
    expect(buckets[0].adjustedAverage).toBe(93);
    expect(buckets[1].adjustedAverage).toBeCloseTo(120.9, 6);
  }));

  it('applies and stores typed global margin for single and multi price codes immediately', fakeAsync(() => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [
      {
        mst: '123',
        company: 'Cao Thành',
        process: true,
        value: 'CT',
        selected_products: new Set(['Ống inox 1', 'Ống inox 2']),
        all_products: [
          { name: 'Ống inox 1', priceRows: [{ excelRow: 1, name: 'Ống inox 1', price: 100, quantity: 1, amount: 100 }] },
          {
            name: 'Ống inox 2',
            priceRows: [
              { excelRow: 2, name: 'Ống inox 2', price: 200, quantity: 1, amount: 200 },
              { excelRow: 3, name: 'Ống inox 2', price: 260, quantity: 1, amount: 260 }
            ]
          }
        ]
      }
    ];
    component.manualCodeOverrides = {
      '123|||Ống inox 1': 'SINGLE',
      '123|||Ống inox 2': 'MULTI'
    };
    component.refreshPriceGroups();
    component.priceAdjustAllPercent = 7;
    component.onPriceAdjustAllPercentChange();

    component.applyPriceAdjustPercentToAll();
    tick(0);
    tick(0);

    const multiSection = component.priceCodeSections[0].priceSubsections[0];
    const singleSection = component.priceCodeSections[0].priceSubsections[1];
    expect(multiSection.title).toBe('Mã VT nhiều đơn giá');
    expect(singleSection.title).toBe('Mã VT 1 đơn giá');
    expect(multiSection.groups[0].buckets.every((bucket: any) => bucket.marginPercent === 7)).toBeTrue();
    expect(singleSection.groups[0].buckets[0].marginPercent).toBe(7);
    expect(singleSection.groups[0].buckets[0].adjustedAverage).toBe(93);
    expect(component.currentProfileSnapshot().price_adjust_all_percent).toBe(7);
    expect(component.priceRangeRules['SINGLE'].groups[0].adjust_percent).toBe(7);
    component.refreshPriceReportRows();
    expect(component.priceReportSummaryTotals().profitRatio).toBeCloseTo(7, 6);
  }));

  it('restores global and single-price margins from saved Stage 3.2 config', fakeAsync(() => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [{
      mst: '123',
      company: 'Cao Thành',
      process: true,
      value: 'CT',
      selected_products: new Set(['Ống inox 1']),
      all_products: [
        { name: 'Ống inox 1', priceRows: [{ excelRow: 1, name: 'Ống inox 1', price: 100, quantity: 1, amount: 100 }] }
      ]
    }];
    component.manualCodeOverrides = { '123|||Ống inox 1': 'SINGLE' };
    component.refreshPriceGroups();
    component.priceAdjustAllPercent = 7;
    component.onPriceAdjustAllPercentChange();
    component.applyPriceAdjustPercentToAll();
    tick(0);
    tick(0);
    const savedProfile = component.currentProfileSnapshot();

    component.config = {
      ...config,
      profiles: {
        ...config.profiles,
        cao_thanh: savedProfile
      }
    } as unknown as typeof component.config;
    component.applyProfileConfig();
    component.refreshPriceGroups();

    expect(component.priceAdjustAllPercent).toBe(7);
    expect(component.priceConflictRows[0].buckets[0].marginPercent).toBe(7);
    expect(component.priceConflictRows[0].buckets[0].adjustedAverage).toBe(93);
  }));

  it('uses restored global margin for a single-price code without its own saved rule', () => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.priceAdjustAllPercent = 7;
    component.companies = [{
      mst: '123',
      company: 'Cao Thành',
      process: true,
      value: 'CT',
      selected_products: new Set(['Ống inox một giá']),
      all_products: [
        { name: 'Ống inox một giá', priceRows: [{ excelRow: 1, name: 'Ống inox một giá', price: 100, quantity: 2, amount: 200 }] }
      ]
    }];
    component.manualCodeOverrides = { '123|||Ống inox một giá': 'SINGLE' };

    component.refreshPriceGroups();

    expect(component.priceConflictRows[0].buckets[0].marginPercent).toBe(7);
    expect(component.priceConflictRows[0].buckets[0].adjustedAverage).toBe(93);
    component.refreshPriceReportRows();
    expect(component.priceReportSummaryRows[0].profitRatio).toBeCloseTo(7, 6);
  });

  it('uses filtered price-code count for Stage 3.2 total mã VT', () => {
    const component = createComponent();
    component.productCodeGroups = [{ key: 'code-1' }, { key: 'code-2' }, { key: 'code-3' }];
    component.priceConflictRows = [{ key: 'price-1' }];

    component.stage3Phase = 'customize';
    expect(component.stageProductCodeTotal()).toBe(3);

    component.stage3Phase = 'price';
    expect(component.stageProductCodeTotal()).toBe(1);
  });

  it('keeps expanded bucket details in sync after rebuilding price sections', () => {
    const component = createComponent();
    const bucket: Parameters<AppComponent['togglePriceBucket']>[0] = {
      key: 'price-code|||ABC|||all-products|||bucket|||1',
      label: 'Nhóm 1',
      count: 1,
      min: 100,
      max: 100,
      averagePrice: 100,
      marginPercent: 5,
      adjustedAverage: 95,
      rows: [
        { key: 'row-1', excelRow: '1', companyName: 'A', productName: 'P1', price: 100, quantity: 2 }
      ],
      details: null
    };

    component.togglePriceBucket(bucket);

    const rebuiltBucket = { ...bucket, details: null };
    expect(component.priceBucketExpanded(rebuiltBucket)).toBeTrue();
    expect(component.expandedPriceBucketDetails(rebuiltBucket).length).toBe(1);
  });

  it('refreshes Stage 3.2 price sections after applying price rules', fakeAsync(() => {
    const component = createComponent();
    component.selectedProfile = 'cao_thanh';
    component.includeCompanyPrefix = false;
    component.companies = [
      {
        mst: '123',
        company: 'Cao Thành',
        process: true,
        value: 'CT',
        selected_products: new Set(['Ống inox 1']),
        all_products: [
          {
            name: 'Ống inox 1',
            priceRows: [
              { excelRow: 1, name: 'Ống inox 1', price: 100 },
              { excelRow: 2, name: 'Ống inox 1', price: 130 }
            ]
          }
        ]
      }
    ];
    component.manualCodeOverrides = { '123|||Ống inox 1': 'ONGA' };
    component.refreshPriceGroups();
    component.priceConflictRows[0].buckets[0].draftMarginPercent = 10;
    component.onPriceBucketMarginChange(component.priceConflictRows[0].buckets[0]);

    component.applyPriceGroupRules();
    tick(0);
    tick(0);

    expect(component.priceCodeSections.length).toBe(1);
    expect(component.priceCodeSections[0].groups[0].buckets[0].marginPercent).toBe(10);
    expect(component.priceCodeSections[0].groups[0].buckets[0].adjustedAverage).toBe(90);
  }));

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

  it('keeps Stage 3.2 filter percent edits as drafts until an apply action rebuilds buckets', fakeAsync(() => {
    const component = createComponent();
    const row: {
      key: string;
      code: string;
      filterPercent: number;
      draftFilterPercent: number;
      sourceRows: Array<Record<string, unknown>>;
      buckets: ReturnType<AppComponent['buildPriceBuckets']>;
    } = {
      key: 'price-code|||ABC',
      code: 'ABC',
      filterPercent: 5,
      draftFilterPercent: 5,
      sourceRows: [
        { key: '1', excelRow: 1, company: { company: 'A' }, product: { name: 'P1' }, name: 'P1', price: 100 },
        { key: '2', excelRow: 2, company: { company: 'A' }, product: { name: 'P2' }, name: 'P2', price: 104 },
        { key: '3', excelRow: 3, company: { company: 'A' }, product: { name: 'P3' }, name: 'P3', price: 130 }
      ],
      buckets: []
    };
    row.buckets = component.buildPriceBuckets(row);
    component.priceConflictRows = [row];
    const originalBuckets = row.buckets;
    spyOn(component, 'buildPriceBuckets').and.callThrough();

    row.draftFilterPercent = 1;
    component.onPriceGroupPercentChange(row);

    expect(row.filterPercent).toBe(5);
    expect(row.buckets).toBe(originalBuckets);
    expect(component.buildPriceBuckets).not.toHaveBeenCalled();

    component.priceFilterAllPercent = 1;
    component.applyPriceFilterPercentToAll();
    tick(0);
    tick(0);

    expect(row.filterPercent).toBe(1);
    expect(row.buckets).not.toBe(originalBuckets);
    expect(component.buildPriceBuckets).toHaveBeenCalled();
  }));

  it('keeps Stage 3.2 bucket margin edits as drafts until apply commits summaries', fakeAsync(() => {
    const component = createComponent();
    const bucket: Parameters<AppComponent['onPriceBucketMarginChange']>[0] = {
      key: 'b1',
      label: 'Nhóm 1',
      count: 1,
      min: 100,
      max: 100,
      averagePrice: 100,
      marginPercent: 0,
      draftMarginPercent: 0,
      adjustedAverage: 100,
      lossCount: 0,
      hasLoss: false,
      details: [
        {
          key: 'a',
          excelRow: '1',
          stt: '1',
          invoiceNo: '',
          invoiceDate: '',
          customerCode: '',
          companyName: 'A',
          productName: 'P1',
          unit: '',
          price: 100,
          quantity: 1,
          totalAmount: 100,
          costAmount: 100,
          deltaAmount: 0,
          deltaTotal: 0,
          deltaPercent: 0
        }
      ],
      rows: []
    };
    const row = {
      bulkAdjustPercent: 5,
      buckets: [bucket]
    };
    const originalDetails = bucket.details;

    row.buckets[0].draftMarginPercent = 10;
    component.onPriceBucketMarginChange(row.buckets[0]);

    expect(row.buckets[0].marginPercent).toBe(0);
    expect(row.buckets[0].adjustedAverage).toBe(100);
    expect(row.buckets[0].details).toBe(originalDetails);

    component.applyPriceAdjustPercentToRow(row);
    tick(0);
    tick(0);

    expect(row.buckets[0].marginPercent).toBe(5);
    expect(row.buckets[0].draftMarginPercent).toBe(5);
    expect(row.buckets[0].adjustedAverage).toBe(95);
    expect(row.buckets[0].details).toBeNull();
  }));

  it('calculates cost amount and profit percent from unit price', () => {
    const component = createComponent();
    const bucket: Parameters<AppComponent['priceBucketDetails']>[0] = {
      key: 'bucket-1',
      label: 'Nhóm 1',
      count: 2,
      min: 100,
      max: 110,
      averagePrice: 105,
      marginPercent: 10,
      adjustedAverage: 90,
      rows: [
        { key: '1', excelRow: '1', companyName: 'A', productName: 'P1', price: 100, quantity: 2 },
        { key: '2', excelRow: '2', companyName: 'A', productName: 'P2', price: 110, quantity: 3 }
      ],
      details: null
    };

    const details = component.priceBucketDetails(bucket);

    expect(details[0].costAmount).toBe(180);
    expect(details[0].deltaTotal).toBe(20);
    expect(details[0].deltaPercent).toBe(10);
    expect(details[1].costAmount).toBe(270);
    expect(details[1].deltaTotal).toBe(60);
    expect(details[1].deltaPercent).toBeCloseTo(18.18, 2);
  });

  it('applies adjustment percent to all buckets across all codes', fakeAsync(() => {
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
    tick(0);
    tick(0);

    expect(component.priceConflictRows[0].buckets[0].marginPercent).toBe(7);
    expect(component.priceConflictRows[0].buckets[1].marginPercent).toBe(7);
    expect(component.priceConflictRows[1].buckets[0].marginPercent).toBe(7);
    expect(component.priceConflictRows[0].buckets[0].details).toBeNull();
  }));

  it('applies filter percent to all price rows', fakeAsync(() => {
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
    tick(0);
    tick(0);

    expect(component.priceConflictRows[0].filterPercent).toBe(6);
    expect(component.priceConflictRows[1].filterPercent).toBe(6);
  }));

  it('applies adjustment percent only within one code and bucket edits still override later', fakeAsync(() => {
    const component = createComponent();
    const buckets: Parameters<AppComponent['applyPriceAdjustPercentToBuckets']>[0] = [
      { key: 'b1', label: 'Nhóm 1', count: 1, min: 100, max: 100, averagePrice: 100, marginPercent: 0, adjustedAverage: 100, details: null, rows: [] },
      { key: 'b2', label: 'Nhóm 2', count: 1, min: 200, max: 200, averagePrice: 200, marginPercent: 0, adjustedAverage: 200, details: null, rows: [] }
    ];
    const row = {
      bulkAdjustPercent: 5,
      buckets
    };

    component.applyPriceAdjustPercentToRow(row);
    tick(0);
    tick(0);
    row.buckets[1].draftMarginPercent = 9;
    component.onPriceBucketMarginChange(row.buckets[1]);
    component.applyPriceRowBucketDrafts(row);

    expect(row.buckets[0].marginPercent).toBe(5);
    expect(row.buckets[1].marginPercent).toBe(9);
  }));

  it('stores bucket adjustment percents under each price rule object', fakeAsync(() => {
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
    tick(0);
    tick(0);

    expect(component.priceGroupRules['0105011506|||Bulon inox 304 M14x80'].groups.length).toBe(2);
    expect(component.priceGroupRules['0105011506|||Bulon inox 304 M14x80'].groups[0].adjust_percent).toBe(2.5);
    expect(component.priceGroupRules['0105011506|||Bulon inox 304 M14x80'].groups[1].adjust_percent).toBe(3);
  }));

  it('flags buckets with loss rows', () => {
    const component = createComponent();
    const bucket: Parameters<AppComponent['bucketHasLoss']>[0] = {
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
    };

    expect(component.bucketHasLoss(bucket)).toBeTrue();
    expect(component.bucketLossCount(bucket)).toBe(1);
  });

  it('summarizes only loss rows for the Stage 3.2 loss report', () => {
    const component = createComponent();
    const firstBucket: Parameters<AppComponent['refreshPriceBucketSummary']>[0] = {
      key: 'bucket-1',
      label: 'Nhóm 1',
      count: 2,
      min: 80,
      max: 100,
      averagePrice: 100,
      marginPercent: 10,
      adjustedAverage: 90,
      rows: [
        { key: '1', excelRow: '1', companyName: 'A', productName: 'P1', price: 80, quantity: 2, totalAmount: 160 },
        { key: '2', excelRow: '2', companyName: 'A', productName: 'P2', price: 100, quantity: 1, totalAmount: 100 }
      ],
      details: null
    };
    const secondBucket: Parameters<AppComponent['refreshPriceBucketSummary']>[0] = {
      key: 'bucket-2',
      label: 'Nhóm 2',
      count: 1,
      min: 180,
      max: 180,
      averagePrice: 200,
      marginPercent: 5,
      adjustedAverage: 190,
      rows: [
        { key: '3', excelRow: '3', companyName: 'A', productName: 'P3', price: 180, quantity: 3, totalAmount: 540 }
      ],
      details: null
    };
    component.refreshPriceBucketSummary(firstBucket);
    component.refreshPriceBucketSummary(secondBucket);
    component.priceConflictRows = [{ key: 'row-1', code: 'LOSS', buckets: [firstBucket, secondBucket] }];

    component.refreshPriceConflictViews();

    expect(component.priceLossReport.rowCount).toBe(2);
    expect(component.priceLossReport.revenue).toBe(700);
    expect(component.priceLossReport.lossAmount).toBe(50);
    expect(component.priceLossReport.lossPercent).toBeCloseTo((50 / 700) * 100, 6);
    expect(component.priceLossReport.rowPercent).toBeCloseTo((2 / 3) * 100, 6);
    expect(component.priceLossReport.revenuePercent).toBeCloseTo((700 / 800) * 100, 6);
    expect(component.formatLossPrice(component.priceLossReport.lossAmount)).toBe('-50');
    expect(component.formatLossPercent(component.priceLossReport.lossPercent)).toBe('-7.14%');
  });

  it('filters Stage 3.2 sections to only Mã VT with loss rows when requested', () => {
    const component = createComponent();
    component.includeCompanyPrefix = false;
    const lossBucket: Parameters<AppComponent['refreshPriceBucketSummary']>[0] = {
      key: 'loss-bucket', label: 'Nhóm 1', count: 1, min: 80, max: 80, averagePrice: 100, marginPercent: 10, adjustedAverage: 90,
      rows: [{ key: '1', excelRow: '1', companyName: 'A', productName: 'P1', price: 80, quantity: 1, totalAmount: 80 }], details: null
    };
    const okBucket: Parameters<AppComponent['refreshPriceBucketSummary']>[0] = {
      key: 'ok-bucket', label: 'Nhóm 1', count: 1, min: 120, max: 120, averagePrice: 100, marginPercent: 10, adjustedAverage: 90,
      rows: [{ key: '2', excelRow: '2', companyName: 'A', productName: 'P2', price: 120, quantity: 1, totalAmount: 120 }], details: null
    };
    component.refreshPriceBucketSummary(lossBucket);
    component.refreshPriceBucketSummary(okBucket);
    component.priceConflictRows = [
      { key: 'row-loss', code: 'LOSS', hasMultiplePrices: false, buckets: [lossBucket], sourceRows: [], products: [] },
      { key: 'row-ok', code: 'OK', hasMultiplePrices: false, buckets: [okBucket], sourceRows: [], products: [] }
    ];

    component.refreshPriceConflictViews();
    expect(component.priceCodeSections[0].groups.length).toBe(2);

    component.priceLossOnly = true;
    component.refreshPriceCodeSections();

    expect(component.priceCodeSections[0].groups.length).toBe(1);
    expect(component.priceCodeSections[0].groups[0].code).toBe('LOSS');
  });

  it('builds sales report rows sorted by date and invoice number with split codes first', () => {
    const component = createComponent();
    component.priceConflictRows = [
      {
        key: 'row-split',
        code: 'SPLIT',
        buckets: [
          {
            key: 'split-bucket-2', label: 'Nhóm 2', finalCode: 'SPLIT.002', count: 2, min: 100, max: 120, averagePrice: 110, marginPercent: 10, adjustedAverage: 90,
            rows: [
              { key: '2', excelRow: '2', stt: '2', invoiceNo: 'HD10', invoiceDate: '02/05/2024', customerCode: 'MST1', companyName: 'A', productName: 'Vật tư B', unit: 'Cái', price: 120, quantity: 1, totalAmount: 120 },
              { key: '1', excelRow: '1', stt: '1', invoiceNo: 'HD2', invoiceDate: '01/05/2024', customerCode: 'MST1', companyName: 'A', productName: 'Vật tư A', unit: 'Cái', price: 100, quantity: 2, totalAmount: 200 }
            ],
            details: null
          }
        ]
      },
      {
        key: 'row-single',
        code: 'SINGLE',
        buckets: [
          {
            key: 'single-bucket', label: 'Nhóm 1', finalCode: 'SINGLE', count: 1, min: 80, max: 80, averagePrice: 80, marginPercent: 0, adjustedAverage: 100,
            rows: [{ key: '3', excelRow: '3', stt: '3', invoiceNo: 'HD1', invoiceDate: '01/05/2024', customerCode: 'MST2', companyName: 'B', productName: 'Vật tư C', unit: 'Mét', price: 80, quantity: 1, totalAmount: 80 }],
            details: null
          }
        ]
      }
    ];

    component.refreshPriceReportRows();

    expect(component.priceReportInvoiceRows.map(row => row.invoiceNo)).toEqual(['HD1', 'HD2', 'HD10']);
    expect(component.priceReportInvoiceRows.find(row => row.invoiceNo === 'HD2')?.customerCode).toBe('MST1');
    expect(component.priceReportSummaryRows.map(row => row.finalCode)).toEqual(['SPLIT.002', 'SINGLE']);
    expect(component.priceReportSummaryRows[0].quantity).toBe(3);
    expect(component.priceReportSummaryRows[0].saleAmount).toBe(320);
    expect(component.priceReportSummaryRows[0].costAmount).toBe(270);
    expect(component.priceReportSummaryRows[0].profitAmount).toBe(50);
    expect(component.priceReportSummaryRows[0].netAmount).toBe(50);
    expect(component.priceReportSummaryRows[1].lossAmount).toBe(-20);
    expect(component.priceReportSummaryRows[1].netAmount).toBe(-20);
    expect(component.priceReportSummaryTotals().saleAmount).toBe(400);
    expect(component.priceReportSummaryTotals().costAmount).toBe(370);
    expect(component.priceReportSummaryTotals().netAmount).toBe(30);
    expect(component.priceReportSummaryTotals().profitRatio).toBe(7.5);
  });

  it('shows loading before opening filtered invoice rows from the summary action', fakeAsync(() => {
    const component = createComponent();
    component.priceReportInvoiceRows = [
      { key: 'a', summaryKey: 'split', finalCode: 'SPLIT.001', excelRow: '1', stt: '1', invoiceNo: 'HD1', invoiceDate: '01/05/2024', customerCode: 'MST1', companyName: 'A', productName: 'A', unit: 'Cái', price: 100, salePrice: 100, quantity: 1, totalAmount: 100, saleAmount: 100, costUnitPrice: 90, costAmount: 90, deltaAmount: 10, deltaTotal: 10, deltaPercent: 10, profitRatio: 10 },
      { key: 'b', summaryKey: 'single', finalCode: 'SINGLE', excelRow: '2', stt: '2', invoiceNo: 'HD2', invoiceDate: '02/05/2024', customerCode: 'MST2', companyName: 'B', productName: 'B', unit: 'Cái', price: 80, salePrice: 80, quantity: 1, totalAmount: 80, saleAmount: 80, costUnitPrice: 100, costAmount: 100, deltaAmount: -20, deltaTotal: -20, deltaPercent: -25, profitRatio: -25 }
    ];
    component.priceReportSummaryRows = component.buildPriceReportSummaryRows(component.priceReportInvoiceRows);

    component.selectPriceReportSummary(component.priceReportSummaryRows[0]);

    expect(component.priceReportTransitionLoading).toBeTrue();
    expect(component.priceReportTransitionLabel).toContain('hóa đơn liên quan');
    expect(component.activePriceReportTab).toBe('summary');
    tick(0);
    tick(0);
    expect(component.visiblePriceReportInvoiceRows().length).toBe(1);
    expect(component.visiblePriceReportInvoiceRows()[0].summaryKey).toBe('split');
    expect(component.priceReportInvoiceTitle()).toContain('SPLIT.001');
    expect(component.activePriceReportTab).toBe('invoice');
    expect(component.priceReportInvoiceTotals().netAmount).toBe(10);
    expect(component.priceReportInvoiceTotals().profitRatio).toBe(10);
    spyOn(component, 'scrollSelectedPriceReportSummaryIntoView');
    component.selectPriceReportTab('summary');
    tick(0);
    expect(component.activePriceReportTab).toBe('summary');
    expect(component.selectedPriceReportSummaryKey).toBe(component.priceReportSummaryRows[0].key);
    expect(component.scrollSelectedPriceReportSummaryIntoView).toHaveBeenCalled();
    component.selectedPriceReportSummaryKey = '';
    expect(component.visiblePriceReportInvoiceRows().length).toBe(2);
    expect(component.priceReportInvoiceTotals().netAmount).toBe(-10);
    expect(component.priceReportInvoiceTotals().profitRatio).toBeCloseTo(-5.56, 2);
  }));

  it('shows loading when opening the invoice report tab', fakeAsync(() => {
    const component = createComponent();

    component.selectPriceReportTab('invoice');

    expect(component.priceReportTransitionLoading).toBeTrue();
    expect(component.activePriceReportTab).toBe('summary');
    tick(0);
    tick(0);
    expect(component.priceReportTransitionLoading).toBeFalse();
    expect(component.activePriceReportTab).toBe('invoice');
  }));

  it('builds an Excel workbook with both sales report sheets', () => {
    const component = createComponent();
    component.priceReportInvoiceRows = [
      { key: 'a', summaryKey: 'split', finalCode: 'SPLIT.001', excelRow: '8', stt: '8', invoiceNo: 'HD1', invoiceDate: '01/05/2024', customerCode: 'MST1', companyName: 'A', productName: 'A & B', unit: 'Cái', price: 100, salePrice: 100, quantity: 1, totalAmount: 100, saleAmount: 100, costUnitPrice: 90, costAmount: 90, deltaAmount: 10, deltaTotal: 10, deltaPercent: 10, profitRatio: 10 },
      { key: 'b', summaryKey: 'single', finalCode: 'SINGLE', excelRow: '15', stt: '15', invoiceNo: 'HD2', invoiceDate: '02/05/2024', customerCode: 'MST2', companyName: 'B', productName: 'B', unit: 'Cái', price: 80, salePrice: 80, quantity: 1, totalAmount: 80, saleAmount: 80, costUnitPrice: 100, costAmount: 100, deltaAmount: -20, deltaTotal: -20, deltaPercent: -25, profitRatio: -25 }
    ];
    component.priceReportSummaryRows = component.buildPriceReportSummaryRows(component.priceReportInvoiceRows);

    const workbook = component.excelWorkbookHtml([
      { name: 'Báo cáo tổng hợp bán hàng', rows: component.priceReportSummaryExportRows() },
      { name: 'Kết xuất hóa đơn', rows: component.priceReportInvoiceExportRows() }
    ]);

    expect(workbook).toContain('<x:Name>Báo cáo tổng hợp bán hàng</x:Name>');
    expect(workbook).toContain('<x:Name>Kết xuất hóa đơn</x:Name>');
    expect(workbook).toContain('<th>Mã VT</th>');
    expect(workbook).toContain('<th>Lãi/lỗ</th>');
    expect(component.priceReportSummaryExportRows()[0]['Lãi/lỗ']).toBe(10);
    expect(component.priceReportSummaryExportRows()[0]['Lãi']).toBeUndefined();
    expect(component.priceReportSummaryExportRows()[0]['Lỗ']).toBeUndefined();
    expect(component.priceReportInvoiceExportRows().map(row => row['STT'])).toEqual([1, 2, 'TỔNG CỘNG']);
    expect(component.priceReportInvoiceExportRows()[0]['Tên công ty']).toBe('A');
    const summaryExportRows = component.priceReportSummaryExportRows();
    expect(summaryExportRows[summaryExportRows.length - 1]['Tỷ lệ lãi/lỗ'] as number).toBeCloseTo(-5.56, 2);
    expect(component.priceReportExportFileName()).toBe('bao-cao-ban-hang_bao_cao_ban_hang.xlsx');
    expect(workbook).toContain('<th>Tên công ty</th>');
    expect(workbook).toContain('<td>SPLIT.001</td>');
    expect(workbook).toContain('<td>A &amp; B</td>');
    expect(workbook).toContain('<td>HD1</td>');
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
