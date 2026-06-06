import math
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app import (
    APP_VERSION,
    CONFIG_PATH,
    DEFAULT_INVOICE_STATUS_COL,
    DEFAULT_INVOICE_STATUS_SKIP_VALUES,
    ICON_PATH,
    OUTPUT_DIR,
    PROFILE_LABELS,
    create_up_ban_ra_workbook,
    default_config,
    empty_profile_config,
    index_to_excel_col,
    invoice_status_options,
    load_config,
    make_code,
    make_product_part,
    normalize_phrase_list,
    normalize_rule_key,
    normalize_token,
    preview_data,
    process_workbook,
    profile_key,
    read_workbook,
    resolve_output_path,
    save_config,
    suggest_prefix,
    up_ban_ra_output_path,
    validate_payload,
    analyze,
)


class ProductCodeFormatterWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Product Code Formatter")
        self.resize(1280, 820)
        if ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(ICON_PATH)))

        self.config = load_config()
        self.source_path = None
        self.original_name = ""
        self.columns = []
        self.preview_rows = []
        self.invoice_statuses = []
        self.companies = []
        self.current_company_index = -1
        self.company_filter_dirty = False
        self.price_conflict_rows = []
        self.price_group_rules = {}
        self.price_range_rules = {}
        self.price_adjust_all_percent = 0
        self.price_rules_initialized = False
        self.price_table_rendering = False
        self.word_rules = {}
        self.first_word_rules = {}
        self.repeated_phrase_removals = []
        self.word_rule_table_rendering = False
        self.inventory_pairs = []
        self.use_default_inventory_pair = False
        self.default_inventory_pair_id = ""
        self.inventory_pair_rules = []
        self.inventory_table_rendering = False
        self.output_path_syncing = False

        self._build_ui()
        self._load_config_to_ui()

    def _build_ui(self):
        root = QWidget()
        main = QVBoxLayout(root)

        header = QHBoxLayout()
        title = QLabel(f"Product Code Formatter - Native Desktop  |  Version {APP_VERSION}")
        title.setStyleSheet("font-size: 18px; font-weight: 700;")
        header.addWidget(title, 1)
        header.addWidget(QLabel("Công ty áp dụng"))
        self.profile_combo = QComboBox()
        for key, label in PROFILE_LABELS.items():
            self.profile_combo.addItem(label, key)
        self.profile_combo.currentIndexChanged.connect(self.on_profile_changed)
        header.addWidget(self.profile_combo)
        self.save_config_btn = QPushButton("Lưu cấu hình")
        self.save_config_btn.clicked.connect(self.save_current_config)
        self.export_config_btn = QPushButton("Export cấu hình")
        self.import_config_btn = QPushButton("Import cấu hình")
        self.clear_cache_btn = QPushButton("Xóa cache profile")
        self.export_config_btn.clicked.connect(self.export_config)
        self.import_config_btn.clicked.connect(self.import_config)
        self.clear_cache_btn.clicked.connect(self.clear_profile_cache)
        header.addWidget(self.save_config_btn)
        header.addWidget(self.export_config_btn)
        header.addWidget(self.import_config_btn)
        header.addWidget(self.clear_cache_btn)
        main.addLayout(header)

        self.stage_labels = [
            "1. Tải file",
            "2. Chọn cột / preview / trạng thái",
            "3. Công ty & prefix",
            "4. Lọc đơn giá",
        ]
        stage_bar = QHBoxLayout()
        self.stage_buttons = []
        for index, label in enumerate(self.stage_labels):
            button = QPushButton(label)
            button.setCheckable(True)
            button.clicked.connect(lambda _checked=False, stage=index: self.set_stage(stage))
            self.stage_buttons.append(button)
            stage_bar.addWidget(button)
        main.addLayout(stage_bar)

        self.pages = QStackedWidget()
        main.addWidget(self.pages, 1)

        file_page = QWidget()
        file_page_layout = QVBoxLayout(file_page)
        file_group = QGroupBox("1. Tải file Excel")
        file_layout = QHBoxLayout(file_group)
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        choose_btn = QPushButton("Chọn file .xlsx/.xlsm")
        choose_btn.clicked.connect(self.choose_file)
        file_layout.addWidget(self.file_path_edit, 1)
        file_layout.addWidget(choose_btn)
        file_page_layout.addWidget(file_group)
        file_page_layout.addWidget(QLabel("Chọn file trước, sau đó bấm Tiếp tục để sang trang chọn cột."))
        file_page_layout.addStretch(1)
        self.pages.addWidget(file_page)

        mapping_page = QWidget()
        mapping_page_layout = QVBoxLayout(mapping_page)
        mapping_group = QGroupBox("2. Chọn cột và trạng thái hóa đơn")
        mapping_layout = QHBoxLayout(mapping_group)
        form = QFormLayout()
        self.company_col = self._column_combo()
        self.mst_col = self._column_combo()
        self.address_col = self._column_combo(allow_empty=True)
        self.product_col = self._column_combo()
        self.qty_col = self._column_combo(allow_empty=True)
        self.price_col = self._column_combo(allow_empty=True)
        self.output_col = self._column_combo()
        self.invoice_status_col = self._column_combo(allow_empty=True)
        self.invoice_status_col.currentIndexChanged.connect(self.refresh_invoice_statuses)
        for label, combo in [
            ("Tên đơn vị bán", self.company_col),
            ("Mã số thuế", self.mst_col),
            ("Địa chỉ", self.address_col),
            ("Tên hàng hóa", self.product_col),
            ("Số lượng", self.qty_col),
            ("Đơn giá", self.price_col),
            ("Cột xuất Mã VT", self.output_col),
            ("Trạng thái hóa đơn", self.invoice_status_col),
        ]:
            form.addRow(label, combo)
        self.output_path_edit = QLineEdit()
        self.output_path_edit.textChanged.connect(lambda text: self.sync_output_path_fields(self.output_path_edit, text))
        output_btn = QPushButton("Chọn nơi lưu file xử lý")
        output_btn.clicked.connect(self.choose_output_path)
        output_row = QHBoxLayout()
        output_row.addWidget(self.output_path_edit, 1)
        output_row.addWidget(output_btn)
        form.addRow("File kết quả xử lý", output_row)
        mapping_layout.addLayout(form, 1)

        right_mapping = QVBoxLayout()
        self.invoice_table = QTableWidget(0, 3)
        self.invoice_table.setHorizontalHeaderLabels(["Bỏ qua", "Trạng thái", "Dòng"])
        self.invoice_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        right_mapping.addWidget(QLabel("Trạng thái hóa đơn"))
        right_mapping.addWidget(self.invoice_table, 1)
        self.preview_table = QTableWidget(0, 0)
        right_mapping.addWidget(QLabel("Xem trước dữ liệu"))
        right_mapping.addWidget(self.preview_table, 2)
        mapping_layout.addLayout(right_mapping, 2)
        mapping_page_layout.addWidget(mapping_group, 1)

        inventory_group = QGroupBox("Cặp Mã kho / TK vật tư")
        inventory_layout = QVBoxLayout(inventory_group)
        inventory_help = QLabel("Cấu hình theo profile hiện tại. Nếu chỉ có một cặp, backend tự dùng cặp đó cho mọi dòng xử lý dù không bật mặc định.")
        inventory_help.setWordWrap(True)
        inventory_layout.addWidget(inventory_help)

        inventory_defaults = QHBoxLayout()
        self.use_default_inventory_pair_check = QCheckBox("Dùng cặp mặc định")
        self.use_default_inventory_pair_check.stateChanged.connect(lambda _state: self.on_inventory_default_changed())
        self.default_inventory_pair_combo = QComboBox()
        self.default_inventory_pair_combo.currentIndexChanged.connect(lambda _index: self.on_inventory_default_changed())
        self.inventory_default_note = QLabel("")
        self.inventory_default_note.setWordWrap(True)
        inventory_defaults.addWidget(self.use_default_inventory_pair_check)
        inventory_defaults.addWidget(QLabel("Cặp mặc định"))
        inventory_defaults.addWidget(self.default_inventory_pair_combo, 1)
        inventory_defaults.addWidget(self.inventory_default_note, 2)
        inventory_layout.addLayout(inventory_defaults)

        inventory_tables = QHBoxLayout()
        pair_box = QGroupBox("Danh sách cặp")
        pair_layout = QVBoxLayout(pair_box)
        self.inventory_pair_table = QTableWidget(0, 2)
        self.inventory_pair_table.setHorizontalHeaderLabels(["Mã kho", "TK vật tư"])
        self.inventory_pair_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.inventory_pair_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self._configure_word_rule_table(self.inventory_pair_table)
        pair_layout.addWidget(self.inventory_pair_table)
        inventory_tables.addWidget(pair_box, 1)

        rule_box = QGroupBox("Quy tắc gán cặp")
        rule_layout = QVBoxLayout(rule_box)
        self.inventory_rule_table = QTableWidget(0, 5)
        self.inventory_rule_table.setHorizontalHeaderLabels(["Bật", "Cột nguồn", "So sánh", "Giá trị", "Cặp gán"])
        self.inventory_rule_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.inventory_rule_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.inventory_rule_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self._configure_word_rule_table(self.inventory_rule_table)
        rule_layout.addWidget(self.inventory_rule_table)
        inventory_tables.addWidget(rule_box, 2)
        inventory_layout.addLayout(inventory_tables)

        inventory_actions = QHBoxLayout()
        self.add_inventory_pair_btn = QPushButton("Thêm cặp")
        self.add_inventory_pair_btn.clicked.connect(self.add_inventory_pair_row)
        self.delete_inventory_pair_btn = QPushButton("Xóa cặp")
        self.delete_inventory_pair_btn.clicked.connect(self.delete_inventory_pair_rows)
        self.add_inventory_rule_btn = QPushButton("Thêm quy tắc")
        self.add_inventory_rule_btn.clicked.connect(self.add_inventory_rule_row)
        self.delete_inventory_rule_btn = QPushButton("Xóa quy tắc")
        self.delete_inventory_rule_btn.clicked.connect(self.delete_inventory_rule_rows)
        self.apply_inventory_config_btn = QPushButton("Áp dụng cặp kho")
        self.apply_inventory_config_btn.clicked.connect(self.apply_inventory_config_from_tables)
        for btn in [
            self.add_inventory_pair_btn,
            self.delete_inventory_pair_btn,
            self.add_inventory_rule_btn,
            self.delete_inventory_rule_btn,
            self.apply_inventory_config_btn,
        ]:
            inventory_actions.addWidget(btn)
        inventory_actions.addStretch(1)
        inventory_layout.addLayout(inventory_actions)
        mapping_page_layout.addWidget(inventory_group)
        self.pages.addWidget(mapping_page)

        company_page = QWidget()
        company_page_layout = QVBoxLayout(company_page)
        company_group = QGroupBox("3. Công ty & prefix")
        company_group_layout = QVBoxLayout(company_group)
        controls = QHBoxLayout()
        self.include_prefix = QCheckBox("Dùng tiền tố công ty")
        self.include_prefix.setChecked(True)
        self.include_prefix.stateChanged.connect(lambda _state: self.verify_prefixes())
        controls.addWidget(self.include_prefix)
        self.apply_company_filter_btn = QPushButton("Áp dụng lọc công ty")
        self.apply_company_filter_btn.clicked.connect(self.apply_company_filter)
        self.apply_company_filter_btn.setEnabled(False)
        controls.addWidget(self.apply_company_filter_btn)
        controls.addStretch(1)
        company_group_layout.addLayout(controls)

        word_group = QGroupBox("Từ thay riêng")
        word_layout = QVBoxLayout(word_group)
        word_help = QLabel("Cấu hình theo profile hiện tại. Với Cao Thành, bảng 'Hai từ đầu tiên' giữ riêng với bảng 'Từ thứ 3 trở đi'.")
        word_help.setWordWrap(True)
        word_layout.addWidget(word_help)
        word_tables = QHBoxLayout()
        self.first_word_rule_box = QGroupBox("Hai từ đầu tiên (Cao Thành)")
        first_word_layout = QVBoxLayout(self.first_word_rule_box)
        self.first_word_rule_table = QTableWidget(0, 2)
        self.first_word_rule_table.setHorizontalHeaderLabels(["Từ / cụm từ", "Mã thay"])
        self.first_word_rule_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.first_word_rule_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self._configure_word_rule_table(self.first_word_rule_table)
        first_word_layout.addWidget(self.first_word_rule_table)
        word_tables.addWidget(self.first_word_rule_box)
        rest_word_box = QGroupBox("Từ thay riêng")
        rest_word_layout = QVBoxLayout(rest_word_box)
        self.word_rule_table = QTableWidget(0, 2)
        self.word_rule_table.setHorizontalHeaderLabels(["Từ / cụm từ", "Mã thay"])
        self.word_rule_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.word_rule_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self._configure_word_rule_table(self.word_rule_table)
        rest_word_layout.addWidget(self.word_rule_table)
        word_tables.addWidget(rest_word_box)
        repeat_word_box = QGroupBox("Cụm lặp chỉ giữ một lần")
        repeat_word_layout = QVBoxLayout(repeat_word_box)
        self.repeated_phrase_table = QTableWidget(0, 1)
        self.repeated_phrase_table.setHorizontalHeaderLabels(["Cụm từ"])
        self.repeated_phrase_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._configure_word_rule_table(self.repeated_phrase_table)
        repeat_word_layout.addWidget(self.repeated_phrase_table)
        word_tables.addWidget(repeat_word_box)
        word_layout.addLayout(word_tables)
        word_actions = QHBoxLayout()
        self.add_first_word_rule_btn = QPushButton("Thêm hai từ đầu")
        self.add_first_word_rule_btn.clicked.connect(lambda: self.add_word_rule_row("first"))
        self.delete_first_word_rule_btn = QPushButton("Xóa hai từ đầu")
        self.delete_first_word_rule_btn.clicked.connect(lambda: self.delete_word_rule_rows("first"))
        self.add_word_rule_btn = QPushButton("Thêm từ thay")
        self.add_word_rule_btn.clicked.connect(lambda: self.add_word_rule_row("rest"))
        self.delete_word_rule_btn = QPushButton("Xóa dòng từ thay")
        self.delete_word_rule_btn.clicked.connect(lambda: self.delete_word_rule_rows("rest"))
        self.add_repeated_phrase_btn = QPushButton("Thêm cụm lặp")
        self.add_repeated_phrase_btn.clicked.connect(lambda: self.add_word_rule_row("repeat"))
        self.delete_repeated_phrase_btn = QPushButton("Xóa cụm lặp")
        self.delete_repeated_phrase_btn.clicked.connect(lambda: self.delete_word_rule_rows("repeat"))
        self.apply_word_rules_btn = QPushButton("Áp dụng từ thay riêng")
        self.apply_word_rules_btn.clicked.connect(self.apply_word_rules_from_tables)
        for btn in [
            self.add_first_word_rule_btn,
            self.delete_first_word_rule_btn,
            self.add_word_rule_btn,
            self.delete_word_rule_btn,
            self.add_repeated_phrase_btn,
            self.delete_repeated_phrase_btn,
            self.apply_word_rules_btn,
        ]:
            word_actions.addWidget(btn)
        word_actions.addStretch(1)
        word_layout.addLayout(word_actions)
        company_group_layout.addWidget(word_group)
        prefix_controls = QHBoxLayout()
        for label, key in [("Áp MST", "mst"), ("Áp 2 chữ", "initials"), ("Áp 2 chữ + MST", "initials-mst")]:
            btn = QPushButton(label)
            btn.clicked.connect(lambda _checked=False, mode=key: self.apply_prefix_mode(mode))
            prefix_controls.addWidget(btn)
        prefix_controls.addStretch(1)
        company_group_layout.addLayout(prefix_controls)

        self.company_table = QTableWidget(0, 5)
        self.company_table.setHorizontalHeaderLabels(["Xử lý", "MST", "Công ty", "Dòng", "Prefix"])
        self.company_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.company_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.company_table.setSelectionMode(QTableWidget.SingleSelection)
        self.company_table.itemChanged.connect(self.on_company_item_changed)
        self.company_table.itemSelectionChanged.connect(self.on_company_selection_changed)
        company_group_layout.addWidget(self.company_table, 3)
        product_group = QGroupBox("Hàng hóa / mã VT preview")
        product_group_layout = QVBoxLayout(product_group)
        self.product_company_label = QLabel("Chọn một dòng công ty để xem danh sách hàng hóa bên dưới.")
        product_group_layout.addWidget(self.product_company_label)
        self.product_warning_label = QLabel("")
        self.product_warning_label.setStyleSheet("color: #991b1b; font-weight: 600;")
        self.product_warning_label.setVisible(False)
        product_group_layout.addWidget(self.product_warning_label)
        self.product_table = QTableWidget(0, 4)
        self.product_table.setHorizontalHeaderLabels(["Xử lý", "Tên hàng hóa", "Dòng", "Mã VT xem trước"])
        self.product_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.product_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.product_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.product_table.itemChanged.connect(self.on_product_item_changed)
        product_group_layout.addWidget(self.product_table)
        company_group_layout.addWidget(product_group, 2)
        process_output_group = QGroupBox("Lưu file xử lý")
        process_output_layout = QHBoxLayout(process_output_group)
        process_output_layout.addWidget(QLabel("File kết quả xử lý"))
        self.company_output_path_edit = QLineEdit()
        self.company_output_path_edit.textChanged.connect(lambda text: self.sync_output_path_fields(self.company_output_path_edit, text))
        process_output_layout.addWidget(self.company_output_path_edit, 1)
        process_output_btn = QPushButton("Chọn nơi lưu file xử lý")
        process_output_btn.clicked.connect(self.choose_output_path)
        process_output_layout.addWidget(process_output_btn)
        company_group_layout.addWidget(process_output_group)
        company_page_layout.addWidget(company_group, 1)
        self.pages.addWidget(company_page)

        price_page = QWidget()
        price_page_layout = QVBoxLayout(price_page)
        price_group = QGroupBox("4. Lọc đơn giá Cao Thành")
        price_group_layout = QVBoxLayout(price_group)
        price_intro = QLabel("Gộp theo Mã VT cuối cùng, chia nhóm theo độ lệch đơn giá, rồi lưu quy tắc để backend tách mã .001/.002 khi xử lý.")
        price_intro.setWordWrap(True)
        price_group_layout.addWidget(price_intro)

        price_stats = QHBoxLayout()
        self.price_total_label = QLabel("Tổng mã/dòng: 0 / 0")
        self.price_single_label = QLabel("Mã 1 giá: 0")
        self.price_multi_label = QLabel("Mã nhiều giá: 0")
        self.price_split_label = QLabel("Nhóm sau lọc: 0")
        for label in [self.price_total_label, self.price_single_label, self.price_multi_label, self.price_split_label]:
            label.setStyleSheet("font-weight: 600;")
            price_stats.addWidget(label)
        price_stats.addStretch(1)
        price_group_layout.addLayout(price_stats)

        price_controls = QHBoxLayout()
        price_controls.addWidget(QLabel("% lọc tất cả"))
        self.price_filter_all_edit = QLineEdit("8")
        self.price_filter_all_edit.setFixedWidth(80)
        price_controls.addWidget(self.price_filter_all_edit)
        self.apply_price_filter_all_btn = QPushButton("Áp dụng % lọc")
        self.apply_price_filter_all_btn.clicked.connect(self.apply_price_filter_percent_to_all)
        price_controls.addWidget(self.apply_price_filter_all_btn)
        price_controls.addSpacing(18)
        price_controls.addWidget(QLabel("% lãi tất cả"))
        self.price_adjust_all_edit = QLineEdit("0")
        self.price_adjust_all_edit.setFixedWidth(80)
        price_controls.addWidget(self.price_adjust_all_edit)
        self.apply_price_adjust_all_btn = QPushButton("Áp dụng % lãi")
        self.apply_price_adjust_all_btn.clicked.connect(self.apply_price_adjust_percent_to_all)
        price_controls.addWidget(self.apply_price_adjust_all_btn)
        self.apply_price_table_btn = QPushButton("Áp dụng chỉnh sửa bảng")
        self.apply_price_table_btn.clicked.connect(self.apply_price_table_edits)
        price_controls.addWidget(self.apply_price_table_btn)
        price_controls.addStretch(1)
        price_group_layout.addLayout(price_controls)

        self.price_table = QTableWidget(0, 11)
        self.price_table.setHorizontalHeaderLabels(["Loại", "Mã VT / mã sau lọc", "Công ty", "Hàng hóa", "Giá min", "Giá max", "Giá TB", "Dòng", "Nhóm", "% lọc", "% lãi"])
        self.price_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.price_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.price_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.price_table.itemChanged.connect(self.on_price_item_changed)
        price_group_layout.addWidget(self.price_table, 1)
        price_page_layout.addWidget(price_group, 1)
        self.pages.addWidget(price_page)

        actions = QHBoxLayout()
        self.back_btn = QPushButton("Quay lại")
        self.back_btn.clicked.connect(self.go_back)
        self.next_btn = QPushButton("Tiếp tục")
        self.next_btn.clicked.connect(self.go_next)
        self.process_btn = QPushButton("Xác nhận & Xử lý")
        self.process_btn.clicked.connect(self.process_file)
        self.process_btn.setEnabled(False)
        actions.addWidget(self.back_btn)
        actions.addWidget(self.next_btn)
        actions.addStretch(1)
        actions.addWidget(self.process_btn)
        main.addLayout(actions)

        self.status_label = QLabel("Sẵn sàng.")
        main.addWidget(self.status_label)
        self.setCentralWidget(root)
        self.current_stage = 0
        self.update_output_path_hint()
        self.update_stage_navigation()

    def _column_combo(self, allow_empty=False):
        combo = QComboBox()
        combo.setProperty("allow_empty", allow_empty)
        return combo

    def _configure_word_rule_table(self, table):
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.ExtendedSelection)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.setFixedHeight(self._word_rule_table_height(table, 3))

    def _word_rule_table_height(self, table, visible_rows):
        header_height = table.horizontalHeader().sizeHint().height()
        row_height = table.verticalHeader().defaultSectionSize()
        return header_height + (row_height * visible_rows) + (table.frameWidth() * 2) + 2

    def can_enter_stage(self, stage):
        if stage < 0 or stage >= self.pages.count():
            return False
        if stage <= 0:
            return True
        if stage == 1:
            return bool(self.source_path)
        if stage == 3:
            return self.current_profile() == "cao_thanh" and bool(self.companies)
        return bool(self.companies)

    def set_stage(self, stage):
        if not self.can_enter_stage(stage):
            self.update_stage_navigation()
            return
        if stage == 3 and self.current_stage != 3 and not self.prepare_price_stage():
            self.update_stage_navigation()
            return
        self.current_stage = stage
        self.pages.setCurrentIndex(stage)
        self.update_stage_navigation()

    def go_back(self):
        if self.current_stage > 0:
            self.set_stage(self.current_stage - 1)

    def go_next(self):
        if self.current_stage == 0:
            if not self.source_path:
                self.show_error("Hãy chọn file Excel trước.")
                return
            self.set_stage(1)
            return
        if self.current_stage == 1:
            self.check_companies()
            return
        if self.current_stage == 2 and self.current_profile() == "cao_thanh":
            self.go_to_price_stage()
            return

    def update_stage_navigation(self):
        for index, button in enumerate(self.stage_buttons):
            button.setChecked(index == self.current_stage)
            button.setEnabled(self.can_enter_stage(index))
        self.back_btn.setEnabled(self.current_stage > 0)
        next_labels = {
            0: "Tiếp tục chọn cột",
            1: "Kiểm tra & tiếp tục",
            2: "Bước tiếp theo: Lọc đơn giá" if self.current_profile() == "cao_thanh" else "Đã đến bước cuối",
            3: "Đã đến bước cuối",
        }
        self.next_btn.setText(next_labels.get(self.current_stage, "Tiếp tục"))
        next_enabled = self.current_stage < self.pages.count() - 1 and (self.current_stage != 0 or bool(self.source_path))
        if self.current_stage == 2:
            next_enabled = self.current_profile() == "cao_thanh" and bool(self.companies)
        self.next_btn.setEnabled(next_enabled)
        cao_thanh_requires_price = self.current_profile() == "cao_thanh"
        self.process_btn.setEnabled(bool(self.companies) and (not cao_thanh_requires_price or self.current_stage == 3))
        self.profile_combo.setEnabled(self.current_stage <= 1)

    def _load_config_to_ui(self):
        selected = profile_key(self.config.get("selected_profile", "son_phuong"))
        self._set_combo_data(self.profile_combo, selected)
        self.apply_profile_columns()
        self.load_word_rule_state()
        self.load_inventory_config_state()

    def current_profile(self):
        return self.profile_combo.currentData() or "son_phuong"

    def current_profile_config(self):
        return {**empty_profile_config(self.current_profile()), **(self.config.get("profiles", {}).get(self.current_profile()) or {})}

    def on_profile_changed(self):
        self.apply_profile_columns()
        self.load_word_rule_state()
        self.load_inventory_config_state()
        if self.current_stage == 3 and self.current_profile() != "cao_thanh":
            self.set_stage(2)
        if self.source_path:
            self.check_companies()

    def default_columns(self):
        if self.current_profile() == "cao_thanh":
            return {"company_col": "I", "mst_col": "J", "address_col": "K", "product_col": "M", "qty_col": "O", "price_col": "P", "output_col": "L"}
        return {"company_col": "F", "mst_col": "G", "address_col": "H", "product_col": "M", "qty_col": "O", "price_col": "", "output_col": "L"}

    def apply_profile_columns(self):
        profile_cfg = self.current_profile_config()
        columns = {**self.default_columns(), **(profile_cfg.get("columns") or {})}
        columns.setdefault("invoice_status_col", DEFAULT_INVOICE_STATUS_COL)
        self._set_mapping_values(columns)
        self.set_output_path(profile_cfg.get("output_path") or "")
        self.update_output_path_hint()
        if hasattr(self, "include_prefix"):
            self.include_prefix.setChecked(profile_cfg.get("include_company_prefix") is not False)
        if hasattr(self, "first_word_rule_box"):
            show_first_word_rules = self.current_profile() == "cao_thanh"
            self.first_word_rule_box.setVisible(show_first_word_rules)
            self.add_first_word_rule_btn.setVisible(show_first_word_rules)
            self.delete_first_word_rule_btn.setVisible(show_first_word_rules)
        self.load_price_rule_state()

    def _set_mapping_values(self, columns):
        values = {
            self.company_col: columns.get("company_col", "F"),
            self.mst_col: columns.get("mst_col", "G"),
            self.address_col: columns.get("address_col", "H"),
            self.product_col: columns.get("product_col", "M"),
            self.qty_col: columns.get("qty_col", "O"),
            self.price_col: columns.get("price_col", ""),
            self.output_col: columns.get("output_col", "L"),
            self.invoice_status_col: columns.get("invoice_status_col", DEFAULT_INVOICE_STATUS_COL),
        }
        for combo, value in values.items():
            self._set_combo_data(combo, value)

    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Chọn file Excel", "", "Excel files (*.xlsx *.xlsm)")
        if not path:
            return
        self.source_path = Path(path)
        self.original_name = self.source_path.name
        self.file_path_edit.setText(str(self.source_path))
        self.update_output_path_hint()
        try:
            _, df = read_workbook(self.source_path)
            self.columns = []
            for idx in range(df.shape[1]):
                letter = index_to_excel_col(idx)
                samples = []
                for row in range(min(6, len(df))):
                    value = df.iat[row, idx]
                    if str(value).strip() and str(value) != "nan":
                        samples.append(str(value).strip())
                label = letter + ((" - " + " | ".join(samples[:2])[:45]) if samples else "")
                self.columns.append({"letter": letter, "label": label})
            self.preview_rows = preview_data(df)
            self.refresh_column_combos()
            self.refresh_preview()
            self.refresh_invoice_statuses()
            self.companies = []
            self.current_company_index = -1
            self.company_filter_dirty = False
            self.price_conflict_rows = []
            self.price_rules_initialized = False
            self.render_companies()
            self.process_btn.setEnabled(False)
            self.status_label.setText(f"Đã đọc file: {self.original_name}")
            self.set_stage(1)
        except Exception as exc:
            self.show_error(str(exc))

    def refresh_column_combos(self):
        current = self.current_column_settings()
        for combo in [self.company_col, self.mst_col, self.address_col, self.product_col, self.qty_col, self.price_col, self.output_col, self.invoice_status_col]:
            allow_empty = bool(combo.property("allow_empty"))
            combo.blockSignals(True)
            combo.clear()
            if allow_empty:
                combo.addItem("Không dùng", "")
            for column in self.columns:
                combo.addItem(column["label"], column["letter"])
            combo.blockSignals(False)
        self._set_mapping_values(current)
        self.refresh_inventory_rule_column_combos()

    def refresh_preview(self):
        if not self.preview_rows:
            self.preview_table.setRowCount(0)
            self.preview_table.setColumnCount(0)
            return
        keys = list(self.preview_rows[0].keys())
        self.preview_table.setColumnCount(len(keys))
        self.preview_table.setHorizontalHeaderLabels(keys)
        self.preview_table.setRowCount(len(self.preview_rows))
        for row_index, row in enumerate(self.preview_rows):
            for col_index, key in enumerate(keys):
                self.preview_table.setItem(row_index, col_index, QTableWidgetItem(str(row.get(key, ""))))
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

    def refresh_invoice_statuses(self):
        if not self.source_path:
            return
        try:
            _, df = read_workbook(self.source_path)
            statuses = invoice_status_options(df, self._combo_data(self.invoice_status_col), DEFAULT_INVOICE_STATUS_SKIP_VALUES)
            self.invoice_statuses = statuses
            self.invoice_table.setRowCount(len(statuses))
            for row, status in enumerate(statuses):
                check = QTableWidgetItem()
                check.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                check.setCheckState(Qt.Checked if status.get("skip") else Qt.Unchecked)
                self.invoice_table.setItem(row, 0, check)
                self.invoice_table.setItem(row, 1, QTableWidgetItem(str(status.get("value", ""))))
                self.invoice_table.setItem(row, 2, QTableWidgetItem(str(status.get("count", ""))))
        except Exception:
            pass

    def choose_output_path(self):
        initial_path = self.current_output_path() or self.default_output_path_text()
        path, _ = QFileDialog.getSaveFileName(self, "Chọn file kết quả", initial_path, "Excel files (*.xlsx)")
        if path:
            self.set_output_path(path)

    def default_output_path_text(self):
        return str(resolve_output_path(self.original_name or "output.xlsx", ""))

    def update_output_path_hint(self):
        hint = self.default_output_path_text()
        for edit in self.output_path_fields():
            edit.setPlaceholderText(hint)

    def output_path_fields(self):
        fields = []
        if hasattr(self, "output_path_edit"):
            fields.append(self.output_path_edit)
        if hasattr(self, "company_output_path_edit"):
            fields.append(self.company_output_path_edit)
        return fields

    def set_output_path(self, path):
        text = str(path or "")
        for edit in self.output_path_fields():
            if edit.text() != text:
                edit.setText(text)

    def sync_output_path_fields(self, source, text):
        if self.output_path_syncing:
            return
        self.output_path_syncing = True
        try:
            for edit in self.output_path_fields():
                if edit is not source and edit.text() != text:
                    edit.setText(text)
        finally:
            self.output_path_syncing = False

    def current_output_path(self):
        preferred_fields = self.output_path_fields()
        if self.current_stage >= 2 and hasattr(self, "company_output_path_edit"):
            preferred_fields = [self.company_output_path_edit, self.output_path_edit]
        for edit in preferred_fields:
            text = edit.text().strip()
            if text:
                return text
        return ""

    def check_companies(self):
        if not self.source_path:
            self.show_error("Hãy chọn file Excel trước.")
            return
        try:
            profile_cfg = self.current_profile_config()
            data = analyze(
                self.source_path,
                self._combo_data(self.company_col) or "F",
                self._combo_data(self.mst_col) or "G",
                self._combo_data(self.address_col) or "H",
                self._combo_data(self.product_col) or "M",
                self._combo_data(self.qty_col) or "O",
                self._combo_data(self.price_col) or "",
                profile_cfg,
                self._combo_data(self.invoice_status_col) or DEFAULT_INVOICE_STATUS_COL,
                self.selected_invoice_statuses(),
            )
            self.price_conflict_rows = []
            self.price_rules_initialized = False
            self.companies = []
            removed = profile_cfg.get("removed_companies") or {}
            prefixes = profile_cfg.get("prefixes") or {}
            skipped = profile_cfg.get("selected_products") or {}
            for company in data.get("companies", []):
                selected = {item["name"] for item in company.get("all_products", [])}
                if isinstance(skipped.get(company["mst"]), list):
                    selected -= set(skipped[company["mst"]])
                company["selected_products"] = selected
                company["process"] = not removed.get(company["mst"], False)
                company["pending_process"] = company["process"]
                if prefixes.get(company["mst"]):
                    company["value"] = prefixes[company["mst"]]
                self.companies.append(company)
            self.render_companies()
            self.verify_prefixes()
            self.process_btn.setEnabled(bool(self.companies))
            self.status_label.setText(f"Tìm thấy {data.get('company_count', 0)} công ty, {data.get('rows_to_process', 0)} dòng xử lý.")
            self.set_stage(2)
        except Exception as exc:
            self.show_error(str(exc))

    def render_companies(self):
        self.company_table.blockSignals(True)
        self.company_table.clearSpans()
        self.company_table.setRowCount(0)
        duplicate_indexes, active_indexes, skipped_indexes = self.company_groups_by_prefix()
        display_groups = []
        if duplicate_indexes:
            display_groups.append(("Công ty trùng prefix cần kiểm tra", duplicate_indexes, QColor("#fff4ce")))
        if active_indexes:
            display_groups.append(("Các công ty đang xử lý", active_indexes, QColor("#ffffff")))
        if skipped_indexes:
            display_groups.append(("Các công ty đã bỏ qua", skipped_indexes, QColor("#f3f4f6")))

        for title, indexes, color in display_groups:
            self.add_company_group_header(title, len(indexes))
            for company_index in indexes:
                self.add_company_row(company_index, color)

        self.company_table.blockSignals(False)
        self.company_filter_dirty = False
        if hasattr(self, "apply_company_filter_btn"):
            self.apply_company_filter_btn.setEnabled(False)
        selected_row = self.company_table_row_for_company_index(self.current_company_index)
        if selected_row is None:
            selected_row = self.first_company_table_row()
        if selected_row is not None:
            self.company_table.selectRow(selected_row)
            self.current_company_index = self.company_index_for_table_row(selected_row)
            self.render_products(self.current_company_index)
        else:
            self.current_company_index = -1
            self.render_products(-1)
        self.update_stage_navigation()

    def company_groups_by_prefix(self):
        counts = {}
        if self.include_prefix.isChecked():
            for company in self.companies:
                if not company.get("process"):
                    continue
                prefix = str(company.get("value", "")).strip().upper()
                if prefix:
                    counts[prefix] = counts.get(prefix, 0) + 1
        duplicate_indexes = []
        active_indexes = []
        skipped_indexes = []
        for index, company in enumerate(self.companies):
            if not company.get("process"):
                skipped_indexes.append(index)
                continue
            prefix = str(company.get("value", "")).strip().upper()
            if prefix and counts.get(prefix, 0) > 1:
                duplicate_indexes.append(index)
            else:
                active_indexes.append(index)
        return duplicate_indexes, active_indexes, skipped_indexes

    def add_company_group_header(self, title, count):
        row = self.company_table.rowCount()
        self.company_table.insertRow(row)
        item = QTableWidgetItem(f"{title} ({count})")
        item.setFlags(Qt.ItemIsEnabled)
        item.setBackground(QColor("#dbeafe"))
        item.setData(Qt.UserRole, -1)
        self.company_table.setItem(row, 0, item)
        self.company_table.setSpan(row, 0, 1, self.company_table.columnCount())

    def add_company_row(self, company_index, color):
        company = self.companies[company_index]
        row = self.company_table.rowCount()
        self.company_table.insertRow(row)
        process = QTableWidgetItem()
        process.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        process_checked = company.get("pending_process", company.get("process"))
        process.setCheckState(Qt.Checked if process_checked else Qt.Unchecked)
        has_long_code = self.company_has_long_preview(company)
        row_color = QColor("#fee2e2") if has_long_code else color
        company_name = company.get("company", "")
        if has_long_code:
            company_name = f"{company_name}  (Mã VT > 50 ký tự)"
        values = [
            process,
            QTableWidgetItem(company.get("mst", "")),
            QTableWidgetItem(company_name),
            QTableWidgetItem(str(company.get("count", ""))),
            QTableWidgetItem(str(company.get("value", ""))),
        ]
        for column, item in enumerate(values):
            item.setData(Qt.UserRole, company_index)
            item.setBackground(row_color)
            if has_long_code:
                item.setToolTip("Có mã VT xem trước vượt 50 ký tự. Chọn dòng công ty để kiểm tra hàng hóa bên dưới.")
            self.company_table.setItem(row, column, item)

    def company_has_long_preview(self, company):
        for product in company.get("all_products", []) or []:
            product_name = product.get("name", "")
            if product_name in company.get("selected_products", set()) and self.product_code_too_long(company, product_name):
                return True
        return False

    def product_code_too_long(self, company, product_name):
        return len(self.preview_code(company, product_name) or "") > 50

    def first_company_table_row(self):
        for row in range(self.company_table.rowCount()):
            if self.company_index_for_table_row(row) >= 0:
                return row
        return None

    def company_table_row_for_company_index(self, company_index):
        if company_index < 0:
            return None
        for row in range(self.company_table.rowCount()):
            if self.company_index_for_table_row(row) == company_index:
                return row
        return None

    def company_index_for_table_row(self, row):
        item = self.company_table.item(row, 0)
        if not item:
            return -1
        value = item.data(Qt.UserRole)
        return int(value) if isinstance(value, int) else -1

    def render_products(self, company_index):
        self.product_table.blockSignals(True)
        self.product_table.setRowCount(0)
        if company_index < 0 or company_index >= len(self.companies):
            self.product_company_label.setText("Chọn một dòng công ty để xem danh sách hàng hóa bên dưới.")
            self.product_warning_label.setVisible(False)
            self.product_table.blockSignals(False)
            return
        company = self.companies[company_index]
        self.product_company_label.setText(f"Hàng hóa của {company.get('company', '')} - MST {company.get('mst', '')}")
        products = company.get("all_products", [])
        self.product_table.setRowCount(len(products))
        long_count = 0
        for row, product in enumerate(products):
            product_name = product.get("name", "")
            selected = product_name in company.get("selected_products", set())
            code = self.preview_code(company, product_name)
            has_long_code = selected and len(code or "") > 50
            if has_long_code:
                long_count += 1
            check = QTableWidgetItem()
            check.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            check.setCheckState(Qt.Checked if selected else Qt.Unchecked)
            name_item = QTableWidgetItem(product_name)
            count_item = QTableWidgetItem(str(product.get("count", "")))
            code_item = QTableWidgetItem(code)
            if has_long_code:
                for item in [check, name_item, count_item, code_item]:
                    item.setBackground(QColor("#fee2e2"))
                    item.setToolTip("Mã VT xem trước vượt 50 ký tự. Hãy chỉnh tên hàng hóa hoặc quy tắc trước khi xử lý.")
                code_item.setText(f"{code}  (vượt 50 ký tự)")
            self.product_table.setItem(row, 0, check)
            self.product_table.setItem(row, 1, name_item)
            self.product_table.setItem(row, 2, count_item)
            self.product_table.setItem(row, 3, code_item)
        if long_count:
            self.product_warning_label.setText(f"Cảnh báo: {long_count} mã VT xem trước vượt 50 ký tự.")
            self.product_warning_label.setVisible(True)
        else:
            self.product_warning_label.setVisible(False)
        self.product_table.blockSignals(False)

    def on_company_item_changed(self, item):
        company_index = self.company_index_for_table_row(item.row())
        if company_index < 0 or company_index >= len(self.companies):
            return
        company = self.companies[company_index]
        if item.column() == 0:
            company["pending_process"] = item.checkState() == Qt.Checked
            self.company_filter_dirty = True
            self.apply_company_filter_btn.setEnabled(True)
            self.status_label.setText("Đã chọn/bỏ công ty. Bấm Áp dụng lọc công ty để xác nhận thay đổi.")
        elif item.column() == 4:
            company["value"] = item.text().strip().upper()
            self.verify_prefixes()
        self.current_company_index = company_index
        self.render_products(self.current_company_index)

    def on_company_selection_changed(self):
        indexes = self.company_table.selectionModel().selectedRows()
        if not indexes:
            return
        company_index = self.company_index_for_table_row(indexes[0].row())
        if company_index < 0:
            return
        self.current_company_index = company_index
        self.render_products(self.current_company_index)

    def on_product_item_changed(self, item):
        if self.current_company_index < 0 or self.current_company_index >= len(self.companies) or item.column() != 0:
            return
        company = self.companies[self.current_company_index]
        products = company.get("all_products", [])
        if item.row() < 0 or item.row() >= len(products):
            return
        product = products[item.row()]
        selected = company.setdefault("selected_products", set())
        if item.checkState() == Qt.Checked:
            selected.add(product.get("name"))
        else:
            selected.discard(product.get("name"))

    def update_company_statuses(self):
        counts = {}
        if self.include_prefix.isChecked():
            for company in self.companies:
                if not company.get("process"):
                    continue
                prefix = str(company.get("value", "")).strip().upper()
                company["value"] = prefix
                if prefix:
                    counts[prefix] = counts.get(prefix, 0) + 1
        for company in self.companies:
            if not company.get("process"):
                status = "Bỏ qua"
            elif not self.include_prefix.isChecked():
                status = "Không dùng tiền tố"
            else:
                prefix = str(company.get("value", "")).strip().upper()
                company["value"] = prefix
                if not prefix:
                    status = "Thiếu tiền tố"
                elif counts.get(prefix, 0) > 1:
                    status = f"Trùng {prefix}"
                else:
                    status = "Hợp lệ"
            company["status"] = status

    def verify_prefixes(self):
        self.update_company_statuses()
        self.render_companies()

    def apply_company_filter(self):
        for company in self.companies:
            if "pending_process" in company:
                company["process"] = bool(company["pending_process"])
            else:
                company["pending_process"] = bool(company.get("process"))
        self.verify_prefixes()
        self.status_label.setText("Đã áp dụng lọc công ty.")

    def apply_prefix_mode(self, mode):
        for company in self.companies:
            if not company.get("process"):
                continue
            if mode == "mst":
                value = self.company_mst_suffix(company)
            elif mode == "initials":
                value = suggest_prefix(company.get("company", ""))
            else:
                initials = suggest_prefix(company.get("company", ""))
                suffix = self.company_mst_suffix(company)
                value = f"{initials}{suffix}" if initials and suffix else ""
            if value:
                company["value"] = value.upper()
        self.verify_prefixes()

    def load_word_rule_state(self):
        profile_cfg = self.current_profile_config()
        self.word_rules = dict(profile_cfg.get("word_rules") or {})
        self.first_word_rules = dict(profile_cfg.get("first_word_rules") or {})
        self.repeated_phrase_removals = list(profile_cfg.get("repeated_phrase_removals") or [])
        self.render_word_rule_tables()

    def render_word_rule_tables(self):
        if not hasattr(self, "word_rule_table"):
            return
        self.word_rule_table_rendering = True
        for table, rules in [(self.word_rule_table, self.word_rules), (self.first_word_rule_table, self.first_word_rules)]:
            table.blockSignals(True)
            table.setRowCount(0)
            for word, output in sorted((rules or {}).items(), key=lambda item: item[0].casefold()):
                row = table.rowCount()
                table.insertRow(row)
                table.setItem(row, 0, QTableWidgetItem(str(word)))
                table.setItem(row, 1, QTableWidgetItem(str(output)))
            table.blockSignals(False)
        self.repeated_phrase_table.blockSignals(True)
        self.repeated_phrase_table.setRowCount(0)
        for phrase in self.repeated_phrase_removals:
            row = self.repeated_phrase_table.rowCount()
            self.repeated_phrase_table.insertRow(row)
            self.repeated_phrase_table.setItem(row, 0, QTableWidgetItem(str(phrase)))
        self.repeated_phrase_table.blockSignals(False)
        if self.current_profile() == "cao_thanh" and self.first_word_rule_table.rowCount() == 0:
            self.add_word_rule_row("first")
        if self.word_rule_table.rowCount() == 0:
            self.add_word_rule_row("rest")
        show_first_word_rules = self.current_profile() == "cao_thanh"
        self.first_word_rule_box.setVisible(show_first_word_rules)
        self.add_first_word_rule_btn.setVisible(show_first_word_rules)
        self.delete_first_word_rule_btn.setVisible(show_first_word_rules)
        self.word_rule_table_rendering = False

    def add_word_rule_row(self, target):
        table = self.word_rule_table
        if target == "first":
            table = self.first_word_rule_table
        elif target == "repeat":
            table = self.repeated_phrase_table
        row = table.rowCount()
        table.insertRow(row)
        table.setItem(row, 0, QTableWidgetItem(""))
        if target != "repeat":
            table.setItem(row, 1, QTableWidgetItem(""))

    def delete_word_rule_rows(self, target):
        table = self.word_rule_table
        if target == "first":
            table = self.first_word_rule_table
        elif target == "repeat":
            table = self.repeated_phrase_table
        selected_rows = sorted({index.row() for index in table.selectedIndexes()}, reverse=True)
        for row in selected_rows:
            table.removeRow(row)

    def rules_from_table(self, table, label):
        rules = {}
        seen = {}
        for row in range(table.rowCount()):
            word_item = table.item(row, 0)
            output_item = table.item(row, 1)
            word = (word_item.text() if word_item else "").strip()
            output = (output_item.text() if output_item else "").strip() or word
            if not word or not output:
                continue
            key = normalize_rule_key(word)
            if key in seen and seen[key] != word:
                raise ValueError(f"{label}: '{seen[key]}' và '{word}' bị trùng sau chuẩn hóa.")
            seen[key] = word
            rules[word] = output
        return dict(sorted(rules.items(), key=lambda item: item[0].casefold()))

    def phrases_from_table(self, table):
        phrases = []
        seen = set()
        for row in range(table.rowCount()):
            item = table.item(row, 0)
            phrase = (item.text() if item else "").strip()
            if not phrase:
                continue
            key = phrase.casefold()
            if key in seen:
                continue
            seen.add(key)
            phrases.append(phrase)
        return sorted(phrases, key=str.casefold)

    def sync_word_rules_from_tables(self):
        if not hasattr(self, "word_rule_table"):
            return
        self.word_rules = self.rules_from_table(self.word_rule_table, "Từ thay riêng")
        self.first_word_rules = self.rules_from_table(self.first_word_rule_table, "Hai từ đầu tiên") if self.current_profile() == "cao_thanh" else {}
        self.repeated_phrase_removals = self.phrases_from_table(self.repeated_phrase_table)

    def apply_word_rules_from_tables(self):
        try:
            self.sync_word_rules_from_tables()
            self.render_word_rule_tables()
            self.render_companies()
            if self.current_company_index >= 0:
                self.render_products(self.current_company_index)
            self.price_conflict_rows = []
            self.price_rules_initialized = False
            self.status_label.setText("Đã áp dụng từ thay riêng cho preview hiện tại.")
        except Exception as exc:
            self.show_error(exc)

    def load_inventory_config_state(self):
        profile_cfg = self.current_profile_config()
        self.inventory_pairs = [dict(pair) for pair in profile_cfg.get("inventory_pairs") or []]
        self.use_default_inventory_pair = bool(profile_cfg.get("use_default_inventory_pair"))
        self.default_inventory_pair_id = str(profile_cfg.get("default_inventory_pair_id") or "").strip()
        self.inventory_pair_rules = [dict(rule) for rule in profile_cfg.get("inventory_pair_rules") or []]
        self.render_inventory_config_tables()

    def render_inventory_config_tables(self):
        if not hasattr(self, "inventory_pair_table"):
            return
        self.inventory_table_rendering = True
        self.inventory_pair_table.blockSignals(True)
        self.inventory_pair_table.setRowCount(0)
        for pair in self.inventory_pairs:
            self.add_inventory_pair_row(pair)
        self.inventory_pair_table.blockSignals(False)

        self.inventory_rule_table.blockSignals(True)
        self.inventory_rule_table.setRowCount(0)
        for rule in self.inventory_pair_rules:
            self.add_inventory_rule_row(rule)
        self.inventory_rule_table.blockSignals(False)
        self.inventory_table_rendering = False
        self.refresh_inventory_pair_options()
        self.update_inventory_default_note()

    def add_inventory_pair_row(self, pair=None):
        pair = pair or {}
        row = self.inventory_pair_table.rowCount()
        self.inventory_pair_table.insertRow(row)
        ma_kho_item = QTableWidgetItem(str(pair.get("ma_kho") or ""))
        pair_id = str(pair.get("id") or "").strip()
        if pair_id:
            ma_kho_item.setData(Qt.UserRole, pair_id)
        self.inventory_pair_table.setItem(row, 0, ma_kho_item)
        self.inventory_pair_table.setItem(row, 1, QTableWidgetItem(str(pair.get("tk_vat_tu") or "")))
        if not self.inventory_table_rendering:
            self.refresh_inventory_pair_options()

    def delete_inventory_pair_rows(self):
        selected_rows = sorted({index.row() for index in self.inventory_pair_table.selectedIndexes()}, reverse=True)
        for row in selected_rows:
            self.inventory_pair_table.removeRow(row)
        if selected_rows:
            self.sync_inventory_config_from_tables()
            self.refresh_inventory_pair_options()

    def add_inventory_rule_row(self, rule=None):
        rule = rule or {}
        row = self.inventory_rule_table.rowCount()
        self.inventory_rule_table.insertRow(row)

        enabled = QTableWidgetItem()
        enabled.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        enabled.setCheckState(Qt.Checked if rule.get("enabled") is not False else Qt.Unchecked)
        self.inventory_rule_table.setItem(row, 0, enabled)

        source_combo = self.inventory_source_column_combo(str(rule.get("source_col") or ""))
        self.inventory_rule_table.setCellWidget(row, 1, source_combo)

        operator_combo = QComboBox()
        operator_combo.addItem("Chứa", "contains")
        operator_combo.addItem("Bằng", "equals")
        self._set_combo_data(operator_combo, str(rule.get("operator") or "contains"))
        self.inventory_rule_table.setCellWidget(row, 2, operator_combo)

        self.inventory_rule_table.setItem(row, 3, QTableWidgetItem(str(rule.get("value") or "")))

        pair_combo = self.inventory_pair_combo(str(rule.get("pair_id") or ""))
        self.inventory_rule_table.setCellWidget(row, 4, pair_combo)

    def delete_inventory_rule_rows(self):
        selected_rows = sorted({index.row() for index in self.inventory_rule_table.selectedIndexes()}, reverse=True)
        for row in selected_rows:
            self.inventory_rule_table.removeRow(row)
        if selected_rows:
            self.sync_inventory_config_from_tables()

    def inventory_source_column_combo(self, selected_value=""):
        combo = QComboBox()
        combo.addItem("Chọn cột", "")
        for column in self.columns:
            combo.addItem(column["label"], column["letter"])
        selected_value = str(selected_value or "").strip().upper()
        if selected_value and combo.findData(selected_value) < 0:
            combo.addItem(f"{selected_value} - đã lưu", selected_value)
        self._set_combo_data(combo, selected_value)
        return combo

    def inventory_pair_combo(self, selected_value=""):
        combo = QComboBox()
        combo.addItem("Chọn cặp", "")
        for pair in self.inventory_pairs:
            combo.addItem(self.inventory_pair_label(pair), pair.get("id", ""))
        selected_value = str(selected_value or "").strip()
        if selected_value and combo.findData(selected_value) < 0:
            combo.addItem(f"{selected_value} - đã lưu", selected_value)
        self._set_combo_data(combo, selected_value)
        return combo

    def refresh_inventory_rule_column_combos(self):
        if not hasattr(self, "inventory_rule_table"):
            return
        for row in range(self.inventory_rule_table.rowCount()):
            current = self._combo_data(self.inventory_rule_table.cellWidget(row, 1))
            self.inventory_rule_table.setCellWidget(row, 1, self.inventory_source_column_combo(current))

    def refresh_inventory_pair_options(self):
        if not hasattr(self, "default_inventory_pair_combo"):
            return
        pairs = self.inventory_pairs_from_table()
        self.inventory_pairs = pairs
        previous_default = self._combo_data(self.default_inventory_pair_combo) or self.default_inventory_pair_id
        self.default_inventory_pair_combo.blockSignals(True)
        self.default_inventory_pair_combo.clear()
        self.default_inventory_pair_combo.addItem("Chọn cặp", "")
        for pair in pairs:
            self.default_inventory_pair_combo.addItem(self.inventory_pair_label(pair), pair.get("id", ""))
        if previous_default and self.default_inventory_pair_combo.findData(previous_default) >= 0:
            self._set_combo_data(self.default_inventory_pair_combo, previous_default)
        elif len(pairs) == 1:
            self._set_combo_data(self.default_inventory_pair_combo, pairs[0].get("id", ""))
        self.default_inventory_pair_combo.blockSignals(False)

        for row in range(self.inventory_rule_table.rowCount()):
            current = self._combo_data(self.inventory_rule_table.cellWidget(row, 4))
            self.inventory_rule_table.setCellWidget(row, 4, self.inventory_pair_combo(current))
        self.update_inventory_default_note()

    def on_inventory_default_changed(self):
        if self.inventory_table_rendering:
            return
        self.use_default_inventory_pair = self.use_default_inventory_pair_check.isChecked()
        self.default_inventory_pair_id = self._combo_data(self.default_inventory_pair_combo)
        self.update_inventory_default_note()

    def update_inventory_default_note(self):
        if not hasattr(self, "inventory_default_note"):
            return
        pairs = self.inventory_pairs_from_table()
        self.use_default_inventory_pair_check.blockSignals(True)
        self.use_default_inventory_pair_check.setChecked(bool(self.use_default_inventory_pair))
        self.use_default_inventory_pair_check.blockSignals(False)
        if len(pairs) == 1 and not self.use_default_inventory_pair_check.isChecked():
            self.inventory_default_note.setText("Có 1 cặp: backend sẽ tự dùng cặp này cho mọi dòng xử lý.")
        elif self.use_default_inventory_pair_check.isChecked():
            self.inventory_default_note.setText("Dòng không khớp quy tắc sẽ dùng cặp mặc định đã chọn.")
        else:
            self.inventory_default_note.setText("Dòng không khớp quy tắc sẽ để trống Mã kho/TK vật tư.")

    def inventory_pairs_from_table(self):
        pairs = []
        used = set()
        for row in range(self.inventory_pair_table.rowCount()):
            ma_kho_item = self.inventory_pair_table.item(row, 0)
            tk_vat_tu_item = self.inventory_pair_table.item(row, 1)
            ma_kho = (ma_kho_item.text() if ma_kho_item else "").strip()
            tk_vat_tu = (tk_vat_tu_item.text() if tk_vat_tu_item else "").strip()
            if not ma_kho or not tk_vat_tu:
                continue
            raw_pair_id = ma_kho_item.data(Qt.UserRole) if ma_kho_item else ""
            pair_id = str(raw_pair_id or "").strip()
            if not pair_id or pair_id in used:
                pair_id = self.next_inventory_pair_id(used)
                if ma_kho_item:
                    ma_kho_item.setData(Qt.UserRole, pair_id)
            used.add(pair_id)
            pairs.append({"id": pair_id, "ma_kho": ma_kho, "tk_vat_tu": tk_vat_tu})
        return pairs

    def next_inventory_pair_id(self, used):
        index = 1
        while f"inventory_pair_{index}" in used:
            index += 1
        return f"inventory_pair_{index}"

    def inventory_pair_label(self, pair):
        return f"{pair.get('ma_kho', '')} / {pair.get('tk_vat_tu', '')}"

    def inventory_rules_from_table(self, valid_pair_ids):
        rules = []
        for row in range(self.inventory_rule_table.rowCount()):
            enabled_item = self.inventory_rule_table.item(row, 0)
            enabled = enabled_item.checkState() == Qt.Checked if enabled_item else True
            source_col = str(self._combo_data(self.inventory_rule_table.cellWidget(row, 1)) or "").strip().upper()
            operator = str(self._combo_data(self.inventory_rule_table.cellWidget(row, 2)) or "contains").strip().casefold()
            value_item = self.inventory_rule_table.item(row, 3)
            value = (value_item.text() if value_item else "").strip()
            pair_id = str(self._combo_data(self.inventory_rule_table.cellWidget(row, 4)) or "").strip()
            has_any_value = bool(source_col or value or pair_id)
            has_complete_value = bool(source_col and value and pair_id)
            if not has_any_value:
                continue
            if not enabled and not has_complete_value:
                continue
            if not has_complete_value and not enabled:
                continue
            if not has_complete_value and enabled:
                rules.append({
                    "source_col": source_col,
                    "operator": operator if operator in {"contains", "equals"} else "contains",
                    "value": value,
                    "pair_id": pair_id,
                    "enabled": enabled,
                })
                continue
            if pair_id not in valid_pair_ids:
                if enabled:
                    rules.append({
                        "source_col": source_col,
                        "operator": operator if operator in {"contains", "equals"} else "contains",
                        "value": value,
                        "pair_id": pair_id,
                        "enabled": enabled,
                    })
                continue
            rules.append({
                "source_col": source_col,
                "operator": operator if operator in {"contains", "equals"} else "contains",
                "value": value,
                "pair_id": pair_id,
                "enabled": enabled,
            })
        return rules

    def sync_inventory_config_from_tables(self):
        if not hasattr(self, "inventory_pair_table"):
            return
        pairs = self.inventory_pairs_from_table()
        valid_pair_ids = {pair["id"] for pair in pairs}
        self.inventory_pairs = pairs
        self.use_default_inventory_pair = self.use_default_inventory_pair_check.isChecked()
        selected_default = str(self._combo_data(self.default_inventory_pair_combo) or self.default_inventory_pair_id or "").strip()
        if selected_default not in valid_pair_ids:
            if len(pairs) == 1:
                selected_default = pairs[0]["id"]
            elif not self.use_default_inventory_pair:
                selected_default = ""
        self.default_inventory_pair_id = selected_default
        self.inventory_pair_rules = self.inventory_rules_from_table(valid_pair_ids)

    def apply_inventory_config_from_tables(self):
        try:
            self.sync_inventory_config_from_tables()
            self.refresh_inventory_pair_options()
            self.status_label.setText("Đã áp dụng cấu hình cặp Mã kho / TK vật tư.")
        except Exception as exc:
            self.show_error(exc)

    def load_price_rule_state(self):
        profile_cfg = self.current_profile_config()
        self.price_group_rules = dict(profile_cfg.get("price_group_rules") or {})
        self.price_range_rules = dict(profile_cfg.get("price_range_rules") or {})
        self.price_adjust_all_percent = self.percent_value(profile_cfg.get("price_adjust_all_percent"), 0)
        self.price_conflict_rows = []
        self.price_rules_initialized = False
        if hasattr(self, "price_filter_all_edit"):
            self.price_filter_all_edit.setText("8")
        if hasattr(self, "price_adjust_all_edit"):
            self.price_adjust_all_edit.setText(self.format_percent(self.price_adjust_all_percent))

    def go_to_price_stage(self):
        if self.current_profile() != "cao_thanh":
            return
        self.set_stage(3)

    def prepare_price_stage(self):
        if self.company_filter_dirty:
            self.show_error("Đang có thay đổi lọc công ty chưa áp dụng. Bấm Áp dụng lọc công ty trước khi lọc đơn giá.")
            return False
        self.verify_prefixes()
        if self.include_prefix.isChecked():
            invalid = [company for company in self.companies if company.get("process") and company.get("status") != "Hợp lệ"]
            if invalid:
                self.show_error("Còn tiền tố bị thiếu hoặc bị trùng. Hãy xử lý ở phần tùy chỉnh tiền tố trước khi lọc giá.")
                return False
        self.refresh_price_groups()
        return True

    def product_key(self, mst, product_name):
        return f"{mst}|||{product_name}"

    def product_code_for(self, company, product_name, trim=True):
        if not company.get("process") or product_name not in company.get("selected_products", set()):
            return ""
        profile_cfg = self.current_profile_config()
        key = self.product_key(company.get("mst", ""), product_name)
        manual_code = str((profile_cfg.get("manual_code_overrides") or {}).get(key) or "").strip()
        if manual_code:
            return manual_code[:50] if trim else manual_code
        prefix_map = {company.get("mst", ""): company.get("value", "")}
        if not trim:
            body = make_product_part(
                self.current_profile(),
                product_name,
                self.word_rules,
                self.first_word_rules,
                self.repeated_phrase_removals,
            )
            if not self.include_prefix.isChecked():
                return body
            mst = company.get("mst", "")
            if not mst or mst not in prefix_map:
                return ""
            return f"{normalize_token(prefix_map[mst])}.{body}"
        return make_code(
            company.get("mst", ""),
            product_name,
            1,
            prefix_map,
            self.current_profile(),
            self.word_rules,
            self.first_word_rules,
            require_qty=False,
            include_company_prefix=self.include_prefix.isChecked(),
            repeated_phrase_removals=self.repeated_phrase_removals,
        )

    def refresh_price_groups(self):
        self.price_conflict_rows = self.build_price_conflict_rows()
        self.price_rules_initialized = True
        self.store_price_rules_from_rows()
        self.render_price_groups()

    def build_price_conflict_rows(self):
        grouped = {}
        profile_cfg = self.current_profile_config()
        manual_code_overrides = profile_cfg.get("manual_code_overrides") or {}
        for company in self.companies:
            if not company.get("process"):
                continue
            for product in company.get("all_products", []) or []:
                product_name = product.get("name", "")
                if product_name not in company.get("selected_products", set()):
                    continue
                product_key = self.product_key(company.get("mst", ""), product_name)
                code = str(manual_code_overrides.get(product_key) or "").strip() or self.product_code_for(company, product_name)
                if not code:
                    continue
                price_items = []
                for index, item in enumerate(product.get("priceRows", []) or []):
                    price = self.number_value(item.get("price"))
                    if price is None:
                        continue
                    quantity = self.valid_quantity(item.get("quantity"))
                    price_items.append({"item": item, "index": index, "price": price, "quantity": quantity})
                if not price_items:
                    continue
                saved_by_product = self.price_group_rules.get(product_key)
                if code not in grouped:
                    saved_by_code = self.price_range_rules.get(code)
                    saved_rule = saved_by_product or saved_by_code or None
                    grouped[code] = {
                        "key": f"price-code|||{code}",
                        "code": code,
                        "companies": [],
                        "companyKeys": set(),
                        "products": [],
                        "productKeySet": set(),
                        "sourceRows": [],
                        "filterPercent": self.percent_value((saved_rule or {}).get("percent") if saved_rule else None, 8),
                        "draftFilterPercent": self.percent_value((saved_rule or {}).get("percent") if saved_rule else None, 8),
                        "savedRule": saved_rule,
                        "buckets": [],
                    }
                row = grouped[code]
                if company.get("mst") not in row["companyKeys"]:
                    row["companyKeys"].add(company.get("mst"))
                    row["companies"].append(company)
                if product_key not in row["productKeySet"]:
                    row["productKeySet"].add(product_key)
                    row["products"].append({"key": product_key, "name": product_name, "company": company})
                for price_item in price_items:
                    item = price_item["item"]
                    row["sourceRows"].append({
                        "key": f"{product_key}|||{item.get('excelRow') or price_item['index']}",
                        "company": company,
                        "product": product,
                        "productKey": product_key,
                        "excelRow": item.get("excelRow") or "",
                        "stt": item.get("stt") or "",
                        "invoiceNo": item.get("invoiceNo") or "",
                        "invoiceDate": item.get("invoiceDate") or "",
                        "name": item.get("name") or product_name,
                        "unit": item.get("unit") or "",
                        "price": price_item["price"],
                        "quantity": price_item["quantity"],
                        "totalAmount": self.line_total_amount(price_item["price"], price_item["quantity"], item.get("amount")),
                    })
        rows = []
        for row in grouped.values():
            prices = [item["price"] for item in row["sourceRows"]]
            unique_prices = set(prices)
            row["count"] = len(unique_prices)
            row["hasMultiplePrices"] = len(unique_prices) > 1
            row["priceRowCount"] = len(prices)
            row["min"] = min(prices)
            row["max"] = max(prices)
            row["quantity"] = self.total_quantity(row["sourceRows"])
            row["totalAmount"] = self.total_amount(row["sourceRows"])
            row["average"] = self.weighted_average_price(row["sourceRows"])
            row["companyDisplay"] = " | ".join(f"{company.get('mst', '')} - {company.get('company', '')}" for company in row["companies"])
            row["productDisplay"] = " | ".join(product["name"] for product in row["products"])
            row["productCount"] = len(row["products"])
            row["buckets"] = self.build_price_buckets(row)
            rows.append(row)
        return sorted(rows, key=lambda item: item.get("code", ""))

    def build_price_buckets(self, row, previous_buckets=None):
        filter_percent = self.percent_value(row.get("filterPercent"), 8)
        sorted_rows = sorted(row.get("sourceRows", []) or [], key=lambda item: item.get("price", 0))
        grouped_rows = []
        for item in sorted_rows:
            current = grouped_rows[-1] if grouped_rows else []
            if not current:
                grouped_rows.append([item])
                continue
            average = self.weighted_average_price(current)
            deviation = abs((item.get("price", 0) - average) / average) * 100 if average > 0 else 0
            if deviation <= filter_percent:
                current.append(item)
            else:
                grouped_rows.append([item])
        buckets = []
        for index, items in enumerate(grouped_rows):
            saved_group = self.saved_price_bucket_rule(row, index, items)
            margin_percent = self.percent_value(saved_group.get("adjust_percent") if saved_group else None, self.price_adjust_all_percent)
            bucket = {
                "key": f"{row.get('key')}|||bucket|||{index + 1}",
                "label": f"Nhóm {index + 1}",
                "finalCode": self.final_price_bucket_code(row.get("code", ""), index, len(grouped_rows)),
                "count": len(items),
                "min": min(item["price"] for item in items),
                "max": max(item["price"] for item in items),
                "averagePrice": self.weighted_average_price(items),
                "marginPercent": margin_percent,
                "draftMarginPercent": margin_percent,
                "quantity": self.total_quantity(items),
                "totalAmount": self.total_amount(items),
                "rows": items,
            }
            self.refresh_price_bucket_summary(bucket)
            if previous_buckets:
                bucket = self.preserve_bucket_draft(bucket, previous_buckets, index)
            buckets.append(bucket)
        return buckets

    def saved_price_bucket_rule(self, row, index, items):
        groups = row.get("savedRule", {}).get("groups", []) if isinstance(row.get("savedRule"), dict) else []
        if not groups:
            return None
        min_price = min(item["price"] for item in items)
        max_price = max(item["price"] for item in items)
        for group in groups:
            if self.number_value(group.get("min_price")) == min_price and self.number_value(group.get("max_price")) == max_price:
                return group
        for group in groups:
            if int(group.get("index") or 0) == index + 1:
                return group
        return None

    def preserve_bucket_draft(self, bucket, previous_buckets, index):
        previous = None
        for candidate in previous_buckets:
            if candidate.get("min") == bucket.get("min") and candidate.get("max") == bucket.get("max"):
                previous = candidate
                break
        if previous is None and index < len(previous_buckets):
            previous = previous_buckets[index]
        if previous is not None:
            margin_percent = self.percent_value(previous.get("draftMarginPercent", previous.get("marginPercent")), bucket.get("marginPercent", 0))
            bucket["marginPercent"] = margin_percent
            bucket["draftMarginPercent"] = margin_percent
            self.refresh_price_bucket_summary(bucket)
        return bucket

    def final_price_bucket_code(self, code, index, bucket_count):
        if not code or bucket_count <= 1:
            return code or ""
        return f"{code}.{index + 1:03d}"

    def refresh_price_bucket_summary(self, bucket):
        margin_percent = self.percent_value(bucket.get("marginPercent"), 0)
        bucket["adjustedAverage"] = self.price_baseline(bucket.get("averagePrice", 0), margin_percent)
        loss_count = 0
        loss_revenue = 0
        loss_amount = 0
        for item in bucket.get("rows", []) or []:
            price = self.percent_value(item.get("price"), 0)
            if price >= bucket["adjustedAverage"]:
                continue
            quantity = self.valid_quantity(item.get("quantity"))
            loss_count += 1
            loss_revenue += self.line_total_amount(item.get("price"), quantity, item.get("totalAmount"))
            loss_amount += (bucket["adjustedAverage"] - price) * quantity
        bucket["lossCount"] = loss_count
        bucket["lossRevenue"] = loss_revenue
        bucket["lossAmount"] = loss_amount
        bucket["hasLoss"] = loss_count > 0

    def price_baseline(self, average, percent):
        average_value = self.percent_value(average, 0)
        percent_value = self.percent_value(percent, 0)
        return average_value * (1 - percent_value / 100)

    def number_value(self, value):
        if value is None:
            return None
        if isinstance(value, (int, float)):
            parsed = float(value)
        else:
            text = str(value).strip().replace(" ", "")
            if not text:
                return None
            if "," in text and "." in text:
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
            try:
                parsed = float(text)
            except ValueError:
                return None
        return parsed if math.isfinite(parsed) else None

    def percent_value(self, value, fallback):
        parsed = self.number_value(value)
        return parsed if parsed is not None else fallback

    def valid_quantity(self, value):
        quantity = self.number_value(value)
        return quantity if quantity is not None and quantity > 0 else 1

    def line_total_amount(self, price, quantity, total_amount):
        parsed_amount = self.number_value(total_amount)
        if parsed_amount is not None and parsed_amount > 0:
            return parsed_amount
        parsed_price = self.percent_value(price, 0)
        return parsed_price * self.valid_quantity(quantity)

    def total_quantity(self, rows):
        return sum(self.valid_quantity(item.get("quantity")) for item in rows or [])

    def total_amount(self, rows):
        return sum(self.line_total_amount(item.get("price"), item.get("quantity"), item.get("totalAmount")) for item in rows or [])

    def weighted_average_price(self, rows):
        quantity = self.total_quantity(rows)
        return self.total_amount(rows) / quantity if quantity > 0 else 0

    def apply_price_filter_percent_to_all(self):
        percent = self.percent_value(self.price_filter_all_edit.text(), 8)
        self.price_filter_all_edit.setText(self.format_percent(percent))
        for row in self.price_conflict_rows:
            previous_buckets = row.get("buckets", [])
            row["filterPercent"] = percent
            row["draftFilterPercent"] = percent
            row["buckets"] = self.build_price_buckets(row, previous_buckets)
        self.store_price_rules_from_rows()
        self.render_price_groups()
        self.status_label.setText("Đã áp dụng % lọc cho tất cả mã VT.")

    def apply_price_adjust_percent_to_all(self):
        percent = self.percent_value(self.price_adjust_all_edit.text(), 0)
        self.price_adjust_all_percent = percent
        self.price_adjust_all_edit.setText(self.format_percent(percent))
        for row in self.price_conflict_rows:
            for bucket in row.get("buckets", []) or []:
                bucket["marginPercent"] = percent
                bucket["draftMarginPercent"] = percent
                self.refresh_price_bucket_summary(bucket)
        self.store_price_rules_from_rows()
        self.render_price_groups()
        self.status_label.setText("Đã áp dụng % lãi cho tất cả nhóm giá.")

    def apply_price_table_edits(self):
        self.store_price_rules_from_rows()
        self.render_price_groups()
        self.status_label.setText("Đã áp dụng chỉnh sửa lọc đơn giá.")

    def on_price_item_changed(self, item):
        if self.price_table_rendering:
            return
        payload = item.data(Qt.UserRole)
        if not isinstance(payload, dict):
            return
        row_index = payload.get("row_index")
        if row_index is None or row_index < 0 or row_index >= len(self.price_conflict_rows):
            return
        price_row = self.price_conflict_rows[row_index]
        if payload.get("type") == "code" and item.column() == 9:
            percent = self.percent_value(item.text(), price_row.get("filterPercent", 8))
            previous_buckets = price_row.get("buckets", [])
            price_row["filterPercent"] = percent
            price_row["draftFilterPercent"] = percent
            price_row["buckets"] = self.build_price_buckets(price_row, previous_buckets)
            self.store_price_rules_from_rows()
            self.render_price_groups()
        elif payload.get("type") == "bucket" and item.column() == 10:
            bucket_index = payload.get("bucket_index")
            buckets = price_row.get("buckets", []) or []
            if bucket_index is None or bucket_index < 0 or bucket_index >= len(buckets):
                return
            percent = self.percent_value(item.text(), buckets[bucket_index].get("marginPercent", 0))
            buckets[bucket_index]["marginPercent"] = percent
            buckets[bucket_index]["draftMarginPercent"] = percent
            self.refresh_price_bucket_summary(buckets[bucket_index])
            self.store_price_rules_from_rows()
            self.render_price_groups()

    def render_price_groups(self):
        if not hasattr(self, "price_table"):
            return
        rows = self.price_conflict_rows or []
        total_price_rows = sum(int(row.get("priceRowCount", 0)) for row in rows)
        single_count = sum(1 for row in rows if not row.get("hasMultiplePrices"))
        multi_count = sum(1 for row in rows if row.get("hasMultiplePrices"))
        split_count = sum(max(1, len(row.get("buckets", []) or [])) for row in rows)
        self.price_total_label.setText(f"Tổng mã/dòng: {len(rows)} / {total_price_rows}")
        self.price_single_label.setText(f"Mã 1 giá: {single_count}")
        self.price_multi_label.setText(f"Mã nhiều giá: {multi_count}")
        self.price_split_label.setText(f"Nhóm sau lọc: {split_count}")

        self.price_table_rendering = True
        self.price_table.blockSignals(True)
        self.price_table.clearSpans()
        self.price_table.setRowCount(0)
        multi_rows = [(index, row) for index, row in enumerate(rows) if row.get("hasMultiplePrices")]
        single_rows = [(index, row) for index, row in enumerate(rows) if not row.get("hasMultiplePrices")]
        if not rows:
            self.price_table.insertRow(0)
            item = QTableWidgetItem("Chưa có dữ liệu đơn giá để lọc. Kiểm tra cột Đơn giá (mặc định P) và hàng hóa đang xử lý.")
            item.setFlags(Qt.ItemIsEnabled)
            self.price_table.setItem(0, 0, item)
            self.price_table.setSpan(0, 0, 1, self.price_table.columnCount())
        else:
            self.add_price_section("Mã VT nhiều đơn giá", multi_rows)
            self.add_price_section("Mã VT 1 đơn giá", single_rows)
        self.price_table.blockSignals(False)
        self.price_table_rendering = False
        self.update_stage_navigation()

    def add_price_section(self, title, indexed_rows):
        if not indexed_rows:
            return
        row_number = self.price_table.rowCount()
        self.price_table.insertRow(row_number)
        item = QTableWidgetItem(f"{title} ({len(indexed_rows)} mã)")
        item.setFlags(Qt.ItemIsEnabled)
        item.setBackground(QColor("#dbeafe"))
        self.price_table.setItem(row_number, 0, item)
        self.price_table.setSpan(row_number, 0, 1, self.price_table.columnCount())
        for row_index, price_row in indexed_rows:
            self.add_price_code_row(row_index, price_row)
            for bucket_index, bucket in enumerate(price_row.get("buckets", []) or []):
                self.add_price_bucket_row(row_index, bucket_index, price_row, bucket)

    def add_price_code_row(self, row_index, price_row):
        table_row = self.price_table.rowCount()
        self.price_table.insertRow(table_row)
        payload = {"type": "code", "row_index": row_index}
        values = [
            "Nhiều đơn giá" if price_row.get("hasMultiplePrices") else "Một đơn giá",
            price_row.get("code", ""),
            price_row.get("companyDisplay", ""),
            price_row.get("productDisplay", ""),
            self.format_price(price_row.get("min", 0)),
            self.format_price(price_row.get("max", 0)),
            self.format_price(price_row.get("average", 0)),
            str(price_row.get("priceRowCount", 0)),
            str(len(price_row.get("buckets", []) or [])),
            self.format_percent(price_row.get("filterPercent", 8)),
            "",
        ]
        for column, value in enumerate(values):
            item = QTableWidgetItem(value)
            item.setData(Qt.UserRole, payload)
            item.setBackground(QColor("#eff6ff"))
            if column == 9:
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
            else:
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.price_table.setItem(table_row, column, item)

    def add_price_bucket_row(self, row_index, bucket_index, price_row, bucket):
        table_row = self.price_table.rowCount()
        self.price_table.insertRow(table_row)
        payload = {"type": "bucket", "row_index": row_index, "bucket_index": bucket_index}
        values = [
            "  Nhóm giá",
            bucket.get("finalCode") or price_row.get("code", ""),
            "",
            bucket.get("label", ""),
            self.format_price(bucket.get("min", 0)),
            self.format_price(bucket.get("max", 0)),
            self.format_price(bucket.get("averagePrice", 0)),
            str(bucket.get("count", 0)),
            f"SL {self.format_quantity(bucket.get('quantity', 0))}",
            "",
            self.format_percent(bucket.get("marginPercent", 0)),
        ]
        for column, value in enumerate(values):
            item = QTableWidgetItem(value)
            item.setData(Qt.UserRole, payload)
            item.setBackground(QColor("#f8fafc"))
            if column == 10:
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
            else:
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.price_table.setItem(table_row, column, item)

    def store_price_rules_from_rows(self):
        if self.current_profile() != "cao_thanh":
            return
        next_rules = {}
        ranges = dict(self.price_range_rules or {})
        for row in self.price_conflict_rows or []:
            code = row.get("code", "")
            if not code:
                continue
            percent = self.percent_value(row.get("filterPercent"), 8)
            groups = []
            for index, bucket in enumerate(row.get("buckets", []) or []):
                groups.append({
                    "index": index + 1,
                    "label": bucket.get("label") or f"Nhóm {index + 1}",
                    "min_price": bucket.get("min"),
                    "max_price": bucket.get("max"),
                    "average_price": bucket.get("averagePrice"),
                    "adjust_percent": self.percent_value(bucket.get("marginPercent"), 0),
                })
            if row.get("hasMultiplePrices") is not False:
                for product in row.get("products", []) or []:
                    next_rules[product.get("key")] = {
                        "base_code": code,
                        "min_price": row.get("min"),
                        "max_price": row.get("max"),
                        "percent": percent,
                        "groups": groups,
                    }
            ranges[code] = {
                "min_price": row.get("min"),
                "max_price": row.get("max"),
                "percent": percent,
                "groups": groups,
            }
        self.price_group_rules = next_rules
        self.price_range_rules = ranges

    def format_price(self, value):
        parsed = self.number_value(value)
        if parsed is None:
            return ""
        if abs(parsed - round(parsed)) < 0.000001:
            return f"{int(round(parsed)):,}"
        return f"{parsed:,.2f}"

    def format_quantity(self, value):
        parsed = self.number_value(value)
        if parsed is None:
            return "0"
        if abs(parsed - round(parsed)) < 0.000001:
            return f"{int(round(parsed)):,}"
        return f"{parsed:,.2f}"

    def format_percent(self, value):
        parsed = self.percent_value(value, 0)
        if abs(parsed - round(parsed)) < 0.000001:
            return str(int(round(parsed)))
        return f"{parsed:.2f}".rstrip("0").rstrip(".")

    def company_mst_suffix(self, company, length=3):
        digits = "".join(ch for ch in str(company.get("mst", "")) if ch.isdigit())
        return digits[-length:] if digits else ""

    def preview_code(self, company, product_name):
        return self.product_code_for(company, product_name, trim=False)

    def selected_invoice_statuses(self):
        values = []
        for row in range(self.invoice_table.rowCount()):
            check = self.invoice_table.item(row, 0)
            value = self.invoice_table.item(row, 1)
            if check and value and check.checkState() == Qt.Checked:
                values.append(value.text())
        return values or DEFAULT_INVOICE_STATUS_SKIP_VALUES[:]

    def current_column_settings(self):
        profile_cfg = self.current_profile_config()
        configured = {**self.default_columns(), **(profile_cfg.get("columns") or {})}
        configured.setdefault("invoice_status_col", DEFAULT_INVOICE_STATUS_COL)
        configured.setdefault("invoice_status_skip_values", DEFAULT_INVOICE_STATUS_SKIP_VALUES[:])

        def combo_or_saved(combo, key):
            if combo.count():
                return self._combo_data(combo)
            return configured.get(key, "")

        return {
            "company_col": combo_or_saved(self.company_col, "company_col") or "F",
            "mst_col": combo_or_saved(self.mst_col, "mst_col") or "G",
            "address_col": combo_or_saved(self.address_col, "address_col") or "H",
            "product_col": combo_or_saved(self.product_col, "product_col") or "M",
            "qty_col": combo_or_saved(self.qty_col, "qty_col") or "O",
            "price_col": combo_or_saved(self.price_col, "price_col") or "",
            "output_col": combo_or_saved(self.output_col, "output_col") or "L",
            "invoice_status_col": combo_or_saved(self.invoice_status_col, "invoice_status_col") or DEFAULT_INVOICE_STATUS_COL,
            "invoice_status_skip_values": self.selected_invoice_statuses(),
        }

    def config_payload(self):
        self.sync_word_rules_from_tables()
        self.sync_inventory_config_from_tables()
        profile_cfg = self.current_profile_config()
        columns = self.current_column_settings()
        if self.current_profile() == "cao_thanh" and self.price_rules_initialized:
            self.store_price_rules_from_rows()
        use_price_rows = self.current_profile() == "cao_thanh" and self.price_rules_initialized and self.current_stage == 3
        price_group_rules = self.price_group_rules if use_price_rows else profile_cfg.get("price_group_rules") or {}
        price_range_rules = self.price_range_rules if use_price_rows else profile_cfg.get("price_range_rules") or {}
        price_adjust_all_percent = self.price_adjust_all_percent if use_price_rows else profile_cfg.get("price_adjust_all_percent") or 0
        data = {
            "saved_name": str(self.source_path or ""),
            "original_name": self.original_name,
            "profile": self.current_profile(),
            **columns,
            "output_path": self.current_output_path(),
            "word_rules": dict(self.word_rules),
            "first_word_rules": dict(self.first_word_rules),
            "repeated_phrase_removals": list(self.repeated_phrase_removals),
            "inventory_pairs": list(self.inventory_pairs),
            "use_default_inventory_pair": bool(self.use_default_inventory_pair),
            "default_inventory_pair_id": self.default_inventory_pair_id,
            "inventory_pair_rules": list(self.inventory_pair_rules),
            "include_company_prefix": self.include_prefix.isChecked(),
            "price_group_rules": price_group_rules,
            "price_range_rules": price_range_rules,
            "price_adjust_all_percent": price_adjust_all_percent,
            "manual_code_overrides": profile_cfg.get("manual_code_overrides") or {},
            "all_mst": [],
            "process_mst": [],
            "mst_safe_id": [],
            "prefixes": dict(profile_cfg.get("prefixes") or {}) if not self.companies else {},
            "removed_companies": dict(profile_cfg.get("removed_companies") or {}) if not self.companies else {},
            "skipped_products_map": dict(profile_cfg.get("selected_products") or {}) if not self.companies else {},
        }
        for index, company in enumerate(self.companies):
            mst = company.get("mst", "")
            sid = str(index)
            data["all_mst"].append(mst)
            data["mst_safe_id"].append(f"{mst}|||{sid}")
            if company.get("process"):
                data["process_mst"].append(mst)
                data[f"prefix_{sid}"] = str(company.get("value", "")).strip().upper()
                selected = sorted(company.get("selected_products", set()))
                data[f"selected_products_{sid}"] = selected
                default_prefix = str(company.get("default_prefix", "")).strip().upper()
                if data[f"prefix_{sid}"] and data[f"prefix_{sid}"] != default_prefix:
                    data["prefixes"][mst] = data[f"prefix_{sid}"]
                all_products = [item.get("name") for item in company.get("all_products", [])]
                skipped = [name for name in all_products if name not in set(selected)]
                if skipped:
                    data["skipped_products_map"][mst] = skipped
            else:
                data["removed_companies"][mst] = True
        return data

    def process_payload(self):
        data = self.config_payload()
        validate_payload(data)
        return data

    def process_file(self):
        if not self.source_path:
            self.show_error("Hãy chọn file Excel trước.")
            return
        try:
            data = self.process_payload()
            output = resolve_output_path(self.original_name, data.get("output_path", ""))
            try:
                processed = process_workbook(self.source_path, output, data)
            except PermissionError:
                self.show_error(f"Không thể ghi file kết quả. Hãy đóng file đang mở rồi thử lại:\n{output}")
                return
            up_stream = create_up_ban_ra_workbook(processed)
            up_output = up_ban_ra_output_path(output)
            try:
                up_output.write_bytes(up_stream.getvalue())
            except PermissionError:
                self.show_error(f"Không thể ghi file kết quả. Hãy đóng file đang mở rồi thử lại:\n{up_output}")
                return
            self.persist_config(data)
            QMessageBox.information(
                self,
                "Hoàn tất",
                f"Đã tạo file:\n{output}\n\n{up_output}",
            )
            self.status_label.setText(f"Đã xử lý xong: {output}")
        except Exception as exc:
            self.show_error(str(exc))

    def save_current_config(self):
        try:
            if self.companies:
                self.verify_prefixes()
            data = self.config_payload()
            self.persist_config(data)
            QMessageBox.information(self, "Lưu cấu hình", "Đã lưu cấu hình hiện tại.")
            self.status_label.setText("Đã lưu cấu hình.")
        except Exception as exc:
            self.show_error(str(exc))

    def persist_config(self, data):
        cfg = load_config()
        profile = profile_key(data.get("profile"))
        cfg["selected_profile"] = profile
        cfg["columns"].update(self.current_column_settings())
        cfg["profiles"].setdefault(profile, empty_profile_config(profile))
        profile_cfg = cfg["profiles"].get(profile) or empty_profile_config(profile)
        profile_cfg.update({
            "prefixes": data.get("prefixes", {}),
            "selected_products": data.get("skipped_products_map", {}),
            "removed_companies": data.get("removed_companies", {}),
            "include_company_prefix": data.get("include_company_prefix") is not False,
            "output_path": data.get("output_path", ""),
            "columns": self.current_column_settings(),
            "word_rules": data.get("word_rules", {}),
            "first_word_rules": data.get("first_word_rules", {}),
            "repeated_phrase_removals": data.get("repeated_phrase_removals", []),
            "inventory_pairs": data.get("inventory_pairs", []),
            "use_default_inventory_pair": bool(data.get("use_default_inventory_pair")),
            "default_inventory_pair_id": str(data.get("default_inventory_pair_id") or "").strip(),
            "inventory_pair_rules": data.get("inventory_pair_rules", []),
        })
        if profile == "cao_thanh":
            profile_cfg.update({
                "price_group_rules": data.get("price_group_rules", {}),
                "price_range_rules": data.get("price_range_rules", {}),
                "price_adjust_all_percent": self.percent_value(data.get("price_adjust_all_percent"), 0),
                "manual_code_overrides": data.get("manual_code_overrides", {}),
            })
        cfg["profiles"][profile] = profile_cfg
        self.config = save_config(cfg)

    def export_config(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export cấu hình", "product_code_config.json", "JSON files (*.json)")
        if path:
            Path(path).write_text(CONFIG_PATH.read_text(encoding="utf-8"), encoding="utf-8")

    def import_config(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import cấu hình", "", "JSON files (*.json)")
        if not path:
            return
        try:
            import json
            self.config = save_config(json.loads(Path(path).read_text(encoding="utf-8")))
            self._load_config_to_ui()
            QMessageBox.information(self, "Import cấu hình", "Đã import cấu hình.")
        except Exception as exc:
            self.show_error(str(exc))

    def clear_profile_cache(self):
        profile = self.current_profile()
        self.config.setdefault("profiles", {})[profile] = empty_profile_config(profile)
        self.config = save_config(self.config)
        self.apply_profile_columns()
        QMessageBox.information(self, "Xóa cache", "Đã xóa cache profile hiện tại.")

    def _combo_data(self, combo):
        return combo.currentData() if combo.currentIndex() >= 0 else ""

    def _set_combo_data(self, combo, value):
        index = combo.findData(value)
        if index >= 0:
            combo.setCurrentIndex(index)
        elif combo.count() and value == "":
            combo.setCurrentIndex(0)

    def show_error(self, message):
        QMessageBox.critical(self, "Thông báo", str(message))
        self.status_label.setText(str(message))


def main():
    qt_app = QApplication(sys.argv)
    window = ProductCodeFormatterWindow()
    window.show()
    sys.exit(qt_app.exec())


if __name__ == "__main__":
    main()
