import json
import subprocess
from pathlib import Path
from typing import Any

from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


ROOT_DIR = Path(__file__).resolve().parent
LICENSE_SERVER_DIR = ROOT_DIR / "license_server"
ADMIN_OPTIONS_PATH = ROOT_DIR / "license_admin_options.json"

DEFAULT_ADMIN_OPTIONS = {
    "profiles": [
        {"key": "son_phuong", "label": "Sơn Phương"},
        {"key": "cao_thanh", "label": "Cao Thành"},
        {"key": "quang_thinh", "label": "Quang Thịnh"},
        {"key": "vietmax", "label": "Vietmax"},
        {"key": "ho_guom", "label": "Hồ Gươm"},
    ],
    "products": [
        {"code": "product-code-formatter", "name": "Product Code Formatter", "label": "Product Code Formatter", "uses_profiles": True, "application": "ProductCodeFormatter"},
        {"code": "inventory-allocator", "name": "Inventory Allocator", "label": "Inventory Allocator", "uses_profiles": False, "application": "InventoryAllocator"},
    ],
    "company_presets": [],
}

CURRENT_DESKTOP_BUILD = "ProductCodeFormatter_v27"


def load_admin_options():
    options = DEFAULT_ADMIN_OPTIONS
    if ADMIN_OPTIONS_PATH.exists():
        with ADMIN_OPTIONS_PATH.open("r", encoding="utf-8") as handle:
            loaded = json.load(handle)
        if isinstance(loaded, dict):
            options = {**DEFAULT_ADMIN_OPTIONS, **loaded}

    profiles = []
    for item in options.get("profiles") or []:
        if isinstance(item, dict):
            key = str(item.get("key") or "").strip()
            label = str(item.get("label") or key).strip()
        else:
            key = str(item[0] if isinstance(item, (list, tuple)) and item else "").strip()
            label = str(item[1] if isinstance(item, (list, tuple)) and len(item) > 1 else key).strip()
        if key:
            profiles.append((key, label or key))
    if not profiles:
        profiles = [(item["key"], item["label"]) for item in DEFAULT_ADMIN_OPTIONS["profiles"]]

    products = []
    for item in options.get("products") or []:
        if not isinstance(item, dict):
            continue
        code = str(item.get("code") or "").strip()
        if not code:
            continue
        products.append((
            code,
            str(item.get("name") or item.get("label") or code).strip(),
            str(item.get("label") or item.get("name") or code).strip(),
            bool(item.get("uses_profiles")),
            str(item.get("application") or code).strip(),
        ))
    if not products:
        products = [
            (item["code"], item["name"], item["label"], bool(item["uses_profiles"]), item["application"])
            for item in DEFAULT_ADMIN_OPTIONS["products"]
        ]

    company_presets = []
    for item in options.get("company_presets") or []:
        if isinstance(item, dict):
            value = str(item.get("value") or item.get("mst") or item.get("name") or "").strip()
            label = str(item.get("label") or value).strip()
        else:
            value = str(item or "").strip()
            label = value
        if value:
            company_presets.append({"value": value, "label": label or value})
    return profiles, products, company_presets


class LicenseServerAdminWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("License Server Admin")
        self.resize(860, 680)
        self.profile_options, self.product_options, self.company_presets = load_admin_options()

        root = QWidget()
        layout = QVBoxLayout(root)

        intro = QLabel(
            "Tạo license key trên máy license server. Docker Desktop phải đang chạy, "
            "và Keygen stack nằm trong thư mục license_server. Vietmax hiện dùng profile chung vietmax."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        form = QFormLayout()
        self.product_combo = QComboBox()
        for code, _name, label, _uses_profiles, _application in self.product_options:
            self.product_combo.addItem(label, code)
        self.product_combo.currentIndexChanged.connect(self.update_product_mode)
        self.license_name_edit = QLineEdit("Customer license")
        self.policy_name_edit = QLineEdit("Default policy")
        form.addRow("Ứng dụng", self.product_combo)
        form.addRow("Tên license", self.license_name_edit)
        form.addRow("Tên policy", self.policy_name_edit)
        layout.addLayout(form)

        self.issued_licenses: list[dict[str, Any]] = []

        self.profiles_box = QGroupBox("Profile Product Code Formatter được phép")
        profiles_layout = QVBoxLayout(self.profiles_box)
        self.profile_checks = []
        for key, label in self.profile_options:
            check = QCheckBox(f"{label} ({key})")
            check.setChecked(True)
            check.setProperty("profile_key", key)
            self.profile_checks.append(check)
            profiles_layout.addWidget(check)
        layout.addWidget(self.profiles_box)

        issued_box = QGroupBox("License đã cấp / công ty được phép")
        issued_layout = QVBoxLayout(issued_box)
        issued_row = QHBoxLayout()
        self.issued_license_combo = QComboBox()
        self.issued_license_combo.currentIndexChanged.connect(lambda _index: self.load_selected_license_companies())
        self.reload_licenses_btn = QPushButton("Tải danh sách license")
        self.reload_licenses_btn.clicked.connect(self.reload_issued_licenses)
        issued_row.addWidget(self.issued_license_combo, 1)
        issued_row.addWidget(self.reload_licenses_btn)
        issued_layout.addLayout(issued_row)
        self.current_license_details = QTextEdit()
        self.current_license_details.setReadOnly(True)
        self.current_license_details.setMinimumHeight(110)
        issued_layout.addWidget(self.current_license_details)
        preset_row = QHBoxLayout()
        self.company_preset_combo = QComboBox()
        self.company_preset_combo.addItem("Chọn công ty mẫu để thêm...", "")
        for item in self.company_presets:
            self.company_preset_combo.addItem(item["label"], item["value"])
        self.add_company_preset_btn = QPushButton("Thêm công ty mẫu")
        self.add_company_preset_btn.clicked.connect(self.add_selected_company_preset)
        self.reload_admin_options_btn = QPushButton("Nạp lại danh sách")
        self.reload_admin_options_btn.clicked.connect(self.reload_admin_options)
        preset_row.addWidget(self.company_preset_combo, 1)
        preset_row.addWidget(self.add_company_preset_btn)
        preset_row.addWidget(self.reload_admin_options_btn)
        issued_layout.addLayout(preset_row)
        self.allowed_companies_edit = QTextEdit()
        self.allowed_companies_edit.setPlaceholderText("Nhập MST hoặc tên công ty được phép, mỗi dòng một giá trị. Để trống = không giới hạn công ty.")
        issued_layout.addWidget(self.allowed_companies_edit)
        self.update_license_companies_btn = QPushButton("Cập nhật profile/công ty cho license đã chọn")
        self.update_license_companies_btn.clicked.connect(self.update_selected_license_companies)
        issued_layout.addWidget(self.update_license_companies_btn)
        layout.addWidget(issued_box)

        actions = QHBoxLayout()
        self.create_btn = QPushButton("Tạo license")
        self.create_btn.clicked.connect(self.create_license)
        self.update_metadata_btn = QPushButton("Cập nhật metadata server")
        self.update_metadata_btn.clicked.connect(self.update_server_metadata)
        self.check_btn = QPushButton("Kiểm tra server")
        self.check_btn.clicked.connect(self.check_server)
        actions.addWidget(self.create_btn)
        actions.addWidget(self.update_metadata_btn)
        actions.addWidget(self.check_btn)
        actions.addStretch(1)
        layout.addLayout(actions)

        self.output = QTextEdit()
        self.output.setReadOnly(True)
        layout.addWidget(self.output, 1)

        self.setCentralWidget(root)
        self.update_product_mode()

    def selected_product(self):
        code = self.product_combo.currentData()
        for option in self.product_options:
            if option[0] == code:
                return option
        return self.product_options[0]

    def update_product_mode(self):
        _code, _name, _label, uses_profiles, _application = self.selected_product()
        self.profiles_box.setVisible(bool(uses_profiles))

    def selected_profiles(self):
        return [check.property("profile_key") for check in self.profile_checks if check.isChecked()]

    def selected_allowed_companies(self):
        text = self.allowed_companies_edit.toPlainText()
        values = []
        for line in text.replace(";", "\n").replace(",", "\n").splitlines():
            value = line.strip()
            if value:
                values.append(value)
        return values

    def reload_admin_options(self):
        try:
            self.profile_options, self.product_options, self.company_presets = load_admin_options()
            self.company_preset_combo.clear()
            self.company_preset_combo.addItem("Chọn công ty mẫu để thêm...", "")
            for item in self.company_presets:
                self.company_preset_combo.addItem(item["label"], item["value"])
            self.output.setPlainText(
                f"Đã nạp {len(self.company_presets)} công ty mẫu từ {ADMIN_OPTIONS_PATH.name}.\n"
                "Nếu vừa sửa danh sách profile/product, hãy mở lại admin app để dựng lại checkbox/dropdown."
            )
        except Exception as exc:
            QMessageBox.critical(self, "Nạp danh sách", str(exc))

    def add_selected_company_preset(self):
        value = str(self.company_preset_combo.currentData() or "").strip()
        if not value:
            return
        current = self.selected_allowed_companies()
        normalized = {item.casefold() for item in current}
        if value.casefold() not in normalized:
            current.append(value)
            self.allowed_companies_edit.setPlainText("\n".join(current))

    def product_metadata(self, product_code, uses_profiles, application, profiles):
        metadata: dict[str, Any] = {"application": application, "product_code": product_code}
        if uses_profiles:
            metadata["allowed_profiles"] = profiles
            metadata["supported_profiles"] = [key for key, _label in self.profile_options]
        if product_code == "product-code-formatter":
            metadata["desktop_build"] = CURRENT_DESKTOP_BUILD
        return metadata

    def profile_label(self, key):
        labels = {profile_key: label for profile_key, label in self.profile_options}
        return labels.get(str(key), str(key))

    def render_license_details(self, item):
        if not item:
            return "Chưa chọn license."
        profiles = item.get("allowed_profiles") or []
        companies = item.get("allowed_companies") or []
        supported = item.get("supported_profiles") or []
        if isinstance(profiles, str):
            profiles = [profiles]
        if isinstance(companies, str):
            companies = [companies]
        if isinstance(supported, str):
            supported = [supported]
        profile_text = ", ".join(f"{self.profile_label(value)} ({value})" for value in profiles) or "Không giới hạn"
        supported_text = ", ".join(str(value) for value in supported) or "Theo product metadata/default"
        company_text = "\n".join(f"  - {value}" for value in companies) if companies else "  Không giới hạn công ty"
        return (
            f"License: {item.get('name') or ''}\n"
            f"Key: {item.get('key') or ''}\n"
            f"Trạng thái: {item.get('status') or ''}\n"
            f"Sản phẩm: {item.get('product_code') or ''} | Policy: {item.get('policy') or ''}\n"
            f"Profile được phép: {profile_text}\n"
            f"Profile product hỗ trợ: {supported_text}\n"
            f"Công ty/MST được phép:\n{company_text}"
        )

    def run_docker(self, args):
        result = subprocess.run(
            args,
            cwd=LICENSE_SERVER_DIR,
            text=True,
            capture_output=True,
            timeout=120,
        )
        output = (result.stdout or "") + (result.stderr or "")
        if result.returncode != 0:
            raise RuntimeError(output.strip() or f"Lệnh lỗi: {' '.join(args)}")
        return output.strip()

    def marked_json(self, output, marker):
        prefix = f"{marker}="
        for line in str(output or "").splitlines():
            if line.startswith(prefix):
                return json.loads(line[len(prefix):])
        raise ValueError(f"Không tìm thấy dữ liệu {marker} trong output server:\n{output}")

    def check_server(self):
        try:
            output = self.run_docker(["docker", "compose", "ps"])
            self.output.setPlainText(output)
        except Exception as exc:
            QMessageBox.critical(self, "Kiểm tra server", str(exc))

    def create_license(self):
        product_code, product_name, _label, uses_profiles, application = self.selected_product()
        profiles = self.selected_profiles()
        if uses_profiles and not profiles:
            QMessageBox.warning(self, "Tạo license", "Chọn ít nhất một profile được phép.")
            return

        self.create_btn.setEnabled(False)
        QApplication.processEvents()
        try:
            license_name = self.license_name_edit.text().strip() or "Customer license"
            policy_name = self.policy_name_edit.text().strip() or "Default policy"
            profiles_json = json.dumps(profiles)
            license_name_json = json.dumps(license_name)
            policy_name_json = json.dumps(policy_name)
            product_code_json = json.dumps(product_code)
            product_name_json = json.dumps(product_name)
            product_metadata = self.product_metadata(product_code, uses_profiles, application, profiles)
            metadata_json = json.dumps(product_metadata, ensure_ascii=False)
            ruby = f"""
account = Account.find(ENV.fetch('KEYGEN_ACCOUNT_ID'))
profiles = JSON.parse({profiles_json!r})
metadata = JSON.parse({metadata_json!r})
product = account.products.find_or_create_by!(code: JSON.parse({product_code_json!r})) do |p|
  p.name = JSON.parse({product_name_json!r})
  p.metadata = {{}}
end
product.update!(name: JSON.parse({product_name_json!r}), metadata: metadata)
policy = account.policies.find_or_create_by!(product: product, name: JSON.parse({policy_name_json!r})) do |p|
  p.authentication_strategy = 'LICENSE'
  p.floating = false
  p.max_machines = 1
  p.metadata = metadata
end
policy.update!(metadata: metadata)
license = account.licenses.create!(
  policy: policy,
  name: JSON.parse({license_name_json!r}),
  metadata: metadata
)
keygen_host = ENV.fetch('KEYGEN_HOST', 'license-server.local')
client_server = keygen_host.start_with?('http') ? keygen_host : "http://#{{keygen_host}}:3000"
puts "CLIENT_APP_FIELDS=License server/IP + APP_LICENSE_KEY"
puts "APP_LICENSE_SERVER=#{{client_server}}"
puts "APP_ACCOUNT_ID=#{{ENV.fetch('KEYGEN_ACCOUNT_ID')}}"
puts "APP_LICENSE_KEY=#{{license.key}}"
puts "LICENSE_KEY=#{{license.key}}"
puts "LICENSE_ID=#{{license.id}}"
puts "PRODUCT_CODE=#{{product.code}}"
puts "ALLOWED_PROFILES=#{{profiles.join(',')}}" if profiles.any?
""".strip()
            output = self.run_docker(["docker", "compose", "exec", "-T", "keygen", "bin/rails", "runner", ruby])
            self.output.setPlainText(output)
        except Exception as exc:
            QMessageBox.critical(self, "Tạo license", str(exc))
        finally:
            self.create_btn.setEnabled(True)

    def update_server_metadata(self):
        product_code, product_name, _label, uses_profiles, application = self.selected_product()
        profiles = self.selected_profiles()
        if uses_profiles and not profiles:
            QMessageBox.warning(self, "Cập nhật metadata", "Chọn ít nhất một profile được phép.")
            return

        self.update_metadata_btn.setEnabled(False)
        QApplication.processEvents()
        try:
            policy_name = self.policy_name_edit.text().strip() or "Default policy"
            product_code_json = json.dumps(product_code)
            product_name_json = json.dumps(product_name)
            policy_name_json = json.dumps(policy_name)
            metadata_json = json.dumps(self.product_metadata(product_code, uses_profiles, application, profiles), ensure_ascii=False)
            ruby = f"""
account = Account.find(ENV.fetch('KEYGEN_ACCOUNT_ID'))
metadata = JSON.parse({metadata_json!r})
product = account.products.find_or_create_by!(code: JSON.parse({product_code_json!r})) do |p|
  p.name = JSON.parse({product_name_json!r})
  p.metadata = {{}}
end
product.update!(name: JSON.parse({product_name_json!r}), metadata: metadata)
policy = account.policies.find_or_create_by!(product: product, name: JSON.parse({policy_name_json!r})) do |p|
  p.authentication_strategy = 'LICENSE'
  p.floating = false
  p.max_machines = 1
  p.metadata = metadata
end
policy.update!(metadata: metadata)
updated_licenses = 0
account.licenses.includes(:policy).find_each do |lic|
  next unless lic.policy && lic.policy.product_id == product.id
  license_metadata = lic.metadata || {{}}
  metadata.each do |key, value|
    next if ['allowed_profiles', 'allowedProfiles', 'profiles', 'allowed_companies', 'allowedCompanies', 'companies'].include?(key)
    license_metadata[key] = value
  end
  lic.update!(metadata: license_metadata)
  updated_licenses += 1
end
puts "UPDATED_PRODUCT=#{{product.code}}"
puts "UPDATED_POLICY=#{{policy.name}}"
puts "UPDATED_LICENSES=#{{updated_licenses}}"
puts "METADATA=#{{metadata.to_json}}"
""".strip()
            output = self.run_docker(["docker", "compose", "exec", "-T", "keygen", "bin/rails", "runner", ruby])
            self.output.setPlainText(output)
        except Exception as exc:
            QMessageBox.critical(self, "Cập nhật metadata", str(exc))
        finally:
            self.update_metadata_btn.setEnabled(True)

    def reload_issued_licenses(self):
        product_code, _product_name, _label, _uses_profiles, _application = self.selected_product()
        self.reload_licenses_btn.setEnabled(False)
        QApplication.processEvents()
        try:
            product_code_json = json.dumps(product_code)
            ruby = f"""
require 'json'
account = Account.find(ENV.fetch('KEYGEN_ACCOUNT_ID'))
product_code = JSON.parse({product_code_json!r})
licenses = []
account.licenses.includes(:policy).find_each do |lic|
  policy = lic.policy
  product = policy && policy.product
  product_metadata = product && product.metadata || {{}}
  next unless product && (product.code == product_code || product_metadata['product_code'] == product_code)
  metadata = lic.metadata || {{}}
  policy_metadata = policy && policy.metadata || {{}}
  allowed_companies = metadata['allowed_companies'] || metadata['allowedCompanies'] || metadata['companies'] || []
  allowed_companies = allowed_companies.split(/[;,\n]+/).map(&:strip).reject(&:empty?) if allowed_companies.is_a?(String)
  allowed_profiles = metadata['allowed_profiles'] || metadata['allowedProfiles'] || metadata['profiles'] || []
  allowed_profiles = allowed_profiles.split(/[;,\n]+/).map(&:strip).reject(&:empty?) if allowed_profiles.is_a?(String)
  supported_profiles = product_metadata['supported_profiles'] || policy_metadata['supported_profiles'] || []
  supported_profiles = supported_profiles.split(/[;,\n]+/).map(&:strip).reject(&:empty?) if supported_profiles.is_a?(String)
    licenses << {{
    id: lic.id,
    key: lic.key,
    name: lic.name,
    status: lic.status,
    policy: policy && policy.name,
    product_code: product.code,
    allowed_profiles: allowed_profiles,
    allowed_companies: allowed_companies,
    supported_profiles: supported_profiles,
    metadata: metadata
    }}
end
puts "LICENSES_JSON=#{{JSON.generate(licenses)}}"
""".strip()
            output = self.run_docker(["docker", "compose", "exec", "-T", "keygen", "bin/rails", "runner", ruby])
            licenses = self.marked_json(output, "LICENSES_JSON")
            if not isinstance(licenses, list):
                licenses = []
            self.issued_licenses = [item for item in licenses if isinstance(item, dict)]
            self.issued_license_combo.clear()
            for item in self.issued_licenses:
                label = f"{item.get('name') or 'License'} | {item.get('key') or ''} | {item.get('status') or ''}"
                self.issued_license_combo.addItem(label, item.get("id"))
            if self.issued_licenses:
                self.issued_license_combo.setCurrentIndex(0)
                self.load_selected_license_companies()
            else:
                self.allowed_companies_edit.clear()
                self.current_license_details.setPlainText("Chưa có license nào cho sản phẩm đang chọn.")
            self.output.setPlainText(f"Đã tải {len(self.issued_licenses)} license đã cấp cho {product_code}.")
        except Exception as exc:
            QMessageBox.critical(self, "Tải danh sách license", str(exc))
        finally:
            self.reload_licenses_btn.setEnabled(True)

    def current_issued_license(self):
        license_id = self.issued_license_combo.currentData()
        for item in self.issued_licenses:
            if item.get("id") == license_id:
                return item
        return None

    def load_selected_license_companies(self):
        item = self.current_issued_license()
        self.current_license_details.setPlainText(self.render_license_details(item))
        profiles = item.get("allowed_profiles") if item else []
        if isinstance(profiles, str):
            profiles = [profiles]
        if not isinstance(profiles, list):
            profiles = []
        profile_set = set(str(value).strip() for value in profiles if str(value).strip())
        for check in self.profile_checks:
            check.setChecked(str(check.property("profile_key")) in profile_set if profile_set else bool(item))
        companies = item.get("allowed_companies") if item else []
        if isinstance(companies, str):
            companies = [companies]
        if not isinstance(companies, list):
            companies = []
        self.allowed_companies_edit.setPlainText("\n".join(str(value) for value in companies if str(value).strip()))

    def update_selected_license_companies(self):
        license_id = self.issued_license_combo.currentData()
        if not license_id:
            QMessageBox.warning(self, "Cập nhật công ty", "Bấm Tải danh sách license và chọn một license trước.")
            return
        companies = self.selected_allowed_companies()
        profiles = self.selected_profiles()
        if not profiles:
            QMessageBox.warning(self, "Cập nhật license", "Chọn ít nhất một profile được phép.")
            return
        self.update_license_companies_btn.setEnabled(False)
        QApplication.processEvents()
        try:
            license_id_json = json.dumps(str(license_id))
            companies_json = json.dumps(companies, ensure_ascii=False)
            profiles_json = json.dumps(profiles, ensure_ascii=False)
            ruby = f"""
require 'json'
account = Account.find(ENV.fetch('KEYGEN_ACCOUNT_ID'))
lic = account.licenses.find(JSON.parse({license_id_json!r}))
metadata = lic.metadata || {{}}
metadata['allowed_profiles'] = JSON.parse({profiles_json!r})
metadata.delete('allowedProfiles')
metadata.delete('profiles')
metadata['allowed_companies'] = JSON.parse({companies_json!r})
metadata.delete('allowedCompanies')
metadata.delete('companies')
lic.update!(metadata: metadata)
puts "UPDATED_LICENSE=#{{lic.key}}"
puts "ALLOWED_PROFILES=#{{metadata['allowed_profiles'].join(',')}}"
puts "ALLOWED_COMPANIES=#{{metadata['allowed_companies'].join(',')}}"
""".strip()
            output = self.run_docker(["docker", "compose", "exec", "-T", "keygen", "bin/rails", "runner", ruby])
            self.output.setPlainText(output)
            item = self.current_issued_license()
            if item is not None:
                item["allowed_profiles"] = profiles
                item["allowed_companies"] = companies
                self.current_license_details.setPlainText(self.render_license_details(item))
        except Exception as exc:
            QMessageBox.critical(self, "Cập nhật công ty", str(exc))
        finally:
            self.update_license_companies_btn.setEnabled(True)


def main():
    app = QApplication([])
    window = LicenseServerAdminWindow()
    window.show()
    app.exec()


if __name__ == "__main__":
    main()
