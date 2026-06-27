"""Small workflow registry for profile ownership.

This module is deliberately dependency-free. It gives future changes a stable
place to describe whether a profile belongs to the product-code formatter,
inventory allocator, or a future workflow such as estimate extraction.
"""

WORKFLOW_PRODUCT_CODE_FORMATTER = "product_code_formatter"
WORKFLOW_INVENTORY_ALLOCATION = "inventory_allocation"
WORKFLOW_ESTIMATE_EXTRACTION = "estimate_extraction"

WORKFLOW_REGISTRY = {
    "vietmax": {
        "label": "Vietmax",
        "workflow": WORKFLOW_PRODUCT_CODE_FORMATTER,
        "native_react": True,
        "notes": "Unified purchase/sales formatter plus inventory allocation.",
    },
    "cao_thanh": {
        "label": "Cao Thanh",
        "workflow": WORKFLOW_PRODUCT_CODE_FORMATTER,
        "native_react": False,
        "notes": "Sales price filtering workflow; migrate in isolated slices.",
    },
    "son_phuong": {
        "label": "Son Phuong",
        "workflow": WORKFLOW_PRODUCT_CODE_FORMATTER,
        "native_react": False,
        "notes": "Legacy-compatible company profile; confirm flow before changes.",
    },
    "quang_thinh": {
        "label": "Quang Thinh",
        "workflow": WORKFLOW_PRODUCT_CODE_FORMATTER,
        "native_react": False,
        "notes": "Legacy-compatible company profile; confirm flow before changes.",
    },
    "yen_thanh": {
        "label": "Yen Thanh",
        "workflow": WORKFLOW_INVENTORY_ALLOCATION,
        "native_react": True,
        "notes": "Inventory/report behavior used as allocation reference.",
    },
    "boc_tach_du_toan": {
        "label": "Boc tach du toan",
        "workflow": WORKFLOW_ESTIMATE_EXTRACTION,
        "native_react": False,
        "notes": "Future workflow; keep separate from product-code stages.",
    },
}


def workflow_for_profile(profile_key):
    """Return registry metadata for a normalized profile key."""

    return WORKFLOW_REGISTRY.get(str(profile_key or "").strip().lower())
