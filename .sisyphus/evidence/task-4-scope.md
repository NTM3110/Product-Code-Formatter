# Task 4 Template Scope Evidence

Converted to cached fields:
- `selectedProfileNote()` -> `selectedProfileNoteText`
- `selectedProfileLabel()` template usages -> `selectedProfileLabelText`
- `wordRuleCount()` template usages -> `wordRuleCountValue`
- `skippedCompanies()` template usages -> `skippedCompanyList`
- `productRowSummary(prod)` / `productPriceSummary(prod)` in the product modal -> precomputed `prod.rowSummary` / `prod.priceSummary`

Remaining methods:
- The original methods remain as compatibility wrappers for TypeScript code paths, but high-frequency template bindings no longer call them.
