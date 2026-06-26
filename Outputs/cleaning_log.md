# Cleaning Log — Capstone Part 1

**Dataset:** `raw_orders.xlsx`  →  `cleaned_orders.xlsx`
**Raw row count:** 932
**Cleaned row count:** 912
**Tools used:** Microsoft Excel 365 (formulas, Remove Duplicates, conditional formatting, pivot tables)

---

## 1. Issues Found in Raw Data

### Text formatting
- `segment` column had **16 variants** of 4 underlying values (whitespace + case inconsistencies, e.g., "Consumer", "  Consumer ", "CONSUMER", "consumer")
- `region`: 15 variants of 4 values
- `category`: 13 variants of 3 values (including double-space "Office  Supplies")
- `sub_category`: 23 variants
- `ship_mode`: 11 variants of 4 values
- `payment_status`: 8 variants of 4 values
- `order_status`: 10 variants of 3 values
- `customer_name`: 5 rows with leading/trailing whitespace

### Date formatting
Raw `order_date` and `ship_date` contained 4 mixed formats simultaneously:

| Format       | order_date count | ship_date count |
|--------------|------------------|-----------------|
| yyyy-mm-dd   | 226              | 227             |
| dd-mm-yyyy   | 232              | 253             |
| mm/dd/yyyy   | 247              | 216             |
| dd Mon yyyy  | 227              | 236             |

- **22 rows** had `ship_date` earlier than `order_date` (impossible — flagged)
- 0 missing dates, 0 unparseable

### Duplicates
- **20 exact duplicate rows** (identical across all 21 columns)
- **24 rows** with duplicate `order_id` but conflicting data (12 unique IDs after dedup)

### Missing values
- `region`: 26 blanks (25 after dedup)
- `ship_mode`: 22 blanks (21 after dedup)
- `discount`: 18 blanks

### Discount issues
- **16 negative discount values** (e.g., -0.19, -0.23) — invalid
- **8 text strings** like `"70%"`, `"85%"` — out of range and stored as text
- **2 numeric values above 0.5** (0.55, 0.65) — out of allowed range

### Sales/profit calculation issues
- **54 rows** where `sales ≠ quantity × unit_price × (1 - discount)`
- **40 rows** where `profit ≠ sales − cost`

### Status logic anomalies
- **2 rows**: payment_status="Failed" with order_status="Completed" — logical contradiction
- Various combinations of cancelled/returned with paid/refunded statuses (expected mix; documented in QA report)

---

## 2. Cleaning Actions Performed

| # | Task | Action |
|---|------|--------|
| 1 | Text cleaning | Applied `TRIM + PROPER` to segment, region, state, city, category, sub_category, ship_mode, payment_status, order_status. Applied `TRIM` only to customer_name. Collapsed multiple internal spaces. |
| 2 | Date parsing | Detected format by character position, rebuilt with `DATE()`. All 932 rows parsed successfully into ISO format `yyyy-mm-dd`. |
| 3 | Date validation | Added `shipping_delay_days` (ship - order) and `date_issue_flag` columns. |
| 4 | Exact duplicates | Removed 20 exact duplicates using Excel's Remove Duplicates feature (kept first occurrence). |
| 5 | Conflicting duplicates | Flagged 24 rows with duplicate order_id using `duplicate_id_flag` column. NOT deleted. |
| 6 | Missing region/ship_mode | Filled blanks with "Unknown". |
| 7 | Cleaned discount | Added `cleaned_discount` column with valid range 0–50%. Out-of-range, negative, blank, and unparseable values → 0. |
| 8 | Discount flag | Added `discount_flag` column to track which rows were modified. |
| 9 | Order validity | Added `order_validity_flag` classifying each row for downstream reporting per business rules. |
| 10 | Recalculated sales/profit | Added `calculated_sales`, `calculated_profit`, `profit_margin` columns. Original `sales`/`profit` preserved for audit. |
| 11 | Date extraction | Added `order_month` and `order_year` for monthly trend analysis. |
| 12 | Quality rollup | Added `data_quality_flag` combining all individual flags into Clean / Warning / Invalid. |

---

## 3. Business Rules Applied

| Rule | Decision |
|------|----------|
| Missing region | Filled as "Unknown", flagged in quality report |
| Missing ship_mode | Filled as "Unknown", flagged in quality report |
| Missing discount | cleaned_discount = 0 |
| Negative discount | cleaned_discount = 0, flagged "Warning" |
| Discount above 50% | cleaned_discount = 0, flagged "Invalid" |
| Cancelled orders | order_validity_flag = "Cancelled - exclude" (not in completed sales) |
| Failed payments | order_validity_flag = "Payment failed - exclude" |
| Refunded payments | order_validity_flag = "Refunded - separate" (separate summary) |
| Returned orders | order_validity_flag = "Returned - review" |
| Ship date < order date | Flagged in date_issue_flag (22 rows); shipping_delay_days negative |
| Sales calculation mismatch | NOT auto-corrected. Original sales preserved; calculated_sales used for analysis. |

---

## 4. Assumptions Made

1. **Allowed discount range: 0% to 50%.** The brief did not specify; inferred from common retail practice and observed valid values (0–25%). Two valid-looking values (0.55, 0.65) are treated as out of range.
2. **Date format inference:**
   - `/`-separated dates are US `mm/dd/yyyy` (verified: no first-position values > 12)
   - `-`-separated 2-digit-first dates are EU `dd-mm-yyyy` (verified: first-position values up to 31)
   - 4-digit-first dates are ISO `yyyy-mm-dd`
   - `dd Mon yyyy` format dates are zero-padded (verified)
3. **Exact duplicates convey no extra information** — safe to remove silently.
4. **Conflicting order_ids cannot be programmatically resolved** without a source-of-truth system — flagged for manual review rather than auto-deduplicated.
5. **Sales/profit mismatches** may reflect manual overrides, promotions, or system errors. NOT auto-corrected; both original and calculated values preserved.
6. **Text "Unknown" for missing categorical data** is more useful for pivot tables than blank cells.

---

## 5. Records Removed

- **20 exact duplicate rows** removed via Remove Duplicates (kept first occurrence)

## 6. Records Flagged (NOT Removed)

| Flag | Count | Column |
|------|-------|--------|
| Ship before Order | 21 | `date_issue_flag` |
| Duplicate Order ID | 24 | `duplicate_id_flag` |
| Discount: Missing | 18 | `discount_flag` |
| Discount: Negative | 15 | `discount_flag` |
| Discount: Above range | 15 | `discount_flag` |
| Quality: Warning | 58 | `data_quality_flag` |
| Quality: Invalid | 15 | `data_quality_flag` |
| Order validity: Cancelled | 145 | `order_validity_flag` |
| Order validity: Returned | 94 | `order_validity_flag` |
| Order validity: Refunded | 34 | `order_validity_flag` |
| Order validity: Payment failed | 37 | `order_validity_flag` |

---

## 7. Limitations

1. **Conflicting duplicate order_ids cannot be auto-resolved** — manual review required by a domain expert.
2. **Sales/profit mismatches preserved as-is** — without access to source systems we cannot determine whether the original or recalculated value is correct.
3. **Discount range threshold of 50% is an assumption** — could be revised with stakeholder input.
4. **No validation against external systems** (e.g., product catalog, customer database) — internal consistency only.
5. **Date format inference relies on positional patterns observed in this dataset** — would need adjustment if new formats appear in future loads.
6. **"Unknown" for missing region/ship_mode** loses information — ideal would be to look up via order_id from a master system.

---

## 8. Files Produced

- `data/cleaned_orders.xlsx` — cleaned dataset with all 33 columns
- `outputs/data_quality_report.xlsx` — 8-sheet quality report
- `outputs/pivot_summary.xlsx` — pivot analyses (Task 8)
- `outputs/cleaning_log.md` — this file
- `screenshots/` — visual confirmations
