# Retail Orders — Data Cleaning Capstone (Part 1)

## Problem Summary

A retail company exports order-level sales data from multiple internal systems. The raw export contains numerous data quality issues — inconsistent text formatting, four different date formats mixed within the same columns, duplicate records, missing values, invalid discounts, sales/profit calculation mismatches, and logically contradictory order/payment statuses.

This project produces an analysis-ready cleaned dataset, documents every issue and decision, and delivers pivot summaries for business review.

---

## Dataset Description

| Property | Value |
|---|---|
| Source file | `raw_orders.xlsx` |
| Raw rows | 932 |
| Cleaned rows | 912 (20 exact duplicates removed) |
| Original columns | 21 |
| Final columns | 33 (12 added during cleaning) |
| Date range | 2024-01-02 to 2025-10-31 |
| Geographic scope | 4 regions (East, North, South, West) + Unknown |
| Categories | Furniture, Office Supplies, Technology |
| Segments | Consumer, Corporate, Home Office, Small Business |

**Original columns:** order_id, order_date, ship_date, customer_id, customer_name, segment, region, state, city, category, sub_category, product_name, ship_mode, quantity, unit_price, discount, sales, cost, profit, payment_status, order_status

**Columns added during cleaning:** shipping_delay_days, date_issue_flag, duplicate_id_flag, cleaned_discount, discount_flag, order_validity_flag, calculated_sales, calculated_profit, profit_margin, order_month, order_year, data_quality_flag

---

## Tools Used

- **Microsoft Excel 365** — primary cleaning environment
  - Formulas: `LET`, `IF`, `IFS`, `IFERROR`, `COUNTIF`, `MATCH`, `DATE`, `MONTH`, `YEAR`, `VALUE`, `SUBSTITUTE`, `MID`, `LEFT`, `RIGHT`, `ISNUMBER`, `ISBLANK`
  - Features: Remove Duplicates, Go To Special → Blanks, Paste Special (Multiply), Custom Number Formats, Conditional Formatting, PivotTables, PivotCharts
- **No external tools or scripts used** — all transformations are reproducible from formulas visible in the cleaned file

---

## Cleaning Steps Performed

| Step | Task | Technique |
|---|---|---|
| 1 | Standardize text fields | `TRIM` + `PROPER` collapsed 16 segment variants → 4, 15 region variants → 4, etc. across 10 text columns |
| 2 | Parse mixed-format dates | Position-based format detection (4 formats), rebuilt with `DATE()` into ISO `yyyy-mm-dd` |
| 3 | Add date validation | `shipping_delay_days` and `date_issue_flag` columns; flagged 21 ship-before-order rows |
| 4 | Remove exact duplicates | Excel Remove Duplicates feature; 20 rows removed (kept first occurrence) |
| 5 | Flag conflicting duplicates | `COUNTIF` formula identifies 24 rows with same order_id but different data — preserved and flagged for review |
| 6 | Fill missing categorical | `region` and `ship_mode` blanks set to `"Unknown"` via Go To Special → Blanks |
| 7 | Clean discount values | Multi-branch formula handles numeric values, text percentages (`"70%"`), blanks, negatives, and out-of-range values |
| 8 | Classify order validity | Combined order_status × payment_status logic into single `order_validity_flag` |
| 9 | Recalculate sales/profit | `calculated_sales = qty × unit_price × (1 − cleaned_discount)`; `calculated_profit = calculated_sales − cost`; originals preserved for audit |
| 10 | Add temporal fields | `order_month` and `order_year` for monthly trend analysis |
| 11 | Rollup quality flag | `data_quality_flag` combines all individual flags → Clean / Warning / Invalid |

---

## Business Rules Applied

| Rule Area | Decision |
|---|---|
| Missing `region` | Filled as `"Unknown"`, flagged in quality report |
| Missing `ship_mode` | Filled as `"Unknown"`, flagged in quality report |
| Missing `discount` | `cleaned_discount = 0` (assume no discount applied) |
| Negative `discount` | `cleaned_discount = 0`, classified as Warning |
| `Discount > 50%` | `cleaned_discount = 0`, classified as Invalid |
| Cancelled orders | Excluded from completed sales summaries |
| Failed payments | Excluded from completed sales summaries |
| Refunded orders | Tracked in a separate summary |
| `ship_date < order_date` | Flagged in `date_issue_flag`, shipping_delay_days shows as negative |
| Sales/profit mismatches | NOT auto-corrected. Original values preserved; calculated values used for analysis |

---

## Summary of Data Quality Issues Found

| Issue category | Count | Resolution |
|---|---|---|
| Text formatting (whitespace, case) | All 10 text columns affected | Standardized via TRIM + PROPER |
| Mixed date formats (4 different formats) | 932 rows, all columns | Parsed to ISO yyyy-mm-dd |
| Exact duplicate rows | 20 | Removed |
| Conflicting duplicate order_ids | 24 rows (12 unique IDs) | Flagged for review |
| Missing region | 26 raw / 25 after dedup | Filled "Unknown" |
| Missing ship_mode | 22 raw / 21 after dedup | Filled "Unknown" |
| Missing discount | 18 raw | Set to 0 |
| Negative discount | 16 raw / 15 after dedup | Set to 0, flagged |
| Discount > 50% (incl. text "70%", "85%") | 10 raw | Set to 0, flagged |
| Sales calculation mismatch | 54 rows | Preserved; calculated_sales used downstream |
| Profit calculation mismatch | 40 rows | Preserved; calculated_profit used downstream |
| Ship date before order date | 21 rows | Flagged |
| Payment Failed + Order Completed | 2 rows | Flagged as logical contradiction |

**Final quality breakdown:** Clean: **839 (92.0%)**, Warning: **58 (6.4%)**, Invalid: **15 (1.6%)**

Full details in `outputs/data_quality_report.xlsx`.

---

## Summary of Final Pivot Reports

`outputs/pivot_summary.xlsx` contains six pivot summaries:

| # | Pivot | Insight focus |
|---|---|---|
| 1 | Sales & Profit by Region | Compares revenue and profitability across 4 regions; filtered to exclude cancelled/failed orders |
| 2 | Sales & Profit by Category & Sub-category | Hierarchical drill-down; sub-categories sorted by profit |
| 3 | Order Count by Ship Mode | Volume distribution across shipping options |
| 4 | Avg Profit Margin by Customer Segment | Identifies most profitable customer types |
| 5 | Problem Orders by Region | Crosstab of refunded/cancelled/failed counts by region |
| 6 | Monthly Sales Trend | Time-series view with line chart |

---

## Key Business Insights

### Regional performance is remarkably even
The four main regions all generate within ~$300K of each other:

| Region | Sales | Profit | Margin |
|---|---|---|---|
| South | $1.88M | $546K | 29.0% |
| West | $1.80M | $515K | 29.0% |
| East | $1.77M | $541K | 30.0% |
| North | $1.59M | $484K | 30.0% |
| Unknown | $245K | $81K | 33.0% |

**Implication:** No region requires emergency intervention, but North is the weakest performer and could be a growth opportunity. The 26 orders with "Unknown" region should be reassigned to improve reporting accuracy.

### Technology leads, but margins are tight across categories
| Category | Sales | Profit | Margin |
|---|---|---|---|
| Technology | $2.57M | $779K | 30.3% |
| Furniture | $2.44M | $724K | 29.6% |
| Office Supplies | $2.27M | $663K | 29.2% |

**Implication:** Margins are flat at ~30% across all three categories — pricing or discount strategy is uniform across the catalog.

### Top sub-categories by profit
1. **Copiers** — $221K profit
2. **Chairs** — $207K
3. **Furnishings** — $200K
4. **Accessories** — $190K
5. **Phones** — $188K

### Customer segments are tightly clustered
| Segment | Total Sales | Avg Margin | Orders |
|---|---|---|---|
| Home Office | $1.86M | 30.3% | 190 |
| Consumer | $1.84M | 30.2% | 179 |
| Corporate | $1.55M | 29.6% | 165 |
| Small Business | $2.04M | 28.8% | 196 |

**Implication:** Small Business is the highest-revenue segment but with the lowest margin — likely higher discounting or higher-cost product mix. Worth investigating whether tighter discount discipline could boost segment profitability.

### Almost a quarter of orders are problematic
**216 orders (23.7% of total)** fall into cancelled, payment-failed, or refunded categories. This is a notable operational concern:
- 145 cancelled
- 37 failed payments
- 34 refunded

These are spread fairly evenly across regions, suggesting the issue is process-related rather than geographic. Recommended next step: investigate root causes of cancellations and failed payments.

### Sales trend is seasonal with no clear growth trajectory
- **Peak month:** June 2024 ($472K)
- **2024 total:** ~$4.00M
- **2025 total (Jan–Oct):** ~$3.29M (on pace for similar full-year if Nov/Dec match 2024 pattern)
- October 2025 ($133K) is unusually low — likely a partial month in the dataset rather than a business issue

---

## Assumptions and Limitations

### Assumptions
1. **Allowed discount range: 0% to 50%.** The brief did not specify; inferred from observed valid values (0–25% being typical) and common retail practice.
2. **Date format conventions:** `/` separators are US `mm/dd/yyyy`; `-` separators with 2-digit-first are EU `dd-mm-yyyy`; verified by checking that no first-positions exceed 12 (for `/`) or 31 (for `-`).
3. **Exact duplicates** convey no information and were removed silently.
4. **Conflicting duplicate order_ids** cannot be programmatically resolved without source-system access — preserved and flagged.
5. **Sales/profit mismatches** may reflect legitimate manual overrides, promotions, or pricing exceptions; not auto-corrected.
6. **"Unknown"** is more useful than blank for downstream pivot tables (blank rows often display awkwardly).

### Limitations
1. **Conflicting order_ids require manual reconciliation** by a domain expert with access to source systems.
2. **Calculation mismatches preserved** — we cannot determine which of (original sales, calculated sales) is correct without source-of-truth data.
3. **Discount range threshold** is an assumption — subject to revision with stakeholder input.
4. **No cross-system validation** — internal consistency only; no validation against product catalog or customer master data.
5. **Partial October 2025 data** in monthly trend should not be interpreted as a sales decline.
6. **"Unknown" region/ship_mode** loses information; ideally these would be looked up via order_id from a master operations system.

---

## Repository Contents

```
part1_data_cleaning/
├── data/
│   ├── raw_orders.xlsx              Original data (untouched)
│   └── cleaned_orders.xlsx          Cleaned dataset, 912 rows × 33 cols
├── outputs/
│   ├── data_quality_report.xlsx     8-sheet quality audit
│   ├── pivot_summary.xlsx           6 pivot summaries + chart
│   └── cleaning_log.md              Full audit trail
├── screenshots/
│   ├── raw_data_preview.png         Raw data before cleaning
│   ├── cleaned_data_preview.png     Cleaned data with calculated columns
│   ├── pivot_summary_1.png          By Region pivot (filtered)
│   └── pivot_summary_2.png          Monthly Trend pivot + chart
└── README.md                        This file
```

### Screenshots Included

| File | Shows |
|---|---|
| `raw_data_preview.png` | First ~15 rows of raw_orders.xlsx before any cleaning — visible mixed date formats, whitespace issues, case inconsistencies |
| `cleaned_data_preview.png` | Cleaned data showing standardized text, ISO-format dates, and the calculated columns (V–AG) including data_quality_flag with colour coding |
| `pivot_summary_1.png` | "By Region" pivot table with `order_validity_flag` filter applied (Cancelled and Payment failed excluded), sorted by sales descending |
| `pivot_summary_2.png` | "Monthly Trend" pivot table alongside its line chart showing sales trajectory across 22 months |

---

## Reproducing This Work

1. Open `data/cleaned_orders.xlsx` in Microsoft Excel 365
2. Click any cell in columns V–AG to inspect the formula behind each calculated value
3. The formulas are self-explanatory and reference only the original raw columns plus other calculated columns
4. To re-run the cleaning on new data: replace `data/raw_orders.xlsx` and re-execute the steps documented in `outputs/cleaning_log.md`
