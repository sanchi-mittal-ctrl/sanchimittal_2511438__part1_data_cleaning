# Orders Data Cleaning & Analysis — README

## 1. Problem Summary

The raw `orders` export contains inconsistent text formatting, mixed date formats, missing values, duplicate records, and discount/sales/profit figures that don't always agree with each other. The goal of this project is to clean and validate the dataset, document every cleaning decision, apply the agreed business rules, and produce analysis-ready outputs (a cleaned dataset, a data quality report, and pivot-style summaries) for downstream reporting.

## 2. Dataset Description

| | |
|---|---|
| Source file | `raw_data/raw_orders.xlsx` (sheet `raw_orders`) |
| Rows | 932 |
| Columns | 21 — order/ship dates, customer & location info, product category/sub-category, ship mode, quantity, unit price, discount, sales, cost, profit, payment status, order status |
| Reference sheet | `business_rules` (the agreed cleaning rules, also reproduced in Section 5 below) |
| Granularity | One row per order line |

## 3. Tools Used

- **Python** (pandas, openpyxl) to inspect, clean, and rebuild the workbook
- **Excel formulas** for every calculated/flag column in `cleaned_orders.xlsx` (so the file stays dynamic and re-checkable — not hardcoded values)
- **LibreOffice** (headless) to recalculate and verify all formulas — final file has **zero formula errors** across 12,768 formulas

## 4. Cleaning Steps Performed

1. Trimmed extra spaces and standardized casing (Title Case) across `customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`, `product_name`.
2. Parsed `order_date` / `ship_date` out of four mixed raw formats into a single consistent date type.
3. Standardized `discount` into a decimal fraction (`cleaned_discount`), handling percentage-text values like `"70%"`.
4. Removed 20 exact duplicate rows (identical across every column).
5. Flagged — but did not delete — 24 rows (12 order_ids) where the same `order_id` appears with conflicting data.
6. Added calculated columns: `cleaned_discount`, `calculated_sales`, `calculated_profit`, `profit_margin`, `shipping_delay_days`, `order_month`, `order_year`, plus supporting validity/mismatch flags and an overall `data_quality_flag`.

Full details and rationale are in [`outputs/cleaning_log.md`](outputs/cleaning_log.md).

## 5. Business Rules Applied

| Rule Area | Action |
|---|---|
| Missing region | Filled `"Unknown"`, flagged |
| Missing ship_mode | Filled `"Unknown"`, flagged |
| Missing discount | Filled `0` (only since all other sales fields were valid for those rows) |
| Negative / out-of-range discount | Flagged invalid (valid range adopted: 0%–25%) |
| Cancelled orders | Excluded from completed-sales pivots |
| Failed payments | Excluded from completed-sales pivots |
| Refunded orders | Summarized separately |
| Ship date before order date | Flagged invalid |
| Duplicates | Exact duplicates removed; conflicting duplicates flagged for manual review, never silently deleted |

## 6. Summary of Data Quality Issues Found

| Issue | Count |
|---|---|
| Exact duplicate rows (removed) | 20 |
| Conflicting duplicate order_ids (flagged) | 12 order_ids / 24 rows |
| Missing region (filled Unknown) | 25 |
| Missing ship_mode (filled Unknown) | 21 |
| Missing discount (filled 0) | 18 |
| Invalid discount (negative or >25%) | 30 |
| Ship date earlier than order date | 21 |
| Sales ≠ quantity × unit price × (1 − discount) | 64 |
| Profit ≠ calculated sales − cost | 64 |
| order_status = Completed but payment ≠ Paid | 2 |
| **Overall record status** | **756 Clean / 44 Warning / 112 Invalid** (of 912 records after dedup) |

See [`outputs/data_quality_report.xlsx`](outputs/data_quality_report.xlsx) for the full breakdown across 7 sheets.

## 7. Summary of Final Pivot Reports

All "completed sales" pivots use the basis: `order_status = Completed` AND `payment_status = Paid` AND `data_quality_flag ≠ Invalid` (537 of 912 records).

- **Sales & profit by region** — South leads on both total sales (₹14.27L) and total profit (₹4.08L); Unknown-region orders (25 records with missing region) are isolated so they don't distort regional totals.
- **Sales & profit by category/sub-category** — Technology is the strongest category overall; Copiers is the single best-selling sub-category, Labels the weakest.
- **Order count by ship mode** — Standard Class is the most-used ship mode (242 of 912 orders); 21 orders had a missing ship mode (filled "Unknown").
- **Profit margin by customer segment** — Home Office customers carry the highest profit margin (~30.2%); Small Business the lowest (~28.1%), though the spread across segments is narrow.
- **Refunded / Cancelled / Failed orders by region** — Broken out by issue type per region so non-completed activity doesn't get mixed into revenue figures.
- **Monthly sales trend** — Chronological month-by-month view of completed sales and profit across 2024–2025.

See [`outputs/pivot_summary.xlsx`](outputs/pivot_summary.xlsx) for all 6 pivot sheets (sortable/filterable via Excel AutoFilter).

## 8. Key Business Insights

- About **17% of completed orders carried an invalid discount, ship-date, or sales/profit mismatch** — these were excluded from the "Clean" bucket and should be reviewed before being trusted for financial reporting.
- **Technology** is the highest-revenue category, driven mainly by Copiers and Phones.
- **Standard Class** dominates shipping volume, but **Same Day** orders make up roughly a fifth of all shipments — worth checking if Same Day capacity matches that demand.
- Profit margins are fairly consistent (~28–31%) across customer segments and regions, suggesting pricing/discounting discipline is broadly even rather than concentrated in one segment.
- The 25 "Unknown"-region orders and 21 "Unknown"-ship-mode orders are small in volume but should be traced back to source systems to close the data-entry gap going forward.

## 9. Assumptions and Limitations

- Discount valid range (0%–25%) is inferred from the data's own tiering, not a confirmed written policy — see `cleaning_log.md` Section 4 for the full assumptions list.
- Conflicting duplicate order_ids (12 of them) could not be auto-resolved; all versions are kept and flagged for manual review since there's no timestamp field to tell which is authoritative.
- Sales/profit mismatch tolerance is ₹1, to absorb rounding rather than treat every fractional difference as an error.
- "Completed sales" basis requires `payment_status = Paid`, which excludes the 2 records marked Completed with a Failed payment — these are flagged as inconsistent rather than corrected, since neither value can be assumed correct.

## 10. Screenshots

| File | Shows |
|---|---|
| `screenshots/raw_data_preview.png` | Raw dataset rows illustrating typical issues (spacing, case, missing values, percentage-text discounts) before cleaning |
| `screenshots/cleaned_data_preview.png` | The same rows after cleaning, with calculated columns and quality flags |
| `screenshots/pivot_summary_1.png` | Sales & profit by region |
| `screenshots/pivot_summary_2.png` | Sales & profit by category and sub-category |

## 11. Repository Structure

```
.
├── README.md
├── raw_data/
│   └── raw_orders.xlsx
├── outputs/
│   ├── cleaned_orders.xlsx
│   ├── data_quality_report.xlsx
│   ├── pivot_summary.xlsx
│   └── cleaning_log.md
└── screenshots/
    ├── raw_data_preview.png
    ├── cleaned_data_preview.png
    ├── pivot_summary_1.png
    └── pivot_summary_2.png
```
