# Cleaning Log — Orders Dataset

**Source file:** `raw_data/raw_orders.xlsx` (sheet `raw_orders`, 932 rows, 21 columns)
**Output file:** `outputs/cleaned_orders.xlsx` (912 rows, 40 columns)
**Tools used:** Python (pandas, openpyxl) to build the workbook; all derived/calculated columns are live Excel formulas, recalculated and verified with LibreOffice (zero formula errors).

---

## 1. Issues Found

| # | Issue | Where |
|---|---|---|
| 1 | Leading/trailing spaces and double spaces in text fields | `segment`, `region`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`, `customer_name` |
| 2 | Inconsistent capitalization (UPPER / lower / Mixed) | Same fields as above |
| 3 | Missing values | `region` (26 rows), `ship_mode` (22 rows), `discount` (18 rows) |
| 4 | Four different date text formats mixed in the same column | `order_date`, `ship_date` |
| 5 | `ship_date` earlier than `order_date` | 22 rows (raw) / 21 rows (after dedup) |
| 6 | Discount stored inconsistently — plain decimal (e.g. `0.2`) vs. percentage text (e.g. `"70%"`) | `discount` |
| 7 | Negative discount values | 16 rows (raw) |
| 8 | Discount values far above the normal tier (0/5/10/15/20/25%) — `0.55`, `0.65`, `0.70`, `0.85` | 15 rows (raw) |
| 9 | Exact duplicate rows (identical on every column) | 20 rows |
| 10 | Same `order_id` reused with conflicting field values (after removing exact duplicates) | 12 order_ids / 24 rows |
| 11 | Reported `sales` not matching `quantity × unit_price × (1 − discount)` | 64 rows |
| 12 | Reported `profit` not matching `calculated_sales − cost` | 64 rows |
| 13 | `order_status = Completed` but `payment_status ≠ Paid` (e.g. Failed) | 2 rows |

---

## 2. Cleaning Actions Performed

- **Text fields** (`customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`, `product_name`): trimmed leading/trailing spaces, collapsed repeated internal spaces to a single space, and standardized casing to Title Case (equivalent to Excel `=PROPER(TRIM(SUBSTITUTE(x," ","  ")))`-style cleanup). This alone resolved every "similar category written differently" case observed (e.g. `"SMALL BUSINESS"`, `"  Small  Business "`, `"small business"` all became `"Small Business"`).
- **Dates**: every value in `order_date` / `ship_date` was tested against four known raw formats — `d Mon yyyy`, `MM/DD/YYYY`, `DD-MM-YYYY`, `YYYY-MM-DD` — and converted to a single Excel date type, displayed as `yyyy-mm-dd`. No row failed to parse (0 invalid/unrecognized dates), so no row needed to be dropped for this reason.
- **Discount**: standardized into `cleaned_discount`, a decimal fraction. Percentage text such as `"70%"` is converted to `0.70`; plain numbers are passed through; missing values are filled with `0`.
- **Duplicates**: removed only true exact-duplicate rows (20). Conflicting duplicate `order_id`s were **not deleted** — every row is kept and flagged (see Section 3 and the `duplicate_summary` sheet of `data_quality_report.xlsx`) for manual review, since it isn't possible to know which of the conflicting versions is correct without going back to source systems.
- **Calculated columns added** (all live Excel formulas, see `column_notes` sheet in `cleaned_orders.xlsx` for the exact logic): `cleaned_discount`, `discount_valid_flag`, `calculated_sales`, `sales_mismatch_flag`, `calculated_profit`, `profit_mismatch_flag`, `profit_margin`, `shipping_delay_days`, `invalid_shipping_flag`, `order_month`, `order_year`, `status_inconsistency_flag`, `duplicate_order_id_flag`, `data_quality_flag`.

---

## 3. Business Rules Applied

| Rule Area | Action Taken |
|---|---|
| Missing `region` | Filled as `"Unknown"`; flagged `region_filled_flag = Yes` (25 rows in the cleaned file) |
| Missing `ship_mode` | Filled as `"Unknown"`; flagged `ship_mode_filled_flag = Yes` (21 rows) |
| Missing `discount` | Treated as `0` **only after confirming** `quantity`, `unit_price`, `sales`, `cost`, and `profit` were all present and valid for every one of these 18 rows (verified — none were missing or zero/negative). Flagged `discount_filled_flag = Yes`. |
| Negative discount | Flagged `discount_valid_flag = Invalid` |
| Discount above allowed range | Allowed range adopted as **0%–25%**, based on the clear natural tiering in the data (0, 5, 10, 15, 20, 25%); values of 55%, 65%, 70%, 85% sit far outside this tiering and are flagged `Invalid`. *(Assumption — see Section 4.)* |
| Cancelled orders | Excluded from the completed-sales pivots in `pivot_summary.xlsx` |
| Failed payments | Excluded from the completed-sales pivots |
| Refunded orders | Summarized separately in the `refund_cancel_fail_by_region` pivot sheet, not mixed into completed sales |
| Ship date before order date | Flagged `invalid_shipping_flag = Invalid` (21 rows after dedup) |
| Duplicates | Exact duplicates removed; conflicting duplicates flagged, never silently deleted |

The "final completed sales summary" basis used throughout `pivot_summary.xlsx` is: **`order_status = Completed` AND `payment_status = Paid` AND `data_quality_flag ≠ Invalid`** (537 of 912 rows). This deliberately also excludes the 2 rows where `order_status = Completed` but `payment_status = Failed`, since a completed order with a failed payment should not count as a completed sale.

---

## 4. Assumptions Made

1. **Discount valid range = 0%–25%.** Not stated explicitly in the source brief; inferred from the data's own discount tiers. If the real business rule allows higher discounts (e.g. seasonal sales up to 50%), this threshold should be adjusted and the `discount_valid_flag` formula updated accordingly.
2. **Sales/profit mismatch tolerance = ₹1.** Reported `sales`/`profit` differing from the recalculated value by ₹1 or less is treated as floating-point rounding, not an error.
3. **"Completed sales" basis** requires `payment_status = Paid` in addition to `order_status = Completed`, since a completed order with a failed/pending/refunded payment is not really a completed sale.
4. Text standardization assumes Title Case is the desired final format for all categorical fields (e.g. `"Office Supplies"`, `"First Class"`).
5. `customer_id` and `product_name` were left otherwise untouched apart from whitespace trimming — no case standardization was applied since they are identifiers/free-text product labels, not fixed categories.

---

## 5. Records Removed

- **20 exact duplicate rows** removed (kept the first occurrence of each fully-identical row). No other rows were deleted. The raw file is preserved unmodified at `raw_data/raw_orders.xlsx`.

## 6. Records Flagged

- **24 rows (12 order_ids)** flagged `duplicate_order_id_flag = Yes` — conflicting duplicate order IDs, kept for manual review.
- **30 rows** flagged `discount_valid_flag = Invalid` (negative or > 25%).
- **21 rows** flagged `invalid_shipping_flag = Invalid` (ship date before order date).
- **64 rows** flagged `sales_mismatch_flag = Mismatch` (and the same 64 also flagged `profit_mismatch_flag = Mismatch`, since profit is derived from sales).
- **2 rows** flagged `status_inconsistency_flag = Inconsistent`.
- **Overall `data_quality_flag`:** 756 Clean / 44 Warning / 112 Invalid (of 912 total).

## 7. Limitations of This Cleaning Process

- Conflicting duplicate order_ids could not be resolved automatically — there is no timestamp or "last updated" field to determine which version of a duplicated order is authoritative, so all versions are retained and flagged rather than guessed at.
- The 25%-discount threshold is a data-driven inference, not a confirmed business policy; it should be validated with the actual discount policy if available.
- Rows with both a sales mismatch *and* an invalid discount cannot be cleanly attributed to a single root cause from the data alone (e.g. it is unclear whether the discount or the sales figure was the original data-entry error).
- `customer_id` was not cross-checked against `customer_name` for consistency (e.g. whether the same customer_id always maps to the same name) — this was outside the scope of the stated cleaning tasks.
