# Cleaning Log

## Issues Found
- Raw file contained 932 rows and 21 columns.
- Exact duplicate copies found: 20.
- Conflicting duplicate order ID records flagged: 25.
- Missing region values filled as Unknown: 26.
- Missing ship_mode values filled as Unknown: 22.
- Missing discounts treated as 0 where sales fields were valid: 18.
- Invalid discounts flagged: 30.
- Date issues flagged: 93.
- Sales/profit calculation mismatches flagged: 83.

## Cleaning Actions Performed
- Preserved `data/raw_orders.xlsx` and created a separate cleaned file.
- Standardized whitespace, casing, and unwanted special characters in customer, location, category, shipping, payment, and order status text fields.
- Converted order and ship dates to consistent Excel date values.
- Added calculated fields for cleaned discount, calculated sales, calculated profit, profit margin, shipping delay, order month, order year, and data quality flag.
- Removed exact duplicate copies only after retaining the first occurrence.
- Kept conflicting duplicate order IDs and flagged them for review.

## Business Rules Applied
- Missing region and ship mode were replaced with Unknown and kept visible in the quality report.
- Missing discount was treated as 0 only when quantity, unit price, sales, cost, and profit were present.
- Negative discounts and discounts above 50% were marked invalid.
- Cancelled, returned, failed-payment, and refunded records were excluded from completed-sales summaries.
- Refunded orders were summarized separately in `outputs/pivot_summary.xlsx`.
- Ship dates earlier than order dates were marked invalid.

## Assumptions Made
- Discounts are expected to be decimal percentages from 0 to 0.50 inclusive; values such as 70% and 85% were treated as invalid.
- Returned orders are non-completed orders and therefore excluded from completed-sales summaries.
- `sales` should equal `quantity * unit_price * (1 - cleaned_discount)`.
- `profit` should equal `calculated_sales - cost`.

## Records Removed or Flagged
- Records removed: 20 exact duplicate copies.
- Records flagged invalid: 179.
- Records flagged warning: 279.
- Final rows in cleaned file: 912.

## Limitations
- Text normalization uses deterministic rules and simple mappings; it does not infer ambiguous customer or product identity changes.
- Invalid records are flagged rather than manually corrected when no reliable replacement value is available.
- Pivot summaries are generated as evaluator-friendly summary sheets rather than native Excel PivotTable cache objects.
