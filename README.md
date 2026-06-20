# Part 1: Business Data Cleaning, Validation & Excel Reporting

## Business Problem Summary
This repository cleans and validates retail order-level sales data exported from multiple internal systems. The goal is to preserve the original raw file, create an analysis-ready cleaned workbook, document quality issues, and provide Excel summary reports for business review.

## Dataset Used
- Source dataset: `data/raw_orders.xlsx`
- Cleaned dataset: `data/cleaned_orders.xlsx`
- Raw sheet used: `raw_orders`
- Business rules sheet used: `business_rules`
- Raw dataset size: 932 rows and 21 columns

## Tools Used
- Microsoft Excel-compatible `.xlsx` workbooks
- Python for repeatable cleaning, validation, pivot summaries, and screenshots
- pandas, openpyxl, Pillow

## Steps Performed
1. Created the required repository folders: `data/`, `outputs/`, and `screenshots/`.
2. Copied the original dataset into `data/raw_orders.xlsx` without modifying it.
3. Standardized text fields including customer, segment, region, city, category, ship mode, payment status, and order status.
4. Converted `order_date` and `ship_date` into consistent date values.
5. Calculated `shipping_delay_days`, `cleaned_discount`, `calculated_sales`, `calculated_profit`, `profit_margin`, `order_month`, `order_year`, and `data_quality_flag`.
6. Removed exact duplicate copies and flagged duplicate order IDs with conflicting information.
7. Applied business rules for missing region, missing ship mode, missing discount, invalid discounts, cancelled orders, failed payments, refunded orders, and invalid shipping dates.
8. Created `outputs/data_quality_report.xlsx` with missing value, duplicate, discount, date, status, mismatch, and clean-vs-flagged summaries.
9. Created `outputs/pivot_summary.xlsx` with required business pivot summaries.
10. Created screenshots as evidence of raw data, cleaned data, and final pivot outputs.

## Key Outputs
- `data/cleaned_orders.xlsx`: cleaned, validated, analysis-ready order data.
- `outputs/data_quality_report.xlsx`: quality summaries and issue lists.
- `outputs/pivot_summary.xlsx`: business summary tables for sales, profit, orders, margins, exceptions, and monthly trends.
- `outputs/cleaning_log.md`: detailed cleaning decisions and assumptions.
- `screenshots/`: required evidence screenshots.

## Business Insights
- Highest completed-sales region: West with sales of 1,394,860.73.
- Highest completed-sales profit margin segment: Home Office with margin of 30.15%.
- Refunded records identified for separate review: 71.
- Cancelled records identified: 145; failed-payment records identified: 69.
- Invalid discount records flagged: 30; these should be reviewed before any final revenue reporting.

## Assumptions Made
- Discount values should be between 0 and 50% inclusive.
- Missing discounts can be treated as 0 only when quantity, unit price, sales, cost, and profit are present.
- Returned orders should not contribute to completed-sales summaries.
- Completed sales summaries include only completed, paid, non-invalid, non-duplicate records.
- Sales and profit validation uses a small rounding tolerance of 0.05.

## Known Limitations
- Conflicting duplicate order IDs were flagged for review instead of manually merged.
- Invalid records were not corrected unless a rule gave a reliable correction.
- Pivot summaries are summary worksheets, not native Excel PivotTable cache objects.
- The cleaning process uses deterministic text standardization and does not infer unknown values beyond the provided rules.

## Screenshots
- [Raw data preview](screenshots/raw_data_preview.png)
- [Cleaned data preview](screenshots/cleaned_data_preview.png)
- [Pivot summary 1: sales and profit by region](screenshots/pivot_summary_1.png)
- [Pivot summary 2: monthly sales trend](screenshots/pivot_summary_2.png)

## Repository Checklist
- Raw dataset preserved: `data/raw_orders.xlsx`
- Cleaned dataset created: `data/cleaned_orders.xlsx`
- Data quality report created: `outputs/data_quality_report.xlsx`
- Pivot summary report created: `outputs/pivot_summary.xlsx`
- Cleaning log created: `outputs/cleaning_log.md`
- Required screenshots created inside `screenshots/`
