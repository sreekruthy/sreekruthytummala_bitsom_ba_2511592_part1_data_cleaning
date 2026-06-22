# Cleaning Log

## 1. Source and scope

- Source: `data/raw_orders.xlsx`
- Raw records reviewed: **932**
- Exact duplicate copies removed: **20**
- Final retained records: **912**
- The raw workbook was preserved unchanged. All transformations were performed in `data/cleaned_orders.xlsx`.

The basic reconciliation is:

```text
932 raw records − 20 later exact copies = 912 retained records
```
---

## 2. Issues found 

| Issue | Count | How the count was obtained |
|---|---:|---|
| Exact duplicate row groups | 20 | Grouped raw rows using all 21 original columns and counted signatures occurring more than once. |
| Exact duplicate copies removed | 20 | For every identical row signature, kept the first occurrence and counted every later occurrence. |
| Distinct duplicate order IDs | 31 | Counted distinct `order_id` values appearing more than once in the raw data. |
| Conflicting duplicate order IDs | 12 | Among repeated IDs, counted IDs having more than one distinct full-row version. |
| Conflicting records retained for review | 24 | Summed the distinct retained versions belonging to the 12 conflicting IDs. |
| Missing region | 25 | `COUNTIF(clean_region,"Unknown")` after exact-copy removal. |
| Missing ship mode | 21 | `COUNTIF(clean_ship_mode,"Unknown")` after exact-copy removal. |
| Missing discount imputed to 0 | 18 | Counted records where raw discount was blank and `cleaned_discount = 0`, after confirming the other sales fields were numeric. |
| Negative discount | 15 | `COUNTIF(cleaned_discount,"<0")`. |
| Discount above 50% | 15 | `COUNTIF(cleaned_discount,">0.5")` after converting text such as `70%` to `0.70`. |
| Missing order date | 0 | Counted blank/unrecognized values in `clean_order_date`. |
| Missing ship date | 0 | Counted blank/unrecognized values in `clean_ship_date`. |
| Invalid date text | 0 | Counted date values that could not be parsed under the documented format rules. |
| Ship date before order date | 21 | Counted rows where `clean_ship_date < clean_order_date`. |
| Cancelled orders | 145 | `COUNTIF(clean_order_status,"Cancelled")`. |
| Returned orders | 163 | `COUNTIF(clean_order_status,"Returned")`. |
| Failed payments | 69 | `COUNTIF(clean_payment_status,"Failed")`. |
| Refunded payments | 71 | `COUNTIF(clean_payment_status,"Refunded")`. |
| Pending payments | 86 | `COUNTIF(clean_payment_status,"Pending")`. |
| Sales mismatches | 64 | Counted rows where `ABS(source sales − calculated_sales) > 0.05`. |
| Profit mismatches | 64 | Counted rows where `ABS(source profit − calculated_profit) > 0.05`. |
| Clean records | 505 | `COUNTIF(data_quality_flag,"Clean")`. |
| Warning records | 305 | `COUNTIF(data_quality_flag,"Warning")`. |
| Invalid records | 102 | `COUNTIF(data_quality_flag,"Invalid")`. |
| Eligible completed-sales records | 537 | `COUNTIF(completed_sales_include,"Yes")`. |

Issue counts overlap. For example, one cancelled order can also have a refunded payment. Therefore, issue counts must not be added together to calculate the total number of flagged records. The mutually exclusive `data_quality_flag` gives the correct reconciliation:

```text
505 Clean + 305 Warning + 102 Invalid = 912 retained records
```

---

## 3. Cleaning actions performed

### Text fields

The following fields were standardized:

`customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, and `order_status`.

The cleaning pattern removed leading/trailing spaces, repeated spaces, non-breaking spaces, and control characters, then standardized case. A representative Excel formula is:

```text
=PROPER(TRIM(CLEAN(SUBSTITUTE(source_cell,CHAR(160)," "))))
```

Missing `region` and `ship_mode` values were filled with `Unknown` and retained as Warning records. They were not inferred from other rows.

### Date conversion and validation

The source contained four date styles. Each was handled explicitly:

- Slash dates: `MM/DD/YYYY`
- Hyphen dates: `DD-MM-YYYY`
- ISO dates: `YYYY-MM-DD`
- Text dates: `DD Mon YYYY`

The conventions are supported by unambiguous source examples such as `08/31/2024` and `28-11-2024`.

```text
shipping_delay_days = clean_ship_date − clean_order_date
```

A negative result means the ship date precedes the order date and the record is Invalid. This identified **21** reversed shipping records.

### Duplicate handling

A full-row signature was built from all original business fields.

- If the signature repeated, only the first identical row was retained.
- If the same `order_id` had different signatures, all distinct versions were retained and marked `Review conflicting duplicate`.
- Conflicting records were excluded from completed-sales reporting pending business review.

This removed **20** exact copies without silently deleting any of the **24** conflicting retained records.

### Discount standardization

Percentage text was converted before validation. For example:

```text
70% to 0.70
85% to 0.85
```

A blank discount was converted to zero only when quantity, unit price, source sales, cost, and source profit were numeric. The valid range was assumed to be **0%–50%, inclusive**.

### Calculated columns

```text
calculated_sales = quantity × unit_price × (1 − cleaned_discount)
calculated_profit = calculated_sales − cost
profit_margin = calculated_profit ÷ calculated_sales
shipping_delay_days = clean_ship_date − clean_order_date
order_month = month extracted from clean_order_date
order_year = year extracted from clean_order_date
```

The calculation columns are formula-backed through all 912 retained records.

### Calculation mismatches

```text
sales_mismatch = ABS(source sales − calculated_sales) > 0.05
profit_mismatch = ABS(source profit − calculated_profit) > 0.05
```

The 0.05 tolerance prevents small rounding differences from being incorrectly flagged. Invalid-discount records can also be calculation mismatches, so mismatch and discount counts overlap.

### Data-quality classification

- **Invalid (102):** invalid discount, reversed date sequence, sales mismatch, or profit mismatch.
- **Warning (305):** a non-invalid issue such as missing-value replacement, non-completed/payment status, or conflicting duplicate review.
- **Clean (505):** no warning or invalid condition.

### Completed-sales inclusion

A record contributes to completed-sales reports only when all conditions are true:

```text
order_status = "Completed"
payment_status = "Paid"
data_quality_flag <> "Invalid"
exact_duplicate_action = "Keep"
duplicate_review_flag = blank
```

This produces **537** eligible completed-sales records.

Looking at your cleaning log, I can see sections 3 (Cleaning Actions) and 5 (Assumptions) already contain embedded business rule info. Here's the dedicated text to add as new sections:

---

## 4. Business rules applied

| Rule Area | Business Rule | Action Taken |
|---|---|---|
| Missing region | If `region` is blank, fill with `Unknown` | Filled 25 records; flagged as Warning |
| Missing ship mode | If `ship_mode` is blank, fill with `Unknown` | Filled 21 records; flagged as Warning |
| Missing discount | Treat blank discount as 0 only when quantity, unit price, source sales, cost, and source profit are all numeric | Applied to 18 records; records with mismatching sales/profit after imputation were flagged Invalid |
| Negative discount | Any `cleaned_discount < 0` is invalid | Flagged 15 records as Invalid |
| Discount above 50% | Any `cleaned_discount > 0.5` is invalid (percentage text such as `70%` converted to `0.70` before check) | Flagged 15 records as Invalid |
| Ship date before order date | If `clean_ship_date < clean_order_date`, the record has an invalid date sequence | Flagged 21 records as Invalid |
| Cancelled orders | Cancelled order status does not contribute to completed-sales summaries | 145 records excluded from completed-sales reports; retained in audit data |
| Returned orders | Returned order status does not contribute to completed-sales summaries | 163 records excluded; retained in audit data |
| Failed payments | Failed payment status does not contribute to completed-sales summaries | 69 records excluded; retained in audit data |
| Refunded payments | Refunded payment status does not contribute to completed-sales summaries but is separately summarized by region | 71 records excluded from completed sales; summarized in pivot report |
| Pending payments | Pending payment status does not contribute to completed-sales summaries | 86 records excluded; retained in audit data |
| Exact duplicate rows | Only the first occurrence of a fully identical row is retained; later copies are removed | 20 exact copies removed |
| Conflicting duplicate order IDs | All distinct versions of a repeated order ID with differing field values are retained and flagged for review | 24 conflicting records flagged; excluded from completed-sales reporting pending business review |
| Calculation mismatch tolerance | A record is flagged for sales or profit mismatch only when the absolute difference exceeds 0.05, to prevent rounding differences from being incorrectly flagged | Applied to all 912 retained records |
| Completed-sales inclusion | A record contributes to completed-sales summaries only when order status is Completed, payment status is Paid, data quality flag is not Invalid, exact duplicate action is Keep, and no conflicting duplicate review flag is present | 537 records qualify |

---

## 5. Assumptions

- Valid discounts range from 0% through 50%, inclusive.
- Missing discounts are set to 0 only when all other required sales fields are valid.
- Calculation mismatch tolerance is 0.05.
- Only Completed and Paid, non-invalid, non-conflicting records contribute to completed-sales summaries.
- Returned and Pending records are excluded because the final inclusion rule requires Completed and Paid.

---

## 6. Records removed

| Action | Count | Reason |
|---|---|---|
| Exact duplicate copies removed | 20 | Later occurrences of a row whose full 21-field signature appeared more than once in the raw data. Only the first occurrence was kept. |
| **Total records removed** | **20** | |

The raw workbook was not modified. All removals are reflected only in `data/cleaned_orders.xlsx`.

Reconciliation:
```
932 raw records − 20 exact duplicate copies = 912 retained records
```

---

## 7. Records flagged

| Flag type | Count | Reason |
|---|---|---|
| Invalid - negative discount | 15 | `cleaned_discount < 0` |
| Invalid - discount above 50% | 15 | `cleaned_discount > 0.5` |
| Invalid - ship date before order date | 21 | `clean_ship_date < clean_order_date` |
| Invalid - sales mismatch | 64 | `ABS(source sales − calculated_sales) > 0.05` |
| Invalid - profit mismatch | 64 | `ABS(source profit − calculated_profit) > 0.05` |
| Warning - missing region filled | 25 | `clean_region = Unknown` |
| Warning - missing ship mode filled | 21 | `clean_ship_mode = Unknown` |
| Warning - missing discount imputed to 0 | 18 | Raw discount blank; other sales fields valid |
| Warning - cancelled order | 145 | `clean_order_status = Cancelled` |
| Warning - returned order | 163 | `clean_order_status = Returned` |
| Warning - failed payment | 69 | `clean_payment_status = Failed` |
| Warning - refunded payment | 71 | `clean_payment_status = Refunded` |
| Warning - pending payment | 86 | `clean_payment_status = Pending` |
| Warning - conflicting duplicate flagged for review | 24 | Repeated order ID with more than one distinct full-row version |

Issue categories overlap - one record can trigger multiple flags. The mutually exclusive `data_quality_flag` gives the correct reconciliation:

```
505 Clean + 305 Warning + 102 Invalid = 912 retained records
```

---

## 8. Limitations

- No external master data was available to infer missing region or ship mode values.
- The dataset does not identify which conflicting duplicate version is authoritative; all distinct versions are retained for review.
- Issue categories overlap, so individual issue counts cannot be summed to obtain total flagged records.
- The summary workbooks use transparent formula-backed Excel tables rather than native PivotTable cache objects.

--- 


## 9. Report methodology

`data_quality_report.xlsx` contains the final 912-row audit data, duplicate-ID detail, issue details, all required summaries, and a methodology sheet. Summary count cells use bounded `COUNTIF`, `COUNTIFS`, `SUM`, and `SUMIF` formulas.

`pivot_summary.xlsx` contains a 912-row reporting data sheet and six required formula-backed pivot summaries. Completed-sales reports use the common filter:

```text
completed_sales_include = "Yes"
```
