# mx-sales

Cleaning, transformation and reconciliation for MercadoLibre MX seller reports.

ML splits a single sale's economics across several exports and does not publish a
figure you can tie out directly. This project cleans each export into a typed table and
then applies the accountant's rules to produce one auditable row per sale.

## Layout

```
mx_sales/
  config.py          paths and settings (all disk locations resolve from here)
  registry.py        the catalogue of known report types
  pipeline.py        discover raw files -> clean -> Parquet
  warehouse.py       DuckDB views over the Parquet output
  readers/excel.py   header detection, two-row grouped headers
  clean/             text / numbers / dates / booleans primitives
  reports/           one module per report type (loader + transformer + spec)
  accounting/        the reconciliation rules and their runner
data/
  processed/<report>/    Parquet output, partitioned
  mx_sales.duckdb        views over the above
tests/
```

Raw downloads stay where the download pipeline puts them and are never modified.

## Usage

Work one file at a time — a snapshot tree holds the same report many times over:

```bash
python -m mx_sales discover --unmatched          # what is here, and what has no spec yet
python -m mx_sales build --file <export name>    # clean one download to Parquet
python -m mx_sales validate <export name> --excel out.xlsx
python -m mx_sales query "SELECT * FROM ventas_mx LIMIT 5"
```

Drop `--file` to process everything discovered under the raw directory.

## What the cleaning handles

Found in the real exports, all covered by tests:

| Problem | Where |
|---|---|
| Blanks written as a single space, so every column looks fully populated | every report |
| Two-row header with repeating names (`Unidades` under three groups) | Ventas MX |
| Header row at a different offset in each report (row 3, 6, 8 or 10) | all XLSX |
| Spanish long-form dates, some with no year (`11 de agosto`) | Ventas MX |
| 16-digit order ids arriving as `float64` | Facturación |
| Composite fields (`RFC: …`, `S01 Sin efectos fiscales.`, `Color : Negro \| …`) | Ventas MX |
| One order spanning several rows (packages, exchanges) | Ventas MX |
| The same report re-downloaded across snapshots | all |

Year-less dates are resolved against the sale date on the same row, choosing the
nearest year — which handles both a December sale delivered in January and an exchange
whose shipment predates the replacement sale.

## The validation rules

`mx_sales/accounting/ventas_mx.py`. Rules are added one at a time, as each is confirmed.

| # | Rule | Output |
|---|---|---|
| 1 | `H..O` must sum to `P` | `components_sum_mxn`, `total_diff_mxn`, `total_matches` |
| 2 | Package money shared across the products inside it | `allocated_from_package`, `package_split` |
| 3 | A returned product is left out of that split | `is_returned_item`, `package_note` |
| 4 | The package row is removed once it has been split | — |
| 5 | Tax withheld: `H / 1.16 * 0.105` | `tax_withholding_mxn` |
| 6 | Platform fee: `\|I\| - tax` | `platform_fee_mxn` |
| 7 | Accounting period from the sale date | `sale_period` |
| 8 | Buyer-identity blocks dropped | — |
| 9 | Unnecessary columns dropped | — |
| 10 | Fully-reversed sales excluded from revenue | `counts_toward_revenue`, `revenue_exclusion_reason` |
| 11 | Per-SKU sales and billed fees | the `sku` sheet |

Rules 1–10 run **2, 3, 4, 1, 10, 5, 6, 7, 8, then 9**: the packages are split and their rows removed first, so by
the time the H..O check runs every remaining row carries its own money.

**Rule 2.** A multi-product sale is written as a `Paquete de N productos` row holding all
the money, followed by N product rows holding the SKU and unit price but nothing else.
Each product's share is `(units * unit_price) / sum(units * unit_price)` — the `W / H`
ratio, expressed against the summed basis so it stays right if a line ever carries a
quantity above one. Every column H..P is split on that same ratio, and rounding residue
goes to the largest share so the products always sum back to the package to the cent.

**Rule 4.** The package row is dropped once its money sits on the products, since keeping
it would count everything twice. Only rows marked `package_split` go: a package rule 3
left intact still holds the only copy of its money, so it stays.

**Rule 3.** A refund in column O marks a return. The refund is folded into revenue —
the amount shared out is `H + O` — and only the products that did *not* come back divide
it, so a return cannot distort the others. Returned products keep no money and are
flagged `is_returned_item`. A package is left intact, with a `package_note`, when every
product came back (`fully_returned`, so the payout stays in the accounts), when a refund
is booked but no product row says which item it belongs to
(`refund_without_identifiable_item`), or when there are no unit prices to divide on.

Which product came back is read from `order_status`, because product rows carry no money
of their own. Statuses where the seller *kept* the money — "Te dimos el dinero",
"Descartamos el producto" — book no refund and are not treated as returns.

> **Caveat.** ML sets O to the reversal of the whole net (revenue less fees and
> shipping), not to the returned product's list price: on a standalone return
> `O = -(H+I+K)` exactly. So `H + O` is revenue-net-of-reversal, not the surviving
> products' combined price.

**Rule 5.** H is quoted gross of the 16% IVA, so `H / 1.16` is the amount net of it and
10.5% of that is the withholding. It runs after the packages
are split, so a product that came out of a package is taxed on its own share of the
revenue rather than on the package's.

**Rule 6.** Column I is *Cargo por venta e impuestos* — commission and tax in one figure
— so taking rule 5's tax back out leaves the commission. Reported positive, like the tax.

> **Sign.** Revenue is positive and **every cost is negative** — platform fee, withheld
> tax, shipping and refunds alike. That matches the accountant's workbook and makes the
> money columns add up: `revenue + fee + tax + shipping_revenue + shipping_cost +
> refunds = payout`, exact on every period tested. Rule 6's arithmetic is done on
> magnitudes and then negated; subtracting the signed values would give −184.33 where ML
> bills 69.67.

Two independent checks that these are right: `platform_fee_mxn + tax_withholding_mxn`
rebuilds `|I|` on every row, and `platform_fee_mxn` reproduces the "Cargo por venta" that
`Reporte_Facturacion` bills, to the cent, on 98 of the 104 orders in both reports. The
implied commission rate lands on exactly 11.00% or 14.50% for 230 of 264 rows — ML's
published tiers.

**Rule 7.** `sale_period` is `YYYYMM` from the sale date — the month the sale happened,
not the month ML billed for it. In this data those agree 98.9% of the time, but not
always, so a fee-based period would have to come from `Reporte_Facturacion` instead.

Strictly the column is derived — `sold_at` is a real timestamp, so any tool that can
group by month can produce it. It is stored because the people reading the workbook pivot
in Excel, where a text key beats a date expression, and a stored key cannot drift between
one person's query and another's. `rules.monthly_summary()` is what the column exists to
support, and lands in the workbook as the `mensual` sheet.

**Rule 8.** The two buyer blocks come out: *Facturación al comprador* (yellow in the
sheet, `FFFFD966`, columns Y..AF) and *Compradores* (green, `FFB1DA9F`, AG..AN) — 16
source columns, 18 in the cleaned schema once the composite fields are split. The set is
derived from the sheet's own group headers via
`reports.ventas_mx.IDENTITY_COLUMNS`, so it follows the cleaner's mapping rather than
being restated by hand. It removes the report's only personal data (names, RFCs,
addresses) — and its only geography, `buyer_state` and `buyer_municipality`, so skip this
step if sales by region is ever wanted.

**Rule 9.** Columns carrying no accounting signal, listed in `UNNECESSARY_COLUMNS`:

| source col | column | why it goes |
|---|---|---|
| E `Paquete de varios productos` | `is_multi_product_package` | Not the package marker it looks like — `Sí` on 147 standalone rows in BOCINA and 7,466 in EWTTO, far more than the actual package products, and nothing else in the report explains which rows get it. The package logic uses `row_role` instead. |
| F `Pertenece a un kit` | `belongs_to_kit` | `False` on every product row in both stores — no ML kits sold, so the column is constant. |
| V `Variante` (parsed) | `listing_variant_attributes` | Restates `listing_variant` as JSON — `Color : Negro \| Voltaje : 127V` vs `{"Color": "Negro", "Voltaje": "127V"}`. Same content, so the raw text is kept and the copy goes. |
| — | `sale_year` | Already the first four characters of `sale_period`. |
| — | `store`, `seller_id` | Pipeline lineage. One export belongs to one account, so they are constant within a run. |

> **On `store` / `seller_id`.** Deduplication and Parquet partitioning key on `seller_id`
> and read the *cleaned* table, which keeps both, so those are unaffected. But order ids
> are unique per account and not across them — once these are dropped, validated outputs
> from two stores can no longer be safely concatenated. `source_file` still identifies
> the origin of every row.

Add to that list as more are agreed. Like rule 8, it only affects the accounting output:
the cleaned `ventas_mx` table keeps every source column, so a dropped column can be
reinstated without re-reading the workbook. That also keeps a *blank* E available there —
it marks package parents and exchange rows, a useful cross-check on the row-structure
detection.

**Rule 10.** A cancelled or fully-returned sale keeps its `product_revenue_mxn` while its
payout falls to zero, so summing revenue over every row counts sales that earned nothing.
Rows with `total_mxn == 0` are flagged `counts_toward_revenue = False` and reason-coded
from their status (`cancelled` / `returned` / `mediation`, or `zero_payout` when the
status says nothing useful). A *partial* refund that still paid out is kept — the revenue
was real — as is a null payout and a loss-making sale.

Nothing is deleted: the rows go to the `excluidos` tab, and `mensual` reports
`excluded_orders` / `excluded_revenue_mxn` beside the counted figures so that
`revenue_mxn + excluded_revenue_mxn = gross_revenue_mxn`, the raw export total.

This is the accountant's `零金额` filter. With it, BOCINA_TA02 August reproduces every
figure on their `订单统计` tab exactly — 123 units, 116,109.73 revenue, 154.56 shipping
income, −12,964.27 fee, −10,510.14 tax, −10,905.06 shipping, −113.75 refunds,
81,771.07 payout.

**Account-level costs on `mensual`.** `monthly_summary(frame, charges)` adds the two
costs the accountant carries as `广告费` and `平台仓租`:

| column | from |
|---|---|
| `advertising_mxn` | Facturación `advertising` charges (Product Ads + Display Ads) |
| `storage_mxn` | `storage` + `pickup` — Full storage and colecta, grouped as they group them |
| `net_after_account_costs_mxn` | `payout + advertising + storage` |

Credit-note voids are dropped here too: advertising bills 6,394.60 in August, 117.26 of it
voided, leaving the **6,277.34** they publish. Storage + pickup gives **5,160.66**, and
the net comes to **70,333.07** — all three match their sheet exactly.

> **These sit outside the payout identity on purpose.** ML bills them to the account, not
> to a sale, so they reduce the month's profit without touching any order. They also carry
> no sale date — none of the 122 account-level charges has one — so they key on the month
> ML billed them, and on a straddling month they will not line up with the sales columns.

**Rule 11.** `sku_summary()` rolls the finished rows up to one row per `sale_period` +
`sku` — the shape of the accountant's `销售` tab — and joins in what ML actually billed.
It is a summary over the completed rows, not a step in the chain above.

Where each column comes from:

| column | source |
|---|---|
| `units`, `revenue_mxn`, `shipping_revenue_mxn` | the sales export, rule 10's filter applied |
| `tax_withheld_mxn` | `-(revenue / 1.16 * 0.105)` on the **group** total |
| `fee_commission_mxn`, `fee_shipping_mxn`, their `_reversal` pairs, `fee_return_mxn` | `Reporte_Facturacion`, pivoted by SKU |
| `fee_total_mxn` | sum of those fee columns |
| `settlement_amount_mxn` | `revenue + tax + shipping_revenue + fee_total` |
| `settlement_unit_price_mxn` | `payout / units` |
| `k3_material` | left blank — an accounting code, not in any ML export |

Three things make this work:

**A charge reaches a SKU through the listing.** Facturación publishes no SKU but does
publish `Número de publicación`, and the sales export maps every listing to a SKU
(`listing_sku_map()`). All 314 charge rows with a listing resolve on this data. Where
colour variants share a listing the busiest SKU wins, which reproduces the accountant's
own product master.

**Account-level charges are left out.** Storage, advertising, pickup and page maintenance
name no listing — 122 rows here — so they belong to no product and are not spread across
SKUs on a guess. They stay visible on the `costos` sheet.

**Credit-note voids are dropped.** A charge marked `Anulado en nota de crédito` was
refunded, so counting it would overstate fees. Charges voided *on the invoice* are kept —
ML has already netted those, and the accountant does not deduct them either. This is
their `调减费用` flag, and it is what takes the fee match from 4/9 to 9/9.

`tax_withheld_mxn` is rounded **once on the group** rather than summed from the per-row
column: rule 5 rounds each row to the cent and adding hundreds of those drifts, by up to
0.16 per SKU here.

Verified against the accountant's `销售` tab for BOCINA_TA02 August — all 8 SKUs match
exactly on revenue, commission, shipping and `费用合计`, and `结算金额` matches on 7 of 8
with one cent on the smallest.

> **Fees need the right month downloaded.** They join on the charge's own sale date, so a
> period with no matching `Reporte_Facturacion` in the snapshot folder shows blank fee
> columns rather than zeros.

`total_matches` is `True` where the row reconciles, `False` where it does not, and null
where there is nothing to compare — a null must not be reported as a reconciling `0.00`.
`rules.packages()` returns the rows rule 2 removed, for audit.

## Reference: how the accountant builds 汇总 and 销售

`data/reports/2026.7-8月MX-TA02平台销售及费用统计.xlsx` is the report we are approximating.
Its two headline tabs are the same table: **`销售`** keyed `年月 + SKU`, and **`汇总`** the
same filtered to one period with a `序号` column and a totals block on top. Everything
below was reverse-engineered from their file and checked against it.

They feed those two tabs from two upstream tabs of their own, each a pasted ML export
with extra columns added on the right:

| their tab | is | columns they add |
|---|---|---|
| `订单明细` | the Ventas MX export | `年月`, `店铺`, `组合`, `标记`, `零金额`, `销售金额/1.16*0.105` |
| `报告明细` | `Reporte_Facturacion` | `年月`, `店铺`, `费用项目`, `标志`, `SKU`, `引用R`, `引用AB`, `重次`, `税金`, `调减费用`, `抵扣年月` |

### Where each 销售 column comes from

| column | source |
|---|---|
| `销量`, `销售金额`, `客人支付运费` | `订单明细`, grouped by SKU, zero-payout rows filtered out (`零金额`) |
| `税金` | `-(销售金额 / 1.16 * 0.105)`, computed on the **SKU total** |
| `销售费`, `取消销售费`, `Mercado Libre 运费`, `取消 Mercado Libre 运费`, `退款费用`, `广告费用`, `仓储服务费`, `页面维护费`, `MKP销售代理佣金` | `报告明细` pivoted by SKU × `费用项目`, sign flipped |
| `费用合计` | sum of those nine fee columns |
| `结算金额` | `销售金额 + 税金 + 客人支付运费 + 费用合计` |
| `结算单价` | `结算金额 / 销量` |
| `K3物料` | lookup from their `产品ID` tab |

Worked example, `BH00335GY` 202607: `2594.00 − 234.80 + 0.00 − 855.80 = 1503.40` is their
`结算金额`, and `1503.40 / 6 = 250.57` is their `结算单价`.

### 费用项目 — their fee categories

`费用项目` is a rename of ML's `Detalle`, and maps onto our `fee_category`:

| 费用项目 | ML `Detalle` | ours |
|---|---|---|
| 销售费 | Cargo por venta | `commission` |
| 取消销售费 | Anulación del cargo por venta | `commission_reversal` |
| Mercado Libre 运费 | Cargo por envíos de Mercado Libre | `shipping` |
| 取消 Mercado Libre 运费 | Anulación del cargo por envíos de Mercado Libre | `shipping_reversal` |
| 广告费用 | Product Ads + Display Ads, and their reversals | `advertising` |
| 仓储服务费 | Cargo por servicio de almacenamiento Full **and** de colecta Full | `storage` + `pickup` |
| 页面维护费 | Cargo por mantenimiento de Mi página, and its reversal | `maintenance` |

Note they fold `colecta` (pickup) into storage, and ad reversals into advertising, where
we keep those separate.

### The three mechanisms that are easy to miss

**1. A charge gets its SKU from the listing.** Facturación names no SKU. They look
`Número de publicación` (`MLM…`) up in their `产品ID` tab — verified **304 of 304 exact**.
The 108 charge rows with no SKU are precisely the account-level ones (storage 58, ads 48,
maintenance 2), which is why those columns read 0 on every SKU row.

*We do the same thing without their product master*: the sales export already maps every
listing to a SKU, so `listing_sku_map()` derives the lookup from data we download.

**2. Charges voided by a credit note are deducted.** Their `调减费用 = Y` flag drops 16
rows. Filtering on it reproduces the tab **9 of 9** on both fee columns; without it only
**4 of 9**. The flag is mechanical — `收费状态 = "Anulado en nota de crédito"` **and** the
row has a SKU. Charges voided *on the invoice* (`Anulado en factura`, 28 rows) are **not**
deducted: ML has already netted those.

**3. Colour variants collapse into the listing's canonical SKU.** Their July shows 6 units
of `BH00335GY`; the raw export has 5 `BH00335GY` + 1 `BH00335PK`, both on listing
`MLM3137580303`, which `产品ID` maps to `BH00335GY`. Our rule 11 reproduces this by giving
a shared listing to whichever SKU sold most.

### What our `sku` sheet matches

Verified for BOCINA_TA02, 202608 — all 8 SKUs exact on `销售金额`, `销售费`,
`Mercado Libre 运费` and `费用合计`; `结算金额` exact on 7 of 8, one cent on the smallest;
totals 123 units and 116,109.73 revenue.

`K3物料` is left blank: it is an internal ERP material code, not published in any ML
export.

## Adding a report type

Write a module under `reports/` exposing `load`, `transform` and a `SPEC`, then add the
spec to `registry.py`. Discovery, Parquet output and the DuckDB view follow from it.
Still unspecified: `settlement_v2`, `reserve-release`, `account_statement`,
`collection`, `withdraw` (CSV, `;`-delimited), and `Notas_Credito`, `Cargos_Full`,
`Pagos_Facturas`, `Costos_por_servicio_almacenamiento`, `stock_general_full`,
`conciliation`, `Returns` (XLSX).

## Tests

```bash
python -m pytest tests -q
```

Tests that need a real export skip themselves when one is not present.
