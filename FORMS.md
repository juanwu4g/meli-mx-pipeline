# Downloaded Forms — Catalogue

Every form this pipeline pulls from the MercadoLibre seller panel: where it comes
from, how it is obtained, and what is inside it.

**This document grows.** As routes are added, append a section under the matching
route and add a row to the index. Keep the "verified" date honest — the panel's
DOM changes, and a stale note is worse than no note.

- Store used for verification: `BOCINA_SM` (MercadoLibre México, seller `3426670806`)
- Last verified: **2026-09-02**, full pipeline run
- Entry point: `python run_downloads.py [--only billing|ventas|stock|mercadopago|all]`
- Output: `downloads/<STORE>/<timestamp>/`

---

## Index

| # | Form | Route | Files per run | Period control |
|---|---|---|---|---|
| 1 | Facturación de Mercado Libre | billing | 0–1 per month | month card |
| 2 | Cargos de Full | billing | 0–1 per month | month card |
| 3 | Pagos de facturas | billing | 0–1 per month | month card |
| 4 | Notas de crédito | billing | 0–1 per month | month card |
| 5 | Excel de ventas | ventas | 1 | period dropdown (6 months) |
| 6 | Costos por servicio de almacenamiento | stock | **2** | period dropdown |
| 7 | Reporte general de stock | stock | 1 | none — snapshot |
| 8 | Reporte consolidado de movimientos | stock | 1 | date-range calendar |
| 9 | Reporte de devoluciones a la bodega | stock | 1 | none |
| 10 | Todas las transacciones (MercadoPago) | mercadopago | 1 (.csv) | 60-day range, or the month's automatic |
| 11 | Liberaciones (MercadoPago) | mercadopago | 1 (.csv) | same |
| 12 | Cobros (MercadoPago) | mercadopago | 1 (.csv) | 30-day range |
| 13 | Retiros (MercadoPago) | mercadopago | 1 (.csv) | 30-day range + estado filter |
| 14 | Poscobro (MercadoPago) | mercadopago | 1 (.csv) | 60-day range + 2-step wizard |
| 15 | Estados de saldos y movimientos (MercadoPago) | mercadopago | 1 (.csv) | newest whole month |

A full run yields **18 files** on a good day: up to 8 billing + 1 ventas +
5 stock + 6 MercadoPago. The billing count varies — see route 1.

Last verified end-to-end **2026-09-02**: 6 billing (September had only 2 report
types), 1 ventas, 5 stock, 6 MercadoPago.

## Run phases — why slow reports do not cost time

Two forms are generated server-side and are not ready when requested, so the run
is split into three phases:

```
PHASE 1  request   ventas Excel  (~30 s to generate)
                   MercadoPago settlement  (~6.5 min to generate)
                        ↓
PHASE 2  main      billing → stock          (downloads immediately, ~6-8 min)
                        ↓
PHASE 3  collect   return to each page, poll until ready, download
```

### Phase 3 is ONE shared wait, not a budget split

The reports generate **concurrently** on MercadoPago's side, so time spent
waiting on one advances all of them. Phase 3 therefore cycles over everything
outstanding, checking each cheaply, until all are collected or
`--collect-timeout` (default **1800s**) expires:

```
>>> collect cycle 1 - 1800s left, 7 outstanding: ventas, mercadopago/settlement, ...
```

An earlier version divided the budget between pending items instead. With seven
reports and 600s that gave each 100s, and nothing taking 2.6–6.5 minutes could
finish in its slice — most timed out no matter how long the total was. Raising
the timeout was not the fix; sharing it was.

Measured 2026-09-02: **all 7 pending reports collected in cycle 1**, because
phase 2 had already absorbed the generation time.

Anything still not ready is listed under `[pending]` and the run continues —
those reports stay on their pages and can be collected on a later run, so a slow
generation never fails the run.

---

## Route 1 — Facturación (billing)

**Page:** <https://vendedores.mercadolibre.com.mx/billing/resume> (title: *Facturación*)
**Menu path:** Tarifas y pagos → month card → `Ir al detalle` → `Reportes` tab
**Code:** `meli_forms.download_billing_reports()`

### How it is obtained

1. Open `/billing/resume`. Each month is a card showing its state (`EN CURSO` /
   `VENCIDO` / closed).
2. Click that card's `Ir al detalle` → `/billing/detail/<YYYYMMDD>`, keyed on the
   billing close date.
3. Click the `Reportes` tab (`button.billing-detail_tab-reports`).
4. Tick **Seleccionar todos los reportes**, then `Descargar`.

**The count is not fixed.** MercadoLibre only offers report types that have data
for that period, so select-all yields a different number each month:

| Period | Types offered | Files |
|---|---|---|
| Agosto 2026 (EN CURSO) | BILL_ML, FULL, PAYMENT, NC_ML | 4 |
| Julio 2026 | BILL_ML, FULL, PAYMENT, NC_ML | 4 |
| Junio 2026 | BILL_ML, PAYMENT | 2 |

### 1. `Reporte_Facturacion_MercadoLibre_<Mmm><YYYY>.xlsx`

The master charge ledger — every fee MercadoLibre billed in the period.

- Sheet `REPORT`, **header row 8**, 33 columns
- Key columns: `Detalle` (charge type), `Valor del cargo`, `Fecha del cargo`,
  `N° de factura fiscal`, `Estado del cargo`
- Charge types seen: `Cargo por venta`, `Cargo por envíos de Mercado Libre`,
  `Cargo por campaña de publicidad de Product Ads` / `Display Ads`,
  `Cargo por servicio de almacenamiento Full`, `Cargo por servicio de colecta Full`,
  `Cargo por retiro de stock Full`, `Cargo por devolución`,
  `Cargo por mantenimiento de Mi página`, `Anulación del cargo por …`
- **Amounts are positive** — a 406.00 ad charge means 406 pesos billed *to you*.
  `report_build.load_billing()` signs the amounts (`amount`) for P&L use.

### 2. `Reporte_Cargos_Full_<Mmm><YYYY>.xlsx`

Dimensional breakdown of the Full warehouse charges.

- 3 sheets, **header row 6**: `Cargo por almacenamiento`,
  `Cargos por servicio de colecta`, `Cargo por retiro de stock`
- Key columns: `Monto del cargo`, `Unidades almacenadas`, `Tamaño de la unidad`,
  `Volumen total (m3)`
- ⚠️ **These rows are already inside `Reporte_Facturacion`.** Verified identical
  for August 2026 — storage 514.74, pickup 2,140.24, withdrawal 19,129.46 appear
  in both files. **Never sum both**, or every Full fee is counted twice. Use this
  file for breakdowns only.

### 3. `Reporte_Pagos_Facturas_<Mmm><YYYY>.xlsx`

Payments made against invoices.

- 2 sheets, **header row 10**: `Pagos y notas de crédito`, `Detalle de Pagos del mes`
- Key columns: `Número de pago`, `Tipo de pago`, `Medio de pago`,
  `Fecha de pago / Emisión`, `Estado`, `Importe total`
- Not yet used by either analysis method. Joining it to `Facturacion` on invoice
  number would show billed vs settled — relevant, since July 2026 shows
  64,447.19 invoiced against only 3,420.54 collected.

### 4. `Reporte_Notas_Credito_MercadoLibre_<Mmm><YYYY>.xlsx`

Credit notes — reversals of earlier charges.

- Sheet `REPORT`, **header row 8**
- Same shape as `Facturacion`; `Valor del cargo` is negative,
  `Cargo que bonifica` points at the reversed charge

---

## Route 2 — Ventas (sales)

**Page:** <https://vendedores.mercadolibre.com.mx/ventas/omni/listado>
**Menu path:** Ventas
**Code:** `meli_forms.request_sales_excel()` / `collect_sales_excel()`
(`download_sales_excel()` does both, for running this route alone)

### 5. `<YYYYMMDD>_Ventas_MX_Mercado_Libre_y_Mercado_Shops_<date>_<time>hs_<seller>.xlsx`

One row per order line — the backbone of all sales analysis.

- Sheet `Ventas MX`, **header row 6**, 63 columns
- Key columns: `# de venta`, `Fecha de venta`, `Estado`, `Unidades`,
  `Ingresos por productos (MXN)`, `Cargo por venta e impuestos (MXN)`,
  `Ingresos por envío (MXN)`, `Costos de envío (MXN)`,
  `Anulaciones y reembolsos (MXN)`, `Total (MXN)`, `SKU`,
  `Título de la publicación`, `Estado.1` (buyer state), `Forma de entrega`,
  `Venta por publicidad`

### How it is obtained

1. Open the Ventas list.
2. **Remove the `Envíos de hoy` filter tag.** Without this the list shows only
   today's shipments — often 0 sales — and the Excel button stays `disabled`.
3. **Set the period dropdown to `Últimos 6 meses`.** This dropdown *is* the export
   range. Left at its default of 2 months it silently truncates: the 2-month
   export started 29 June and showed 2 orders for June; 6 months recovered 23
   orders from 17 June.
4. Click `Descargar Excel de ventas`.
5. Generation is **asynchronous**. A row appears in the corner notification
   widget reading `Generando archivo`, then `Archivo generado con éxito`
   (~25–30 s).
6. Click `Descargar` on that newest widget row.

### Things that bite

- **Dates are Spanish free text** — `3 de agosto de 2026 09:24 hs.` —
  `pd.to_datetime` returns `NaT` for every row, silently. Use
  `report_build.parse_spanish_date()`.
- **Revenue is gross and is never reduced on return.** A returned order keeps its
  full `Ingresos por productos`; only `Estado` changes and the money comes back in
  `Anulaciones y reembolsos`. `Total (MXN)` nets to 0 for a fully reversed order.
- **A month is never final.** Refunds land on the original order row whenever they
  settle, so re-downloading changes a closed month.
- **Multi-product packages split across rows.** The parent row carries the money
  with `Unidades = 0`; each item row carries `Unidades = 1` and **blank** money
  columns. Blank, not zero.

---

## Route 3 — Reportes de stock (Full inventory)

**Page:** <https://vendedores.mercadolibre.com.mx/publicaciones/listado/space_management>
(title: *Control de stock*)
**Menu path:** Publicaciones → Control de stock → `Descargar reportes de stock`
**Code:** `meli_forms.download_stock_reports()`

The button is a dropdown offering four reports, each behaving differently.

### 6. `<DD-MM-YY>_<DD-MM-YY>_Costos_por_servicio_almacenamiento.xlsx` ×2

Storage cost for one billing period. **Downloaded twice** — the current period
plus the newest one preceding it.

- 2 sheets, **header row 5**: `Resumen`, `Detalle`
- `Resumen`: cost by unit size — `Tamaño de la unidad`, `Tarifa diaria`,
  `Costos acumulados`, `Total`
- `Detalle`: per-SKU — `Código ML`, `Código universal`, `SKU`, `Producto`,
  `Tamaño de la unidad`, `Medidas`, `Peso` (44 columns)
- Filename carries the period, e.g. `01-08-26_31-08-26_…`

**How:** open the dropdown → `Reporte de costos por el servicio de almacenamiento`
→ open the `Período` dropdown → select a period → `Descargar`. Repeated once per
period, reopening the modal each time (it holds only one period at a time).

Period options come in two shapes, newest first:

```
Período actual - Del 1 de agosto al 31 de agosto del 2026 Acumulas $ 415.81 hasta el momento
Del 1 de julio al 31 de julio del 2026
Del 17 de junio al 30 de junio del 2026
```

⚠️ The `Período actual` entry is **multi-line** and carries a live accruing
amount. Match it with whitespace collapsed on both sides, or it is silently
skipped while the closed periods still work. It also renders **asynchronously** —
a short wait after opening the dropdown can show only the closed periods.

### 7. `stock_general_full_<id>_<hash>.xlsx`

Point-in-time snapshot of Full inventory. **Not a period report** — the header
reads `Actualizado el 31 de agosto`, which is why it has no calendar.

- 5 sheets, **header row 6**: `Resumen`, `Buena calidad`, `Para impulsar ventas`,
  `Para poner en venta`, `Para evitar descarte`
- Key columns: `Código ML`, `Código universal`, `SKU`, `# Publicación`,
  `Producto`, `Ventas últimos 30 días`
- `Resumen` also carries space usage (`Pequeños y medianos`, `Grandes y extragrandes`)

**How:** open the dropdown → click the entry. Downloads directly after
`Estamos preparando tu planilla` (~12 s).

### 8. `conciliation_<seller>_<id>_<hash>.xlsx`

Consolidated stock movements over a chosen date range.

- 1 sheet (literally named `{sheet_name}` — an unsubstituted template variable
  on MercadoLibre's side), **header row 3**, 29 columns
- Key columns: `Código universal`, `SKU`, `Código ML`, `ID de publicación`,
  `Producto`, `Ofrece Full`, `Stock total almacenado`

**How:** open the dropdown → click the entry → a range datepicker appears →
select start then end → `Aplicar` → `Descargar`.

⚠️ **The calendar has a rolling floor of roughly 60 days.** On 2026-08-29 the
earliest selectable day was 1 July; two days later it was 3 July. June was fully
disabled both times. "Two months back" is therefore not always selectable, so the
code clamps forward to the earliest *enabled* day rather than hard-coding a date.
Last run selected `2026-07-03 → 2026-08-31`.

⚠️ **Pick the earlier date first.** Once a start is selected every earlier day
becomes `--disabled`, so doing it backwards silently selects only one end.

### 9. `Returns_<DD-MM-YYYY>.xlsx`

Warehouse returns and their inspection outcomes.

- Sheet `Triages`, **header row 3**, 7 columns
- Columns: `Número de orden`, `SKU`, `IMEI`, `Fecha de revisión`,
  `Estado del producto`, `Resultado de la revisión`, `Estado del dinero`
- Values seen: `No funciona correctamente`, `Sello dañado` /
  `Producto para retirar` / `Reembolsamos el dinero`

**How:** open the dropdown → click the entry. Downloads directly (~3 s).

### Mechanics shared by route 3

- **The dropdown opens ONLY via a JS click.** A native Selenium click and an
  ActionChains click both leave `aria-expanded="false"`. This is the *opposite* of
  the billing `Reportes` tab, which needs a real click. Do not unify them.
- **Each menu `<li>` is inert.** The real target is an empty overlay
  `button.andes-list__item-actionable` inside it.
- **Calendar days carry full Spanish `aria-label`s** (`sábado 29 de agosto de
  2026`), so dates are selected by name, not by counting grid cells. The
  datepicker's own input is `readOnly`, so typing is impossible.

---

## Route 4 — MercadoPago reports

**Page:** <https://www.mercadopago.com.mx/balance/reports>
**Code:** `mercadopago.request_report(driver, report=...)` / `collect_report()`

Four report types live under `/balance/reports/<kind>`, in **three different UI
flavours**. They are not interchangeable, and the flavour is recorded per type in
`mercadopago.REPORTS`.

| key | URL | Menu path | Flavour |
|---|---|---|---|
| `settlement` | `/balance/reports/settlement_v2` | Reportes → Todas las transacciones | `typed` |
| `release` | `/balance/reports/release` | Reportes → Liberaciones | `typed` |
| `collection` | `/balance/reports/collection` | Reportes → Cobros | `simple` |
| `withdraw` | `/balance/reports/withdraw` | Reportes → Retiros | `simple` |
| `after_collection` | `/balance/reports/after_collection` | Reportes → Poscobro | `wizard` |
| `account_statement` | `/balance/reports/account_statement_generic` | Reportes → Estados de saldos y movimientos | `accordion` |

**`typed`** — results table; `Crear reporte` is a split button with a
`data-testid`; menu item `Manual` (also offers `Programado`); format chosen with
`#idCsv` / `#idXlsx` radios; the list has a `Tipo de reporte` filter
(`Manuales` / `Automáticos`); one row per format.

**`simple`** — results table; `Crear reporte` is a plain button found by text;
menu item `Crear` (also offers `Ajustes`); **no format choice** — the report is
produced in both and the row carries a download button for each; the third
column is `Estado del cobro`, so there is no `Automáticos` filter.

**`accordion`** — **no results table at all**. See report 15.

**`wizard`** — a two-step modal with its own row helpers. See report 14.

`typed` and `simple` share the range datepicker and the
`En preparación` → download-button transition, which is why one collect routine
serves both; `accordion` needs its own.

A different domain from the rest — the MercadoLibre session carries over via SSO,
no separate login.

### 10. `settlement_v2-<seller>-manual-<YYYY-MM-DD>-<HHMMSS>.csv`

Every money movement in the period — the payments-side counterpart to the
MercadoLibre sales and billing data.

- Requested in **csv**
- Filename encodes seller id, type (`manual`), and creation timestamp

### 11. `reserve-release-<seller>-manual-<YYYY-MM-DD>-<HHMMSS>.csv`

*Liberaciones* — when money held in reserve is released to the available
balance. Page title: **Reportes de Liberaciones**.

- Requested in **csv**, same 60-day window
- Obtained by exactly the same steps as report 10; only the URL differs
- Verified 2026-08-31: 83,920 bytes

### 12. Cobros — `collection`

Incoming payments. Same table-and-status mechanism as reports 10 and 11, but the
`simple` flavour: the create menu item is `Crear` rather than `Manual`, and there
is no format radio — MercadoPago produces both formats and the row carries a
download button for each, so the collect step picks the `.csv` one.

Verified end-to-end 2026-09-02. Requested with a 30-day window; generation took
~2.6 minutes.

### 13. `withdraw-<timestamp>-<hash>.csv`

*Retiros* — withdrawals to a bank account. Page title: **Reportes de Retiros**.
Same `simple` flavour as Cobros, 30-day window, plus one filter.

**The `Estado del cobro` filter does not exist until a period is chosen.** Before
that the modal has only `Período`, `Generar`, `Cancelar`; after `Aplicar` it
reveals *"Si lo requieres, puedes aplicar los siguientes filtros"* with a
`Seleccionar` dropdown. The code sets it to **`Todos los estados`** (options:
Todos los estados / Aprobados / En proceso / Rechazados / Cancelados).

Two things to know: MercadoPago labels it **"Estado del cobro"** in the modal
while this page's own column reads **"Estado del retiro"**; and the filter is
optional, so `Generar` is enabled without it — a run that silently failed to set
it would still produce a report, just differently scoped. The code verifies the
dropdown's value changed rather than trusting the click.

### 14. `after_collection-<timestamp>-<hash>.csv`

*Poscobro* — claims, chargebacks and returns. Page title: **Reportes de
Poscobro**. A **two-step wizard**, and the only report of its shape.

**Step 1** — `Crear reporte` → `Crear`, then a range datepicker on the left and
a multi-select `Tipo de operación` on the right (**Reclamos**, **Contracargos**,
**Devoluciones**). That dropdown has its own `Aplicar`. `Siguiente` stays
disabled until *both* a period and at least one operation type are set.

**Step 2** — five dropdowns (estado del reclamo / contracargo / devolución,
canal, herramienta de cobro) which already default to their `Todos …` values.
The code verifies each and only changes one that isn't. Then `Generar`.

⚠️ **Ticking the operations must be done one at a time.** Clicking re-renders the
list, so a JS loop over a NodeList captured up front holds stale nodes after the
first click — a single pass ticked only `Devoluciones`, and the report came back
scoped to one operation type. Worse, step 2 then showed only three filters
instead of five, which is the visible symptom to watch for.

⚠️ **The results table uses plain `.csv` / `.xlsx` buttons**, not
`button[class*=statement-button-download]`, so the shared row helpers see
nothing here — it has its own.

⚠️ **The row's period text is unlike any other page** — `4 julio 2026 -
2 septiembre 2026`, full month names — **and its end date is recorded a day
later than the one selected**. Rather than model that, the request reads the top
row's period back after `Generar` and matches on it verbatim.

### 15. `account_statement_generic-<uuid>.csv`

*Estados de saldos y movimientos* — the monthly balance-and-movements statement.
Page title: **Estados de saldos y movimientos**. Verified 2026-08-31.

Unlike every other MercadoPago report, this page has **no results table**.
Statements are **accordions**, one per period, newest first.

**Request:**

1. Click the button that opens the modal — see the trap below.
2. In `Generar estado de cuenta`: pick the newest **Período**, set **Formato**
   to `.csv`, click `Generar`.
3. A new accordion appears immediately reading **`EN PREPARACIÓN`**.

**Período is a discrete month dropdown**, not the range datepicker the other
types use: `Agosto de 2026` / `Julio de 2026` / `Junio de 2026` /
`Periodo personalizado`. The last is a range option, not a month, and is
excluded when picking "newest". The modal says the range may be at most 31 days.

**Formato defaults to `.pdf`.** Leaving it alone silently produces the wrong
file. Options are `.pdf` / `.xlsx` / `.csv`.

**Collect:** find the accordion whose header contains the month
(`1 agosto 2026 - 31 agosto 2026` for `Agosto de 2026`), expand it, and read the
format rows:

```
.pdf    Generar     <- that format has not been produced
.xlsx   Generar
.csv    Abrir       <- ready; clicking Abrir downloads it
```

**`Abrir` is both the readiness signal and the download.** There is no separate
status text once the accordion exists.

### Things that bite on report 15

- **Three different things are labelled "Generar"** — the promo-card button that
  opens the modal, the modal's own submit (`andes-ui-button--large`), and the
  per-format link inside every accordion. Plus a fourth button,
  `Generar nuevo estado`.
- **The exact-`Generar` button is conditional.** It sits on a promo card shown
  only while the newest period has *no* statement yet; once one exists it
  disappears and only `Generar nuevo estado` remains. Both open the same modal,
  so the code tries the exact one and falls back to the other.
- **`Generar nuevo estado` is itself `--large`.** Excluding that class to avoid
  the modal's submit also excludes the toolbar button — which is exactly how
  this failed twice during development. `--large` may only be excluded for the
  ambiguous exact-`Generar`.
- **40+ zero-width `Generar` textlinks** exist on the page, one per format in
  every collapsed accordion, so a width test is mandatory.
- **The format row is `.card-list__item`.** `closest('div')` from the link stops
  at an inner wrapper containing only the link text, so the label is not visible
  from there.

### How reports 10-14 are obtained

The report list is the whole mechanism. Every report ever generated is listed
with `Fecha de creación | Período de fechas | Tipo de reporte | Formato`, and the
**format column doubles as the status column**:

| While generating | When ready |
|---|---|
| `En preparación` (plain text) | a download button (`.csv` / `.xlsx`) |

That transition is the only completion signal — polling that cell is how phase 3
knows the file exists.

**Day 6 onwards** (`request`): `Crear reporte` → `Manual` → pick the last 60 days
in the range datepicker → select csv → `Generar`. Records the range string
(`2/jul/2026 a 31/ago/2026`) as the row identity.

**Days 1–5**: prefer the automatic report MercadoPago generates on the 1st —
filter `Tipo de reporte` → `Automáticos` and take the newest. **If none exists,
fall back to requesting a manual one.** That fallback is not theoretical: on this
account the Automáticos filter returns *"No encontramos reportes con los filtros
que aplicaste"* — every existing report is typed `Manual`, including the ones on
2/jul and 3/ago that look exactly like scheduled monthly reports. The `Programado`
option beside `Manual` under `Crear reporte` is where a schedule would be set up;
it has not been touched.

**Collect:** return to the page, find the row by **range and format**, poll until
its download button exists, click it.

### Things that bite

- **A period can be listed twice**, once per format — the existing list has two
  rows for `1/ago/2026 a 28/ago/2026`, one `.xlsx` and one `.csv`. Matching on
  range alone grabs whichever comes first, so the row lookup matches range **and**
  format.
- **Never grab "the newest row".** During testing that downloaded a three-day-old
  report while the freshly requested one still said `En preparación`.
- **Datepicker aria-labels embed the date mid-string** —
  `lunes 27 de julio de 2026, Día no seleccionado, Inicio de rango`. `endsWith`
  fails here (it works on the stock picker), and a bare `includes` would match
  `2 de julio` inside `22 de julio`, so the match uses a leading space.
- Two months are displayed at once, unlike the stock picker's single month.
- Generation took **~6.5 minutes** for 60 days when requested alone. With both
  report types requested together it ran longer — one was still `En preparación`
  after 10 minutes while the other was ready. Budget accordingly, and do not
  treat a timeout as failure.
- MercadoPago also emails when ready (*"Te enviaremos un mail cuando el reporte
  esté listo"*), but the list is the reliable signal.
- **Collect in more than one pass.** Draining the whole timeout on the first
  pending report starves the rest — in testing the settlement report consumed all
  600s while the Liberaciones report sat ready behind it. `run_downloads.py`
  gives each pending item an equal slice first, then spends what is left on the
  stragglers.

---

## Cross-cutting traps

These are not tied to one form. Each cost a failed run.

### A ready report can still read as "not ready"

Three ways the collect loop waited on a report that was sitting there finished.
All three printed the same thing - `not ready yet`, once per cycle, against a
page plainly showing the report ready.

**One row hid another.** The matcher answered from the *first* row whose period
matched and returned immediately, so a row with no download buttons ended the
search and a later row with the same period that *was* ready never got looked
at. Duplicate periods are the norm: every run asks for the same rolling window,
so a second run the same day creates a second row for it. `JS_FIND_ROW` now
scans **every** matching row and answers with the best of them - ready in the
wanted format beats ready in another, which beats a terminal state, which beats
still-generating.

**The format we asked for never arrived.** MercadoPago decides which formats a
report comes out in. A finished report offering only `.xlsx` when we asked for
`.csv` was skipped with `continue`, the loop fell off the end returning `null`,
and `probe()` read that as "still generating" - forever, because that report is
done and will never grow a `.csv` button. Seen live on UNIT_TA04:

```
Todas las transacciones
  [2] state=available btns=.xlsx   ['2/sep/2026', '1/ago/2026 a 31/ago/2026']
```

`probe()` now takes whatever format exists, switches `pending["fmt"]` to it and
says so. The file is the point, not its extension.

**The site recorded a different period than we chose.** Matching is on the
period string, so if the row says something else, nothing ever matches. The
pages do not even agree with each other - same store, same day, 2026-09-03:

| Page | Created | Period recorded |
|---|---|---|
| Liberaciones | 3/sep | `5/jul/2026 a 2/sep/2026` |
| Retiros | 3/sep | `4/ago/2026 a 3/sep/2026` |

and Poscobro records an end date a day *later* than the one selected.
`_confirm_row_period()` now gives the row a few seconds to appear under the
period we picked and, failing that, adopts whatever the newest row says - which
is what the wizard flavour always had to do.

A fourth row state turned up while checking this: **`--processed`**, on
Liberaciones, carrying a normal download button. It needs no special handling,
but it confirms the state list is open-ended, which is why readiness is decided
by "does a download button exist" rather than by enumerating known states.

### "Still generating" and "will never generate" look identical

Two different failures render as a row with an orange `!` and no download
button. Neither is coming while a run waits, and they are **not the same
failure**:

| Row class | Tooltip | Meaning |
|---|---|---|
| `statements-table__row--empty` | *No pudimos generar tu reporte porque no hay **movimientos** en este rango de fechas.* | Permanent for that range |
| `statements-table__row--delayed` | *No pudimos generar tu reporte porque hay **datos en proceso** en este momento. Espera unas horas a que te notifiquemos.* | Transient, on a scale of **hours** |
| `statements-table__row--available` | — | Ready; has a real download button |

Neither failed row looks distinctive to a naive check. The `--empty` one still
renders the text `.csv .xlsx` inside an inert
`[data-testid=download-error-container]` (so matching on cell text calls it
ready), and neither has a `button[class*=statement-button-download]` (so
matching on buttons calls it *still generating*, and polls it for the entire
1800s budget while genuinely-generating reports queue behind it).

`probe()` returns `READY` / `WAITING` / `EMPTY` / `DELAYED`, reading the row's
own `statements-table__row--<state>` class. Two details matter:

- **Match the bare one-word suffix.** The row also carries a long variant
  (`statements-table__row--settlement_v2-delayed-manual--not-seen`), so the
  regex is anchored: `^statements-table__row--([a-z]+)$`.
- **Read the class, not the tooltip.** The tooltip needs a hover and is Spanish
  prose MercadoPago can reword; the class is already in the DOM.

The collect loop then drops both from polling, reported apart:

```
    mercadopago/withdraw         no movements in range - skipping
    mercadopago/settlement       data in process at MercadoPago - retry on a later run
```

`--empty` lands under **`[no data]`** and leaves the store `ok` — nothing went
wrong, there was nothing to report. `--delayed` lands under **`[pending]`**
with `- datos en proceso`, because the request stands on MercadoPago's side and
a later run collects it.

Verified on UNIT_TA04 2026-09-03, all four states including a ready row, which
must keep returning `READY` — a false positive here silently discards a good
report, which is worse than waiting.

### A store can be blocked from reports entirely

MercadoPago reports are gated per **collaborator login**, not per store. If the
login behind a store lacks the permission, every report page renders

> No tienes permisos para ver los reportes
> Pedile al administrador de la cuenta que desde "Colaboradores" te otorgue
> permiso para *Acceder a reportes de tus cobros y facturación* y *Acceder a
> reportes de operaciones*.

and there is no `Crear reporte` button anywhere. Found on **EWTTO_SM**
(`angela.a@shenming.mx`) on 2026-09-02 while BOCINA_SM ran clean.

Without a check this looks exactly like a broken selector — six consecutive
`[warn] 'Crear reporte' not found` lines — and sends you hunting for a markup
change that never happened. `_require_access()` now reads the page before
anything is clicked and raises `NoReportAccess`, which names the permission and
who has to grant it. The caller stops after the first report, since the wall is
on the login and the other five would fail identically: 33 seconds instead of
3m25s, with an actionable message.

**This is an account-admin fix, not a code fix.** Nothing in the pipeline can
work around it.

### A route that quietly returns nothing is a failure

The `request_*` helpers warn and return `None` rather than raising, so one bad
report cannot sink the other five. That is right, but it must be *recorded*: on
2026-09-02 all six MercadoPago requests returned `None` on EWTTO_SM and the
store still reported `ok, 13 files`. A route asked to run and producing nothing
now counts as an error, so the store reports `partial` and the batch exits 1.

### The site's "today" is not this machine's today

This box runs on **CST (UTC+8)**; MercadoPago México is **UTC-6**. For much of
the day the local clock is a **day ahead**, so an end date built from
`datetime.date.today()` does not exist yet on the site — the calendar cell comes
back `disabled` or `not-found` and the whole range selection fails:

```
local clock          2026-09-02
site calendar        "Hoy, martes 1 de septiembre de 2026"
end 2026-09-02 -> not-found
```

`_select_range()` reads the calendar's own `--today` cell, clamps the end to it,
and **returns the dates actually selected** so callers rebuild their row-match
string from reality:

```
end 2026-09-02 is ahead of the site's today (2026-09-01) - clamping
range actually selected: 4/jul/2026 a 1/sep/2026
```

This affects every MercadoPago date-range report. It only went unnoticed because
early runs happened when the two dates coincided.

### The download path must be absolute

CDP's `Browser.setDownloadBehavior` resolves a **relative** `downloadPath`
against the *browser process's* working directory, not ours. Passing
`--out downloads/_fix` made every route appear to click correctly and then
report `got 0 files`, with the files nowhere findable. `run_downloads.py` now
runs `--out` through `os.path.abspath()`.

### MercadoLibre changes markup without notice

The stock-reports menu entries were `li.andes-button-dropdown__menu-item` on
2026-08-31 and plain `li.andes-list__item` on 2026-09-02. The whole stock route
returned `menu has: []` while the menu was demonstrably open
(`aria-expanded="true"`). Selectors there now accept **either** class.

The lesson: when a route reports "found nothing", check whether the control is
present under a *different* class before assuming a timing problem. Polling
longer would never have fixed this one.

### React re-renders invalidate a NodeList mid-loop

Ticking several checkboxes in one JS pass only ticks the first — the remaining
nodes in the captured NodeList are stale after the first click re-renders the
list. Click one per call, from Python, with a pause between. See report 14.

---

## Not yet automated

Seen in the panel but not downloaded. Add sections here as they are built.

| Form | Page | Note |
|---|---|---|
| Cargos por stock antiguo | `/publicaciones/listado/fee_storage` | Tab next to Control de stock |
| Reporte de planificación de envíos | stock reports menu | **New** — appeared in the menu on 2026-09-02, a 5th entry |
| Facturas fiscales | `/billing/detail/<period>` | Tab beside Reportes; PDF/XML rather than xlsx |
| Reporte de publicidad | Product Ads console | Ad spend detail beyond the billing line |

---

## Change log

| Date | Change |
|---|---|
| 2026-08-28 | Routes 1 and 2 built and verified |
| 2026-08-29 | Ventas period widened from 2 to 6 months, recovering early June |
| 2026-08-31 | Route 3 (4 stock reports) built and verified |
| 2026-08-31 | Report 6 extended to 2 periods (current + preceding) |
| 2026-08-31 | Route 4 (MercadoPago settlement) added; run split into request/main/collect phases |
| 2026-08-31 | Liberaciones added; MercadoPago route generalised to any report type |
| 2026-08-31 | Cobros (`simple`) and Estados de saldos y movimientos (`accordion`) added |
| 2026-09-02 | Retiros and Poscobro added; full pipeline verified |
| 2026-09-02 | Phase 3 reworked into one shared wait; timeout default 600s → 1800s |
| 2026-09-02 | Fixed: timezone-ahead end dates, relative download path, changed stock menu markup, stale-NodeList ticking |
| 2026-09-02 | `run_batch.py` added: several stores per client session. Both entry points now share `run_store()`, and a broken route no longer aborts the run |
| 2026-09-02 | MercadoPago permission wall detected explicitly (`NoReportAccess`); a request that returns nothing now counts as a failure |
| 2026-09-03 | Reports with no movements in range are flagged `[no data]` and skipped instead of polled to timeout |
| 2026-09-03 | `--delayed` rows ("datos en proceso") also skipped, reported under `[pending]` for a later run |
| 2026-09-03 | Each run now writes `REPORT.md` into its own download folder |
| 2026-09-03 | `run_batch.py` now runs the `downloads/mx_sales` cleaning pipeline after the stores (`--no-transform` to skip) |
| 2026-09-04 | `run_batch.py --workbooks`: download → clean → one accounting workbook per store, unattended |
| 2026-09-09 | Batch split into per-store download + a late collect pass; slow MercadoPago reports no longer block other stores. New `logs/pending.json` registry and `--collect-only`. |
| 2026-09-09 | Ventas verified to survive a store close/reopen, so it is deferrable too |
| 2026-09-09 | Accounting workbooks now written per store as each finishes, not batched at the end |
| 2026-09-03 | Fixed three ways a finished report read as "not ready": first-row-wins matching, an unavailable format, and a period the site recorded differently |
