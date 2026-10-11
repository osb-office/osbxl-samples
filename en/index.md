[日本語](../ja/index.md)

# osbxl Sample Collection

These osbxl samples have all been executed and verified. The code on each page runs as-is. Every code block switches between **five language tabs — Node.js / Python / Java / PHP / WebAssembly** — and the install instructions shown per tab are language-specific too.

**If you are new here, start with [general-01 Overview — object model and basic flow](general-01-overview.md).**

## Find a sample by what you want to do

| What you want to do | Sample |
| --- | --- |
| Understand the big picture of osbxl and the class hierarchy | [general-01 Overview](general-01-overview.md) |
| The flow of creating a new book / opening an existing book / save / saveAs / close | [general-01 Overview — basic flow](general-01-overview.md#basic-flow) |
| What methods App / WorkBook / WorkSheet / Range have | [general-01 Overview — main methods](general-01-overview.md#main-methods) |
| Build a table with values and formulas | [general-02 Values and formulas](general-02-values-formulas.md) |
| Enter per-row shifting formulas (relative references) in bulk | [general-02 Values and formulas](general-02-values-formulas.md#code) / [general-06 Bulk input](general-06-bulk-input.md) |
| Add, delete, move or rename sheets | [general-03 Sheet operations](general-03-sheets.md) |
| List all sheets of a book whose sheet names you don't know | [general-03 Sheet operations](general-03-sheets.md#code) |
| Enter dates, times and booleans / set number formats | [general-04 Cell value types](general-04-value-types.md) |
| The difference between `getValue()` and `getValue(true)` | [general-04 Cell value types](general-04-value-types.md#getvalue-vs-getvalue-true) |
| Insert, delete, clear or copy rows and columns | [general-05 Rows, columns and cell ranges](general-05-rows-cols.md) |
| Set row height, column width and hidden state | [general-05 Rows, columns and cell ranges](general-05-rows-cols.md#code-2-row-col-objects) |
| Populate tens of thousands of cells fast | [general-06 Bulk input](general-06-bulk-input.md) |
| Filter or sort data (AutoFilter / Sort) | [general-07 AutoFilter and sorting](general-07-autofilter-sort.md) |
| Use tables (totals row, structured references) and comments | [general-08 Tables and comments](general-08-tables-comments.md) |
| Set conditional formatting (color scales, data bars, top N) | [general-09 Conditional formatting](general-09-conditional-format.md) |
| Create charts (column, line) | [general-10 Charts](general-10-charts.md) |
| Aggregate with pivot tables | [general-11 Pivot tables](general-11-pivottables.md) |
| Detailed Feature Comparison & Benchmarks vs POI / openpyxl / PHP | [Benchmark Comparison Report](benchmark-comparison.md) |

## general — overview and basic operations

| # | Title | level | Main APIs |
| --- | --- | --- | --- |
| 01 | [Overview — object model and basic flow](general-01-overview.md) | basic | App / WorkBook / WorkSheet / Range |
| 02 | [Building a Table with Values and Formulas](general-02-values-formulas.md) | basic | Range / InputFormulaObject |
| 03 | [Working with books and sheets — add, delete, move, rename](general-03-sheets.md) | basic | WorkBook / WorkSheet |
| 04 | [Cell value types and number formats — strings, numbers, dates, times, booleans](general-04-value-types.md) | basic | Range / DateTimeObject |
| 05 | [Rows, columns and cell ranges — insert, delete, clear, copy](general-05-rows-cols.md) | basic | WorkSheet / Row / Col |
| 06 | [Bulk input of large data — setValueArray / setFormulaArray](general-06-bulk-input.md) | intermediate | WorkSheet / InputValueObject / InputFormulaObject |
| 07 | [AutoFilter and sorting — conditional filters, sort](general-07-autofilter-sort.md) | intermediate | AutoFilter / Sort |
| 08 | [Tables and comments — ListObjects, totals row, structured references](general-08-tables-comments.md) | intermediate | ListObjects / Table / Comments |
| 09 | [Conditional formatting — cell value rules, color scales, data bars](general-09-conditional-format.md) | intermediate | FormatConditions / ColorScaleObject / DataBarObject |
| 10 | [Charts basics — column, line, title, axes, legend](general-10-charts.md) | intermediate | Chart / chart.\* / enums.XlChartType |
| 11 | [Pivot tables — create, aggregate, change aggregation method](general-11-pivottables.md) | advanced | PivotTable / PivotField |

## Planned (not yet written)

### general

Book properties (`setCompanyName` / `setCreateAuthor` etc.) / named ranges (`getNames`) / custom styles (`addCustomStyle` / `addCustomTableStyle`) / external workbook integration (`addExternalWorkBook` / `updateExternalWorkBook`) / password-protected books (`openPasswordWorkBook` / `save(password)`) / outlining in detail

### format

Fonts and colors / borders / fills and gradients / alignment and wrapping / number formats (numeric, date, currency) / built-in and custom styles (for conditional formatting see [general-09](general-09-conditional-format.md))

### data

Cell merging / hyperlinks (for tables, sorting, AutoFilter and comments see [general-07](general-07-autofilter-sort.md) / [general-08](general-08-tables-comments.md))

### formula

Formula basics and array formulas / financial functions (ACCINT etc.) / lookup functions (XLOOKUP) / forecasting (FORECAST.ETS) / WEBSERVICE + FILTERXML

### pivot

Aggregation and grouping (for the basics of creating and aggregating see [general-11](general-11-pivottables.md))

### chart

Chart varieties / data labels / formatting (colors, lines, shadows) (for basic creation, axes and legend see [general-10](general-10-charts.md))

### advanced

Page setup and printing / protection / windows and gridlines / Shapes
