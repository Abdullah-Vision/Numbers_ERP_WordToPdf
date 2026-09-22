# Digital Invoice Template Engine — Design

**Date:** 2026-09-16
**Status:** Approved (SQL Server first, Oracle-ready)

## Goal

Fill a user-uploaded Word template whose placeholders match the column aliases of the
`SpArDigitalInvoices` stored procedure, then return a downloadable PDF. When a new alias is
added to the procedure, no code change is required — the engine discovers columns at runtime.

## Requirements (from user)

1. Placeholders in the Word template use the **same names as the procedure aliases**
   (`{InvoiceNo}`, `{CustomerName}`, `{Qty}`, ...).
2. The engine fetches the procedure's result columns **at runtime** (ADO.NET
   `GetColumnSchema()`); no DTOs, no manual mapping, no rebuild when the proc changes.
3. Line-item table rows in Word are **auto-detected**: the engine computes which aliases
   have values that **vary across the result set** (detail columns, e.g. Qty, Rate); any
   table row containing at least one detail-column token is cloned and **appended once per
   data row** (N items = N rows); the template row is removed afterwards. Rows containing
   only invariant master aliases (e.g. `{CompanyName}`, `{GrandTotal}`) are filled once
   from the first row, so supplier/buyer/totals tables are never repeated.
4. Placeholders outside repeated rows take the **first row's** values (master fields).
5. Optional format specifier in the placeholder: `{InvoiceDate:dd-MMM-yy}`, `{Qty:N2}`.
   No format → smart defaults (dates `dd-MMM-yy`, decimals with thousand separators).
6. Matching is **case-insensitive**. Unknown placeholders are left visible and reported
   back as warnings.
7. **SQL Server now**; Oracle-ready design. Deliverable includes an Oracle PL/SQL script
   (`SP_AR_DIGITAL_INVOICES(p_id IN NUMBER, p_result OUT SYS_REFCURSOR)`) with quoted
   aliases preserving SQL Server casing.
8. Output: single PDF, returned as a file download.
9. Procedure name is fixed **in configuration** (invoice-specific endpoints), not code.

## Non-Goals

- Generic "any procedure" API (user chose invoice-specific).
- Oracle execution testing (no Oracle access available now).
- Saving generated PDFs to disk/DB — response-only download.

## Architecture

Existing solution: `WordToPDF.sln` (ASP.NET Core 8). Reuses existing patterns:
template storage via `SqlTemplateService`/`WordTemplates` table, `{Token}` replacement via
OpenXML (`ProcessParagraphForPlaceholders`-style run merging), LibreOffice headless
`soffice.exe` DOCX→PDF conversion.

### Components

```
DigitalInvoiceController            (new, [Route("DigitalInvoice")])
  ├─ GET  placeholders?dbType       → column alias list for template authoring
  ├─ POST templates                 → upload/overwrite template (reuses SqlTemplateService)
  └─ POST generate                  → invoiceId + dbType + templateId → PDF download

IDigitalInvoiceService / DigitalInvoiceService   (new)
  ├─ Task<List<Dictionary<string, object>>> GetInvoiceRowsAsync(int invoiceId, string dbType)
  └─ Task<byte[]> FillTemplateAsync(byte[] templateBytes, List<rows>)

IInvoiceDbExecutor
  ├─ SqlServerInvoiceExecutor       (EXEC SpArDigitalInvoices @id via Microsoft.Data.SqlClient)
  └─ OracleInvoiceExecutor          (SP_AR_DIGITAL_INVOICES with SYS_REFCURSOR out param)
```

### Configuration (appsettings.json)

```json
"DigitalInvoice": {
  "SqlServerProcedure": "SpArDigitalInvoices",
  "OracleProcedure": "SP_AR_DIGITAL_INVOICES"
}
```

## Data Flow (generate)

1. Load template bytes by `templateId` (`ITemplateService.GetTemplateStreamAsync`).
2. Execute proc for `invoiceId` via the executor chosen by `dbType`.
3. Build `List<Dictionary<string, object>>` keyed by column alias (as returned).
4. OpenXML pass:
   a. Find all table rows whose text contains at least one `{...}` token → repeated rows.
   b. For each repeated row: clone per data row, fill row-level placeholders
      (row values), insert clones before template row; remove template row.
   c. Replace remaining `{...}` tokens in body/headers/footers whose aliases exist with
      first-row values (run-merging replacement so tokens split across runs still match);
      unknown aliases stay untouched and are collected for the warnings header.
5. Save DOCX to temp → LibreOffice `--headless --convert-to pdf` → read PDF bytes.
6. Return `File(pdfBytes, "application/pdf", $"Invoice_{invoiceId}.pdf")`; cleanup temp dir.

## Placeholder Rules

- Syntax: `{Alias}` or `{Alias:format}` (`.NET` standard format string).
- Value formatting: `DateTime` uses format or default `dd-MMM-yy`; `decimal`/numeric uses
  format or default `N2`; everything else `ToString()`. `null`/DBNull → empty string.
- Alias matching is case-insensitive (`invoiceno` ≡ `InvoiceNo`).
- A placeholder whose alias does not exist in the result set is left **untouched** and its
  name is included in the response's `X-Unresolved-Placeholders` header (comma-separated).
- Repeated-row detection: a column is a **detail column** when at least two rows exist and
  its values differ across rows (case/format-insensitive string comparison); a table row is
  a **detail row** when it contains a token for a detail column. Single-row result sets
  produce no detail columns, so the items row is filled once. Totals rows must therefore
  use master-level aliases (e.g. `{GrandTotal}`); add SUM aliases to the procedure if line
  aggregates are needed — demonstrating the zero-code-change workflow.

## Error Handling

| Case | Result |
|---|---|
| Template id not found / file missing | 404 |
| `invoiceId` has no rows | 404 `"No data found for invoice {id}"` |
| Non-.docx upload | 400 |
| LibreOffice missing / conversion failure | 500 with soffice stderr |
| DB failure | 500 with message (no connection string leaked) |

## Testing

- Manual via Swagger with the delivered `DigitalInvoiceTemplate.docx` (SALES TAX INVOICE
  layout mirroring the user's screenshot) against `invoiceId=5410` — values must match the
  filled screenshot (1,271,187 / 228,814 / 1,500,001 etc.).
- `GET placeholders` must list every alias of the current proc; adding a temp alias to the
  proc and re-calling must show it without code changes.
- Multi-item invoice → one PDF table row per item.
- Unit-testable core kept pure: token parsing/format resolution + row detection helpers.

## Deliverables

1. `DigitalInvoiceTemplate.docx` in `WordTemplates/` — SALES TAX INVOICE layout with
   placeholders (header block, supplier/buyer block, items table with repeating row,
   totals/footer block).
2. `Database/SpArDigitalInvoicesOracle.sql` — PL/SQL procedure (REF CURSOR, quoted aliases).
3. Controller, service, executors, DI wiring in `Program.cs`, config section.
