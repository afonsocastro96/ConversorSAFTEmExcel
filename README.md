# Conversor SAF-T → Excel

A standalone tool that converts a Portuguese SAF-T (Standard Audit File for Tax) XML export into
a single Excel workbook, one sheet per SAF-T section and one row per record.

![Conversor SAF-T → Excel](https://i.imgur.com/ZCKkq7z.png)

## Contents

- `saft_to_excel.py` — the converter itself (also runnable as a CLI).
- `saft_to_excel_gui_tk.py` — a small desktop GUI (tkinter/ttk) wrapping the converter.
- `generate_exe.py` — builds a standalone Windows `.exe` of the GUI via PyInstaller.

## Usage

### Command line

```
python saft_to_excel.py <saft.xml> [output.xlsx]
```

If `output.xlsx` is omitted, the workbook is written next to the input file with the same base
name.

### GUI

```
python saft_to_excel_gui_tk.py
```

Pick the SAF-T XML file and a destination folder, then click "Gerar Excel". The conversion runs
in a background process, so "Cancelar" can stop it immediately. There's an option to open the
generated workbook automatically once it's done.

### Standalone executable

```
python generate_exe.py
```

Builds a onefile, windowed `.exe` (no console window) named `Conversor SAFT em Excel.exe`,
requiring nothing but Windows to run — no Python install needed on the target machine.

## What's in the workbook

Every run produces 18 sheets, in this order:

| Sheet | Contents |
|---|---|
| Header | The SAF-T `Header` block (company/file metadata), one field per row |
| Totals | The file's own declared totals and checksums — `GeneralLedgerAccounts`' `TaxonomyReference`, and each of `GeneralLedgerEntries`/`SalesInvoices`/`MovementOfGoods`/`WorkingDocuments`/`Payments`' `NumberOfEntries`/`TotalDebit`/`TotalCredit`/etc. |
| Journals | Every `Journal`'s ID and description |
| GeneralLedgerAccounts | The chart of accounts, with opening/closing balances |
| Customers | Customer master data, including billing/shipping addresses |
| Suppliers | Supplier master data, including billing/shipping addresses |
| Products | Product master data |
| TaxTable | The tax code table |
| Transactions | General ledger transactions (journal entries) |
| MovementLines | Each transaction's debit/credit lines |
| Invoices | Sales invoices |
| InvoiceLines | Invoice line items |
| StockMovements | Stock movement documents (delivery notes, etc.) |
| StockMovementsLines | Stock movement line items |
| WorkDocuments | Working documents (quotes, proformas, etc.) |
| WorkDocumentsLines | Working document line items |
| Payments | Payment documents |
| PaymentLines | Payment line items |

A file with no `SourceDocuments` data (a common export shape for accounting-only SAF-Ts) still
gets all 18 sheets — the document-related ones are simply empty except for their header row.

Nested single-occurrence blocks (an address, a document's status, a tax breakdown, ...) are
flattened onto the parent record's row as prefixed columns (e.g. `BillingAddress_City`). Elements
that genuinely repeat within one record (an invoice line's `References`, for example) are
concatenated into a single cell rather than spread across numbered columns.

Money, quantity and percentage columns are written as real numbers (2 decimal places, no currency
symbol); counts are written as whole numbers; date/datetime columns are written as real Excel
dates. Everything else — IDs, codes, free text — stays as plain text.

Sheet tabs are colour-coded by section, columns are auto-fit to their content, and every sheet has
a bold header row. If a sheet would exceed Excel's 1,048,576-row limit, the converter continues
onto an additional sheet (`_2`, `_3`, ...) with the same header row rather than dropping rows.

## Integrity check

After conversion, the script cross-checks the row counts and debit/credit sums it wrote against
each section's own declared totals (the same figures shown on the Totals sheet) and prints a
warning to stderr for any mismatch. This is a sanity check on the workbook's own output, not proof
that the source file itself is correct — a mismatch can come from either.

It also warns (to stderr, without stopping) whenever a field appears in the source file under a
known record type but isn't one of that sheet's columns, so a gap never passes silently.

## Requirements

- `lxml` and `xlsxwriter` for the converter itself.
- `sv-ttk` for the GUI's theming.
- `pyinstaller` to build the standalone executable.

```
pip install -r requirements.txt
```
