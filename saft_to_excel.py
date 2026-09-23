"""
saft_to_excel.py -- standalone converter from a Portuguese SAF-T (Standard
Audit File for Tax) XML export into a single Excel workbook, one sheet per
SAF-T section, one row per record.

Standalone: this script does not import anything from AuditPilot's app/
package and has no dependency on engagement-folder/template conventions --
it is a generic "dump the file into readable tabs" tool, not part of the
audit pipeline (no prior-year comparison, no commentary).

Design notes:
- Namespace-agnostic: the SAF-T XML declares a version-specific default
  namespace (e.g. urn:OECD:StandardAuditFile-Tax:PT_1.04_01). Every tag is
  matched with the lxml wildcard form "{*}TagName" so the script keeps
  working regardless of schema version.
- Streaming: some real Portuguese SAF-T exports run to hundreds of MB. The
  XML is read with lxml.etree.iterparse, clearing each top-level record
  element right after it is written so memory stays bounded by one
  record's own subtree rather than the whole file. The workbook is written
  with xlsxwriter in "constant_memory" mode -- streams each worksheet to
  disk as rows are written rather than holding it in memory, the same idea
  as openpyxl's write_only mode this script used before (see git history);
  switched because xlsxwriter turned out to be both faster (no per-cell
  style-dedup lookup to fight -- formats are plain objects this script
  creates once and reuses by reference, see Formats) and simpler (native
  set_column() for auto-fit widths and set_tab_color(), rather than the
  post-save zip-editing openpyxl's write_only mode needed for the first
  and an 8-hex-digit alpha-channel gotcha for the second).
- Column headers are a fixed list per sheet, authored from the SAF-T PT
  schema (not discovered per file), so column order is stable across runs.
  A field that isn't in these lists -- because it belongs to some future
  schema revision, or because the schema allows more than one occurrence of
  an element that every sample file only ever had once (e.g. a customer's
  ShipToAddress) -- won't get a column; SheetWriter prints a one-time
  warning per sheet/field instead of silently dropping it, so a gap is
  never silent. Add the field to the relevant header list if it's ever hit.
- A handful of schema elements are genuinely repeatable within one record
  (e.g. an invoice line's References, up to 13 deep in real files) even
  though almost everything else in the schema that technically allows
  repetition never does in practice. Rather than a separate lookup sheet or
  numbered columns, repeated occurrences are concatenated into one cell.
- Four sheets describe the file/workbook as a whole rather than a list of
  records -- Estrutura (a static legend explaining what every other sheet
  contains and where it comes from, transcribed from Estrutura.xlsx (the
  reference sheet this was authored from, not included in this repo); see
  ESTRUTURA_ROWS/populate_estrutura_sheet()), Header (the SAF-T Header
  block), Totals (the file's own declared totals, one section per
  MasterFiles/SourceDocuments/GeneralLedgerEntries wrapper -- including
  MasterFiles/GeneralLedgerAccounts's own TaxonomyReference, which isn't a
  total/count like its five siblings but sits in the same "field(s) before
  the record list" position), and Journals (ID + description) -- the latter
  three laid out vertically, one row per field/entry, since e.g. Header's
  ~30 fields read far worse as one wide row than as a two-column Field/Value
  table. See _KeyValueSheetBuilder and populate_header_sheet()/
  populate_totals_sheet()/populate_journals_sheet(). All four share one tab
  colour and sit first in the workbook, Estrutura ahead of the other three.
  Estrutura also has its own screen+print gridlines turned off and every row/
  column past its own content hidden (see populate_estrutura_sheet()),
  unlike every other sheet in this workbook.
- Every sheet has a bold header row, left-aligned cells throughout, and a
  fixed font (Calibri 10pt, chosen for being both easy to read and a
  standard Windows/Office font). A column whose name marks it as a count, a
  money/quantity/rate figure, or a date gets converted from XML text to a
  real Excel number/date with a matching format (integer, 2-decimal, or
  date/datetime); see _column_kind()/_cell_value_and_format(). Tab colours
  group sheets by section per SHEET_TAB_COLORS. Column widths are auto-fit
  to each column's longest value, applied once all rows are written (see
  SheetWriter.finalize_column_widths()).
- A sheet that would exceed Excel's 1,048,576-row limit doesn't drop the
  overflow: SheetWriter starts a new worksheet part (base name + "_2",
  "_3", ...) with the same header row and keeps going, so every record
  ends up in the workbook somewhere -- split across sheets an auditor
  needs to page through together, never silently missing.

Usage:
    python saft_to_excel.py <saft.xml> [output.xlsx]
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import date, datetime
from pathlib import Path

import xlsxwriter
from lxml import etree

# Calibri is Excel's own long-time default -- legible on screen, ships with
# every Windows/Office install, and (unlike Verdana, already used throughout
# AuditPilot's own working papers) gives this standalone tool its own look.
# Every cell is left-aligned, including numbers/dates that Excel would
# otherwise right-align by default under "General" alignment -- requested
# so the whole workbook reads as a consistent left-aligned table.
FONT_NAME = "Calibri"
FONT_SIZE = 10

# Column-width auto-fit, mirroring AuditPilot's own Teste._autofit_column
# (app/excel_writer/teste.py): openpyxl can only count characters, not
# measure rendered pixels, so this is an approximation, calibrated the same
# way -- by eye against real output, not derived from font metrics.
CHARACTER_WIDTH_FACTOR = 0.67
MIN_COLUMN_WIDTH = 8
MAX_COLUMN_WIDTH = 60

# Plain number formats -- no currency symbol, per the "not as currency"
# instruction: just a decimal count for whole numbers and 2 decimal places
# for everything else numeric (money, quantities, percentages, rates alike).
FORMAT_INTEGER = "#,##0"
FORMAT_DECIMAL = "#,##0.00"
FORMAT_DATE = "yyyy-mm-dd"
FORMAT_DATETIME = "yyyy-mm-dd hh:mm:ss"

# Joins a single occurrence's own sub-fields when a repeatable element (see
# module docstring) is rendered into one cell, e.g. an OrderReferences with
# OriginatingON="GR 187024/5" and OrderDate="2024-01-05" becomes
# "GR 187024/5: 2024-01-05".
FIELD_SEPARATOR = ": "
# Joins multiple occurrences of the same repeatable element into one cell.
OCCURRENCE_SEPARATOR = " | "

# Excel's hard per-sheet row limit (2**20 = 1,048,576); one row is the header.
MAX_DATA_ROWS_PER_SHEET = 1_048_576 - 1

# Wrapper elements under MasterFiles/SourceDocuments/GeneralLedgerEntries that
# carry their own summary field(s) before the list of actual records -- these
# feed the Totals sheet. Most carry counts/totals (NumberOfEntries,
# TotalDebit, ...); MasterFiles/GeneralLedgerAccounts is the odd one out --
# its only summary field is TaxonomyReference (a single descriptive code, not
# a count), sitting in exactly the same "before the record list" position.
# Mapped to a human label for the "Table" field.
TOTALS_WRAPPER_TAGS = {
    "GeneralLedgerAccounts": "GeneralLedgerAccounts",
    "GeneralLedgerEntries": "GeneralLedgerEntries",
    "SalesInvoices": "SalesInvoices",
    "MovementOfGoods": "MovementOfGoods",
    "WorkingDocuments": "WorkingDocuments",
    "Payments": "Payments",
}
# The tag that starts a wrapper's list of actual records -- summary fields
# always appear before this tag, so flattening stops there.
_TOTALS_RECORD_TAGS = {"Account", "Journal", "Invoice", "StockMovement", "WorkDocument", "Payment"}

# Expected immediate parent tag for each top-level record type, used to tell
# a real record apart from a same-named element nested elsewhere in the
# schema (see the comment where this is checked, in convert()).
_EXPECTED_PARENT_TAG = {
    "Journal": "GeneralLedgerEntries",
    "Account": "GeneralLedgerAccounts",
    "Customer": "MasterFiles",
    "Supplier": "MasterFiles",
    "Product": "MasterFiles",
    "TaxTableEntry": "TaxTable",
    "Transaction": "Journal",
    "Invoice": "SalesInvoices",
    "StockMovement": "MovementOfGoods",
    "WorkDocument": "WorkingDocuments",
    "Payment": "Payments",
}


def local_name(tag) -> str | None:
    """Strip the namespace from an lxml tag, e.g. '{urn:...}Header' -> 'Header'.

    Returns None for comments/processing instructions, whose .tag is a
    callable rather than a string -- the one place this script touches a
    client-supplied file directly, so this guard matters.
    """
    return tag.rpartition("}")[2] if isinstance(tag, str) else None


# A small set of schema-repeatable elements whose observed count genuinely
# varies from record to record within the same file (an invoice line's
# OrderReferences might have 1 occurrence on one line and 3 on another).
# Flattening such a
# tag "as recursed columns when there's 1, concatenated when there's >1"
# would give different rows in the same sheet different shapes, so these
# always take the concatenated form, even for a single occurrence -- unlike
# e.g. a customer's ShipToAddress, which the schema also allows more than
# one of but which never actually repeats in any real file seen so far, so
# it's left as the more readable expanded-columns form (see module
# docstring: an unexpected second occurrence there is caught by
# SheetWriter's unknown-field warning rather than silently mis-shaped).
ALWAYS_CONCATENATE_TAGS = {"OrderReferences", "References", "ProductSerialNumber"}


def flatten_record(element, skip_tags: frozenset[str] = frozenset()) -> dict[str, str]:
    """Flatten one XML record element into a flat {column: value} dict.

    A child tag that appears once among its siblings recurses into a
    prefixed set of columns if it has children of its own (e.g.
    BillingAddress/City -> 'BillingAddress_City'), or becomes '<tag>': text
    if it's a leaf.

    A child tag that appears more than once, or is in ALWAYS_CONCATENATE_TAGS,
    is a repeatable group: each occurrence is rendered as its own leaf values
    joined by FIELD_SEPARATOR, and the occurrences are joined by
    OCCURRENCE_SEPARATOR into one cell under the plain tag name.

    `skip_tags` names direct children to leave out entirely -- used to keep
    a document's own <Line>/<Lines> children (handled separately, as their
    own sheet) out of the parent record's row.
    """
    children_by_tag: dict[str, list] = {}
    for child in element:
        tag = local_name(child.tag)
        if tag is None or tag in skip_tags:
            continue
        children_by_tag.setdefault(tag, []).append(child)

    row: dict[str, str] = {}
    for tag, occurrences in children_by_tag.items():
        if len(occurrences) == 1 and tag not in ALWAYS_CONCATENATE_TAGS:
            child = occurrences[0]
            if len(child):
                for sub_column, value in flatten_record(child).items():
                    row[f"{tag}_{sub_column}"] = value
            else:
                row[tag] = (child.text or "").strip()
        else:
            row[tag] = OCCURRENCE_SEPARATOR.join(_render_occurrence(child) for child in occurrences)
    return row


def _render_occurrence(element) -> str:
    if not len(element):
        return (element.text or "").strip()
    return FIELD_SEPARATOR.join(value for value in flatten_record(element).values() if value)


def _child_text(element, tag: str) -> str:
    for child in element:
        if local_name(child.tag) == tag:
            return (child.text or "").strip()
    return ""


# --- Cell typing/formatting -------------------------------------------------
# Every value starts life as XML text. A column whose name marks it as a
# whole-number count, a monetary/quantity/rate figure, or a date gets
# converted to a real Excel number/date (so it sorts, sums and displays
# properly) with the matching format; anything else -- IDs, codes, free
# text -- stays plain text, since forcing e.g. an AccountID through int()
# would silently drop a meaningful leading zero.

# Columns holding a whole-number count -- rendered with no decimal places.
_INTEGER_COLUMNS = {
    "LineNumber", "Period", "HashControl", "FiscalYear",
    "NumberOfEntries", "NumberOfMovementLines",
}
# Substrings in a column name meaning "a decimal number, 2 decimal places,
# no currency symbol" -- covers monetary amounts, quantities, tax
# percentages and exchange rates alike, per the user's two-bucket
# formatting request. Matched as a substring, not a suffix, since several
# real column names carry the keyword in the middle rather than at the end
# (Totals-sheet fields like TotalDebit/TotalCredit/TotalQuantityIssued).
_DECIMAL_KEYWORDS = (
    "Amount", "Total", "Payable", "Balance", "Price", "Percentage",
    "Quantity", "ExchangeRate", "Debit", "Credit",
)

_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_DATETIME_RE = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$")


def _column_kind(header: str) -> str:
    """Classify a column by name into how its values should be typed:
    'integer', 'decimal', 'date_like' (a date or a datetime -- decided per
    value at write time, since the same column family, e.g. *StatusDate,
    holds either depending on the record), or 'text' (left alone).
    """
    if header in _INTEGER_COLUMNS:
        return "integer"
    if any(keyword in header for keyword in _DECIMAL_KEYWORDS):
        return "decimal"
    if "Date" in header:
        return "date_like"
    return "text"


class Formats:
    """The handful of distinct cell styles this workbook ever uses, each
    created once via workbook.add_format() and reused by reference for
    every cell -- xlsxwriter formats are plain objects the caller manages
    explicitly (no per-cell attribute assignment triggering an internal
    style-dedup lookup the way openpyxl's write_only mode did; see this
    file's git history for the hash-caching workaround that needed).
    """

    def __init__(self, workbook: xlsxwriter.Workbook):
        base = {"font_name": FONT_NAME, "font_size": FONT_SIZE, "align": "left"}
        self.text = workbook.add_format(base)
        self.bold = workbook.add_format({**base, "bold": True})
        self.integer = workbook.add_format({**base, "num_format": FORMAT_INTEGER})
        self.decimal = workbook.add_format({**base, "num_format": FORMAT_DECIMAL})
        self.date = workbook.add_format({**base, "num_format": FORMAT_DATE})
        self.datetime = workbook.add_format({**base, "num_format": FORMAT_DATETIME})
        # The Estrutura sheet's header row (see populate_estrutura_sheet) --
        # "pattern": 1 is a solid fill, whose *visible* colour is fg_color,
        # not bg_color (an OOXML quirk: bg_color only matters for non-solid
        # patterns) -- confirmed against Estrutura.xlsx's own header cells.
        self.estrutura_header = workbook.add_format(
            {**base, "bold": True, "pattern": 1, "fg_color": "#D9D9D9"}
        )


def _cell_value_and_format(header: str, text: str, formats: Formats):
    """Convert `text` to a real number/date with the matching format when
    the column calls for it and the value actually parses that way;
    otherwise returns it as plain text. Never raises -- a value that
    doesn't fit its column's expected shape is just left as text rather
    than dropped or crashing the run.
    """
    text = text or ""
    if not text:
        return None, formats.text

    kind = _column_kind(header)
    if kind == "integer":
        try:
            return int(float(text)), formats.integer
        except ValueError:
            pass
    elif kind == "decimal":
        try:
            return float(text), formats.decimal
        except ValueError:
            pass
    elif kind == "date_like":
        if _DATETIME_RE.match(text):
            return datetime.fromisoformat(text), formats.datetime
        if _DATE_RE.match(text):
            return date.fromisoformat(text), formats.date

    return text, formats.text


# --- Tab colours -------------------------------------------------------------
# The "Theme Colors, Darker 25%" row from Excel's classic Office theme
# palette, pixel-sampled directly from the user's own reference screenshot
# rather than derived from theory (the row doesn't reduce to one uniform
# lumMod/lumOff transform across every column, so sampling was the reliable
# path). xlsxwriter's set_tab_color() wants a plain "#RRGGBB" string --
# unlike openpyxl, which needed an explicit "FF" alpha-channel prefix or it
# defaulted to "00" (fully transparent, no visible tab colour at all), and
# silently produces a corrupt file if given that 8-digit aRGB form instead.
_COLOR_BLUE = "#8DB3E2"
_COLOR_GREEN = "#D7E3BC"
_COLOR_RED = "#E5B9B7"
_COLOR_AQUA = "#B7DDE8"
_COLOR_ORANGE = "#FBD5B5"
_COLOR_TAN = "#C4BD97"
_COLOR_PURPLE = "#CCC1D9"
_COLOR_LIGHT_GREY = "#D8D8D8"

# Header isn't named individually in the user's colour scheme -- blue, same
# as the rest of the "describes the file as a whole" content it now holds
# (the workbook's own Estrutura legend, the SAF-T Header block, the file's
# declared Totals, and the Journal list).
SHEET_TAB_COLORS = {
    "Header": _COLOR_BLUE,
    "GeneralLedgerAccounts": _COLOR_GREEN,
    "Customers": _COLOR_RED,
    "Suppliers": _COLOR_RED,
    "Products": _COLOR_RED,
    "TaxTable": _COLOR_RED,
    "Transactions": _COLOR_AQUA,
    "MovementLines": _COLOR_AQUA,
    "Invoices": _COLOR_ORANGE,
    "InvoiceLines": _COLOR_ORANGE,
    "StockMovements": _COLOR_TAN,
    "StockMovementsLines": _COLOR_TAN,
    "WorkDocuments": _COLOR_PURPLE,
    "WorkDocumentsLines": _COLOR_PURPLE,
    "Payments": _COLOR_LIGHT_GREY,
    "PaymentLines": _COLOR_LIGHT_GREY,
}


def totals_row(wrapper_element, section_name: str) -> dict[str, str]:
    """Flatten a SourceDocuments/GeneralLedgerEntries wrapper's own summary
    fields (everything before its list of actual records) into one row.
    """
    row = {"Table": section_name}
    for child in wrapper_element:
        tag = local_name(child.tag)
        if tag is None:
            continue
        if tag in _TOTALS_RECORD_TAGS:
            break
        row[tag] = (child.text or "").strip()
    return row


def _estimated_width(longest_length: int) -> float:
    return min(max(longest_length * CHARACTER_WIDTH_FACTOR + 2, MIN_COLUMN_WIDTH), MAX_COLUMN_WIDTH)


class SheetWriter:
    """Wraps one or more xlsxwriter worksheets for a single logical sheet:
    writes the fixed header row up front, and maps each row dict onto that
    column order, leaving unknown/absent fields blank. Warns once per
    sheet/field for a value that doesn't match any known column.

    If a sheet would exceed Excel's per-sheet row limit, a new worksheet
    part is started automatically (base name + "_2", "_3", ...) with the
    same header row, rather than dropping rows past the limit -- every
    record ends up in the workbook somewhere, just split across sheets an
    auditor needs to page through together.

    Also tracks each column's longest value as rows come in, for
    finalize_column_widths() -- xlsxwriter's set_column() can be called at
    any time, unlike openpyxl's write_only mode, which streamed each
    worksheet's XML to disk as soon as the first row was appended and
    needed a post-save workaround for this (see this file's git history).
    """

    def __init__(
        self, workbook: xlsxwriter.Workbook, base_title: str, headers: list[str],
        formats: Formats, tab_color: str | None = None,
    ):
        self._workbook = workbook
        self._base_title = base_title
        self._headers = headers
        self._header_set = set(headers)
        self._formats = formats
        self._tab_color = tab_color
        self._sheets = []
        self._next_row = 0  # next row index (0-based) to write in the current part
        self._rows_in_current_sheet = 0
        self._column_max_len = [len(header) for header in headers]
        self._warned_unknown_fields: set[str] = set()
        self._start_new_sheet()

    def _start_new_sheet(self) -> None:
        part_number = len(self._sheets) + 1
        # "_2", "_3", ... for further parts -- comfortably inside Excel's
        # own 31-character sheet-name limit even for the longest base name
        # this script uses ("GeneralLedgerAccounts", 22 chars).
        title = self._base_title if part_number == 1 else f"{self._base_title}_{part_number}"
        if part_number > 1:
            print(
                f"INFO: sheet '{self._base_title}' reached Excel's {MAX_DATA_ROWS_PER_SHEET:,} "
                f"row limit -- continuing on '{title}' rather than dropping rows.",
                file=sys.stderr,
            )
        sheet = self._workbook.add_worksheet(title)
        if self._tab_color:
            sheet.set_tab_color(self._tab_color)
        for col, header in enumerate(self._headers):
            sheet.write(0, col, header, self._formats.bold)
        self._sheets.append(sheet)
        self._next_row = 1
        self._rows_in_current_sheet = 0

    def append(self, row: dict[str, str]) -> None:
        unknown_fields = set(row) - self._header_set - self._warned_unknown_fields
        for field in unknown_fields:
            print(
                f"WARNING: sheet '{self._base_title}' has a value for "
                f"'{field}', which isn't one of its known columns -- add it "
                "to the header list in tools/saft_to_excel.py if this "
                "matters. Value dropped for this and any further row.",
                file=sys.stderr,
            )
        self._warned_unknown_fields |= unknown_fields

        if self._rows_in_current_sheet >= MAX_DATA_ROWS_PER_SHEET:
            self._start_new_sheet()

        sheet = self._sheets[-1]
        for col, header in enumerate(self._headers):
            text = row.get(header, "")
            value, fmt = _cell_value_and_format(header, text, self._formats)
            sheet.write(self._next_row, col, value, fmt)
        self._next_row += 1
        self._rows_in_current_sheet += 1

        for i, header in enumerate(self._headers):
            value = row.get(header, "")
            if value:
                self._column_max_len[i] = max(self._column_max_len[i], len(str(value)))

    def finalize_column_widths(self) -> None:
        """Applies the auto-fit widths computed from every row seen so far
        to every part sheet this logical sheet ended up split across.
        Called once, after the very last append() for this sheet.
        """
        widths = [_estimated_width(n) for n in self._column_max_len]
        for sheet in self._sheets:
            for col, width in enumerate(widths):
                sheet.set_column(col, col, width)


# --- Fixed, schema-authored header lists, one per sheet --------------------


def _prefixed(prefix: str, fields: list[str]) -> list[str]:
    return [f"{prefix}_{field}" for field in fields]


# Some exporters add structured BuildingNumber/StreetName fields alongside
# the free-text AddressDetail (both present, same content); always kept,
# never assumed.
_ADDRESS_FIELDS = [
    "BuildingNumber", "StreetName", "AddressDetail", "City", "PostalCode", "Region", "Country",
]

HEADERS_HEADER = [
    "AuditFileVersion", "CompanyID", "TaxRegistrationNumber", "TaxAccountingBasis",
    "CompanyName", "BusinessName",
    *_prefixed("CompanyAddress", _ADDRESS_FIELDS),
    "FiscalYear", "StartDate", "EndDate", "CurrencyCode", "DateCreated", "TaxEntity",
    "ProductCompanyTaxID", "SoftwareCertificateNumber", "ProductID", "ProductVersion",
    "HeaderComment", "Telephone", "Fax", "Email", "Website",
]

# The fields a Totals-sheet row can carry (see totals_row()), in display
# order. Not every section has every field (e.g. MovementOfGoods has no
# TotalDebit) -- populate_totals_sheet skips absent ones.
_TOTALS_FIELD_ORDER = [
    "TaxonomyReference", "NumberOfEntries", "TotalDebit", "TotalCredit",
    "NumberOfMovementLines", "TotalQuantityIssued",
]

HEADERS_GENERAL_LEDGER_ACCOUNTS = [
    "AccountID", "AccountDescription",
    "OpeningDebitBalance", "OpeningCreditBalance",
    "ClosingDebitBalance", "ClosingCreditBalance",
    "GroupingCategory", "GroupingCode", "TaxonomyCode",
]

HEADERS_CUSTOMERS = [
    "CustomerID", "AccountID", "CustomerTaxID", "CompanyName", "Contact",
    *_prefixed("BillingAddress", _ADDRESS_FIELDS),
    *_prefixed("ShipToAddress", _ADDRESS_FIELDS),
    "Telephone", "Fax", "Email", "Website", "SelfBillingIndicator",
]

HEADERS_SUPPLIERS = [
    "SupplierID", "AccountID", "SupplierTaxID", "CompanyName", "Contact",
    *_prefixed("BillingAddress", _ADDRESS_FIELDS),
    *_prefixed("ShipFromAddress", _ADDRESS_FIELDS),
    "Telephone", "Fax", "Email", "Website", "SelfBillingIndicator",
]

HEADERS_PRODUCTS = [
    "ProductType", "ProductCode", "ProductGroup", "ProductDescription", "ProductNumberCode",
]

HEADERS_TAX_TABLE = [
    "TaxType", "TaxCountryRegion", "TaxCode", "Description",
    "TaxExpirationDate", "TaxPercentage", "TaxAmount",
]

HEADERS_TRANSACTIONS = [
    "JournalID", "TransactionID", "Period", "TransactionDate", "SourceID", "Description",
    "DocArchivalNumber", "TransactionType", "GLPostingDate", "CustomerID", "SupplierID",
]

HEADERS_MOVEMENT_LINES = [
    "TransactionID", "RecordID", "AccountID", "SourceDocumentID", "SystemEntryDate",
    "Description", "DebitAmount", "CreditAmount",
    "CurrencyCode", "CurrencyAmount", "ExchangeRate",
]

# Shared "delivery" block used by ShipTo/ShipFrom on Invoice/StockMovement/WorkDocument.
_SHIP_FIELDS = ["DeliveryID", "DeliveryDate", "WarehouseID", "LocationID", *_prefixed("Address", _ADDRESS_FIELDS)]
# Note: the line above pre-prefixes "Address" onto the address sub-fields, then
# ShipTo/ShipFrom prefixes the whole block again below -- see _prefixed usage.

_DOCUMENT_TOTALS_FIELDS = [
    "TaxPayable", "NetTotal", "GrossTotal",
    "Currency_CurrencyCode", "Currency_CurrencyAmount", "Currency_ExchangeRate",
    "Settlement_SettlementAmount", "Settlement_PaymentTerms",
    "Payment_PaymentMechanism", "Payment_PaymentAmount", "Payment_PaymentDate",
]

HEADERS_INVOICES = [
    "InvoiceNo", "ATCUD",
    "DocumentStatus_InvoiceStatus", "DocumentStatus_InvoiceStatusDate", "DocumentStatus_Reason",
    "DocumentStatus_SourceID", "DocumentStatus_SourceBilling",
    "Hash", "HashControl", "Period", "InvoiceDate", "InvoiceType",
    "SpecialRegimes_SelfBillingIndicator", "SpecialRegimes_CashVATSchemeIndicator",
    "SpecialRegimes_ThirdPartiesBillingIndicator",
    "SourceID", "EACCode", "SystemEntryDate", "TransactionID", "CustomerID",
    *_prefixed("ShipTo", _SHIP_FIELDS),
    *_prefixed("ShipFrom", _SHIP_FIELDS),
    "MovementEndTime", "MovementStartTime",
    "WithholdingTax_WithholdingTaxType", "WithholdingTax_WithholdingTaxAmount",
    *_prefixed("DocumentTotals", _DOCUMENT_TOTALS_FIELDS),
]

HEADERS_INVOICE_LINES = [
    "InvoiceNo", "LineNumber", "OrderReferences", "ProductCode", "ProductDescription",
    "Quantity", "UnitOfMeasure", "UnitPrice", "TaxPointDate", "References", "Description",
    "ProductSerialNumber", "DebitAmount", "CreditAmount",
    "Tax_TaxType", "Tax_TaxCountryRegion", "Tax_TaxCode", "Tax_TaxPercentage", "Tax_TaxAmount",
    "TaxExemptionReason", "TaxExemptionCode", "SettlementAmount", "CustomsInformation_IECAmount",
]

HEADERS_STOCK_MOVEMENTS = [
    "DocumentNumber", "ATCUD",
    "DocumentStatus_MovementStatus", "DocumentStatus_MovementStatusDate", "DocumentStatus_Reason",
    "DocumentStatus_SourceID", "DocumentStatus_SourceBilling",
    "Hash", "HashControl", "Period", "MovementDate", "MovementType", "SystemEntryDate",
    "TransactionID", "CustomerID", "SupplierID", "SourceID", "EACCode",
    *_prefixed("ShipTo", _SHIP_FIELDS),
    *_prefixed("ShipFrom", _SHIP_FIELDS),
    "MovementEndTime", "MovementStartTime", "ATDocCodeID", "MovementComments",
    *_prefixed("DocumentTotals", _DOCUMENT_TOTALS_FIELDS),
]

HEADERS_STOCK_MOVEMENTS_LINES = [
    "DocumentNumber", "LineNumber", "ProductCode", "ProductDescription", "Quantity",
    "UnitOfMeasure", "UnitPrice", "Description", "OrderReferences", "References",
    "ProductSerialNumber", "DebitAmount", "CreditAmount",
    "Tax_TaxType", "Tax_TaxCountryRegion", "Tax_TaxCode", "Tax_TaxPercentage", "Tax_TaxAmount",
    "TaxExemptionReason", "TaxExemptionCode", "SettlementAmount", "CustomsInformation_IECAmount",
]

HEADERS_WORK_DOCUMENTS = [
    "DocumentNumber", "ATCUD",
    "DocumentStatus_WorkStatus", "DocumentStatus_WorkStatusDate", "DocumentStatus_Reason",
    "DocumentStatus_SourceID", "DocumentStatus_SourceBilling",
    "Hash", "HashControl", "Period", "WorkDate", "WorkType",
    "SourceID", "EACCode", "SystemEntryDate", "TransactionID", "CustomerID",
    *_prefixed("ShipTo", _SHIP_FIELDS),
    *_prefixed("ShipFrom", _SHIP_FIELDS),
    "MovementEndTime", "MovementStartTime",
    *_prefixed("DocumentTotals", _DOCUMENT_TOTALS_FIELDS),
]

HEADERS_WORK_DOCUMENTS_LINES = [
    "DocumentNumber", "LineNumber", "ProductCode", "ProductDescription", "Quantity",
    "UnitOfMeasure", "UnitPrice", "TaxPointDate", "References", "Description",
    "ProductSerialNumber", "DebitAmount", "CreditAmount",
    "Tax_TaxType", "Tax_TaxCountryRegion", "Tax_TaxCode", "Tax_TaxPercentage", "Tax_TaxAmount",
    "TaxExemptionReason", "TaxExemptionCode", "SettlementAmount", "CustomsInformation_IECAmount",
]

HEADERS_PAYMENTS = [
    "PaymentRefNo", "ATCUD", "Period", "TransactionID", "TransactionDate", "PaymentType",
    "SystemID",
    "DocumentStatus_PaymentStatus", "DocumentStatus_PaymentStatusDate", "DocumentStatus_Reason",
    "DocumentStatus_SourceID", "DocumentStatus_SourcePayment",
    "PaymentMethod_PaymentMechanism", "PaymentMethod_PaymentAmount", "PaymentMethod_PaymentDate",
    # A few real exporters write PaymentMechanism/PaymentAmount/PaymentDate as
    # direct children of Payment instead of wrapped in PaymentMethod -- kept
    # as separate bare columns so that variant is captured too, not dropped.
    "PaymentMechanism", "PaymentAmount", "PaymentDate",
    "SourceID", "SystemEntryDate", "CustomerID", "SupplierID",
    *_prefixed("DocumentTotals", _DOCUMENT_TOTALS_FIELDS),
]

HEADERS_PAYMENT_LINES = [
    "PaymentRefNo", "LineNumber",
    "SourceDocumentID_OriginatingON", "SourceDocumentID_InvoiceDate", "SourceDocumentID_Description",
    "SettlementAmount", "DebitAmount", "CreditAmount",
    "Tax_TaxType", "Tax_TaxCountryRegion", "Tax_TaxCode", "Tax_TaxPercentage", "Tax_TaxAmount",
    "TaxExemptionReason", "TaxExemptionCode",
]

# Section order the wrapper elements appear in within the file -- used to
# order the Header sheet's Totals block the same way.
_TOTALS_SECTION_ORDER = [
    "GeneralLedgerAccounts", "GeneralLedgerEntries", "SalesInvoices",
    "MovementOfGoods", "WorkingDocuments", "Payments",
]


# --- Estrutura sheet (workbook legend) --------------------------------------
# Transcribed from Estrutura.xlsx -- the reference sheet this was authored
# from (not included in this repo) describing what each generated sheet
# contains and which SAF-T section(s) it's sourced from. A
# 1-tuple is a bold section header (e.g. "1. Informação geral"); a 4-tuple is
# a data row: (Folha, Conteúdo, "SAF-T onde consta", Descrição).
#
# One correction from the source workbook: its last row names the sheet
# "PaymentsLines" (with an s after Payment), but the sheet this script
# actually writes is "PaymentLines" (see HEADERS_PAYMENT_LINES/SHEET_TAB_
# COLORS) -- fixed here since a legend that misnames the sheet it's
# describing would mislead more than it helps.
ESTRUTURA_ROWS: list[tuple[str, ...]] = [
    ("1. Informação geral",),
    ("Header", "Cabeçalho", "Ambos", "Informação geral alusiva à Entidade a que respeita o SAF-T."),
    ("Totals", "Totais", "Ambos", "Informação sobre todos os totalizadores constantes no SAF-T."),
    ("Journals", "Diários", "Contabilidade", "Informação sobre os diários da contabilidade."),
    ("2. Tabelas mestres",),
    (
        "GeneralLedgerAccounts", "Tabela de código de contas", "Contabilidade",
        "Tabela com o código de contas previsto pelo SNC. Atua como balancete analítico.",
    ),
    ("Customers", "Tabela de clientes", "Ambos", "Informação do ficheiro de clientes da Entidade."),
    ("Suppliers", "Tabela de fornecedores", "Faturação", "Informação do ficheiro de fornecedores da Entidade."),
    (
        "Product", "Tabela de produtos/serviços", "Faturação",
        "Informação com o catálogo de produtos e tipos de serviços prestados que foram objeto de movimentação.",
    ),
    (
        "TaxTable", "Tabela de impostos", "Ambos",
        "Informação com os registos fiscais de IVA e as rúbricas de imposto de selo a liquidar.",
    ),
    ("3. Movimentos contabilísticos",),
    (
        "Transactions", "Lançamentos contabilísticos", "Contabilidade",
        "Lançamentos contabilísticos correspondentes ao período de exportação. Atua em conjunto com a "
        "tabela MovementLines como extrato.",
    ),
    (
        "MovementLines", "Linhas dos lançamentos contabilísticos", "Contabilidade",
        "Linhas dos lançamentos constantes na tabela Transactions. Atua em conjunto com esta como extrato.",
    ),
    ("4. Documentos comerciais",),
    (
        "Invoices", "Documentos comerciais a clientes", "Faturação",
        "Documentos de venda e retificativos emitidos pela Entidade, como faturas e notas de crédito.",
    ),
    (
        "InvoiceLines", "Linhas dos documentos comerciais", "Faturação",
        "Linhas dos documentos listados na tabela Invoices.",
    ),
    (
        "StockMovements", "Documentos de movimentação de mercadorias", "Faturação",
        "Guias de transporte ou de remessa que sirvam de documento de transporte.",
    ),
    (
        "StockMovementsLines", "Linhas dos documentos de movimentação", "Faturação",
        "Linhas dos documentos listados na tabela StockMovements.",
    ),
    (
        "WorkDocuments", "Documentos de conferência", "Faturação",
        "Documentos apresentados ao cliente para conferência de mercadorias ou prestação de serviços.",
    ),
    (
        "WorkDocumentsLines", "Linhas dos documentos de conferência", "Faturação",
        "Linhas dos documentos listados na tabela WorkDocuments.",
    ),
    ("Payments", "Documentos de recibos emitidos", "Ambos", "Recibos emitidos."),
    (
        "PaymentLines", "Linhas dos documentos de pagamento", "Ambos",
        "Linhas dos documentos listados na tabela Payment.",
    ),
]

# Character-width column widths A-F, matching Estrutura.xlsx's own layout --
# A and F are narrow margin columns, E is wide enough to fit the longest
# Descrição on one line without wrapping.
_ESTRUTURA_COLUMN_WIDTHS = [0.6, 19.4, 35.46, 14.4, 100.4, 0.93]
_ESTRUTURA_HEADER_LABELS = ("Folha", "Conteúdo", "SAF-T onde consta", "Descrição")
# Blank spacer row below the last data row, matching Estrutura.xlsx's own
# compact row height there (its other rows are left at Excel's default
# height). The source workbook also has one above the header row, but that's
# deliberately dropped here -- the header starts at row 1 instead.
_ESTRUTURA_BOTTOM_SPACER_HEIGHT = 5.75
# 1-based Excel row: every row below this is hidden (see populate_estrutura_
# sheet's use of set_default_row(hide_unused_rows=True)).
_ESTRUTURA_LAST_VISIBLE_ROW = 24


def populate_estrutura_sheet(worksheet, formats: Formats) -> None:
    """Fills the Estrutura sheet: a static legend (not derived from the SAF-T
    being converted) describing what each of this workbook's other sheets
    contains. See ESTRUTURA_ROWS for the content and where it comes from.
    """
    worksheet.hide_gridlines(2)  # off on screen and when printing
    for col, width in enumerate(_ESTRUTURA_COLUMN_WIDTHS):
        worksheet.set_column(col, col, width)
    # Columns after F: hidden in one call, native to set_column -- no need
    # for the per-row workaround below.
    worksheet.set_column(6, 16383, None, None, {"hidden": True})

    row = 0  # Excel row 1: the header row
    # The grey fill spans the full A:F width, including the narrow A/F margin
    # columns either side of the labelled B:E ones -- written as blank cells
    # rather than left untouched so the fill still shows on those columns.
    worksheet.write_blank(row, 0, None, formats.estrutura_header)
    for col, label in enumerate(_ESTRUTURA_HEADER_LABELS):
        worksheet.write(row, 1 + col, label, formats.estrutura_header)
    worksheet.write_blank(row, 5, None, formats.estrutura_header)
    row += 1

    for item in ESTRUTURA_ROWS:
        if len(item) == 1:
            worksheet.write(row, 1, item[0], formats.bold)
        else:
            for col, text in enumerate(item):
                worksheet.write(row, 1 + col, text, formats.text)
        row += 1

    # One more blank row stays visible below the data, matching the source
    # workbook's own padding, before everything from _ESTRUTURA_LAST_VISIBLE_
    # ROW onward gets hidden. Also gets a blank written cell, not just
    # set_row(): in constant_memory mode a row's properties are only ever
    # flushed to the file when some later write() call passes through it (see
    # worksheet.py's _write_single_row) -- a set_row() with no write() to
    # that row, or to a row after it, is silently dropped rather than shown
    # with default formatting. write_blank() gives this row that later
    # write() to hang the flush on.
    worksheet.set_row(row, _ESTRUTURA_BOTTOM_SPACER_HEIGHT)
    worksheet.write_blank(row, 1, None, formats.text)
    assert row + 1 == _ESTRUTURA_LAST_VISIBLE_ROW, "content grew past the sheet's hidden-row cutoff"

    # Rows past _ESTRUTURA_LAST_VISIBLE_ROW are hidden by leaving them
    # untouched: hide_unused_rows zeroes the height of any row that was never
    # written to or explicitly set_row()'d, which is one worksheet-level flag
    # rather than an explicit set_row(..., hidden=True) call for each of the
    # ~1,048,550 rows that would otherwise need it.
    worksheet.set_default_row(hide_unused_rows=True)


class _KeyValueSheetBuilder:
    """Shared builder for the three small vertical Field/Value-style sheets
    that describe the file as a whole (Header, Totals, Journals) rather
    than a list of records. Created early (see convert(): all three sit
    first in the workbook, but their data isn't known until the main
    parsing loop finishes) and populated late -- xlsxwriter worksheets can
    be written to in any order relative to each other, so this needs no
    equivalent of openpyxl's move_sheet() workaround. Tracks each column's
    longest value as rows come in, for finalize_column_widths(). Never
    needs to split across multiple sheets in practice (a file's own
    Header/Totals/Journal data is at most a few dozen rows), unlike
    SheetWriter, which real per-record sheets do need.
    """

    def __init__(self, workbook: xlsxwriter.Workbook, title: str, formats: Formats, tab_color: str | None):
        self.sheet = workbook.add_worksheet(title)
        if tab_color:
            self.sheet.set_tab_color(tab_color)
        self._formats = formats
        self._row = 0
        self._column_max_len = [0, 0]

    def _track(self, *texts: str) -> None:
        for i, text in enumerate(texts):
            if text:
                self._column_max_len[i] = max(self._column_max_len[i], len(str(text)))

    def field_value_row(self, field: str, value: str) -> None:
        self._track(field, value)
        self.sheet.write(self._row, 0, field, self._formats.bold)
        cell_value, fmt = _cell_value_and_format(field, value, self._formats)
        self.sheet.write(self._row, 1, cell_value, fmt)
        self._row += 1

    def data_row(self, first_column: str, first_value: str, second_column: str, second_value: str) -> None:
        """Like field_value_row, but neither cell is bold -- for a row that's
        an actual record (e.g. one journal's ID and description) rather than
        a field label paired with its value.
        """
        self._track(first_value, second_value)
        first, first_fmt = _cell_value_and_format(first_column, first_value, self._formats)
        second, second_fmt = _cell_value_and_format(second_column, second_value, self._formats)
        self.sheet.write(self._row, 0, first, first_fmt)
        self.sheet.write(self._row, 1, second, second_fmt)
        self._row += 1

    def bold_row(self, *texts: str) -> None:
        self._track(*texts)
        for col, text in enumerate(texts):
            self.sheet.write(self._row, col, text, self._formats.bold)
        self._row += 1

    def finalize_column_widths(self) -> None:
        widths = [_estimated_width(n) for n in self._column_max_len]
        for col, width in enumerate(widths):
            self.sheet.set_column(col, col, width)


def populate_header_sheet(builder: _KeyValueSheetBuilder, header_row: dict[str, str]) -> None:
    """The SAF-T Header block, laid out vertically (Field/Value pairs, one
    row per field) rather than one ~30-column-wide row.
    """
    builder.bold_row("Field", "Value")
    for field in HEADERS_HEADER:
        builder.field_value_row(field, header_row.get(field, ""))
    builder.finalize_column_widths()


def populate_totals_sheet(builder: _KeyValueSheetBuilder, declared_totals: dict[str, dict]) -> None:
    """The file's own declared Totals, one section per SourceDocuments/
    GeneralLedgerEntries wrapper (see totals_row()).
    """
    for section in _TOTALS_SECTION_ORDER:
        declared = declared_totals.get(section)
        if declared is None:
            continue
        builder.bold_row("Table", section)
        for field in _TOTALS_FIELD_ORDER:
            if field in declared:
                builder.field_value_row(field, declared[field])
    builder.finalize_column_widths()


def populate_journals_sheet(builder: _KeyValueSheetBuilder, journals: list[dict]) -> None:
    """The Journal list (ID + description) from GeneralLedgerEntries."""
    builder.bold_row("JournalID", "Description")
    for journal in journals:
        builder.data_row("JournalID", journal.get("JournalID", ""), "Description", journal.get("Description", ""))
    builder.finalize_column_widths()


def _append_document_lines(document_element, key_column: str, key_value: str, sheet: SheetWriter, on_row=None) -> None:
    """Flatten each direct <Line> child of a document element (Invoice,
    StockMovement, WorkDocument, Payment) into its own Lines sheet row,
    carrying the parent document's own business key as a plain column.
    `on_row`, if given, is also called with each written row (used to feed
    the integrity check's running sums).
    """
    for child in document_element:
        if local_name(child.tag) == "Line":
            row = flatten_record(child)
            row[key_column] = key_value
            sheet.append(row)
            if on_row is not None:
                on_row(row)


def _append_movement_lines(transaction_element, transaction_id: str, sheet: SheetWriter, on_row=None) -> None:
    for child in transaction_element:
        if local_name(child.tag) != "Lines":
            continue
        for line in child:
            if local_name(line.tag) in ("DebitLine", "CreditLine"):
                row = flatten_record(line)
                row["TransactionID"] = transaction_id
                sheet.append(row)
                if on_row is not None:
                    on_row(row)


def _to_float(text: str) -> float:
    try:
        return float(text) if text else 0.0
    except ValueError:
        return 0.0


# A document's own DocumentStatus/*Status field holding "A" (Anulado --
# cancelled) is what each SourceDocuments wrapper's declared TotalDebit/
# TotalCredit excludes, confirmed against a real file (GFC 2024: excluding
# the 8 cancelled invoices' lines from the raw sum lands exactly on the
# file's own declared SalesInvoices.TotalCredit). NumberOfEntries still
# counts every document, cancelled or not -- only the money sums exclude
# them. Applied the same way to WorkingDocuments/Payments by the same
# "A" convention, though only the SalesInvoices case has a confirmed
# real-file example.
_STATUS_FIELD = {
    "SalesInvoices": "DocumentStatus_InvoiceStatus",
    "WorkingDocuments": "DocumentStatus_WorkStatus",
    "Payments": "DocumentStatus_PaymentStatus",
}
_CANCELLED_STATUS = "A"


class IntegrityCheck:
    """Accumulates row counts and debit/credit sums while records are being
    written, so they can be cross-checked afterwards against each section's
    own declared totals (the Totals sheet, see totals_row()) -- the same
    kind of check an auditor already runs by hand on a SAF-T's declared
    totals, done here for free as a sanity check on this workbook's own
    output. A mismatch can mean a bug in this script as easily as a genuine
    inconsistency in the source file; it's reported, not assumed either way.
    """

    _SECTIONS = ("GeneralLedgerEntries", "SalesInvoices", "WorkingDocuments", "Payments")

    def __init__(self):
        self.entries = {section: 0 for section in self._SECTIONS}
        self.debit = {section: 0.0 for section in self._SECTIONS}
        self.credit = {section: 0.0 for section in self._SECTIONS}
        self.movement_lines = 0
        self.movement_quantity = 0.0

    def record(self, section: str) -> None:
        self.entries[section] += 1

    def line(self, section: str, row: dict) -> None:
        self.debit[section] += _to_float(row.get("DebitAmount", ""))
        self.credit[section] += _to_float(row.get("CreditAmount", ""))

    def movement_line(self, row: dict) -> None:
        self.movement_lines += 1
        self.movement_quantity += _to_float(row.get("Quantity", ""))


def _line_accumulator(integrity: IntegrityCheck, section: str, document_row: dict):
    """Build the on_row callback for a document's lines: skip a cancelled
    document's lines (see _STATUS_FIELD/_CANCELLED_STATUS), since the file's
    own declared totals for `section` don't count them either.
    """
    status_field = _STATUS_FIELD[section]
    if document_row.get(status_field) == _CANCELLED_STATUS:
        return None
    return lambda row: integrity.line(section, row)


def _report_integrity_check(declared_totals: dict[str, dict], integrity: IntegrityCheck) -> None:
    tolerance = 0.01  # cents
    for section in IntegrityCheck._SECTIONS:
        declared = declared_totals.get(section)
        if declared is None:
            continue
        mismatches = []
        declared_count = declared.get("NumberOfEntries", "")
        if declared_count and int(_to_float(declared_count)) != integrity.entries[section]:
            mismatches.append(f"NumberOfEntries: declared {declared_count}, workbook has {integrity.entries[section]}")
        for field, computed in (("TotalDebit", integrity.debit[section]), ("TotalCredit", integrity.credit[section])):
            declared_value = declared.get(field, "")
            if abs(_to_float(declared_value) - computed) > tolerance:
                mismatches.append(f"{field}: declared {declared_value}, workbook sums to {computed:.2f}")
        if mismatches:
            print(f"WARNING: {section} totals don't match this workbook's own rows:", file=sys.stderr)
            for mismatch in mismatches:
                print(f"    {mismatch}", file=sys.stderr)

    movement = declared_totals.get("MovementOfGoods")
    if movement is not None:
        declared_lines = movement.get("NumberOfMovementLines", "")
        if declared_lines and int(_to_float(declared_lines)) != integrity.movement_lines:
            print(
                f"WARNING: MovementOfGoods.NumberOfMovementLines declared {declared_lines}, "
                f"workbook has {integrity.movement_lines} StockMovementsLines rows",
                file=sys.stderr,
            )
        declared_quantity = movement.get("TotalQuantityIssued", "")
        if abs(_to_float(declared_quantity) - integrity.movement_quantity) > tolerance:
            print(
                f"WARNING: MovementOfGoods.TotalQuantityIssued declared {declared_quantity}, "
                f"workbook's StockMovementsLines Quantity sums to {integrity.movement_quantity:.2f} "
                "-- 'Issued' may only count outbound movements, so this isn't necessarily a bug here.",
                file=sys.stderr,
            )


def _sheet(workbook: xlsxwriter.Workbook, title: str, headers: list[str], formats: Formats) -> SheetWriter:
    return SheetWriter(workbook, title, headers, formats, tab_color=SHEET_TAB_COLORS[title])


def convert(saft_path: str, output_path: str) -> None:
    workbook = xlsxwriter.Workbook(output_path, {"constant_memory": True})
    formats = Formats(workbook)

    # The workbook legend, sitting even before Header -- it explains what
    # every other sheet is before the reader gets to them. Static content
    # (see ESTRUTURA_ROWS), so unlike Header/Totals/Journals below it's
    # filled in immediately rather than deferred to the end of convert().
    header_color = SHEET_TAB_COLORS["Header"]
    estrutura_sheet = workbook.add_worksheet("Estrutura")
    estrutura_sheet.set_tab_color(header_color)
    populate_estrutura_sheet(estrutura_sheet, formats)

    # Created now, before the main parsing loop, so they sit first in the
    # workbook -- but not populated until the very end, once the loop has
    # gathered their data (header_row/declared_totals/journals). xlsxwriter
    # worksheets can be written to in any order relative to each other, so
    # this needs no equivalent of openpyxl's move_sheet() workaround.
    header_builder = _KeyValueSheetBuilder(workbook, "Header", formats, header_color)
    totals_builder = _KeyValueSheetBuilder(workbook, "Totals", formats, header_color)
    journals_builder = _KeyValueSheetBuilder(workbook, "Journals", formats, header_color)

    sheets = {
        "GeneralLedgerAccounts": _sheet(workbook, "GeneralLedgerAccounts", HEADERS_GENERAL_LEDGER_ACCOUNTS, formats),
        "Customers": _sheet(workbook, "Customers", HEADERS_CUSTOMERS, formats),
        "Suppliers": _sheet(workbook, "Suppliers", HEADERS_SUPPLIERS, formats),
        "Products": _sheet(workbook, "Products", HEADERS_PRODUCTS, formats),
        "TaxTable": _sheet(workbook, "TaxTable", HEADERS_TAX_TABLE, formats),
        "Transactions": _sheet(workbook, "Transactions", HEADERS_TRANSACTIONS, formats),
        "MovementLines": _sheet(workbook, "MovementLines", HEADERS_MOVEMENT_LINES, formats),
        "Invoices": _sheet(workbook, "Invoices", HEADERS_INVOICES, formats),
        "InvoiceLines": _sheet(workbook, "InvoiceLines", HEADERS_INVOICE_LINES, formats),
        "StockMovements": _sheet(workbook, "StockMovements", HEADERS_STOCK_MOVEMENTS, formats),
        "StockMovementsLines": _sheet(workbook, "StockMovementsLines", HEADERS_STOCK_MOVEMENTS_LINES, formats),
        "WorkDocuments": _sheet(workbook, "WorkDocuments", HEADERS_WORK_DOCUMENTS, formats),
        "WorkDocumentsLines": _sheet(workbook, "WorkDocumentsLines", HEADERS_WORK_DOCUMENTS_LINES, formats),
        "Payments": _sheet(workbook, "Payments", HEADERS_PAYMENTS, formats),
        "PaymentLines": _sheet(workbook, "PaymentLines", HEADERS_PAYMENT_LINES, formats),
    }

    tags_of_interest = tuple(
        f"{{*}}{tag}"
        for tag in (
            "Header",
            *TOTALS_WRAPPER_TAGS,
            "Account", "Customer", "Supplier", "Product", "TaxTableEntry",
            "Journal", "Transaction", "Invoice", "StockMovement", "WorkDocument", "Payment",
        )
    )
    integrity = IntegrityCheck()
    declared_totals: dict[str, dict] = {}
    header_row: dict[str, str] = {}
    journals: list[dict] = []

    context = etree.iterparse(saft_path, events=("end",), tag=tags_of_interest, recover=True)
    for _event, elem in context:
        tag = local_name(elem.tag)

        # A handful of these tag names also occur nested somewhere else in
        # the schema with a different meaning -- e.g. an Invoice's own
        # DocumentTotals/Payment block, which is not a Payments-sheet
        # record. iterparse's tag filter matches by local name anywhere in
        # the document, so a same-named nested element fires this loop just
        # like a real top-level record. Checking the immediate parent tag
        # tells the two apart; a false match is left untouched (not
        # dispatched, not cleared) so it's read correctly later as part of
        # its real ancestor's own flatten_record call.
        expected_parent = _EXPECTED_PARENT_TAG.get(tag)
        if expected_parent is not None:
            parent = elem.getparent()
            if parent is None or local_name(parent.tag) != expected_parent:
                continue

        if tag == "Header":
            header_row = flatten_record(elem)
        elif tag == "Journal":
            journals.append({
                "JournalID": _child_text(elem, "JournalID"),
                "Description": _child_text(elem, "Description"),
            })
        elif tag == "Account":
            sheets["GeneralLedgerAccounts"].append(flatten_record(elem))
        elif tag == "Customer":
            sheets["Customers"].append(flatten_record(elem))
        elif tag == "Supplier":
            sheets["Suppliers"].append(flatten_record(elem))
        elif tag == "Product":
            sheets["Products"].append(flatten_record(elem))
        elif tag == "TaxTableEntry":
            sheets["TaxTable"].append(flatten_record(elem))
        elif tag == "Transaction":
            row = flatten_record(elem, skip_tags=frozenset({"Lines"}))
            row["JournalID"] = _child_text(elem.getparent(), "JournalID")
            sheets["Transactions"].append(row)
            integrity.record("GeneralLedgerEntries")
            _append_movement_lines(
                elem, row.get("TransactionID", ""), sheets["MovementLines"],
                on_row=lambda r: integrity.line("GeneralLedgerEntries", r),
            )
        elif tag == "Invoice":
            row = flatten_record(elem, skip_tags=frozenset({"Line"}))
            sheets["Invoices"].append(row)
            integrity.record("SalesInvoices")
            _append_document_lines(
                elem, "InvoiceNo", row.get("InvoiceNo", ""), sheets["InvoiceLines"],
                on_row=_line_accumulator(integrity, "SalesInvoices", row),
            )
        elif tag == "StockMovement":
            row = flatten_record(elem, skip_tags=frozenset({"Line"}))
            sheets["StockMovements"].append(row)
            _append_document_lines(
                elem, "DocumentNumber", row.get("DocumentNumber", ""), sheets["StockMovementsLines"],
                on_row=integrity.movement_line,
            )
        elif tag == "WorkDocument":
            row = flatten_record(elem, skip_tags=frozenset({"Line"}))
            sheets["WorkDocuments"].append(row)
            integrity.record("WorkingDocuments")
            _append_document_lines(
                elem, "DocumentNumber", row.get("DocumentNumber", ""), sheets["WorkDocumentsLines"],
                on_row=_line_accumulator(integrity, "WorkingDocuments", row),
            )
        elif tag == "Payment":
            row = flatten_record(elem, skip_tags=frozenset({"Line"}))
            sheets["Payments"].append(row)
            integrity.record("Payments")
            _append_document_lines(
                elem, "PaymentRefNo", row.get("PaymentRefNo", ""), sheets["PaymentLines"],
                on_row=_line_accumulator(integrity, "Payments", row),
            )
        elif tag in TOTALS_WRAPPER_TAGS:
            section = TOTALS_WRAPPER_TAGS[tag]
            declared_totals[section] = totals_row(elem, section)

        # Clear only this element's own subtree (never a sibling or an
        # ancestor), so context still needed later -- a Journal's JournalID
        # for the next Transaction, a wrapper's summary fields once all its
        # records are done -- stays intact. This bounds memory to roughly
        # one record's own subtree rather than the whole file, which is
        # what matters for the ~1GB files this needs to handle, without the
        # sibling-deletion trick that would risk deleting that context.
        elem.clear()

    _report_integrity_check(declared_totals, integrity)

    for sheet in sheets.values():
        sheet.finalize_column_widths()

    # Populated last, now that the main loop has gathered their data --
    # see the comment where these builders were created, above.
    populate_header_sheet(header_builder, header_row)
    populate_totals_sheet(totals_builder, declared_totals)
    populate_journals_sheet(journals_builder, journals)

    workbook.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convert a Portuguese SAF-T (PT) XML file into an Excel workbook, one sheet per section."
    )
    parser.add_argument("saft_path", help="Path to the SAF-T XML file")
    parser.add_argument(
        "output_path",
        nargs="?",
        help="Path to the output .xlsx file (default: same name as the input, .xlsx extension)",
    )
    args = parser.parse_args(argv)

    output_path = args.output_path or str(Path(args.saft_path).with_suffix(".xlsx"))
    convert(args.saft_path, output_path)
    print(f"Wrote {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
