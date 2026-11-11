from __future__ import annotations

import importlib
import os
import re
import tempfile
import warnings
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from html import escape
from pathlib import Path
from urllib.parse import quote
from xml.etree import ElementTree as ET
import zipfile

import streamlit as st

WORKBOOK_NAME = "BCCommittee saving system.xlsx"
LEDGER_SHEET = "App Collections"
START_MONTH = date(2026, 7, 1)
SCHEDULE_LENGTH = 30

LEDGER_HEADERS = (
    "Date",
    "Member ID",
    "Member Name",
    "Due Month",
    "Amount Collected",
    "Note",
)

# Page configuration
st.set_page_config(
    page_title="Committee Savings",
    page_icon="💰",
    layout="wide",
)

st.markdown(
    """
    <style>
    :root {
        color-scheme: light;
        --page: #F8FAFC;
        --card: #FFFFFF;
        --border: #E2E8F0;
        --text: #172033;
        --muted: #53657D;
        --green: #10B981;
        --red: #EF4444;
        --blue: #2563EB;
    }

    .stApp,
    [data-testid="stAppViewContainer"] {
        background: var(--page);
        color: var(--text);
    }

    .block-container {
        max-width: 1200px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    h1, h2, h3, p, label {
        color: var(--text);
    }

    [data-testid="stVerticalBlockBorderWrapper"] {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        padding: 1rem 1.1rem;
        margin-bottom: 0.8rem;
    }

    .metric-card {
        min-height: 132px;
        padding: 1.1rem 1.2rem;
        background: var(--card);
        border: 1px solid var(--border);
        border-top: 5px solid var(--accent);
        border-radius: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }

    .metric-label {
        color: var(--muted);
        font-size: 0.9rem;
        font-weight: 700;
        letter-spacing: 0.04em;
        text-transform: uppercase;
    }

    .metric-value {
        margin-top: 0.35rem;
        color: var(--text);
        font-size: clamp(1.45rem, 2.4vw, 2rem);
        font-weight: 800;
        line-height: 1.2;
        overflow-wrap: anywhere;
    }

    .metric-detail {
        margin-top: 0.35rem;
        color: var(--muted);
        font-size: 0.9rem;
    }

    .member-id {
        margin-top: -0.5rem;
        margin-bottom: 0.55rem;
        color: var(--muted);
        font-size: 0.95rem;
        font-weight: 700;
    }

    .status-badge {
        display: inline-block;
        margin: 0.35rem 0;
        padding: 0.45rem 0.9rem;
        border-radius: 999px;
        font-size: 0.9rem;
        font-weight: 800;
        letter-spacing: 0.03em;
    }

    .status-paid { background: #DCFCE7; color: #15803D; }
    .status-partial { background: #FEF3C7; color: #B45309; }
    .status-unpaid { background: #FEE2E2; color: #B91C1C; }

    div[data-testid="stButton"] > button,
    div[data-testid="stLinkButton"] > a {
        min-height: 3.15rem;
        border-radius: 12px;
        font-size: 1rem;
        font-weight: 750;
        transition: filter 120ms ease, transform 120ms ease;
    }

    div[data-testid="stButton"] > button[kind="primary"] {
        background: var(--green);
        border-color: var(--green);
        color: #FFFFFF;
    }

    div[data-testid="stButton"] > button[kind="secondary"] {
        background: #FFFFFF;
        border-color: var(--border);
        color: var(--text);
    }

    div[class*="st-key-member-unpaid-action-"] button {
        background: var(--red) !important;
        border-color: var(--red) !important;
        color: #FFFFFF !important;
    }

    div[class*="st-key-whatsapp-action-"] a {
        background: #25D366 !important;
        border-color: #25D366 !important;
        color: #FFFFFF !important;
    }

    div[data-testid="stButton"] > button:hover,
    div[data-testid="stLinkButton"] > a:hover {
        filter: brightness(0.94);
        transform: translateY(-1px);
    }

    @media (max-width: 640px) {
        .block-container {
            padding: 1rem 0.8rem 2rem;
        }

        [data-testid="stVerticalBlockBorderWrapper"] {
            padding: 0.75rem;
        }

        .metric-card {
            min-height: 112px;
            padding: 0.85rem;
        }

        .metric-label, .metric-detail {
            font-size: 0.78rem;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


class Member:
    def __init__(self, member_id: str, name: str, amount: Decimal, phone: str, latest_phone_cell: str | None = None):
        self.member_id = member_id
        self.name = name
        self.amount = amount
        self.phone = phone
        self.latest_phone_cell = latest_phone_cell


class Collection:
    def __init__(self, date_recorded: str, member_id: str, member_name: str, due_month: str, amount: Decimal, note: str = "Collection"):
        self.date_recorded = date_recorded
        self.member_id = member_id
        self.member_name = member_name
        self.due_month = due_month
        self.amount = amount
        self.note = note


@st.cache_resource
def load_openpyxl():
    return importlib.import_module("openpyxl")


class CommitteeAppManager:
    def __init__(self, workbook_path: Path, openpyxl_module: object) -> None:
        self.workbook_path = workbook_path
        self.openpyxl = openpyxl_module
        self.months = self._make_months()
        self.month_labels = [month.strftime("%B %Y") for month in self.months]
        self.members: list[Member] = []
        self.member_by_id: dict[str, Member] = {}
        self.collections: list[Collection] = []
        self.recorded_month_label: str | None = None
        self._load_workbook_data()

    @staticmethod
    def _make_months() -> list[date]:
        months = []
        year, month = START_MONTH.year, START_MONTH.month
        for _ in range(SCHEDULE_LENGTH):
            months.append(date(year, month, 1))
            year, month = (year + 1, 1) if month == 12 else (year, month + 1)
        return months

    @staticmethod
    def _normalize_header(value: object) -> str:
        return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()

    @staticmethod
    def _normalize_id(value: object) -> str:
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value or "").strip()

    @staticmethod
    def _as_amount(value: object, context: str) -> Decimal:
        try:
            amount = Decimal(str(value or 0))
        except (InvalidOperation, ValueError) as exc:
            raise ValueError(f"Invalid amount in {context}: {value!r}") from exc
        if not amount.is_finite():
            raise ValueError(f"Invalid amount in {context}: {value!r}")
        return amount

    @staticmethod
    def _cell_text(cell: object) -> str:
        value = cell.value
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    @staticmethod
    def _headers_on_sheet(sheet: object, row: int) -> dict[str, int]:
        return {
            CommitteeAppManager._normalize_header(sheet.cell(row, column).value): column
            for column in range(1, sheet.max_column + 1)
            if sheet.cell(row, column).value is not None
        }

    def _load_workbook_data(self) -> None:
        warning_text = "Data Validation extension is not supported and will be removed"
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore",
                message=warning_text,
                category=UserWarning,
                module="openpyxl.worksheet._reader",
            )
            workbook = self.openpyxl.load_workbook(
                self.workbook_path, data_only=True, read_only=False
            )

        try:
            if "Members" not in workbook.sheetnames:
                raise ValueError('The workbook must contain a "Members" sheet.')
            if "Latest" not in workbook.sheetnames:
                raise ValueError('The workbook must contain a "Latest" sheet.')

            self.recorded_month_label = self._month_for_sheet(workbook["Latest"])
            self._load_members(workbook["Members"], workbook["Latest"])
            if not self.members:
                raise ValueError('No active members were found on the "Members" sheet.')

            if LEDGER_SHEET in workbook.sheetnames:
                self.collections = self._read_ledger(workbook[LEDGER_SHEET])
            else:
                self.collections = self._import_existing_history(workbook)
                self._save_data(self.collections, create_ledger=True)
        finally:
            workbook.close()

    def _load_members(self, members_sheet: object, latest_sheet: object) -> None:
        member_headers = self._headers_on_sheet(members_sheet, 1)
        for required in ("member id", "member name", "monthly amount"):
            if required not in member_headers:
                raise ValueError(f'Missing "{required}" heading on the Members sheet.')
        id_col = member_headers["member id"]
        name_col = member_headers["member name"]
        amount_col = member_headers["monthly amount"]

        latest_header_row = None
        latest_headers: dict[str, int] = {}
        for row in range(1, min(latest_sheet.max_row, 12) + 1):
            found = self._headers_on_sheet(latest_sheet, row)
            if "member id" in found and "phone number" in found:
                latest_header_row = row
                latest_headers = found
                break

        phone_by_id: dict[str, tuple[str, str | None]] = {}
        if latest_header_row is not None:
            for row in range(latest_header_row + 1, latest_sheet.max_row + 1):
                member_id = self._normalize_id(latest_sheet.cell(row, latest_headers["member id"]).value)
                if not member_id:
                    continue
                phone_cell = latest_sheet.cell(row, latest_headers["phone number"])
                phone = self._cell_text(phone_cell)
                phone_by_id[member_id] = (
                    phone if phone != "0" else "",
                    phone_cell.coordinate,
                )

        seen_ids: set[str] = set()
        for row in range(2, members_sheet.max_row + 1):
            member_id = self._normalize_id(members_sheet.cell(row, id_col).value)
            name = self._cell_text(members_sheet.cell(row, name_col))
            if not member_id or not name or member_id in seen_ids:
                continue
            amount = self._as_amount(
                members_sheet.cell(row, amount_col).value,
                f"Members row {row}",
            )
            if amount <= 0:
                continue
            phone, phone_cell = phone_by_id.get(member_id, ("", None))
            member = Member(member_id, name, amount, phone, phone_cell)
            self.members.append(member)
            self.member_by_id[member_id] = member
            seen_ids.add(member_id)

    def _month_for_sheet(self, sheet: object) -> str | None:
        title = str(sheet.title).strip()
        candidate = self._normalize_header(sheet["B1"].value)
        if title.lower() != "latest":
            normalized_title = self._normalize_header(title)
            if normalized_title == "octuber":
                normalized_title = "october"
            if any(month.strftime("%B").lower() == normalized_title for month in self.months):
                candidate = normalized_title
        elif candidate:
            candidate = candidate.lower()

        if candidate == "octuber":
            candidate = "october"
        for month, label in zip(self.months, self.month_labels):
            if candidate in {month.strftime("%B").lower(), label.lower()}:
                return label
        return None

    def _import_existing_history(self, workbook: object) -> list[Collection]:
        imported: dict[tuple[str, str], Collection] = {}
        ignored_sheets = {"members", "dashboard", "payment", LEDGER_SHEET.lower()}
        for sheet in workbook.worksheets:
            if sheet.title.strip().lower() in ignored_sheets:
                continue
            due_month = self._month_for_sheet(sheet)
            if due_month is None:
                continue

            header_row = None
            headers: dict[str, int] = {}
            for row in range(1, min(sheet.max_row, 12) + 1):
                found = self._headers_on_sheet(sheet, row)
                if "member id" in found and "member name" in found:
                    header_row, headers = row, found
                    break
            if header_row is None:
                continue
            paid_col = headers.get("paid amount")
            status_col = headers.get("paid")
            amount_col = headers.get("amount")
            if amount_col is None:
                continue

            for row in range(header_row + 1, sheet.max_row + 1):
                member_id = self._normalize_id(sheet.cell(row, headers["member id"]).value)
                member = self.member_by_id.get(member_id)
                if member is None:
                    continue
                amount_due = self._as_amount(
                    sheet.cell(row, amount_col).value,
                    f"{sheet.title} row {row}",
                )
                paid_value = sheet.cell(row, paid_col).value if paid_col else None
                marker = self._cell_text(sheet.cell(row, paid_col)) if paid_col else ""
                status = str(sheet.cell(row, status_col).value if status_col else "").strip().lower()

                if sheet.title.strip().lower() == "latest":
                    collected = amount_due if status == "paid" else Decimal("0")
                elif marker.upper() == "Y" or status == "paid":
                    collected = amount_due
                elif paid_col and isinstance(paid_value, (int, float)):
                    collected = min(amount_due, Decimal(str(paid_value)))
                else:
                    collected = Decimal("0")

                if collected <= 0:
                    continue
                key = (member_id, due_month)
                imported[key] = Collection(
                    date_recorded=f"{due_month} (imported)",
                    member_id=member.member_id,
                    member_name=member.name,
                    due_month=due_month,
                    amount=collected,
                    note=f"Imported from {sheet.title}",
                )
        return list(imported.values())

    def _read_ledger(self, sheet: object) -> list[Collection]:
        headers = self._headers_on_sheet(sheet, 1)
        required = {self._normalize_header(header) for header in LEDGER_HEADERS}
        if not required.issubset(headers):
            raise ValueError(f'The "{LEDGER_SHEET}" sheet exists but headings are not recognized.')
        result = []
        for row in range(2, sheet.max_row + 1):
            member_id = self._normalize_id(sheet.cell(row, headers["member id"]).value)
            due_month = self._cell_text(sheet.cell(row, headers["due month"]))
            if not member_id or not due_month:
                continue
            member = self.member_by_id.get(member_id)
            name = self._cell_text(sheet.cell(row, headers["member name"]))
            amount = self._as_amount(
                sheet.cell(row, headers["amount collected"]).value,
                f"{LEDGER_SHEET} row {row}",
            )
            if amount == 0:
                continue
            recorded = self._cell_text(sheet.cell(row, headers["date"]))
            note = self._cell_text(sheet.cell(row, headers["note"]))
            result.append(
                Collection(
                    recorded,
                    member_id,
                    member.name if member else name,
                    due_month,
                    amount,
                    note or "Collection",
                )
            )
        return result

    def _sheet_path(self, archive: zipfile.ZipFile, sheet_name: str) -> str:
        workbook_xml = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relationships.findall(f"{{{PKG_REL_NS}}}Relationship")
        }
        for sheet in workbook_xml.findall(f".//{{{MAIN_NS}}}sheet"):
            if sheet.attrib.get("name") != sheet_name:
                continue
            target = targets[sheet.attrib[f"{{{REL_NS}}}id"]]
            return target.lstrip("/") if target.startswith("/") else f"xl/{target}"
        raise ValueError(f'Could not locate the "{sheet_name}" worksheet in the file.')

    @staticmethod
    def _new_ledger_path(archive: zipfile.ZipFile) -> tuple[str, int, int]:
        names = archive.namelist()
        sheet_numbers = [
            int(match.group(1))
            for name in names
            if (match := re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", name))
        ]
        next_sheet_number = max(sheet_numbers, default=0) + 1
        workbook_xml = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet_ids = [int(sheet.attrib["sheetId"]) for sheet in workbook_xml.findall(f".//{{{MAIN_NS}}}sheet")]
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationship_ids = [
            int(match.group(1))
            for relation in relationships.findall(f"{{{PKG_REL_NS}}}Relationship")
            if (match := re.fullmatch(r"rId(\d+)", relation.attrib.get("Id", "")))
        ]
        return (
            f"xl/worksheets/sheet{next_sheet_number}.xml",
            max(sheet_ids, default=0) + 1,
            max(relationship_ids, default=0) + 1,
        )

    @staticmethod
    def _make_ledger_xml(collections: list[Collection]) -> bytes:
        ET.register_namespace("", MAIN_NS)
        worksheet = ET.Element(f"{{{MAIN_NS}}}worksheet")
        last_row = len(collections) + 1
        ET.SubElement(worksheet, f"{{{MAIN_NS}}}dimension", {"ref": f"A1:F{max(last_row, 1)}"})
        views = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetViews")
        ET.SubElement(views, f"{{{MAIN_NS}}}sheetView", {"workbookViewId": "0"})
        ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetFormatPr", {"defaultRowHeight": "18"})
        data = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetData")

        def append_row(row_number: int, values: tuple[object, ...]) -> None:
            row = ET.SubElement(data, f"{{{MAIN_NS}}}row", {"r": str(row_number)})
            for column, value in enumerate(values, 1):
                reference = CommitteeAppManager._excel_column(column) + str(row_number)
                if isinstance(value, Decimal):
                    cell = ET.SubElement(row, f"{{{MAIN_NS}}}c", {"r": reference})
                    ET.SubElement(cell, f"{{{MAIN_NS}}}v").text = str(value)
                else:
                    cell = ET.SubElement(row, f"{{{MAIN_NS}}}c", {"r": reference, "t": "inlineStr"})
                    inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
                    ET.SubElement(inline, f"{{{MAIN_NS}}}t").text = str(value)

        append_row(1, LEDGER_HEADERS)
        for row_number, entry in enumerate(collections, 2):
            append_row(
                row_number,
                (entry.date_recorded, entry.member_id, entry.member_name, entry.due_month, entry.amount, entry.note),
            )
        return ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)

    @staticmethod
    def _excel_column(number: int) -> str:
        name = ""
        while number:
            number, remainder = divmod(number - 1, 26)
            name = chr(65 + remainder) + name
        return name

    @staticmethod
    def _add_ledger_metadata(
        workbook_xml: bytes,
        relationships_xml: bytes,
        content_types_xml: bytes,
        archive_sheet_path: str,
        sheet_id: int,
        relationship_id: int,
    ) -> tuple[bytes, bytes, bytes]:
        sheet = (f'<sheet name="{LEDGER_SHEET}" sheetId="{sheet_id}" r:id="rId{relationship_id}"/>').encode()
        workbook_xml, count = re.subn(rb"</sheets>", sheet + b"</sheets>", workbook_xml, count=1)
        
        relationship = (
            f'<Relationship Id="rId{relationship_id}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="{archive_sheet_path.removeprefix("xl/")}"/>'
        ).encode()
        relationships_xml, count = re.subn(rb"</Relationships>", relationship + b"</Relationships>", relationships_xml, count=1)

        override = (
            f'<Override PartName="/{archive_sheet_path}" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        ).encode()
        content_types_xml, count = re.subn(rb"</Types>", override + b"</Types>", content_types_xml, count=1)
        return workbook_xml, relationships_xml, content_types_xml

    def _save_data(
        self,
        collections: list[Collection],
        create_ledger: bool = False,
    ) -> None:
        temporary_path: str | None = None
        try:
            with tempfile.NamedTemporaryFile(
                prefix="committee-save-",
                suffix=".xlsx",
                dir=str(self.workbook_path.parent),
                delete=False,
            ) as temporary_file:
                temporary_path = temporary_file.name

            with zipfile.ZipFile(self.workbook_path, "r") as source:
                workbook_xml = ET.fromstring(source.read("xl/workbook.xml"))
                existing_sheets = {sheet.attrib.get("name") for sheet in workbook_xml.findall(f".//{{{MAIN_NS}}}sheet")}
                if LEDGER_SHEET in existing_sheets:
                    ledger_path = self._sheet_path(source, LEDGER_SHEET)
                    new_metadata = None
                elif create_ledger:
                    ledger_path, sheet_id, relationship_id = self._new_ledger_path(source)
                    new_metadata = (sheet_id, relationship_id)
                else:
                    raise ValueError(f'The "{LEDGER_SHEET}" worksheet is missing.')

                metadata_changes = {}
                if new_metadata:
                    (
                        metadata_changes["xl/workbook.xml"],
                        metadata_changes["xl/_rels/workbook.xml.rels"],
                        metadata_changes["[Content_Types].xml"],
                    ) = self._add_ledger_metadata(
                        source.read("xl/workbook.xml"),
                        source.read("xl/_rels/workbook.xml.rels"),
                        source.read("[Content_Types].xml"),
                        ledger_path,
                        new_metadata[0],
                        new_metadata[1],
                    )

                with zipfile.ZipFile(temporary_path, "w") as destination:
                    destination.comment = source.comment
                    ledger_xml = self._make_ledger_xml(collections)
                    for entry in source.infolist():
                        contents = source.read(entry.filename)
                        if entry.filename == ledger_path:
                            contents = ledger_xml
                        elif entry.filename in metadata_changes:
                            contents = metadata_changes[entry.filename]
                        destination.writestr(entry, contents)
                    if new_metadata:
                        destination.writestr(ledger_path, ledger_xml)

            os.replace(temporary_path, self.workbook_path)
        except Exception as exc:
            if temporary_path and os.path.exists(temporary_path):
                try:
                    os.remove(temporary_path)
                except OSError:
                    pass
            raise OSError("Could not save the Excel workbook.") from exc

    def _collected_by_month(self, member: Member) -> dict[str, Decimal]:
        totals: dict[str, Decimal] = {}
        for entry in self.collections:
            if entry.member_id == member.member_id:
                totals[entry.due_month] = totals.get(entry.due_month, Decimal("0")) + entry.amount
        return totals

    def _member_dues(self, member: Member, through_index: int) -> tuple[int, Decimal, Decimal]:
        totals = self._collected_by_month(member)
        outstanding_count = 0
        outstanding_amount = Decimal("0")
        selected_balance = Decimal("0")
        for index, label in enumerate(self.month_labels[: through_index + 1]):
            balance = max(Decimal("0"), member.amount - totals.get(label, Decimal("0")))
            if balance > 0:
                outstanding_count += 1
                outstanding_amount += balance
            if index == through_index:
                selected_balance = balance
        return outstanding_count, outstanding_amount, selected_balance


@st.cache_resource
def get_manager():
    openpyxl = load_openpyxl()
    workbook_path = Path(__file__).resolve().parent / WORKBOOK_NAME
    if not workbook_path.exists():
        st.error(f"Workbook '{WORKBOOK_NAME}' not found beside app.py!")
        st.stop()
    return CommitteeAppManager(workbook_path, openpyxl)


def main():
    manager = get_manager()

    st.title("💰 Committee Savings Dashboard")
    st.markdown("Manage monthly committee collections easily with one click.")

    # Sidebar / Top Controls
    col_m, col_s, col_f = st.columns([2, 2, 1])
    with col_m:
        selected_month = st.selectbox("Select Month", manager.month_labels, index=0)
    with col_s:
        search_query = st.text_input("Find Member", "").strip().casefold()
    with col_f:
        st.markdown("<br>", unsafe_allow_html=True)
        owes_only = st.checkbox("Owes Only", value=False)

    selected_index = manager.month_labels.index(selected_month)

    # Calculate Summary Statistics
    total_members = len(manager.members)
    total_outstanding_all = Decimal("0")
    paid_count = 0
    partial_count = 0
    unpaid_count = 0
    month_collected = Decimal("0")

    for member in manager.members:
        _, outstanding, month_balance = manager._member_dues(member, selected_index)
        total_outstanding_all += outstanding
        collected_this_month = member.amount - month_balance
        month_collected += collected_this_month
        if month_balance <= 0:
            paid_count += 1
        elif collected_this_month > 0:
            partial_count += 1
        else:
            unpaid_count += 1

    # Total cumulative collection across all recorded entries
    total_collection_all = sum((entry.amount for entry in manager.collections if entry.amount > 0), Decimal("0"))

    # Display Top Metric Cards
    metrics = (
        ("Total Members", f"{total_members:,}", "Active members", "#2563EB"),
        ("Total Collection", f"Rs {total_collection_all:,.0f}", "All collections", "#10B981"),
        (f"Paid ({selected_month})", f"{paid_count:,}", f"Rs {month_collected:,.0f} collected", "#059669"),
        ("Total Outstanding", f"Rs {total_outstanding_all:,.0f}", f"{unpaid_count} unpaid this month", "#EF4444"),
    )
    metric_columns = st.columns(4)
    for column, (label, value, detail, accent) in zip(metric_columns, metrics):
        with column:
            st.markdown(
                f"""
                <div class="metric-card" style="--accent: {accent}">
                    <div class="metric-label">{escape(label)}</div>
                    <div class="metric-value">{escape(value)}</div>
                    <div class="metric-detail">{escape(detail)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.divider()

    # Display Member Cards
    for member_index, member in enumerate(manager.members):
        dues_count, outstanding, month_balance = manager._member_dues(member, selected_index)
        collected_this_month = member.amount - month_balance

        if search_query and search_query not in f"{member.name} {member.member_id}".casefold():
            continue
        if owes_only and outstanding <= 0:
            continue

        with st.container(border=True):
            col_info, col_status = st.columns([3, 1])
            with col_info:
                st.subheader(member.name)
                st.markdown(
                    f'<div class="member-id">MEMBER ID &nbsp; {escape(member.member_id)}</div>',
                    unsafe_allow_html=True,
                )
                st.caption(
                    f"{selected_month}: Rs {member.amount:,.0f} due  ·  "
                    f"{dues_count} installment(s) outstanding  ·  "
                    f"Total owed: Rs {outstanding:,.0f}"
                )
            with col_status:
                if month_balance <= 0:
                    badge_class, status = "status-paid", "PAID"
                elif collected_this_month > 0:
                    badge_class, status = "status-partial", "PART PAID"
                else:
                    badge_class, status = "status-unpaid", "UNPAID"
                st.markdown(
                    f'<span class="status-badge {badge_class}">{status}</span>',
                    unsafe_allow_html=True,
                )

            # Action Buttons
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                button_label = "Mark Unpaid" if collected_this_month > 0 else "Mark Paid"
                action_scope = (
                    f"member_unpaid_action_{member_index}"
                    if collected_this_month > 0
                    else f"member_paid_action_{member_index}"
                )
                with st.container(key=action_scope):
                    clicked = st.button(
                        button_label,
                        key=f"toggle_{member_index}_{selected_month}",
                        type="secondary" if collected_this_month > 0 else "primary",
                        use_container_width=True,
                    )
                if clicked:
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    amount = -collected_this_month if collected_this_month > 0 else member.amount - collected_this_month
                    note = "Payment reversed" if collected_this_month > 0 else "Marked Paid"
                    
                    manager.collections.append(Collection(timestamp, member.member_id, member.name, selected_month, amount, note))
                    try:
                        manager._save_data(manager.collections)
                    except OSError as exc:
                        manager.collections.pop()
                        st.error(f"Could not save payment status: {exc}")
                        continue
                    st.rerun()

            with btn_col2:
                phone_digits = re.sub(r"\D", "", member.phone) if member.phone else ""
                whatsapp_msg = (
                    f"Hello {member.name}, this is a polite reminder that "
                    f"{dues_count} committee installment(s) are outstanding through "
                    f"{selected_month}, totaling Rs {outstanding:,.0f}. "
                    "Please pay at your earliest convenience. Thank you!"
                )
                whatsapp_url = f"https://wa.me/{phone_digits}?text={quote(whatsapp_msg)}"
                with st.container(key=f"whatsapp_action_{member_index}"):
                    st.link_button(
                        "💬 WhatsApp",
                        whatsapp_url,
                        disabled=(phone_digits == ""),
                        use_container_width=True,
                    )


if __name__ == "__main__":
    main()