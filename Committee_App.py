"""Touch-friendly committee collection app for the supplied Excel workbook.

Install:
    py -m pip install openpyxl

Run from the Excel Folder:
    py committee_saving_app.py

The app manages 30 monthly installments starting July 2026. Collections are
written to an "App Collections" worksheet; the original workbook sheets,y

formulas, and Excel dropdown validations are preserved.
"""

from __future__ import annotations

import importlib
import os
import re
import tempfile
import tkinter as tk
import warnings
import webbrowser
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from urllib.parse import quote
from xml.etree import ElementTree as ET


WORKBOOK_NAME = "BCCommittee saving system.xlsx"
LEDGER_SHEET = "App Collections"
START_MONTH = date(2026, 7, 1)
SCHEDULE_LENGTH = 30

BACKGROUND = "#F8FAFC"
CARD = "#FFFFFF"
CARD_BORDER = "#E2E8F0"
TEXT = "#172033"
MUTED = "#53657D"
GREEN = "#10B981"
GREEN_LIGHT = "#DCFCE7"
RED = "#EF4444"
RED_LIGHT = "#FEE2E2"
AMBER = "#B45309"
AMBER_LIGHT = "#FEF3C7"
BLUE = "#2563EB"
WHATSAPP_BLUE = "#2563EB"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

LEDGER_HEADERS = (
    "Date",
    "Member ID",
    "Member Name",
    "Due Month",
    "Amount Collected",
    "Note",
)


@dataclass
class Member:
    member_id: str
    name: str
    amount: Decimal
    phone: str
    latest_phone_cell: str | None = None


@dataclass
class Collection:
    date_recorded: str
    member_id: str
    member_name: str
    due_month: str
    amount: Decimal
    note: str = "Collection"


class CommitteeApp:
    def __init__(self, root: tk.Tk, openpyxl_module: object) -> None:
        self.root = root
        self.openpyxl = openpyxl_module
        self.root.title("Committee Savings")
        self.root.configure(bg=BACKGROUND)
        self.root.minsize(800, 650)
        try:
            self.root.state("zoomed")
        except tk.TclError:
            self.root.geometry("1180x820")

        self.workbook_path = self._find_workbook()
        self.months = self._make_months()
        self.month_labels = [month.strftime("%B %Y") for month in self.months]
        self.members: list[Member] = []
        self.member_by_id: dict[str, Member] = {}
        self.collections: list[Collection] = []
        self._load_workbook_data()
        self._build_ui()
        self.refresh()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    @staticmethod
    def _find_workbook() -> Path:
        expected = Path(__file__).resolve().parent / WORKBOOK_NAME
        if expected.exists():
            return expected
        selected = filedialog.askopenfilename(
            title=f"Select {WORKBOOK_NAME}",
            initialdir=str(expected.parent),
            filetypes=[("Excel workbooks", "*.xlsx")],
        )
        if not selected:
            raise FileNotFoundError(
                f"Could not find {WORKBOOK_NAME} beside this application."
            )
        return Path(selected)

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
            CommitteeApp._normalize_header(sheet.cell(row, column).value): column
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
                marker = (
                    self._cell_text(sheet.cell(row, paid_col))
                    if paid_col
                    else ""
                )
                status = str(
                    sheet.cell(row, status_col).value if status_col else ""
                ).strip().lower()

                if sheet.title.strip().lower() == "latest":
                    # The user selected the cached Paid? result for November.
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
            raise ValueError(
                f'The "{LEDGER_SHEET}" sheet exists but its headings are not '
                "recognized. Please do not rename the ledger columns."
            )
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
        sheet_ids = [
            int(sheet.attrib["sheetId"])
            for sheet in workbook_xml.findall(f".//{{{MAIN_NS}}}sheet")
        ]
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
        ET.SubElement(
            worksheet,
            f"{{{MAIN_NS}}}dimension",
            {"ref": f"A1:F{max(last_row, 1)}"},
        )
        views = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetViews")
        ET.SubElement(views, f"{{{MAIN_NS}}}sheetView", {"workbookViewId": "0"})
        ET.SubElement(
            worksheet, f"{{{MAIN_NS}}}sheetFormatPr", {"defaultRowHeight": "18"}
        )
        data = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetData")

        def append_row(row_number: int, values: tuple[object, ...]) -> None:
            row = ET.SubElement(data, f"{{{MAIN_NS}}}row", {"r": str(row_number)})
            for column, value in enumerate(values, 1):
                reference = CommitteeApp._excel_column(column) + str(row_number)
                if isinstance(value, Decimal):
                    cell = ET.SubElement(row, f"{{{MAIN_NS}}}c", {"r": reference})
                    ET.SubElement(cell, f"{{{MAIN_NS}}}v").text = str(value)
                else:
                    cell = ET.SubElement(
                        row,
                        f"{{{MAIN_NS}}}c",
                        {"r": reference, "t": "inlineStr"},
                    )
                    inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
                    ET.SubElement(inline, f"{{{MAIN_NS}}}t").text = str(value)

        append_row(1, LEDGER_HEADERS)
        for row_number, entry in enumerate(collections, 2):
            append_row(
                row_number,
                (
                    entry.date_recorded,
                    entry.member_id,
                    entry.member_name,
                    entry.due_month,
                    entry.amount,
                    entry.note,
                ),
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
        sheet = (
            f'<sheet name="{LEDGER_SHEET}" sheetId="{sheet_id}" '
            f'r:id="rId{relationship_id}"/>'
        ).encode()
        workbook_xml, count = re.subn(
            rb"</sheets>", sheet + b"</sheets>", workbook_xml, count=1
        )
        if count != 1:
            raise ValueError("Could not register the App Collections worksheet.")

        relationship = (
            '<Relationship '
            f'Id="rId{relationship_id}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="{archive_sheet_path.removeprefix("xl/")}"/>'
        ).encode()
        relationships_xml, count = re.subn(
            rb"</Relationships>",
            relationship + b"</Relationships>",
            relationships_xml,
            count=1,
        )
        if count != 1:
            raise ValueError("Could not register the App Collections relationship.")

        override = (
            '<Override '
            f'PartName="/{archive_sheet_path}" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        ).encode()
        content_types_xml, count = re.subn(
            rb"</Types>", override + b"</Types>", content_types_xml, count=1
        )
        if count != 1:
            raise ValueError("Could not register the App Collections worksheet type.")
        return workbook_xml, relationships_xml, content_types_xml

    @staticmethod
    def _replace_cell_xml(
        worksheet_xml: bytes, coordinate: str, value: str
    ) -> bytes:
        def make_cell(tag: bytes) -> bytes:
            body = tag[:-2] if tag.endswith(b"/>") else tag[:-1]
            body = re.sub(rb'\s+t="[^"]*"', b"", body)
            safe_value = (
                value.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .encode("utf-8")
            )
            return (
                body
                + b' t="inlineStr"><is><t>'
                + safe_value
                + b"</t></is></c>"
            )

        ref = re.compile(rb'\br="' + re.escape(coordinate.encode("ascii")) + rb'"')
        cell_pattern = re.compile(rb"<c\b[^>]*>")
        for match in cell_pattern.finditer(worksheet_xml):
            if not ref.search(match.group(0)):
                continue
            start = match.start()
            opening_end = match.end()
            if match.group(0).endswith(b"/>"):
                end = opening_end
                body = match.group(0)[:-2]
            else:
                close = worksheet_xml.find(b"</c>", opening_end)
                if close < 0:
                    raise ValueError(f"Malformed Excel cell {coordinate}.")
                end = close + 4
            replacement = make_cell(match.group(0))
            return worksheet_xml[:start] + replacement + worksheet_xml[end:]

        reference = re.fullmatch(r"([A-Z]+)(\d+)", coordinate)
        if reference is None:
            raise ValueError(f"Invalid Excel cell reference: {coordinate}")
        column_number = CommitteeApp._column_index(reference.group(1))
        target_row = reference.group(2).encode("ascii")
        row_pattern = re.compile(rb"<row\b[^>]*>")
        row_match = next(
            (
                match
                for match in row_pattern.finditer(worksheet_xml)
                if re.search(rb'\br="' + target_row + rb'"', match.group(0))
            ),
            None,
        )
        if row_match is None:
            raise ValueError(f"Could not find Excel row {target_row.decode()}.")
        row_end = worksheet_xml.find(b"</row>", row_match.end())
        if row_end < 0:
            raise ValueError(f"Malformed Excel row {target_row.decode()}.")
        row_cells = worksheet_xml[row_match.end():row_end]
        insertion = row_end
        for match in cell_pattern.finditer(row_cells):
            found = re.search(rb'\br="([A-Z]+)\d+"', match.group(0))
            if found and CommitteeApp._column_index(
                found.group(1).decode("ascii")
            ) > column_number:
                insertion = row_match.end() + match.start()
                break
        new_cell = make_cell(
            b'<c r="' + coordinate.encode("ascii") + b'">'
        )
        return worksheet_xml[:insertion] + new_cell + worksheet_xml[insertion:]

    @staticmethod
    def _column_index(letters: str) -> int:
        number = 0
        for letter in letters:
            number = number * 26 + ord(letter) - ord("A") + 1
        return number

    def _save_data(
        self,
        collections: list[Collection],
        create_ledger: bool = False,
        phone_update: tuple[str, str, str] | None = None,
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
                existing_sheets = {
                    sheet.attrib.get("name")
                    for sheet in workbook_xml.findall(f".//{{{MAIN_NS}}}sheet")
                }
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
                        elif phone_update and entry.filename == phone_update[0]:
                            contents = self._replace_cell_xml(
                                contents, phone_update[1], phone_update[2]
                            )
                        destination.writestr(entry, contents)
                    if new_metadata:
                        destination.writestr(ledger_path, ledger_xml)

            with zipfile.ZipFile(temporary_path, "r") as verify:
                bad_entry = verify.testzip()
                if bad_entry:
                    raise OSError(f"Workbook validation failed at {bad_entry}.")
            os.replace(temporary_path, self.workbook_path)
        except Exception as exc:
            if temporary_path and os.path.exists(temporary_path):
                try:
                    os.remove(temporary_path)
                except OSError:
                    pass
            raise OSError(
                "Could not save the Excel workbook. Close it in Excel and try again."
            ) from exc

    def _selected_index(self) -> int:
        return self.month_labels.index(self.month_var.get())

    def _collected_by_month(self, member: Member) -> dict[str, Decimal]:
        totals: dict[str, Decimal] = {}
        for entry in self.collections:
            if entry.member_id == member.member_id:
                totals[entry.due_month] = totals.get(
                    entry.due_month, Decimal("0")
                ) + entry.amount
        return totals

    def _member_dues(
        self, member: Member, through_index: int | None = None
    ) -> tuple[int, Decimal, Decimal]:
        if through_index is None:
            through_index = self._selected_index()
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

    def _month_counts(self) -> tuple[int, int, int, Decimal]:
        paid = partial = unpaid = 0
        collected_total = Decimal("0")
        label = self.month_labels[self._selected_index()]
        for member in self.members:
            collected = self._collected_by_month(member).get(label, Decimal("0"))
            collected_total += collected
            if collected >= member.amount:
                paid += 1
            elif collected > 0:
                partial += 1
            else:
                unpaid += 1
        return paid, partial, unpaid, collected_total

    

    def _total_collection_through_selected_month(self) -> Decimal:
        selected_month = self.months[self._selected_index()]
    
        total = Decimal("0")
    
        for entry in self.collections:
            try:
                recorded_date = datetime.strptime(
                    entry.date_recorded,
                    "%Y-%m-%d %H:%M:%S"
                )
            except ValueError:
                # Skip imported historical records because
                # they do not contain an actual collection date.
                continue
    
            if (
                recorded_date.year == selected_month.year
                and recorded_date.month == selected_month.month
            ):
                total += entry.amount
    
        return total



    def _build_ui(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background=BACKGROUND)
        style.configure("Card.TFrame", background=CARD)
        style.configure(
            "TLabel", background=BACKGROUND, foreground=TEXT, font=("Segoe UI", 12)
        )
        style.configure(
            "Card.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 12)
        )
        style.configure(
            "Title.TLabel", background=BACKGROUND, foreground=TEXT,
            font=("Segoe UI", 28, "bold"),
        )

        outer = ttk.Frame(self.root, padding=(22, 18))
        outer.pack(fill="both", expand=True)
        title_bar = ttk.Frame(outer)
        title_bar.pack(fill="x", pady=(0, 12))
        ttk.Label(title_bar, text="Committee Savings", style="Title.TLabel").pack(
            side="left"
        )
        tk.Button(
            title_bar,
            text="FULL SCREEN",
            command=self._toggle_fullscreen,
            font=("Segoe UI", 11, "bold"),
            bg=CARD,
            fg=BLUE,
            relief="flat",
            padx=18,
            pady=12,
        ).pack(side="right")

        self.summary_frame = ttk.Frame(outer)
        self.summary_frame.pack(fill="x", pady=(0, 12))
        self.summary_values: dict[str, ttk.Label] = {}
        self.summary_details: dict[str, ttk.Label] = {}
        self.summary_cards: list[tk.Frame] = []
        self.summary_columns = 0
        summary_cards = (
            ("members", "TOTAL MEMBERS", BLUE),
            ("collection", "TOTAL COLLECTION", GREEN),
            ("paid", "PAID THIS MONTH", "#059669"),
            ("unpaid", "TOTAL OUTSTANDING", RED),
        )
        for key, label, accent in summary_cards:
            card = tk.Frame(
                self.summary_frame,
                bg=CARD,
                highlightbackground=CARD_BORDER,
                highlightthickness=1,
                bd=0,
            )
            card.grid(
                row=0,
                column=len(self.summary_cards),
                sticky="nsew",
                padx=(0, 9),
                pady=(0, 8),
            )
            self.summary_cards.append(card)
            tk.Frame(card, bg=accent, height=5).pack(fill="x")
            body = ttk.Frame(card, style="Card.TFrame", padding=(15, 11))
            body.pack(fill="both", expand=True)
            ttk.Label(
                body, text=label, style="Card.TLabel", foreground=MUTED,
                font=("Segoe UI", 11, "bold"),
            ).pack(anchor="w")
            value = ttk.Label(
                body, text="0", style="Card.TLabel", foreground=accent,
                font=("Segoe UI", 25, "bold"),
            )
            value.pack(anchor="w", pady=(3, 0))
            self.summary_values[key] = value
            detail = ttk.Label(
                body, text="", style="Card.TLabel",
                foreground=MUTED, font=("Segoe UI", 11),
            )
            detail.pack(anchor="w", pady=(2, 0))
            self.summary_details[key] = detail
        self.summary_frame.bind("<Configure>", self._layout_summary_cards)

        toolbar = ttk.Frame(outer)
        toolbar.pack(fill="x", pady=(0, 12))
        ttk.Label(
            toolbar, text="MONTH", font=("Segoe UI", 11, "bold")
        ).pack(side="left", padx=(0, 8))
        latest_month = self._default_month()
        self.month_var = tk.StringVar(value=latest_month)
        month_picker = ttk.Combobox(
            toolbar,
            textvariable=self.month_var,
            values=self.month_labels,
            state="readonly",
            font=("Segoe UI", 13, "bold"),
            width=18,
        )
        month_picker.pack(side="left", ipady=7, padx=(0, 16))
        month_picker.bind("<<ComboboxSelected>>", lambda _event: self.refresh())

        ttk.Label(
            toolbar, text="FIND MEMBER", font=("Segoe UI", 11, "bold")
        ).pack(side="left", padx=(0, 8))
        self.search_var = tk.StringVar()
        search = ttk.Entry(toolbar, textvariable=self.search_var, font=("Segoe UI", 14))
        search.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 12))
        search.bind("<KeyRelease>", lambda _event: self._render_members())
        self.filter_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            toolbar,
            text="OWES ONLY",
            variable=self.filter_var,
            command=self._render_members,
            font=("Segoe UI", 11, "bold"),
            bg=BACKGROUND,
            activebackground=BACKGROUND,
            padx=8,
        ).pack(side="right")

        list_card = ttk.Frame(outer, style="Card.TFrame", padding=(10, 8))
        list_card.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(
            list_card, bg=CARD, highlightthickness=0, borderwidth=0
        )
        scrollbar = ttk.Scrollbar(
            list_card, orient="vertical", command=self.canvas.yview
        )
        self.rows_frame = ttk.Frame(self.canvas, style="Card.TFrame")
        self.rows_window = self.canvas.create_window(
            (0, 0), window=self.rows_frame, anchor="nw"
        )
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.rows_frame.bind("<Configure>", self._update_scroll_region)
        self.canvas.bind(
            "<Configure>",
            lambda event: self.canvas.itemconfigure(
                self.rows_window, width=event.width
            ),
        )
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<ButtonPress-1>", self._start_touch_scroll)
        self.canvas.bind("<B1-Motion>", self._drag_touch_scroll)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.footer = ttk.Label(
            outer,
            text=(
                "30 monthly installments: July 2026 – December 2028. "
                "Collections are allocated to the oldest unpaid month first."
            ),
            font=("Segoe UI", 10),
            foreground=MUTED,
        )
        self.footer.pack(anchor="w", pady=(10, 0))

    def _layout_summary_cards(self, event: tk.Event) -> None:
        columns = 2 if event.width < 1050 else 4
        if columns == self.summary_columns:
            return
        self.summary_columns = columns
        for column in range(4):
            self.summary_frame.columnconfigure(
                column, weight=1 if column < columns else 0
            )
        for index, card in enumerate(self.summary_cards):
            card.grid_configure(
                row=index // columns,
                column=index % columns,
                sticky="nsew",
            )

    def _default_month(self) -> str:
        now = date.today()
        today_index = max(
            0,
            min(
                (now.year - START_MONTH.year) * 12 + now.month - START_MONTH.month,
                len(self.month_labels) - 1,
            ),
        )
        recorded_index = (
            self.month_labels.index(self.recorded_month_label)
            if self.recorded_month_label in self.month_labels
            else 0
        )
        return self.month_labels[max(today_index, recorded_index)]

    def _toggle_fullscreen(self) -> None:
        current = bool(self.root.attributes("-fullscreen"))
        self.root.attributes("-fullscreen", not current)

    def _update_scroll_region(self, _event: tk.Event) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_mousewheel(self, event: tk.Event) -> None:
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _start_touch_scroll(self, event: tk.Event) -> None:
        canvas_y = event.y_root - self.canvas.winfo_rooty()
        self.canvas.scan_mark(0, canvas_y)

    def _drag_touch_scroll(self, event: tk.Event) -> None:
        canvas_y = event.y_root - self.canvas.winfo_rooty()
        self.canvas.scan_dragto(0, canvas_y, gain=1)

    @staticmethod
    def _format_amount(amount: Decimal) -> str:
        if amount == amount.to_integral_value():
            return f"Rs {amount:,.0f}"
        return f"Rs {amount:,.2f}"

    def refresh(self) -> None:
        owing_members = 0
        outstanding_installments = 0
        outstanding_total = Decimal("0")
        for member in self.members:
            count, balance, _selected = self._member_dues(member)
            outstanding_installments += count
            owing_members += balance > 0
            outstanding_total += balance
        paid, partial, unpaid, month_collected = self._month_counts()

        self.summary_values["members"].configure(text=str(len(self.members)))
        self.summary_details["members"].configure(text="Active committee members")
        self.summary_values["collection"].configure(
            text=self._format_amount(self._total_collection_through_selected_month())
        )
        self.summary_details["collection"].configure(
    text=f"Collected in {self.month_var.get()}"
)
        
        self.summary_values["paid"].configure(text=str(paid))
        self.summary_details["paid"].configure(
            text=f"Paid in {self.month_var.get()}: {self._format_amount(month_collected)}"
            f"  ·  {partial} partial"
        )
        self.summary_values["unpaid"].configure(
            text=self._format_amount(outstanding_total)
        )
        self.summary_details["unpaid"].configure(
            text=f"{owing_members} members · {outstanding_installments} installments"
            f"  ·  {unpaid} unpaid this month"
        )
        self.footer.configure(
            text=(
                f"Showing {self.month_var.get()} · {paid} paid, {partial} partial, "
                f"{unpaid} unpaid this month. "
                "Use the green/red status button to mark PAID or reverse to UNPAID."
            )
        )
        self._render_members()

    def _render_members(self) -> None:
        for child in self.rows_frame.winfo_children():
            child.destroy()

        query = self.search_var.get().strip().casefold()
        only_owing = self.filter_var.get()
        month_label = self.month_labels[self._selected_index()]
        shown = 0
        for member in self.members:
            dues_count, outstanding, month_balance = self._member_dues(member)
            collected_this_month = member.amount - month_balance
            if query and query not in f"{member.name} {member.member_id}".casefold():
                continue
            if only_owing and outstanding <= 0:
                continue

            card = tk.Frame(
                self.rows_frame,
                bg=CARD,
                highlightbackground=CARD_BORDER,
                highlightthickness=1,
                bd=0,
            )
            card.pack(fill="x", pady=(3, 6), padx=3)
            card.columnconfigure(0, weight=1)

            content = ttk.Frame(card, style="Card.TFrame", padding=(16, 14))
            content.grid(row=0, column=0, sticky="ew")
            content.columnconfigure(0, weight=1)
            top = tk.Frame(content, bg=CARD)
            top.grid(row=0, column=0, sticky="ew")
            name_label = tk.Label(
                top,
                text=member.name,
                bg=CARD,
                fg=TEXT,
                anchor="w",
                font=("Segoe UI", 18, "bold"),
            )
            name_label.pack(side="left")

            if month_balance <= 0:
                status, status_bg, status_fg = "PAID", GREEN_LIGHT, GREEN
            elif collected_this_month > 0:
                status, status_bg, status_fg = "PART PAID", AMBER_LIGHT, AMBER
            else:
                status, status_bg, status_fg = "UNPAID", RED_LIGHT, RED
            tk.Label(
                top, text=status, bg=status_bg, fg=status_fg,
                font=("Segoe UI", 12, "bold"), padx=15, pady=8,
            ).pack(side="right")

            member_id_label = tk.Label(
                content,
                text=f"MEMBER ID  {member.member_id}",
                bg=CARD,
                fg=MUTED,
                anchor="w",
                font=("Segoe UI", 11, "bold"),
            )
            member_id_label.grid(row=1, column=0, sticky="w", pady=(2, 0))
            detail = (
                f"{month_label}: {self._format_amount(member.amount)} due"
                f"     ·     {dues_count} installment(s) outstanding"
                f"     ·     Total owed: {self._format_amount(outstanding)}"
            )
            detail_label = tk.Label(
                content,
                text=detail,
                bg=CARD,
                fg=MUTED,
                anchor="w",
                justify="left",
                font=("Segoe UI", 12),
            )
            detail_label.grid(row=2, column=0, sticky="w", pady=(8, 12))

            actions = tk.Frame(content, bg=CARD)
            actions.grid(row=3, column=0, sticky="ew")
            tk.Button(
                actions,
                text="MARK UNPAID" if collected_this_month > 0 else "MARK PAID",
                command=lambda item=member: self.toggle_payment(item),
                font=("Segoe UI", 13, "bold"),
                bg=RED if collected_this_month > 0 else GREEN,
                fg="white",
                activebackground="#DC2626" if collected_this_month > 0 else "#059669",
                activeforeground="white",
                relief="flat",
                padx=14,
                pady=14,
            ).pack(side="left", padx=(0, 10))
            tk.Button(
                actions,
                text="COLLECT PART",
                command=lambda item=member: self.collect_payment(item),
                font=("Segoe UI", 12, "bold"),
                bg="#E7EEF8",
                fg=TEXT,
                activebackground="#D6E2F2",
                relief="flat",
                padx=14,
                pady=14,
                state="normal" if outstanding > 0 else "disabled",
            ).pack(side="left", padx=(0, 10))
            tk.Button(
                actions,
                text="WHATSAPP REMINDER",
                command=lambda item=member: self.send_reminder(item),
                font=("Segoe UI", 12, "bold"),
                bg=WHATSAPP_BLUE,
                fg="white",
                activebackground="#1D4ED8",
                activeforeground="white",
                relief="flat",
                padx=14,
                pady=14,
                state="normal" if outstanding > 0 else "disabled",
            ).pack(side="left")

            for widget in (
                card, content, top, name_label, member_id_label, detail_label,
            ):
                widget.bind("<MouseWheel>", self._on_mousewheel)
                widget.bind("<ButtonPress-1>", self._start_touch_scroll)
                widget.bind("<B1-Motion>", self._drag_touch_scroll)
            ttk.Separator(self.rows_frame, orient="horizontal").pack(
                fill="x", padx=8
            )
            shown += 1

        if shown == 0:
            ttk.Label(
                self.rows_frame,
                text="No members match these filters.",
                style="Card.TLabel",
                foreground=MUTED,
                font=("Segoe UI", 14),
            ).pack(pady=32)

    def _ask_amount(self, member: Member, due: Decimal) -> Decimal | None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Record a collection")
        dialog.configure(bg=CARD)
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("440x620")

        ttk.Label(
            dialog,
            text=f"Collect from {member.name}",
            style="Title.TLabel",
            font=("Segoe UI", 20, "bold"),
            background=CARD,
        ).pack(pady=(18, 5))
        ttk.Label(
            dialog,
            text=f"Outstanding through {self.month_var.get()}: "
            f"{self._format_amount(due)}",
            style="Card.TLabel",
            font=("Segoe UI", 12),
        ).pack(pady=(0, 10))
        value = tk.StringVar()
        entry = ttk.Entry(
            dialog, textvariable=value, justify="right", font=("Segoe UI", 24)
        )
        entry.pack(fill="x", padx=18, ipady=10, pady=(0, 12))
        entry.focus_set()

        keypad = ttk.Frame(dialog, padding=(14, 4))
        keypad.pack(fill="both", expand=True)
        keys = (
            ("7", "8", "9", "⌫"),
            ("4", "5", "6", "C"),
            ("1", "2", "3", "."),
            ("0", "00", "CANCEL", "SAVE"),
        )
        result: list[Decimal | None] = [None]

        def press(key: str) -> None:
            current = value.get()
            if key == "⌫":
                value.set(current[:-1])
            elif key == "C":
                value.set("")
            elif key == "CANCEL":
                dialog.destroy()
            elif key == "SAVE":
                try:
                    amount = Decimal(current.replace(",", "").strip())
                except InvalidOperation:
                    messagebox.showerror(
                        "Enter an amount", "Please enter a valid collection amount.",
                        parent=dialog,
                    )
                    return
                if not amount.is_finite() or amount <= 0:
                    messagebox.showerror(
                        "Invalid amount", "The collection must be greater than zero.",
                        parent=dialog,
                    )
                    return
                if amount > due:
                    messagebox.showerror(
                        "Amount exceeds dues",
                        f"Enter no more than {self._format_amount(due)}. "
                        "This prevents recording an overpayment.",
                        parent=dialog,
                    )
                    return
                result[0] = amount
                dialog.destroy()
            elif key == ".":
                if "." not in current:
                    value.set(current + ".")
            else:
                value.set(current + key)

        for row_index, keys_row in enumerate(keys):
            keypad.rowconfigure(row_index, weight=1)
            for column_index, key in enumerate(keys_row):
                keypad.columnconfigure(column_index, weight=1)
                is_action = key in {"CANCEL", "SAVE"}
                tk.Button(
                    keypad,
                    text=key,
                    command=lambda item=key: press(item),
                    font=("Segoe UI", 16, "bold"),
                    bg=BLUE if key == "SAVE" else ("#E7EEF8" if is_action else CARD),
                    fg="white" if key == "SAVE" else TEXT,
                    relief="flat",
                    padx=5,
                    pady=12,
                ).grid(
                    row=row_index, column=column_index, sticky="nsew",
                    padx=4, pady=4,
                )
        dialog.wait_window()
        return result[0]

    def toggle_payment(self, member: Member) -> None:
        month = self.month_labels[self._selected_index()]
        collected = self._collected_by_month(member).get(month, Decimal("0"))
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if collected > 0:
            amount = -collected
            note = "Payment reversed to UNPAID"
        else:
            amount = member.amount - collected
            note = "Marked PAID"

        updated = self.collections + [
            Collection(
                timestamp,
                member.member_id,
                member.name,
                month,
                amount,
                note,
            )
        ]
        try:
            self._save_data(updated)
        except OSError as exc:
            messagebox.showerror("Could not save status", str(exc), parent=self.root)
            return
        self.collections = updated
        self.refresh()

    def collect_payment(self, member: Member) -> None:
        _count, due, _month_balance = self._member_dues(member)
        if due <= 0:
            return
        amount = self._ask_amount(member, due)
        if amount is None:
            return

        remaining = amount
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_entries: list[Collection] = []
        totals = self._collected_by_month(member)
        for month in self.month_labels[: self._selected_index() + 1]:
            balance = max(Decimal("0"), member.amount - totals.get(month, Decimal("0")))
            if balance <= 0:
                continue
            applied = min(balance, remaining)
            new_entries.append(
                Collection(
                    timestamp,
                    member.member_id,
                    member.name,
                    month,
                    applied,
                    "App collection",
                )
            )
            remaining -= applied
            if remaining <= 0:
                break

        if not new_entries:
            return
        updated = self.collections + new_entries
        try:
            self._save_data(updated)
        except OSError as exc:
            messagebox.showerror("Could not save collection", str(exc), parent=self.root)
            return
        self.collections = updated
        self.refresh()
        messagebox.showinfo(
            "Collection saved",
            f"Saved {self._format_amount(amount)} for {member.name}.\n"
            f"Applied across {len(new_entries)} installment(s), oldest first.",
            parent=self.root,
        )

    @staticmethod
    def _normalize_phone(phone: str) -> str | None:
        value = phone.strip()
        if value.startswith("00"):
            value = value[2:]
        digits = re.sub(r"\D", "", value)
        if not 8 <= len(digits) <= 15 or digits.startswith("0"):
            return None
        return digits

    def send_reminder(self, member: Member) -> None:
        _count, outstanding, _balance = self._member_dues(member)
        if outstanding <= 0:
            return
        phone = self._normalize_phone(member.phone)
        phone_update = None
        entered_phone = None
        if phone is None:
            entered = simpledialog.askstring(
                "WhatsApp number",
                f"Enter {member.name}'s number with country code "
                "(example: 923001234567):",
                parent=self.root,
            )
            if entered is None:
                return
            phone = self._normalize_phone(entered)
            if phone is None:
                messagebox.showerror(
                    "Invalid number",
                    "Enter 8–15 digits including the country code, without a "
                    "leading local 0.",
                    parent=self.root,
                )
                return
            entered_phone = entered.strip()
            if member.latest_phone_cell:
                try:
                    latest_path = self._sheet_path_from_file("Latest")
                    phone_update = (latest_path, member.latest_phone_cell, entered_phone)
                except (OSError, ValueError, KeyError) as exc:
                    messagebox.showerror(
                        "Could not save phone number", str(exc), parent=self.root
                    )
                    return

        due_count, _amount, _selected = self._member_dues(member)
        message = (
            f"Hello {member.name}, this is a polite reminder that "
            f"{due_count} committee installment(s) are outstanding through "
            f"{self.month_var.get()}, totaling {self._format_amount(outstanding)}. "
            "Please pay at your earliest convenience. Thank you!"
        )
        url = f"https://wa.me/{phone}?text={quote(message)}"
        try:
            if phone_update:
                self._save_data(
                    self.collections,
                    phone_update=phone_update,
                )
            if entered_phone is not None:
                member.phone = entered_phone
            import webbrowser
            opened = webbrowser.open_new_tab(url)
        except (OSError, webbrowser.Error) as exc:
            messagebox.showerror(
                "Could not open WhatsApp",
                f"Could not open the WhatsApp reminder in your browser: {exc}",
                parent=self.root,
            )
            return
        if not opened:
            messagebox.showerror(
                "Could not open WhatsApp",
                "The browser did not open the reminder link. Check that a default "
                "browser is configured.",
                parent=self.root,
            )
            return
        messagebox.showinfo(
            "WhatsApp reminder ready",
            "WhatsApp opened with the reminder filled in. Review the message and "
            "tap Send.",
            parent=self.root,
        )

    def _sheet_path_from_file(self, sheet_name: str) -> str:
        with zipfile.ZipFile(self.workbook_path, "r") as archive:
            return self._sheet_path(archive, sheet_name)

    def _on_close(self) -> None:
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    try:
        openpyxl = importlib.import_module("openpyxl")
    except ModuleNotFoundError:
        messagebox.showerror(
            "Install openpyxl",
            "This app needs openpyxl to read the Excel workbook.\n\n"
            "Open Command Prompt in the Excel Folder and run:\n"
            "py -m pip install openpyxl",
            parent=root,
        )
        root.destroy()
        return
    try:
        app = CommitteeApp(root, openpyxl)
    except (OSError, ValueError, ImportError, zipfile.BadZipFile) as exc:
        messagebox.showerror("Could not start Committee Savings", str(exc), parent=root)
        root.destroy()
        return
    root.deiconify()
    root.mainloop()


if __name__ == "__main__":
    main()
