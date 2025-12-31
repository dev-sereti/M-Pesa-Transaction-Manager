from __future__ import annotations
from typing import Any, Dict, List, Union
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Union

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from scr.models.transaction import Transaction


Record = Union[Transaction, Dict[str, Any]]


class ExcelHandler:
    """
    Updates only one worksheet (default: 'Transaction') in an existing workbook.

    - Does NOT delete/overwrite other sheets.
    - Can append new transactions.
    - Can update existing rows matched by Transaction Code (upsert).
    """

    DEFAULT_HEADERS = [
        "Transaction Code",
        "Date",
        "Type",
        "Amount",
        "Fee",
        "Sender",
        "Recipient",
        "Balance",
    ]

    def __init__(self, excel_path: str, sheet_name: str = "Transaction"):
        self.excel_path = Path(excel_path)
        self.sheet_name = sheet_name

        # Create file if missing, and ensure the Transaction sheet exists
        self._ensure_workbook_exists()
        self._ensure_sheet_exists()

    # Public API
    def append_transactions(
        self,
        transactions: List[Transaction],
        update_existing: bool = False,
    ) -> bool:
    
        # Append transactions into the Transaction sheet.

        # If update_existing=True:
        # If Transaction Code already exists, update known columns in that row
        #     without touching any other columns (e.g., Category/Notes remain).
        # Otherwise append a new row.
        # If update_existing=False:
        #   - Skip duplicates (existing Transaction Code).
        try:
            wb = load_workbook(self.excel_path)
            ws = self._get_or_create_sheet(wb)

            header_map = self._get_or_create_headers(ws)
            code_col = header_map["Transaction Code"]

            code_to_row = self._build_code_index(ws, code_col)

            appended = 0
            updated = 0
            skipped = 0

            for tx in transactions:
                record = self._tx_to_row_values(tx)

                tx_code = str(record.get("Transaction Code") or "").strip()
                if not tx_code:
                    continue

                existing_row = code_to_row.get(tx_code)

                if existing_row is not None:
                    if update_existing:
                        self._write_record_to_row(ws, header_map, existing_row, record)
                        updated += 1
                    else:
                        skipped += 1
                    continue

                # Append new row
                new_row = ws.max_row + 1
                self._write_record_to_row(ws, header_map, new_row, record)

                code_to_row[tx_code] = new_row
                appended += 1

            self._auto_fit_columns(ws, header_map)
            wb.save(self.excel_path)

            return True
        except Exception as e:
            print(f"Error updating Excel workbook: {e}")
            return False

    def upsert_from_database(
        self,
        db_rows: List[Dict[str, Any]],
        update_existing: bool = True,
    ) -> bool:
        
        # Ensures all DB transactions exist in the Transaction sheet.
        # By default it updates existing rows and appends missing ones.

        # This does NOT clear the sheet (so it won't delete manual columns/data).
        try:
            wb = load_workbook(self.excel_path)
            ws = self._get_or_create_sheet(wb)

            header_map = self._get_or_create_headers(ws)
            code_col = header_map["Transaction Code"]
            code_to_row = self._build_code_index(ws, code_col)

            for row in db_rows:
                record = self._db_dict_to_row_values(row)
                tx_code = str(record.get("Transaction Code") or "").strip()
                if not tx_code:
                    continue

                existing_row = code_to_row.get(tx_code)
                if existing_row is not None:
                    if update_existing:
                        self._write_record_to_row(ws, header_map, existing_row, record)
                    continue

                new_row = ws.max_row + 1
                self._write_record_to_row(ws, header_map, new_row, record)
                code_to_row[tx_code] = new_row

            self._auto_fit_columns(ws, header_map)
            wb.save(self.excel_path)
            return True
        except Exception as e:
            print(f"Error upserting DB rows to Excel: {e}")
            return False

    # Backward compatible name for your existing GUI button
    def export_from_database(self, db_manager) -> bool:
    
        # Keeps your existing GUI flow:
        # - Pull all rows from DB
        # - Upsert into Transaction sheet
        # - Does not touch other sheets
    
        rows = db_manager.get_all_transactions()
        return self.upsert_from_database(rows, update_existing=True)

    # Workbook / Sheet helpers
  
    def _ensure_workbook_exists(self) -> None:
        if self.excel_path.exists():
            return

        wb = Workbook()
        # Remove default sheet if present and create the desired one
        default_name = wb.active.title
        ws = wb.active
        ws.title = self.sheet_name

        # Headers
        for col_idx, h in enumerate(self.DEFAULT_HEADERS, start=1):
            ws.cell(row=1, column=col_idx, value=h)

        self._format_header(ws, len(self.DEFAULT_HEADERS))
        self._auto_fit_columns(ws, {h: i for i, h in enumerate(self.DEFAULT_HEADERS, start=1)})

        wb.save(self.excel_path)

    def _ensure_sheet_exists(self) -> None:
        wb = load_workbook(self.excel_path)
        _ = self._get_or_create_sheet(wb)
        wb.save(self.excel_path)

    def _get_or_create_sheet(self, wb) -> Any:
        # Case-insensitive match
        for name in wb.sheetnames:
            if name.strip().lower() == self.sheet_name.strip().lower():
                return wb[name]
        return wb.create_sheet(self.sheet_name)

    def _get_or_create_headers(self, ws) -> Dict[str, int]:
        """
        Returns a mapping: header name -> column index.
        If sheet is empty or missing headers, create missing headers without removing others.
        """
        # If sheet looks empty, write headers
        if ws.max_row < 1 or all((ws.cell(1, c).value is None for c in range(1, ws.max_column + 1))):
            for col_idx, h in enumerate(self.DEFAULT_HEADERS, start=1):
                ws.cell(row=1, column=col_idx, value=h)
            self._format_header(ws, len(self.DEFAULT_HEADERS))
            return {h: i for i, h in enumerate(self.DEFAULT_HEADERS, start=1)}

        # Read existing headers
        existing_headers: Dict[str, int] = {}
        max_col = max(ws.max_column, len(self.DEFAULT_HEADERS))
        for col_idx in range(1, max_col + 1):
            v = ws.cell(1, col_idx).value
            if v is None:
                continue
            header = str(v).strip()
            if header:
                existing_headers[header] = col_idx

        # Add missing required headers to the end (do not delete/overwrite existing)
        next_col = ws.max_column + 1
        changed = False
        for h in self.DEFAULT_HEADERS:
            if h not in existing_headers:
                ws.cell(row=1, column=next_col, value=h)
                existing_headers[h] = next_col
                next_col += 1
                changed = True

        if changed:
            self._format_header(ws, ws.max_column)

        # Ensure required headers exist
        for h in self.DEFAULT_HEADERS:
            if h not in existing_headers:
                raise ValueError(f"Failed to create required header: {h}")

        return existing_headers

    def _format_header(self, ws, header_col_count: int) -> None:
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_align = Alignment(horizontal="center", vertical="center")

        for col_idx in range(1, header_col_count + 1):
            cell = ws.cell(row=1, column=col_idx)
            if cell.value is None:
                continue
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align

    def _build_code_index(self, ws, code_col: int) -> Dict[str, int]:
        """
        Builds an index of existing Transaction Code -> row number.
        """
        code_to_row: Dict[str, int] = {}
        if ws.max_row < 2:
            return code_to_row

        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, code_col).value
            if v is None:
                continue
            code = str(v).strip()
            if code and code not in code_to_row:
                code_to_row[code] = r
        return code_to_row

    # -----------------------------
    # Writing helpers
    # -----------------------------
    def _write_record_to_row(self, ws, header_map: Dict[str, int], row_idx: int, record: Dict[str, Any]) -> None:
        """
        Writes only known mapped columns. Does not touch other columns on that row.
        """
        for header, col_idx in header_map.items():
            if header not in record:
                continue

            value = record[header]
            cell = ws.cell(row=row_idx, column=col_idx, value=value)

            # Set some basic formats (won't affect other sheets)
            if header == "Date":
                # If you pass datetime, Excel will recognize it properly
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
            elif header in ("Amount", "Fee", "Balance"):
                cell.number_format = "#,##0.00"

    def _auto_fit_columns(self, ws, header_map: Dict[str, int]) -> None:
        """
        Best-effort width adjustment for the Transaction sheet only.
        """
        for header, col_idx in header_map.items():
            max_len = len(header)
            for r in range(2, min(ws.max_row, 5000) + 1):  # guard for very large sheets
                v = ws.cell(r, col_idx).value
                if v is None:
                    continue
                max_len = max(max_len, len(str(v)))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 50)

    # -----------------------------
    # Normalization helpers
    # -----------------------------
    def _tx_to_row_values(self, tx: Transaction) -> Dict[str, Any]:
        return {
            "Transaction Code": tx.transaction_code,
            "Date": tx.date,  # datetime object is ideal for Excel
            "Type": (tx.transaction_type or "").capitalize(),
            "Amount": float(tx.amount),
            "Fee": float(tx.fee),
            "Sender": tx.sender or "",
            "Recipient": tx.recipient or "",
            "Balance": float(tx.balance) if tx.balance is not None else "",
        }

    def _db_dict_to_row_values(self, row: Dict[str, Any]) -> Dict[str, Any]:
        # row['date'] is stored as string 'YYYY-mm-dd HH:MM:SS' in your DB manager
        raw_date = row.get("date")
        dt_value: Any = raw_date
        if isinstance(raw_date, str):
            try:
                dt_value = datetime.strptime(raw_date, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                dt_value = raw_date  # keep as text if parsing fails

        return {
            "Transaction Code": row.get("transaction_code", ""),
            "Date": dt_value,
            "Type": str(row.get("transaction_type", "")).capitalize(),
            "Amount": float(row.get("amount") or 0.0),
            "Fee": float(row.get("fee") or 0.0),
            "Sender": row.get("sender") or "",
            "Recipient": row.get("recipient") or "",
            "Balance": float(row.get("balance")) if row.get("balance") is not None else "",
        }