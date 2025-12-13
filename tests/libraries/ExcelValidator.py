"""
Excel Validator Library for Robot Framework
Validates Excel data structure and content for timeline generation.
"""
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


class ExcelValidator:
    """Validates Excel files for timeline data structure."""

    def __init__(self):
        self.required_columns = [
            "Task Name",
            "Start Date",
            "Type"
        ]
        self.optional_columns = ["End Date", "Color", "Swimlane"]
        self.valid_types = ["Milestone", "Feature", "Phase"]
        self.valid_colors = [
            "red", "blue", "green", "orange", "purple", "yellow", "gray", "grey"
        ]

    def validate_excel_file(
        self,
        file_path: str,
        sheet_name: str = "TimelineData"
    ) -> Tuple[bool, List[str]]:
        """
        Validate Excel file structure and content.

        Args:
            file_path: Path to the Excel file
            sheet_name: Name of the sheet to validate

        Returns:
            Tuple of (is_valid, list_of_errors)
        """
        errors = []

        # Check if file exists
        if not os.path.exists(file_path):
            errors.append(f"Excel file not found: {file_path}")
            return False, errors

        try:
            workbook = load_workbook(file_path, data_only=True)
        except Exception as e:
            errors.append(f"Failed to open Excel file: {str(e)}")
            return False, errors

        # Check if sheet exists
        if sheet_name not in workbook.sheetnames:
            errors.append(
                f"Sheet '{sheet_name}' not found. Available sheets: "
                f"{', '.join(workbook.sheetnames)}"
            )
            return False, errors

        sheet = workbook[sheet_name]

        # Validate headers
        header_errors = self._validate_headers(sheet)
        errors.extend(header_errors)

        # Validate data rows
        if not header_errors:  # Only validate data if headers are valid
            data_errors = self._validate_data_rows(sheet)
            errors.extend(data_errors)

        workbook.close()
        return len(errors) == 0, errors

    def _validate_headers(self, sheet) -> List[str]:
        """Validate that required headers exist."""
        errors = []
        headers = []

        # Read headers from row 1
        for cell in sheet[1]:
            if cell.value:
                headers.append(str(cell.value).strip())

        # Check for required columns
        missing_columns = []
        for required_col in self.required_columns:
            if required_col not in headers:
                missing_columns.append(required_col)

        if missing_columns:
            errors.append(
                f"Missing required columns: {', '.join(missing_columns)}"
            )

        return errors

    def _validate_data_rows(self, sheet) -> List[str]:
        """Validate data rows for correctness."""
        errors = []
        headers = [str(cell.value).strip() if cell.value else "" 
                   for cell in sheet[1]]

        # Get column indices
        col_indices = {}
        for idx, header in enumerate(headers, start=1):
            col_indices[header] = idx

        # Validate each data row
        for row_num in range(2, sheet.max_row + 1):
            row_errors = self._validate_row(
                sheet, row_num, col_indices, headers
            )
            if row_errors:
                errors.extend([f"Row {row_num}: {err}" for err in row_errors])

        return errors

    def _validate_row(
        self,
        sheet,
        row_num: int,
        col_indices: Dict[str, int],
        headers: List[str]
    ) -> List[str]:
        """Validate a single data row."""
        errors = []

        # Get cell values
        task_name = self._get_cell_value(sheet, row_num, col_indices, "Task Name")
        start_date = self._get_cell_value(sheet, row_num, col_indices, "Start Date")
        end_date = self._get_cell_value(sheet, row_num, col_indices, "End Date")
        event_type = self._get_cell_value(sheet, row_num, col_indices, "Type")
        color = self._get_cell_value(sheet, row_num, col_indices, "Color")
        swimlane = self._get_cell_value(sheet, row_num, col_indices, "Swimlane")

        # Validate required fields
        if not task_name:
            errors.append("Task Name is required")
        if not start_date:
            errors.append("Start Date is required")
        if not event_type:
            errors.append("Type is required")
        elif event_type.lower() not in [t.lower() for t in self.valid_types]:
            errors.append(
                f"Type must be one of: {', '.join(self.valid_types)}"
            )
        if color and color.lower() not in [c.lower() for c in self.valid_colors]:
            errors.append(
                f"Color must be one of: {', '.join(self.valid_colors)}"
            )

        # Validate date logic
        if start_date and end_date:
            try:
                start = self._parse_date(start_date)
                end = self._parse_date(end_date)
                if start and end and start > end:
                    errors.append("Start Date must be before End Date")
            except Exception:
                pass  # Date parsing errors handled elsewhere

        # Validate Feature and Phase require End Date
        if event_type:
            event_type_lower = event_type.lower()
            if event_type_lower in ["feature", "phase"] and not end_date:
                errors.append(f"{event_type} events require an End Date")

        return errors

    def _get_cell_value(
        self,
        sheet,
        row_num: int,
        col_indices: Dict[str, int],
        column_name: str
    ) -> Optional[str]:
        """Get cell value by column name."""
        if column_name not in col_indices:
            return None
        col_idx = col_indices[column_name]
        cell = sheet.cell(row=row_num, column=col_idx)
        return str(cell.value).strip() if cell.value else None

    def _parse_date(self, date_value: str) -> Optional[datetime]:
        """Parse date value from Excel cell."""
        from datetime import datetime
        if isinstance(date_value, datetime):
            return date_value
        # Try to parse string dates
        try:
            from dateutil import parser
            return parser.parse(str(date_value))
        except Exception:
            return None

    def get_data_summary(
        self,
        file_path: str,
        sheet_name: str = "TimelineData"
    ) -> Dict[str, any]:
        """
        Get summary statistics about the Excel data.

        Args:
            file_path: Path to the Excel file
            sheet_name: Name of the sheet

        Returns:
            Dictionary with summary statistics
        """
        summary = {
            "total_rows": 0,
            "milestones": 0,
            "features": 0,
            "phases": 0,
            "swimlanes": set(),
            "colors": set()
        }

        try:
            workbook = load_workbook(file_path, data_only=True)
            sheet = workbook[sheet_name]

            headers = [str(cell.value).strip() if cell.value else "" 
                      for cell in sheet[1]]
            col_indices = {header: idx for idx, header in enumerate(headers, start=1)}

            for row_num in range(2, sheet.max_row + 1):
                task_name = self._get_cell_value(sheet, row_num, col_indices, "Task Name")
                if task_name:
                    summary["total_rows"] += 1
                    event_type = self._get_cell_value(sheet, row_num, col_indices, "Type")
                    if event_type:
                        event_type_lower = event_type.lower()
                        if event_type_lower == "milestone":
                            summary["milestones"] += 1
                        elif event_type_lower == "feature":
                            summary["features"] += 1
                        elif event_type_lower == "phase":
                            summary["phases"] += 1

                    color = self._get_cell_value(sheet, row_num, col_indices, "Color")
                    if color:
                        summary["colors"].add(color.lower())

                    swimlane = self._get_cell_value(sheet, row_num, col_indices, "Swimlane")
                    if swimlane:
                        summary["swimlanes"].add(swimlane)

            workbook.close()

            # Convert sets to lists for JSON serialization
            summary["swimlanes"] = list(summary["swimlanes"])
            summary["colors"] = list(summary["colors"])

        except Exception as e:
            summary["error"] = str(e)

        return summary

    def read_excel_data(
        self,
        file_path: str,
        sheet_name: str = "TimelineData"
    ) -> List[Dict[str, any]]:
        """
        Read data from Excel file and return as list of dictionaries.

        Args:
            file_path: Path to the Excel file
            sheet_name: Name of the sheet to read

        Returns:
            List of dictionaries, each representing a row
        """
        data = []

        try:
            workbook = load_workbook(file_path, data_only=True)
            sheet = workbook[sheet_name]

            # Get headers from first row
            headers = [str(cell.value).strip() if cell.value else f"Column{idx}" 
                      for idx, cell in enumerate(sheet[1], start=1)]

            # Read data rows
            for row_num in range(2, sheet.max_row + 1):
                row_data = {}
                has_data = False
                for col_idx, header in enumerate(headers, start=1):
                    cell = sheet.cell(row=row_num, column=col_idx)
                    value = cell.value
                    if value is not None:
                        has_data = True
                        # Convert datetime to string for consistency
                        if isinstance(value, datetime):
                            value = value.strftime("%Y-%m-%d")
                        row_data[header] = str(value).strip() if value else ""
                    else:
                        row_data[header] = ""

                if has_data and row_data.get("Task Name", "").strip():
                    data.append(row_data)

            workbook.close()

        except Exception as e:
            raise RuntimeError(f"Error reading Excel file: {str(e)}")

        return data

    def get_sheet_names(self, file_path: str) -> List[str]:
        """
        Get list of sheet names in the Excel file.

        Args:
            file_path: Path to the Excel file

        Returns:
            List of sheet names
        """
        try:
            workbook = load_workbook(file_path, data_only=True)
            sheets = workbook.sheetnames
            workbook.close()
            return sheets
        except Exception as e:
            raise RuntimeError(f"Error reading Excel file: {str(e)}")

