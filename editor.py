import os
import pandas as pd

from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt6.QtWidgets import (
    QWidget, QTableView, QVBoxLayout, QHBoxLayout,
    QPushButton, QMessageBox, QFileDialog, QLabel
)

from constants import *

class PandasTableModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame):
        super().__init__()
        self._df = df

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.index)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None

        value = self._df.iat[index.row(), index.column()]

        # display/edit as text; keep underlying dtype handling simple
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole):
            if pd.isna(value):
                return ""
            return str(value)

        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return str(self._df.columns[section])
        return str(self._df.index[section])

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return (
            Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsEditable
        )

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if role != Qt.ItemDataRole.EditRole or not index.isValid():
            return False

        r, c = index.row(), index.column()
        col = self._df.columns[c]
        current = self._df.iat[r, c]

        text = "" if value is None else str(value).strip()

        # Write empty string as NA (common spreadsheet behavior)
        if text == "":
            new_val = pd.NA
        else:
            # Try to preserve dtype when possible
            try:
                if pd.isna(current):
                    # if current is NA, try infer by column dtype
                    dtype = self._df[col].dtype
                    if dtype.kind in ("i", "u"):
                        new_val = int(text)
                    elif dtype.kind == "f":
                        new_val = float(text)
                    elif dtype.kind == "b":
                        new_val = text.lower() in ("1", "true", "yes", "y")
                    else:
                        new_val = text
                else:
                    if isinstance(current, (int,)):
                        new_val = int(text)
                    elif isinstance(current, (float,)):
                        new_val = float(text)
                    elif isinstance(current, (bool,)):
                        new_val = text.lower() in ("1", "true", "yes", "y")
                    else:
                        new_val = text
            except Exception:
                # If conversion fails, fall back to string
                new_val = text

        self._df.iat[r, c] = new_val
        self.dataChanged.emit(index, index, [Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole])
        return True

    def insertRows(self, row: int, count: int, parent=QModelIndex()) -> bool:
        """
        Qt calls this when rows are inserted. We'll insert `count` blank rows
        starting at position `row`.
        """
        if count <= 0:
            return False

        row = max(0, min(row, self.rowCount()))
        self.beginInsertRows(QModelIndex(), row, row + count - 1)

        # Build `count` new rows with NA values for all columns
        new_rows = pd.DataFrame([{col: pd.NA for col in self._df.columns} for _ in range(count)])

        if self._df.empty:
            self._df = new_rows.copy()
        else:
            top = self._df.iloc[:row, :]
            bottom = self._df.iloc[row:, :]
            self._df = pd.concat([top, new_rows, bottom], ignore_index=True)

        self.endInsertRows()
        return True

    def add_row(self):
        """Convenience: append one blank row at the end."""
        self.insertRows(self.rowCount(), 1)

    def removeRows(self, row: int, count: int, parent=QModelIndex()) -> bool:
        """Optional but useful."""
        if count <= 0 or self.rowCount() == 0:
            return False

        row = max(0, min(row, self.rowCount() - 1))
        last = min(self.rowCount() - 1, row + count - 1)

        self.beginRemoveRows(QModelIndex(), row, last)
        drop_idx = list(range(row, last + 1))
        self._df = self._df.drop(self._df.index[drop_idx]).reset_index(drop=True)
        self.endRemoveRows()
        return True

    def df(self) -> pd.DataFrame:
        return self._df


class EditorWindow(QWidget):
    def __init__(self, excel_file: str):
        super().__init__()
        self.excel_file = excel_file
        self.df = None
        self.model = None

        self.init_ui()
        self.load_excel(excel_file)

    def init_ui(self):
        self.setWindowTitle(f"Editor - {os.path.basename(self.excel_file)}")
        self.setMinimumSize(900, 600)

        root = QVBoxLayout(self)

        top = QHBoxLayout()
        self.path_label = QLabel(self.excel_file)
        self.btn_reload = QPushButton("Reload")
        self.btn_save = QPushButton("Save")
        self.btn_save_as = QPushButton("Save As…")
        self.btn_add_row = QPushButton("Add Row")
        self.btn_rm_row = QPushButton("Remove Selected Row(s)")
        
        self.btn_reload.clicked.connect(self.on_reload)
        self.btn_save.clicked.connect(self.on_save)
        self.btn_save_as.clicked.connect(self.on_save_as)
        self.btn_add_row.clicked.connect(self.on_add_row)
        self.btn_rm_row.clicked.connect(self.on_rm_row)
        
        top.addWidget(self.path_label, stretch=1)
        top.addWidget(self.btn_reload)
        top.addWidget(self.btn_save)
        top.addWidget(self.btn_save_as)
        top.addWidget(self.btn_add_row)
        top.addWidget(self.btn_rm_row)
        root.addLayout(top)

        # Table
        self.table_view = QTableView(self)
        self.table_view.setEditTriggers(
            QTableView.EditTrigger.DoubleClicked
            | QTableView.EditTrigger.SelectedClicked
            | QTableView.EditTrigger.EditKeyPressed
        )
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table_view.setSelectionMode(QTableView.SelectionMode.ExtendedSelection)
        root.addWidget(self.table_view)

    def load_excel(self, path: str):
        try:
            df = pd.read_excel(path)

            self.df = df
            self.model = PandasTableModel(self.df)
            self.table_view.setModel(self.model)
            self.table_view.resizeColumnsToContents()

            self.excel_file = path
            self.path_label.setText(path)
            self.setWindowTitle(f"Editor - {os.path.basename(path)}")

        except Exception as e:
            QMessageBox.critical(self, "Load Failed", str(e))

    def on_reload(self):
        self.load_excel(self.excel_file)

    def on_save(self):
        if self.model is None:
            return
        try:
            # Write back exactly what the user sees
            self.model.df().to_excel(self.excel_file, index=False)
            QMessageBox.information(self, "Saved", f"Saved to:\n{self.excel_file}")
        except Exception as e:
            QMessageBox.critical(self, "Save Failed", str(e))

    def on_save_as(self):
        if self.model is None:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save As", self.excel_file, "Excel Files (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            self.model.df().to_excel(path, index=False)
            self.excel_file = path
            self.path_label.setText(path)
            self.setWindowTitle(f"Editor - {os.path.basename(path)}")
            QMessageBox.information(self, "Saved", f"Saved to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Save Failed", str(e))
    
    def on_add_row(self):
        if self.model is None:
            return

        # If a row is selected, insert below it; otherwise append to bottom
        sel = self.table_view.selectionModel().selectedRows()
        if sel:
            insert_at = sel[0].row() + 1
            self.model.insertRows(insert_at, 1)
            self.table_view.selectRow(insert_at)
        else:
            self.model.add_row()
            self.table_view.selectRow(self.model.rowCount() - 1)

        self.table_view.scrollToBottom()
    
    def on_rm_row(self):
        if self.model is None:
            return

        selection = self.table_view.selectionModel().selectedRows()
        if not selection:
            QMessageBox.information(self, "No Selection", "Select one or more rows to remove.")
            return

        rows = sorted((idx.row() for idx in selection), reverse=True)

        confirm = QMessageBox.question(
            self,
            "Remove Rows",
            f"Remove {len(rows)} selected row(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if confirm != QMessageBox.StandardButton.Yes:
            return

        for row in rows:
            self.model.removeRows(row, 1)


def launch_editor(excel_file: str):
    # Keep a reference so the window doesn't get GC'd.
    win = EditorWindow(excel_file)
    win.show()
    return win
    
