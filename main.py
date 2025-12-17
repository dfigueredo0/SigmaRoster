import sys
import os
import subprocess
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QLabel, QFileDialog, QLineEdit, QMessageBox
)
from PyQt6.QtCore import Qt

from constants import *
from minutes import *
from roster import create_roster
from editor import launch_editor

class ExcelDropLineEdit(QLineEdit):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setPlaceholderText("Drag & drop an Excel file here or click to browse...")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            path = event.mimeData().urls()[0].toLocalFile()
            if path.endswith('.xlsx'):
                event.acceptProposedAction()

    def dropEvent(self, event):
        path = event.mimeData().urls()[0].toLocalFile()
        if os.path.isfile(path) and path.endswith('.xlsx'):
            self.setText(path)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
            if file:
                self.setText(file)

class MinutesGeneratorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Meeting Minutes Generator")
        self.setMinimumWidth(500)

        self.layout = QVBoxLayout()
        self.label_file = QLabel("Select an Excel Roster File:")
        self.excel_input = ExcelDropLineEdit()

        self.label_folder = QLabel("Select an Output Folder:")
        self.output_folder_input = QLineEdit()
        self.output_folder_input.setPlaceholderText("Click to select folder")
        self.output_folder_input.setReadOnly(True)
        self.output_folder_input.mousePressEvent = self.select_output_folder

        self.run_button = QPushButton("Generate Documents")
        self.status_label = QLabel("")

        self.run_button.clicked.connect(self.run_generator)
        
        self.run_e_button = QPushButton("Open Editor")
        self.run_e_button.clicked.connect(self.run_editor)
        
        self.layout.addWidget(self.label_file)
        self.layout.addWidget(self.excel_input)
        self.layout.addWidget(self.label_folder)
        self.layout.addWidget(self.output_folder_input)
        self.layout.addWidget(self.run_button)
        self.layout.addWidget(self.run_e_button)
        self.layout.addWidget(self.status_label)
        self.setLayout(self.layout)

    def select_output_folder(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
            if folder:
                self.output_folder_input.setText(folder)

    def run_generator(self):
        excel_file = self.excel_input.text().strip()
        base_output_dir = self.output_folder_input.text().strip()

        if not os.path.isfile(excel_file) or not excel_file.endswith('.xlsx'):
            QMessageBox.critical(self, "Invalid File", "Please select a valid Excel (.xlsx) file.")
            return

        if not base_output_dir or not os.path.isdir(base_output_dir):
            QMessageBox.critical(self, "Invalid Folder", "Please select a valid output folder.")
            return

        try:
            active_df, advisor_df = read(excel_file)

            docx_output = os.path.join(base_output_dir, 'Minutes')
            xlsx_output = os.path.join(base_output_dir, 'Rosters')
            write(active_df, advisor_df, docx_output_dir=docx_output, xlsx_output_dir=xlsx_output)

            self.status_label.setText("Documents generated!")
            QMessageBox.information(self, "Success", "Minutes and Rosters created.")

            # Open the output folder
            if sys.platform == 'win32':
                os.startfile(base_output_dir)
            elif sys.platform == 'darwin':
                subprocess.run(['open', base_output_dir])
            else:
                subprocess.run(['xdg-open', base_output_dir])
        except Exception as e:
            self.status_label.setText("An error occurred.")
            QMessageBox.critical(self, "Error", str(e))
    
    def run_editor(self):
        excel_file = self.excel_input.text().strip()

        if not os.path.isfile(excel_file) or not excel_file.endswith('.xlsx'):
            QMessageBox.critical(self, "Invalid File", "Please select a valid Excel (.xlsx) file.")
            return

        # keep reference on self to prevent garbage collection
        self._editor_window = launch_editor(excel_file)


def read(excel_file):
    df = pd.read_excel(excel_file, header=1)
    active_df = df[df['Status'] == 'Active'][['Last Name', 'First Name', 'Current Office']]
    advisor_df = df[df['Current Office'].isin(advisors)][['Last Name', 'First Name', 'Current Office']]

    return active_df, advisor_df

def write(active_df, advisor_df, docx_output_dir='Minutes', xlsx_output_dir='Rosters'):
    os.makedirs(docx_output_dir, exist_ok=True)
    os.makedirs(xlsx_output_dir, exist_ok=True)

    with pd.ExcelWriter(os.path.join(xlsx_output_dir, 'Officer Roster and Minutes Rosters.xlsx'), engine='openpyxl') as writer:
        create_roster(writer, xlsx_output_dir, active_df, advisor_df)

    create_bylaws_minutes(docx_output_dir, active_df)
    create_chapter_minutes(docx_output_dir, active_df, advisor_df)
    create_events_minutes(docx_output_dir, active_df)
    create_exec_minutes(docx_output_dir, active_df)
    create_finance_minutes(docx_output_dir, active_df)
    create_house_minutes(docx_output_dir, active_df, advisor_df)
    create_IOC_minutes(docx_output_dir, active_df)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MinutesGeneratorApp()
    window.show()
    sys.exit(app.exec())