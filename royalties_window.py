from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                           QFileDialog, QTextEdit, QMessageBox, QListWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from datetime import datetime
import os
from retro_style import RetroWindow, create_retro_central_widget
from royalties_processor import RoyaltiesProcessThread

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    import sys
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class RoyaltiesWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_files = []

    def initUI(self):
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel("Royalties Processing", self)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setObjectName("title_label")
        layout.addWidget(title_label)

        # Instructions button
        self.instructions_button = QPushButton('Instructions')
        self.instructions_button.clicked.connect(self.show_instructions)
        layout.addWidget(self.instructions_button)

        # Input Files button and list
        input_layout = QVBoxLayout()
        self.input_button = QPushButton('Input Files')
        self.input_button.clicked.connect(self.select_files)
        input_layout.addWidget(self.input_button)
        self.file_list = QListWidget()
        self.file_list.setFixedHeight(150)  # Adjust this value as needed
        input_layout.addWidget(self.file_list)
        layout.addLayout(input_layout)

        # Run button
        self.run_button = QPushButton('RUN')
        self.run_button.clicked.connect(self.run_processing)
        layout.addWidget(self.run_button)

        # Console output
        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)
        layout.addWidget(self.console_output)

        self.setWindowTitle('Royalties Processing')
        self.resize(1000, 738)
        self.center()

    def center(self):
        from PyQt5.QtWidgets import QApplication
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Input Files", "", "CSV/Excel Files (*.csv *.xlsx *.xls)"
        )
        if files:
            self.selected_files.extend(files)
            self.update_file_list()
            self.console_output.append(f"Selected {len(files)} file(s)")

    def update_file_list(self):
        self.file_list.clear()
        for file in self.selected_files:
            self.file_list.addItem(os.path.basename(file))

    def run_processing(self):
        if not self.selected_files:
            QMessageBox.warning(self, "Error", "Please select input files.")
            return

        # Get Downloads folder path
        downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')

        self.console_output.clear()
        self.console_output.append("Starting royalties processing...")
        self.run_button.setEnabled(False)

        self.process_thread = RoyaltiesProcessThread(self.selected_files, downloads_path)
        self.process_thread.update_signal.connect(self.update_console)
        self.process_thread.finished_signal.connect(self.processing_finished)
        self.process_thread.start()

    def update_console(self, message):
        self.console_output.append(message)

    def processing_finished(self, success, message):
        self.run_button.setEnabled(True)

        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.critical(self, "Error", message)

        self.console_output.append(message)

    def show_instructions(self):
        from PyQt5.QtWidgets import QScrollArea, QWidget, QDialog

        instructions = """
Royalties Processing Instructions:

This tool processes sales data to calculate royalties and create various reports.

Required Files:
- GroupOverview CSV files (contain net sales data by location)
- Order CSV files (contain details on orders)
- GL CSV files (contain R365 sales tax data)
- Profit Loss CSV files (contain P&L data by location)
- Tax CSV files (contain tax-exempt sales)
- DoorDash CSV files (DoorDash sales data)
- GrubHub CSV files (GrubHub sales data)
- UE CSV files (UberEats data)

Steps:
1. Click "Input Files" to select all input files (you can select multiple files at once)
2. Click "RUN" to process the files
3. Output will be saved to your Downloads folder in a "Royalties" directory with the date range

The tool will generate:
- Royalties Summary workbook with detailed calculations
- AR Invoices CSV file for accounts receivable
- AP Invoices CSV file for accounts payable

Note: The process may take a few minutes depending on the number of files.
"""

        # Create a custom dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Instructions")

        # Create main layout
        main_layout = QVBoxLayout(dialog)

        # Create scroll area to handle potential overflow
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        # Add text instructions
        text_label = QLabel(instructions)
        text_label.setWordWrap(True)
        scroll_layout.addWidget(text_label)

        # Set up scroll area
        scroll.setWidget(scroll_widget)
        main_layout.addWidget(scroll)

        # Create button layout
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(dialog.accept)
        button_layout.addStretch()
        button_layout.addWidget(ok_button)

        # Add button layout to main layout
        main_layout.addLayout(button_layout)

        # Set dialog size
        dialog.setMinimumWidth(800)
        dialog.setMinimumHeight(500)
        dialog.exec_()
