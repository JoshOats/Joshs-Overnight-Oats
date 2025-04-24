# payroll_window.py

import os
import sys
from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                            QFileDialog, QMessageBox, QTextEdit, QApplication, QListWidget)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon
from retro_style import RetroWindow, create_retro_central_widget
from payroll_automation import PayrollAutomationThread
from datetime import datetime

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class PayrollWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.selected_files = {}

        # Initialize output_dir BEFORE initUI
        try:
            self.output_dir = os.path.join(os.path.expanduser("~"), "Downloads")
            # Verify the path exists
            if not os.path.exists(self.output_dir):
                self.output_dir = os.getcwd()  # Fallback to current directory
        except Exception:
            # Fallback if any error occurs
            self.output_dir = os.getcwd()

        # Only call initUI after setting output_dir
        self.initUI()

        icon_path = resource_path(os.path.join('assets', 'icon.png'))
        self.setWindowIcon(QIcon(icon_path))

    def initUI(self):
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel("Payroll Automation", self)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setObjectName("title_label")
        layout.addWidget(title_label)

        # Instructions button
        self.instructions_button = QPushButton('Instructions')
        self.instructions_button.clicked.connect(self.show_instructions)
        layout.addWidget(self.instructions_button)

        # Time Entries File
        input_layout = QVBoxLayout()
        time_entries_layout = QHBoxLayout()
        self.time_entries_button = QPushButton('Time Entries File')
        self.time_entries_button.clicked.connect(self.select_time_entries)
        time_entries_layout.addWidget(self.time_entries_button)
        self.time_entries_label = QLabel('No file selected')
        time_entries_layout.addWidget(self.time_entries_label)
        input_layout.addLayout(time_entries_layout)

        # Payroll Dictionary File
        payroll_dict_layout = QHBoxLayout()
        self.payroll_dict_button = QPushButton('Payroll Dictionary File')
        self.payroll_dict_button.clicked.connect(self.select_payroll_dict)
        payroll_dict_layout.addWidget(self.payroll_dict_button)
        self.payroll_dict_label = QLabel('No file selected')
        payroll_dict_layout.addWidget(self.payroll_dict_label)
        input_layout.addLayout(payroll_dict_layout)

        # Tips File (Optional)
        tips_layout = QHBoxLayout()
        self.tips_button = QPushButton('Tips File')
        self.tips_button.clicked.connect(self.select_tips_file)
        tips_layout.addWidget(self.tips_button)
        self.tips_label = QLabel('No file selected')
        tips_layout.addWidget(self.tips_label)
        input_layout.addLayout(tips_layout)

        layout.addLayout(input_layout)

        # Output notification - show where files will be saved
        output_label = QLabel(f"Files will be saved to: {self.output_dir}")
        layout.addWidget(output_label)

        # Run button
        self.run_button = QPushButton('RUN')
        self.run_button.clicked.connect(self.run_payroll_automation)
        layout.addWidget(self.run_button)

        # Console output
        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)
        layout.addWidget(self.console_output)

        self.setWindowTitle('Payroll Automation')
        self.setFixedSize(1000, 738)
        self.center()

    def center(self):
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def select_time_entries(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Time Entries File", "", "CSV Files (*.csv)"
        )
        if file:
            self.selected_files['time_entries'] = file
            self.time_entries_label.setText(os.path.basename(file))
            self.console_output.append(f"Selected Time Entries file: {file}")

    def select_payroll_dict(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Payroll Dictionary File", "", "Excel Files (*.xlsx)"
        )
        if file:
            self.selected_files['payroll_dict'] = file
            self.payroll_dict_label.setText(os.path.basename(file))
            self.console_output.append(f"Selected Payroll Dictionary file: {file}")

    def select_tips_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Tips File", "", "CSV Files (*.csv)"
        )
        if file:
            self.selected_files['tips'] = file
            self.tips_label.setText(os.path.basename(file))
            self.console_output.append(f"Selected Tips file: {file}")

    def select_output_directory(self):
        self.output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if self.output_dir:
            self.output_label.setText(f"Output directory: {self.output_dir}")
            self.console_output.append(f"Selected output directory: {self.output_dir}")

    def run_payroll_automation(self):
        if 'time_entries' not in self.selected_files or 'payroll_dict' not in self.selected_files:
            QMessageBox.warning(self, "Error", "Please select Time Entries file and Payroll Dictionary file.")
            return

        # Create timestamp for output folder
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Create single output folder in Downloads
        output_folder_name = f"Payroll_Automation_{timestamp}"
        output_folder_path = os.path.join(self.output_dir, output_folder_name)
        os.makedirs(output_folder_path, exist_ok=True)

        self.console_output.clear()
        self.console_output.append(f"Starting payroll automation process...")
        self.console_output.append(f"Output will be saved to: {output_folder_path}")
        self.run_button.setEnabled(False)

        # Pass the original file paths directly to the thread
        time_entries_path = self.selected_files['time_entries']
        payroll_dict_path = self.selected_files['payroll_dict']
        tips_path = self.selected_files.get('tips', None)  # Optional tips file

        self.payroll_thread = PayrollAutomationThread(
            time_entries_path=time_entries_path,
            payroll_dict_path=payroll_dict_path,
            tips_path=tips_path,
            output_dir=output_folder_path
        )
        self.payroll_thread.update_signal.connect(self.update_console)
        self.payroll_thread.finished_signal.connect(self.payroll_finished)
        self.payroll_thread.start()

    def update_console(self, message):
        self.console_output.append(message)

    def payroll_finished(self, success, message):
        self.run_button.setEnabled(True)
        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.critical(self, "Error", message)
        self.console_output.append(message)

    def show_instructions(self):
        instructions = """
1. Select the required input files:
   - Time Entries file (CSV format)
   - Payroll Dictionary file (Excel [XLSX] format)
   - Tips file (CSV format)

2. Output folder:
   - Files will be saved to a "Payroll_Automation_[timestamp]" folder in your Downloads directory
   - No need to select an output location

3. Click RUN to process the payroll automation

The program will:
- Process time entries and calculate gross pay
- Apply holiday hours and overtime calculations
- Process tips if a tips file is provided
- Generate Payroll Summary, ADP Cargue files, and Workers Comp report
- Create warnings file with detailed information
    """
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Instructions")
        msg_box.setText(instructions)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()
