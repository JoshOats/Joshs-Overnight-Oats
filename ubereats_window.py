from PyQt5.QtWidgets import (QVBoxLayout, QPushButton, QLabel,
                           QFileDialog, QTextEdit, QMessageBox, QListWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from datetime import datetime
import os
from retro_style import RetroWindow, create_retro_central_widget


class UberEatsProcessThread(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, files):
        super().__init__()
        self.files = files

    def run(self):
        import pandas as pd
        try:
            self.update_signal.emit("Starting UberEats processing...")

            # Categorize files
            toast_files = []
            uber_file = None

            for file in self.files:
                filename = os.path.basename(file).lower()
                if filename.startswith('order'):
                    toast_files.append(file)
                else:
                    if uber_file:
                        raise ValueError("Multiple non-Order files detected. Please select only one Uber file.")
                    uber_file = file

            if not toast_files:
                raise ValueError("No Toast files found. Please select at least one Toast file.")
            if not uber_file:
                raise ValueError("No Uber file found. Please select a non-Order CSV file.")

            self.update_signal.emit(f"Found {len(toast_files)} Toast files and 1 Uber file")

            # Import the functions
            from uber_entries import read_toast_files, read_uber_file, create_journal_entries, create_deposit_journal_entries

            # Process files
            self.update_signal.emit("Reading Toast files...")
            toast_df = read_toast_files(toast_files)

            self.update_signal.emit("Reading Uber file...")
            uber_df = read_uber_file(uber_file)

            # Create journal entries
            self.update_signal.emit("Creating journal entries...")
            je_df = create_journal_entries(toast_df, uber_df)

            # Create deposit journal entries
            self.update_signal.emit("Creating deposit journal entries...")
            deposit_df = create_deposit_journal_entries(je_df)

            # Combine both sets of entries
            combined_df = pd.concat([je_df, deposit_df], ignore_index=True) if not deposit_df.empty else je_df

            # Create output directory in Downloads
            downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
            today = datetime.now().strftime("%m%d%Y")
            output_file = os.path.join(downloads_path, f"UE_PayoutImport_{today}.csv")

            # Save to CSV
            combined_df.to_csv(output_file, index=False, float_format='%.2f')

            success_msg = f"Successfully processed UberEats transactions!\nOutput saved to: {output_file}"
            self.finished_signal.emit(True, success_msg)

        except Exception as e:
            error_msg = f"Error processing files: {str(e)}"
            self.finished_signal.emit(False, error_msg)


class UberEatsWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_files = []

    def initUI(self):
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel("UberEats Payout Processing", self)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setObjectName("title_label")
        layout.addWidget(title_label)

        # Instructions button
        self.instructions_button = QPushButton('Instructions')
        self.instructions_button.clicked.connect(self.show_instructions)
        layout.addWidget(self.instructions_button)

        # Input File button and list
        input_layout = QVBoxLayout()
        self.input_button = QPushButton('Input Files')
        self.input_button.clicked.connect(self.select_files)
        input_layout.addWidget(self.input_button)
        self.file_list = QListWidget()
        self.file_list.setFixedHeight(150)
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

        self.setWindowTitle('UberEats Transaction Processing')
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
            self, "Select Input Files", "", "CSV Files (*.csv)"
        )
        if files:
            self.selected_files = files
            self.file_list.clear()
            for file in files:
                self.file_list.addItem(os.path.basename(file))
            self.console_output.append(f"Selected {len(files)} files")

    def run_processing(self):
        if not self.selected_files:
            QMessageBox.warning(self, "Error", "Please select input files.")
            return

        self.console_output.clear()
        self.console_output.append("Starting processing...")
        self.run_button.setEnabled(False)

        self.process_thread = UberEatsProcessThread(self.selected_files)
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
        instructions = """
    Instructions for UberEats Transaction Processing:

    1. Click 'Input Files' and select all required files:
    - One or more Toast CSV files (filenames start with "Order")
    - One UberEats "Payment details" CSV file (any CSV file that doesn't start with "Order")

    2. Click RUN to process the transactions

    The output file will be saved to your Downloads folder as "UE_PayoutImport_MMDDYYYY.csv"

    Note:
    - You can select multiple Toast files
    - You must select exactly one UberEats file
    - Each order date will have its own journal entry
    """
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Instructions")
        msg_box.setText(instructions)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()
