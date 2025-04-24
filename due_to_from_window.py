import os
from datetime import datetime
from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                            QFileDialog, QMessageBox, QTextEdit, QApplication, QDialog, QScrollArea, QWidget)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon, QPixmap
from retro_style import RetroWindow, create_retro_central_widget
import sys


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class DueToFromWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_file = None
        self.output_dir = None

        icon_path = resource_path(os.path.join('assets', 'icon.png'))
        self.setWindowIcon(QIcon(icon_path))

    def initUI(self):
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Add title
        title_label = QLabel("Due To/From Analysis", self)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setObjectName("title_label")
        layout.addWidget(title_label)

        # Instructions button
        self.instructions_button = QPushButton('Instructions')
        self.instructions_button.clicked.connect(self.show_instructions)
        layout.addWidget(self.instructions_button)

        # Input File button and label
        input_layout = QHBoxLayout()
        self.input_button = QPushButton('Input File')
        self.input_button.clicked.connect(self.select_file)
        input_layout.addWidget(self.input_button)
        self.file_label = QLabel('No file selected')
        input_layout.addWidget(self.file_label)
        layout.addLayout(input_layout)

        # Output Directory button and label
        output_layout = QHBoxLayout()
        self.output_button = QPushButton('Output Directory')
        self.output_button.clicked.connect(self.select_output_directory)
        output_layout.addWidget(self.output_button)
        self.output_label = QLabel('No output directory selected')
        output_layout.addWidget(self.output_label)
        layout.addLayout(output_layout)

        # Run button
        self.run_button = QPushButton('RUN')
        self.run_button.clicked.connect(self.run_analysis)
        layout.addWidget(self.run_button)

        # Console output
        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)
        layout.addWidget(self.console_output)

        self.setWindowTitle('Due To/From Analysis')
        self.setFixedSize(1000, 738)
        self.center()

    def center(self):
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def select_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Input File", "", "CSV Files (*.csv)"
        )
        if file:
            if os.path.basename(file).startswith("GL"):
                self.selected_file = file
                self.file_label.setText(f"Selected file: {os.path.basename(file)}")
                self.console_output.append(f"Selected input file: {file}")
            else:
                QMessageBox.warning(self, "Error", "File name must start with 'GL'")

    def select_output_directory(self):
        self.output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if self.output_dir:
            self.output_label.setText(f"Output directory: {self.output_dir}")
            self.console_output.append(f"Selected output directory: {self.output_dir}")

    def run_analysis(self):
        if not self.selected_file or not self.output_dir:
            QMessageBox.warning(self, "Error", "Please select both input file and output directory.")
            return

        self.console_output.clear()
        self.console_output.append("Starting Due To/From analysis...")
        self.run_button.setEnabled(False)

        # Create dated output directory
        today = datetime.now().strftime('%m%d%Y')
        output_subdir = os.path.join(self.output_dir, f"Due_to_from_{today}")
        os.makedirs(output_subdir, exist_ok=True)

        self.analysis_thread = AnalysisThread(self.selected_file, output_subdir)
        self.analysis_thread.update_signal.connect(self.update_console)
        self.analysis_thread.finished_signal.connect(self.analysis_finished)
        self.analysis_thread.start()

    def update_console(self, message):
        self.console_output.append(message)

    def analysis_finished(self, success, message):
        self.run_button.setEnabled(True)
        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.critical(self, "Error", message)
        self.console_output.append(message)

    def show_instructions(self):
        instructions = """
1. Download the following file from "My Reports". You may choose the date range.
   Choose the "View" called "Due to/from"
   Run report and save to CSV.
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

        # Add image
        image_label = QLabel()
        image_path = resource_path(os.path.join('assets', 'due_instructions.png'))
        pixmap = QPixmap(image_path)
        scaled_pixmap = pixmap.scaled(1600, 1200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        image_label.setPixmap(scaled_pixmap)
        scroll_layout.addWidget(image_label)

        # Add remaining instructions
        remaining_text = """
2. Make sure not to modify the file once downloaded.

3. Click the "Input Files" button and select the CSV file

4. Click the "Output Directory" button and select where you want the output folder to be saved.

5. Click RUN to process the file
        """
        remaining_label = QLabel(remaining_text)
        remaining_label.setWordWrap(True)
        scroll_layout.addWidget(remaining_label)
        # Add some spacing at the bottom
        scroll_layout.addSpacing(10)
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
        dialog.setMinimumWidth(1600)
        dialog.setMinimumHeight(1000)
        dialog.exec_()

class AnalysisThread(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, input_file, output_dir):
        super().__init__()
        self.input_file = input_file
        self.output_dir = output_dir

    def run(self):
        try:
            self.update_signal.emit("Processing input file...")

            # Import the analysis functions
            from due_to_from_analysis import (
                process_gl_data,
                find_matches_and_mismatches,
                create_je_file,
                create_cnb_transfer_files,
                create_discrepancy_file,
                create_external_transfer_file,
                create_discrepancy_summary  # Add this import
            )

            # Process the data
            gl_data = process_gl_data(self.input_file)
            matches, mismatches = find_matches_and_mismatches(gl_data)

            # Generate output files
            today = datetime.now().strftime('%m%d%Y')

            self.update_signal.emit("Generating R365 Journal Entry file...")
            create_je_file(
                matches,
                os.path.join(self.output_dir, f'R365_JE_{today}.csv')
            )

            self.update_signal.emit("Generating CNB transfer files...")
            create_cnb_transfer_files(
                matches,
                self.output_dir
            )

            self.update_signal.emit("Generating discrepancy report...")
            create_discrepancy_file(
                mismatches,
                os.path.join(self.output_dir, f'Discrepancies_{today}.csv')
            )

            self.update_signal.emit("Generating discrepancy summary...")
            create_discrepancy_summary(
                mismatches,
                os.path.join(self.output_dir, f'Discrepancies_Summary_{today}.csv')
            )

            self.update_signal.emit("Generating external transfer file...")
            create_external_transfer_file(
                matches,
                os.path.join(self.output_dir, f'External_Transfer_{today}.csv')
            )

            message = f"Analysis complete. Output files created in:\n{self.output_dir}"
            success = True

        except Exception as e:
            success = False
            message = f"An error occurred: {str(e)}"

        self.finished_signal.emit(success, message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DueToFromWindow()
    icon_path = resource_path(os.path.join('assets', 'icon.png'))
    app.setWindowIcon(QIcon(icon_path))
    ex.show()
    sys.exit(app.exec_())
