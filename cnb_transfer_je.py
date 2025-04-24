import os
import csv
from datetime import datetime
from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                             QFileDialog, QMessageBox, QTextEdit, QApplication)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from retro_style import RetroWindow, create_retro_central_widget
import sys
from PyQt5.QtGui import QIcon

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class CNBTransferJEWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_files = []  # Changed from single file to list
        self.output_dir = None

        icon_path = resource_path(os.path.join('assets', 'icon.png'))
        self.setWindowIcon(QIcon(icon_path))

    def initUI(self):
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Add a retro-style title
        title_label = QLabel("CNB Transfer JE Import", self)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setObjectName("title_label")
        layout.addWidget(title_label)

        # Instructions button
        self.instructions_button = QPushButton('Instructions')
        self.instructions_button.clicked.connect(self.show_instructions)
        layout.addWidget(self.instructions_button)

        # Input File button and label
        input_layout = QHBoxLayout()
        self.input_button = QPushButton('Input Files')
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
        self.run_button.clicked.connect(self.run_transfer)
        layout.addWidget(self.run_button)

        # Console output
        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)
        layout.addWidget(self.console_output)

        self.setWindowTitle('CNB Transfer JE Import')
        self.setFixedSize(1000, 738)
        self.center()

    def center(self):
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def select_file(self):
        files, _ = QFileDialog.getOpenFileNames(  # Changed to getOpenFileNames
            self, "Select Input Files", "", "CSV Files (*.csv)"
        )
        if files:
            valid_files = []
            for file in files:
                filename = os.path.basename(file)
                if filename.startswith("CNB") or filename.startswith("Funds"):
                    valid_files.append(file)
                else:
                    QMessageBox.warning(self, "Error", f"File {filename} must start with 'CNB' or 'Funds'")

            if valid_files:
                self.selected_files = valid_files
                file_names = [os.path.basename(f) for f in valid_files]
                self.file_label.setText(f"Selected files: {', '.join(file_names)}")
                self.console_output.append(f"Selected input files: {', '.join(valid_files)}")

    def select_output_directory(self):
        self.output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if self.output_dir:
            self.output_label.setText(f"Output directory: {self.output_dir}")
            self.console_output.append(f"Selected output directory: {self.output_dir}")

    def run_transfer(self):
        if not self.selected_files or not self.output_dir:
            QMessageBox.warning(self, "Error", "Please select both input file(s) and output directory.")
            return

        self.console_output.clear()
        self.console_output.append("Starting CNB Transfer JE process...")
        self.run_button.setEnabled(False)

        self.transfer_thread = TransferThread(self.selected_files, self.output_dir)
        self.transfer_thread.update_signal.connect(self.update_console)
        self.transfer_thread.finished_signal.connect(self.transfer_finished)
        self.transfer_thread.start()

    def update_console(self, message):
        self.console_output.append(message)

    def transfer_finished(self, success, message):
        self.run_button.setEnabled(True)
        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.critical(self, "Error", message)
        self.console_output.append(message)

    def show_instructions(self):
        instructions = """
1. Download bank transfer file from CNB or use the "CNB_Transfer" file from Due to/from analysis.
2. Make sure the name of the file starts with "Funds" or "CNB".
3. Select the input files. For "CNB_Transfer" files, you may choose multiple.
4. Select Output destination.
5. Click RUN to process the file
        """
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Instructions")
        msg_box.setText(instructions)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

class TransferThread(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, input_files, output_dir):  # Changed from single input_file to input_files
        super().__init__()
        self.input_files = input_files  # Store list of files
        self.output_dir = output_dir

        # Dictionaries for CNB format
        self.cnb_due_to_from_dict = {
            "Carrot Leadership LLC": "Due To/From Carrot Leadership LLC",
            "Carrot Express Franchise System LLC": "Due To/From Carrot Express Franchise System LLC",
            "Carrot Global LLC": "Due To/From Carrot Global",
            "Carrot Express Commissary LLC": "Due To/From Carrot Express Commissary LLC",
            "Carrot Coral GablesLove LLC (Coral Gabes)": "Due To/From Carrot Love LLC",
            "Carrot Aventura Love LLC (Aventura)": "Due To/From Carrot Love LLC",
            "Carrot North Beach Love LL (North Beach)": "Due To/From Carrot Love LLC",
            "Carrot Downtown Love Two LLC": "Due To/From Carrot Love Two LLC",
            "Carrot Love City Place Doral Operating LLC": "Due To/From CARROT LOVE CITY PLACE DORAL OPERATING LLC",
            "Carrot Love Palmetto Park Operating LLC": "Due To/From CARROT LOVE PALMETTO PARK OPERATING LLC",
            "Carrot Love Brickell Operating LLC": "Due To/From CARROT LOVE BRICKELL OPERATING LLC",
            "Carrot Love West Boca Operating LLC": "Due To/From CARROT LOVE WEST BOCA OPERATING LLC",
            "Carrot Love Aventura Mall Operating LLC": "Due To/From CARROT LOVE AVENTURA MALL OPERATING LLC",
            "Carrot Love Coconut Creek Operating LLC": "Due To/From CARROT LOVE COCONUT CREEK OPERATING LLC",
            "Carrot Love Coconut Grove Operating LLC": "Due To/From CARROT LOVE COCONUT GROVE OPERATING LLC",
            "Carrot Love Sunset Operating LLC": "Due To/From CARROT LOVE SUNSET OPERATING LLC",
            "Carrot Love Pembroke Pines Operating LLC": "Due To/From CARROT LOVE PEMBROKE PINES OPERATING LLC",
            "Carrot Love Plantation Operating LLC": "Due To/From CARROT LOVE PLANTATION OPERATING LLC",
            "Carrot Love River Lading Operating LLC": "Due To/From CARROT LOVE RIVER LANDING LLC",
            "Carrot Love Las Olas Operating LLC": "Due To/From CARROT LOVE LAS OLAS OPERATING LLC",
            "Carrot Love Hollywood Operating LLC": "Due To/From CARROT LOVE HOLLYWOOD OPERATING LLC",
            "Carrot Sobe Love South Florida Operating C LLC": "Due To/From Carrot Love South Florida Operating C LLC",
            "Carrot Love South Florida Operating A LLC": "Due To/From Carrot Love South Florida Operating A LLC",
            "Carrot Flatiron Love Manhattan Operating LLC": "Due To/From Carrot Love Manhattan Operating LLC",
            "Carrot Love Bryant Park Operating LLC": "Due To/From CARROT LOVE BRYANT PARK OPERATING LLC",
            "Carrot Love 600 Lexington LLC": "Due To/From Carrot Love Lexington 52 LLC",
            "CARROT LOVE LIBERTY STREET LLC": "Due To/From CARROT LOVE LIBERTY STREET LLC",
            "Carrot Holdings LLC": "Due To/From Carrot Holdings LLC",
            "Carrot Gem LLC": "Due To/From Carrot Gem LLC",
            "Carrot Dream LLC": "Due To/From Carrot Dream LLC",
            "Carrot Love Dadeland Operating LLC": "Due To/From CARROT LOVE DADELAND OPERATING  LLC",
            "Beyond Branding LLC": "Due To/From Beyond Branding"
        }

        self.cnb_checking_account_dict = {
            "Carrot Leadership LLC": "Checking Carrot Leadership LLC",
            "Carrot Express Franchise System LLC": "Checking Carrot Express Franchise System LLC",
            "Carrot Global LLC": "Checking Carrot Global LLC",
            "Carrot Express Commissary LLC": "Checking Carrot Express Commissary LLC",
            "Carrot Coral GablesLove LLC (Coral Gabes)": "Checking Carrot Love LLC",
            "Carrot Aventura Love LLC (Aventura)": "Checking Carrot Love LLC",
            "Carrot North Beach Love LL (North Beach)": "Checking Carrot Love LLC",
            "Carrot Downtown Love Two LLC": "Checking Carrot Love Two LLC",
            "Carrot Love City Place Doral Operating LLC": "Checking Carrot Love City Place Doral Operating LLC",
            "Carrot Love Palmetto Park Operating LLC": "Checking Carrot Love Palmetto Park Operating LLC",
            "Carrot Love Brickell Operating LLC": "Checking Carrot Love Brickell Operating LLC",
            "Carrot Love West Boca Operating LLC": "Checking Carrot Love West Boca Operating LLC",
            "Carrot Love Aventura Mall Operating LLC": "Checking Carrot Love Aventura Mall Operating LLC",
            "Carrot Love Coconut Creek Operating LLC": "Checking Carrot Love Coconut Creek Operating LLC",
            "Carrot Love Coconut Grove Operating LLC": "Checking Carrot Love Coconut Grove Operating LLC",
            "Carrot Love Sunset Operating LLC": "Checking Carrot Love Sunset Operating LLC",
            "Carrot Love Pembroke Pines Operating LLC": "Checking Carrot Love Pembroke Pines Operating LLC",
            "Carrot Love Plantation Operating LLC": "Checking Carrot Love Plantation Operating LLC",
            "Carrot Love River Lading Operating LLC": "Checking Carrot Love River Landing LLC",
            "Carrot Love Las Olas Operating LLC": "Checking Carrot Love Las Olas Operating LLC",
            "Carrot Love Hollywood Operating LLC": "Checking Carrot Love Hollywood Operating LLC",
            "Carrot Sobe Love South Florida Operating C LLC": "Checking Carrot Love South Florida Operating C LLC",
            "Carrot Love South Florida Operating A LLC": "Checking Carrot Love South Florida Operating A LLC",
            "Carrot Flatiron Love Manhattan Operating LLC": "Checking Carrot Love Manhattan Operating LLC",
            "Carrot Love Bryant Park Operating LLC": "Checking Carrot Love Bryant Park Operating LLC",
            "Carrot Love 600 Lexington LLC": "Checking Carrot Love Lexington 52 LLC",
            "CARROT LOVE LIBERTY STREET LLC": "Checking Carrot Love Liberty Street LLC",
            "Carrot Holdings LLC": "Checking Carrot Holdings LLC",
            "Carrot Gem LLC": "Checking Carrot Gem LLC",
            "Carrot Dream LLC": "Checking Carrot Dream LLC",
            "Carrot Love Dadeland Operating LLC": "Checking Carrot Love Dadeland Operating LLC",
            "Beyond Branding LLC": "Checking Beyond Branding LLC"
        }

        self.cnb_je_location_dict = {
            "Carrot Leadership LLC": "Carrot Leadership LLC",
            "Carrot Express Franchise System LLC": "Carrot Express Franchise System LLC",
            "Carrot Global LLC": "Carrot Global LLC",
            "Carrot Express Commissary LLC": "Carrot Express Commissary LLC",
            "Carrot Coral GablesLove LLC (Coral Gabes)": "Carrot Aventura Love LLC (Aventura)",
            "Carrot Aventura Love LLC (Aventura)": "Carrot Aventura Love LLC (Aventura)",
            "Carrot North Beach Love LL (North Beach)": "Carrot Aventura Love LLC (Aventura)",
            "Carrot Downtown Love Two LLC": "Carrot Downtown Love Two LLC",
            "Carrot Love City Place Doral Operating LLC": "Carrot Love City Place Doral Operating LLC",
            "Carrot Love Palmetto Park Operating LLC": "Carrot Love Palmetto Park Operating LLC",
            "Carrot Love Brickell Operating LLC": "Carrot Love Brickell Operating LLC",
            "Carrot Love West Boca Operating LLC": "Carrot Love West Boca Operating LLC",
            "Carrot Love Aventura Mall Operating LLC": "Carrot Love Aventura Mall Operating LLC",
            "Carrot Love Coconut Creek Operating LLC": "Carrot Love Coconut Creek Operating LLC",
            "Carrot Love Coconut Grove Operating LLC": "Carrot Love Coconut Grove Operating LLC",
            "Carrot Love Sunset Operating LLC": "Carrot Love Sunset Operating LLC",
            "Carrot Love Pembroke Pines Operating LLC": "Carrot Love Pembroke Pines Operating LLC",
            "Carrot Love Plantation Operating LLC": "Carrot Love Plantation Operating LLC",
            "Carrot Love River Lading Operating LLC": "Carrot Love River Lading Operating LLC",
            "Carrot Love Las Olas Operating LLC": "Carrot Love Las Olas Operating LLC",
            "Carrot Love Hollywood Operating LLC": "Carrot Love Hollywood Operating LLC",
            "Carrot Sobe Love South Florida Operating C LLC": "Carrot Sobe Love South Florida Operating C LLC",
            "Carrot Love South Florida Operating A LLC": "Carrot Love South Florida Operating A LLC",
            "Carrot Flatiron Love Manhattan Operating LLC": "Carrot Flatiron Love Manhattan Operating LLC",
            "Carrot Love Bryant Park Operating LLC": "Carrot Love Bryant Park Operating LLC",
            "Carrot Love 600 Lexington LLC": "Carrot Love 600 Lexington LLC",
            "CARROT LOVE LIBERTY STREET LLC": "CARROT LOVE LIBERTY STREET LLC",
            "Carrot Holdings LLC": "Carrot Holdings LLC",
            "Carrot Gem LLC": "Carrot Gem LLC",
            "Carrot Dream LLC": "Carrot Dream LLC",
            "Carrot Love Dadeland Operating LLC": "Carrot Love Dadeland Operating LLC",
            "Beyond Branding LLC": "Beyond Branding LLC"
        }

        self.cnb_detail_location_dict = self.cnb_je_location_dict.copy()

    def run(self):
        try:
            self.update_signal.emit("Processing input files...")
            output_file = os.path.join(self.output_dir, f"Modified_Transfer_{datetime.now().strftime('%m%d%Y')}.csv")

            all_empty_details = set()
            all_empty_jes = set()

            # Sort files into CNB and Funds
            cnb_files = []
            funds_files = []
            for file in self.input_files:
                if os.path.basename(file).startswith("CNB"):
                    cnb_files.append(file)
                else:
                    funds_files.append(file)

            # Process all CNB files together
            if cnb_files:
                cnb_empty_details, cnb_empty_jes = self.transform_multiple_cnb_transfers(cnb_files, output_file)
                all_empty_details.update(cnb_empty_details)
                all_empty_jes.update(cnb_empty_jes)

            # Process Funds files one by one (maintaining original behavior)
            for funds_file in funds_files:
                funds_output = os.path.join(self.output_dir, f"Modified_Transfer_Funds_{datetime.now().strftime('%m%d%Y')}.csv")
                empty_details, empty_jes = self.transform_funds_transfer(funds_file, funds_output)
                all_empty_details.update(empty_details)
                all_empty_jes.update(empty_jes)

            message = f"Transformation complete."
            if all_empty_details or all_empty_jes:
                message += "\n\nATTENTION: Empty values detected!"
                if all_empty_details:
                    message += "\n\nEmpty DetailLocation for:\n" + "\n".join(f"  - {account}" for account in all_empty_details)
                if all_empty_jes:
                    message += "\n\nEmpty JELocation for:\n" + "\n".join(f"  - {account}" for account in all_empty_jes)
                message += "\n\nPlease fill in these values before attempting to upload into R365."
            else:
                message += "\n\nSuccess!"

            success = True
        except Exception as e:
            success = False
            message = f"An error occurred: {str(e)}"

        self.finished_signal.emit(success, message)

    def transform_multiple_cnb_transfers(self, input_files, output_file):
        empty_detail_locations = set()
        empty_je_locations = set()

        # Initialize the output file with headers
        with open(output_file, 'w', newline='') as outfile:
            fieldnames = ["JENumber", "Type", "DetailComment", "Reversal Date", "JEComment", "JELocation",
                         "Account", "Debit", "Credit", "DetailLocation", "Date"]
            writer = csv.DictWriter(outfile, fieldnames=fieldnames)
            writer.writeheader()

        je_counter = 1
        current_date = datetime.now().strftime('%m%d%y')

        # Process each input file and append to the same output file
        for input_file in input_files:
            self.update_signal.emit(f"Processing {os.path.basename(input_file)}...")

            with open(input_file, 'r') as infile, open(output_file, 'a', newline='') as outfile:
                reader = csv.DictReader(infile)
                writer = csv.DictWriter(outfile, fieldnames=fieldnames)

                for row in reader:
                    companies = row['From company ---> To company'].split(' ---> ')
                    from_company = companies[0].strip()
                    to_company = companies[1].strip()

                    if from_company not in self.cnb_je_location_dict:
                        empty_je_locations.add(from_company)

                    amount = float(row['Amount'])
                    je_number = f"Transfer {current_date}-{je_counter:02d}"
                    je_location = self.cnb_je_location_dict.get(from_company, "")

                    rows = [
                        {
                            "JENumber": je_number,
                            "Type": "Standard",
                            "DetailComment": "",
                            "Reversal Date": "",
                            "JEComment": "",
                            "JELocation": je_location,
                            "Account": self.cnb_checking_account_dict.get(from_company, ""),
                            "Debit": "0",
                            "Credit": f"{amount:.2f}",
                            "DetailLocation": self.cnb_detail_location_dict.get(from_company, ""),
                            "Date": datetime.now().strftime('%m/%d/%Y')
                        },
                        {
                            "JENumber": je_number,
                            "Type": "Standard",
                            "DetailComment": "",
                            "Reversal Date": "",
                            "JEComment": "",
                            "JELocation": je_location,
                            "Account": self.cnb_due_to_from_dict.get(to_company, ""),
                            "Debit": f"{amount:.2f}",
                            "Credit": "0",
                            "DetailLocation": self.cnb_detail_location_dict.get(from_company, ""),
                            "Date": datetime.now().strftime('%m/%d/%Y')
                        },
                        {
                            "JENumber": je_number,
                            "Type": "Standard",
                            "DetailComment": "",
                            "Reversal Date": "",
                            "JEComment": "",
                            "JELocation": je_location,
                            "Account": self.cnb_due_to_from_dict.get(from_company, ""),
                            "Debit": "0",
                            "Credit": f"{amount:.2f}",
                            "DetailLocation": self.cnb_detail_location_dict.get(to_company, ""),
                            "Date": datetime.now().strftime('%m/%d/%Y')
                        },
                        {
                            "JENumber": je_number,
                            "Type": "Standard",
                            "DetailComment": "",
                            "Reversal Date": "",
                            "JEComment": "",
                            "JELocation": je_location,
                            "Account": self.cnb_checking_account_dict.get(to_company, ""),
                            "Debit": f"{amount:.2f}",
                            "Credit": "0",
                            "DetailLocation": self.cnb_detail_location_dict.get(to_company, ""),
                            "Date": datetime.now().strftime('%m/%d/%Y')
                        }
                    ]

                    # Check for empty DetailLocations
                    if not self.cnb_detail_location_dict.get(from_company, ""):
                        empty_detail_locations.add(from_company)
                    if not self.cnb_detail_location_dict.get(to_company, ""):
                        empty_detail_locations.add(to_company)

                    # Write all four rows
                    for new_row in rows:
                        writer.writerow(new_row)

                    je_counter += 1

        return empty_detail_locations, empty_je_locations

    def transform_funds_transfer(self, input_file, output_file):
        due_to_from_dict = {
            "Carrot Leadership, LLC 30000480952": "Due To/From Carrot Leadership LLC",
            "Carrot Franchise Systems, LLC 30000481015": "Due To/From Carrot Express Franchise System LLC",
            "Carrot Global LLC30000482356": "Due To/From Carrot Global",
            "Commissary 30000488431": "Due To/From Carrot Express Commissary LLC",
            "CARROT LOVE LLC 30000481123": "Due To/From Carrot Love LLC",
            "Carrot Love Two LLC 30000481258": "Due To/From Carrot Love Two LLC",
            "Carrot Love City Place Doral Operating, LLC 30000481978": "Due To/From CARROT LOVE CITY PLACE DORAL OPERATING LLC",
            "Carrot Love Palmetto Park Operating LLC 30000482122": "Due To/From CARROT LOVE PALMETTO PARK OPERATING LLC",
            "Carrot Love Brickell Operating LLC 30000482104": "Due To/From CARROT LOVE BRICKELL OPERATING LLC",
            "Carrot Love West Boca Operating LLC 30000482140": "Due To/From CARROT LOVE WEST BOCA OPERATING LLC",
            "Carrot Love Aventura Mall Operating, LLC 30000482023": "Due To/From CARROT LOVE AVENTURA MALL OPERATING LLC",
            "Carrot Love Coconut Creek Operating, LLC 30000482167": "Due To/From CARROT LOVE COCONUT CREEK OPERATING LLC",
            "Carrot Love Coconut Grove Operating LLC 30000482176": "Due To/From CARROT LOVE COCONUT GROVE OPERATING LLC",
            "Carrot Love Sunset Operating, LLC 30000482212": "Due To/From CARROT LOVE SUNSET OPERATING LLC",
            "New Pembroke Pines 30000594757": "Due To/From CARROT LOVE PEMBROKE PINES OPERATING LLC",
            "Carrot Love Plantation Operating ?LLC 30000482149": "Due To/From CARROT LOVE PLANTATION OPERATING LLC",
            "Carrot Love River Landing LLC 30000482230": "Due To/From CARROT LOVE RIVER LANDING LLC",
            "Carrot Love Las Olas Operating LLC 30000482158": "Due To/From CARROT LOVE LAS OLAS OPERATING LLC",
            "Carrot Love Hollywood Operating, LLC 30000482203": "Due To/From CARROT LOVE HOLLYWOOD OPERATING LLC",
            "So Flo C 30000633502": "Due To/From Carrot Love South Florida Operating C LLC",
            "Carrot Love So Flo A 30000633448": "Due To/From Carrot Love South Florida Operating A LLC",
            "Carrot Love Manhattan Operating, LLC 30000482131": "Due To/From Carrot Love Manhattan Operating LLC",
            "Carrot Love Bryant Park Operating LLC  30000482410": "Due To/From CARROT LOVE BRYANT PARK OPERATING LLC",
            "Carrot Love Lexington 52 30000510616": "Due To/From Carrot Love Lexington 52 LLC",
            "Carrot love Liberty Street LLC. 30000674938": "Due To/From CARROT LOVE LIBERTY STREET LLC",
            "Carrot Holdings LLC 30000469729": "Due To/From Carrot Holdings LLC",
            "Carrot Gem 30000488503": "Due To/From Carrot Gem LLC",
            "Carrot Dream LLC 30000482266": "Due To/From Carrot Dream LLC",
            "Carrot Love Dadeland Operating LLC 30000481834": "Due To/From CARROT LOVE DADELAND OPERATING  LLC",
            "Beyond Branding LLC 30000566218": "Due To/From Beyond Branding"
        }

        checking_account_dict = {
            "Carrot Leadership, LLC 30000480952": "Checking Carrot Leadership LLC",
            "Carrot Franchise Systems, LLC 30000481015": "Checking Carrot Express Franchise System LLC",
            "Carrot Global LLC30000482356": "Checking Carrot Global LLC",
            "Commissary 30000488431": "Checking Carrot Express Commissary LLC",
            "CARROT LOVE LLC 30000481123": "Checking Carrot Love LLC",
            "Carrot Love Two LLC 30000481258": "Checking Carrot Love Two LLC",
            "Carrot Love City Place Doral Operating, LLC 30000481978": "Checking Carrot Love City Place Doral Operating LLC",
            "Carrot Love Palmetto Park Operating LLC 30000482122": "Checking Carrot Love Palmetto Park Operating LLC",
            "Carrot Love Brickell Operating LLC 30000482104": "Checking Carrot Love Brickell Operating LLC",
            "Carrot Love West Boca Operating LLC 30000482140": "Checking Carrot Love West Boca Operating LLC",
            "Carrot Love Aventura Mall Operating, LLC 30000482023": "Checking Carrot Love Aventura Mall Operating LLC",
            "Carrot Love Coconut Creek Operating, LLC 30000482167": "Checking Carrot Love Coconut Creek Operating LLC",
            "Carrot Love Coconut Grove Operating LLC 30000482176": "Checking Carrot Love Coconut Grove Operating LLC",
            "Carrot Love Sunset Operating, LLC 30000482212": "Checking Carrot Love Sunset Operating LLC",
            "New Pembroke Pines 30000594757": "Checking Carrot Love Pembroke Pines Operating LLC",
            "Carrot Love Plantation Operating ?LLC 30000482149": "Checking Carrot Love Plantation Operating LLC",
            "Carrot Love River Landing LLC 30000482230": "Checking Carrot Love River Landing LLC",
            "Carrot Love Las Olas Operating LLC 30000482158": "Checking Carrot Love Las Olas Operating LLC",
            "Carrot Love Hollywood Operating, LLC 30000482203": "Checking Carrot Love Hollywood Operating LLC",
            "So Flo C 30000633502": "Checking Carrot Love South Florida Operating C LLC",
            "Carrot Love So Flo A 30000633448": "Checking Carrot Love South Florida Operating A LLC",
            "Carrot Love Manhattan Operating, LLC 30000482131": "Checking Carrot Love Manhattan Operating LLC",
            "Carrot Love Bryant Park Operating LLC  30000482410": "Checking Carrot Love Bryant Park Operating LLC",
            "Carrot Love Lexington 52 30000510616": "Checking Carrot Love Lexington 52 LLC",
            "Carrot love Liberty Street LLC. 30000674938": "Checking Carrot Love Liberty Street LLC",
            "Carrot Holdings LLC 30000469729": "Checking Carrot Holdings LLC",
            "Carrot Gem3 0000488503": "Checking Carrot Gem LLC",
            "Carrot Dream LLC 30000482266": "Checking Carrot Dream LLC",
            "Carrot Leadership, LLC 30000480952": "Savings Chase Carrot Leadership LLC",
            "Carrot Love Dadeland Operating LLC 30000481834": "Checking Carrot Love Dadeland Operating LLC",
            "Beyond Branding LLC 30000566218": "Checking Beyond Branding LLC"
        }

        je_location_dict = {
            "Carrot Leadership, LLC 30000480952": "Carrot Leadership LLC",
            "Carrot Franchise Systems, LLC 30000481015": "Carrot Express Franchise System LLC",
            "Carrot Global LLC30000482356": "Carrot Global LLC",
            "Commissary 30000488431": "Carrot Express Commissary LLC",
            "CARROT LOVE LLC 30000481123": "Carrot Aventura Love LLC (Aventura)",
            "Carrot Love Two LLC 30000481258": "Carrot Downtown Love Two LLC",
            "Carrot Love City Place Doral Operating, LLC 30000481978": "Carrot Love City Place Doral Operating LLC",
            "Carrot Love Palmetto Park Operating LLC 30000482122": "Carrot Love Palmetto Park Operating LLC",
            "Carrot Love Brickell Operating LLC 30000482104": "Carrot Love Brickell Operating LLC",
            "Carrot Love West Boca Operating LLC 30000482140": "Carrot Love West Boca Operating LLC",
            "Carrot Love Aventura Mall Operating, LLC 30000482023": "Carrot Love Aventura Mall Operating LLC",
            "Carrot Love Coconut Creek Operating, LLC 30000482167": "Carrot Love Coconut Creek Operating LLC",
            "Carrot Love Coconut Grove Operating LLC 30000482176": "Carrot Love Coconut Grove Operating LLC",
            "Carrot Love Sunset Operating, LLC 30000482212": "Carrot Love Sunset Operating LLC",
            "New Pembroke Pines 30000594757": "Carrot Love Pembroke Pines Operating LLC",
            "Carrot Love Plantation Operating ?LLC 30000482149": "Carrot Love Plantation Operating LLC",
            "Carrot Love River Landing LLC 30000482230": "Carrot Love River Lading Operating LLC",
            "Carrot Love Las Olas Operating LLC 30000482158": "Carrot Love Las Olas Operating LLC",
            "Carrot Love Hollywood Operating, LLC 30000482203": "Carrot Love Hollywood Operating LLC",
            "So Flo C 30000633502": "Carrot Sobe Love South Florida Operating C LLC",
            "Carrot Love So Flo A 30000633448": "Carrot Love South Florida Operating A LLC",
            "Carrot Love Manhattan Operating, LLC 30000482131": "Carrot Flatiron Love Manhattan Operating LLC",
            "Carrot Love Bryant Park Operating LLC  30000482410": "Carrot Love Bryant Park Operating LLC",
            "Carrot Love Lexington 52 30000510616": "Carrot Love 600 Lexington LLC",
            "Carrot love Liberty Street LLC. 30000674938": "CARROT LOVE LIBERTY STREET LLC",
            "Carrot Holdings LLC 30000469729": "Carrot Holdings LLC",
            "Carrot Gem 30000488503": "Carrot Gem LLC",
            "Carrot Dream LLC 30000482266": "Carrot Dream LLC",
            "Carrot Love Dadeland Operating LLC 30000481834": "Carrot Love Dadeland Operating LLC",
            "Beyond Branding LLC 30000566218": "Beyond Branding LLC"
        }

        detail_location_dict = je_location_dict.copy()

        empty_detail_locations = set()
        empty_je_locations = set()

        with open(input_file, 'r') as infile, open(output_file, 'w', newline='') as outfile:
            reader = csv.DictReader(infile)
            fieldnames = ["JENumber", "Type", "DetailComment", "Reversal Date", "JEComment", "JELocation", "Account", "Debit", "Credit", "DetailLocation", "Date"]
            writer = csv.DictWriter(outfile, fieldnames=fieldnames)
            writer.writeheader()

            je_counter = 1
            current_date = datetime.now().strftime('%m%d%y')

            for row in reader:
                from_account = row['From Account']
                to_account = row['To Account']
                amount = float(row['Amount'].replace('$', '').replace(',', '').strip())
                date = row['Will Process On']

                je_location = je_location_dict.get(from_account, "")
                if not je_location:
                    empty_je_locations.add(from_account)

                je_number = f"Transfer {current_date}-{je_counter:02d}"

                from_due_to_from = due_to_from_dict.get(from_account, "")
                to_due_to_from = due_to_from_dict.get(to_account, "")

                rows = [
                    {
                        "JENumber": je_number,
                        "Type": "Standard",
                        "DetailComment": "",
                        "Reversal Date": "",
                        "JEComment": "",
                        "JELocation": je_location,
                        "Account": checking_account_dict.get(from_account, ""),
                        "Debit": "0",
                        "Credit": f"{amount:.2f}",
                        "DetailLocation": detail_location_dict.get(from_account, ""),
                        "Date": date
                    },
                    {
                        "JENumber": je_number,
                        "Type": "Standard",
                        "DetailComment": "",
                        "Reversal Date": "",
                        "JEComment": "",
                        "JELocation": je_location,
                        "Account": to_due_to_from,
                        "Debit": f"{amount:.2f}",
                        "Credit": "0",
                        "DetailLocation": detail_location_dict.get(from_account, ""),
                        "Date": date
                    },
                    {
                        "JENumber": je_number,
                        "Type": "Standard",
                        "DetailComment": "",
                        "Reversal Date": "",
                        "JEComment": "",
                        "JELocation": je_location,
                        "Account": from_due_to_from,
                        "Debit": "0",
                        "Credit": f"{amount:.2f}",
                        "DetailLocation": detail_location_dict.get(to_account, ""),
                        "Date": date
                    },
                    {
                        "JENumber": je_number,
                        "Type": "Standard",
                        "DetailComment": "",
                        "Reversal Date": "",
                        "JEComment": "",
                        "JELocation": je_location,
                        "Account": checking_account_dict.get(to_account, ""),
                        "Debit": f"{amount:.2f}",
                        "Credit": "0",
                        "DetailLocation": detail_location_dict.get(to_account, ""),
                        "Date": date
                    }
                ]

                if not detail_location_dict.get(from_account, ""):
                    empty_detail_locations.add(from_account)
                if not detail_location_dict.get(to_account, ""):
                    empty_detail_locations.add(to_account)

                for new_row in rows:
                    writer.writerow(new_row)

                je_counter += 1

        return empty_detail_locations, empty_je_locations


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CNBTransferJEWindow()
    icon_path = resource_path(os.path.join('assets', 'icon.png'))
    app.setWindowIcon(QIcon(icon_path))
    ex.show()
    sys.exit(app.exec_())
