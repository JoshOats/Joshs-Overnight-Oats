# ap_reconcile.py

import os
import csv, sys
from datetime import datetime
from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                            QFileDialog, QMessageBox, QTextEdit, QApplication, QListWidget)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon
from retro_style import RetroWindow, create_retro_central_widget


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class APReconcileWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_files = []
        self.output_dir = None

        icon_path = resource_path(os.path.join('assets', 'icon.png'))
        self.setWindowIcon(QIcon(icon_path))

    def initUI(self):
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel("AP Reconciliation", self)
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
        self.run_button.clicked.connect(self.run_reconciliation)
        layout.addWidget(self.run_button)

        # Console output
        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)
        layout.addWidget(self.console_output)

        self.setWindowTitle('AP Reconciliation')
        self.setFixedSize(1000, 738)
        self.center()

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Input Files", "", "CSV Files (*.csv)"
        )
        if files and len(files) == 3:  # Expecting 3 CSV files now
            # Identify which file is which
            ach_file = None
            r365_file = None
            balance_file = None

            for file in files:
                if os.path.basename(file).startswith('Ach'):
                    ach_file = file
                elif os.path.basename(file).lower().startswith('balance'):
                    balance_file = file
                else:
                    r365_file = file

            if ach_file and r365_file and balance_file:
                self.selected_files = [r365_file, ach_file, balance_file]
                self.update_file_list()
                self.console_output.append(f"Selected R365 file: {r365_file}")
                self.console_output.append(f"Selected ACH file: {ach_file}")
                self.console_output.append(f"Selected Balance file: {balance_file}")
            else:
                QMessageBox.warning(self, "Error",
                    "Please select:\n- One balance CSV file\n- One file starting with 'Ach'\n- One other CSV file")
                self.selected_files = []
                self.file_list.clear()
        else:
            QMessageBox.warning(self, "Error", "Please select exactly three CSV files:\n- Balance file\n- ACH file\n- R365 file")

    def update_file_list(self):
        self.file_list.clear()
        for file in self.selected_files:
            self.file_list.addItem(os.path.basename(file))

    def center(self):
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def select_output_directory(self):
        self.output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if self.output_dir:
            self.output_label.setText(f"Output directory: {self.output_dir}")
            self.console_output.append(f"Selected output directory: {self.output_dir}")

    def run_reconciliation(self):
        if not self.selected_files or len(self.selected_files) != 3 or not self.output_dir:
            QMessageBox.warning(self, "Error", "Please select all input files and output directory.")
            return

        self.console_output.clear()
        self.console_output.append("Starting AP reconciliation process...")
        self.run_button.setEnabled(False)

        self.reconcile_thread = ReconcileThread(
            self.selected_files[0],  # R365 file
            self.selected_files[1],  # ACH file
            self.selected_files[2],  # CSV file
            self.output_dir
        )
        self.reconcile_thread.update_signal.connect(self.update_console)
        self.reconcile_thread.finished_signal.connect(self.reconciliation_finished)
        self.reconcile_thread.start()

    def update_console(self, message):
        self.console_output.append(message)

    def reconciliation_finished(self, success, message):
        self.run_button.setEnabled(True)
        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.critical(self, "Error", message)
        self.console_output.append(message)

    def show_instructions(self):
        instructions = """
1. Select three input files:
   - One CSV balance file from CNB
   - One file starting with 'Ach' downloaded
     from CNB -> Activity Center -> Star ->
     "ACHB Reconcile" -> Choose Date Range ->
     Download(ACH batch file)
   - One other CSV file (R365 AP report saved
     as a csv)
2. Select output destination directory
3. Click RUN to process the reconciliation

The program will:
- Compare payments between R365 and ACH files
- Check available balances against ACH amounts
- Generate a report of any discrepancies
- Show warnings for accounts not found in either file
        """
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Instructions")
        msg_box.setText(instructions)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

class ReconcileThread(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, r365_file, ach_file, balance_file, output_dir):
        super().__init__()
        self.r365_file = r365_file
        self.ach_file = ach_file
        self.balance_file = balance_file  # Changed from pdf_file
        self.output_dir = output_dir

        self.account_to_store = {
            "CARROT LOVE LLC 30000481123": ["Carrot Coral GablesLove LLC (Coral Gabes)",
                                       "Carrot Aventura Love LLC (Aventura)",
                                       "Carrot North Beach Love LL (North Beach)"],
            "Commissary 30000488431": ["Carrot Express Commissary LLC",
                                       "NY - Carrot Express Commissary"],
            "Carrot Love Coconut Creek Operating, LLC 30000482167": "Carrot Love Coconut Creek Operating LLC",
            "Carrot Love Coconut Grove Operating LLC 30000482176": "Carrot Love Coconut Grove Operating LLC",
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
            "Carrot Love South Florida C LLC  30000633502": "Carrot Sobe Love South Florida Operating C LLC",
            "Carrot Love So Flo A30000633448": "Carrot Love South Florida Operating A LLC",
            "Carrot Love Manhattan Operating, LLC 30000482131": "Carrot Flatiron Love Manhattan Operating LLC",
            "Carrot Love Manhattan Operating, LLC 30000482131": "Carrot Flatiron Love Manhattan Operating LLC",
            "Carrot Love Bryant Park Operating LLC  30000482410": "Carrot Love Bryant Park Operating LLC",
            "Carrot Love Bryant Park Operating LLC  30000482410": "Carrot Love Bryant Park Operating LLC",
            "Carrot Love Lexington 52 30000510616": "Carrot Love 600 Lexington LLC",
            "Carrot Love Lexington 52 30000510616": "Carrot Love 600 Lexington LLC",
            "Carrot love Liberty Street LLC.  30000674938": 'CARROT "BROOKFIED" LIBERTY STREET LLC',
            "Carrot Holdings LLC30000469729": "Carrot Holdings LLC",
            "Carrot Gem30000488503": "Carrot Gem LLC",
            "Carrot Dream LLC30000482266": "Carrot Dream LLC",
            "Carrot Leadership, LLC 30000480952": "Carrot Leadership LLC",
            "Carrot Love Dadeland Operating LLC 30000481834": "Carrot Love Dadeland Operating LLC",
            "Beyond Branding  30000566218": "Beyond Branding LLC",
            "Carrot Franchise Systems, LLC 30000481015": "Carrot Express Franchise System LLC"
        }

        self.vendor_name_mapping = {
            "Action Plumbing and Heating Blackflow Corp": "ACTION PLUMBING AND HEATING BACKFLO",
            "Choice Mechanical Refrigeration Services": "Choice Mechanical Refrigeration Ser",
            "Fire Zone Ventilation & Suppression Inc.": "Fire Zone Ventilation & Suppression",
            "Duke Martin Refrigeration & Air Cond Inc": "Duke Martin Refrigeration & Air Con",
            "Sunshine Cleaning Contractor & Services": "Sunshine Cleaning Contractor & Serv",
            "ALFRED I DUPONT BUILDING PARTNERSHIP LLP": "ALFRED I DUPONT BUILDING PARTNERSHI",
            "Universal Environmental Consulting, Inc": "Universal Environmental Consulting,",
            "Hernan Gonzalez - Petit Cash SoFLC": "Hernan Gonzalez Petit Cash SoFLC",
            "Ginette Salas Petit Cash River Landing": "Ginette Salas PC River Landing",
            "Currus Group, LLC": "Currus Group LLC",
            "Fabiola Cavalier PC  Commissary": "Fabiola Cavalier PC Commissary",
        }

    def extract_balances_from_csv(self):
        balances = {}
        try:
            current_balance = None
            with open(self.balance_file, 'r') as f:
                lines = f.readlines()

                # Process lines in groups of 3 (Available Balance, Current Balance, Account Name)
                i = 0
                while i < len(lines):
                    line = lines[i].strip().strip('"')

                    # If we find an Available Balance line
                    if 'Available Balance' in line:
                        try:
                            # Get the Available Balance amount
                            amount_str = line.split('$')[1].strip()
                            amount = float(amount_str.replace(',', ''))

                            # Skip Current Balance line
                            i += 1

                            # Get the account name from the next line
                            if i + 1 < len(lines):
                                account_line = lines[i + 1].strip().strip('"')
                                if account_line.startswith('City National Bank of Florida'):
                                    account_info = account_line.replace('City National Bank of Florida ', '').strip()

                                    # Skip loan accounts
                                    if not any(loan in account_info for loan in
                                        ['COMMERCIAL TERM LOAN- US ADDRESSEE 149800',
                                         'COMMERCIAL TERM LOAN- US ADDRESSEE 152110',
                                         '3rd Loan CNB 153570']):

                                        # Extract just the text part before the account number
                                        account_name = ' '.join(account_info.split()[:-1])

                                        # Handle special cases
                                        if account_name == "REGULAR COMMERCIAL CHECKING":
                                            account_name = "Beyond Branding"
                                        elif account_name == "Carrot Love Plantation Operating LLC":
                                            account_name = "Carrot Love Plantation Operating ?LLC"

                                        balances[account_name] = amount
                                        self.update_signal.emit(f"Stored balance for {account_name}: {amount}")
                        except (ValueError, IndexError) as e:
                            self.update_signal.emit(f"Error processing balance: {str(e)}")
                    i += 1

            self.update_signal.emit(f"Extracted balances: {balances}")
            return balances

        except Exception as e:
            self.update_signal.emit(f"Error processing balance file: {str(e)}")
            return {}

    def run(self):
        import pandas as pd
        try:
            self.update_signal.emit("Reading input files...")

            # Extract balances from CSV
            balances = self.extract_balances_from_csv()

            # Initialize summary data
            summary_data = []

            # Create store to account mapping
            store_to_account = {}
            for account, stores in self.account_to_store.items():
                if isinstance(stores, list):
                    for store in stores:
                        store_to_account[store] = account
                else:
                    store_to_account[stores] = account

            # Create reverse mapping
            reverse_account_to_store = {}
            for store, account in store_to_account.items():
                if account not in reverse_account_to_store:
                    reverse_account_to_store[account] = []
                reverse_account_to_store[account].append(store)

            # Create case-insensitive vendor mapping
            vendor_mapping_lower = {k.lower(): v for k, v in self.vendor_name_mapping.items()}

            # Read R365 file and debug columns
            r365_df = pd.read_csv(self.r365_file, skiprows=1)
            self.update_signal.emit(f"R365 columns found: {list(r365_df.columns)}")

            # Process R365 data
            excluded_stores = ["Carrot Express Midtown LLC", "Carrot Express Miami Shores LLC"]
            store_column = next((col for col in r365_df.columns if 'store' in col.lower()), None)
            if not store_column:
                raise ValueError("Could not find Store column in R365 file")

            payment_type_column = next((col for col in r365_df.columns if 'payment type' in col.lower()), None)
            if not payment_type_column:
                raise ValueError("Could not find Payment Type column in R365 file")

            vendor_column = next((col for col in r365_df.columns if 'vendor' in col.lower()), None)
            if not vendor_column:
                raise ValueError("Could not find Vendor column in R365 file")

            total_column = next((col for col in r365_df.columns if 'total' in col.lower()), None)
            if not total_column:
                raise ValueError("Could not find Total column in R365 file")

            invoice_number_column = next((col for col in r365_df.columns if 'invoice number' in col.lower()), None)
            if not invoice_number_column:
                raise ValueError("Could not find Invoice Number column in R365 file")

            # Find the Approved Payment Date column
            approved_payment_date_column = next((col for col in r365_df.columns if 'approved payment date' in col.lower()), None)
            if not approved_payment_date_column:
                raise ValueError("Could not find Approved Payment Date column in R365 file")

            # Filter R365 data
            r365_df = r365_df[~r365_df[store_column].isin(excluded_stores)]

            # Calculate ACHB totals
            achb_df = r365_df[r365_df[payment_type_column] == 'ACHB']

            # Get WIRE and RENT entries
            wire_rent_df = r365_df[r365_df[payment_type_column].isin(['WIRE', 'RENT'])]

            # Handle XFR entries separately - EXCLUDE if Approved Payment Date contains "Paid"
            xfr_df = r365_df[r365_df[payment_type_column] == 'XFR']
            valid_xfr_df = xfr_df[
                ~xfr_df[approved_payment_date_column].astype(str).str.contains('paid', case=False, na=False)
            ]

            # Combine valid XFR entries with WIRE and RENT entries
            wire_xfr_df = pd.concat([wire_rent_df, valid_xfr_df])

            # Group R365 data and calculate net amounts
            r365_totals = {}
            r365_nets = {}
            wire_xfr_nets = {}

            # Track all accounts that have any kind of data
            all_accounts = set()

            # Process ACHB data
            for store in achb_df[store_column].unique():
                if store in store_to_account:
                    account = store_to_account[store]
                    all_accounts.add(account)
                    store_data = achb_df[achb_df[store_column] == store]

                    # Convert Total column to float
                    store_data[total_column] = store_data[total_column].apply(lambda x: float(str(x).replace('$', '').replace(',', '')))

                    if account not in r365_totals:
                        r365_totals[account] = {}
                        r365_nets[account] = 0

                    for vendor, amount in store_data.groupby(vendor_column)[total_column].sum().items():
                        if vendor in r365_totals[account]:
                            r365_totals[account][vendor] += amount
                        else:
                            r365_totals[account][vendor] = amount
                        r365_nets[account] += amount

            # Process WIRE/XFR/RENT data (excluding XFR entries with "Paid")
            for store in wire_xfr_df[store_column].unique():
                if store in store_to_account:
                    account = store_to_account[store]
                    all_accounts.add(account)
                    store_data = wire_xfr_df[wire_xfr_df[store_column] == store]

                    # Convert Total column to float
                    store_data[total_column] = store_data[total_column].apply(lambda x: float(str(x).replace('$', '').replace(',', '')))

                    if account not in wire_xfr_nets:
                        wire_xfr_nets[account] = 0

                    wire_xfr_nets[account] += store_data[total_column].sum()

            # Process ACH data
            self.update_signal.emit("Processing ACH data...")
            ach_totals = {}
            ach_nets = {}
            current_account = None
            account_totals = {}
            current_net = 0

            with open(self.ach_file, 'r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row['From Account'].strip():
                        # Only save and reset when we find a new account
                        new_account = row['From Account'].strip()
                        if current_account and new_account != current_account:
                            all_accounts.add(current_account)
                            if current_account in ach_totals:
                                for vendor, amount in account_totals.items():
                                    ach_totals[current_account][vendor] = ach_totals[current_account].get(vendor, 0) + amount
                                ach_nets[current_account] = ach_nets.get(current_account, 0) + current_net
                            else:
                                ach_totals[current_account] = account_totals.copy()
                                ach_nets[current_account] = current_net
                            account_totals = {}
                            current_net = 0
                        current_account = new_account
                    elif row['Recipient Payment Amount'].strip():
                        recipient = row['Recipient Name']
                        amount_str = row['Recipient Payment Amount'].strip().replace('$', '').replace(',', '')
                        amount = float(amount_str)

                        if recipient in account_totals:
                            account_totals[recipient] += amount
                        else:
                            account_totals[recipient] = amount
                        current_net += amount

            # Save the last account's totals
            if current_account and account_totals:
                all_accounts.add(current_account)
                if current_account in ach_totals:
                    for vendor, amount in account_totals.items():
                        ach_totals[current_account][vendor] = ach_totals[current_account].get(vendor, 0) + amount
                    ach_nets[current_account] = ach_nets.get(current_account, 0) + current_net
                else:
                    ach_totals[current_account] = account_totals.copy()
                    ach_nets[current_account] = current_net

            # Compare and find discrepancies
            self.update_signal.emit("Comparing files and checking for discrepancies...")
            discrepancies = []

            # Update to include all accounts with any data
            for account in sorted(all_accounts):
                r365_net = r365_nets.get(account, 0)
                ach_net = ach_nets.get(account, 0)
                wire_xfr_net = wire_xfr_nets.get(account, 0)

                # Extract the account name without the number for matching
                account_name = ' '.join(account.split()[:-1])
                available_balance = balances.get(account_name, 0)

                # Calculate total payment amount including WIRE/XFR
                total_payment = ach_net + wire_xfr_net
                payment_status = "Pay" if available_balance >= total_payment else "Insufficient Funds"

                # Process discrepancies if there's a net difference
                if abs(r365_net - ach_net) > 0.001:
                    bank_discrepancies = []
                    r365_vendor_totals = r365_totals.get(account, {})
                    ach_vendor_totals = ach_totals.get(account, {})

                    # Check R365 against ACH
                    for vendor, r365_amount in r365_vendor_totals.items():
                        ach_vendor = vendor_mapping_lower.get(vendor.lower(), vendor)
                        ach_amount = 0
                        for ach_v, amt in ach_vendor_totals.items():
                            if ach_v.lower() == ach_vendor.lower():
                                ach_amount = amt
                                break

                        if abs(r365_amount - ach_amount) > 0.001:
                            invoice_number = ''
                            vendor_data = r365_df[(r365_df[store_column].map(store_to_account) == account) &
                                                (r365_df[vendor_column] == vendor)]
                            if not vendor_data.empty:
                                invoice_number = vendor_data[invoice_number_column].iloc[0]

                            bank_discrepancies.append({
                                'Bank': account,
                                'Vendor': vendor,
                                'Invoice Number': invoice_number,
                                'R365 Amount': f"${r365_amount:,.2f}",
                                'ACH Amount': f"${ach_amount:,.2f}",
                                'Difference': f"${r365_amount - ach_amount:,.2f}",
                                'Source': 'R365'
                            })

                    # Check ACH against R365
                    for ach_vendor, ach_amount in ach_vendor_totals.items():
                        r365_vendor = ach_vendor
                        r365_amount = 0
                        for k, v in vendor_mapping_lower.items():
                            if v.lower() == ach_vendor.lower():
                                r365_vendor = next(orig_k for orig_k in self.vendor_name_mapping.keys()
                                                if orig_k.lower() == k)
                                break
                        for r365_v, amt in r365_vendor_totals.items():
                            if r365_v.lower() == r365_vendor.lower():
                                r365_amount = amt
                                break

                        if abs(ach_amount - r365_amount) > 0.001:
                            already_recorded = any(
                                d['Bank'] == account and
                                d['Vendor'].lower() == r365_vendor.lower() and
                                abs(float(d['ACH Amount'].replace('$', '').replace(',', '')) - ach_amount) <= 0.01
                                for d in bank_discrepancies
                            )

                            if not already_recorded:
                                bank_discrepancies.append({
                                    'Bank': account,
                                    'Vendor': ach_vendor,
                                    'Invoice Number': '',
                                    'R365 Amount': f"${r365_amount:,.2f}",
                                    'ACH Amount': f"${ach_amount:,.2f}",
                                    'Difference': f"${r365_amount - ach_amount:,.2f}",
                                    'Source': 'ACH'
                                })

                    if bank_discrepancies:
                        bank_discrepancies[-1].update({
                            'R365 Net': f"${r365_net:,.2f}",
                            'ACH Net': f"${ach_net:,.2f}",
                            'Net Difference': f"${r365_net - ach_net:,.2f}"
                        })

                        if account != sorted(all_accounts)[-1]:
                            bank_discrepancies.append({
                                'Bank': '',
                                'Vendor': '',
                                'Invoice Number': '',
                                'R365 Amount': '',
                                'ACH Amount': '',
                                'Difference': '',
                                'Source': '',
                                'R365 Net': '',
                                'ACH Net': '',
                                'Net Difference': ''
                            })

                        discrepancies.extend(bank_discrepancies)

                # Add to summary data if any amount exists OR if the account appears in the ACH file
                if r365_net != 0 or ach_net != 0 or wire_xfr_net != 0 or account in ach_totals:
                    summary_data.append({
                        'Bank': account,
                        'R365 Net': f"${r365_nets.get(account, 0):,.2f}",
                        'ACH Net': f"${ach_nets.get(account, 0):,.2f}",
                        'Difference': f"${r365_nets.get(account, 0) - ach_nets.get(account, 0):,.2f}",
                        'WIRE/XFR/RENT Net': f"${wire_xfr_nets.get(account, 0):,.2f}",
                        'Available Balance': f"${available_balance:,.2f}",
                        'Payment Status': payment_status
                    })

            # Create folder for outputs
            folder_name = f"AP_Reconciliation_{datetime.now().strftime('%m-%d-%y')}"
            output_folder = os.path.join(self.output_dir, folder_name)
            os.makedirs(output_folder, exist_ok=True)

            # Write summary file
            summary_filename = os.path.join(output_folder, f"AP_Summary_{datetime.now().strftime('%m-%d-%y')}.csv")
            with open(summary_filename, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['Bank', 'R365 Net', 'ACH Net', 'Difference',
                                    'WIRE/XFR/RENT Net', 'Available Balance', 'Payment Status'])
                writer.writeheader()
                writer.writerows(summary_data)

            # Write discrepancies file if there are any
            if discrepancies:
                output_filename = os.path.join(output_folder, f"AP_Reconcile_{datetime.now().strftime('%m-%d-%y')}.csv")
                with open(output_filename, 'w', newline='') as f:
                    writer = csv.DictWriter(f, fieldnames=['Bank', 'Vendor', 'Invoice Number', 'R365 Amount',
                                                        'ACH Amount', 'Difference', 'Source', 'R365 Net',
                                                        'ACH Net', 'Net Difference'])
                    writer.writeheader()
                    writer.writerows(discrepancies)

                self.update_signal.emit("\nWarnings:")
                for account in r365_totals.keys():
                    if account not in ach_totals:
                        self.update_signal.emit(f"Account {account} from R365 not found in ACH file")
                for account in ach_totals.keys():
                    if account not in r365_totals:
                        self.update_signal.emit(f"Account {account} from ACH not found in R365")

                self.finished_signal.emit(True, f"Discrepancies found! Check {folder_name} folder for details.")
            else:
                self.finished_signal.emit(True, f"Good news! No discrepancies found. Summary file created in {folder_name} folder.")

        except Exception as e:
                import traceback
                self.update_signal.emit(f"Error details: {traceback.format_exc()}")
                self.finished_signal.emit(False, f"An error occurred: {str(e)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = APReconcileWindow()
    ex.show()
    sys.exit(app.exec_())
