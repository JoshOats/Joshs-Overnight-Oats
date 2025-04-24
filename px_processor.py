import os
import sys
from datetime import datetime
from pathlib import Path
from PyQt5.QtWidgets import (QVBoxLayout, QPushButton, QLabel,
                            QFileDialog, QMessageBox, QTextEdit, QApplication,
                            QListWidget)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon
from retro_style import RetroWindow, create_retro_central_widget

# Import all functions from your original code
from px_functions import (process_special_stores, save_special_stores_excel,
                        calculate_monthly_fee_percentages, create_redemption_entries,
                        create_payout_entries, create_transfer_files, create_ap_invoices, create_achb_payment)

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class PXGiftCardsWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_files = []
        icon_path = resource_path(os.path.join('assets', 'icon.png'))
        self.setWindowIcon(QIcon(icon_path))

    def initUI(self):
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel("Paytronix Gift Cards", self)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setObjectName("title_label")
        layout.addWidget(title_label)

        # Instructions button
        self.instructions_button = QPushButton('INSTRUCTIONS')
        self.instructions_button.clicked.connect(self.show_instructions)
        layout.addWidget(self.instructions_button)

        # Single file selection button
        self.input_button = QPushButton('INPUT FILES')
        self.input_button.clicked.connect(self.select_files)
        layout.addWidget(self.input_button)

        # File list with fixed height
        self.file_list = QListWidget()
        self.file_list.setFixedHeight(150)  # Adjust this value as needed
        layout.addWidget(self.file_list)

        # Run button
        self.run_button = QPushButton('RUN')
        self.run_button.clicked.connect(self.run_processing)
        layout.addWidget(self.run_button)

        # Console output
        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)
        layout.addWidget(self.console_output)

        self.setWindowTitle('Paytronix Gift Cards')
        self.resize(1000, 738)
        self.center()

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Input Files", "",
            "CSV Files (*.csv)"
        )
        if files:
            self.selected_files = files
            self.update_file_list()

    def update_file_list(self):
        self.file_list.clear()
        for file in self.selected_files:
            self.file_list.addItem(os.path.basename(file))

    def center(self):
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def show_instructions(self):
        instructions = """
Please select the following input files:
- Chase CSV file
- Payouts CSV file
- One or more StoredValueRedemption CSV files

The program will:
- Process gift card redemptions
- Generate summary reports
- Create transfer files
- Save all output to your Downloads folder
        """
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Instructions")
        msg_box.setText(instructions)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def run_processing(self):
        if not self.selected_files:
            QMessageBox.warning(self, "Error",
                              "Please select input files.")
            return

        self.console_output.clear()
        self.run_button.setEnabled(False)

        # Create processing thread
        self.process_thread = ProcessThread(self.selected_files)
        self.process_thread.update_signal.connect(self.update_console)
        self.process_thread.finished_signal.connect(self.processing_finished)
        self.process_thread.start()

    def update_console(self, message):
        self.console_output.append(message)
        # Scroll to bottom
        scrollbar = self.console_output.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def processing_finished(self, success, message):
        self.run_button.setEnabled(True)
        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.critical(self, "Error", message)
        self.console_output.append(message)

class ProcessThread(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, files):
        super().__init__()
        self.files = files

    def categorize_files(self):
        chase_file = None
        payouts_file = None
        redemption_files = []

        for file in self.files:
            filename = os.path.basename(file).lower()
            if 'chase' in filename:
                chase_file = file
            elif 'payout' in filename:
                payouts_file = file
            elif 'storedvalueredemption' in filename:
                redemption_files.append(file)

        return chase_file, payouts_file, redemption_files

    def run(self):
        import pandas as pd
        try:
            # Categorize files
            chase_file, payouts_file, redemption_files = self.categorize_files()

            # Validate files
            if not chase_file:
                raise ValueError("Chase CSV file not found in selected files")
            if not payouts_file:
                raise ValueError("Payouts CSV file not found in selected files")
            if not redemption_files:
                raise ValueError("No StoredValueRedemption CSV files found")

            # Create output directory in Downloads
            downloads_path = str(Path.home() / "Downloads")
            today = datetime.now().strftime('%m%d%Y')
            output_dir = os.path.join(downloads_path, f"PX Gift Cards - {today}")
            os.makedirs(output_dir, exist_ok=True)

            self.update_signal.emit("Processing files...")

            # Load Chase data with proper parsing
            chase_df = pd.read_csv(
                chase_file,
                sep=',',
                quotechar='"',
                dtype=str,
                skipinitialspace=True
            )

            # Look for Paytronix in the Posting Date column where it actually appears
            chase_df = chase_df[
                chase_df['Posting Date'].str.contains('ORIG CO NAME:Paytronix', case=False, na=False)
            ].copy()

            # Now realign the columns properly
            chase_df['Amount'] = pd.to_numeric(chase_df['Description'].str.strip(), errors='coerce')
            chase_df['Posting Date'] = chase_df['Details']  # Get the actual date

            # Load Paytronix data
            px_df = pd.read_csv(
                payouts_file,
                dtype={
                    'Unnamed: 0': str,
                    'Payout ID': str,
                    'Payout Status': str,
                    'Description': str,
                    'Payout Created Date': str,
                    'Payout Arrival Date': str,
                    'Gross': str,
                    'Fees': str,
                    'Total': str
                }
            )

            # Clean up Paytronix data
            px_df['Payout Created Date'] = pd.to_datetime(px_df['Payout Created Date']).dt.strftime('%m/%d/%Y')
            for col in ['Gross', 'Fees', 'Total']:
                px_df[col] = pd.to_numeric(
                    px_df[col].astype(str).str.replace('$', '').str.replace(',', ''),
                    errors='coerce'
                )

            # Load redemptions data
            all_redemptions = []
            for file in redemption_files:
                df = pd.read_csv(
                    file,
                    skiprows=1,
                    dtype={'Store Name': str, 'Card Template': str, 'Dollars Redeemed': float, 'Date': str}
                )
                all_redemptions.append(df)
            redemptions_df = pd.concat(all_redemptions, ignore_index=True)
            redemptions_df['Date'] = pd.to_datetime(redemptions_df['Date']).dt.strftime('%m/%d/%Y')

            # Get date range
            date_range = pd.to_datetime(redemptions_df['Date'])
            start_date = date_range.min()
            end_date = date_range.max()
            start_month1 = start_date.strftime('%B %Y')
            end_month1 = end_date.strftime('%B %Y')
            month_range1 = start_month1 if start_month1 == end_month1 else f"{start_month1}-{end_month1}"

            # Calculate fee percentages
            fee_percentages = calculate_monthly_fee_percentages(px_df)

            # Process and save files
            # 1. Process special stores (Maduro file)
            special_stores_df = process_special_stores(redemptions_df, fee_percentages)
            special_filename = f"MaduroPX_{month_range1.replace(' to ', '-')}.xlsx"
            save_special_stores_excel(special_stores_df, redemptions_df, os.path.join(output_dir, special_filename))

            # Create AP Invoices file
            ap_invoices_filename = create_ap_invoices(special_stores_df, output_dir)

            # 2. Create redemption entries
            redemption_entries = create_redemption_entries(redemptions_df, fee_percentages)
            redemption_filename = f"PX_Redemptions_{today}.csv"
            redemption_entries.to_csv(os.path.join(output_dir, redemption_filename), index=False)

            # 3. Create payout entries
            payout_entries, monthly_totals, grand_total = create_payout_entries(chase_df, px_df, redemptions_df)

            if payout_entries.empty:
                self.update_signal.emit("\nWARNING: No payout entries were created!")
            else:
                payout_filename = f"PX_LeadershipPayouts_{month_range1.replace(' to ', '-')}.csv"
                payout_filepath = os.path.join(output_dir, payout_filename)
                payout_entries.to_csv(payout_filepath, index=False)

                # Display transfer amounts information
                self.update_signal.emit("\nTransfer Amounts Summary:")
                for month, amount in monthly_totals.items():
                    month_date = datetime.strptime(month, '%m/%Y')
                    month_name = month_date.strftime('%B %Y')
                    self.update_signal.emit(f"{month_name}: ${amount:,.2f}")
                self.update_signal.emit(f"\nTotal Transfer Amount from Chase to CNB: ${grand_total:,.2f}")

            # 4. Create transfer files and ACHB payment file
            cnb_dfs, today = create_transfer_files(redemption_entries)

            # Create CNB transfer files
            if cnb_dfs:
                for cnb_file in cnb_dfs:
                    cnb_filepath = os.path.join(output_dir, cnb_file['filename'])
                    cnb_file['df'].to_csv(cnb_filepath, index=False)

            # Create ACHB payment file
            achb_filename = create_achb_payment(redemption_entries, output_dir)
            if achb_filename:
                self.update_signal.emit(f"\nCreated ACHB Payment file: {achb_filename}")


            # Generate email text based on time of day
            current_hour = datetime.now().hour
            greeting = (
                "morning" if 4 <= current_hour < 12 else
                "afternoon" if 12 <= current_hour < 18 else
                "evening"
            )

            email_text = f"""
eGift Cards - Paytronix // {month_range1}

Good {greeting} Carolina,

Attached please find the information relating to the gift card redemptions of your stores from {month_range1}. On the first tab you will find the summary of the money to be paid out to you and on the second tab you will find all the transactions relating to the eGift cards of your stores for the time period.
We have created these invoices in our system and will be paid out when we make our next weekly batch of AP payments.

Please let me know if you have any questions.

Thank you!
"""
            self.update_signal.emit("\nEmail Template:" + email_text)

            success_message = (
                f"Files have been created successfully in:\n{output_dir}\n\n"
                "NEXT STEPS:\n"
                """
1. Make manual transfer from Chase Leadership to CNB Leadership (Transfer amounts are at the top of output log ^)
2. Copy email from the output log, attach "MaduroPX" file, and send to Carolina Maduro and copy Alejandro, Isabel, Josh, and Gustavo
3. Uplaod "ACHB_PX" file into CNB. ELectronic Payments -> New Payment -> ACH Batch -> Upload File -> ACHB -> "ACHB_PX" file -> *2 BUSINESS DAYS LATER*
4. Upload "PX_CNB_Transfer" file into CNB to make the transfers from Leadership to the restaurants. Payments & Transfers -> Internal Transfer -> Multi-Account Transfers -> Transfer Funds -> UPLOAD FROM FILE -> *TOMORROW'S DATE*
5. Import "AP_Invoices" file for Maduros restaurants in R365. Top banner -> Vendor -> Import AP Transaction
6. Import "PX_Redemptions" file and "PX_LeadershipPayouts" file in R365. Top banner -> Account -> Import Journal Entry
            """
            )
            self.finished_signal.emit(True, success_message)

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.update_signal.emit(f"Error details:\n{error_details}")
            self.finished_signal.emit(False, f"An error occurred: {str(e)}")
