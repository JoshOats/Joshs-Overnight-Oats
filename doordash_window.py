from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                           QFileDialog, QTextEdit, QMessageBox, QListWidget, QApplication)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from datetime import datetime, timedelta
import os
import csv
import numpy as np
from calendar import monthrange
from collections import defaultdict
from retro_style import RetroWindow, create_retro_central_widget


def get_deposit_date(order_date_str):
    """
    Calculate deposit date based on the order date
    Orders from Monday-Sunday are deposited on the Friday of the following week
    """
    # Parse the order date
    try:
        if '-' in order_date_str:
            order_date = datetime.strptime(order_date_str, '%Y-%m-%d')
        else:
            order_date = datetime.strptime(order_date_str, '%m/%d/%Y')
    except ValueError:
        raise ValueError(f"Could not parse date: {order_date_str}")

    # Determine the Sunday at the end of the current week
    days_until_sunday = 6 - order_date.weekday()  # 0=Monday, 6=Sunday
    if days_until_sunday < 0:
        days_until_sunday += 7

    end_of_week = order_date + timedelta(days=days_until_sunday)

    # Deposit date is the Friday of the following week
    deposit_date = end_of_week + timedelta(days=5)  # 5 days from Sunday to Friday

    return deposit_date.strftime('%m/%d/%Y')


def load_doordash_data(file):
    """
    Load data from Doordash transaction file
    """
    try:
        # Try different encodings
        for encoding in ['utf-8', 'latin1', 'cp1252']:
            try:
                with open(file, 'r', encoding=encoding) as f:
                    reader = csv.DictReader(f)
                    return list(reader)
            except UnicodeDecodeError:
                continue

        # If all encodings fail
        raise ValueError(f"Unable to read file {file} with any supported encoding")
    except Exception as e:
        raise Exception(f"Error loading DoorDash data: {str(e)}")


def preprocess_toast_data(toast_files):
    import pandas as pd
    """
    Load Toast data and preprocess it for direct use
    """
    # Dictionary to hold data by location, date, and type
    toast_summary = defaultdict(lambda: defaultdict(lambda: {'delivery': 0, 'pickup': 0, 'tax': 0}))

    for file in toast_files:
        try:
            # Try different encodings
            df = None
            for encoding in ['utf-8', 'cp1252', 'latin1']:
                try:
                    df = pd.read_csv(file, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue

            if df is None:
                continue

            # Process all DoorDash orders row by row
            for _, row in df.iterrows():
                if 'Dining Options' not in df.columns or pd.isna(row.get('Dining Options', '')):
                    continue

                dining_option = str(row.get('Dining Options', '')).strip()

                # Only process DoorDash orders
                if 'DoorDash' not in dining_option:
                    continue

                # Extract date and location
                date_str = str(row.get('Opened', '')).split(' ')[0]
                location = str(row.get('Location', '')).strip()

                # Use the actual location from Toast
                doordash_location = f"Carrot Express ({location})"

                # Get amount and tax
                amount = float(row.get('Amount', 0) or 0)
                tax = float(row.get('Tax', 0) or 0)

                # Add up values based on dining option
                if dining_option == 'DoorDash (Delivery)':
                    toast_summary[doordash_location][date_str]['delivery'] += amount
                    toast_summary[doordash_location][date_str]['tax'] += tax
                elif dining_option == 'DoorDash (Pickup)':
                    toast_summary[doordash_location][date_str]['pickup'] += amount
                    toast_summary[doordash_location][date_str]['tax'] += tax

        except Exception as e:
            pass  # Continue to next file if there's an error

    return toast_summary


def process_transactions(doordash_data, toast_summary):
    """
    Process transactions and create journal entries
    Handle FEE transactions with no corresponding sales by combining with other dates
    """
    # Mapping dictionaries
    location_mapping = {
        "Carrot Express (Flatiron)": "Carrot Flatiron Love Manhattan Operating LLC",
        "Carrot Express (Bryant Park)": "Carrot Love Bryant Park Operating LLC",
        "Carrot Express (Lexington)": "Carrot Love 600 Lexington LLC"
    }

    # Group doordash data by location and date
    grouped_data = defaultdict(lambda: defaultdict(list))
    all_dates = set()
    all_locations = set()

    # First pass: organize all transactions by location and date
    for row in doordash_data:
        store_name = row.get('Store Name', '')
        if store_name not in location_mapping:
            continue

        # Use Timestamp Local Date for the date
        date = row.get('Timestamp Local Date', '')
        if not date:
            continue

        grouped_data[store_name][date].append(row)
        all_dates.add(date)
        all_locations.add(store_name)

    # Find isolated FEE transactions (dates with only FEE transactions, no sales)
    isolated_fees = {}  # Store location -> date -> fee amount
    deposit_dates = {}  # Store date -> deposit date mapping

    # Calculate deposit dates for all order dates
    for date in all_dates:
        deposit_dates[date] = get_deposit_date(date)

    # Identify dates with only FEE transactions for each location
    for location in all_locations:
        for date in all_dates:
            transactions = grouped_data[location][date]
            if not transactions:
                continue

            # Check if all transactions are FEE type
            all_fee = all(t.get('Transaction Type') == 'FEE' for t in transactions)
            has_fees = any(t.get('Transaction Type') == 'FEE' for t in transactions)

            if all_fee and has_fees:
                # Calculate total fee amount
                fee_amount = sum(float(t.get('Debit', 0) or 0) for t in transactions)
                if fee_amount > 0:
                    if location not in isolated_fees:
                        isolated_fees[location] = {}
                    isolated_fees[location][date] = fee_amount
                    # Remove these transactions so they won't be processed normally
                    grouped_data[location][date] = []

    # Redistribute isolated fees to other dates with the same deposit date
    for location, fee_dates in isolated_fees.items():
        for fee_date, fee_amount in fee_dates.items():
            fee_deposit_date = deposit_dates[fee_date]

            # Find another date for this location with the same deposit date
            target_date = None
            for date in all_dates:
                # Check if this date has transactions for this location
                if (date != fee_date and  # Not the fee date
                    grouped_data[location][date] and  # Has transactions
                    deposit_dates[date] == fee_deposit_date):  # Same deposit date
                    target_date = date
                    break

            if target_date:
                # Create a synthetic FEE transaction and add it to the target date
                synthetic_fee = {
                    'Transaction Type': 'FEE',
                    'Debit': str(fee_amount),
                    'Credit': '0',
                    'Store Name': location
                }
                grouped_data[location][target_date].append(synthetic_fee)
            else:
                # If no target date found, restore the original transactions
                fee_transactions = [t for t in doordash_data
                                   if t.get('Store Name') == location
                                   and t.get('Timestamp Local Date') == fee_date]
                grouped_data[location][fee_date] = fee_transactions

    # Generate journal entries
    journal_entries = []
    # Keep these variables for compatibility but we're not tracking exchange totals anymore
    exchange_totals_by_month = defaultdict(lambda: defaultdict(float))
    exchange_totals = defaultdict(float)
    transaction_dates = []  # Track transaction dates for reporting
    je_counter = 1

    today = datetime.now().strftime("%m%d%Y")

    for location in sorted(all_locations):
        je_location = location_mapping.get(location, location)
        je_suffix = je_location.split('Love')[1].strip().split()[0][0] if 'Love' in je_location else 'X'

        for date in sorted(all_dates):
            transactions = grouped_data[location][date]
            if not transactions:
                continue

            # Convert date string to datetime object for month tracking
            date_obj = None
            try:
                if '-' in date:
                    date_obj = datetime.strptime(date, '%Y-%m-%d')
                else:
                    date_obj = datetime.strptime(date, '%m/%d/%Y')
                transaction_dates.append(date_obj)
            except ValueError:
                pass  # Skip invalid dates

            # Format deposit date
            deposit_date = get_deposit_date(date)

            # Pre-calculate some common sums
            subtotal_sum = sum(float(t.get('Subtotal', 0) or 0) for t in transactions)
            subtotal_tax_sum = sum(float(t.get('Subtotal Tax Passed by DoorDash to Merchant', 0) or 0) for t in transactions)

            picked_up_transactions = [t for t in transactions if t.get('Final Order Status') == 'Picked Up']
            not_picked_up_transactions = [t for t in transactions if t.get('Final Order Status') != 'Picked Up']

            picked_up_subtotal = sum(float(t.get('Subtotal', 0) or 0) for t in picked_up_transactions)
            not_picked_up_subtotal = sum(float(t.get('Subtotal', 0) or 0) for t in not_picked_up_transactions)

            # Calculate merchant funded discount for pickup and delivery
            picked_up_discount = sum(float(t.get('Merchant funded subtotal discount amount', 0) or 0) for t in picked_up_transactions)
            not_picked_up_discount = sum(float(t.get('Merchant funded subtotal discount amount', 0) or 0) for t in not_picked_up_transactions)

            # Calculate commission for pickup and delivery, but don't subtract discount yet
            picked_up_commission_before_discount = sum(float(t.get('Commission', 0) or 0) + float(t.get('Marketing Fees', 0) or 0) for t in picked_up_transactions)
            not_picked_up_commission_before_discount = sum(float(t.get('Commission', 0) or 0) + float(t.get('Marketing Fees', 0) or 0) for t in not_picked_up_transactions)

            # Subtract merchant funded discount from commissions
            picked_up_commission = picked_up_commission_before_discount - picked_up_discount
            not_picked_up_commission = not_picked_up_commission_before_discount - not_picked_up_discount

            error_charge_sum = sum(float(t.get('Error Charge', 0) or 0) for t in transactions)

            # Calculate fees from FEE transaction types (including synthetic ones)
            fee_debit_sum = sum(float(t.get('Debit', 0) or 0) for t in transactions if t.get('Transaction Type') == 'FEE')

            # Convert the date format directly
            toast_lookup_date = None

            # DoorDash data has dates in YYYY-MM-DD format (e.g., 2025-02-10)
            # Toast data has dates in M/D/YY format (e.g., 2/10/25)
            if '-' in date:
                # Split the date into year, month, day
                year, month, day = date.split('-')

                # Remove leading zeros from month and day
                month = month.lstrip('0')
                day = day.lstrip('0')

                # Take only the last 2 digits of the year
                short_year = year[2:]

                # Format as M/D/YY
                toast_lookup_date = f"{month}/{day}/{short_year}"

            # Get Toast data for this location and date
            delivery_amount = 0
            pickup_amount = 0
            tax_amount = 0

            if toast_lookup_date and location in toast_summary and toast_lookup_date in toast_summary[location]:
                delivery_amount = toast_summary[location][toast_lookup_date]['delivery']
                pickup_amount = toast_summary[location][toast_lookup_date]['pickup']
                tax_amount = toast_summary[location][toast_lookup_date]['tax']

            # Row 1: A/R DoorDash
            row1_val = subtotal_sum + subtotal_tax_sum
            row1_debit = 0 if row1_val >= 0 else abs(row1_val)
            row1_credit = row1_val if row1_val >= 0 else 0

            # Row 2: Commission to Fee - Doordash (Picked Up)
            row2_val = picked_up_commission
            row2_debit = row2_val if row2_val >= 0 else 0
            row2_credit = 0 if row2_val >= 0 else abs(row2_val)

            # Row 3: Commission to Fee - Doordash (Not Picked Up)
            row3_val = not_picked_up_commission
            row3_debit = row3_val if row3_val >= 0 else 0
            row3_credit = 0 if row3_val >= 0 else abs(row3_val)

            # Row 4: NEW Combined Doordash Discount (Pickup + Delivery)
            total_discount = picked_up_discount + not_picked_up_discount
            row4_val = total_discount
            row4_debit = row4_val if row4_val >= 0 else 0
            row4_credit = 0 if row4_val >= 0 else abs(row4_val)

            # Row 5: A/R DoorDash (complex calculation) - Updated row number
            row5_val = (subtotal_sum + subtotal_tax_sum - not_picked_up_commission - picked_up_commission - error_charge_sum - total_discount) - fee_debit_sum
            row5_debit = row5_val if row5_val >= 0 else 0
            row5_credit = 0 if row5_val >= 0 else abs(row5_val)

            # Row 6: Exchange - Updated row number
            row6_val = error_charge_sum
            row6_debit = row6_val if row6_val >= 0 else 0
            row6_credit = 0 if row6_val >= 0 else abs(row6_val)

            # Row 7: Dues and Subscriptions - Updated row number
            row7_val = fee_debit_sum
            row7_debit = row7_val if row7_val >= 0 else 0
            row7_credit = 0 if row7_val >= 0 else abs(row7_val)

            # Row 8: Food Sales (DoorDash Delivery) - Updated row number
            row8_val = delivery_amount
            row8_debit = row8_val if row8_val >= 0 else 0
            row8_credit = 0 if row8_val >= 0 else abs(row8_val)

            # Row 9: Food Sales (DoorDash Pickup) - Updated row number
            row9_val = pickup_amount
            row9_debit = row9_val if row9_val >= 0 else 0
            row9_credit = 0 if row9_val >= 0 else abs(row9_val)

            # Row 10: Sales Tax Payable (Toast) - Updated row number
            row10_val = tax_amount
            row10_debit = row10_val if row10_val >= 0 else 0
            row10_credit = 0 if row10_val >= 0 else abs(row10_val)

            # Row 11: DD Delivery (Not Picked Up) - Updated row number
            row11_val = not_picked_up_subtotal
            row11_debit = 0 if row11_val >= 0 else abs(row11_val)
            row11_credit = row11_val if row11_val >= 0 else 0

            # Row 12: Pickup (Picked Up) - Updated row number
            row12_val = picked_up_subtotal
            row12_debit = 0 if row12_val >= 0 else abs(row12_val)
            row12_credit = row12_val if row12_val >= 0 else 0

            # Row 13: Sales Tax Payable (DoorDash) - Updated row number
            row13_val = subtotal_tax_sum
            row13_debit = 0 if row13_val >= 0 else abs(row13_val)
            row13_credit = row13_val if row13_val >= 0 else 0

            # Row 14: A/R DoorDash (Balancing row) - Updated row number
            row14_val = (row8_val + row9_val + row10_val) - (row11_val + row12_val + row13_val)
            row14_debit = 0 if row14_val >= 0 else abs(row14_val)
            row14_credit = row14_val if row14_val >= 0 else 0

            # Create the journal entry number
            je_number = f"DD{today}{je_suffix}{je_counter}"

            # Format date for output
            formatted_date = date
            try:
                if '-' in date:
                    date_obj = datetime.strptime(date, '%Y-%m-%d')
                    formatted_date = date_obj.strftime('%m/%d/%Y')
            except ValueError:
                pass  # Keep the original format if parsing fails

            # Base JE comment
            je_comment = f"Deposited {deposit_date} // DoorDash Orders {formatted_date}"

            # Create the journal entries for all rows
            journal_entry = [
                # Row 1: A/R DoorDash
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'A/R DoorDash',
                    'Debit': f"{row1_debit:.2f}",
                    'Credit': f"{row1_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Subtotal + Subtotal Tax"
                },
                # Row 2: Commission to Fee - Doordash (Picked Up)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'Doordash Pickup Commission',
                    'Debit': f"{row2_debit:.2f}",
                    'Credit': f"{row2_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Pick-Up Commission + Marketing Fees - Merchant Discount"
                },
                # Row 3: Commission to Fee - Doordash (Not Picked Up)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'Doordash Delivery Commission',
                    'Debit': f"{row3_debit:.2f}",
                    'Credit': f"{row3_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Delivery Commission + Marketing Fees - Merchant Discount"
                },
                # Row 4: NEW Combined Doordash Discount
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'Doordash Discount',
                    'Debit': f"{row4_debit:.2f}",
                    'Credit': f"{row4_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Merchant funded subtotal discount amount"
                },
                # Row 5: A/R DoorDash (Updated row number)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'A/R DoorDash',
                    'Debit': f"{row5_debit:.2f}",
                    'Credit': f"{row5_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Net Payout to be deposited"
                },
                # Row 6: Exchange (Updated row number)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'Refunds',
                    'Debit': f"{row6_debit:.2f}",
                    'Credit': f"{row6_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Error Charge (Refund)"
                },
                # Row 7: Dues and Subscriptions (Updated row number)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'Dues and Subscriptions',
                    'Debit': f"{row7_debit:.2f}",
                    'Credit': f"{row7_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Tablet Fee"
                },
                # Row 8: Food Sales (DoorDash Delivery) (Updated row number)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'UberEats Sales',
                    'Debit': f"{row8_debit:.2f}",
                    'Credit': f"{row8_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Toast Delivery Amount"
                },
                # Row 9: Food Sales (DoorDash Pickup) (Updated row number)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'UberEats Sales',
                    'Debit': f"{row9_debit:.2f}",
                    'Credit': f"{row9_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Toast Pickup Amount"
                },
                # Row 10: Sales Tax Payable (Toast) (Updated row number)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'Sales Tax Payable',
                    'Debit': f"{row10_debit:.2f}",
                    'Credit': f"{row10_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Toast Tax Amount"
                },
                # Row 11: DD Delivery (Updated row number)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'DD Delivery',
                    'Debit': f"{row11_debit:.2f}",
                    'Credit': f"{row11_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Delivery Subtotal"
                },
                # Row 12: Pickup (Updated row number)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'DD Pickup',
                    'Debit': f"{row12_debit:.2f}",
                    'Credit': f"{row12_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Pickup Subtotal"
                },
                # Row 13: Sales Tax Payable (DoorDash) (Updated row number)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'Sales Tax Payable',
                    'Debit': f"{row13_debit:.2f}",
                    'Credit': f"{row13_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Subtotal Tax"
                },
                # Row 14: A/R DoorDash (Balancing row) (Updated row number)
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'Date': formatted_date,
                    'ReversalDate': '',
                    'JEComment': je_comment,
                    'JELocation': je_location,
                    'Account': 'A/R DoorDash',
                    'Debit': f"{row14_debit:.2f}",
                    'Credit': f"{row14_credit:.2f}",
                    'DetailLocation': je_location,
                    'DetailComment': f"Deposited {deposit_date} // Difference in A/R DoorDash in Toast vs DoorDash"
                }
            ]

            journal_entries.extend(journal_entry)
            je_counter += 1

    # We still return these values to maintain compatibility with existing code
    # that expects these return values, but they're not used for exchange refund tracking
    return journal_entries, exchange_totals, exchange_totals_by_month, transaction_dates





def create_summary_journal_entries(journal_entries):
    """
    Create summary journal entries by location and deposit date
    """
    checking_account_mapping = {
        "Carrot Flatiron Love Manhattan Operating LLC": "Checking Carrot Love Manhattan Operating LLC",
        "Carrot Love Bryant Park Operating LLC": "Checking Carrot Love Bryant Park Operating LLC",
        "Carrot Love 600 Lexington LLC": "Checking Carrot Love Lexington 52 LLC"
    }

    summary_entries = []

    # Group journal entries by location and deposit date
    by_location_and_deposit = defaultdict(lambda: defaultdict(list))

    for entry in journal_entries:
        if "Net Payout" in entry.get('DetailComment', ''):
            # Extract deposit date from DetailComment
            detail_comment = entry.get('DetailComment', '')
            if 'Deposited ' in detail_comment:
                deposit_date = detail_comment.split('Deposited ')[1].split(' // ')[0]
                location = entry.get('JELocation', '')
                by_location_and_deposit[location][deposit_date].append(entry)

    # Generate summary entries
    je_counter = 1
    today = datetime.now().strftime("%m%d%Y")

    for location, deposits in by_location_and_deposit.items():
        for deposit_date, entries in deposits.items():
            # Calculate total deposit amount
            total_amount = sum(float(entry.get('Debit', 0)) for entry in entries)

            # Find the date range of orders for this deposit
            order_dates = []
            for entry in entries:
                comment = entry.get('JEComment', '')
                if ' // DoorDash Orders ' in comment:
                    order_date = comment.split(' // DoorDash Orders ')[1]
                    order_dates.append(order_date)

            # Get the min and max dates in mm/dd/yyyy format
            formatted_dates = []
            for date_str in order_dates:
                try:
                    if '-' in date_str:
                        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                    else:
                        date_obj = datetime.strptime(date_str, '%m/%d/%Y')
                    formatted_dates.append(date_obj)
                except ValueError:
                    pass

            date_range_str = ""
            if formatted_dates:
                min_date = min(formatted_dates)
                max_date = max(formatted_dates)
                date_range_str = f"{min_date.strftime('%m/%d/%Y')}-{max_date.strftime('%m/%d/%Y')}"

            if total_amount > 0:
                # Create a unique JE number
                location_suffix = location.split('Love')[1].strip().split()[0][0] if 'Love' in location else 'X'
                je_number = f"DD-DEP-{today}{location_suffix}{je_counter}"

                je_comment = f"DoorDash Deposit {deposit_date}"
                if date_range_str:
                    je_comment += f" // Orders {date_range_str}"

                # Create the summary journal entry
                summary_entry = [
                    # Debit to checking account
                    {
                        'JENumber': je_number,
                        'Type': 'Standard',
                        'Date': deposit_date,
                        'ReversalDate': '',
                        'JEComment': je_comment,
                        'JELocation': location,
                        'Account': checking_account_mapping.get(location, f"Checking {location}"),
                        'Debit': f"{total_amount:.2f}",
                        'Credit': "0.00",
                        'DetailLocation': location,
                        'DetailComment': je_comment
                    },
                    # Credit to A/R DoorDash
                    {
                        'JENumber': je_number,
                        'Type': 'Standard',
                        'Date': deposit_date,
                        'ReversalDate': '',
                        'JEComment': je_comment,
                        'JELocation': location,
                        'Account': 'A/R DoorDash',
                        'Debit': "0.00",
                        'Credit': f"{total_amount:.2f}",
                        'DetailLocation': location,
                        'DetailComment': je_comment
                    }
                ]

                summary_entries.extend(summary_entry)
                je_counter += 1

    return summary_entries


class DoorDashProcessThread(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, doordash_files, toast_files):
        super().__init__()
        self.doordash_files = doordash_files
        self.toast_files = toast_files

    def run(self):
        try:
            self.update_signal.emit("Starting DoorDash transaction processing...")

            # Count files
            self.update_signal.emit(f"Processing {len(self.doordash_files)} DoorDash files and {len(self.toast_files)} Toast files...")

            # Load DoorDash data from all files
            self.update_signal.emit("Loading DoorDash data...")
            doordash_data = []
            for file in self.doordash_files:
                data = load_doordash_data(file)
                doordash_data.extend(data)
                self.update_signal.emit(f"Loaded {len(data)} records from {os.path.basename(file)}")

            self.update_signal.emit(f"Total DoorDash records loaded: {len(doordash_data)}")

            # Process Toast data
            self.update_signal.emit("Processing Toast data...")
            toast_summary = preprocess_toast_data(self.toast_files)

            # Process transactions and create journal entries
            self.update_signal.emit("Creating journal entries...")
            journal_entries, _, _, transaction_dates = process_transactions(doordash_data, toast_summary)
            self.update_signal.emit(f"Generated {len(journal_entries)} journal entries")

            # Report on any isolated fees that were processed
            self.update_signal.emit("Checking for isolated fee transactions...")

            # Create summary journal entries
            self.update_signal.emit("Creating summary journal entries...")
            summary_entries = create_summary_journal_entries(journal_entries)
            self.update_signal.emit(f"Generated {len(summary_entries)} summary journal entries")

            # Determine date range for file naming
            all_dates = set()
            for entry in journal_entries:
                entry_date = entry.get('Date', '')
                if entry_date:
                    try:
                        if '-' in entry_date:
                            date_obj = datetime.strptime(entry_date, '%Y-%m-%d')
                        else:
                            date_obj = datetime.strptime(entry_date, '%m/%d/%Y')
                        all_dates.add(date_obj)
                    except ValueError:
                        pass

            if all_dates:
                start_date = min(all_dates)
                end_date = max(all_dates)
                date_range = f"{start_date.strftime('%m%d%Y')}-{end_date.strftime('%m%d%Y')}"
            else:
                # Fallback to current date if no valid dates found
                today = datetime.now()
                date_range = f"{today.strftime('%m%d%Y')}"

            # Get the downloads folder for output
            downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')

            # Output filename with date range
            output_file = os.path.join(downloads_path, f'DoorDash_Payout_{date_range}.csv')

            # CSV field names
            fieldnames = ['JENumber', 'Type', 'Date', 'ReversalDate', 'JEComment', 'JELocation',
                        'Account', 'Debit', 'Credit', 'DetailLocation', 'DetailComment']

            # Write journal entries to CSV - including summary entries in the same file
            with open(output_file, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(journal_entries)
                writer.writerows(summary_entries)

            self.update_signal.emit(f"Created journal entries file: {output_file}")

            self.finished_signal.emit(True, f"DoorDash processing complete!\nFile saved to:\n{output_file}")

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.update_signal.emit(f"Error: {str(e)}\n{error_details}")
            self.finished_signal.emit(False, f"Error processing files: {str(e)}")


class DoorDashWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.doordash_files = []
        self.toast_files = []

    def initUI(self):
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel("DoorDash Payout Processing", self)
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

        self.setWindowTitle('DoorDash Transaction Processing')
        self.setFixedSize(1000, 738)
        self.center()

    def center(self):
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select All Files", "", "CSV Files (*.csv)"
        )
        if files:
            # Clear previous files
            self.doordash_files = []
            self.toast_files = []

            # Split files by type
            for file in files:
                basename = os.path.basename(file)
                if basename.startswith("Order"):
                    self.toast_files.append(file)
                elif 'financials_detailed_transactions' in basename:
                    self.doordash_files.append(file)
                else:
                    # Try to infer file type
                    try:
                        with open(file, 'r', encoding='utf-8') as f:
                            header = f.readline()
                            if 'DoorDash' in header or 'Store Name' in header:
                                self.doordash_files.append(file)
                            elif 'Dining Options' in header or 'Toast' in header:
                                self.toast_files.append(file)
                            else:
                                # Default to DoorDash if can't determine
                                self.doordash_files.append(file)
                    except:
                        # Default to DoorDash if can't read the file
                        self.doordash_files.append(file)

            # Update file list
            self.file_list.clear()
            for file in self.doordash_files:
                self.file_list.addItem(f"DoorDash file: {os.path.basename(file)}")
            for file in self.toast_files:
                self.file_list.addItem(f"Toast file: {os.path.basename(file)}")

            self.console_output.append(f"Selected {len(files)} files: {len(self.doordash_files)} DoorDash files and {len(self.toast_files)} Toast files")

    def run_processing(self):
        if not self.doordash_files:
            QMessageBox.warning(self, "Error", "Please select at least one DoorDash transaction file.")
            return

        self.console_output.clear()
        self.console_output.append("Starting processing...")
        self.run_button.setEnabled(False)

        self.process_thread = DoorDashProcessThread(self.doordash_files, self.toast_files)
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
Instructions for DoorDash Transaction Processing:

1. Export your DoorDash financial files from the DoorDash merchant portal
   - Go to Reports > Create Report > Financial report > By Date Range > Transactions Breakdown
   - Select date range and download the CSV file(s)
   - Use the file that starts with "financials_detailed_transactions_us"

2. Export your Toast Order Details CSV files
   - "Orders" tab
   - Download the report for the same date range

3. Click "Input Files" to select ALL files at once
   - Files with names starting with "Order" will be treated as Toast files
   - Files with "financials_detailed_transactions" in the name will be treated as DoorDash files

4. Click RUN to process the transactions

5. The processed file will be saved to your Downloads folder:
   - DoorDash_Payout_[DATERANGE].csv - Journal entries
"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Instructions")
        msg_box.setText(instructions)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()
