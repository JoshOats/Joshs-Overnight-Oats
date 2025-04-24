from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
                           QFileDialog, QTextEdit, QMessageBox, QListWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from datetime import datetime, timedelta
import os
from retro_style import RetroWindow, create_retro_central_widget
import calendar


def get_deposit_week(date):
    import pandas as pd
    date = pd.to_datetime(date)
    weekday = date.weekday()

    # Check if this is the last day of the month and if it's a Tuesday
    last_day = pd.Timestamp(date.year, date.month, calendar.monthrange(date.year, date.month)[1])

    # Special case: Month ends on Tuesday
    if last_day.weekday() == 1:  # Tuesday
        last_tuesday = last_day
        prev_tuesday = last_tuesday - timedelta(days=7)

        # If date falls in the special Tuesday-to-Tuesday range
        if prev_tuesday <= date <= last_tuesday:
            deposit_date = last_tuesday + timedelta(days=3)  # Friday of same week
            return prev_tuesday, last_tuesday, deposit_date, ""

    # Regular case: Find Tuesday-Monday period
    if weekday < 1:  # Monday
        tuesday = date - timedelta(days=weekday + 6)
    else:  # Tuesday-Sunday
        tuesday = date - timedelta(days=weekday - 1)

    monday = tuesday + timedelta(days=6)
    deposit_date = tuesday + timedelta(days=10)

    # First day of month special case: If month starts on Wednesday, treat normally
    first_day = pd.Timestamp(date.year, date.month, 1)
    if first_day.weekday() == 2:  # Wednesday
        return tuesday, monday, deposit_date, ""

    # Check for month transition (excluding end-of-month Tuesday case)
    if tuesday.month != monday.month:
        if date.month == tuesday.month:
            return tuesday, monday, deposit_date, "A // "
        else:
            return tuesday, monday, deposit_date, "B // "

    return tuesday, monday, deposit_date, ""


def create_journal_entries(transactions_df, order_files=None):
    import pandas as pd
    import os
    from datetime import datetime

    LOCATION_MAPPING = {
        "Carrot Express -  Lexington Avenue": ["Carrot Love 600 Lexington LLC", "Carrot Love 600 Lexington LLC", "LX", "Checking Carrot Love Lexington 52 LLC"],
        "Carrot Express Bryant Park - West 41st Street": ["Carrot Love Bryant Park Operating LLC", "Carrot Love Bryant Park Operating LLC", "BP", "Checking Carrot Love Bryant Park Operating LLC"],
        "Carrot Express Flatiron - West 23rd Street": ["Carrot Flatiron Love Manhattan Operating LLC", "Carrot Flatiron Love Manhattan Operating LLC", "FI", "Checking Carrot Love Manhattan Operating LLC"]
    }

    # Process order files if provided - with optimization
    grubhub_order_data = {}
    if order_files:
        for file_path in order_files:
            if not os.path.exists(file_path):
                continue

            # Try different encodings
            orders_df = None
            for encoding in ['utf-8', 'latin1']:
                try:
                    # Only read necessary columns
                    orders_df = pd.read_csv(file_path, encoding=encoding,
                                          usecols=['Location', 'Opened', 'Dining Options', 'Amount', 'Tax'])
                    break
                except Exception as e:
                    continue

            if orders_df is None:
                continue

            # Filter to just Grubhub orders
            orders_df = orders_df[orders_df['Dining Options'].str.contains('Grubhub', na=False)]

            # Batch process rows
            for loc_group in orders_df.groupby('Location'):
                location_name = loc_group[0]
                location_df = loc_group[1]

                # Map location
                location = None
                if "Bryant Park" in location_name:
                    location = "Carrot Love Bryant Park Operating LLC"
                elif "Flatiron" in location_name:
                    location = "Carrot Flatiron Love Manhattan Operating LLC"
                elif "Lexington" in location_name:
                    location = "Carrot Love 600 Lexington LLC"
                else:
                    continue

                # Process by date and dining option
                for date_str, date_group in location_df.groupby(location_df['Opened'].str.split(' ').str[0]):
                    try:
                        date_obj = pd.to_datetime(date_str)
                        date_formatted = date_obj.strftime('%m/%d/%Y')

                        key = f"{date_formatted}-{location}"

                        if key not in grubhub_order_data:
                            grubhub_order_data[key] = {
                                'date': date_obj,
                                'location': location,
                                'pickup': 0,
                                'delivery': 0,
                                'tax': 0
                            }

                        # Sum by dining option
                        pickup_mask = date_group['Dining Options'] == 'Grubhub (Takeout)'
                        delivery_mask = date_group['Dining Options'] == 'Grubhub (Delivery)'

                        grubhub_order_data[key]['pickup'] += date_group[pickup_mask]['Amount'].sum()
                        grubhub_order_data[key]['delivery'] += date_group[delivery_mask]['Amount'].sum()
                        grubhub_order_data[key]['tax'] += date_group['Tax'].sum()

                    except Exception as e:
                        continue

    journal_entries = []

    grouped = transactions_df.groupby(['Date', 'Restaurant'])

    for (date, restaurant), group in grouped:
        start_tuesday, end_monday, deposit_date, prefix = get_deposit_week(date)
        je_location, detail_location, suffix, checking_account = LOCATION_MAPPING[restaurant]

        deposit_total = round(group['Restaurant Total'].sum() + (group['Commission'].sum()) + \
                       (group['Processing Fee'].sum()) + group['Targeted Promotion'].sum() + \
                       group['Rewards'].sum(), 2)

        transaction_date = pd.to_datetime(date)
        je_number = f"GH{transaction_date.strftime('%m%d%Y')}-{suffix}"
        date_str = transaction_date.strftime('%m/%d/%Y')
        deposit_date_str = deposit_date.strftime('%m/%d/%Y')

        # For the output file, use transaction date as the date
        output_date = transaction_date

        # Split subtotal into pickup and delivery based on fulfillment type
        pickup_subtotal = round(group[group['Fulfillment Type'] == 'Pick-Up']['Subtotal'].sum(), 2)
        delivery_subtotal = round(group[group['Fulfillment Type'] == 'Self Delivery']['Subtotal'].sum(), 2)

        # Split Commission into pickup and delivery based on fulfillment type
        pickup_commission = round(group[group['Fulfillment Type'] == 'Pick-Up']['Commission'].sum(), 2)
        delivery_commission = round(group[group['Fulfillment Type'] == 'Self Delivery']['Commission'].sum(), 2)

        # Split Processing Fee into pickup and delivery based on fulfillment type
        pickup_processing_fee = round(group[group['Fulfillment Type'] == 'Pick-Up']['Processing Fee'].sum(), 2)
        delivery_processing_fee = round(group[group['Fulfillment Type'] == 'Self Delivery']['Processing Fee'].sum(), 2)

        # Verify that the sum matches the total subtotal
        total_subtotal = round(group['Subtotal'].sum(), 2)
        if abs(pickup_subtotal + delivery_subtotal - total_subtotal) > 0.01:
            remaining_subtotal = total_subtotal - (pickup_subtotal + delivery_subtotal)
            delivery_subtotal += remaining_subtotal

        # Get order data for the second Food Sales split
        order_key = f"{date_str}-{je_location}"
        order_pickup_amount = 0
        order_delivery_amount = 0
        toast_tax = 0

        if order_key in grubhub_order_data:
            order_pickup_amount = round(grubhub_order_data[order_key]['pickup'], 2)
            order_delivery_amount = round(grubhub_order_data[order_key]['delivery'], 2)
            toast_tax = round(grubhub_order_data[order_key]['tax'], 2)

        # If no matching order data, use the proportions from the transactions
        if order_pickup_amount == 0 and order_delivery_amount == 0:
            if total_subtotal > 0:
                ratio_pickup = pickup_subtotal / total_subtotal
                ratio_delivery = delivery_subtotal / total_subtotal
                order_pickup_amount = round(total_subtotal * ratio_pickup, 2)
                order_delivery_amount = round(total_subtotal * ratio_delivery, 2)
            else:
                order_pickup_amount = round(total_subtotal / 2, 2)
                order_delivery_amount = round(total_subtotal / 2, 2)

        # Get delivery fee amount
        delivery_fee = round(group['Delivery Fee'].sum(), 2)

        # Adjust accounts and detail comments lists based on location
        if je_location == "Carrot Love Bryant Park Operating LLC":
            # For Bryant Park, split row 3 into two separate rows (row 3 and row 4)
            # But don't include Delivery Fee Income (which was row 5 in original)
            accounts = ["A/R Grubhub", "A/R GrubHub", "A/R GrubHub", "A/R GrubHub",
                     "Grubhub Delivery Commission", "Grubhub Delivery Commission", "Grubhub Pickup Commission", "Grubhub Pickup Commission",
                     "GrubHub Discount", "Rewards", "GH Takeout", "GH Delivery",
                     "GrubHub Sales", "GrubHub Sales",
                     "Sales Tax Payable", "Sales Tax Payable", "A/R Grubhub", "A/R Grubhub"]

            detail_comments = [
                f"{prefix}Deposited {deposit_date_str} // Restaurant Total + Commission + Processing Fee + Targeted Promotion + Rewards // Cash to be deposited",
                f"{prefix}Deposited {deposit_date_str} // Tips",
                f"{prefix}Deposited {deposit_date_str} // Subtotal + Tax",
                f"{prefix}Deposited {deposit_date_str} // Delivery Fee",
                f"{prefix}Deposited {deposit_date_str} // Commission - Delivery",
                f"{prefix}Deposited {deposit_date_str} // Processing Fee - Delivery",
                f"{prefix}Deposited {deposit_date_str} // Commission - Pickup",
                f"{prefix}Deposited {deposit_date_str} // Processing Fee - Pickup",
                f"{prefix}Deposited {deposit_date_str} // Targeted Promotion",
                f"{prefix}Deposited {deposit_date_str} // Rewards",
                f"{prefix}Deposited {deposit_date_str} // Pickup Subtotal From Grubhub",
                f"{prefix}Deposited {deposit_date_str} // Delivery Subtotal From Grubhub",
                f"{prefix}Deposited {deposit_date_str} // GrubHub Pickup From Toast",
                f"{prefix}Deposited {deposit_date_str} // GrubHub Delivery From Toast",
                f"{prefix}Deposited {deposit_date_str} // Grubhub Sales Tax - Toast",
                f"{prefix}Deposited {deposit_date_str} // Grubhub Sales Tax - Grubhub",
                f"{prefix}Deposited {deposit_date_str} // Tax Difference",
                f"{prefix}Deposited {deposit_date_str} // Sales/Subtotal Difference"
            ]
        else:
            # For other locations, keep the existing structure intact (19 accounts total)
            accounts = ["A/R Grubhub", "A/R GrubHub", "A/R GrubHub", "Delivery Fee Income",
                     "GH Delivery", "Grubhub Delivery Commission", "Grubhub Delivery Commission", "Grubhub Pickup Commission", "Grubhub Pickup Commission",
                     "GrubHub Discount", "Rewards", "GH Takeout", "GH Delivery",
                     "GrubHub Sales", "GrubHub Sales",
                     "Sales Tax Payable", "Sales Tax Payable", "A/R Grubhub", "A/R Grubhub"]

            detail_comments = [
                f"{prefix}Deposited {deposit_date_str} // Restaurant Total + Commission + Processing Fee + Targeted Promotion + Rewards // Cash to be deposited",
                f"{prefix}Deposited {deposit_date_str} // Tips",
                f"{prefix}Deposited {deposit_date_str} // Subtotal + Tax",
                f"{prefix}Deposited {deposit_date_str} // Delivery Fee",
                f"{prefix}Deposited {deposit_date_str} // Delivery Fee Income",
                f"{prefix}Deposited {deposit_date_str} // Commission - Delivery",
                f"{prefix}Deposited {deposit_date_str} // Processing Fee - Delivery",
                f"{prefix}Deposited {deposit_date_str} // Commission - Pickup",
                f"{prefix}Deposited {deposit_date_str} // Processing Fee - Pickup",
                f"{prefix}Deposited {deposit_date_str} // Targeted Promotion",
                f"{prefix}Deposited {deposit_date_str} // Rewards",
                f"{prefix}Deposited {deposit_date_str} // Pickup Subtotal From Grubhub",
                f"{prefix}Deposited {deposit_date_str} // Delivery Subtotal From Grubhub + Delivery Fee",
                f"{prefix}Deposited {deposit_date_str} // GrubHub Pickup From Toast",
                f"{prefix}Deposited {deposit_date_str} // GrubHub Delivery From Toast",
                f"{prefix}Deposited {deposit_date_str} // Grubhub Sales Tax - Toast",
                f"{prefix}Deposited {deposit_date_str} // Grubhub Sales Tax - Grubhub",
                f"{prefix}Deposited {deposit_date_str} // Tax Difference",
                f"{prefix}Deposited {deposit_date_str} // Sales/Subtotal Difference"
            ]

        # Calculate commission debits and credits
        commission_delivery_debit = abs(delivery_commission) if delivery_commission < 0 else 0
        commission_delivery_credit = delivery_commission if delivery_commission > 0 else 0

        commission_pickup_debit = abs(pickup_commission) if pickup_commission < 0 else 0
        commission_pickup_credit = pickup_commission if pickup_commission > 0 else 0

        # Calculate processing fee debits and credits
        processing_delivery_debit = abs(delivery_processing_fee) if delivery_processing_fee < 0 else 0
        processing_delivery_credit = delivery_processing_fee if delivery_processing_fee > 0 else 0

        processing_pickup_debit = abs(pickup_processing_fee) if pickup_processing_fee < 0 else 0
        processing_pickup_credit = pickup_processing_fee if pickup_processing_fee > 0 else 0

        # Split the last Food Sales into Pickup and Delivery based on order data
        food_sales_pickup = order_pickup_amount
        food_sales_delivery = order_delivery_amount

        # Set up credits and debits based on location
        if je_location == "Carrot Love Bryant Park Operating LLC":
            # For Bryant Park, modify how Row 3 is handled by separating Subtotal+Tax from Delivery Fee
            credits = [
                0,                                              # Cash to be deposited
                round(group['Tip'].sum(), 2),                   # Tips
                round(group['Subtotal'].sum() + group['Tax Fee'].sum(), 2),  # Subtotal + Tax (row 3)
                delivery_fee,                                   # Delivery Fee (new row 4)
                commission_delivery_credit,                     # Commission - Delivery
                processing_delivery_credit,                     # Processing Fee - Delivery
                commission_pickup_credit,                       # Commission - Pickup
                processing_pickup_credit,                       # Processing Fee - Pickup
                0,                                              # Targeted Promotion
                0,                                              # Rewards
                pickup_subtotal,                                # Pickup Subtotal From Grubhub
                delivery_subtotal,                              # Delivery Subtotal From Grubhub
                0,                                              # GrubHub Pickup From Toast
                0                                               # GrubHub Delivery From Toast
            ]

            debits = [
                deposit_total,                                  # Cash to be deposited
                0,                                              # Tips
                0,                                              # Subtotal + Tax
                0,                                              # Delivery Fee (new row)
                commission_delivery_debit,                      # Commission - Delivery
                processing_delivery_debit,                      # Processing Fee - Delivery
                commission_pickup_debit,                        # Commission - Pickup
                processing_pickup_debit,                        # Processing Fee - Pickup
                round(abs(group['Targeted Promotion'].sum()), 2),  # Targeted Promotion
                round(abs(group['Rewards'].sum()), 2),          # Rewards
                0,                                              # Pickup Subtotal From Grubhub
                0,                                              # Delivery Subtotal From Grubhub
                food_sales_pickup,                              # GrubHub Pickup From Toast
                food_sales_delivery - delivery_fee               # GrubHub Delivery From Toast (adjusted for Bryant Park)
            ]
        else:
            # For other locations, keep the exact original structure with 19 rows
            credits = [
                0,                                              # Cash to be deposited
                round(group['Tip'].sum(), 2),                   # Tips
                round(group['Subtotal'].sum() + group['Tax Fee'].sum(), 2),  # Subtotal + Tax
                delivery_fee,                                   # Delivery Fee
                0,                                              # Delivery Fee Income
                commission_delivery_credit,                     # Commission - Delivery
                processing_delivery_credit,                     # Processing Fee - Delivery
                commission_pickup_credit,                       # Commission - Pickup
                processing_pickup_credit,                       # Processing Fee - Pickup
                0,                                              # Targeted Promotion
                0,                                              # Rewards
                pickup_subtotal,                                # Pickup Subtotal From Grubhub
                delivery_subtotal + delivery_fee,               # Delivery Subtotal From Grubhub + Delivery Fee
                0,                                              # GrubHub Pickup From Toast
                0                                               # GrubHub Delivery From Toast
            ]

            debits = [
                deposit_total,                                  # Cash to be deposited
                0,                                              # Tips
                0,                                              # Subtotal + Tax
                0,                                              # Delivery Fee
                round(group['Delivery Fee'].sum(), 2),          # Delivery Fee Income
                commission_delivery_debit,                      # Commission - Delivery
                processing_delivery_debit,                      # Processing Fee - Delivery
                commission_pickup_debit,                        # Commission - Pickup
                processing_pickup_debit,                        # Processing Fee - Pickup
                round(abs(group['Targeted Promotion'].sum()), 2),  # Targeted Promotion
                round(abs(group['Rewards'].sum()), 2),          # Rewards
                0,                                              # Pickup Subtotal From Grubhub
                0,                                              # Delivery Subtotal From Grubhub
                food_sales_pickup,                              # GrubHub Pickup From Toast
                food_sales_delivery                             # GrubHub Delivery From Toast
            ]

        # Get Grubhub tax amount
        grubhub_tax = round(group['Tax Fee'].sum(), 2)

        # Calculate tax difference
        tax_difference = toast_tax - grubhub_tax

        # Calculate sales/subtotal difference
        delivery_fee_grubhub = round(group['Delivery Fee'].sum(), 2)
        if je_location == "Carrot Love Bryant Park Operating LLC":
            sales_diff = food_sales_pickup - pickup_subtotal + food_sales_delivery - delivery_subtotal - delivery_fee_grubhub
        else:
            sales_diff = food_sales_pickup - pickup_subtotal + food_sales_delivery + debits[4] - delivery_subtotal - delivery_fee_grubhub

        # Add the new rows to debits and credits
        debits.extend([
            toast_tax,                                           # Sales Tax Payable (Toast)
            0,                                                   # Sales Tax Payable (Grubhub)
            0 if tax_difference >= 0 else abs(tax_difference),   # Tax Difference
            0 if sales_diff >= 0 else abs(sales_diff)            # Sales/Subtotal Difference
        ])

        credits.extend([
            0,                                                  # Sales Tax Payable (Toast)
            grubhub_tax,                                        # Sales Tax Payable (Grubhub)
            tax_difference if tax_difference >= 0 else 0,       # Tax Difference
            sales_diff if sales_diff >= 0 else 0                # Sales/Subtotal Difference
        ])

        # Handle negative values
        for i in range(len(accounts)):  # Use the length of accounts to be safe
            if i < len(debits) and i < len(credits):  # Ensure we don't go out of range
                if debits[i] < 0:
                    credits[i] = abs(debits[i])
                    debits[i] = 0
                if credits[i] < 0:
                    debits[i] = abs(credits[i])
                    credits[i] = 0

        # For non-Bryant Park locations, don't skip any rows
        # For Bryant Park, we include all rows (since we've manually constructed the arrays with 18 items)
        row_indices = range(len(accounts))

        for i in row_indices:
            if i < len(accounts) and i < len(detail_comments) and i < len(debits) and i < len(credits):
                journal_entries.append({
                    'JENumber': je_number,
                    'Type': "",
                    'DetailComment': detail_comments[i],
                    'Reversal Date': None,
                    'JEComment': f"{prefix}Deposited {deposit_date_str} // Orders on {date_str}",
                    'JELocation': je_location,
                    'Account': accounts[i],
                    'Debit': debits[i],
                    'Credit': credits[i],
                    'DetailLocation': detail_location,
                    'Date': output_date  # Changed to use transaction date
                })

    return pd.DataFrame(journal_entries)

def create_deposit_journal_entries(transactions_df):
    import pandas as pd
    from datetime import datetime


    LOCATION_MAPPING = {
        "Carrot Express -  Lexington Avenue": ["Carrot Love 600 Lexington LLC", "Carrot Love 600 Lexington LLC", "LX", "Checking Carrot Love Lexington 52 LLC"],
        "Carrot Express Bryant Park - West 41st Street": ["Carrot Love Bryant Park Operating LLC", "Carrot Love Bryant Park Operating LLC", "BP", "Checking Carrot Love Bryant Park Operating LLC"],
        "Carrot Express Flatiron - West 23rd Street": ["Carrot Flatiron Love Manhattan Operating LLC", "Carrot Flatiron Love Manhattan Operating LLC", "FI", "Checking Carrot Love Manhattan Operating LLC"]
    }

    deposit_entries = []

    # Group transactions by deposit date
    transactions_df['deposit_date'] = None
    transactions_df['deposit_date_str'] = None
    transactions_df['start_tuesday'] = None
    transactions_df['end_monday'] = None

    # Calculate deposit dates for all transactions
    for idx, row in transactions_df.iterrows():
        start_tuesday, end_monday, deposit_date, prefix = get_deposit_week(row['Date'])
        transactions_df.at[idx, 'deposit_date'] = deposit_date
        transactions_df.at[idx, 'deposit_date_str'] = deposit_date.strftime('%m/%d/%Y')
        transactions_df.at[idx, 'start_tuesday'] = start_tuesday
        transactions_df.at[idx, 'end_monday'] = end_monday

    # Group by deposit date and restaurant
    deposit_groups = transactions_df.groupby(['deposit_date', 'Restaurant'])

    for (deposit_date, restaurant), group in deposit_groups:
        je_location, detail_location, suffix, checking_account = LOCATION_MAPPING[restaurant]

        # Calculate deposit total
        deposit_total = round(group['Restaurant Total'].sum() + (group['Commission'].sum()) + \
                       (group['Processing Fee'].sum()) + group['Targeted Promotion'].sum() + \
                       group['Rewards'].sum(), 2)

        # Get date range for comment
        min_date = group['Date'].min().strftime('%m/%d/%Y')
        max_date = group['Date'].max().strftime('%m/%d/%Y')
        date_range = f"{min_date}"
        if min_date != max_date:
            date_range += f"-{max_date}"

        # Format deposit date
        deposit_date_str = deposit_date.strftime('%m/%d/%Y')

        # Create JE number
        je_number = f"GH-DEP-{deposit_date.strftime('%m%d%Y')}-{suffix}"

        # Calculate prefix
        _, _, _, prefix = get_deposit_week(group['Date'].iloc[0])

        # Create comment
        comment = f"{prefix}Deposited {deposit_date_str} // Orders on {date_range}"

        # Create the two entries
        deposit_entries.append({
            'JENumber': je_number,
            'Type': "",
            'DetailComment': comment,
            'Reversal Date': None,
            'JEComment': comment,
            'JELocation': je_location,
            'Account': checking_account,
            'Debit': deposit_total,
            'Credit': 0,
            'DetailLocation': detail_location,
            'Date': deposit_date
        })

        deposit_entries.append({
            'JENumber': je_number,
            'Type': "",
            'DetailComment': comment,
            'Reversal Date': None,
            'JEComment': comment,
            'JELocation': je_location,
            'Account': "A/R GrubHub",
            'Debit': 0,
            'Credit': deposit_total,
            'DetailLocation': detail_location,
            'Date': deposit_date
        })

    return pd.DataFrame(deposit_entries)


def create_tip_adjustment_entries(transactions_df, order_files=None):
    import pandas as pd
    import os
    from datetime import datetime

    LOCATION_MAPPING = {
        "Carrot Express -  Lexington Avenue": ["Carrot Love 600 Lexington LLC", "Carrot Love 600 Lexington LLC", "LX", "Checking Carrot Love Lexington 52 LLC"],
        "Carrot Express Bryant Park - West 41st Street": ["Carrot Love Bryant Park Operating LLC", "Carrot Love Bryant Park Operating LLC", "BP", "Checking Carrot Love Bryant Park Operating LLC"],
        "Carrot Express Flatiron - West 23rd Street": ["Carrot Flatiron Love Manhattan Operating LLC", "Carrot Flatiron Love Manhattan Operating LLC", "FI", "Checking Carrot Love Manhattan Operating LLC"]
    }

    # Process order files to get tip data
    grubhub_order_tips = {}
    if order_files:
        for file_path in order_files:
            if not os.path.exists(file_path):
                continue

            # Try different encodings
            orders_df = None
            for encoding in ['utf-8', 'latin1']:
                try:
                    orders_df = pd.read_csv(file_path, encoding=encoding,
                                          usecols=['Location', 'Opened', 'Dining Options', 'Tip'])
                    break
                except Exception as e:
                    continue

            if orders_df is None:
                continue

            # Filter to just Grubhub orders
            orders_df = orders_df[orders_df['Dining Options'].str.contains('Grubhub', na=False)]

            # Process by location and date
            for loc_group in orders_df.groupby('Location'):
                location_name = loc_group[0]
                location_df = loc_group[1]

                # Map location
                location = None
                if "Bryant Park" in location_name:
                    location = "Carrot Love Bryant Park Operating LLC"
                elif "Flatiron" in location_name:
                    location = "Carrot Flatiron Love Manhattan Operating LLC"
                elif "Lexington" in location_name:
                    location = "Carrot Love 600 Lexington LLC"
                else:
                    continue

                # Process by date
                for date_str, date_group in location_df.groupby(location_df['Opened'].str.split(' ').str[0]):
                    try:
                        date_obj = pd.to_datetime(date_str)
                        date_formatted = date_obj.strftime('%m/%d/%Y')

                        key = f"{date_formatted}-{location}"

                        if key not in grubhub_order_tips:
                            grubhub_order_tips[key] = {
                                'date': date_obj,
                                'location': location,
                                'delivery_tips': 0,
                                'pickup_tips': 0
                            }

                        # Sum tips by dining option
                        delivery_tips = date_group[date_group['Dining Options'] == 'Grubhub (Delivery)']['Tip'].sum()
                        pickup_tips = date_group[date_group['Dining Options'] == 'Grubhub (Takeout)']['Tip'].sum()

                        grubhub_order_tips[key]['delivery_tips'] += round(delivery_tips, 2)
                        grubhub_order_tips[key]['pickup_tips'] += round(pickup_tips, 2)

                    except Exception as e:
                        continue

    # Create tip adjustment entries
    tip_entries = []
    grouped = transactions_df.groupby(['Date', 'Restaurant'])

    for (date, restaurant), group in grouped:
        je_location, detail_location, suffix, _ = LOCATION_MAPPING[restaurant]

        transaction_date = pd.to_datetime(date)
        je_number = f"GH-T-{transaction_date.strftime('%m%d%Y')}-{suffix}"
        date_str = transaction_date.strftime('%m/%d/%Y')

        # Get Toast tip amounts
        order_key = f"{date_str}-{je_location}"
        toast_delivery_tips = 0
        toast_pickup_tips = 0

        if order_key in grubhub_order_tips:
            toast_delivery_tips = grubhub_order_tips[order_key]['delivery_tips']
            toast_pickup_tips = grubhub_order_tips[order_key]['pickup_tips']

        # Get Grubhub tip amounts
        grubhub_delivery_tips = round(group[group['Fulfillment Type'] == 'Self Delivery']['Tip'].sum(), 2)
        grubhub_pickup_tips = round(group[group['Fulfillment Type'] == 'Pick-Up']['Tip'].sum(), 2)

        # Calculate tip difference
        tip_difference = (toast_delivery_tips + toast_pickup_tips) - (grubhub_delivery_tips + grubhub_pickup_tips)

        # Only create entry if there's a difference
        if abs(tip_difference) > 0.01:  # Using 0.01 to account for rounding
            accounts = [
                "Payable Delivery Tips",
                "Payable Delivery Tips",
                "Employee Tips Payable",
                "Employee Tips Payable",
                "A/R Grubhub"
            ]

            detail_comments = [
                "Delivery Tips - Toast",
                "Delivery Tips - Grubhub",
                "Pickup Tips - Toast",
                "Pickup Tips - Grubhub",
                "Tips adjustment due to difference"
            ]

            debits = [
                toast_delivery_tips,     # Delivery Tips - Toast
                0,                       # Delivery Tips - Grubhub
                toast_pickup_tips,       # Pickup Tips - Toast
                0,                       # Pickup Tips - Grubhub
                0 if tip_difference >= 0 else abs(tip_difference)  # Tips adjustment
            ]

            credits = [
                0,                       # Delivery Tips - Toast
                grubhub_delivery_tips,   # Delivery Tips - Grubhub
                0,                       # Pickup Tips - Toast
                grubhub_pickup_tips,     # Pickup Tips - Grubhub
                tip_difference if tip_difference >= 0 else 0  # Tips adjustment
            ]

            # Create the 5 rows for this journal entry
            for i in range(5):
                tip_entries.append({
                    'JENumber': je_number,
                    'Type': "",
                    'DetailComment': detail_comments[i],
                    'Reversal Date': None,
                    'JEComment': "Grubhub Tip Adjustment due to Refunded/Cancelled Orders",
                    'JELocation': je_location,
                    'Account': accounts[i],
                    'Debit': debits[i],
                    'Credit': credits[i],
                    'DetailLocation': detail_location,
                    'Date': transaction_date
                })

    return pd.DataFrame(tip_entries)

class GrubHubProcessThread(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, input_file, order_files=None):
        super().__init__()
        self.input_file = input_file
        self.order_files = order_files or []

    def run(self):
        import pandas as pd
        try:
            self.update_signal.emit("Starting GrubHub transaction processing...")

            # Read input file with various encodings to handle potential issues
            encodings_to_try = ['utf-8', 'latin1', 'cp1252', 'ISO-8859-1']
            df = None
            for encoding in encodings_to_try:
                try:
                    self.update_signal.emit(f"Trying to read file with {encoding} encoding...")
                    df = pd.read_csv(self.input_file, encoding=encoding)
                    self.update_signal.emit(f"Successfully read file with {encoding} encoding")
                    break
                except UnicodeDecodeError:
                    continue

            if df is None:
                raise Exception("Could not read the file with any encoding. The file might be corrupted.")

            df['Date'] = pd.to_datetime(df['Date'])

            # Report number of order files being processed
            self.update_signal.emit(f"Processing {len(self.order_files)} Order files...")

            # Create regular journal entries
            je_df = create_journal_entries(df, self.order_files)

            # Create deposit journal entries
            deposit_je_df = create_deposit_journal_entries(df)

            # Create tip adjustment entries
            tip_je_df = create_tip_adjustment_entries(df, self.order_files)

            # Combine the entries
            combined_df = pd.concat([je_df, tip_je_df, deposit_je_df])

            # Add row numbers for sorting
            combined_df['row_num'] = combined_df.groupby('JENumber').cumcount()
            combined_df = combined_df.sort_values(['JENumber', 'row_num'])
            combined_df = combined_df.drop('row_num', axis=1)

            # Create output directory in Downloads
            downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
            today = datetime.now().strftime("%m%d%Y")
            output_file = os.path.join(downloads_path, f"GrubHub_JE_{today}.csv")

            # Save to CSV
            combined_df.to_csv(output_file, index=False, float_format='%.2f')

            success_msg = f"Successfully processed GrubHub transactions!\nOutput saved to: {output_file}"
            self.finished_signal.emit(True, success_msg)

        except Exception as e:
            error_msg = f"Error processing file: {str(e)}"
            self.finished_signal.emit(False, error_msg)


class GrubHubWindow(RetroWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_file = None
        self.order_files = []

    def initUI(self):
        central_widget = create_retro_central_widget(self)
        layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel("GrubHub Payout Processing", self)
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

        self.setWindowTitle('GrubHub Transaction Processing')
        self.setFixedSize(1000, 738)
        self.center()

    def center(self):
        from PyQt5.QtWidgets import QApplication
        frame_geometry = self.frameGeometry()
        center_point = QApplication.desktop().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select All Files", "", "CSV Files (*.csv)"
        )
        if files:
            # Split files by type
            self.order_files = []
            for file in files:
                if os.path.basename(file).startswith("Order"):
                    self.order_files.append(file)
                else:
                    self.selected_file = file

            # Update file list
            self.file_list.clear()
            if self.selected_file:
                self.file_list.addItem(f"GrubHub file: {os.path.basename(self.selected_file)}")
            for file in self.order_files:
                self.file_list.addItem(f"Order file: {os.path.basename(file)}")

            self.console_output.append(f"Selected {len(files)} files: 1 GrubHub file and {len(self.order_files)} Order files")

    def run_processing(self):
        if not self.selected_file:
            QMessageBox.warning(self, "Error", "Please select a GrubHub transaction file.")
            return

        self.console_output.clear()
        self.console_output.append("Starting processing...")
        self.run_button.setEnabled(False)

        self.process_thread = GrubHubProcessThread(self.selected_file, self.order_files)
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
Instructions for GrubHub Transaction Processing:

1. Go to GrubHub -> Financials -> Reports -> Request CSV Report -> Choose all locations & Date Range
2. Also have your OrderDetails CSV files from Toast ready
3. Click "Select All Files" to choose ALL files at once (both GrubHub and Order files)
   - Files with names starting with "Order" will be treated as Order files
   - Any other CSV file will be treated as the GrubHub transaction file
4. Click RUN to process the transactions
5. The processed file will be saved to your Downloads folder

The output file will be named "GrubHub_JE_MMDDYYYY.csv"
"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Instructions")
        msg_box.setText(instructions)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()
