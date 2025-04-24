import pandas as pd
import os
from datetime import datetime, timedelta
import glob
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

# Location mapping dictionary stays the same
LOCATION_MAPPING = {
    'Carrot Express (Aventura)': {'toast': 'Aventura (Miami Gardens)', 'je': 'Carrot Aventura Love LLC (Aventura)', 'checking': 'Checking Carrot Love LLC'},
    'Carrot Express (Coral Gables)': {'toast': 'Coral Gables', 'je': 'Carrot Coral GablesLove LLC (Coral Gabes)', 'checking': 'Checking Carrot Love LLC'},
    'Carrot Express (Downtown)': {'toast': 'Downtown', 'je': 'Carrot Downtown Love Two LLC', 'checking': 'Checking Carrot Love Two LLC'},
    'Carrot Express (Miami Shores)': {'toast': 'Miami Shores', 'je': 'Carrot Express Miami Shores LLC', 'checking': 'Checking Carrot Express Miami Shores LLC'},
    'Carrot Express (Midtown)': {'toast': 'Midtown', 'je': 'Carrot Express Midtown LLC', 'checking': 'Checking Carrot Express Midtown LLC'},
    'Carrot Express (Flatiron)': {'toast': 'Flatiron', 'je': 'Carrot Flatiron Love Manhattan Operating LLC', 'checking': 'Checking Carrot Love Manhattan Operating LLC'},
    'Carrot Express (Lexington)': {'toast': 'Lexington', 'je': 'Carrot Love 600 Lexington LLC', 'checking': 'Checking Carrot Love Lexington 52 LLC'},
    'Carrot Express (Aventura Mall)': {'toast': 'Aventura Mall', 'je': 'Carrot Love Aventura Mall Operating LLC', 'checking': 'Checking Carrot Love Aventura Mall Operating LLC'},
    'Carrot Express (Brickell)': {'toast': 'Brickell', 'je': 'Carrot Love Brickell Operating LLC', 'checking': 'Checking Carrot Love Brickell Operating LLC'},
    'Carrot Express (Bryant Park)': {'toast': 'Bryant Park', 'je': 'Carrot Love Bryant Park Operating LLC', 'checking': 'Checking Carrot Love Bryant Park Operating LLC'},
    'Carrot Express (Doral)': {'toast': 'Doral', 'je': 'Carrot Love City Place Doral Operating LLC', 'checking': 'Checking Carrot Love City Place Doral Operating LLC'},
    'Carrot Express (Coconut Creek)': {'toast': 'Coconut Creek', 'je': 'Carrot Love Coconut Creek Operating LLC', 'checking': 'Checking Carrot Love Coconut Creek Operating LLC'},
    'Carrot Express (Coconut Grove)': {'toast': 'Coconut Grove', 'je': 'Carrot Love Coconut Grove Operating LLC', 'checking': 'Checking Carrot Love Coconut Grove Operating LLC'},
    'Carrot Express (Dadeland)': {'toast': 'Dadeland', 'je': 'Carrot Love Dadeland Operating LLC', 'checking': 'Checking Carrot Love Dadeland Operating LLC'},
    'Carrot Express (Hollywood)': {'toast': 'Hollywood', 'je': 'Carrot Love Hollywood Operating LLC', 'checking': 'Checking Carrot Love Hollywood Operating LLC'},
    'Carrot Express (Las Olas)': {'toast': 'Las Olas', 'je': 'Carrot Love Las Olas Operating LLC', 'checking': 'Checking Carrot Love Las Olas Operating LLC'},
    'Carrot Express (Boca)': {'toast': 'Boca Palmetto Park', 'je': 'Carrot Love Palmetto Park Operating LLC', 'checking': 'Checking Carrot Love Palmetto Park Operating LLC'},
    'Carrot Express (Pembroke Pines)': {'toast': 'Pembroke Pines', 'je': 'Carrot Love Pembroke Pines Operating LLC', 'checking': 'Checking Carrot Love Pembroke Pines Operating LLC'},
    'Carrot Express (Plantation Walk)': {'toast': 'Plantation', 'je': 'Carrot Love Plantation Operating LLC', 'checking': 'Checking Carrot Love Plantation Operating LLC'},
    'Carrot Express (River Landing)': {'toast': 'River Landing', 'je': 'Carrot Love River Lading Operating LLC', 'checking': 'Checking Carrot Love River Landing LLC'},
    'Carrot Express (South Miami)': {'toast': 'South Miami (Sunset)', 'je': 'Carrot Love Sunset Operating LLC', 'checking': 'Checking Carrot Love Sunset Operating LLC'},
    'Carrot Express (West Boca)': {'toast': 'West Boca', 'je': 'Carrot Love West Boca Operating LLC', 'checking': 'Checking Carrot Love West Boca Operating LLC'},
    'Carrot Express (Miami Beach)': {'toast': 'North Beach', 'je': 'Carrot North Beach Love LL (North Beach)', 'checking': 'Checking Carrot Love LLC'},
    'Carrot Express (South Beach)': {'toast': 'South Beach', 'je': 'Carrot Sobe Love South Florida Operating C LLC', 'checking': 'Checking Carrot Love South Florida Operating C LLC'}
}


def is_proper_date_range(payout_date):
    """
    Check if a payout date results in the proper Monday-Sunday order date range.
    Returns True if the range is correct, False if it needs adjustment.
    """
    end_date = payout_date - timedelta(days=1)  # Day before payout
    start_date = end_date - timedelta(days=6)   # 6 days before end_date

    # Check if start_date is a Monday (weekday() == 0) and end_date is a Sunday (weekday() == 6)
    return start_date.weekday() == 0 and end_date.weekday() == 6


def read_toast_files(toast_files):
    """
    Read and process Toast files directly from a list of file paths

    Args:
        toast_files (list): List of paths to Toast CSV files

    Returns:
        DataFrame: Concatenated DataFrame from all Toast files
    """
    toast_dfs = []
    for file in toast_files:
        df = pd.read_csv(file, encoding='cp1252')
        df['Opened'] = pd.to_datetime(df['Opened'], format='mixed')
        toast_dfs.append(df)

    return pd.concat(toast_dfs, ignore_index=True) if toast_dfs else pd.DataFrame()

def read_uber_file(uber_file):
    """
    Read and process the Uber file directly from its file path

    Args:
        uber_file (str): Path to Uber CSV file

    Returns:
        DataFrame: Processed Uber DataFrame
    """
    # Read the file skipping the first row since headers start on second row
    df = pd.read_csv(uber_file, skiprows=1, encoding='utf-8')
    df['Payout Date'] = pd.to_datetime(df['Payout Date'])
    df['Order Date'] = pd.to_datetime(df['Order Date'])
    return df

def round_to_cents(value):
    if pd.isna(value) or value == '':
        return 0.00
    try:
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').replace(' ', '')
        float_val = float(value)
        return float(Decimal(str(float_val)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
    except (ValueError, InvalidOperation) as e:
        print(f"Warning: Could not convert value '{value}' (type: {type(value)}) to decimal: {str(e)}")
        return 0.00

def get_date_range(payout_date):
    # End date is 1 day before payout date (Sunday)
    end_date = payout_date - timedelta(days=1)
    # Start date is 6 days before end date (previous Monday)
    start_date = end_date - timedelta(days=6)
    return start_date, end_date

def create_journal_entries(toast_df, uber_df):
    journal_entries = []
    journal_totals = {}

    # Special locations list
    special_locations = [
        'Carrot Express (Lexington)',
        'Carrot Express (Flatiron)',
        'Carrot Express (Bryant Park)',
        'Carrot Express (South Beach)',
        'Carrot Express (Miami Beach)'
    ]

    # ADD THIS NEW CODE BLOCK HERE - before grouping
    # Handle empty Payout Dates
    missing_payout_mask = uber_df['Payout Date'].isna()
    if missing_payout_mask.any():
        # Define refund-related statuses
        refund_statuses = ["Refund", "Refund Disputed", "Unfulfilled"]

        # Process each location separately
        for location, location_group in uber_df[missing_payout_mask].groupby('Store Name'):
            # Separate regular orders from refund/unfulfilled orders
            regular_orders = location_group[~location_group['Order Status'].isin(refund_statuses)]
            special_orders = location_group[location_group['Order Status'].isin(refund_statuses)]

            if not regular_orders.empty:
                # Get the latest order date for regular orders
                latest_order_date = regular_orders['Order Date'].max()

                # Calculate payout date for regular orders
                for idx in regular_orders.index:
                    order_date = regular_orders.loc[idx, 'Order Date']
                    # Find the next Monday after the order date
                    days_until_monday = (7 - order_date.weekday()) % 7
                    next_monday = order_date + timedelta(days=days_until_monday)
                    # Payout date is the following Monday
                    calculated_payout_date = next_monday + timedelta(days=7)
                    # Update the DataFrame
                    uber_df.loc[idx, 'Payout Date'] = calculated_payout_date

                # Update all special orders to use the same payout date and latest order date
                if not special_orders.empty:
                    calculated_payout_date = uber_df.loc[regular_orders.index[0], 'Payout Date']
                    for idx in special_orders.index:
                        uber_df.loc[idx, 'Payout Date'] = calculated_payout_date
                        uber_df.loc[idx, 'Order Date'] = latest_order_date
                        print(f"Updated refund/unfulfilled order for {location}: using Payout Date {calculated_payout_date.date()} and Order Date {latest_order_date.date()}")

    # First group by Store Name and Payout Date
    payout_groups = uber_df.groupby(['Store Name', 'Payout Date'])

    # Counter for journal entries per deposit date and order date
    je_counters = {}

    for (uber_location, payout_date), payout_group in payout_groups:
        if uber_location not in LOCATION_MAPPING:
            continue

        # Check if payout date needs adjustment
        adjusted_payout_date = payout_date
        if not is_proper_date_range(payout_date):
            adjusted_payout_date = payout_date - timedelta(days=1)
            print(f"Warning: Adjusted payout date for {uber_location} from {payout_date.date()} to {adjusted_payout_date.date()} to maintain Monday-Sunday order range")

        # Calculate relevant dates using adjusted payout date
        deposit_date = adjusted_payout_date + timedelta(days=1)
        start_date, end_date = get_date_range(adjusted_payout_date)

        # Create date range string for comment
        date_range_str = f"{start_date.strftime('%m/%d/%y')}-{end_date.strftime('%m/%d/%y')}"

        # Copy the group to avoid modifying the original DataFrame
        group_copy = payout_group.copy()

        # Reassign out-of-range order dates to the start_date before grouping
        for idx, row in group_copy.iterrows():
            order_date = row['Order Date']
            if order_date.date() < start_date.date() or order_date.date() > end_date.date():
                group_copy.at[idx, 'Order Date'] = start_date
                print(f"Warning: Order date {order_date.date()} is outside expected range {date_range_str} for {uber_location}. Using {start_date.date()}.")

        # First calculate sales by order date to identify dates with zero sales
        sales_by_date = {}
        for order_date, date_group in group_copy.groupby('Order Date'):
            sales_by_date[order_date] = round_to_cents(date_group['Sales (excl. tax)'].sum())

        # Find latest date with non-zero sales for reassignment
        latest_date_with_sales = None
        for date in sorted(sales_by_date.keys(), reverse=True):
            if sales_by_date[date] > 0:
                latest_date_with_sales = date
                break

        # If we found no dates with sales, use the latest date
        if latest_date_with_sales is None and sales_by_date:
            latest_date_with_sales = max(sales_by_date.keys())

        # Create a new DataFrame to store the reassigned orders
        reassigned_group = group_copy.copy()

        # Reassign order dates with zero sales to the latest date with sales
        if latest_date_with_sales is not None:
            for date, sales in sales_by_date.items():
                if sales == 0:
                    print(f"Warning: Order date {date.date()} has zero sales for {uber_location}. "
                        f"Reassigning to {latest_date_with_sales.date()} (latest date with sales).")
                    reassigned_group.loc[reassigned_group['Order Date'] == date, 'Order Date'] = latest_date_with_sales

        # Now group by Order Date with the potentially reassigned dates
        order_date_groups = reassigned_group.groupby('Order Date')

        for order_date, order_group in order_date_groups:
            # Format deposit date for JE number
            deposit_date_str = deposit_date.strftime('%m%d%y')
            order_date_str = order_date.strftime('%m%d%y')
            je_key = f"{deposit_date_str}-{order_date_str}"

            if je_key not in je_counters:
                je_counters[je_key] = 1

            je_number = f"UE{deposit_date_str}-{order_date_str}-{je_counters[je_key]:02d}"
            je_counters[je_key] += 1

            # Create comment with order date info
            je_comment = f"Deposited {deposit_date.strftime('%m/%d/%Y')} // Orders {order_date.strftime('%m/%d/%y')}"

            # Calculate uber values from the order group
            # Filter by Dining Mode
            uber_pickup = order_group[order_group['Dining Mode'] == 'Pickup']
            uber_delivery = order_group[order_group['Dining Mode'] == 'Delivery - Partner Using Uber App']

            uber_values = {
                'pickup_sales': round_to_cents(uber_pickup['Sales (excl. tax)'].sum()),
                'delivery_sales': round_to_cents(uber_delivery['Sales (excl. tax)'].sum()),
                'promotions': round_to_cents(order_group['Promotions on items'].sum()),
                'marketing_adjustment': round_to_cents(order_group['Marketing adjustment'].sum()),
                'other_payments': round_to_cents(order_group['Other payments'].sum()),
                'marketplace_fee': round_to_cents(order_group['Marketplace fee'].sum()),
                'refunds_excl_tax': round_to_cents(order_group['Refunds (excl tax)'].sum()),
                'sales_incl_tax': round_to_cents(order_group['Sales (incl. tax)'].sum()),
                'total_payout': round_to_cents(order_group['Total payout '].sum()),
                'tax_on_promotion': round_to_cents(order_group['Tax on Promotion on items'].sum()),
                'tax_on_sales': round_to_cents(order_group['Tax on sales'].sum()),
                'tax_on_refunds': round_to_cents(order_group['Tax on Refunds'].sum()),
                'marketplace_facilitator_tax': round_to_cents(order_group['Marketplace Facilitator Tax'].sum()),
                # Add new values for price adjustments
                'pickup_price_adjustments': round_to_cents(uber_pickup['Price adjustments (excl. tax)'].sum()),
                'delivery_price_adjustments': round_to_cents(uber_delivery['Price adjustments (excl. tax)'].sum()),
                'tax_on_price_adjustments': round_to_cents(order_group['Tax on price adjustments'].sum())
            }

            # Get Toast data for this specific order date
            toast_location = LOCATION_MAPPING[uber_location]['toast']

            toast_pickup = toast_df[
                (toast_df['Location'] == toast_location) &
                (toast_df['Opened'].dt.date == order_date.date()) &
                (toast_df['Dining Options'] == 'UberEats (Pickup)')
            ]

            toast_delivery = toast_df[
                (toast_df['Location'] == toast_location) &
                (toast_df['Opened'].dt.date == order_date.date()) &
                (toast_df['Dining Options'].isin(['Uber Eats - Delivery!']))
            ]

            toast_values = {
                'pickup_amount': round_to_cents(toast_pickup['Amount'].sum()),
                'delivery_amount': round_to_cents(toast_delivery['Amount'].sum()),
                'pickup_tax': round_to_cents(toast_pickup['Tax'].sum()),
                'delivery_tax': round_to_cents(toast_delivery['Tax'].sum())
            }

            # Initialize journal totals for this journal entry
            journal_totals[je_number] = {'debit': 0, 'credit': 0}

            # Initialize rows list
            rows = []

            # Common rows for both regular and special locations
            # Instead of splitting the marketplace fee proportionally, calculate it directly from order data
            uber_pickup_fees = order_group[order_group['Dining Mode'] == 'Pickup']['Marketplace fee'].sum()
            uber_delivery_fees = order_group[order_group['Dining Mode'] == 'Delivery - Partner Using Uber App']['Marketplace fee'].sum()

            # Ensure we're using rounded values
            pickup_marketplace_fee = round_to_cents(uber_pickup_fees)
            delivery_marketplace_fee = round_to_cents(uber_delivery_fees)

            # Verify that the sum matches the total (within rounding tolerance)
            total_marketplace_fee = round_to_cents(uber_values['marketplace_fee'])
            calculated_total = round_to_cents(pickup_marketplace_fee + delivery_marketplace_fee)

            # Log any discrepancies for debugging
            if abs(calculated_total - total_marketplace_fee) > 0.01:
                print(f"Warning: Calculated marketplace fees ({calculated_total}) don't match total ({total_marketplace_fee})")
                print(f"Pickup fees: {pickup_marketplace_fee}, Delivery fees: {delivery_marketplace_fee}")

            # Calculate the total UberEats AR
            uber_ar_total = (uber_values['sales_incl_tax'] +
                            uber_values['pickup_price_adjustments'] +
                            uber_values['delivery_price_adjustments'] +
                            uber_values['tax_on_price_adjustments'])

            # Calculate the total Toast AR
            toast_ar_total = (toast_values['pickup_tax'] +
                            toast_values['delivery_tax'] +
                            toast_values['pickup_amount'] +
                            toast_values['delivery_amount'])

            # Common rows for both regular and special locations
            rows = [
                {
                    'account': "UE Pickup & Takeout",
                    'detail': f"{je_comment} // Pickup Sales (excl. tax) including Price adjustments",
                    'debit': 0,
                    'credit': uber_values['pickup_sales'] + uber_values['pickup_price_adjustments']
                },
                {
                    'account': "UE Delivery",
                    'detail': f"{je_comment} // Delivery Sales (excl. tax) including Price adjustments",
                    'debit': 0,
                    'credit': uber_values['delivery_sales'] + uber_values['delivery_price_adjustments']
                },
                {
                    'account': "UberEats Discount",
                    'detail': f"{je_comment} // Promotions on items",
                    'debit': abs(uber_values['promotions']) if uber_values['promotions'] < 0 else 0,
                    'credit': uber_values['promotions'] if uber_values['promotions'] > 0 else 0
                },
                {
                    'account': "UberEats Promotion Adjustment",
                    'detail': f"{je_comment} // Marketing adjustment",
                    'debit': abs(uber_values['marketing_adjustment']) if uber_values['marketing_adjustment'] < 0 else 0,
                    'credit': uber_values['marketing_adjustment'] if uber_values['marketing_adjustment'] > 0 else 0
                },
                {
                    'account': "UberEats Ad Spend",
                    'detail': f"{je_comment} // Ad Spend (Net of Other payments)",
                    'debit': abs(uber_values['other_payments']) if uber_values['other_payments'] < 0 else 0,
                    'credit': uber_values['other_payments'] if uber_values['other_payments'] > 0 else 0
                },
                {
                    'account': "UberEats Delivery Commission",
                    'detail': f"{je_comment} // Delivery Marketplace fee",
                    'debit': abs(delivery_marketplace_fee) if delivery_marketplace_fee < 0 else 0,
                    'credit': delivery_marketplace_fee if delivery_marketplace_fee > 0 else 0
                },
                {
                    'account': "UberEats Pickup Commission",
                    'detail': f"{je_comment} // Pickup Marketplace fee",
                    'debit': abs(pickup_marketplace_fee) if pickup_marketplace_fee < 0 else 0,
                    'credit': pickup_marketplace_fee if pickup_marketplace_fee > 0 else 0
                },
                {
                    'account': "Refunds",
                    'detail': f"{je_comment} // Refunds (excl tax)",
                    'debit': abs(uber_values['refunds_excl_tax']) if uber_values['refunds_excl_tax'] < 0 else 0,
                    'credit': uber_values['refunds_excl_tax'] if uber_values['refunds_excl_tax'] > 0 else 0
                },
                {
                    'account': "Sales Tax Adjustement",
                    'detail': je_comment,
                    'debit': 0,
                    'credit': 0
                },
                {
                    'account': "Sales Tax Payable",
                    'detail': f"{je_comment} // Tax calculated from Toast",
                    'debit': toast_values['pickup_tax'] + toast_values['delivery_tax'],
                    'credit': 0
                },
                {
                    'account': "UberEats Sales",
                    'detail': f"{je_comment} // UberEats Pickup Sales calculated from Toast",
                    'debit': toast_values['pickup_amount'],
                    'credit': 0
                },
                {
                    'account': "UberEats Sales",
                    'detail': f"{je_comment} // UberEats Delivery Sales calculated from Toast",
                    'debit': toast_values['delivery_amount'],
                    'credit': 0
                },
                {
                    'account': "A/R UberEats",
                    'detail': f"{je_comment} // Sales (incl. tax) including Price adjustments and Tax on price adjustments",
                    'debit': abs(uber_values['sales_incl_tax'] + uber_values['pickup_price_adjustments'] + uber_values['delivery_price_adjustments'] + uber_values['tax_on_price_adjustments'])
                        if (uber_values['sales_incl_tax'] + uber_values['pickup_price_adjustments'] + uber_values['delivery_price_adjustments'] + uber_values['tax_on_price_adjustments']) < 0 else 0,
                    'credit': (uber_values['sales_incl_tax'] + uber_values['pickup_price_adjustments'] + uber_values['delivery_price_adjustments'] + uber_values['tax_on_price_adjustments'])
                        if (uber_values['sales_incl_tax'] + uber_values['pickup_price_adjustments'] + uber_values['delivery_price_adjustments'] + uber_values['tax_on_price_adjustments']) > 0 else 0
                },
                {
                    'account': "A/R UberEats",
                    'detail': f"{je_comment} // Cash to be deposited",
                    'debit': uber_values['total_payout'] if uber_values['total_payout'] > 0 else 0,
                    'credit': abs(uber_values['total_payout']) if uber_values['total_payout'] < 0 else 0
                },
                {
                    'account': "A/R UberEats",
                    'detail': f"{je_comment} // A/R UberEats from Toast - A/R UberEats from UberEats",
                    'debit': abs(toast_ar_total - uber_ar_total) if (toast_ar_total - uber_ar_total) < 0 else 0,
                    'credit': (toast_ar_total - uber_ar_total) if (toast_ar_total - uber_ar_total) > 0 else 0
                }
            ]
            # Add final row based on location type
            if uber_location in special_locations:
                # Special calculation for special locations
                tax_calc = (uber_values['tax_on_sales'] + uber_values['tax_on_refunds'] +
                        uber_values['tax_on_promotion'] + uber_values['marketplace_facilitator_tax'] +
                        uber_values['tax_on_price_adjustments'])  # Add tax on price adjustments

                rows.append({
                    'account': "Sales Tax Payable",
                    'detail': f"{je_comment} // Tax on sales + Tax on Refunds + Tax on Promotion on items + Marketplace Facilitator Tax + Tax on price adjustments",
                    'debit': abs(tax_calc) if tax_calc < 0 else 0,
                    'credit': tax_calc if tax_calc > 0 else 0
                })
            else:
                # Calculate totals before final row
                temp_total_debit = sum(round_to_cents(row['debit']) for row in rows)
                temp_total_credit = sum(round_to_cents(row['credit']) for row in rows)
                difference = abs(round_to_cents(temp_total_debit - temp_total_credit))
                tax_on_promotion_abs = abs(round_to_cents(uber_values['tax_on_promotion']))

                if abs(difference - tax_on_promotion_abs) < 0.01:  # Allow for small rounding differences
                    rows.append({
                        'account': "Sales Tax Payable",
                        'detail': f"{je_comment} // Tax on Promotion on items paid to us by UberEats",
                        'debit': 0,
                        'credit': tax_on_promotion_abs
                    })
                else:
                    rows.append({
                        'account': "Sales Tax Payable",
                        'detail': f"{je_comment} // Tax on Promotion on items paid to us by UberEats",
                        'debit': 0,
                        'credit': 0
                    })

            # Calculate initial totals
            temp_totals = {'debit': 0, 'credit': 0}
            for row in rows:
                debit = round_to_cents(row['debit'])
                credit = round_to_cents(row['credit'])
                temp_totals['debit'] += debit
                temp_totals['credit'] += credit

            # Check for small imbalance
            imbalance = round_to_cents(temp_totals['debit'] - temp_totals['credit'])

            if 0 < abs(imbalance) <= 0.02:
                # Find the UberEats Delivery marketplace fee row and adjust it
                for row in rows:
                    if row['account'] == "UberEats Delivery Commission" and "Marketplace fee" in row['detail']:
                        if imbalance > 0:
                            # Need to decrease debit
                            row['debit'] = round_to_cents(row['debit'] - imbalance)
                        else:
                            # Need to increase debit
                            row['debit'] = round_to_cents(row['debit'] + abs(imbalance))
                        break

            # Now create the final journal entries with potentially adjusted values
            for row in rows:
                debit = round_to_cents(row['debit'])
                credit = round_to_cents(row['credit'])

                journal_totals[je_number]['debit'] += debit
                journal_totals[je_number]['credit'] += credit

                journal_entries.append({
                    'JENumber': je_number,
                    'Type': '',
                    'DetailComment': row['detail'],
                    'JEComment': je_comment,
                    'JELocation': LOCATION_MAPPING[uber_location]['je'],
                    'Account': row['account'],
                    'Debit': debit,
                    'Credit': credit,
                    'DetailLocation': LOCATION_MAPPING[uber_location]['je'],
                    'Date': order_date.strftime('%Y-%m-%d')  # Use actual_order_date
                })

    # Check for unbalanced journals
    unbalanced_journals = []
    for je_number, totals in journal_totals.items():
        debit_sum = round_to_cents(totals['debit'])
        credit_sum = round_to_cents(totals['credit'])
        if abs(debit_sum - credit_sum) > 0.001:  # Allow for small rounding differences
            unbalanced_journals.append(je_number)
            print(f"Warning: Journal {je_number} is unbalanced:")
            print(f"    Total Debits: {debit_sum:.2f}")
            print(f"    Total Credits: {credit_sum:.2f}")
            print(f"    Difference: {abs(debit_sum - credit_sum):.2f}")

    if unbalanced_journals:
        print("\nUnbalanced Journal Entries:", ', '.join(unbalanced_journals))

    return pd.DataFrame(journal_entries)

def create_deposit_journal_entries(journal_entries_df):
    deposit_entries = []

    # Create a dictionary to track deposits by location and payout date
    deposits = {}

    # Find all entries with "Cash to be deposited"
    for _, row in journal_entries_df.iterrows():
        if "A/R UberEats" in row['Account'] and "Cash to be deposited" in row['DetailComment']:
            # Extract location
            location = row['JELocation']

            # Extract deposit date from JEComment (format: "Deposited MM/DD/YYYY // Orders...")
            je_comment = row['JEComment']
            deposit_date_str = je_comment.split("Deposited ")[1].split(" //")[0]
            deposit_date = datetime.strptime(deposit_date_str, "%m/%d/%Y")

            # Find corresponding checking account
            checking_account = None
            for store, accounts in LOCATION_MAPPING.items():
                if accounts['je'] == location:
                    checking_account = accounts['checking']
                    break

            if checking_account is None:
                print(f"Warning: Could not find checking account for location {location}")
                continue

            # Group key
            key = (location, deposit_date.strftime('%Y-%m-%d'))

            # Initialize if needed
            if key not in deposits:
                deposits[key] = {
                    'location': location,
                    'deposit_date': deposit_date,
                    'checking_account': checking_account,
                    'amount': 0,
                    'start_date': None,
                    'end_date': None
                }

            # Add to total amount (debit minus credit for this entry)
            deposits[key]['amount'] += row['Debit'] - row['Credit']

            # Get the week's date range from journal entry comment
            try:
                # This should match the format we're using for date ranges in JEComment
                # "Deposited MM/DD/YYYY // Orders MM/DD/YY"
                if "Deposited" in je_comment and "//" in je_comment and "Orders" in je_comment:
                    order_date_str = je_comment.split("Orders ")[1].split(" //")[0] if " //" in je_comment.split("Orders ")[1] else je_comment.split("Orders ")[1]
                    order_date = datetime.strptime(order_date_str, "%m/%d/%y")

                    # Track earliest and latest order dates
                    if deposits[key]['start_date'] is None or order_date < deposits[key]['start_date']:
                        deposits[key]['start_date'] = order_date

                    if deposits[key]['end_date'] is None or order_date > deposits[key]['end_date']:
                        deposits[key]['end_date'] = order_date
            except Exception as e:
                print(f"Warning: Could not parse date range from comment: {je_comment}. Error: {str(e)}")

    # Create deposit journal entries
    je_counters = {}
    for key, info in deposits.items():
        location = info['location']
        deposit_date = info['deposit_date']
        deposit_date_str = deposit_date.strftime('%m%d%y')
        checking_account = info['checking_account']
        amount = round_to_cents(info['amount'])

        # Skip if amount is zero
        if abs(amount) < 0.01:
            continue

        # Get counter for this deposit date
        if deposit_date_str not in je_counters:
            je_counters[deposit_date_str] = 1

        je_number = f"UE{deposit_date_str}-{je_counters[deposit_date_str]:02d}"
        je_counters[deposit_date_str] += 1

        # Create date range string
        date_range = "Unknown date range"
        if info['start_date'] and info['end_date']:
            # If start and end dates are the same, just show one date
            if info['start_date'].date() == info['end_date'].date():
                date_range = info['start_date'].strftime('%m/%d/%y')
            else:
                date_range = f"{info['start_date'].strftime('%m/%d/%y')}-{info['end_date'].strftime('%m/%d/%y')}"

        je_comment = f"Deposit for {date_range}"

        # Create the two rows
        deposit_entries.append({
            'JENumber': je_number,
            'Type': '',
            'DetailComment': je_comment,
            'JEComment': je_comment,
            'JELocation': location,
            'Account': checking_account,
            'Debit': amount if amount > 0 else 0,
            'Credit': abs(amount) if amount < 0 else 0,
            'DetailLocation': location,
            'Date': deposit_date.strftime('%Y-%m-%d')
        })

        deposit_entries.append({
            'JENumber': je_number,
            'Type': '',
            'DetailComment': je_comment,
            'JEComment': je_comment,
            'JELocation': location,
            'Account': 'A/R UberEats',
            'Debit': abs(amount) if amount < 0 else 0,
            'Credit': amount if amount > 0 else 0,
            'DetailLocation': location,
            'Date': deposit_date.strftime('%Y-%m-%d')
        })

    return pd.DataFrame(deposit_entries) if deposit_entries else pd.DataFrame()


def main(toast_files, uber_file, output_dir=None):
    """
    Process Toast and Uber files and create journal entries

    Args:
        toast_files (list): List of paths to Toast CSV files
        uber_file (str): Path to Uber CSV file
        output_dir (str, optional): Directory to save output file. Defaults to Downloads folder.

    Returns:
        str: Path to the output file
    """
    if output_dir is None:
        output_dir = os.path.join(os.path.expanduser('~'), 'Downloads')

    # Read input files
    toast_df = read_toast_files(toast_files)
    uber_df = read_uber_file(uber_file)

    # Create regular journal entries
    journal_entries_df = create_journal_entries(toast_df, uber_df)

    # Create deposit journal entries
    deposit_entries_df = create_deposit_journal_entries(journal_entries_df)

    # Combine both sets of entries
    combined_df = pd.concat([journal_entries_df, deposit_entries_df], ignore_index=True) if not deposit_entries_df.empty else journal_entries_df

    today = datetime.now().strftime('%m%d%Y')
    output_filename = f"UE_PayoutImport_{today}.csv"
    output_path = os.path.join(output_dir, output_filename)
    combined_df.to_csv(output_path, index=False, float_format='%.2f')

    return output_path
