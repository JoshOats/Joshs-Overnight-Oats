import pandas as pd
import os
from datetime import datetime
import glob
import logging

# CNB Dictionary embedded in code
CNB_DICTIONARY = {
    'location name': [
        'Carrot Leadership LLC',
        'Carrot Express Franchise System LLC',
        'Carrot Global LLC',
        'Carrot Express Commissary LLC',
        'Carrot Coral GablesLove LLC (Coral Gabes)',
        'Carrot Aventura Love LLC (Aventura)',
        'Carrot North Beach Love LL (North Beach)',
        'Carrot Downtown Love Two LLC',
        'Carrot Love City Place Doral Operating LLC',
        'Carrot Love Palmetto Park Operating LLC',
        'Carrot Love Brickell Operating LLC',
        'Carrot Love West Boca Operating LLC',
        'Carrot Love Aventura Mall Operating LLC',
        'Carrot Love Coconut Creek Operating LLC',
        'Carrot Love Coconut Grove Operating LLC',
        'Carrot Love Sunset Operating LLC',
        'Carrot Love Pembroke Pines Operating LLC',
        'Carrot Love Plantation Operating LLC',
        'Carrot Love River Lading Operating LLC',
        'Carrot Love Las Olas Operating LLC',
        'Carrot Love Hollywood Operating LLC',
        'Carrot Sobe Love South Florida Operating C LLC',
        'Carrot Love South Florida Operating A LLC',
        'Carrot Flatiron Love Manhattan Operating LLC',
        'Carrot Love Bryant Park Operating LLC',
        'Carrot Love 600 Lexington LLC',
        'Carrot Holdings LLC',
        'Carrot Gem LLC',
        'Carrot Dream LLC',
        'Carrot Love Dadeland Operating LLC',
        'Beyond Branding LLC',
        'CARROT "BROOKFIED" LIBERTY STREET LLC',
        'Carrot Express Miami Shores LLC',
        'Carrot Express Midtown LLC',
        'CARROT LOVE OAKLAND PARK LLC'
    ],
    'CNB Account': [
        '30000480952',
        '30000481015',
        '30000482356',
        '30000488431',
        '30000481123',
        '30000481123',
        '30000481123',
        '30000481258',
        '30000481978',
        '30000482122',
        '30000482104',
        '30000482140',
        '30000482023',
        '30000482167',
        '30000482176',
        '30000482212',
        '30000594757',
        '30000482149',
        '30000482230',
        '30000482158',
        '30000482203',
        '30000633502',
        '30000633448',
        '30000482131',
        '30000482410',
        '30000510616',
        '30000469729',
        '30000488503',
        '30000482266',
        '30000481834',
        '30000566218',
        '30000674938',
        '##N/A##',
        '##N/A##',
        '##N/A##'
    ],
    'bank account': [
        'Checking Carrot Leadership LLC',
        'Checking Carrot Express Franchise System LLC',
        'Checking Carrot Global LLC',
        'Checking Carrot Express Commissary LLC',
        'Checking Carrot Love LLC',
        'Checking Carrot Love LLC',
        'Checking Carrot Love LLC',
        'Checking Carrot Love Two LLC',
        'Checking Carrot Love City Place Doral Operating LLC',
        'Checking Carrot Love Palmetto Park Operating LLC',
        'Checking Carrot Love Brickell Operating LLC',
        'Checking Carrot Love West Boca Operating LLC',
        'Checking Carrot Love Aventura Mall Operating LLC',
        'Checking Carrot Love Coconut Creek Operating LLC',
        'Checking Carrot Love Coconut Grove Operating LLC',
        'Checking Carrot Love Sunset Operating LLC',
        'Checking Carrot Love Pembroke Pines Operating LLC',
        'Checking Carrot Love Plantation Operating LLC',
        'Checking Carrot Love River Landing LLC',
        'Checking Carrot Love Las Olas Operating LLC',
        'Checking Carrot Love Hollywood Operating LLC',
        'Checking Carrot Love South Florida Operating C LLC',
        'Checking Carrot Love South Florida Operating A LLC',
        'Checking Carrot Love Manhattan Operating LLC',
        'Checking Carrot Love Bryant Park Operating LLC',
        'Checking Carrot Love Lexington 52 LLC',
        'Checking Carrot Holdings LLC',
        'Checking Carrot Gem LLC',
        'Checking Carrot Dream LLC',
        'Checking Carrot Love Dadeland Operating LLC',
        'Checking Beyond Branding LLC',
        'Checking Carrot Love Liberty Street LLC',
        'Checking Carrot Express Miami Shores LLC',
        'Checking Carrot Express Midtown LLC',
        'Checking Carrot Love Oakland Park LLC'

    ],
    'due to/from account': [
        'Due To/From Carrot Leadership LLC',
        'Due To/From Carrot Express Franchise System LLC',
        'Due To/From Carrot Global',
        'Due To/From Carrot Express Commissary LLC',
        'Due To/From Carrot Love LLC',
        'Due To/From Carrot Love LLC',
        'Due To/From Carrot Love LLC',
        'Due To/From Carrot Love Two LLC',
        'Due To/From CARROT LOVE CITY PLACE DORAL OPERATING LLC',
        'Due To/From CARROT LOVE PALMETTO PARK OPERATING LLC',
        'Due To/From CARROT LOVE BRICKELL OPERATING LLC',
        'Due To/From CARROT LOVE WEST BOCA OPERATING LLC',
        'Due To/From CARROT LOVE AVENTURA MALL OPERATING LLC',
        'Due To/From CARROT LOVE COCONUT CREEK OPERATING LLC',
        'Due To/From CARROT LOVE COCONUT GROVE OPERATING LLC',
        'Due To/From CARROT LOVE SUNSET OPERATING LLC',
        'Due To/From CARROT LOVE PEMBROKE PINES OPERATING LLC',
        'Due To/From CARROT LOVE PLANTATION OPERATING LLC',
        'Due To/From CARROT LOVE RIVER LANDING LLC',
        'Due To/From CARROT LOVE LAS OLAS OPERATING LLC',
        'Due To/From CARROT LOVE HOLLYWOOD OPERATING LLC',
        'Due To/From Carrot Love South Florida Operating C LLC',
        'Due To/From Carrot Love South Florida Operating A LLC',
        'Due To/From Carrot Love Manhattan Operating LLC',
        'Due To/From CARROT LOVE BRYANT PARK OPERATING LLC',
        'Due To/From Carrot Love Lexington 52 LLC',
        'Due To/From Carrot Holdings LLC',
        'Due To/From Carrot Gem LLC',
        'Due To/From Carrot Dream LLC',
        'Due To/From CARROT LOVE DADELAND OPERATING  LLC',
        'Due To/From Beyond Branding',
        'Due To/From CARROT LOVE LIBERTY STREET LLC',
        'Due To/From CARROT EXPRESS MIAMI SHORES LLC',
        'Due To/From CARROT LOVE MIDTOWN LLC',
        'Due To/From CARROT LOVE OAKLAND PARK LLC'
    ]
}
# Special accounts that need different handling
EXTERNAL_LOCATIONS = [
    'Carrot Express Miami Shores LLC',
    'Carrot Express Midtown LLC',
    'CARROT LOVE OAKLAND PARK LLC'
]

# Combined locations that share accounts
COMBINED_LOCATIONS = [
    "Carrot Coral GablesLove LLC (Coral Gabes)",
    "Carrot Aventura Love LLC (Aventura)",
    "Carrot North Beach Love LL (North Beach)"
]


def calculate_difference(balance1, balance2):
    """Calculate absolute value of sum of balances"""
    bal1 = float(balance1)
    bal2 = float(balance2)
    return abs(bal1 + bal2)

def create_discrepancy_summary(mismatches, output_path):
    """Create a summary of discrepancies with beginning/ending balances and differences"""
    try:
        start_date, end_date = get_date_range(mismatches)
        summary_entries = []
        
        for item1, item2 in mismatches:
            if not isinstance(item2, list):
                # Handle normal relationships
                beg_bal1 = convert_string_to_float(item1['transactions']['BegBalAmount2'].iloc[0])
                end_bal1 = convert_string_to_float(item1['amount'])
                
                beg_bal2 = convert_string_to_float(item2['transactions']['BegBalAmount2'].iloc[0])
                end_bal2 = convert_string_to_float(item2['amount'])
                
                # Calculate difference as absolute value of sum
                difference = calculate_difference(end_bal1, end_bal2)
                
                summary_entries.extend([
                    {
                        'GL Account': item1['due_to_from'],
                        'Location': item1['location'],
                        'Beg Balance': beg_bal1,
                        'End Balance': end_bal1,
                        'Difference': difference
                    },
                    {
                        'GL Account': item2['due_to_from'],
                        'Location': item2['location'],
                        'Beg Balance': beg_bal2,
                        'End Balance': end_bal2,
                        'Difference': difference
                    },
                    # Add blank line
                    {
                        'GL Account': "",
                        'Location': "",
                        'Beg Balance': "",
                        'End Balance': "",
                        'Difference': ""
                    }
                ])
            
            else:
                # Handle Carrot Love LLC relationships
                beg_bal1 = convert_string_to_float(item1['transactions']['BegBalAmount2'].iloc[0])
                end_bal1 = convert_string_to_float(item1['amount'])
                
                carrot_love_beg_bal = sum(convert_string_to_float(rel['transactions']['BegBalAmount2'].iloc[0])
                                        for rel in item2)
                carrot_love_end_bal = sum(convert_string_to_float(rel['amount']) for rel in item2)
                
                # Calculate difference as absolute value of sum
                difference = calculate_difference(end_bal1, carrot_love_end_bal)
                
                summary_entries.append({
                    'GL Account': item1['due_to_from'],
                    'Location': item1['location'],
                    'Beg Balance': beg_bal1,
                    'End Balance': end_bal1,
                    'Difference': ""
                })
                
                for rel in item2:
                    summary_entries.append({
                        'GL Account': rel['due_to_from'],
                        'Location': rel['location'],
                        'Beg Balance': convert_string_to_float(rel['transactions']['BegBalAmount2'].iloc[0]),
                        'End Balance': convert_string_to_float(rel['amount']),
                        'Difference': ""
                    })
                
                summary_entries.append({
                    'GL Account': f"CARROT LOVE LLC Total {item2[0]['due_to_from']}",
                    'Location': "",
                    'Beg Balance': "",
                    'End Balance': carrot_love_end_bal,
                    'Difference': difference
                })
                
                summary_entries.append({
                    'GL Account': "",
                    'Location': "",
                    'Beg Balance': "",
                    'End Balance': "",
                    'Difference': ""
                })

        # Write to CSV file
        try:
            with open(output_path, 'w', newline='', encoding='utf-8') as f:
                f.write(f"{start_date} - {end_date},,,,,\n")
                f.write("GL Account,Location,Beg Balance,End Balance,,\n")
                
                i = 0
                while i < len(summary_entries):
                    entry = summary_entries[i]
                    gl_account = entry['GL Account']
                    location = entry['Location']
                    beg_balance = format_number_for_csv(entry['Beg Balance']) if entry['Beg Balance'] != "" else ""
                    end_balance = format_number_for_csv(entry['End Balance']) if entry['End Balance'] != "" else ""
                    
                    # Only show difference for last non-blank entry before a blank line
                    is_last_before_blank = (i + 1 < len(summary_entries) and 
                                          not summary_entries[i + 1]['GL Account'] and 
                                          not summary_entries[i + 1]['Location'])
                    
                    if is_last_before_blank and entry['Difference']:
                        f.write(f"{gl_account},{location},{beg_balance},{end_balance},{format_number_for_csv(entry['Difference'])},Difference (absolute)\n")
                    else:
                        f.write(f"{gl_account},{location},{beg_balance},{end_balance},,\n")
                    i += 1
        except IOError as e:
            raise Exception(f"Error writing to file {output_path}: {str(e)}")
                    
    except Exception as e:
        raise Exception(f"Error creating discrepancy summary: {str(e)}")
      

def get_unique_accounts_and_locations(mismatches):
    """Get all unique accounts and locations from mismatches"""
    accounts = set()
    locations = set()
    
    for item1, item2 in mismatches:
        # Add first item's information
        accounts.add(item1['due_to_from'])
        locations.add(item1['location'])
        
        # Add second item's information
        if isinstance(item2, list):
            for rel in item2:
                accounts.add(rel['due_to_from'])
                locations.add(rel['location'])
        else:
            accounts.add(item2['due_to_from'])
            locations.add(item2['location'])
    
    return sorted(accounts), sorted(locations)

def get_date_range(mismatches):
    """Get min and max transaction dates from all transactions"""
    min_date = None
    max_date = None
    
    for item1, item2 in mismatches:
        # Get dates from first item
        dates = pd.to_datetime(item1['transactions']['TrxDate'])
        if min_date is None or dates.min() < min_date:
            min_date = dates.min()
        if max_date is None or dates.max() > max_date:
            max_date = dates.max()
        
        # Get dates from second item
        if isinstance(item2, list):
            for rel in item2:
                dates = pd.to_datetime(rel['transactions']['TrxDate'])
                if dates.min() < min_date:
                    min_date = dates.min()
                if dates.max() > max_date:
                    max_date = dates.max()
        else:
            dates = pd.to_datetime(item2['transactions']['TrxDate'])
            if dates.min() < min_date:
                min_date = dates.min()
            if dates.max() > max_date:
                max_date = dates.max()
    
    return min_date.strftime('%m/%d/%Y'), max_date.strftime('%m/%d/%Y')


def clean_cell_content(content):
    """Clean cell content by replacing newlines and properly handle apostrophes"""
    if pd.isna(content):
        return ""
    content_str = str(content).replace("\n", " | ")
    if "," in content_str:
        return f'"{content_str}"'
    return content_str

def get_location_from_dict(due_to_from_account):
    """Get the location that owns this due to/from account from the dictionary"""
    cnb_dict = pd.DataFrame(CNB_DICTIONARY)
    matches = cnb_dict[cnb_dict['due to/from account'].apply(standardize_name) == standardize_name(due_to_from_account)]
    if not matches.empty:
        return matches['location name'].iloc[0]
    print(f"\nWARNING: Could not find owner for due to/from account: {due_to_from_account}")
    return None

def format_number_for_csv(number):
    """Format number with commas but prevent CSV from splitting"""
    if pd.isna(number):
        return '"0.00"'
    return f'"{f"{float(number):,.2f}"}"'


def should_be_grouped_under_carrot_love(location, due_to_from):
    """Determine if this relationship should be grouped under Carrot Love LLC"""
    is_carrot_love_location = any(loc.upper() in location.upper() for loc in COMBINED_LOCATIONS)
    is_due_to_from_carrot_love = "DUE TO/FROM CARROT LOVE LLC" in due_to_from.upper()
    return is_carrot_love_location and is_due_to_from_carrot_love

def is_carrot_love_location(location):
    """Check if location is one of the Carrot Love LLC combined locations"""
    return any(loc.upper() in location.upper() for loc in COMBINED_LOCATIONS)

def has_due_to_from_carrot_love(due_to_from):
    """Check if the due to/from account is for Carrot Love LLC"""
    return "DUE TO/FROM CARROT LOVE LLC" in due_to_from.upper()

def get_due_to_from_name_from_dict(location):
    """Get the standardized due to/from account name for a location from the dictionary"""
    cnb_dict = pd.DataFrame(CNB_DICTIONARY)
    matches = cnb_dict[cnb_dict['location name'].apply(standardize_name) == standardize_name(location)]
    if not matches.empty:
        return matches['due to/from account'].iloc[0]
    print(f"WARNING: Could not find dictionary entry for location: {location}")
    return None

def get_carrot_love_group_relationships(relationships, company_location):
    """Get all Carrot Love LLC relationships for a specific company"""
    carrot_love_rels = []
    
    print(f"\nLooking for Carrot Love relationships for: {company_location}")
    
    # First, check if this company has a Due To/From Carrot Love LLC account
    company_due_to_from = next((rel for rel in relationships 
                              if rel['location'] == company_location and 
                              'DUE TO/FROM CARROT LOVE LLC' in rel['due_to_from'].upper()), None)
    
    if company_due_to_from:
        print(f"Found company has Due To/From Carrot Love LLC account with balance: {company_due_to_from['amount']}")
        
        # Get the correct due to/from account name from dictionary
        expected_due_to_from = get_due_to_from_name_from_dict(company_location)
        if expected_due_to_from:
            print(f"Looking for matches with due to/from account: {expected_due_to_from}")
            
            # Now look for Carrot Love locations that have a due to/from account for this company
            for rel in relationships:
                if is_carrot_love_location(rel['location']):
                    print(f"Checking Carrot Love location: {rel['location']}")
                    print(f"Due to/from account: {rel['due_to_from']}")
                    
                    if standardize_name(rel['due_to_from']) == standardize_name(expected_due_to_from):
                        print(f"FOUND MATCH: {rel['location']} -> {rel['due_to_from']} = {rel['amount']}")
                        carrot_love_rels.append(rel)
                    else:
                        print(f"No match - expected: {expected_due_to_from}")
        else:
            print(f"Could not find dictionary entry for {company_location}")
    
    if not carrot_love_rels:
        print(f"WARNING: No Carrot Love relationships found for {company_location}")
    else:
        print(f"Found {len(carrot_love_rels)} matching Carrot Love relationships")
    
    return carrot_love_rels

def get_carrot_love_group_balance(relationships, company_location):
    """Get combined balance for all Carrot Love locations' due to/from accounts with specified company"""
    total_balance = 0
    for rel in relationships:
        if (is_carrot_love_location(rel['location']) and 
            standardize_name(company_location) in standardize_name(rel['due_to_from'])):
            total_balance += convert_string_to_float(rel['amount'])
    return total_balance

def convert_string_to_float(value):
    """Convert string numbers with commas to float"""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    # Remove commas and convert to float
    return float(str(value).replace(',', ''))

def is_external_location(location):
    """Check if location is one of the external locations"""
    return any(ext.upper() in location.upper() for ext in [
        'Carrot Express Miami Shores LLC',
        'Carrot Express Midtown LLC',
        'CARROT LOVE OAKLAND PARK LLC'
    ])

def should_combine_locations(location):
    """Check if location is in the combined locations list"""
    return any(loc.upper() in location.upper() for loc in COMBINED_LOCATIONS)

def clean_due_to_from(account_name):
    """Remove characters before 'Due' in account names"""
    if isinstance(account_name, str) and 'Due' in account_name:
        return account_name[account_name.index('Due'):]
    
    return account_name
def process_gl_data(file_path):
    """Process GL account detail file"""
    print("\nDebugging GL Data Processing:")
    
    # Skip the description rows and use row 4 as headers
    df = pd.read_csv(file_path, header=3)
    
    # Filter out excluded locations first
    excluded_locations = [
        '5A Healthy Restaurant LLC',
        '5AM Healthy Restaurant LLC',
        '5M Healthy Restaurant LLC'
    ]
    df = df[~df.iloc[:, 0].isin(excluded_locations)]
    
    # Complete column mapping based on position
    column_mapping = {
        df.columns[0]: 'LocationName1',
        df.columns[1]: 'ParentAccountName',
        df.columns[2]: 'BegBalAmount2',
        df.columns[3]: 'AccountName',
        df.columns[4]: 'BegBalAmount',
        df.columns[5]: 'TrxDate',
        df.columns[6]: 'TrxType',
        df.columns[7]: 'TrxNumber',
        df.columns[8]: 'TrxCompany',
        df.columns[9]: 'LocationName',
        df.columns[10]: 'Comment',
        df.columns[11]: 'Comment1',
        df.columns[12]: 'Debit',
        df.columns[13]: 'Credit',
        df.columns[14]: 'Textbox17',
        df.columns[23]: 'Textbox50'
    }
    
    # Rename columns
    df = df.rename(columns=column_mapping)
    
    print("\nColumns after mapping:")
    print(df.columns.tolist())
    
    relationships = []
    
    for _, group in df.groupby('LocationName1'):
        location = group['LocationName1'].iloc[0]
        print(f"\nProcessing location: {location}")
        due_to_froms = group.groupby('ParentAccountName')
        
        for account, transactions in due_to_froms:
            print(f"Processing account: {account}")
            if 'Due To/From' in str(account):
                try:
                    # Get the final amount from Textbox50
                    amount = convert_string_to_float(transactions['Textbox50'].iloc[0])
                    clean_account = clean_due_to_from(account)
                    print(f"Found relationship: {location} -> {clean_account} = {amount}")
                    
                    # Make sure transactions DataFrame has all required columns
                    transactions_copy = transactions.copy()
                    relationships.append({
                        'location': location,
                        'due_to_from': clean_account,
                        'amount': amount,
                        'transactions': transactions_copy
                    })
                except Exception as e:
                    print(f"Error processing amount for {account}: {str(e)}")
    
    print(f"\nTotal relationships found: {len(relationships)}")
    return relationships


def standardize_name(name):
    """Standardize location names for comparison"""
    if name is None:
        return ''
    # Remove common words and characters that might interfere with matching
    standardized = (name.upper()
                   .replace('"', '')
                   .replace('CARROT', '')
                   .replace('LLC', '')
                   .replace('OPERATING', '')
                   .replace('(', '')
                   .replace(')', '')
                   .replace('-', ' ')
                   .strip())
    # Remove multiple spaces
    return ' '.join(standardized.split())

def get_location_name_from_dict(due_to_from_account):
    """Get the location name that owns this due to/from account from the dictionary"""
    cnb_dict = pd.DataFrame(CNB_DICTIONARY)
    matches = cnb_dict[cnb_dict['due to/from account'].apply(standardize_name) == standardize_name(due_to_from_account)]
    if not matches.empty:
        return matches['location name'].iloc[0]
    print(f"\nWARNING: Could not find owner for due to/from account: {due_to_from_account}")
    return None

def convert_string_to_float(value):
    """Convert string numbers with commas to float"""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    # Remove commas and convert to float
    return float(str(value).replace(',', ''))


def find_matches_and_mismatches(relationships):
    """Identify matching and mismatching relationships"""
    print("\nDebugging Matches and Mismatches:")
    matches = []
    mismatches = []
    processed_pairs = set()
    
    for rel1 in relationships:
        if rel1 is None:
            continue
            
        rel1_id = f"{standardize_name(rel1['location'])}-{standardize_name(rel1['due_to_from'])}"
        if rel1_id in processed_pairs:
            continue
        
        # Handle companies that have due to/from with Carrot Love LLC
        if has_due_to_from_carrot_love(rel1['due_to_from']):
            print(f"\nAnalyzing relationship with Carrot Love LLC: {rel1['location']}")
            
            carrot_love_rels = get_carrot_love_group_relationships(relationships, rel1['location'])
            
            if carrot_love_rels:
                rel1_balance = convert_string_to_float(rel1['amount'])
                carrot_love_balance = sum(convert_string_to_float(r['amount']) for r in carrot_love_rels)
                
                print(f"Company balance: {rel1_balance}")
                print(f"Combined Carrot Love balance: {carrot_love_balance}")
                
                if abs(rel1_balance + carrot_love_balance) < 0.01:
                    print("MATCH!")
                    for carrot_rel in carrot_love_rels:
                        matches.append((rel1, carrot_rel))
                else:
                    print("MISMATCH!")
                    mismatches.append((rel1, carrot_love_rels))
                
                processed_pairs.add(rel1_id)
                for r in carrot_love_rels:
                    processed_pairs.add(f"{standardize_name(r['location'])}-{standardize_name(r['due_to_from'])}")
            else:
                print(f"No corresponding Carrot Love relationships found - skipping")
        
        # Handle all other relationships
        else:
            for rel2 in relationships:
                if rel2 is None or rel1 == rel2:
                    continue
                    
                rel2_id = f"{standardize_name(rel2['location'])}-{standardize_name(rel2['due_to_from'])}"
                pair_id = tuple(sorted([rel1_id, rel2_id]))
                
                if pair_id in processed_pairs:
                    continue
                
                if (standardize_name(rel1['location']) in standardize_name(rel2['due_to_from']) and
                    standardize_name(rel2['location']) in standardize_name(rel1['due_to_from'])):
                    
                    amount1 = convert_string_to_float(rel1['amount'])
                    amount2 = convert_string_to_float(rel2['amount'])
                    
                    print(f"\nChecking relationship between {rel1['location']} ({amount1}) and {rel2['location']} ({amount2})")
                    
                    if abs(amount1 + amount2) < 0.01:
                        print("MATCH!")
                        matches.append((rel1, rel2))
                    else:
                        print("MISMATCH!")
                        mismatches.append((rel1, rel2))
                    
                    processed_pairs.add(pair_id)
                    break
    
    print(f"\nFound {len(matches)} matching pairs")
    print(f"Found {len(mismatches)} mismatching pairs")
    return matches, mismatches

def process_transfers(matches, mismatches):
    """Process matches into transfers, handling special cases"""
    transfers = []
    
    # Process matches
    for rel1, rel2 in matches:
        # Skip if either amount is effectively zero
        if abs(convert_string_to_float(rel1['amount'])) < 0.01:
            continue
            
        # Determine sender (negative amount) and receiver
        amount1 = convert_string_to_float(rel1['amount'])
        amount2 = convert_string_to_float(rel2['amount'])
        sender = rel1 if amount1 < 0 else rel2
        receiver = rel2 if amount1 < 0 else rel1
        
        transfers.append({
            'sender': sender,
            'receiver': receiver,
            'amount': abs(amount1)
        })
    
    return transfers

def format_number_no_quotes(number):
    """Format number with commas but without quotes"""
    if pd.isna(number):
        return '0.00'
    return f"{float(number):,.2f}"


def create_je_file(matches, output_path):
    """Create journal entry import file"""
    today = datetime.now()
    je_rows = []
    cnb_dict = pd.DataFrame(CNB_DICTIONARY)
    
    transfer_idx = 1
    
    for rel1, rel2 in matches:
        try:
            # Handle Carrot Love LLC cases
            if has_due_to_from_carrot_love(rel1['due_to_from']):
                # rel2 is the Carrot Love location
                carrot_amount = convert_string_to_float(rel2['amount'])
                # Skip if individual Carrot Love location amount is zero
                if abs(carrot_amount) < 0.01:
                    print(f"Skipping zero amount JE for {rel2['location']}")
                    continue
                    
                # If Carrot Love location has positive amount, other company sends money
                if carrot_amount > 0:
                    sender = rel1
                    receiver = rel2
                else:
                    sender = rel2
                    receiver = rel1
                amount = abs(carrot_amount)
            elif has_due_to_from_carrot_love(rel2['due_to_from']):
                # rel1 is the Carrot Love location
                carrot_amount = convert_string_to_float(rel1['amount'])
                # Skip if individual Carrot Love location amount is zero
                if abs(carrot_amount) < 0.01:
                    print(f"Skipping zero amount JE for {rel1['location']}")
                    continue
                    
                # If Carrot Love location has positive amount, other company sends money
                if carrot_amount > 0:
                    sender = rel2
                    receiver = rel1
                else:
                    sender = rel1
                    receiver = rel2
                amount = abs(carrot_amount)
            else:
                # Normal case - determine sender by negative amount
                amount1 = convert_string_to_float(rel1['amount'])
                # Skip if amount is zero
                if abs(amount1) < 0.01:
                    print(f"Skipping zero amount JE between {rel1['location']} and {rel2['location']}")
                    continue
                    
                sender = rel1 if amount1 < 0 else rel2
                receiver = rel2 if amount1 < 0 else rel1
                amount = abs(amount1)
            
            # Get sender/receiver info from dictionary
            sender_info = cnb_dict[cnb_dict['location name'].apply(standardize_name) == 
                                 standardize_name(sender['location'])].iloc[0]
            receiver_info = cnb_dict[cnb_dict['location name'].apply(standardize_name) == 
                                   standardize_name(receiver['location'])].iloc[0]
            
            je_number = f"Transfer {today.strftime('%m%d%y')}-{transfer_idx:02d}"
            print(f"Creating JE {je_number}: {sender['location']} -> {receiver['location']} Amount: {amount}")
            
            je_rows.extend([
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'DetailComment': '',
                    'Reversal Date': '',
                    'JEComment': '',
                    'JELocation': sender['location'],
                    'Account': sender_info['bank account'],
                    'Debit': format_number_no_quotes(0),
                    'Credit': format_number_no_quotes(amount),
                    'DetailLocation': sender['location'],
                    'Date': today.strftime('%m/%d/%Y')
                },
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'DetailComment': '',
                    'Reversal Date': '',
                    'JEComment': '',
                    'JELocation': sender['location'],
                    'Account': receiver_info['due to/from account'],
                    'Debit': format_number_no_quotes(amount),
                    'Credit': format_number_no_quotes(0),
                    'DetailLocation': sender['location'],
                    'Date': today.strftime('%m/%d/%Y')
                },
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'DetailComment': '',
                    'Reversal Date': '',
                    'JEComment': '',
                    'JELocation': sender['location'],
                    'Account': sender_info['due to/from account'],
                    'Debit': format_number_no_quotes(0),
                    'Credit': format_number_no_quotes(amount),
                    'DetailLocation': receiver['location'],
                    'Date': today.strftime('%m/%d/%Y')
                },
                {
                    'JENumber': je_number,
                    'Type': 'Standard',
                    'DetailComment': '',
                    'Reversal Date': '',
                    'JEComment': '',
                    'JELocation': sender['location'],
                    'Account': receiver_info['bank account'],
                    'Debit': format_number_no_quotes(amount),
                    'Credit': format_number_no_quotes(0),
                    'DetailLocation': receiver['location'],
                    'Date': today.strftime('%m/%d/%Y')
                }
            ])
            
            transfer_idx += 1
            
        except Exception as e:
            print(f"Warning: Could not create JE for {rel1['location']} -> {rel2['location']}: {str(e)}")
    
    if je_rows:
        je_df = pd.DataFrame(je_rows)
        je_df.to_csv(output_path, index=False)
        print(f"Created {len(je_rows) // 4} journal entries")

def create_cnb_transfer_files(matches, base_path):
    """Create CNB transfer files with account names and splitting into chunks of 35"""
    transfer_rows = []
    cnb_dict = pd.DataFrame(CNB_DICTIONARY)
    today = datetime.now().strftime('%m%d%Y')
    
    for rel1, rel2 in matches:
        # Skip external locations and zero amounts
        amount1 = convert_string_to_float(rel1['amount'])
        if (abs(amount1) < 0.01 or 
            is_external_location(rel1['location']) or 
            is_external_location(rel2['location'])):
            continue
        
        try:
            # Handle Carrot Love LLC cases
            if has_due_to_from_carrot_love(rel1['due_to_from']):
                carrot_amount = convert_string_to_float(rel2['amount'])
                if abs(carrot_amount) < 0.01:
                    continue
                    
                if carrot_amount > 0:
                    sender = rel1
                    receiver = rel2
                else:
                    sender = rel2
                    receiver = rel1
                amount = abs(carrot_amount)
            elif has_due_to_from_carrot_love(rel2['due_to_from']):
                carrot_amount = convert_string_to_float(rel1['amount'])
                if abs(carrot_amount) < 0.01:
                    continue
                    
                if carrot_amount > 0:
                    sender = rel2
                    receiver = rel1
                else:
                    sender = rel1
                    receiver = rel2
                amount = abs(carrot_amount)
            else:
                sender = rel1 if amount1 < 0 else rel2
                receiver = rel2 if amount1 < 0 else rel1
                amount = abs(amount1)
            
            # Get sender/receiver info from dictionary
            sender_info = cnb_dict[cnb_dict['location name'].apply(standardize_name) == 
                                 standardize_name(sender['location'])].iloc[0]
            receiver_info = cnb_dict[cnb_dict['location name'].apply(standardize_name) == 
                                   standardize_name(receiver['location'])].iloc[0]
            
            transfer_rows.append({
                'From': sender_info['CNB Account'],
                'To': receiver_info['CNB Account'],
                'Amount': amount,
                'From company ---> To company': f"{sender_info['location name']} ---> {receiver_info['location name']}"
            })
            
        except Exception as e:
            print(f"Warning: Could not create CNB transfer for {sender['location']} ---> {receiver['location']}: {str(e)}")
    
    # Split into files of 35 rows each
    if transfer_rows:
        transfer_df = pd.DataFrame(transfer_rows)
        
        # Calculate number of files needed
        chunk_size = 35
        num_chunks = (len(transfer_df) + chunk_size - 1) // chunk_size  # Round up division
        
        print(f"Creating {num_chunks} CNB transfer files...")
        
        # Create separate files
        for i in range(num_chunks):
            start_idx = i * chunk_size
            end_idx = min((i + 1) * chunk_size, len(transfer_df))
            
            chunk_df = transfer_df.iloc[start_idx:end_idx]
            
            # Create filename
            filename = f"CNB-{i+1}_Transfer_{today}.csv"
            output_path = os.path.join(base_path, filename)
            
            # Ensure directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            chunk_df.to_csv(output_path, index=False)
            print(f"Created CNB transfer file {i+1} with {len(chunk_df)} rows: {filename}")


def create_external_transfer_file(matches, output_path):
    """Create external transfer file"""
    transfer_rows = []
    
    for rel1, rel2 in matches:
        # Only process if either location is external and amount is not zero
        amount1 = convert_string_to_float(rel1['amount'])
        if (abs(amount1) < 0.01 or 
            not (is_external_location(rel1['location']) or 
                 is_external_location(rel2['location']))):
            continue
        
        # Determine sender (negative amount) and receiver
        sender = rel1 if amount1 < 0 else rel2
        receiver = rel2 if amount1 < 0 else rel1
        amount = abs(amount1)
        
        transfer_rows.append({
            'From': sender['location'],
            'To': receiver['location'],
            'Amount': amount
        })
    
    if transfer_rows:
        transfer_df = pd.DataFrame(transfer_rows)
        transfer_df.to_csv(output_path, index=False)


def create_discrepancy_file(mismatches, output_path):
    start_date, end_date = get_date_range(mismatches)
    accounts, locations = get_unique_accounts_and_locations(mismatches)
    
    header_info = [
        "Account Detail: " + "; ".join(accounts),
        "",
        f"{start_date} - {end_date}",
        "Location: " + "; ".join(locations),
        ""
    ]
    
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        for line in header_info:
            f.write(f"{line}\n")
        
        f.write("Master Location,Account,Date,Type,Ref. Number,Company,Location,Comment/Item,Debit,Credit,Balance,Mismatch\n")
        
        if not mismatches:
            f.write("Grand Total,,,,,,,,0.00,0.00,0\n")
            return
        
        total_all_debits = 0
        total_all_credits = 0
        
        for item1, item2 in mismatches:
            if not isinstance(item2, list):
                # Write first location
                location_debit = 0
                location_credit = 0
                
                f.write(f"{item1['location']},,,,,,,,,,,\n")
                beg_balance = convert_string_to_float(item1['transactions']['BegBalAmount2'].iloc[0])
                f.write(f",{item1['due_to_from']},,,,,,,,,{format_number_for_csv(beg_balance)},Beg Balance\n")
                
                filtered_trans1, filtered_trans2 = filter_matching_transactions(
                    item1['transactions'],
                    item2['transactions']
                )
                
                for _, row in filtered_trans1.iterrows():
                    debit = convert_string_to_float(row['Debit'])
                    credit = convert_string_to_float(row['Credit'])
                    balance = convert_string_to_float(row['Textbox17'])
                    location_debit += debit
                    location_credit += credit
                    
                    f.write(f",,{row['TrxDate']},{row['TrxType']},{clean_cell_content(row['TrxNumber'])},"
                           f"{clean_cell_content(row['TrxCompany'])},{clean_cell_content(row['LocationName1'])},"
                           f"{clean_cell_content(row['Comment'])},{format_number_for_csv(debit)},{format_number_for_csv(credit)},"
                           f"{format_number_for_csv(balance)},{row['Mismatch']}\n")
                
                total_all_debits += location_debit
                total_all_credits += location_credit
                
                amount = convert_string_to_float(item1['amount'])
                f.write(f",,,,,,,,Total {item1['due_to_from']},,"
                       f"{format_number_for_csv(amount)},End Balance\n\n")
                
                # Write second location
                location_debit = 0
                location_credit = 0
                
                f.write(f"{item2['location']},,,,,,,,,,,\n")
                beg_balance = convert_string_to_float(item2['transactions']['BegBalAmount2'].iloc[0])
                f.write(f",{item2['due_to_from']},,,,,,,,,{format_number_for_csv(beg_balance)},Beg Balance\n")
                
                for _, row in filtered_trans2.iterrows():
                    debit = convert_string_to_float(row['Debit'])
                    credit = convert_string_to_float(row['Credit'])
                    balance = convert_string_to_float(row['Textbox17'])
                    location_debit += debit
                    location_credit += credit
                    
                    f.write(f",,{row['TrxDate']},{row['TrxType']},{clean_cell_content(row['TrxNumber'])},"
                           f"{clean_cell_content(row['TrxCompany'])},{clean_cell_content(row['LocationName1'])},"
                           f"{clean_cell_content(row['Comment'])},{format_number_for_csv(debit)},{format_number_for_csv(credit)},"
                           f"{format_number_for_csv(balance)},{row['Mismatch']}\n")
                
                total_all_debits += location_debit
                total_all_credits += location_credit
                
                amount = convert_string_to_float(item2['amount'])
                f.write(f",,,,,,,,Total {item2['due_to_from']},,"
                       f"{format_number_for_csv(amount)},End Balance\n")
                
                # Calculate difference with sign check
                difference = calculate_difference(item1['amount'], item2['amount'])
                f.write(f",,,,,,,,,,,,DIFFERENCE:,{format_number_for_csv(difference)}\n")
                
                # Add separator after relationship group
                f.write("\n" + "-" * 500 + "\n\n")
            
            else:
                # Handle Carrot Love LLC relationships
                location_debit = 0
                location_credit = 0
                
                f.write(f"{item1['location']},,,,,,,,,,,\n")
                beg_balance = convert_string_to_float(item1['transactions']['BegBalAmount2'].iloc[0])
                f.write(f",{item1['due_to_from']},,,,,,,,,{format_number_for_csv(beg_balance)},Beg Balance\n")
                
                filtered_trans1 = item1['transactions']
                for carrot_loc in item2:
                    filtered1, _ = filter_matching_transactions(filtered_trans1, carrot_loc['transactions'])
                    filtered_trans1 = filtered1
                
                for _, row in filtered_trans1.iterrows():
                    debit = convert_string_to_float(row['Debit'])
                    credit = convert_string_to_float(row['Credit'])
                    balance = convert_string_to_float(row['Textbox17'])
                    location_debit += debit
                    location_credit += credit
                    
                    f.write(f",,{row['TrxDate']},{row['TrxType']},{clean_cell_content(row['TrxNumber'])},"
                           f"{clean_cell_content(row['TrxCompany'])},{clean_cell_content(row['LocationName1'])},"
                           f"{clean_cell_content(row['Comment'])},{format_number_for_csv(debit)},{format_number_for_csv(credit)},"
                           f"{format_number_for_csv(balance)},{row['Mismatch']}\n")
                
                total_all_debits += location_debit
                total_all_credits += location_credit
                
                amount = convert_string_to_float(item1['amount'])
                f.write(f",,,,,,,Total {item1['due_to_from']},,,"
                       f"{format_number_for_csv(amount)},End Balance\n\n")
                
                # Process each Carrot Love location
                carrot_love_total_debit = 0
                carrot_love_total_credit = 0
                carrot_love_total_amount = 0
                
                for rel in item2:
                    location_debit = 0
                    location_credit = 0
                    
                    f.write(f"{rel['location']},,,,,,,,,,,\n")
                    beg_balance = convert_string_to_float(rel['transactions']['BegBalAmount2'].iloc[0])
                    f.write(f",{rel['due_to_from']},,,,,,,,,{format_number_for_csv(beg_balance)},Beg Balance\n")
                    
                    _, filtered_trans = filter_matching_transactions(item1['transactions'], rel['transactions'])
                    
                    for _, row in filtered_trans.iterrows():
                        debit = convert_string_to_float(row['Debit'])
                        credit = convert_string_to_float(row['Credit'])
                        balance = convert_string_to_float(row['Textbox17'])
                        location_debit += debit
                        location_credit += credit
                        
                        f.write(f",,{row['TrxDate']},{row['TrxType']},{clean_cell_content(row['TrxNumber'])},"
                               f"{clean_cell_content(row['TrxCompany'])},{clean_cell_content(row['LocationName1'])},"
                               f"{clean_cell_content(row['Comment'])},{format_number_for_csv(debit)},{format_number_for_csv(credit)},"
                               f"{format_number_for_csv(balance)},{row['Mismatch']}\n")
                    
                    carrot_love_total_debit += location_debit
                    carrot_love_total_credit += location_credit
                    carrot_love_total_amount += convert_string_to_float(rel['amount'])
                    total_all_debits += location_debit
                    total_all_credits += location_credit
                    
                    amount = convert_string_to_float(rel['amount'])
                    f.write(f",,,,,,,,Total {rel['due_to_from']},,"
                           f"{format_number_for_csv(amount)},End Balance\n\n")
                
                # Add Carrot Love LLC total row
                f.write(f",,,,,,,,CARROT LOVE LLC Total {item2[0]['due_to_from']},,"
                       f"{format_number_for_csv(carrot_love_total_amount)},End Balance\n")
                
                # Calculate difference
                difference = calculate_difference(item1['amount'], carrot_love_total_amount)
                f.write(f",,,,,,,,,,,,Difference:,{format_number_for_csv(difference)}\n")
                
                # Add separator after relationship group
                f.write("\n" + "-" * 500 + "\n\n")
        
        f.write(f"Grand Total,,,,,,,,{format_number_for_csv(total_all_debits)},{format_number_for_csv(total_all_credits)},0\n")

def filter_matching_transactions(transactions1_df, transactions2_df):
    """Remove matching transactions with two-pass matching"""
    df1 = transactions1_df.copy()
    df2 = transactions2_df.copy()
    
    # Pre-process numeric columns
    for df in [df1, df2]:
        for col in ['Debit', 'Credit']:
            df[col] = df[col].fillna('0')
            df[col] = df[col].replace('', '0')
            df[col] = df[col].apply(lambda x: str(x).replace(',', '') if isinstance(x, str) else x)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Remove zero-value transactions
    df1 = df1[~((df1['Debit'] == 0) & (df1['Credit'] == 0))]
    df2 = df2[~((df2['Debit'] == 0) & (df2['Credit'] == 0))]
    
    df1['TrxDate'] = pd.to_datetime(df1['TrxDate'])
    df2['TrxDate'] = pd.to_datetime(df2['TrxDate'])
    
    df1['Mismatch'] = "***MISMATCH***"
    df2['Mismatch'] = "***MISMATCH***"
    
    keep_mask1 = pd.Series(True, index=df1.index)
    keep_mask2 = pd.Series(True, index=df2.index)
    
    # First pass - match dates and amounts
    for idx1, row1 in df1.iterrows():
        if not keep_mask1[idx1]:
            continue
        
        debit1 = float(row1['Debit'])
        credit1 = float(row1['Credit'])
        
        for idx2, row2 in df2.iterrows():
            if not keep_mask2[idx2]:
                continue
            
            debit2 = float(row2['Debit'])
            credit2 = float(row2['Credit'])
            
            if row1['TrxDate'] == row2['TrxDate']:
                if (abs(debit1 - credit2) < 0.01 and debit1 > 0) or (abs(credit1 - debit2) < 0.01 and credit1 > 0):
                    keep_mask1[idx1] = False
                    keep_mask2[idx2] = False
                    df1.at[idx1, 'Mismatch'] = ""
                    df2.at[idx2, 'Mismatch'] = ""
                    break
    
    # Second pass - match only amounts for remaining unmatched transactions
    for idx1, row1 in df1.iterrows():
        if not keep_mask1[idx1]:
            continue
        
        debit1 = float(row1['Debit'])
        credit1 = float(row1['Credit'])
        
        for idx2, row2 in df2.iterrows():
            if not keep_mask2[idx2]:
                continue
            
            debit2 = float(row2['Debit'])
            credit2 = float(row2['Credit'])
            
            if (abs(debit1 - credit2) < 0.01 and debit1 > 0) or (abs(credit1 - debit2) < 0.01 and credit1 > 0):
                keep_mask1[idx1] = False
                keep_mask2[idx2] = False
                df1.at[idx1, 'Mismatch'] = ""
                df2.at[idx2, 'Mismatch'] = ""
                break
    
    return df1[keep_mask1], df2[keep_mask2]

def main():
    # Create output directory if it doesn't exist
    output_dir = '../Modified'
    os.makedirs(output_dir, exist_ok=True)
    
    print("\n=== Starting Processing ===")
    
    try:
        # Process files as before
        gl_data = process_gl_data('../Original/GL Account Detail (15).csv')
        matches, mismatches = find_matches_and_mismatches(gl_data)
        
        # Generate output files
        today = datetime.now().strftime('%m%d%Y')
        
        print("\nGenerating output files...")
        create_je_file(
            matches,
            os.path.join(output_dir, f'R365_JE_{today}.csv')
        )
        
        create_cnb_transfer_files(  # Note: Changed to new function name
            matches,
            output_dir  # Pass directory path instead of full file path
        )
        
        create_discrepancy_file(
            mismatches,
            os.path.join(output_dir, f'Discrepancies_{today}.csv')
        )
        
        create_external_transfer_file(
            matches,
            os.path.join(output_dir, f'External_Transfer_{today}.csv')
        )
        
        print("\nProcessing complete!")
        
    except Exception as e:
        print(f"\nError during processing: {str(e)}")
        raise

if __name__ == "__main__":
    main()