# payroll_automation.py

import os
from datetime import datetime
from PyQt5.QtCore import QThread, pyqtSignal

# List of employee values to exclude
exclude_employees = [
    'Cashier, AM','CASHIER, AM', 'Cashier, All Day', 'Cashier AM, Cashier AM',
    'Cashier PM, Cashier PM', 'Cashier, PM', 'Online Ordering, Default',
    'Casher, AM', 'Card Swipe, PM', 'Login, KDS', 'KDS, Login', 'CASHIER, PM',
    'Login (Hollywood), KDS', 'Login, Toast Generic', 'KDS Login, Aventura Mall',
    'KDS Login, Toast', 'Login, Toast KDS', 'Sucre, Ana','Screens, KDS','User, Generic'
]

class PayrollAutomationThread(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, time_entries_path, payroll_dict_path, tips_path, output_dir):
        super().__init__()
        self.time_entries_path = time_entries_path
        self.payroll_dict_path = payroll_dict_path
        self.tips_path = tips_path
        self.output_dir = output_dir

    def parse_date(self, date_str):
        import pandas as pd
        """
        Parse date string to date object using specific common formats to avoid warnings.

        Args:
            date_str (str): Date string to parse

        Returns:
            datetime.date: Parsed date or None if parsing fails
        """
        if pd.isna(date_str):
            return None

        date_formats = [
            '%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y',  # Common formats
            '%m-%d-%Y', '%d-%m-%Y', '%Y/%m/%d',  # Alternative formats
            '%m/%d/%y', '%d/%m/%y', '%y/%m/%d',  # Short year formats
            '%m/%d/%Y %H:%M:%S', '%m/%d/%Y %H:%M:%S %p'  # With time
        ]

        for fmt in date_formats:
            try:
                return pd.to_datetime(date_str, format=fmt).date()
            except (ValueError, TypeError):
                continue

        # If all explicit formats fail, try the default parser as fallback
        try:
            return pd.to_datetime(date_str).date()
        except (ValueError, TypeError):
            return None

    def load_time_entries(self, file_path):
        import pandas as pd
        """
        Load and preprocess the time entries file.

        Args:
            file_path (str): Path to the time entries CSV file

        Returns:
            pandas.DataFrame: Preprocessed time entries data
        """
        self.update_signal.emit(f"Loading time entries from: {file_path}")

        # Load the CSV file
        time_entries = pd.read_csv(file_path)

        # Parse the date columns
        if 'In Date' in time_entries.columns:
            # Convert 'In Date' column to datetime object to avoid repeated warnings
            time_entries['In_Date_Parsed'] = time_entries['In Date'].apply(self.parse_date)

        # List of locations to exclude
        exclude_locations = [
            'Weston', 'West Kendall (London Square)', 'Pinecrest'
        ]

        # Filter out the excluded employees and locations
        filtered_entries = time_entries[
            (~time_entries['Employee'].isin(exclude_employees)) &
            (~time_entries['Location'].isin(exclude_locations))
        ]

        self.update_signal.emit(f"Loaded {len(time_entries)} entries, filtered down to {len(filtered_entries)} entries")

        return filtered_entries

    def load_tips_file(self, file_path):
        import pandas as pd
        """
        Load and process the tips file to extract employee tips by location.

        Args:
            file_path (str): Path to the tips file CSV

        Returns:
            dict: Dictionary mapping (location, employee_name) to tips amount
        """
        self.update_signal.emit(f"Loading tips from: {file_path}")

        # Load the CSV file
        tips_df = pd.read_csv(file_path, encoding='cp1252')

        # Dictionary to store tips by (location, employee)
        employee_tips = {}

        # Find all occurrences of "Empleado registro Toast" in column B
        empleado_indices = tips_df.index[tips_df.iloc[:, 1] == "Empleado registro Toast"].tolist()

        if not empleado_indices:
            self.update_signal.emit("Warning: No 'Empleado registro Toast' entries found in the tips file")
            return employee_tips

        # List of markers that indicate the end of a section
        end_markers = [
            "Cashier, AM",
            "TOT HORAS",
            "Cashier, PM",
            "Online Ordering, Online Ordering",
            "Cashier AM, Cashier AM",
            "Casher, AM"
        ]

        # Process each section
        for i, start_idx in enumerate(empleado_indices):
            # Determine end index by finding any of the end markers after this section start
            end_idx = len(tips_df)  # Default to end of file

            for marker in end_markers:
                marker_indices = tips_df.index[(tips_df.iloc[:, 1] == marker) &
                                              (tips_df.index > start_idx)].tolist()
                if marker_indices:
                    # Find the closest marker
                    if marker_indices[0] < end_idx:
                        end_idx = marker_indices[0]

            # Skip the header row with "Empleado registro Toast"
            start_row = start_idx + 1

            # Process employees in this section (from start_row to end_idx, not including end markers)
            for row_idx in range(start_row, end_idx):
                if row_idx >= len(tips_df):
                    break

                employee_name = tips_df.iloc[row_idx, 1]  # Column B

                # Skip if this row is empty or an end marker
                if pd.isna(employee_name) or employee_name in end_markers:
                    continue

                # Get location from column A and tips from column AZ (index 51)
                location = tips_df.iloc[row_idx, 0]  # Column A

                # Strip whitespace from both location and employee name
                if pd.notna(location):
                    location = location.strip()
                if pd.notna(employee_name):
                    employee_name = employee_name.strip()

                # Get tips amount from column AZ
                tips_amount = tips_df.iloc[row_idx, 51]  # Column AZ

                # Convert tips to float
                try:
                    if pd.notna(tips_amount):
                        tips_amount = float(tips_amount)
                    else:
                        tips_amount = 0.0
                except (ValueError, TypeError):
                    tips_amount = 0.0

                # Store in dictionary
                if pd.notna(location) and pd.notna(employee_name):
                    employee_tips[(location, employee_name)] = tips_amount

        self.update_signal.emit(f"Loaded tips for {len(employee_tips)} employees")

        return employee_tips

    def load_payroll_dictionary(self, file_path):
        import pandas as pd
        """
        Load the payroll dictionary from Excel file.

        Args:
            file_path (str): Path to the payroll dictionary Excel file

        Returns:
            pandas.DataFrame: Payroll dictionary data
        """
        self.update_signal.emit(f"Loading payroll dictionary from: {file_path}")

        # Load the Excel file
        payroll_dict = pd.read_excel(file_path)

        # Parse the Holiday Date column if it exists
        if 'Holiday Date' in payroll_dict.columns:
            payroll_dict['Holiday_Date_Parsed'] = payroll_dict['Holiday Date'].apply(self.parse_date)

        # Round the Rate column to 2 decimal places if it exists
        if 'Rate' in payroll_dict.columns:
            payroll_dict['Rate'] = payroll_dict['Rate'].round(2)

        self.update_signal.emit(f"Loaded payroll dictionary with {len(payroll_dict)} entries")

        return payroll_dict

    def calculate_gross_pay(self, time_entries_df, payroll_dict_df, employee_tips=None):
        import pandas as pd
        """
        Calculate gross pay for each employee at each location.

        Args:
            time_entries_df (pandas.DataFrame): Filtered time entries data
            payroll_dict_df (pandas.DataFrame): Payroll dictionary data
            employee_tips (dict, optional): Dictionary mapping (location, employee) to tips amount

        Returns:
            pandas.DataFrame: Result containing all required columns
            list: List of employees missing from payroll dictionary
            list: List of row indices for salaried employees without time entries
            list: List of employees with tips who weren't found in the summary
            list: List of salaried employees without time entries
            list: List of employees marked not to be paid
        """
        # Remove the input_dir parameter from the function definition
        # Create lists to track results
        self.update_signal.emit("Calculating gross pay...")

        # Create a tracking list for unmatched tips employees
        unmatched_tips_employees = []

        # Initialize employee_tips if not provided
        if employee_tips is None:
            employee_tips = {}

        # Group by Location and Employee to sum the Total Hours
        employee_hours = time_entries_df.groupby(['Location', 'Employee'])['Total Hours'].sum().reset_index()

        # Create a result dataframe with all the required columns
        result_columns = [
            'Location', 'Number', 'Num-Loc', 'ID', 'Code', 'ADP Company Code',
            'ADP Employee Code', 'Employee', 'Total Hours', 'Gross Pay',
            'Tips', 'Gross + Tips', 'Holiday Holiday', 'Regular Hours',
            'Overtime Hours', 'Wage Basis', 'Rate', 'Tip Adjustment', 'Bonus',
            'Wage Owed', 'Reimbursements', 'Comment',
            'Recordatorio futuros', 'Raw Overtime', 'Holiday Hours'  # Added working columns
        ]

        # Track missing employees for logging
        missing_employees = []
        # Track employees in time entries for later comparison with dictionary
        employees_processed = set()
        # Track employees not to be paid
        not_to_be_paid = []

        # Check for holiday dates in payroll dictionary
        holiday_dates = []
        if 'Holiday_Date_Parsed' in payroll_dict_df.columns:
            holiday_dates = payroll_dict_df['Holiday_Date_Parsed'].dropna().unique()
            holiday_dates = [date for date in holiday_dates if date is not None]
            if holiday_dates:
                self.update_signal.emit(f"Found holiday dates: {holiday_dates}")

        # First create a dictionary to store all processed employees by location
        all_employees_by_location = {}

        # Process employees with time entries first
        location_groups = employee_hours.groupby('Location')
        for location, location_group in location_groups:
            # Initialize the location entry if not already in the dictionary
            if location not in all_employees_by_location:
                all_employees_by_location[location] = []

            # Start numbering from 1 for each location
            employee_number = 1

            # Process each employee at this location
            for index, row in location_group.iterrows():
                employee = row['Employee']
                total_hours = row['Total Hours']

                # Skip employees with zero total hours
                if total_hours <= 0:
                    continue

                # Add to employees processed set
                employees_processed.add((location, employee))

                # Find the employee in the payroll dictionary
                matched_employee = payroll_dict_df[
                    (payroll_dict_df['Location'] == location) &
                    (payroll_dict_df['Employee'] == employee)
                ]

                # Check if employee should be paid
                should_be_paid = True
                if not matched_employee.empty and 'PAY?' in matched_employee.columns:
                    pay_status = matched_employee['PAY?'].iloc[0]
                    if isinstance(pay_status, str) and pay_status.lower() == 'no':
                        should_be_paid = False
                        not_to_be_paid.append(f"'{employee}' in '{location}' is set NOT to be paid EVEN THOUGH they have hours logged in Time Entries file")
                        continue  # Skip this employee

                # Create a row for this employee
                employee_row = {col: None for col in result_columns}

                # Populate the basic information
                employee_row['Number'] = employee_number
                employee_row['Num-Loc'] = f"{employee_number}{location}"
                employee_row['Employee'] = employee
                employee_row['Total Hours'] = total_hours
                employee_row['Location'] = location  # Ensure Location is always set

                # Calculate Holiday Hours - using the parsed date column for efficiency
                holiday_hours = 0
                if holiday_dates and 'In_Date_Parsed' in time_entries_df.columns:
                    # Filter time entries for this employee on holiday dates
                    for date in holiday_dates:
                        holiday_entries = time_entries_df[
                            (time_entries_df['Location'] == location) &
                            (time_entries_df['Employee'] == employee) &
                            (time_entries_df['In_Date_Parsed'] == date)
                        ]
                        if not holiday_entries.empty:
                            holiday_hours += holiday_entries['Total Hours'].sum()

                employee_row['Holiday Hours'] = holiday_hours

                # Calculate Raw Overtime
                raw_overtime = max(0, total_hours - 40)
                employee_row['Raw Overtime'] = raw_overtime

                # Calculate Overtime Hours using the formula: IF(D52-F52>0,D52-F52,0)
                # D52 = Raw Overtime, F52 = Holiday Hours
                overtime_hours = max(0, raw_overtime - holiday_hours)
                employee_row['Overtime Hours'] = overtime_hours

                # Flag to track if employee was found in payroll dictionary
                employee_found = False

                if not matched_employee.empty:
                    # Employee found in payroll dictionary
                    employee_found = True
                    employee_info = matched_employee.iloc[0]

                    # Create ID from ADP Company Code, ADP Employee Code, and Employee
                    adp_company_code = str(employee_info.get('ADP Company Code', '')) if not pd.isna(employee_info.get('ADP Company Code', '')) else ''
                    adp_employee_code = str(employee_info.get('ADP Employee Code', '')) if not pd.isna(employee_info.get('ADP Employee Code', '')) else ''
                    employee_id = f"{adp_company_code}{adp_employee_code}{employee}"
                    employee_row['ID'] = employee_id

                    # Copy all matching columns from payroll dictionary
                    for col in result_columns:
                        if col in employee_info and col not in ['Number', 'Num-Loc', 'Total Hours', 'Location', 'ID',
                                                               'Raw Overtime', 'Holiday Hours', 'Overtime Hours',
                                                               'Regular Hours', 'Gross Pay', 'Gross + Tips', 'Holiday Holiday']:
                            employee_row[col] = employee_info[col]

                    # Get the adjustment value from the dictionary
                    if 'Tip Adjustment' in employee_info and not pd.isna(employee_info['Tip Adjustment']):
                        adjustment = employee_info['Tip Adjustment']
                        employee_row['Tip Adjustment'] = adjustment
                    else:
                        employee_row['Tip Adjustment'] = 0

                    # Set default values for calculation fields if they don't exist
                    for field in ['Bonus', 'Wage Owed', 'Reimbursements']:
                        if field not in employee_info or pd.isna(employee_info[field]):
                            employee_row[field] = 0
                        else:
                            employee_row[field] = round(employee_info[field], 2)

                    # Calculate Regular Hours using the formula:
                    # =IF(H52-IF(F52-D52>0,F52-D52,0)-(IF(D52-F52>0,D52-F52,0))>40,40,(H52-F52-D52))
                    # H52 = Total Hours, F52 = Holiday Hours, D52 = Raw Overtime
                    overtime_part1 = max(0, holiday_hours - raw_overtime)
                    overtime_part2 = max(0, raw_overtime - holiday_hours)
                    regular_hours_calc = total_hours - overtime_part1 - overtime_part2
                    regular_hours = min(40, regular_hours_calc) if regular_hours_calc > 40 else regular_hours_calc
                    employee_row['Regular Hours'] = regular_hours

                    # Copy Holiday Hours to Holiday Holiday for the output
                    employee_row['Holiday Holiday'] = holiday_hours

                    # Calculate gross pay based on wage basis
                    if 'Wage Basis' in employee_info and employee_info['Wage Basis'].lower() == 'hourly':
                        if 'Rate' in employee_info:
                            rate = round(employee_info['Rate'], 2)
                            employee_row['Rate'] = rate

                            # Calculate Gross Pay for hourly employees using the modified formula:
                            # Gross = (Regular Hours * Rate) + (Overtime Hours * Rate * 1.5) + (Holiday Hours * Rate * 1.5) + Wage Owed + Bonus
                            # NOTE: Reimbursements is no longer included in this calculation
                            regular_pay = regular_hours * rate
                            overtime_pay = overtime_hours * rate * 1.5
                            holiday_pay = holiday_hours * rate * 1.5
                            wage_owed = employee_row['Wage Owed'] or 0
                            bonus = employee_row['Bonus'] or 0

                            gross_pay = regular_pay + overtime_pay + holiday_pay + wage_owed + bonus
                            employee_row['Gross Pay'] = round(gross_pay, 2)
                        else:
                            employee_row['Gross Pay'] = 0
                            self.update_signal.emit(f"Warning: No rate found for hourly employee '{employee}' at '{location}'")
                    elif 'Wage Basis' in employee_info and employee_info['Wage Basis'].lower() == 'salary':
                        if 'Rate' in employee_info:
                            rate = round(employee_info['Rate'], 2)
                            employee_row['Rate'] = rate

                            # Calculate Gross Pay for salary employees using the modified formula:
                            # Gross = Rate + Wage Owed + Bonus
                            # NOTE: Reimbursements is no longer included in this calculation
                            wage_owed = employee_row['Wage Owed'] or 0
                            bonus = employee_row['Bonus'] or 0

                            gross_pay = rate + wage_owed + bonus
                            employee_row['Gross Pay'] = round(gross_pay, 2)
                        else:
                            employee_row['Gross Pay'] = 0
                            self.update_signal.emit(f"Warning: No rate found for salaried employee '{employee}' at '{location}'")
                    else:
                        if 'Wage Basis' in employee_info:
                            self.update_signal.emit(f"Warning: Unknown wage basis '{employee_info['Wage Basis']}' for '{employee}' at '{location}'")
                        else:
                            self.update_signal.emit(f"Warning: No wage basis found for '{employee}' at '{location}'")
                        employee_row['Gross Pay'] = 0
                else:
                    # Employee not found in payroll dictionary
                    employee_row['Gross Pay'] = 0
                    employee_row['Wage Basis'] = 'Unknown'
                    employee_row['Rate'] = 0
                    employee_row['ID'] = employee  # Just use employee name as ID if no payroll entry
                    employee_row['Holiday Holiday'] = holiday_hours
                    employee_row['Regular Hours'] = total_hours - overtime_hours - holiday_hours

                    # Add to missing employees list
                    missing_employees.append(f"'{employee}' at '{location}'")

                # Add placeholder values for columns we'll calculate later
                employee_row['Tips'] = 0
                employee_row['Gross + Tips'] = employee_row['Gross Pay']  # Just Gross Pay until Tips are implemented

                # Add row to the location's employee list instead of directly to result_df
                all_employees_by_location[location].append(employee_row)

                # Increment employee number for this location
                employee_number += 1

        # Find salaried employees in the dictionary that are not in time entries
        # Group the payroll dictionary by location
        payroll_locations = payroll_dict_df.groupby('Location')
        salaried_without_time = []

        # Process each location separately
        for location, location_group in payroll_locations:
            # Initialize the location entry if not already in the dictionary
            if location not in all_employees_by_location:
                all_employees_by_location[location] = []

            # Get the highest employee number for this location
            employee_number = 1
            if all_employees_by_location[location]:
                employee_numbers = [row['Number'] for row in all_employees_by_location[location] if 'Number' in row]
                if employee_numbers:
                    employee_number = max(employee_numbers) + 1

            # Check each employee in the payroll dictionary
            for index, row in location_group.iterrows():
                employee = row['Employee']

                # Skip if not a salaried employee
                if 'Wage Basis' not in row or pd.isna(row['Wage Basis']) or row['Wage Basis'].lower() != 'salary':
                    continue

                # Skip if already processed (has time entries)
                if (location, employee) in employees_processed:
                    continue

                # Check if employee should be paid
                should_be_paid = True
                if 'PAY?' in row and isinstance(row['PAY?'], str) and row['PAY?'].lower() == 'no':
                    should_be_paid = False
                    not_to_be_paid.append(f"'{employee}' in '{location}' is set NOT to be paid AND they have no time entries")
                    continue  # Skip this employee

                # Found a salaried employee without time entries
                salaried_without_time.append(f"'{employee}' at '{location}'")

                # Create a row for this employee
                employee_row = {col: None for col in result_columns}

                # Populate the basic information
                employee_row['Number'] = employee_number
                employee_row['Num-Loc'] = f"{employee_number}{location}"
                employee_row['Employee'] = employee
                employee_row['Total Hours'] = 0  # No time entries
                employee_row['Location'] = location
                employee_row['Wage Basis'] = 'Salary'

                # Special flag for highlighting
                employee_row['Missing Time Entry'] = True

                # Create ID
                adp_company_code = str(row.get('ADP Company Code', '')) if not pd.isna(row.get('ADP Company Code', '')) else ''
                adp_employee_code = str(row.get('ADP Employee Code', '')) if not pd.isna(row.get('ADP Employee Code', '')) else ''
                employee_id = f"{adp_company_code}{adp_employee_code}{employee}"
                employee_row['ID'] = employee_id

                # Copy all matching columns from payroll dictionary
                for col in result_columns:
                    if col in row and col not in ['Number', 'Num-Loc', 'Total Hours', 'Location', 'ID',
                                               'Raw Overtime', 'Holiday Hours', 'Overtime Hours',
                                               'Regular Hours', 'Gross Pay', 'Gross + Tips', 'Holiday Holiday',
                                               'Missing Time Entry']:
                        employee_row[col] = row[col]

                # Set default values for calculation fields
                for field in ['Tip Adjustment', 'Bonus', 'Wage Owed', 'Reimbursements']:
                    if field not in row or pd.isna(row[field]):
                        employee_row[field] = 0
                    else:
                        employee_row[field] = round(row[field], 2)

                # Set placeholder values for hours-related fields
                employee_row['Holiday Hours'] = 0
                employee_row['Raw Overtime'] = 0
                employee_row['Overtime Hours'] = 0
                employee_row['Regular Hours'] = 0
                employee_row['Holiday Holiday'] = 0

                # Calculate gross pay for salaried employee
                if 'Rate' in row and not pd.isna(row['Rate']):
                    rate = round(row['Rate'], 2)
                    employee_row['Rate'] = rate

                    # Calculate Gross Pay for salary employees
                    wage_owed = employee_row['Wage Owed'] or 0
                    bonus = employee_row['Bonus'] or 0
                    reimbursements = employee_row['Reimbursements'] or 0

                    gross_pay = rate + wage_owed + bonus + reimbursements
                    employee_row['Gross Pay'] = round(gross_pay, 2)
                else:
                    employee_row['Gross Pay'] = 0
                    self.update_signal.emit(f"Warning: No rate found for salaried employee '{employee}' at '{location}'")

                # Add placeholder values
                employee_row['Tips'] = 0
                employee_row['Gross + Tips'] = employee_row['Gross Pay']

                # Add row to the location's employee list instead of directly to result_df
                all_employees_by_location[location].append(employee_row)

                # Increment employee number
                employee_number += 1

        # Log employees not to be paid
        if not_to_be_paid:
            self.update_signal.emit("\nINFO: The following employees are marked NOT to be paid:")
            for emp in not_to_be_paid:
                self.update_signal.emit(f"  - {emp}")
            self.update_signal.emit(f"Total employees marked not to be paid: {len(not_to_be_paid)}")

        # Log salaried employees without time entries
        if salaried_without_time:
            self.update_signal.emit("\nINFO: The following salaried employees were found in payroll dictionary but had no time entries:")
            for emp in salaried_without_time:
                self.update_signal.emit(f"  - {emp}")
            self.update_signal.emit(f"Total salaried employees without time entries: {len(salaried_without_time)}")

        # Log missing employees (excluding the ones in the exclude_employees list)
        filtered_missing_employees = []
        for emp in missing_employees:
            # Extract employee name from the "'{employee}' at '{location}'" format
            employee_name = emp.split("'")[1]
            if employee_name not in exclude_employees:
                filtered_missing_employees.append(emp)

        if filtered_missing_employees:
            self.update_signal.emit("\nWARNING: The following employees were found in time entries but not in the payroll dictionary:")
            for emp in filtered_missing_employees:
                self.update_signal.emit(f"  - {emp}")
            self.update_signal.emit(f"Total missing employees (excluding known entries to ignore): {len(filtered_missing_employees)}")

        # Apply tips to employees in the summary
        tips_applied_count = 0
        for location in all_employees_by_location:
            for employee_row in all_employees_by_location[location]:
                employee_name = employee_row['Employee']

                # Get adjustment value (default to 0 if not present)
                adjustment = employee_row.get('Tip Adjustment', 0)
                if pd.isna(adjustment):
                    adjustment = 0

                # Strip whitespace for matching
                employee_name_stripped = employee_name.strip() if isinstance(employee_name, str) else employee_name
                location_stripped = location.strip() if isinstance(location, str) else location

                # Check for exact match first
                tips_key = (location, employee_name)
                tips_amount = 0

                # Try to find matching tips
                if tips_key in employee_tips:
                    tips_amount = employee_tips[tips_key]
                    employee_tips.pop(tips_key)
                    tips_applied_count += 1
                else:
                    # Try with stripped values
                    tips_key = (location_stripped, employee_name_stripped)

                    # If still not found, try searching all keys with stripped values
                    if tips_key in employee_tips:
                        tips_amount = employee_tips[tips_key]
                        employee_tips.pop(tips_key)
                        tips_applied_count += 1
                    else:
                        found = False
                        for (tip_loc, tip_emp), tip_amount in list(employee_tips.items()):
                            # Strip whitespace from tips file values for comparison
                            tip_loc_stripped = tip_loc.strip() if isinstance(tip_loc, str) else tip_loc
                            tip_emp_stripped = tip_emp.strip() if isinstance(tip_emp, str) else tip_emp

                            if tip_loc_stripped == location_stripped and tip_emp_stripped == employee_name_stripped:
                                tips_amount = tip_amount
                                employee_tips.pop((tip_loc, tip_emp))
                                tips_applied_count += 1
                                found = True
                                break

                # Add adjustment to tips amount (now we DO include the adjustment)
                total_tips = round(tips_amount + adjustment, 2)
                employee_row['Tips'] = total_tips
                employee_row['Gross + Tips'] = round(employee_row['Gross Pay'] + total_tips, 2)

        self.update_signal.emit(f"Applied tips to {tips_applied_count} employees")

        # Check for unmatched tips employees
        for (location, employee), tips_amount in employee_tips.items():
            if tips_amount > 0:  # Only warn about non-zero tips
                unmatched_tips_employees.append(f"'{employee}' at '{location}' has ${tips_amount:.2f} in tips")

        # Create the result_df by adding employees from each location in order
        all_employee_rows = []
        for location in sorted(all_employees_by_location.keys()):
            location_employees = all_employees_by_location[location]
            # Sort employees within each location by Number
            location_employees.sort(key=lambda x: x['Number'])
            all_employee_rows.extend(location_employees)

        # Create result_df from all rows at once
        result_df = pd.DataFrame(all_employee_rows)

        # Keep track of rows that need yellow highlighting (missing time entry)
        missing_time_entry_rows = []
        if 'Missing Time Entry' in result_df.columns:
            missing_time_entry_rows = result_df[result_df['Missing Time Entry'] == True].index.tolist()

        # Remove the calculation and temporary columns for final output
        final_columns = [
            'Location', 'Number', 'Num-Loc', 'ID', 'Code', 'ADP Company Code',
            'ADP Employee Code', 'Employee', 'Total Hours', 'Gross Pay',
            'Tips', 'Gross + Tips', 'Holiday Holiday', 'Regular Hours',
            'Overtime Hours', 'Wage Basis', 'Rate', 'Tip Adjustment', 'Bonus',
            'Wage Owed', 'Reimbursements', 'Comment',
            'Recordatorio futuros'
        ]
        result_df = result_df[final_columns]

        self.update_signal.emit(f"Processed {len(result_df)} employee records")

        return result_df, filtered_missing_employees, missing_time_entry_rows, unmatched_tips_employees, salaried_without_time, not_to_be_paid

    def create_adp_cargue_file(self, result_df, output_dir, timestamp):
        import pandas as pd
        """
        Create the ADP_Cargue CSV file based on the summary data.

        Args:
            result_df (pandas.DataFrame): The summary data
            output_dir (str): Directory to save the output file
            timestamp (str): Timestamp for the filename

        Returns:
            str: Path to the created file
        """
        self.update_signal.emit("Creating ADP_Cargue file...")

        # Create a list to hold the rows for the ADP_Cargue file
        adp_rows = []

        # Process each employee in the summary file
        for _, row in result_df.iterrows():
            # Skip employees with ADP Employee Code of "QB", "QBS", or "Run"
            adp_employee_code = row.get('ADP Employee Code', '')
            if pd.isna(adp_employee_code) or adp_employee_code in ["QB", "QBS", "Run"]:
                continue

            # Get values from summary
            co_code = row.get('ADP Company Code', '')
            file_num = adp_employee_code
            reg_hours = row.get('Regular Hours', 0)
            ot_hours = row.get('Overtime Hours', 0)
            holiday_hours = row.get('Holiday Holiday', 0)
            tips = row.get('Tips', 0)

            # Get values for new columns
            reimbursements = row.get('Reimbursements', 0)
            bonus = row.get('Bonus', 0)
            wage_owed = row.get('Wage Owed', 0)

            # Create Batch ID
            batch_id = f"PR{co_code}EPI"

            # Create first row (Pay # = 1)
            adp_rows.append({
                'Co Code': co_code,
                'Batch ID': batch_id,
                'File #': file_num,
                'Pay #': 1,
                'Reg Hours': round(reg_hours, 2),
                'O/T Hours': round(ot_hours, 2),
                'Earnings 3 Code': 'T',
                'Earnings 3 Amount': tips,  # Now this is already tips + tip_adjustment
                'Adjust Ded Code': 'NTR',
                'Adjust Ded Amount': -1 * reimbursements,  # Get reimbursements from row and negate it
                'Earnings 5 Code': 'BN',
                'Earnings 5 Amount': bonus,
                'Reg Earnings': wage_owed
            })

            # If employee has holiday hours, create second row (Pay # = 2)
            if holiday_hours > 0:
                adp_rows.append({
                    'Co Code': co_code,
                    'Batch ID': batch_id,
                    'File #': file_num,
                    'Pay #': 2,
                    'Reg Hours': 0,  # Always 0 for Pay # = 2
                    'O/T Hours': round(holiday_hours, 2),
                    'Earnings 3 Code': 'T',
                    'Earnings 3 Amount': 0,  # Tips only included in first row
                    'Adjust Ded Code': 'NTR',
                    'Adjust Ded Amount': 0,  # Only include in first row
                    'Earnings 5 Code': 'BN',
                    'Earnings 5 Amount': 0,  # Only include in first row
                    'Reg Earnings': 0  # Only include in first row
                })

        # Create DataFrame from rows
        adp_df = pd.DataFrame(adp_rows)

        # Define output file path
        output_file = os.path.join(output_dir, f"ADP_Cargue_{timestamp}.csv")

        # Write to CSV file without index
        adp_df.to_csv(output_file, index=False)

        self.update_signal.emit(f"ADP_Cargue file saved to: {output_file}")
        self.update_signal.emit(f"Created {len(adp_rows)} rows for {len(adp_df['File #'].unique())} employees")

        return output_file

    def run(self):
        import pandas as pd
        try:
            # Create timestamp for the output file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Define output path
            output_dir = self.output_dir

            # Create ADP Cargue directory with timestamp inside the output folder
            adp_cargue_dir = os.path.join(output_dir, f"ADP_Cargue_{timestamp}")
            os.makedirs(adp_cargue_dir, exist_ok=True)

            # Use the provided file paths directly
            time_entries_path = self.time_entries_path
            payroll_dict_path = self.payroll_dict_path

            # Load tips if provided
            employee_tips = {}
            tips_df = None

            if self.tips_path:
                tips_path = self.tips_path
                self.update_signal.emit(f"Found tips file: {os.path.basename(tips_path)}")

                # Load the raw tips data for the Excel tab
                try:
                    tips_df = pd.read_csv(tips_path, encoding='cp1252')
                    self.update_signal.emit(f"Loaded tips file with {len(tips_df)} rows for Excel tab")
                except Exception as e:
                    self.update_signal.emit(f"Warning: Could not load tips file for Excel tab: {str(e)}")

                # Process tips data
                employee_tips = self.load_tips_file(tips_path)
            else:
                self.update_signal.emit("Note: No tips file provided, tips will be set to 0 for all employees")

            # Load the data
            time_entries_df = self.load_time_entries(time_entries_path)
            payroll_dict_df = self.load_payroll_dictionary(payroll_dict_path)

            # Calculate the gross pay without ADP data
            result_df, missing_employees, missing_time_entry_rows, unmatched_tips_employees, salaried_without_time, not_to_be_paid = self.calculate_gross_pay(
                time_entries_df, payroll_dict_df, employee_tips)  # Remove the input_dir parameter

            # Define the summary output file path
            summary_output_file = os.path.join(output_dir, f"Payroll_Summary_{timestamp}.xlsx")

            # Create a Pandas Excel writer using XlsxWriter
            with pd.ExcelWriter(summary_output_file, engine='xlsxwriter') as writer:
                # Write the DataFrame to the Excel file
                result_df.to_excel(writer, sheet_name='Summary', index=False)

                # Add the TimeEntries as a separate tab
                time_entries_df.to_excel(writer, sheet_name='TimeEntries', index=False)

                # Add the Payroll Dictionary as a separate tab
                # Round numeric columns in payroll dictionary before writing to Excel
                numeric_columns = ['Rate', 'Tip Adjustment', 'Bonus', 'Wage Owed', 'Reimbursements']
                for col in numeric_columns:
                    if col in payroll_dict_df.columns:
                        payroll_dict_df[col] = payroll_dict_df[col].round(2)
                payroll_dict_df.to_excel(writer, sheet_name='Payroll Dictionary', index=False)

                # Add the Tips file as a separate tab if available
                if tips_df is not None:
                    # Limit to columns up to "BE" (index 56)
                    max_col_index = 56  # Column BE
                    num_cols = min(tips_df.shape[1], max_col_index + 1)

                    # Write the tips data without pandas-generated index or header
                    worksheet = writer.book.add_worksheet('Tips')

                    # Write each row of the dataframe directly, but only up to column BE
                    for r_idx, row in tips_df.iterrows():
                        for c_idx in range(num_cols):
                            if c_idx < len(row):
                                val = row[c_idx]
                                if pd.isna(val):
                                    worksheet.write(r_idx, c_idx, "")
                                elif isinstance(val, (int, float)) and not pd.isna(val):
                                    try:
                                        worksheet.write_number(r_idx, c_idx, val)
                                    except TypeError:
                                        worksheet.write_string(r_idx, c_idx, str(val))
                                else:
                                    worksheet.write_string(r_idx, c_idx, str(val) if not pd.isna(val) else "")

                # Add a new tab for Location Gross Pay Summary
                # Filter out rows with QB or QBS in ADP Employee Code
                filtered_result_df = result_df[~result_df['ADP Employee Code'].isin(["QB", "QBS"])]

                # Group by Location and sum Gross + Tips
                location_pay_summary = filtered_result_df.groupby(['Location', 'ADP Company Code'])['Gross + Tips'].sum().reset_index()

                # Create a dictionary to map locations that need to be combined
                location_mapping = {
                    'Aventura (Miami Gardens)': 'Carrot Love LLC',
                    'North Beach': 'Carrot Love LLC',
                    'Coral Gables': 'Carrot Love LLC'
                }

                # Apply the mapping to create a new column with the correct location names
                location_pay_summary['Mapped Location'] = location_pay_summary['Location'].map(
                    lambda x: location_mapping.get(x, x)
                )

                # Group by the mapped location and sum the gross + tips pay
                location_pay_summary = location_pay_summary.groupby(['Mapped Location', 'ADP Company Code'])['Gross + Tips'].sum().reset_index()
                location_pay_summary.columns = ['Location', 'ADP Company Code', 'Total Gross + Tips Pay']

                # Round the Total Gross + Tips Pay column to 2 decimal places
                location_pay_summary['Total Gross + Tips Pay'] = location_pay_summary['Total Gross + Tips Pay'].round(2)

                # Sort by location name
                location_pay_summary = location_pay_summary.sort_values('Location')

                # Add to Excel
                location_pay_summary.to_excel(writer, sheet_name='Location Pay Summary', index=False)

                # Create ADP_Cargue data
                adp_rows = []

                # Process each employee in the summary file
                for _, row in result_df.iterrows():
                    # Skip employees with ADP Employee Code of "QB", "QBS", or "Run"
                    adp_employee_code = row.get('ADP Employee Code', '')
                    if pd.isna(adp_employee_code) or adp_employee_code in ["QB", "QBS", "Run"]:
                        continue

                    # Get values from summary
                    co_code = row.get('ADP Company Code', '')
                    file_num = adp_employee_code
                    reg_hours = row.get('Regular Hours', 0)
                    ot_hours = row.get('Overtime Hours', 0)
                    holiday_hours = row.get('Holiday Holiday', 0)
                    tips = row.get('Tips', 0)

                    # Get values for new columns
                    reimbursements = row.get('Reimbursements', 0)
                    bonus = row.get('Bonus', 0)
                    wage_owed = row.get('Wage Owed', 0)

                    # Create Batch ID
                    batch_id = f"PR{co_code}EPI"

                    # Create first row (Pay # = 1)
                    adp_rows.append({
                        'Co Code': co_code,
                        'Batch ID': batch_id,
                        'File #': file_num,
                        'Pay #': 1,
                        'Reg Hours': round(reg_hours, 2),
                        'O/T Hours': round(ot_hours, 2),
                        'Earnings 3 Code': 'T',
                        'Earnings 3 Amount': tips,  # Now this is already tips + tip_adjustment
                        'Adjust Ded Code': 'NTR',
                        'Adjust Ded Amount': -1 * reimbursements,  # Use negative of reimbursements
                        'Earnings 5 Code': 'BN',
                        'Earnings 5 Amount': bonus,
                        'Reg Earnings': wage_owed
                    })

                    # If employee has holiday hours, create second row (Pay # = 2)
                    if holiday_hours > 0:
                        adp_rows.append({
                            'Co Code': co_code,
                            'Batch ID': batch_id,
                            'File #': file_num,
                            'Pay #': 2,
                            'Reg Hours': 0,  # Always 0 for Pay # = 2
                            'O/T Hours': round(holiday_hours, 2),
                            'Earnings 3 Code': 'T',
                            'Earnings 3 Amount': 0,  # Tips only included in first row
                            'Adjust Ded Code': 'NTR',
                            'Adjust Ded Amount': 0,  # Only include in first row
                            'Earnings 5 Code': 'BN',
                            'Earnings 5 Amount': 0,  # Only include in first row
                            'Reg Earnings': 0  # Only include in first row
                        })

                # Create DataFrame from rows
                adp_df = pd.DataFrame(adp_rows)

                # Add the ADP_Cargue data as a separate tab
                adp_df.to_excel(writer, sheet_name='ADP_Cargue', index=False)

                # Create individual ADP Cargue files for each Batch ID
                batch_ids = adp_df['Batch ID'].unique()
                self.update_signal.emit(f"Creating {len(batch_ids)} individual ADP Cargue files, one for each Batch ID")

                for batch_id in batch_ids:
                    # Filter the dataframe for just this batch ID
                    batch_df = adp_df[adp_df['Batch ID'] == batch_id]

                    # Create the output file with the batch ID as the filename
                    batch_file = os.path.join(adp_cargue_dir, f"{batch_id}.csv")

                    # Save to CSV
                    batch_df.to_csv(batch_file, index=False)
                    self.update_signal.emit(f"Created ADP Cargue file for {batch_id}")

                # Get the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                workbook.nan_inf_to_errors = True
                worksheet = writer.sheets['Summary']

                # Define formats for different row types with borders added
                red_format = workbook.add_format({
                    'bg_color': '#FFC7CE',
                    'font_color': '#9C0006',
                    'border': 1  # Add border to all cells
                })
                light_blue_format = workbook.add_format({
                    'bg_color': '#DDEBF7',
                    'border': 1  # Add border to all cells
                })
                light_orange_format = workbook.add_format({
                    'bg_color': '#FCE4D6',
                    'border': 1  # Add border to all cells
                })
                light_yellow_format = workbook.add_format({
                    'bg_color': '#FFEB9C',
                    'border': 1  # Add border to all cells
                })

                # Also add a header format with borders and bold
                header_format = workbook.add_format({
                    'bold': True,
                    'border': 1,
                    'bg_color': '#D9D9D9'  # Light gray background for headers
                })

                # Apply the header format to the first row
                for col_idx, col_name in enumerate(result_df.columns):
                    worksheet.write(0, col_idx, col_name, header_format)

                # Apply conditional formatting based on location and missing employees
                unique_locations = result_df['Location'].unique()

                # Dictionary to track which color to use for each location
                location_colors = {}
                for idx, loc in enumerate(unique_locations):
                    location_colors[loc] = 0 if idx % 2 == 0 else 1  # Alternate between 0 (orange) and 1 (blue)

                # Get missing employees for red highlighting
                missing_employee_indices = []
                if missing_employees:
                    for row_idx, row in result_df.iterrows():
                        employee = row['Employee']
                        location = row['Location']

                        # Check if this employee-location combo is in missing_employees and not in exclude_employees
                        if f"'{employee}' at '{location}'" in missing_employees and employee not in exclude_employees:
                            missing_employee_indices.append(row_idx)

                # Format all rows safely
                for row_idx, row in result_df.iterrows():
                    # Get row format based on conditions
                    is_missing = row_idx in missing_employee_indices
                    is_salaried_without_time = row_idx in missing_time_entry_rows

                    if is_missing:
                        # Use red format for missing employees (highest priority)
                        row_format = red_format
                    elif is_salaried_without_time:
                        # Use yellow format for salaried employees without time entries (second priority)
                        row_format = light_yellow_format
                    else:
                        # Use location-based format (lowest priority)
                        location = row['Location']
                        color_idx = location_colors[location]
                        row_format = light_orange_format if color_idx == 0 else light_blue_format

                    # Apply format to each cell in the row
                    for col_idx, col_name in enumerate(result_df.columns):
                        cell_value = row[col_name]

                        # Handle numeric values separately to avoid NaN/INF errors
                        if pd.isna(cell_value):
                            # Write empty string for NaN values
                            worksheet.write_string(row_idx + 1, col_idx, "", row_format)
                        elif isinstance(cell_value, (int, float)) and not pd.isna(cell_value) and pd.notna(cell_value):
                            # For valid numbers, use write_number
                            try:
                                worksheet.write_number(row_idx + 1, col_idx, cell_value, row_format)
                            except TypeError:
                                # If there's an error (like INF), convert to string
                                worksheet.write_string(row_idx + 1, col_idx, str(cell_value), row_format)
                        else:
                            # For everything else, convert to string
                            worksheet.write_string(row_idx + 1, col_idx, str(cell_value) if not pd.isna(cell_value) else "", row_format)

                # Set left alignment for ADP Employee Code column
                adp_employee_code_col = result_df.columns.get_loc('ADP Employee Code')
                worksheet.set_column(adp_employee_code_col, adp_employee_code_col, None, {'align': 'left'})

                # Auto-adjust column width
                for i, col in enumerate(result_df.columns):
                    # Convert values to string to get their length
                    col_values = result_df[col].astype(str)
                    # Find maximum length (but only check first 100 rows to avoid performance issues)
                    max_len = max(
                        col_values.iloc[:100].map(len).max() if len(col_values) > 0 else 0,
                        len(str(col))
                    ) + 1  # Add a little extra space

                    worksheet.set_column(i, i, max_len)

                # Format the Location Pay Summary tab
                if 'Location Pay Summary' in writer.sheets:
                    worksheet = writer.sheets['Location Pay Summary']
                    # Apply header format
                    for col_idx, col_name in enumerate(location_pay_summary.columns):
                        worksheet.write(0, col_idx, col_name, header_format)

                    # Apply cell formats and currency format for amount column
                    money_format = workbook.add_format({
                        'border': 1,
                        'num_format': '$#,##0.00'  # Currency format
                    })
                    cell_format = workbook.add_format({'border': 1})

                    # Apply formats to each cell
                    for row_idx, row in location_pay_summary.iterrows():
                        worksheet.write(row_idx + 1, 0, row['Location'], cell_format)
                        worksheet.write(row_idx + 1, 1, row['ADP Company Code'], cell_format)
                        worksheet.write_number(row_idx + 1, 2, row['Total Gross + Tips Pay'], money_format)

                    # Set column widths
                    worksheet.set_column(0, 0, 25)  # Location column
                    worksheet.set_column(1, 1, 15)  # ADP Company Code column
                    worksheet.set_column(2, 2, 20)  # Total Gross + Tips Pay column

            # Create a new Excel file with all warnings and information
            warnings_file = os.path.join(output_dir, f"Payroll_Warnings_{timestamp}.xlsx")

            with pd.ExcelWriter(warnings_file, engine='xlsxwriter') as writer:
                workbook = writer.book

                # Create Dashboard worksheet first (as the main tab)
                worksheet = workbook.add_worksheet("Dashboard")

                # Set column widths
                worksheet.set_column(0, 0, 25)
                worksheet.set_column(1, 1, 15)
                worksheet.set_column(2, 2, 15)

                # Format for title
                title_format = workbook.add_format({
                    'bold': True, 'font_size': 16, 'align': 'center',
                    'valign': 'vcenter', 'font_color': '#4472C4'
                })
                worksheet.merge_range('A1:C2', f"Payroll Processing Summary", title_format)

                # Format headers
                header_format = workbook.add_format({
                    'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
                    'border': 1, 'text_wrap': True, 'align': 'center'
                })

                # Format for warning status
                warning_format = workbook.add_format({
                    'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center'
                })

                # Format for OK status
                ok_format = workbook.add_format({
                    'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center'
                })

                # Format for info status
                info_format = workbook.add_format({
                    'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'border': 1, 'align': 'center'
                })

                # Regular cell format
                cell_format = workbook.add_format({
                    'border': 1, 'align': 'center'
                })

                # Summary data
                summary_data = [
                    ["Category", "Count", "Status"],
                    ["Total Employees", len(result_df), "Processed"],
                    ["Missing From Dictionary", len(missing_employees), "Warning" if missing_employees else "OK"],
                    ["Salaried Without Time", len(salaried_without_time), "Info"],
                    ["Not To Be Paid", len(not_to_be_paid), "Info"],
                    ["Unmatched Tips", len(unmatched_tips_employees), "Warning" if unmatched_tips_employees else "OK"],
                ]

                # Write the summary data
                for row_idx, row_data in enumerate(summary_data):
                    for col_idx, cell_data in enumerate(row_data):
                        if row_idx == 0:
                            # Header row
                            worksheet.write(row_idx + 3, col_idx, cell_data, header_format)
                        else:
                            if col_idx == 2:  # Status column
                                if cell_data == "Warning":
                                    worksheet.write(row_idx + 3, col_idx, cell_data, warning_format)
                                elif cell_data == "OK":
                                    worksheet.write(row_idx + 3, col_idx, cell_data, ok_format)
                                else:
                                    worksheet.write(row_idx + 3, col_idx, cell_data, info_format)
                            else:
                                worksheet.write(row_idx + 3, col_idx, cell_data, cell_format)

                # 1. Missing Employees (employees in time entries but not in dictionary)
                if missing_employees:
                    missing_df = pd.DataFrame({
                        'Employee': [emp.split("'")[1] for emp in missing_employees],
                        'Location': [emp.split("'")[3] for emp in missing_employees],
                        'Issue': ['Missing from payroll dictionary'] * len(missing_employees)
                    })
                    missing_df.to_excel(writer, sheet_name='Missing Employees', index=False)

                    # Format the worksheet
                    worksheet = writer.sheets['Missing Employees']
                    header_format = workbook.add_format({
                        'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'text_wrap': True, 'align': 'center'
                    })
                    for col_num, value in enumerate(missing_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, 20)

                # 2. Salaried Without Time
                if salaried_without_time:
                    salaried_df = pd.DataFrame({
                        'Employee': [emp.split("'")[1] for emp in salaried_without_time],
                        'Location': [emp.split("'")[3] for emp in salaried_without_time],
                        'Issue': ['Salaried employee without time entries'] * len(salaried_without_time)
                    })
                    salaried_df.to_excel(writer, sheet_name='Salaried No Time', index=False)

                    # Format the worksheet
                    worksheet = writer.sheets['Salaried No Time']
                    header_format = workbook.add_format({
                        'bold': True, 'bg_color': '#FFEB9C', 'border': 1, 'text_wrap': True, 'align': 'center'
                    })
                    for col_num, value in enumerate(salaried_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, 20)

                # 3. Not To Be Paid
                if not_to_be_paid:
                    notpaid_df = pd.DataFrame({
                        'Details': not_to_be_paid,
                        'Issue': ['Marked NOT to be paid'] * len(not_to_be_paid)
                    })
                    notpaid_df.to_excel(writer, sheet_name='Not To Be Paid', index=False)

                    # Format the worksheet
                    worksheet = writer.sheets['Not To Be Paid']
                    header_format = workbook.add_format({
                        'bold': True, 'bg_color': '#F8CBAD', 'border': 1, 'text_wrap': True, 'align': 'center'
                    })
                    for col_num, value in enumerate(notpaid_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, 50)

                # 4. Unmatched Tips
                if unmatched_tips_employees:
                    tips_df = pd.DataFrame({
                        'Details': unmatched_tips_employees,
                        'Issue': ['Unmatched tips'] * len(unmatched_tips_employees)
                    })
                    tips_df.to_excel(writer, sheet_name='Unmatched Tips', index=False)

                    # Format the worksheet
                    worksheet = writer.sheets['Unmatched Tips']
                    header_format = workbook.add_format({
                        'bold': True, 'bg_color': '#BDD7EE', 'border': 1, 'text_wrap': True, 'align': 'center'
                    })
                    for col_num, value in enumerate(tips_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, 50)

                # 5. Excluded Employees Tab (new)
                excluded_df = pd.DataFrame({
                    'Employee': exclude_employees,
                    'Issue': ['Excluded from processing'] * len(exclude_employees)
                })
                excluded_df.to_excel(writer, sheet_name='Excluded Employees', index=False)

                # Format the worksheet
                worksheet = writer.sheets['Excluded Employees']
                header_format = workbook.add_format({
                    'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'text_wrap': True, 'align': 'center'
                })
                for col_num, value in enumerate(excluded_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 30)

            self.update_signal.emit(f"Payroll summary with all raw data saved to: {summary_output_file}")
            self.update_signal.emit(f"ADP Cargue files saved to: {adp_cargue_dir}")
            self.update_signal.emit(f"Payroll warnings and information saved to: {warnings_file}")

            # Create Workers Comp file
            # Define the workers comp file path
            workers_comp_file = os.path.join(output_dir, f"Workers_Comp_{timestamp}.xlsx")

            # Filter out employees with ADP Employee Code of "QB" or "QBS"
            filtered_result_df = result_df[~result_df['ADP Employee Code'].isin(["QB", "QBS"])]

            # Group by Location and sum Gross Pay
            location_gross_pay = filtered_result_df.groupby('Location')['Gross Pay'].sum().reset_index()

            # Create a dictionary to map locations that need to be combined
            location_mapping = {
                'Aventura (Miami Gardens)': 'Carrot Love LLC',
                'North Beach': 'Carrot Love LLC',
                'Coral Gables': 'Carrot Love LLC'
            }

            # Apply the mapping to create a new column with the correct location names
            location_gross_pay['Mapped Location'] = location_gross_pay['Location'].map(
                lambda x: location_mapping.get(x, x)
            )

            # Group by the mapped location and sum the gross pay
            workers_comp_data = location_gross_pay.groupby('Mapped Location')['Gross Pay'].sum().reset_index()
            workers_comp_data.columns = ['Location', 'Total Gross Pay']

            # Round the Total Gross Pay column to 2 decimal places
            workers_comp_data['Total Gross Pay'] = workers_comp_data['Total Gross Pay'].round(2)

            # Sort by location name
            workers_comp_data = workers_comp_data.sort_values('Location')

            # Create a Pandas Excel writer
            with pd.ExcelWriter(workers_comp_file, engine='xlsxwriter') as writer:
                # Write the DataFrame to the Excel file
                workers_comp_data.to_excel(writer, sheet_name='Workers Comp', index=False)

                # Get the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Workers Comp']

                # Format for header
                header_format = workbook.add_format({
                    'bold': True,
                    'border': 1,
                    'bg_color': '#D9D9D9'  # Light gray background for headers
                })

                # Format for data cells
                cell_format = workbook.add_format({
                    'border': 1
                })

                # Format for numeric values
                number_format = workbook.add_format({
                    'border': 1,
                    'num_format': '$#,##0.00'  # Currency format
                })

                # Apply formats
                for col_idx, col_name in enumerate(workers_comp_data.columns):
                    worksheet.write(0, col_idx, col_name, header_format)

                # Apply formats to data rows
                for row_idx, row in workers_comp_data.iterrows():
                    worksheet.write(row_idx + 1, 0, row['Location'], cell_format)
                    worksheet.write(row_idx + 1, 1, row['Total Gross Pay'], number_format)

                # Set column widths
                worksheet.set_column(0, 0, 25)  # Location column
                worksheet.set_column(1, 1, 15)  # Total Gross Pay column

            self.update_signal.emit(f"Workers Comp file saved to: {workers_comp_file}")

            # Log unmatched tips employees
            if unmatched_tips_employees:
                self.update_signal.emit("\nWARNING: The following employees with tips could not be matched to the summary file:")
                for emp in unmatched_tips_employees:
                    self.update_signal.emit(f"  - {emp}")
                self.update_signal.emit(f"Total unmatched employees with tips: {len(unmatched_tips_employees)}")
            else:
                self.update_signal.emit("\nINFO: All tips matched to an employee")

            self.finished_signal.emit(True, f"Payroll automation completed successfully! Files saved to the output directory.")

        except Exception as e:
            import traceback
            self.update_signal.emit(f"Error during processing: {str(e)}")
            self.update_signal.emit(f"Error details: {traceback.format_exc()}")
            self.finished_signal.emit(False, f"An error occurred: {str(e)}")
