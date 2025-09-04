import pandas as pd
import datetime

from read_sign_in import read_sign_in_sheet
from discrepancies import print_discrepancies, EMPTY_TIMESHEET, INVALID_NAME, TIMESHEET_EXTRA_ENTRY, SIGN_IN_EXTRA_ENTRY


def read_timesheet(df) -> dict[str, tuple[str, datetime.datetime, str]]:
    """
    Read a timesheet excel file and return a list of tuples of
       (number of hours, date, rate).
    """
    
    # Get the row number whose first column is 'Date'
    header_row_index = df[df.iloc[:, 0] == 'Date'].index[0]

    # From df, get all rows after the header row, including the header row
    table_df = df.iloc[header_row_index:]
    
    # Reset the index of the dataframe
    table_df.reset_index(drop=True, inplace=True)
    table_df.columns = table_df.iloc[0]
    table_df = table_df[1:]

    # Remove all rows where the Date, Week day, start time, end time are Nan
    table_df.dropna(subset=['Date', 'Week day', 'Start Time', 'End Time'], inplace=True)
    table_df.reset_index(drop=True, inplace=True)
    
    # Create a tuple of (number of hours, start time, end time, rate)
    timesheet_data = []
    location_list = ["Acton hours", "Admin hours", "Safeguarding hours", "GALA day rate", "House Event day rate"]
    for _, row in table_df.iterrows():
        location_hours = [float(row[loc]) if pd.notna(row[loc]) else 0 for loc in location_list]
        # TODO: error if multiple location hours are not 0
        timesheet_data.append((
            str(sum(location_hours)),
            row["Date"].date(),
            row['Rate of pay (see table below)']
        ))

    first_name = str(df.iloc[3, 2]).strip()
    last_name = str(df.iloc[4, 2]).strip()
    name = first_name + " " + last_name

    return name, timesheet_data


def check_timesheets(amindefied_excel_path, sign_in_sheet_path, rates, month):
    # Check for discrepancies
    discrepancies = []

    # Read sign in sheet
    sign_in_data = read_sign_in_sheet(month, sign_in_sheet_path, rates)

    with pd.ExcelFile(amindefied_excel_path) as xls:
        for sheet_name in xls.sheet_names:
            # Read individual timesheet
            df = pd.read_excel(xls, sheet_name=sheet_name)
            if df.empty:
                discrepancies.append((EMPTY_TIMESHEET, {"sheet name": sheet_name}))

            check_timesheet(df, sign_in_data, discrepancies)
    
    # Check for remaining entries in sign in data
    for name, entries in sign_in_data.items():
        for entry in entries:
            discrepancies.append((SIGN_IN_EXTRA_ENTRY, {"name": name, "entry": entry}))
    
    print_discrepancies(discrepancies)


def check_timesheet(df, sign_in_data: dict[str, set[tuple[str, datetime.datetime, str]]], discrepancies):
    """
    Check a single timesheet against the sign in data and print any discrepancies found.
    """
    # Read the timesheet
    name, data = read_timesheet(df)

    # Check if timesheet name is correct
    if name not in sign_in_data:
        sign_in_names = list(sign_in_data.keys())
        discrepancies.append((INVALID_NAME, {"name": name, "sign in names": sign_in_names}))
    else:
        # Make sets for comparison to not modify the original data
        data_set = set(data)
        sign_in_set = sign_in_data[name]

        # For each entry in the timesheet data, match and remove from the sign in data
        print(f"\nChecking timesheet for {name}...")
        for entry in data:
            if entry not in sign_in_set:
                discrepancies.append((TIMESHEET_EXTRA_ENTRY, {"name": name, "entry": entry}))
            else: 
                # Successfully matched entry
                sign_in_set.remove(entry)
                data_set.remove(entry)
