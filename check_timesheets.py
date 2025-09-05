import pandas as pd

from read_sign_in import read_sign_in_sheet
from discrepancies import print_discrepancies, EMPTY_TIMESHEET, INVALID_NAME, TIMESHEET_EXTRA_ENTRY, SIGN_IN_EXTRA_ENTRY
from entry import Entry


def read_timesheet(df) -> tuple[str, list[Entry]]:
    """
    Read a timesheet excel file and return a set of entries
    """
    # Get the name
    first_name = str(df.iloc[3, 2]).strip()
    last_name = str(df.iloc[4, 2]).strip()
    name = first_name + " " + last_name
    
    # Get the row number whose first column is 'Date'
    header_row_index = df[df.iloc[:, 0] == 'Date'].index[0]

    # From df, get all rows after the header row, including the header row
    table_df = df.iloc[header_row_index:]
    
    # Reset the index of the dataframe
    table_df.reset_index(drop=True, inplace=True)
    table_df.columns = table_df.iloc[0]
    table_df = table_df[1:]

    # Filter rows by those that have a date
    table_df.dropna(subset=['Date', 'Week day'], inplace=True)
    table_df.reset_index(drop=True, inplace=True)
    
    # Create a tuple of (number of hours, start time, end time, rate)
    timesheet_data = []
    location_list = ["Acton hours", "Admin hours", "Safeguarding hours", "GALA day rate", "House Event day rate"]
    for _, row in table_df.iterrows():
        location_hours = [loc for loc in location_list if pd.notna(row[loc])]

        if not location_hours:
            raise ValueError(f"No hours found for {name} on {row['Date']}")
        if len(location_hours) > 1:
            raise ValueError(f"Multiple hours found in a single row for {name} on {row['Date']}")
        
        entry = Entry(
            date=row["Date"].date(),
            hours=float(row[location_hours[0]]),
            rate=row['Rate of pay (see table below)'],
        )
        timesheet_data.append(entry)

    return name, timesheet_data


def check_timesheets(
    amindefied_excel_path,
    sign_in_sheet_path, rates,
    rates_after,
    rate_change_date,
    month,
):
    # Check for discrepancies
    discrepancies = []

    # Read sign in sheet
    sign_in_data = read_sign_in_sheet(month, sign_in_sheet_path, rates, rates_after, rate_change_date)

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


def check_timesheet(df, sign_in_data: dict[str, set[Entry]], discrepancies):
    """
    Check a single timesheet against the sign in data and print any discrepancies found.
    """
    # Read the timesheet
    name, timesheet_entries = read_timesheet(df)

    # Check if timesheet name is correct
    if name not in sign_in_data:
        sign_in_names = list(sign_in_data.keys())
        discrepancies.append((INVALID_NAME, {"name": name, "sign in names": sign_in_names}))
    else:
        # Make sets for comparison to not modify the original data
        timesheet_set = set(timesheet_entries)
        sign_in_set = sign_in_data[name]

        # For each entry in the timesheet data, match and remove from the sign in data
        print(f"\nChecking timesheet for {name}...")
        for entry in timesheet_entries:
            if entry not in sign_in_set:
                discrepancies.append((TIMESHEET_EXTRA_ENTRY, {"name": name, "entry": entry}))
            else:
                # Successfully matched entry
                sign_in_set.remove(entry)
                timesheet_set.remove(entry)
