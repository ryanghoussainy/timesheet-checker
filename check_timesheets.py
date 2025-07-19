import os
import pandas as pd
import datetime
from read_sign_in import read_sign_in_sheet


def read_timesheet(file_path: str) -> dict[str, tuple[str, datetime.datetime, str]]:
    """
    Read a timesheet excel file and return a list of tuples of
       (number of hours, date, rate).
    """
    
    # Read the excel file
    df = pd.read_excel(file_path, header=0)
    
    # Get the row number whose first column is 'Date'
    header_row_index = df[df.iloc[:, 0] == 'Date'].index[0]

    # From df, get all rows after the header row, including the header row
    table_df = df.iloc[header_row_index:]
    
    # Reset the index of the dataframe
    table_df.reset_index(drop=True, inplace=True)
    table_df.columns = table_df.iloc[0]
    table_df = table_df[1:]

    # Remove all rows where the Date, Week day, start time, end time are Nan
    table_df.dropna(subset=['Date', 'Week', 'Start', 'End'], inplace=True)
    table_df.reset_index(drop=True, inplace=True)
    
    # Create a tuple of (number of hours, start time, end time, rate)
    timesheet_data = []
    location_list = ["Masters", "NP", "Chiswick", "Sh. Bush", "Acton", "Water", "Horsenden ", "St Helen's", "Disability", "Squad", "Admin", "Safeguarding "]
    for _, row in table_df.iterrows():
        location_hours = [row[loc] if pd.notna(row[loc]) else 0 for loc in location_list]
        timesheet_data.append((
            str(float(sum(location_hours))),
            row["Date"],
            row['Rate of pay']
        ))
        
    name = df.iloc[3, 3]
    
    return name, timesheet_data

    
def check_timesheet(file_path: str, sign_in_data: dict[str, tuple[str, datetime.datetime, str]]):
    """
    Check a single timesheet against the sign in data and print any discrepancies found.
    """
    # Read the timesheet
    name, data = read_timesheet(file_path)

    # Check if the name exists in the sign in data
    if name in sign_in_data:
        
        # Make sets for comparison to not modify the original data
        data_set = set(data)
        sign_in_set = set(sign_in_data[name])

        # For each entry in the timesheet data, match and remove from the sign in data 
        for entry in data:
            if entry in sign_in_set:
                sign_in_set.remove(entry)
                data_set.remove(entry)

        if len(data_set) == len(sign_in_set) == 0:
            print(f"No discrepancies found for {name}.")
        else:
            print(f"Discrepancy found for {name}: {data_set} (timesheet) vs {sign_in_set} (sign in)")
    else:
        print(f"No sign in data found for {name}")

def check_timesheets(timesheet_folder: str, sign_in_sheet_folder: str):
    """
    Read all timesheet excel files and print any discrepancies found with the sign in sheet.
    """
    sign_in_data = read_sign_in_sheet("July", os.path.join(sign_in_sheet_folder, "TestSheet.xlsx"))
    
    for filename in os.listdir(timesheet_folder):
    
        if filename.endswith(".xlsx"):
            file_path = os.path.join(timesheet_folder, filename)
            check_timesheet(file_path, sign_in_data)
