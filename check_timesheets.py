import pandas as pd
import datetime

level_to_rate = {
    "L1": 11.07,
    "L2": 19.65,
    "NQL2": 17.23,
    "Enhanced L2": 22.44,
    "Lower Enhanced L2": 11.90,
    "LHC": 30.00,
}

def normalise_time(time_str: str) -> str:
    """
    Normalise time string to "HH:MM" format.
    """
    if isinstance(time_str, datetime.time):
        return time_str.strftime("%H:%M")
    return time_str


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
    print(name, (timesheet_data))
    return {name: timesheet_data}
    
def read_sign_in_sheet(month: str, file_path: str) -> dict[str, tuple[str, datetime.datetime, str]]:
    """
    Read a sign in sheet excel file and return a list of tuples of
       (number of hours, date, rate).
    """
    global level_to_rate

    sign_df = pd.read_excel(file_path, month, header=0)

    sign_in_sheet_data = {}

    for _, row in sign_df.iterrows():
        if not pd.isna(row['Name']) and row['Name'] != "Total Hours":
            name = row['Name']
            if name not in sign_in_sheet_data:
                sign_in_sheet_data[name] = []
            for col in sign_df.columns[3:]:
                if not pd.isna(row[col]):
                    sign_in_sheet_data[name].append((
                        str(row[col]),
                        col, 
                        level_to_rate[row["Level"]]
                    ))

    print(sign_in_sheet_data)
    return sign_in_sheet_data

def check_timesheets(file_paths: list[str]):
    """
    Read all timesheet excel files and print any discrepancies found with the sign in sheet.
    """
    pass