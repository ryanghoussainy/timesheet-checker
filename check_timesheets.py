import pandas as pd
import datetime


def normalise_time(time_str: str) -> str:
    """
    Normalise time string to "HH:MM" format.
    """
    if isinstance(time_str, datetime.time):
        return time_str.strftime("%H:%M")
    return time_str


def read_timesheet(file_path: str) -> list[tuple[str, str, str, str]]:
    """
    Read a timesheet excel file and return a list of tuples of
       (number of hours, start time, end time, rate).
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
    print(table_df)
    
    # Create a tuple of (number of hours, start time, end time, rate)
    timesheet_data = []
    location_list = ["Masters", "NP", "Chiswick", "Sh. Bush", "Acton", "Water", "Horsenden ", "St Helen's", "Disability", "Squad", "Admin", "Safeguarding "]
    for _, row in table_df.iterrows():
        location_hours = [row[loc] if pd.notna(row[loc]) else 0 for loc in location_list]
        timesheet_data.append((
            sum(location_hours),
            row["Date"],
            row['Rate of pay']
        ))

    print(df.iloc[3, 2] + " " + df.iloc[4, 2], (timesheet_data))


def check_timesheets(file_paths: list[str]):
    """
    Read all timesheet excel files and print any discrepancies found with the sign in sheet.
    """
    pass