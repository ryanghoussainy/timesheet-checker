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
    df = pd.read_excel(file_path, header=7)

    # Remove all rows where the Date, Week day, start time, end time are Nan
    df.dropna(subset=['Date', 'Week', 'Start', 'End'], inplace=True)

    # Create a tuple of (number of hours, start time, end time, rate)
    timesheet_data = []
    for _, row in df.iterrows():
        timesheet_data.append((
            row['Acton'],
            normalise_time(row['Start']),
            normalise_time(row['End']),
            row['Rate of pay']
        ))

    print(timesheet_data)


def check_timesheets(file_paths: list[str]):
    """
    Read all timesheet excel files and print any discrepancies found with the sign in sheet.
    """
    pass
