import pandas as pd
import datetime


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
    table_df.dropna(subset=['Date', 'Week', 'Start', 'End'], inplace=True)
    table_df.reset_index(drop=True, inplace=True)
    
    # Create a tuple of (number of hours, start time, end time, rate)
    timesheet_data = []
    location_list = ["Masters", "NP", "Chiswick", "Sh. Bush", "Acton", "Water", "Horsenden ", "St Helen's", "Disability", "Squad", "Admin", "Safeguarding "]
    for _, row in table_df.iterrows():
        location_hours = [row[loc] if pd.notna(row[loc]) else 0 for loc in location_list]
        timesheet_data.append((
            str(float(sum(location_hours))),
            row["Date"].date().strftime('%d-%m-%Y'),
            row['Rate of pay']
        ))
        
    name = df.iloc[3, 3]
    
    return name, timesheet_data


def check_timesheet(df, sign_in_data: dict[str, tuple[str, datetime.datetime, str]]):
    """
    Check a single timesheet against the sign in data and print any discrepancies found.
    """
    # Read the timesheet
    name, data = read_timesheet(df)

    # Check if the name exists in the sign in data
    if name in sign_in_data:
        
        # Make sets for comparison to not modify the original data
        data_set = set(data)
        sign_in_set = set(sign_in_data[name])

        # For each entry in the timesheet data, match and remove from the sign in data
        print(f"\nChecking timesheet for {name}...")
        for entry in data:
            if entry in sign_in_set:
                sign_in_set.remove(entry)
                data_set.remove(entry)
            else: 
                print(f"Discrepancy found for {name} on {entry[1]}")
        
        remaining_entries = sorted(data_set, key=lambda x: x[1])
        remaining_sign_in = sorted(sign_in_set, key=lambda x: x[1])
        
        if len(remaining_sign_in) != 0:
            print("\n")
            for i in range(len(remaining_entries)):
                print(f"{name} claims to have worked for {remaining_entries[i][0]} hours on {remaining_entries[i][1]}, with a rate of {remaining_entries[i][2]}")
            print("\n")
            for i in range(len(remaining_sign_in)): 
                print(f"The sign in sheet shows {name} worked for {remaining_sign_in[i][0]} hours on {remaining_sign_in[i][1]}, with a rate of {remaining_sign_in[i][2]}")
        else:
            print(f"All entries for {name} match between timesheet and sign in sheet.")
        print("\n")
        print("Press Enter to continue...")
                

    else:
        print(f"No sign in data found for {name}")

