import pandas as pd
import datetime
from level_to_rate import level_to_rate

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
                        col.date().strftime('%d-%m-%Y'),
                        level_to_rate[row["Level"]]
                    ))

    # print(sign_in_sheet_data)
    return sign_in_sheet_data