import pandas as pd
from entry import Entry
from collections import defaultdict

def read_sign_in_sheet(month: str, file_path: str, rates: dict[str, float]) -> dict[str, set[Entry]]:
    """
    Read a sign in sheet excel file and return a dictionnary from name to set of entries
    """
    sign_df = pd.read_excel(file_path, month, header=0)

    sign_in_sheet_data = defaultdict(set)

    for _, row in sign_df.iterrows():
        # Skip rows below the table and LHC rows
        if pd.isna(row['Level']) or row['Level'] == "LHC":
            continue
    
        name = row['Name']

        for col in sign_df.columns[3:]:
            # Skip empty cells
            if pd.isna(row[col]):
                continue
            
            entry = Entry(
                date=col.date(),
                hours=float(row[col]),
                rate=rates[row['Level']],
            )
            sign_in_sheet_data[name].add(entry)

    return sign_in_sheet_data