import os
import pandas as pd
from datetime import datetime
    
def amindefy_timesheets(timesheet_folder: str, output_file: str):
    """
    Takes in a folder of timesheets and combines them into 1 excel file, each timesheet as one sheet.
    """
    
    with pd.ExcelWriter(output_file) as writer:
        for filename in os.listdir(timesheet_folder):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(timesheet_folder, filename)
                df = pd.read_excel(file_path)
                sheet_name = os.path.splitext(filename)[0]
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"All timesheets have been combined into {output_file}.")