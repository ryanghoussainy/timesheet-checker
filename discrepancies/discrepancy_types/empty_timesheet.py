from discrepancies.discrepancy_types.discrepancy import Discrepancy
from printing import colour_text, RED

class EmptyTimesheet(Discrepancy):
    def __init__(self, sheet_name: str):
        self.sheet_name = sheet_name
    
    def __str__(self):
        return f"- {colour_text(RED, f'Empty timesheet: {self.sheet_name}')}"