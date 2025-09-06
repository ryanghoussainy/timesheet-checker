from discrepancies.discrepancy_types.discrepancy import Discrepancy
from printing import colour_text, RED
from entry import Entry

class TimesheetExtraEntry(Discrepancy):
    def __init__(self, name: str, entry: Entry):
        self.name = name
        self.entry = entry

    def __str__(self):
        return f"- {colour_text(RED, f'Extra entry in timesheet for {self.name}: {self.entry.hours} hours on {self.entry.date} at {self.entry.rate}/hour')}"
