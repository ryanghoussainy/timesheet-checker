from discrepancies.discrepancy_types.discrepancy import Discrepancy
from printing import colour_text, RED, YELLOW

class InvalidName(Discrepancy):
    def __init__(self, name: str, sign_in_names: list[str]):
        self.name = name
        self.sign_in_names = sign_in_names

    def __str__(self):
        return f"- {colour_text(RED, f'Invalid name in timesheet: {self.name}')}\n" \
               f"{colour_text(YELLOW, 'Names in sign in sheet are:')}\n" \
               f"{colour_text(YELLOW, ', '.join(self.sign_in_names))}"
