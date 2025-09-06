from printing import print_colour, RED, YELLOW, GREEN

EMPTY_TIMESHEET = 0
INVALID_NAME = 1
TIMESHEET_EXTRA_ENTRY = 2
SIGN_IN_EXTRA_ENTRY = 3


def print_discrepancies(discrepancies):
    if discrepancies:

        print("Mismatches found:")
        for d in discrepancies:
            print(d)
    else:
        print_colour(GREEN, "No mismatches found.")
