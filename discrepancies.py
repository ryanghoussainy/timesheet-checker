from printing import print_colour, RED, YELLOW, GREEN

EMPTY_TIMESHEET = 0
INVALID_NAME = 1
TIMESHEET_EXTRA_ENTRY = 2
SIGN_IN_EXTRA_ENTRY = 3


def print_discrepancies(discrepancies):
    if discrepancies:

        print("Mismatches found:")
        for d in discrepancies:
            # Extract information
            discrepancy_type, details = d

            print("- ", end='')
            if discrepancy_type == EMPTY_TIMESHEET:
                sheet_name = details["sheet name"]

                print_colour(RED, f"Empty timesheet: {sheet_name}")
            
            elif discrepancy_type == INVALID_NAME:
                name = details["name"]
                sign_in_names = details["sign in names"]

                print_colour(RED, f"Invalid name in timesheet: {name}")
                print_colour(YELLOW, "Names in sign in sheet are:", end=' ')
                print_colour(YELLOW, ", ".join(sign_in_names))

            elif discrepancy_type == TIMESHEET_EXTRA_ENTRY:
                name = details["name"]
                entry = details["entry"]

                # Unpack entry
                hours, date_time, rate = entry

                print_colour(RED, f"Extra entry in timesheet for {name}: {hours} hours on {date_time} at {rate}/hour")
            
            elif discrepancy_type == SIGN_IN_EXTRA_ENTRY:
                name = details["name"]
                entry = details["entry"]

                # Unpack entry
                hours, date_time, rate = entry

                print_colour(RED, f"Extra entry in sign in sheet for {name}: {hours} hours on {date_time} at {rate}/hour")
    else:
        print_colour(GREEN, "No mismatches found.")
