import sys


def main():
    if len(sys.argv) != 1:
        # Invalid format
        print("Usage: python run.py")
        sys.exit(1)
    
    from check_timesheets import read_timesheet, read_sign_in_sheet, check_timesheets
    read_timesheet("examples/WilliamTest.xlsx")
    read_timesheet("examples/Ryan.xlsx")
    read_sign_in_sheet("July", "examples/TestSheet.xlsx")


if __name__ == "__main__":
    main()
