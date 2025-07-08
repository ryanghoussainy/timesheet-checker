import sys


def main():
    if len(sys.argv) != 1:
        # Invalid format
        print("Usage: python run.py")
        sys.exit(1)
    
    from check_timesheets import read_timesheet
    read_timesheet("examples/William.xlsx")
    

if __name__ == "__main__":
    main()
