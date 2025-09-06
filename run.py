import sys


def main():
    if len(sys.argv) != 1:
        # Invalid format
        print("Usage: python run.py")
        sys.exit(1)
    
    from gui_app import main as gui_main
    gui_main()

if __name__ == "__main__":
    main()
