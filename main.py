import tkinter as tk
import sys
import os


def check_packages():
    try:
        import openpyxl
        import requests
        return True
    except ImportError as error:
        print(f"Error: {error}")
        return False


def check_files():
    required_files = ['Interface.py', 'excelbuilder.py', 'skimmer.py']
    missing_files = []

    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)

    if missing_files:
        print("Missing required files:", ", ".join(missing_files))
        return False
    return True


def main():
    if not check_packages():
        sys.exit(1)

    if not check_files():
        sys.exit(1)

    try:
        from Interface import DataScraperGUI

        root = tk.Tk()
        app = DataScraperGUI(root)
        root.mainloop()

    except Exception as e:
        print(f"An error occurred while starting the application: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
