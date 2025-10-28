import importlib, sys, time
import tkinter as tk

sys.path.insert(0, '/home/infelious/Raphael/Python Project')
from dynamic_payroll_gui import DynamicPayrollGUIGenerator


def main():
    try:
        root = tk.Tk()
    except Exception as e:
        print('Could not open Tk root:', e)
        return 2

    app = DynamicPayrollGUIGenerator(root)
    # Try to show the GUI window (this will block until closed)
    try:
        app.file_path_var.set('Payroll.csv')
        app.load_employee_preview()
        root.deiconify()
        print('GUI shown. Please interact manually: toggle "Show all columns" and drag the horizontal scrollbar, or press the ◀ ▶ buttons to page.')
        print('Close the window when done to continue.')
        root.mainloop()
        return 0
    except Exception as e:
        print('Runtime error during interactive session:', e)
        return 3


if __name__ == '__main__':
    sys.exit(main())
