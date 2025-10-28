import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import os
from datetime import datetime
import sys

# Import our existing classes
from main import Employee, Payslip, load_employees_from_csv, generate_payslips_from_csv

class PayslipGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Payslip Generator - CSV & Excel Compatible")
        self.root.geometry("800x600")
        self.root.resizable(True, True)

        # Set theme
        style = ttk.Style()
        style.theme_use('clam')

        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="Payslip Generator",
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="CSV/Excel File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))

        self.browse_button = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2)

        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(10, 20))

        self.load_button = ttk.Button(button_frame, text="Load Employee Data",
                                    command=self.load_employee_data, width=20)
        self.load_button.grid(row=0, column=0, padx=(0, 10))

        self.generate_button = ttk.Button(button_frame, text="Generate Payslips",
                                        command=self.generate_payslips, width=20)
        self.generate_button.grid(row=0, column=1, padx=(0, 10))

        self.template_button = ttk.Button(button_frame, text="Create Template",
                                        command=self.create_template, width=20)
        self.template_button.grid(row=0, column=2)

        # Employee list section
        list_frame = ttk.LabelFrame(main_frame, text="Employee Data", padding="10")
        list_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        # Create treeview for employee data
        columns = ('Name', 'ID', 'Rate', 'Hours', 'Department', 'Position')
        self.employee_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=10)

        # Define headings
        for col in columns:
            self.employee_tree.heading(col, text=col)
            self.employee_tree.column(col, width=100)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.employee_tree.yview)
        self.employee_tree.configure(yscrollcommand=scrollbar.set)

        self.employee_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

        # Initialize variables
        self.employees = []
        self.current_file = None

        # Bind double-click event
        self.employee_tree.bind('<Double-1>', self.show_employee_payslip)

    def browse_file(self):
        """Browse for CSV or Excel file"""
        filetypes = [
            ('CSV files', '*.csv'),
            ('Excel files', '*.xlsx'),
            ('All files', '*.*')
        ]

        filename = filedialog.askopenfilename(
            title="Select Employee Data File",
            filetypes=filetypes
        )

        if filename:
            self.file_path_var.set(filename)
            self.current_file = filename
            self.status_var.set(f"Selected: {os.path.basename(filename)}")

    def load_employee_data(self):
        """Load employee data from selected file"""
        if not self.current_file:
            messagebox.showerror("Error", "Please select a file first")
            return

        try:
            # Clear existing data
            for item in self.employee_tree.get_children():
                self.employee_tree.delete(item)

            # Load data based on file type
            if self.current_file.endswith('.csv'):
                self.employees = self.load_csv_data(self.current_file)
            elif self.current_file.endswith('.xlsx'):
                self.employees = self.load_excel_data(self.current_file)
            else:
                messagebox.showerror("Error", "Unsupported file format")
                return

            # Populate treeview
            for employee in self.employees:
                self.employee_tree.insert('', tk.END, values=(
                    employee.name,
                    employee.employee_id,
                    f"${employee.hourly_rate:.2f}",
                    f"{employee.hours_worked:.1f}",
                    employee.department,
                    employee.position
                ))

            self.status_var.set(f"Loaded {len(self.employees)} employees")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")

    def load_csv_data(self, filename):
        """Load data from CSV file"""
        employees = []
        try:
            with open(filename, 'r') as file:
                csv_reader = csv.DictReader(file)
                for row in csv_reader:
                    employee = Employee(
                        name=row['Employee Name'],
                        employee_id=row['Employee ID'],
                        hourly_rate=row['Hourly Rate'],
                        hours_worked=row['Hours Worked'],
                        department=row.get('Department', ''),
                        position=row.get('Position', '')
                    )
                    employees.append(employee)
        except Exception as e:
            raise Exception(f"CSV loading error: {str(e)}")

        return employees

    def load_excel_data(self, filename):
        """Load data from Excel file (placeholder for future implementation)"""
        # For now, we'll show a message that Excel support requires additional libraries
        messagebox.showinfo("Excel Support",
                          "Excel file support requires the 'openpyxl' library.\n"
                          "Please install it using: pip install openpyxl\n"
                          "For now, please use CSV format.")
        return []

    def generate_payslips(self):
        """Generate payslips for all loaded employees"""
        if not self.employees:
            messagebox.showerror("Error", "No employee data loaded")
            return

        # Ask for output directory
        output_dir = filedialog.askdirectory(title="Select Output Directory")
        if not output_dir:
            return

        try:
            # Generate payslips
            generated_count = 0
            for employee in self.employees:
                payslip = Payslip(employee)
                payslip.calculate_deductions()
                payslip.calculate_net_pay()

                # Save payslip
                filename = f"payslip_{employee.employee_id}_{datetime.now().strftime('%Y%m%d')}.txt"
                filepath = os.path.join(output_dir, filename)

                with open(filepath, 'w') as f:
                    f.write(payslip.generate_payslip())

                generated_count += 1

            messagebox.showinfo("Success",
                              f"Generated {generated_count} payslips in:\n{output_dir}")
            self.status_var.set(f"Generated {generated_count} payslips")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate payslips: {str(e)}")

    def create_template(self):
        """Create a template file for user to fill"""
        template_content = """Employee Name,Employee ID,Hourly Rate,Hours Worked,Department,Position,Tax Exempt,Insurance Plan,Notes
John Doe,EMP001,25.00,40.0,Engineering,Software Developer,No,Standard,
Jane Smith,EMP002,22.50,38.5,Marketing,Marketing Manager,No,Premium,
Bob Johnson,EMP003,20.00,45.0,Sales,Sales Representative,No,Basic,Overtime eligible
Alice Brown,EMP004,28.00,42.0,HR,HR Specialist,Yes,Standard,
Charlie Wilson,EMP005,18.50,37.0,Finance,Accountant,No,Basic,"""

        # Ask where to save template
        filename = filedialog.asksaveasfilename(
            title="Save Template As",
            defaultextension=".csv",
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')],
            initialfile="employee_template.csv"
        )

        if filename:
            try:
                with open(filename, 'w') as f:
                    f.write(template_content)

                messagebox.showinfo("Template Created",
                                  f"Template saved as:\n{filename}\n\n"
                                  "You can open this file in Excel or any spreadsheet application.")

                # Open the template in default application
                os.startfile(filename)

            except Exception as e:
                messagebox.showerror("Error", f"Failed to create template: {str(e)}")

    def show_employee_payslip(self, event):
        """Show payslip for double-clicked employee"""
        selection = self.employee_tree.selection()
        if selection:
            item = self.employee_tree.item(selection[0])
            employee_id = item['values'][1]

            # Find employee
            employee = next((emp for emp in self.employees if emp.employee_id == employee_id), None)
            if employee:
                payslip = Payslip(employee)
                payslip.calculate_deductions()
                payslip.calculate_net_pay()

                # Create popup window
                popup = tk.Toplevel(self.root)
                popup.title(f"Payslip - {employee.name}")
                popup.geometry("600x500")

                # Text widget for payslip
                text_widget = tk.Text(popup, wrap=tk.WORD, padx=10, pady=10)
                text_widget.insert(tk.END, payslip.generate_payslip())
                text_widget.config(state=tk.DISABLED)

                scrollbar = ttk.Scrollbar(popup, orient=tk.VERTICAL, command=text_widget.yview)
                text_widget.configure(yscrollcommand=scrollbar.set)

                text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

                # Save button
                save_button = ttk.Button(popup, text="Save Payslip",
                                       command=lambda: self.save_individual_payslip(payslip, popup))
                save_button.pack(pady=10)

    def save_individual_payslip(self, payslip, parent_window):
        """Save individual payslip"""
        filename = filedialog.asksaveasfilename(
            title="Save Payslip As",
            defaultextension=".txt",
            filetypes=[('Text files', '*.txt'), ('All files', '*.*')],
            initialfile=f"payslip_{payslip.employee.employee_id}.txt"
        )

        if filename:
            try:
                with open(filename, 'w') as f:
                    f.write(payslip.generate_payslip())
                messagebox.showinfo("Saved", f"Payslip saved as:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save: {str(e)}")

def main():
    root = tk.Tk()
    app = PayslipGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()