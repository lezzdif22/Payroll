import csv
import os
from datetime import datetime

class Employee:
    def __init__(self, name, employee_id, hourly_rate, hours_worked, department="", position=""):
        self.name = name
        self.employee_id = employee_id
        self.hourly_rate = float(hourly_rate)
        self.hours_worked = float(hours_worked)
        self.department = department
        self.position = position
        self.overtime_hours = max(0, self.hours_worked - 40)  # Overtime after 40 hours
        self.regular_hours = min(self.hours_worked, 40)

class Payslip:
    def __init__(self, employee, company_name="Tech Solutions Inc."):
        self.employee = employee
        self.company_name = company_name
        self.pay_period = datetime.now().strftime("%B %Y")
        self.pay_date = datetime.now().strftime("%B %d, %Y")

        # Pay calculations
        self.regular_pay = self.employee.regular_hours * self.employee.hourly_rate
        self.overtime_pay = self.employee.overtime_hours * (self.employee.hourly_rate * 1.5)
        self.gross_pay = self.regular_pay + self.overtime_pay

        # Deductions
        self.deductions = {}
        self.net_pay = 0

    def calculate_deductions(self):
        """Calculate various deductions"""
        # Federal Tax (20% of gross pay)
        self.deductions['Federal Tax'] = self.gross_pay * 0.20

        # State Tax (5% of gross pay)
        self.deductions['State Tax'] = self.gross_pay * 0.05

        # Social Security (6.2% of gross pay)
        self.deductions['Social Security'] = self.gross_pay * 0.062

        # Medicare (1.45% of gross pay)
        self.deductions['Medicare'] = self.gross_pay * 0.0145

        # Health Insurance (fixed amount)
        self.deductions['Health Insurance'] = 75.00

        # Retirement (401k - 5% of gross pay)
        self.deductions['Retirement (401k)'] = self.gross_pay * 0.05

        return self.deductions

    def calculate_net_pay(self):
        """Calculate net pay after deductions"""
        total_deductions = sum(self.deductions.values())
        self.net_pay = self.gross_pay - total_deductions
        return self.net_pay

    def generate_payslip(self):
        """Generate formatted payslip based on template"""
        payslip_text = f"""
{'='*50}
                PAYSLIP
{'='*50}

{self.company_name}
PAY PERIOD: {self.pay_period}
PAY DATE: {self.pay_date}

EMPLOYEE INFORMATION:
Name: {self.employee.name}
Employee ID: {self.employee.employee_id}
Department: {self.employee.department}
Position: {self.employee.position}

PAY DETAILS:
{'-'*50}
Regular Hours: {self.employee.regular_hours} hours
Hourly Rate: ${self.employee.hourly_rate:.2f}
"""

        if self.employee.overtime_hours > 0:
            payslip_text += f"""Overtime Hours: {self.employee.overtime_hours} hours
Overtime Rate: ${self.employee.hourly_rate * 1.5:.2f}

"""
        else:
            payslip_text += "\n"

        payslip_text += f"""GROSS PAY CALCULATION:
Regular Pay: ${self.regular_pay:.2f}
Overtime Pay: ${self.overtime_pay:.2f}
Gross Pay: ${self.gross_pay:.2f}

DEDUCTIONS:
{'-'*50}
"""

        for deduction_type, amount in self.deductions.items():
            if 'Tax' in deduction_type or 'Social' in deduction_type or 'Medicare' in deduction_type:
                rate = ""
                if 'Federal' in deduction_type:
                    rate = " (20%)"
                elif 'State' in deduction_type:
                    rate = " (5%)"
                elif 'Social' in deduction_type:
                    rate = " (6.2%)"
                elif 'Medicare' in deduction_type:
                    rate = " (1.45%)"
                payslip_text += f"{deduction_type}{rate}: -${amount:.2f}\n"
            else:
                payslip_text += f"{deduction_type}: -${amount:.2f}\n"

        payslip_text += f"""
Total Deductions: -${sum(self.deductions.values()):.2f}

NET PAY: ${self.net_pay:.2f}
{'='*50}

Payment Method: Direct Deposit
Account: ****-****-****-1234

For questions about this payslip, contact HR at hr@techsolutions.com

{'='*50}
"""
        return payslip_text

    def save_to_file(self, filename=None):
        """Save payslip to a text file"""
        if filename is None:
            filename = f"payslip_{self.employee.employee_id}_{datetime.now().strftime('%Y%m%d')}.txt"

        with open(filename, 'w') as f:
            f.write(self.generate_payslip())

        return filename

def load_employees_from_csv(filename):
    """Load employee data from CSV file"""
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
    except FileNotFoundError:
        print(f"Error: File '{filename}' not found.")
        return []
    except KeyError as e:
        print(f"Error: Missing required column in CSV: {e}")
        return []
    except Exception as e:
        print(f"Error reading CSV file: {e}")
        return []

    return employees

def generate_payslips_from_csv(csv_filename, output_dir="payslips"):
    """Generate payslips for all employees in CSV file"""
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Load employees from CSV
    employees = load_employees_from_csv(csv_filename)

    if not employees:
        print("No employees loaded. Please check your CSV file.")
        return

    print(f"Loaded {len(employees)} employees from {csv_filename}")
    print(f"Generating payslips in '{output_dir}' directory...")

    # Generate payslip for each employee
    for employee in employees:
        payslip = Payslip(employee)
        payslip.calculate_deductions()
        payslip.calculate_net_pay()

        # Save payslip to file
        filename = payslip.save_to_file()
        full_path = os.path.join(output_dir, filename)
        with open(full_path, 'w') as f:
            f.write(payslip.generate_payslip())

        print(f"Generated payslip for {employee.name} ({employee.employee_id})")

    print(f"\nAll payslips generated successfully in '{output_dir}' directory!")

def main():
    """Main function to run the payslip generator"""
    print("Enhanced Payslip Generator")
    print("=" * 40)
    print("Choose an option:")
    print("1. Generate payslips from CSV file")
    print("2. Manual entry (single employee)")
    print("3. View CSV template")
    print("4. Exit")

    choice = input("\nEnter your choice (1-4): ").strip()

    if choice == '1':
        csv_file = input("Enter CSV filename (default: employees.csv): ").strip()
        if not csv_file:
            csv_file = "employees.csv"

        output_dir = input("Enter output directory (default: payslips): ").strip()
        if not output_dir:
            output_dir = "payslips"

        generate_payslips_from_csv(csv_file, output_dir)

    elif choice == '2':
        # Manual entry mode
        print("\nManual Entry Mode")
        print("-" * 20)

        name = input("Enter employee name: ")
        employee_id = input("Enter employee ID: ")
        department = input("Enter department (optional): ")
        position = input("Enter position (optional): ")

        try:
            hourly_rate = float(input("Enter hourly rate: $"))
            hours_worked = float(input("Enter hours worked: "))
        except ValueError:
            print("Invalid input. Please enter numeric values for rate and hours.")
            return

        employee = Employee(name, employee_id, hourly_rate, hours_worked, department, position)
        payslip = Payslip(employee)
        payslip.calculate_deductions()
        payslip.calculate_net_pay()

        print("\n" + payslip.generate_payslip())

        save_option = input("Save this payslip to file? (y/n): ").lower().strip()
        if save_option == 'y':
            filename = payslip.save_to_file()
            print(f"Payslip saved as: {filename}")

    elif choice == '3':
        print("\nCSV Template Format:")
        print("The CSV file should have these columns:")
        print("Employee Name,Employee ID,Hourly Rate,Hours Worked,Department,Position")
        print("\nExample:")
        print("John Doe,EMP001,25.00,40.0,Engineering,Software Developer")
        print("Jane Smith,EMP002,22.50,38.5,Marketing,Marketing Manager")
        print("\nNote: Department and Position columns are optional.")

    elif choice == '4':
        print("Goodbye!")
        return

    else:
        print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()