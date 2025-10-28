# Dynamic Payroll PDF Generator

A comprehensive payroll processing system that dynamically detects pay periods from CSV files and generates professional PDF payslips using ReportLab. Designed for university payroll systems with variable period structures (2-5 periods).

## 🎯 Features

- ✅ **Modern GUI Interface**: User-friendly interface with progress tracking and real-time updates
- ✅ **Dynamic Period Detection**: Automatically detects and processes 2-5 pay periods from CSV headers
- ✅ **Professional PDF Output**: Generates publication-ready PDF payslips with tables and formatting
- ✅ **University Payroll Support**: Specifically designed for PAMANTASAN NG LUNGSOD NG VALENZUELA payroll structure
- ✅ **Multi-Encoding Support**: Handles various CSV encodings (UTF-8, Latin-1, CP1252, ISO-8859-1)
- ✅ **Organized File Structure**: Automatic folder creation for inputs, templates, outputs, and logs
- ✅ **Batch Processing**: Processes hundreds of employees simultaneously
- ✅ **Tax Calculation**: Handles multiple tax types (W/TAX, W/HOLDING, P-TAX) with percentage rates
- ✅ **Error Handling**: Robust error handling with detailed logging and skip mechanisms
- ✅ **Email Address Persistence**: Remembers validated recipient emails and reuses them automatically on future runs

## 📁 Project Structure

```
├── dynamic_payroll_pdf_generator.py  # Main processing script
├── dynamic_payroll_gui.py            # Modern GUI interface
├── gui_payslip.py                     # GUI interface (legacy)
├── main.py                           # Basic payslip generator (legacy)
├── Payroll.csv                       # Sample university payroll data
├── employees.csv                     # Sample employee data
├── requirements.txt                  # Python dependencies
├── run_gui.bat                       # GUI launcher
├── README.md                         # This documentation
├── EXCEL_TEMPLATE_README.md          # Excel template instructions
├── Template Payslip.xlsx             # Excel template file
├── input_csv/                        # Input CSV storage
├── templates/                        # Template files
├── output_pdfs/                      # Generated PDF payslips
├── logs/                            # Processing logs
└── payslips/                        # Legacy text payslips
```

## 🚀 Quick Start

### Prerequisites
```bash
pip install -r requirements.txt
```

### Basic Usage
```bash
python dynamic_payroll_pdf_generator.py
```

The system will:
1. Scan for CSV files in the current directory
2. Display available files for selection
3. Automatically detect pay periods from CSV headers
4. Process all employees and generate individual PDF payslips
5. Save PDFs to `output_pdfs/` folder

### GUI Interface (Recommended)
```bash
python dynamic_payroll_gui.py
# Or simply double-click run_gui.bat
```

The modern GUI provides:
- **Visual File Selection**: Browse and select CSV files with preview
- **Real-time Progress**: Progress bar and status updates during processing
- **Detailed Results**: Live log of period detection and PDF generation
- **Quick Access**: One-click buttons for common files (Payroll.csv)
- **Output Viewer**: Direct access to generated PDF folder

**GUI Features:**
- Progress tracking with percentage completion
- Detailed processing logs with success/failure indicators
- Error handling with user-friendly messages
- Automatic folder organization
- Professional interface with modern styling
- **Employee Data Preview**: Load and preview employee data with period breakdowns before processing
- **Tabbed Interface**: Separate tabs for processing logs and employee data
- **Interactive Data View**: Click employees to see detailed period-by-period breakdowns
- **Data Sorting**: Sort employees by sequence, name, rate, gross pay, or net pay

**GUI Workflow:**
1. **Launch** → GUI opens with file scanning
2. **Select File** → Browse or quick-select Payroll.csv
3. **Preview Data** → Click "Load & Preview Data" to see employee information and period breakdowns
4. **Review Data** → Use the Employee Data tab to sort and examine employee details
5. **Process** → Click "Process Payroll & Generate PDFs" for full processing
6. **Monitor** → Watch real-time progress and detailed logs
7. **View Results** → Click "View Output PDFs" to see generated files

### Command Line Options
- Select from numbered list of CSV files
- Or enter a specific filename
- Press Enter to use default file

## 📊 CSV File Format

The system expects university payroll CSV files with the following structure:

```csv
PAMANTASAN NG LUNGSOD NG VALENZUELA,,,,,,,,,,,,,,,,
PART TIME FACULTY WAGES - (COLLEGE),,,,,,,,,,,,,,,,
,,,,,,,,,,,,,,,,
"January 13-31, 2025",,,,,,,,,,,,,,,,
,,Forwarded: 2/19,,,,,,,,,,,,,,
,,Released: 2/25 ,,,,,,,,,,,,,~ ATM ~,,,,,
,,,,,,,,,,,,,,,,
Seq.,Account No.,NAME,RATE,,,,AMOUNT,ADJUSTMENT,,NET AMOUNT ,W/ TAX,W/HOLDING,P-TAX,PERCENT.,TOTAL TAX, NET AMOUNT,,,,,,,,
,,,per hour,May 1-31,Feb 1-28, Jan 13-31 ,EARNED,No. of Hours,Amount,EARNED,RATE,TAX ,RATE,TAX,DEDUCTIONS,RECEIVED,,,,,,,,
1,,"Abante, Julie H.", 300.00 , 37.42 , 47.14 , 32.33 ," 35,067.00 ",,," 35,067.00 ",10%," 3,506.70 ",3%," 1,052.01 "," 4,558.71 "," 30,508.29 ",,,,,,,,,
```

### Key Columns Detected:
- **Sequence Number**: Employee sequence in payroll
- **Account Number**: Employee account identifier
- **Name**: Employee full name
- **Hourly Rate**: Base pay rate per hour
- **Dynamic Periods**: Variable number of pay periods (detected automatically)
- **Tax Rates**: W/TAX, W/HOLDING, P-TAX percentages

## 🔍 Dynamic Period Detection

The system automatically detects pay period columns by scanning merged header rows with a tolerant pattern:

```python
PERIOD_HEADER_REGEX = r'([A-Za-z]{3,}\.?)\s*\d{0,2}\s*[-–—]\s*\d{1,2}'
```

Key capabilities:

- Supports month names with or without trailing periods (e.g., `Sept. 1-15`)
- Accepts hyphen, en dash, or em dash separators (`Sept 1–15`, `Sept 1—15`)
- Handles optional leading day numbers (`Oct -15` is normalised to `Oct -15`)
- Preserves original header text so the GUI mirrors the CSV exactly
- Works with two-line headers by merging the "per hour" row with the row above it

**Examples that now detect correctly:**
- `Sept. 1–15`
- `Sept 16-30`
- `Oct -15`
- `January 13-31`
- `Feb 1 – 28 (Hours)`
- Any month abbreviation or full name followed by a date range

## � Email Delivery Enhancements

Sending payslips by email is now simpler thanks to a lightweight persistent store:

- **Automatic memory**: Every time you supply or confirm an employee email (via CSV, the missing-email dialog, or during sending) it is saved to a local SQLite database located at:
	- Windows: `%LOCALAPPDATA%\dynamic_payroll\email_addresses.sqlite3`
	- Linux/macOS: `~/.local/share/dynamic_payroll/email_addresses.sqlite3`
- **Legacy compatibility**: Existing `emails.csv` files are imported on startup, and the store exports back to `emails.csv` after updates so you can still edit addresses in a spreadsheet if you prefer.
- **Auto-fill**: When loading employee data, any known address is injected automatically so you do not re-enter it on each run.
- **Multiple keys**: Lookups use sequence number first, then account number, then the normalised employee name giving the best chance of a match.

**Tip:** You can pre-populate addresses by dropping an `emails.csv` (with `seq,account_no,name,email` columns) beside the executable; they will be imported into the store on the next launch.

## �📄 PDF Output Features

Generated payslips include:

### Header Section
- University name: "PAMANTASAN NG LUNGSOD NG VALENZUELA"
- Document title: "PART-TIME FACULTY PAYSLIP"
- Pay period and generation date

### Employee Information
- Sequence number
- Account number (if available)
- Employee name
- Hourly rate

### Pay Details Table
- Period-by-period breakdown (hours and amounts)
- Rate information
- Total calculations

### Payroll Summary
- Gross pay total
- Tax deductions (W/TAX, W/HOLDING, P-TAX)
- Total deductions
- Net pay amount

### Payment Information
- Payment method (ATM Transfer)
- Bank details placeholder
- Account information placeholder

## 🛠️ Technical Details

### Dependencies
- `openpyxl >= 3.0.0` - Excel file handling
- `reportlab >= 4.0.0` - PDF generation
- `csv` - Built-in CSV processing
- `os`, `re`, `datetime` - Standard library modules

### Encoding Support
The system tries multiple encodings in order:
1. UTF-8 (default)
2. Latin-1 (ISO-8859-1)
3. CP1252 (Windows-1252)
4. ISO-8859-1

### Error Handling
- Skips malformed rows with detailed error messages
- Continues processing despite individual employee errors
- Logs encoding detection and processing status
- Graceful handling of missing or invalid data

## 📈 Processing Statistics

**Sample Run Results:**
- **Pay Periods Detected**: 3 (May 1-31, Feb 1-28, Jan 13-31)
- **Employees Processed**: 285
- **PDFs Generated**: 285
- **Processing Time**: ~30 seconds
- **Output Location**: `output_pdfs/` folder

## 🎯 Use Cases

### University Payroll Processing
- Part-time faculty wages
- Multiple pay periods per month
- Complex tax structures
- Batch PDF generation for distribution

### Variable Period Payroll
- Systems with 2-5 pay periods
- Dynamic period detection
- Flexible CSV structures
- Professional PDF output

## 🔧 Customization

### Tax Rate Modification
```python
# In dynamic_payroll_pdf_generator.py
employee['tax_rate'] = float(tax_str.rstrip('%')) / 100
```

### Company Information
```python
# Modify these variables in the PDF generation method
university_name = "YOUR UNIVERSITY NAME"
document_title = "PAYSLIP"
```

### PDF Styling
```python
# Adjust colors, fonts, and layout in generate_pdf_payslip method
title_style.fontSize = 18  # Change title size
table_style = TableStyle([...])  # Modify table appearance
```

## 📋 Requirements

- **Python**: 3.6+
- **Memory**: 512MB minimum (depends on employee count)
- **Disk Space**: ~50MB for dependencies + output PDFs
- **OS**: Windows, macOS, Linux

## 🚨 Troubleshooting

### Common Issues

**"Could not find header row with 'per hour'"**
- Ensure CSV contains the expected header structure
- Check file encoding (system supports multiple encodings)

**"UnicodeDecodeError"**
- File encoding not supported
- System automatically tries alternative encodings

**"PDF generation failed"**
- Check ReportLab installation
- Verify employee data integrity

### File Organization
The system creates these folders automatically:
- `input_csv/` - Store input CSV files here
- `templates/` - Excel templates and formats
- `output_pdfs/` - Generated PDF payslips
- `logs/` - Processing logs and error reports

## 🔄 Future Enhancements

- Excel template integration (`Template Payslip.xlsx`)
- Email distribution system
- Database storage integration
- Advanced reporting and analytics
- Multi-language support
- Custom PDF templates
- Web-based interface

## 📞 Support

For issues or questions:
1. Check the generated log files in `logs/` folder
2. Verify CSV format matches the expected structure
3. Ensure all dependencies are installed
4. Review error messages for specific guidance

## 📝 License

This project is designed for educational and organizational use. Ensure compliance with local data protection regulations when processing payroll information.

## Building a Windows EXE

You can produce a Windows executable for this project using PyInstaller. Two options are provided:

1) GitHub Actions (recommended): a workflow in `.github/workflows/build-windows.yml` builds on `windows-latest` and uploads a `dynamic_payroll_exe` artifact on push to `main`. Download the artifact from the Actions run.

2) Local build on Windows: run `build_windows.bat` on a Windows machine with Python 3.11 installed. This installs dependencies and runs PyInstaller to create a single-file exe in `dist/`.

Notes:
- PyInstaller must run on the target OS to create a native executable. Building a Windows .exe from Linux is not supported directly; use the GitHub Actions workflow, a Windows machine, or a Windows VM/WSL2 with proper GUI support.
- If your app needs extra data files (templates, images, Excel templates), ensure they are included with PyInstaller's `--add-data` option. The workflow and local scripts include example `--add-data` flags for `templates/` and `PLV LOGO.png`.

## Local Linux build

Run `./build_linux.sh` to build a single-file Linux binary (requires PyInstaller). This will not produce a Windows exe.
# Payroll
# Payroll
# Payroll
# Payroll
# Payroll
# Payroll
