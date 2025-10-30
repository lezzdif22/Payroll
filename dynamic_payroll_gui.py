import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os, csv, re, threading, sys

# Import the processor (assumed to exist in same workspace)
from dynamic_payroll_pdf_generator import DynamicPayrollProcessor


class DynamicPayrollGUIGenerator:
    def _email_payslips(self, real_send: bool):
        import threading
        t = threading.Thread(target=self._do_email_payslips, args=(real_send,), daemon=True)
        t.start()

    def _do_email_payslips(self, real_send: bool):
        try:
            self.log_result("Preparing emails..." if real_send else "Dry-run: listing emails...")
        except Exception:
            pass

        output_dir = getattr(self, "output_dir", "output_pdfs")
        subject_tpl = "Payslip for {name}"
        body_tpl    = "Please find your payslip attached."

        results = self.processor.send_all_payslips(
            output_dir=output_dir,
            subject_tpl=subject_tpl,
            body_tpl=body_tpl,
            throttle_seconds=1.0,
            dry_run=not real_send,
            progress_cb=(lambda m: self.root.after(0, lambda: self.log_result(m)))
        )

        sent = sum(1 for r in results if r["status"] == ("sent" if real_send else "dry_run"))
        skipped_no_email = sum(1 for r in results if r["status"].startswith("skipped:no_email"))
        skipped_no_pdf = sum(1 for r in results if r["status"].startswith("skipped:no_pdf"))
        errs = [r for r in results if r["status"] == "error"]

        try:
            # UI-safe logging
            summary = f"Done. {'Sent' if real_send else 'Listed'}: {sent} | no email: {skipped_no_email} | no pdf: {skipped_no_pdf} | errors: {len(errs)}"
            self.root.after(0, lambda s=summary: self.log_result(s))
            import os
            for r in results:
                status = r.get('status', '')
                name = r.get('name', '')
                path = r.get('path', '') or r.get('pdf', '')
                err = r.get('error')
                line = f"{status}: {name} :: {os.path.basename(path)}"
                if err:
                    line += f" Err:{err}"
                # schedule each line on the UI thread
                self.root.after(0, (lambda l=line: self.log_result(l)))
        except Exception:
            pass

    def __init__(self, root):
        self.root = root
        self.root.title("Dynamic Payroll PDF Generator - University System")
        self.root.geometry("1100x700")
        self.root.resizable(True, True)

        # Application state
        self.processor = DynamicPayrollProcessor()
        self.current_file = None
        self.employee_data = []
        self.sort_column = None
        self.sort_reverse = False

        # Period column window state for sideways paging (displaycolumns slicing)
        self.period_view_offset = 0
        self.period_view_window = 6  # number of period columns visible at once (adjustable)

        # Styling
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass

        # Main layout
        main_frame = ttk.Frame(root, padding=10)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # File selection
        file_frame = ttk.LabelFrame(main_frame, text="Payroll Data Selection", padding=8)
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="CSV File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 6))
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var)
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))

        self.browse_button = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=(6, 0))

        # Quick select
        quick_frame = ttk.Frame(file_frame)
        quick_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(8, 0))
        # Replace 'Scan Directory' with a direct 'Browse CSV...' button so user can pick any CSV file
        ttk.Button(quick_frame, text="Browse CSV...", command=self.browse_file).pack(side=tk.LEFT, padx=(8, 0))

        # Actions
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(8, 8))
        self.process_button = ttk.Button(action_frame, text="Process Payroll & Generate PDFs", command=self.start_processing)
        self.process_button.pack(side=tk.LEFT)
        ttk.Button(action_frame, text="Enter Missing Emails…", command=self._prompt_missing_emails).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(action_frame, text="Apply Stored Emails", command=self._apply_stored_emails).pack(side=tk.LEFT, padx=(8, 0))
        # 'Send Test PDF' removed for packaged exe
        ttk.Button(action_frame, text="View Output PDFs", command=self.view_output_folder).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(action_frame, text="Email Dry Run",
           command=lambda: self._email_payslips(False)
        ).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Button(action_frame, text="Send Payslips",
           command=lambda: self._email_payslips(True)
        ).pack(side=tk.LEFT, padx=(8, 0))

        # Notebook for log and data preview
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.rowconfigure(2, weight=1)

        # Log tab
        log_frame = ttk.Frame(self.notebook)
        self.results_text = tk.Text(log_frame, height=12, wrap=tk.WORD)
        results_scroll = ttk.Scrollbar(log_frame, orient='vertical', command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=results_scroll.set)
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.notebook.add(log_frame, text="Processing Log")

        # Data tab
        data_frame = ttk.Frame(self.notebook)
        self.notebook.add(data_frame, text="Employee Data")

        controls = ttk.Frame(data_frame)
        controls.grid(row=0, column=0, sticky=(tk.W, tk.E))

        ttk.Button(controls, text="Load & Preview Data", command=self.load_employee_preview).grid(row=0, column=0)
        ttk.Button(controls, text="Clear Data", command=self.clear_employee_data).grid(row=0, column=1, padx=(6, 0))

        # Allow viewing all columns via native horizontal scroll
        self.show_all_columns = tk.BooleanVar(value=False)
        ttk.Checkbutton(controls, text="Show all columns", variable=self.show_all_columns, command=self._on_show_all_columns_toggle).grid(row=0, column=9, padx=(8, 0))

        # Pan buttons (store refs so we can enable/disable)
        self.pan_prev_btn = ttk.Button(controls, text="◀", width=3, command=lambda: self.pan_xview(-0.25))
        self.pan_prev_btn.grid(row=0, column=3, padx=(12, 2))
        self.page_var = tk.StringVar(value='Page 1/1')
        ttk.Label(controls, textvariable=self.page_var).grid(row=0, column=2, padx=(6, 2))
        self.pan_next_btn = ttk.Button(controls, text="▶", width=3, command=lambda: self.pan_xview(0.25))
        self.pan_next_btn.grid(row=0, column=4)

        # Treeview area with horizontal scrollbar
        tree_frame = ttk.Frame(data_frame, padding=(0, 8, 0, 0))
        tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(1, weight=1)

        self.static_columns = [
            'Seq', 'Name', 'Rate (per hour)',
            # period columns will be inserted here dynamically
            'Adjustment (Hours)', 'Adjustment (Amount)',
            'Net Amount Earned', 'W/ Tax Rate', 'W/Holding Tax', 'P-Tax Rate', 'Percent. Tax',
            'Total Tax Deductions', 'Net Amount Received'
        ]

        self.employee_tree = ttk.Treeview(tree_frame, columns=self.static_columns, show='headings')
        self.h_scroll = ttk.Scrollbar(tree_frame, orient='horizontal', command=self._hscroll_command)
        self.v_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.employee_tree.yview)
        self.employee_tree.configure(xscrollcommand=self._on_tree_xscroll, yscrollcommand=self.v_scroll.set)

        # Block header resizing
        try:
            self.employee_tree.bind('<Button-1>', self._block_header_resize, add='+')
            self.employee_tree.bind('<B1-Motion>', self._block_header_resize, add='+')
        except Exception:
            pass

        self.employee_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.h_scroll.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.v_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        for col in self.static_columns:
            self.employee_tree.heading(col, text=col)
            width = 200 if col == 'Name' else 120
            self.employee_tree.column(col, width=width, anchor='center', stretch=False)

        self.period_text = tk.Text(data_frame, height=6, state=tk.DISABLED)
        self.period_text.grid(row=2, column=0, sticky=(tk.W, tk.E))

        self.status_var = tk.StringVar(value='Ready')
        status = ttk.Label(main_frame, textvariable=self.status_var)
        status.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(6, 0))
    
        try:
            self.root.bind('<Left>', lambda e: self.pan_xview(-0.25))
            self.root.bind('<Right>', lambda e: self.pan_xview(0.25))
        except Exception:
            pass
    def _safe_set_displaycolumns(self, all_cols, desired_displaycols):
        all_cols = list(all_cols) if all_cols else []
        if not all_cols:
            self.employee_tree['displaycolumns'] = tuple(self.static_columns)
            return
        avail = set(all_cols) | {"#0", "#all"}
        cleaned = [c for c in desired_displaycols if c in avail]
        if not cleaned:
            cleaned = list(all_cols)  # fallback: show all
        self.employee_tree['displaycolumns'] = tuple(cleaned)

    def _prompt_missing_emails(self):
        """Prompt for emails that are missing; save to emails.csv and update the table."""
        # Collect employees with missing/invalid emails
        def _valid_email(e: str) -> bool:
            e = (e or "").strip()
            return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", e))

        employees = list(getattr(self.processor, "employees", []) or [])
        if not employees:
            employees = list(getattr(self, "employee_data", []) or [])

        if not employees:
            try:
                messagebox.showwarning("Emails", "Load a payroll file first before entering missing emails.")
            except Exception:
                print("No employees loaded yet; cannot enter emails.")
            return

        def _primary_email(emp_dict):
            for key in ("email", "work_email", "personal_email"):
                val = (emp_dict.get(key) or "").strip()
                if val:
                    return val
            return ""

        missing = []
        for emp in employees:
            if not _valid_email(_primary_email(emp)):
                missing.append(emp)

        if missing:
            rows_source = missing
            dialog_title = "Enter Missing Emails"
            banner_text = None
        else:
            rows_source = employees
            dialog_title = "Review Emails"
            banner_text = "All employees currently have an email on file. You can review or update them below."

        # Build the dialog
        win = tk.Toplevel(self.root)
        win.title(dialog_title)
        win.resizable(False, True)
        win.grab_set()

        # Scrollable frame (in case many entries)
        container = ttk.Frame(win)
        canvas = tk.Canvas(container, height=400, width=560)
        scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        form = ttk.Frame(canvas)

        form.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)

        container.pack(fill="both", expand=True, padx=10, pady=10)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # Headers / banner
        if banner_text:
            ttk.Label(form, text=banner_text, foreground="#555555", wraplength=520, justify="left").grid(row=0, column=0, columnspan=3, sticky="w", padx=6, pady=(4, 0))
            start_row = 1
        else:
            start_row = 0

        header_row = start_row
        ttk.Label(form, text="Employee", font=("TkDefaultFont", 10, "bold")).grid(row=header_row, column=0, sticky="w", padx=6, pady=6)
        ttk.Label(form, text="Email",    font=("TkDefaultFont", 10, "bold")).grid(row=header_row, column=1, sticky="w", padx=6, pady=6)

        # Rows
        rows = []
        for idx, emp in enumerate(rows_source, start=1):
            row = header_row + idx
            name = emp.get("name", "") or f"(No name — seq {emp.get('seq','')})"
            ttk.Label(form, text=name).grid(row=row, column=0, sticky="w", padx=6, pady=4)
            var = tk.StringVar(value=_primary_email(emp))
            entry = ttk.Entry(form, textvariable=var, width=44)
            entry.grid(row=row, column=1, sticky="w", padx=6, pady=4)
            rows.append((emp, var))
        

        # Buttons
        btns = ttk.Frame(win); btns.pack(fill="x", padx=10, pady=(0,10))
        def _save_and_apply():
            # Validate & update in-memory
            bad = []
            for emp, var in rows:
                e = (var.get() or "").strip()
                # Allow empty entries (user may only fill some emails). Only treat non-empty invalids as errors.
                if not e:
                    # leave as missing
                    continue
                if not _valid_email(e):
                    bad.append((emp.get("name",""), e))
                else:
                    emp["email"] = e
                    try:
                        self.processor.remember_employee_email(emp, e)
                    except Exception:
                        pass

            if bad:
                lines = "\n".join([f"• {n}  ({e or 'empty'})" for n, e in bad])
                try:
                    messagebox.showerror("Invalid email(s)", f"Please fix:\n{lines}")
                except Exception:
                    print("Invalid emails:", lines)
                return

            # Persist to emails.csv (overwrite with the full known mapping)
            path = os.path.join(os.getcwd(), "emails.csv")
            try:
                with open(path, "w", newline="", encoding="utf-8") as fh:
                    writer = csv.DictWriter(fh, fieldnames=["seq","account_no","name","email"])
                    writer.writeheader()
                    for emp in self.processor.employees:
                        email = (emp.get("email") or "").strip()
                        if not email:
                            continue
                        writer.writerow({
                            "seq": emp.get("seq",""),
                            "account_no": emp.get("account_no",""),
                            "name": emp.get("name",""),
                            "email": email
                        })
                try:
                    if getattr(self.processor, "email_store", None):
                        self.processor.email_store.export_to_csv(path)
                except Exception:
                    pass
                # Feedback
                try:
                    self.log_result(f"Saved {path} and attached emails to employees.")
                except Exception:
                    print(f"Saved {path} and attached emails to employees.")
            except Exception as e:
                try:
                    messagebox.showerror("Save failed", str(e))
                except Exception:
                    print("Save failed:", e)
                return

            # Refresh the Employee View
            try:
                self.populate_employee_tree()
            except Exception:
                pass
            win.destroy()

        ttk.Button(btns, text="Save",  command=_save_and_apply).pack(side="right", padx=6)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="right", padx=6)

        # ----------------- helper methods -----------------
    def _apply_stored_emails(self):
        store = getattr(self.processor, "email_store", None)
        if store is None:
            try:
                messagebox.showinfo("Emails", "No stored email book is available yet.")
            except Exception:
                self.log_result("Email store not available.")
            return

        if not getattr(self.processor, "employees", None):
            try:
                messagebox.showwarning("Emails", "Load a payroll file first before applying stored emails.")
            except Exception:
                self.log_result("Cannot apply stored emails without loaded employees.")
            return

        employees = list(getattr(self.processor, "employees", []) or [])
        missing_before = sum(1 for emp in employees if not (emp.get("email") or "").strip())

        try:
            store.apply_to_employees(employees)
        except Exception as exc:
            try:
                messagebox.showerror("Emails", f"Could not apply stored emails: {exc}")
            except Exception:
                self.log_result(f"Could not apply stored emails: {exc}")
            return

        # Update GUI copy of data
        self.employee_data = list(employees)
        try:
            self.populate_employee_tree()
        except Exception:
            pass

        missing_after = sum(1 for emp in employees if not (emp.get("email") or "").strip())
        applied = missing_before - missing_after
        try:
            message = f"Applied stored emails. Updated {applied} records; {missing_after} still missing."
            self.log_result(message)
            messagebox.showinfo("Emails", message)
        except Exception:
            self.log_result(f"Stored emails applied. Updated {applied}; remaining {missing_after}.")

    def _reset_tree_columns(self):
        try:
        # base schema
            self.employee_tree['columns'] = tuple(self.static_columns)
            self._safe_set_displaycolumns(self.static_columns, self.static_columns)

        # reapply headings/widths
            for col in self.static_columns:
                try:
                    self.employee_tree.heading(col, text=col)
                    width = 200 if col == 'Name' else 120
                    self.employee_tree.column(col, width=width, anchor='center', stretch=False)
                except Exception:
                    pass

        # reset paging/state
            self.period_view_offset = 0
            self._fixed_period_ids = []
            self.all_period_ids = []
            self.all_columns = list(self.static_columns)
            self._two_page_mode = False
            self._page_number = 1
            self._page_split_index = None
        except Exception:
            pass
    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[('CSV files', '*.csv'), ('All files', '*.*')])
        if path:
            self.current_file = path
            self.file_path_var.set(path)

    def quick_select(self, filename):
        path = os.path.join(os.getcwd(), filename)
        if os.path.exists(path):
            self.current_file = path
            self.file_path_var.set(path)
        else:
            messagebox.showinfo('Not found', f'{filename} not found in current directory')

    def scan_directory(self):
        for f in os.listdir('.'):
            if f.lower().endswith('.csv'):
                self.current_file = os.path.abspath(f)
                self.file_path_var.set(self.current_file)
                messagebox.showinfo('File selected', f'Selected {f}')
                return
        messagebox.showinfo('Scan', 'No CSV files found in current directory')

    def start_processing(self):
        if not self.file_path_var.get():
            messagebox.showerror('Error', 'Please select a CSV file first')
            return
        self.current_file = self.file_path_var.get()
        try:
            from tkinter import simpledialog
            choice = simpledialog.askstring('Withholding placement', "Place withholding on: '15', '30', or 'both' (default: both)", parent=self.root)
            if choice is None:
                return
            choice = choice.strip().lower()
            if choice not in ('15', '30', 'both'):
                choice = 'both'
        except Exception:
            choice = 'both'
        self.withholding_placement = choice
        t = threading.Thread(target=self.process_payroll)
        t.daemon = True
        t.start()

    def process_payroll(self):
        try:
            self.status_var.set('Loading and detecting periods...')
            if not self.processor.detect_periods(self.current_file):
                self.log_result('Failed to detect periods')
            if not self.processor.load_employee_data(self.current_file):
                self.log_result('Failed to load employee data')
                return
            self.employee_data = list(self.processor.employees)
            self.log_result(f'Loaded {len(self.employee_data)} employees')
            self.root.after(0, self.populate_employee_tree)
            self.status_var.set('Generating PDFs...')
            success = 0
            for i, emp in enumerate(self.employee_data, 1):
                outdir = 'output_pdfs'
                os.makedirs(outdir, exist_ok=True)
                seq = int(emp.get('seq') or 0) if str(emp.get('seq', '')).isdigit() else 0
                clean = ''.join(c for c in emp.get('name', '') if c.isalnum() or c in (' ', '-', '_')).strip().replace(' ', '_')
                fname = f'payslip_{seq:03d}_{clean}.pdf'
                path = os.path.join(outdir, fname)
                # Log progress (processed so far / total) before generating
                total = len(self.employee_data)
                try:
                    self.root.after(0, lambda i=i, total=total, name=emp.get('name',''): self.log_result(f'({i-1}/{total}) Processing: {name}'))
                except Exception:
                    pass

                generated = False
                try:
                    generated = bool(self.processor.generate_pdf_payslip(emp, path, withholding_placement=getattr(self, 'withholding_placement', 'both')))
                except Exception as e:
                    # record error for this employee
                    try:
                        self.root.after(0, lambda i=i, total=total, name=emp.get('name',''), err=str(e): self.log_result(f'({i}/{total}) Error for {name}: {err}'))
                    except Exception:
                        pass

                if generated:
                    success += 1
                    try:
                        self.root.after(0, lambda i=i, total=total, fname=fname: self.log_result(f'({i}/{total}) Generated: {fname}'))
                    except Exception:
                        pass
                else:
                    try:
                        self.root.after(0, lambda i=i, total=total, name=emp.get('name',''): self.log_result(f'({i}/{total}) Skipped/No PDF for: {name}'))
                    except Exception:
                        pass
            self.log_result(f'Generated {success}/{len(self.employee_data)} PDFs')
            self.status_var.set('Processing complete')
        except Exception as e:
            self.log_result(f'Error: {e}')
            self.status_var.set('Error')

    def update_progress(self, message, value):
        self.progress_var = getattr(self, 'progress_var', tk.StringVar())
        self.progress_var.set(message)

    def log_result(self, text):
        try:
            self.results_text.insert(tk.END, text + '\n')
            self.results_text.see(tk.END)
        except Exception:
            pass

    def view_output_folder(self):
        out = os.path.abspath('output_pdfs')
        if os.path.exists(out):
            try:
                import subprocess
                if sys.platform.startswith('win'):
                    os.startfile(out)  # type: ignore[attr-defined]
                elif sys.platform == 'darwin':
                    subprocess.run(['open', out], check=False)
                else:
                    subprocess.run(['xdg-open', out], check=False)
            except Exception:
                messagebox.showinfo('Output', out)
        else:
            messagebox.showinfo('Output', f'Output folder: {out}')

    def load_employee_preview(self):
        if not self.file_path_var.get():
            messagebox.showerror('Error', 'Please select a CSV file first')
            return
        for it in self.employee_tree.get_children():
                self.employee_tree.delete(it)
        try:
            self.show_all_columns.set(False)
        except Exception:
            pass
        self._reset_tree_columns()
        self.current_file = self.file_path_var.get()
        temp = DynamicPayrollProcessor()
        if not temp.detect_periods(self.current_file):
            messagebox.showerror('Error', 'Could not detect periods in CSV')
            return
        if not temp.load_employee_data(self.current_file):
            messagebox.showerror('Error', 'Could not load employee data')
            return
        self.processor = temp
        self.employee_data = list(temp.employees)
        self.populate_employee_tree()
        self.show_period_summary(self.processor.periods)

    def show_period_summary(self, periods):
        self.period_text.config(state=tk.NORMAL)
        self.period_text.delete(1.0, tk.END)
        s = f'Detected {len(periods)} pay periods:\n'
        for i, p in enumerate(periods, 1):
            s += f'{i}. {p.get("name","Period") }\n'
        s += f'\nTotal Employees: {len(self.employee_data)}'
        self.period_text.insert(tk.END, s)
        self.period_text.config(state=tk.DISABLED)

    def clear_employee_data(self):
        for it in self.employee_tree.get_children():
            self.employee_tree.delete(it)
        self.employee_data = []
        try:
            self.show_all_columns.set(False)
        except Exception:
            pass
        self._reset_tree_columns()
        self.period_text.config(state=tk.NORMAL)
        self.period_text.delete(1.0, tk.END)
        self.period_text.config(state=tk.DISABLED)
    def populate_employee_tree(self):
        for it in self.employee_tree.get_children():
            self.employee_tree.delete(it)

        if not self.employee_data or not hasattr(self.processor, 'periods'):
            return

        # Build a parallel mapping: period_keys (used to lookup emp['periods']) and period_labels (displayed headings)
        period_keys = []
        period_labels = []
        for i, p in enumerate(self.processor.periods):
            # Prefer a canonical key if processor provides column_index or an explicit key
            key = None
            if isinstance(p.get('column_index'), int):
                key = f'col_{p.get("column_index")}'
            elif p.get('key') is not None:
                key = str(p.get('key'))
            else:
                # fallback to the period name (normalized)
                key = str(p.get('name') or f'Period {i+1}').strip()

            # label: prefer merged_headers when available (friendly display)
            label = None
            try:
                if hasattr(self.processor, 'merged_headers') and self.processor.merged_headers:
                    idx = p.get('column_index')
                    if isinstance(idx, int):
                        label = self.processor.merged_headers[idx]
            except Exception:
                label = None
            short_name = (p.get('name') or '').strip()
            if short_name:
                label = short_name
            elif not label:
                label = p.get('name') or f'Period {i+1}'

            period_keys.append(key)
            period_labels.append(label)

        expanded_period_ids = []
        col_label_map = {}
        # Use period_labels for display, but keep period_keys for lookup when building rows
        for i, name in enumerate(period_labels):
            id_col = f'period_{i}'
            expanded_period_ids.append(id_col)
            col_label_map[id_col] = name

        remaining_static = [c for c in self.static_columns[3:]]

        FIXED_PERIOD_COUNT = 5
        fixed_period_ids = expanded_period_ids[:min(FIXED_PERIOD_COUNT, len(expanded_period_ids))]

        all_columns = [self.static_columns[0], self.static_columns[1], self.static_columns[2]] + expanded_period_ids + remaining_static

        self.all_period_ids = expanded_period_ids
        self._fixed_period_ids = fixed_period_ids
        self.all_columns = all_columns

        try:
            split_idx = None
            if 'W/ Tax Rate' in self.all_columns:
                split_idx = self.all_columns.index('W/ Tax Rate')
            self._page_split_index = split_idx
            self._two_page_mode = bool(split_idx is not None and split_idx < (len(self.all_columns) - 1))
            self._page_number = 1
        except Exception:
            self._page_split_index = None
            self._two_page_mode = False
            self._page_number = 1

        try:
            self._period_heading_map = col_label_map.copy()
            heading_map = {c: c for c in [self.static_columns[0], self.static_columns[1], self.static_columns[2]] + remaining_static}
            heading_map.update(col_label_map)
            self._heading_map = heading_map
        except Exception:
            self._period_heading_map = {}
            self._heading_map = {}

        try:
            total_periods = len(self.all_period_ids)
            if total_periods > 1 and total_periods <= self.period_view_window:
                self.period_view_window = max(1, total_periods - 1)
        except Exception:
            pass
        try:
            total_periods = len(self.all_period_ids)
            self.period_view_window = max(1, min(self.period_view_window, total_periods))
            self.period_view_offset = 0
        except Exception:
            pass

        # Reset columns safely: set to static columns first and ensure displaycolumns matches
        try:
            self.employee_tree['columns'] = tuple(self.static_columns)
            self.employee_tree['displaycolumns'] = tuple(self.static_columns)
        except Exception:
            pass
        # Now set full columns list
        self.employee_tree['columns'] = tuple(all_columns)
        self._update_displaycolumns()

        for cid in expanded_period_ids:
            self.employee_tree.heading(cid, text=col_label_map.get(cid, cid))
            self.employee_tree.column(cid, width=96, anchor='center', stretch=False)

        for col in [self.static_columns[0], self.static_columns[1], self.static_columns[2]] + remaining_static:
            self.employee_tree.heading(col, text=col)
            w = 200 if col == 'Name' else 120
            self.employee_tree.column(col, width=w, anchor='center', stretch=False)

        try:
            self._column_width_map = {}
            for c in self.all_columns:
                try:
                    self._column_width_map[c] = int(self.employee_tree.column(c, option='width'))
                except Exception:
                    pass
        except Exception:
            self._column_width_map = {}

        for emp in self.employee_data:
            payroll = self.processor.calculate_employee_payroll(emp) if hasattr(self.processor, 'calculate_employee_payroll') else {'gross_pay': 0, 'net_pay': 0}
            period_values = []
            # Lookup periods using canonical keys first, then fallback to label names
            for k_idx, pname in enumerate(period_labels):
                lookup_key = period_keys[k_idx] if k_idx < len(period_keys) else pname
                pd = emp.get('periods', {}).get(lookup_key)
                if pd is None:
                    # fallback: try label-based key (older processor formats)
                    canon = ''
                    try:
                        canon = (self.processor.periods[k_idx].get('name') or '').strip()
                    except Exception:
                        canon = ''
                    pd = (
                        emp.get('periods', {}).get(pname)
                        or emp.get('periods', {}).get(pname.strip())
                        or (canon and emp.get('periods', {}).get(canon))
                        or {}
                    )
                if pd is None:
                    pd = {}
                hours = 0.0
                try:
                    hours = float(self.processor.safe_float(pd.get('hours', 0))) if pd else 0.0
                except Exception:
                    try:
                        hours = float(pd) if pd else 0.0
                    except Exception:
                        hours = 0.0
                period_values.append(f"{hours:.2f}")

            gross = emp.get('total_gross') if emp.get('total_gross') is not None else payroll.get('gross_pay', 0)
            adj = emp.get('adjustment_amount', 0) or 0

            def fmt_percent_field(v):
                if v is None:
                    return ''
                try:
                    v = float(v)
                except Exception:
                    return str(v)
                if abs(v) <= 1:
                    return f"{v*100:.0f}%"
                return f"{v}%"

            def fmt_currency(v):
                try:
                    return f"${float(v):,.2f}"
                except Exception:
                    return ''

            net_amount_earned_display = fmt_currency(gross - adj)
            w_tax_rate_display = fmt_percent_field(emp.get('tax_rate'))
            wa = emp.get('withholding_amount', None)
            if wa is not None and wa != 0:
                w_withholding_display = fmt_currency(wa)
            else:
                wh_val = emp.get('withholding_rate')
                if wh_val is None:
                    w_withholding_display = ''
                else:
                    try:
                        wnum = float(wh_val)
                        if abs(wnum) <= 1:
                            w_withholding_display = f"{wnum*100:.0f}%"
                        else:
                            w_withholding_display = fmt_currency(wnum)
                    except Exception:
                        w_withholding_display = str(wh_val)

            p_tax_display = fmt_percent_field(emp.get('p_tax_rate'))
            percent_tax_display = fmt_currency(emp.get('percent_tax', 0))
            total_tax_deductions_display = fmt_currency(emp.get('total_tax_deductions', 0))
            net_amount_received_display = fmt_currency(emp.get('net_amount_received', 0))

            row = [
                emp.get('seq', ''),
                emp.get('name', ''),
                f"${emp.get('hourly_rate', 0):.2f}",
                *period_values,
                emp.get('adjustment_hours', 0),
                fmt_currency(emp.get('adjustment_amount', 0)),
                net_amount_earned_display,
                w_tax_rate_display,
                w_withholding_display,
                p_tax_display,
                percent_tax_display,
                total_tax_deductions_display,
                net_amount_received_display
            ]
            self.employee_tree.insert('', tk.END, values=row)

        try:
            self.employee_tree.update_idletasks()
            self.employee_tree.xview_moveto(0.0)
        except Exception:
            pass

        try:
            left, right = self.employee_tree.xview()
            self.h_scroll.set(left, right)
        except Exception:
            pass

        self.show_period_summary(self.processor.periods)
        self.notebook.select(1)

    def _update_displaycolumns(self):
        fixed_left = [self.static_columns[0], self.static_columns[1], self.static_columns[2]]
        remaining_static = [c for c in self.static_columns[3:]]
        all_columns = getattr(self, 'all_columns', fixed_left + remaining_static)

        # Guard: no periods → just show static columns
        if not getattr(self, 'all_period_ids', None):
            displaycols = fixed_left + remaining_static
            try:
                self.employee_tree['columns'] = tuple(all_columns)
            except Exception:
                pass
            self._safe_set_displaycolumns(all_columns, displaycols)
            return

        try:
            show_all = bool(self.show_all_columns.get())
        except Exception:
            show_all = False
        if getattr(self, '_two_page_mode', False):
            all_cols = list(getattr(self, 'all_columns', fixed_left + remaining_static))
            split_idx = getattr(self, '_page_split_index', None)
            if split_idx is None:
                split_idx = len(all_cols) - 1
            right_cols = all_cols[3:]
            split_rel = max(0, split_idx - 3)
            if getattr(self, '_page_number', 1) == 1:
                page_slice = right_cols[: split_rel + 1]
            else:
                page_slice = right_cols[split_rel + 1 :]
            displaycols = fixed_left + page_slice
        else:
            if show_all and getattr(self, 'all_period_ids', None):
                visible_periods = list(self.all_period_ids)
            else:
                if not hasattr(self, 'all_period_ids') or not self.all_period_ids:
                    visible_periods = []
                else:
                    paged_periods = [p for p in self.all_period_ids if p not in getattr(self, '_fixed_period_ids', [])]
                    num_paged = len(paged_periods)
                    if num_paged <= 0:
                        visible_paged = []
                    else:
                        max_off = max(0, num_paged - self.period_view_window)
                        self.period_view_offset = max(0, min(self.period_view_offset, max_off))
                        start = self.period_view_offset
                        end = start + self.period_view_window
                        visible_paged = paged_periods[start:end]
                    visible_periods = list(getattr(self, '_fixed_period_ids', [])) + visible_paged
            displaycols = fixed_left + visible_periods + remaining_static

        try:
            self.employee_tree['columns'] = getattr(self, 'all_columns', displaycols)
        except Exception:
            pass
        self._safe_set_displaycolumns(all_columns, displaycols)

        try:
            heading_map = getattr(self, '_heading_map', {}) or {}
            for cid in displaycols:
                try:
                    self.employee_tree.heading(cid, text=heading_map.get(cid, cid))
                except Exception:
                    pass
        except Exception:
            pass

        try:
            width_map = getattr(self, '_column_width_map', {}) or {}
            for cid in displaycols:
                try:
                    w = width_map.get(cid)
                    if w is None:
                        if cid.startswith('period_'):
                            w = 96
                        elif cid == 'Name':
                            w = 200
                        else:
                            w = 120
                    self.employee_tree.column(cid, width=w, stretch=False)
                except Exception:
                    pass
        except Exception:
            pass

        try:
            for pid, label in (getattr(self, '_period_heading_map', {}) or {}).items():
                try:
                    self.employee_tree.heading(pid, text=label)
                except Exception:
                    pass
        except Exception:
            pass

        try:
            self.employee_tree.update_idletasks()
        except Exception:
            pass

        try:
            if getattr(self, '_two_page_mode', False):
                try:
                    self.page_var.set(f'Page {getattr(self, "_page_number", 1)}/2')
                except Exception:
                    pass
                try:
                    if getattr(self, 'pan_prev_btn', None):
                        self.pan_prev_btn.state(['disabled' if self._page_number <= 1 else '!disabled'])
                    if getattr(self, 'pan_next_btn', None):
                        self.pan_next_btn.state(['disabled' if self._page_number >= 2 else '!disabled'])
                except Exception:
                    pass
            else:
                paged_periods = [p for p in getattr(self, 'all_period_ids', []) if p not in getattr(self, '_fixed_period_ids', [])]
                num_paged = len(paged_periods)
                page_count = 1 if num_paged <= 0 else (num_paged + max(1, self.period_view_window) - 1) // max(1, self.period_view_window)
                current_page = 1 if num_paged <= 0 else (self.period_view_offset // max(1, self.period_view_window)) + 1
                try:
                    self.page_var.set(f'Page {current_page}/{page_count}')
                except Exception:
                    pass
                try:
                    if getattr(self, 'pan_prev_btn', None):
                        self.pan_prev_btn.state(['disabled' if current_page <= 1 else '!disabled'])
                    if getattr(self, 'pan_next_btn', None):
                        self.pan_next_btn.state(['disabled' if current_page >= page_count else '!disabled'])
                except Exception:
                    pass

            if getattr(self, 'h_scroll', None):
                if show_all:
                    try:
                        left, right = self.employee_tree.xview()
                        self.h_scroll.set(left, right)
                        try:
                            self.status_var.set('Scroll mode: native')
                        except Exception:
                            pass
                    except Exception:
                        self.h_scroll.set(0.0, 1.0)
                else:
                    if getattr(self, '_two_page_mode', False):
                        self.h_scroll.set(0.0, 1.0)
                        try:
                            self.status_var.set('Scroll mode: two-page')
                        except Exception:
                            pass
                    else:
                        paged_periods = [p for p in getattr(self, 'all_period_ids', []) if p not in getattr(self, '_fixed_period_ids', [])]
                        num_paged = len(paged_periods)
                        if num_paged <= 0 or num_paged <= self.period_view_window:
                            self.h_scroll.set(0.0, 1.0)
                            try:
                                self.status_var.set('Scroll mode: none')
                            except Exception:
                                pass
                        else:
                            frac = self.period_view_window / max(1, num_paged)
                            left = (current_page - 1) * self.period_view_window / max(1, num_paged)
                            self.h_scroll.set(left, min(1.0, left + frac))
                            try:
                                self.status_var.set('Scroll mode: paging')
                            except Exception:
                                pass
        except Exception:
            pass

    def _block_header_resize(self, event):
        try:
            region = self.employee_tree.identify_region(event.x, event.y)
            if region == 'separator':
                return "break"
        except Exception:
            pass

    def _on_show_all_columns_toggle(self):
        try:
            self._update_displaycolumns()
            try:
                show_all = bool(self.show_all_columns.get())
            except Exception:
                show_all = False
            if show_all:
                try:
                    left, right = self.employee_tree.xview()
                    self.h_scroll.set(left, right)
                except Exception:
                    pass
            else:
                self.employee_tree.xview_moveto(0.0)
                self.h_scroll.set(0.0, 1.0)
        except Exception:
            pass

    def _on_tree_xscroll(self, left, right):
        try:
            if getattr(self, 'h_scroll', None):
                try:
                    l = float(left)
                    r = float(right)
                except Exception:
                    l, r = left, right
                self.h_scroll.set(l, r)
        except Exception:
            pass

    def _hscroll_command(self, *args):
        try:
            try:
                show_all = bool(self.show_all_columns.get())
            except Exception:
                show_all = False
            if not hasattr(self, 'all_period_ids') or not self.all_period_ids:
                try:
                    self.employee_tree.xview(*args)
                except Exception:
                    pass
                return

            try:
                left, right = self.employee_tree.xview()
                width = right - left
            except Exception:
                left, right, width = 0.0, 1.0, 1.0

            if (not show_all) or abs(width - 1.0) < 1e-6:
                if args[0] == 'moveto':
                    try:
                        frac = float(args[1])
                    except Exception:
                        frac = 0.0
                    paged_periods = [p for p in self.all_period_ids if p not in getattr(self, '_fixed_period_ids', [])]
                    num_paged = len(paged_periods)
                    max_off = max(0, num_paged - self.period_view_window)
                    if max_off > 0:
                        page_count = (num_paged + max(1, self.period_view_window) - 1) // max(1, self.period_view_window)
                        page_index = int(round(frac * (page_count - 1))) if page_count > 1 else 0
                        new_offset = page_index * self.period_view_window
                        new_offset = max(0, min(max_off, new_offset))
                    else:
                        new_offset = 0
                    if new_offset != self.period_view_offset:
                        self.period_view_offset = new_offset
                        self._update_displaycolumns()
                elif args[0] == 'scroll':
                    n = int(args[1])
                    what = args[2] if len(args) > 2 else 'units'
                    step = self.period_view_window if what in ('pages',) else 1
                    new_offset = self.period_view_offset + (n * step)
                    paged_periods = [p for p in self.all_period_ids if p not in getattr(self, '_fixed_period_ids', [])]
                    num_paged = len(paged_periods)
                    max_off = max(0, num_paged - self.period_view_window)
                    new_offset = max(0, min(max_off, new_offset))
                    if new_offset != self.period_view_offset:
                        self.period_view_offset = new_offset
                        self._update_displaycolumns()
                try:
                    paged_periods = [p for p in self.all_period_ids if p not in getattr(self, '_fixed_period_ids', [])]
                    total = len(paged_periods)
                    if total <= 0:
                        self.h_scroll.set(0.0, 1.0)
                    else:
                        frac = self.period_view_window / max(1, total)
                        left = self.period_view_offset / max(1, total)
                        self.h_scroll.set(left, min(1.0, left + frac))
                except Exception:
                    pass
                return

            try:
                self.employee_tree.xview(*args)
            except Exception:
                pass
        except Exception:
            pass

    def pan_xview(self, delta_fraction):
        try:
            left, right = self.employee_tree.xview()
            width = right - left
        except Exception:
            left, right, width = 0.0, 1.0, 1.0

        try:
            self.log_result(
                f'pan_xview: left={left:.4f} right={right:.4f} width={width:.4f} '
                f'offset={self.period_view_offset} total_periods={len(getattr(self, "all_period_ids", []))} window={self.period_view_window}'
            )
        except Exception:
            pass

        if getattr(self, '_two_page_mode', False):
            try:
                if delta_fraction > 0:
                    self._page_number = min(2, getattr(self, '_page_number', 1) + 1)
                else:
                    self._page_number = max(1, getattr(self, '_page_number', 1) - 1)
                self._update_displaycolumns()
            except Exception:
                pass
            return

        try:
            show_all = bool(self.show_all_columns.get())
        except Exception:
            show_all = False
        if (not show_all) or abs(width - 1.0) < 1e-6:
            step_pages = 1 if delta_fraction > 0 else -1
            paged_periods = [p for p in getattr(self, 'all_period_ids', []) if p not in getattr(self, '_fixed_period_ids', [])]
            num_paged = len(paged_periods)
            max_off = max(0, num_paged - self.period_view_window)

            new_offset = self.period_view_offset + (step_pages * self.period_view_window)
            new_offset = max(0, min(max_off, new_offset))

            if new_offset == self.period_view_offset:
                try:
                    self.log_result(f'period_view_offset -> {self.period_view_offset} (clamped; max_off={max_off})')
                except Exception:
                    pass
                return

            self.period_view_offset = new_offset
            self._update_displaycolumns()
            try:
                self.log_result(f'period_view_offset -> {self.period_view_offset} (after move; max_off={max_off})')
            except Exception:
                pass
            return

        try:
            new_left = max(0.0, min(1.0 - width, left + delta_fraction))
            self.employee_tree.xview_moveto(new_left)
            try:
                self.h_scroll.set(new_left, new_left + width)
            except Exception:
                pass
        except Exception as e:
            try:
                self.log_result(f'pan_xview error: {e}')
            except Exception:
                pass

    def sort_employee_data(self, col):
        pass

    def show_employee_periods(self, event):
        sel = self.employee_tree.selection()
        if not sel:
            return
        vals = self.employee_tree.item(sel[0])['values']
        messagebox.showinfo('Employee', f'Selected: {vals}')
    # _send_test_pdf removed - not required in packaged exe


def main():
    root = tk.Tk()
    app = DynamicPayrollGUIGenerator(root)
    root.mainloop()


if __name__ == '__main__':
    main()
