"""Simple Tkinter GUI to run the exam seat allocator.

This module provides a minimal interface to select input files and run
the allocation/export routines from the project's allocator class.

The GUI is only created when run as a script (i.e. guarded by
if __name__ == '__main__'), so importing this module is safe in
headless or test environments.
"""
from __future__ import annotations

import os
import traceback
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Optional

# Prefer tkcalendar if available; fallback handled at runtime
try:
    from tkcalendar import DateEntry
except Exception:
    DateEntry = None


def _safe_import_allocator():
    """Attempt to import ExamSeatAllocator and return it.

    The project contains two similar allocator implementations with
    different constructor signatures. We import the class and let the
    runtime decide which constructor to call.
    """
    try:
        from University_Exam_Seat_Allocater import ExamSeatAllocator
        return ExamSeatAllocator
    except Exception:
        # Bubble up a clearer error
        raise


class ExamGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Exam Seat Allocator")

        self.mode_var = tk.StringVar(value="single")  # 'single' or 'separate'

        frame = tk.Frame(root, padx=10, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # Mode selection
        mode_frame = tk.Frame(frame)
        mode_frame.pack(fill=tk.X, pady=(0, 6))
        tk.Radiobutton(mode_frame, text="Single Excel (Students+Halls sheet)", variable=self.mode_var, value="single").pack(anchor=tk.W)
        tk.Radiobutton(mode_frame, text="Separate Students and Halls files", variable=self.mode_var, value="separate").pack(anchor=tk.W)

        # Single file chooser
        # Default single-excel filename (keeps existing convention)
        default_excel = os.path.join(os.getcwd(), 'exam_seat_allocation.xlsx')
        self.single_path_var = tk.StringVar(value=default_excel)
        single_row = tk.Frame(frame)
        single_row.pack(fill=tk.X, pady=4)
        tk.Label(single_row, text="Excel file:", width=12).pack(side=tk.LEFT)
        tk.Entry(single_row, textvariable=self.single_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(single_row, text="Browse", command=self.browse_single).pack(side=tk.LEFT, padx=6)
        tk.Button(single_row, text="Preview", command=self.preview_single).pack(side=tk.LEFT, padx=6)

        # Separate students/halls
        self.students_path_var = tk.StringVar()
        self.halls_path_var = tk.StringVar()

        stud_row = tk.Frame(frame)
        stud_row.pack(fill=tk.X, pady=4)
        tk.Label(stud_row, text="Students file:", width=12).pack(side=tk.LEFT)
        tk.Entry(stud_row, textvariable=self.students_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(stud_row, text="Browse", command=self.browse_students).pack(side=tk.LEFT, padx=6)
        tk.Button(stud_row, text="Preview", command=self.preview_students).pack(side=tk.LEFT, padx=6)

        halls_row = tk.Frame(frame)
        halls_row.pack(fill=tk.X, pady=4)
        tk.Label(halls_row, text="Halls file:", width=12).pack(side=tk.LEFT)
        tk.Entry(halls_row, textvariable=self.halls_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(halls_row, text="Browse", command=self.browse_halls).pack(side=tk.LEFT, padx=6)
        tk.Button(halls_row, text="Preview", command=self.preview_halls).pack(side=tk.LEFT, padx=6)

        # Output options
        out_row = tk.Frame(frame)
        out_row.pack(fill=tk.X, pady=(8, 4))
        tk.Label(out_row, text="Output folder:", width=12).pack(side=tk.LEFT)
        self.out_dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "pdf_reports"))
        tk.Entry(out_row, textvariable=self.out_dir_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(out_row, text="Browse", command=self.browse_output).pack(side=tk.LEFT, padx=6)

        # Date and Session inputs
        opts_row = tk.Frame(frame)
        opts_row.pack(fill=tk.X, pady=(6, 4))
        tk.Label(opts_row, text="Exam Date:", width=12).pack(side=tk.LEFT)

        # Use tkcalendar.DateEntry if available (dropdown calendar); otherwise fall back to spinboxes
        # Try to create a DateEntry; use a per-instance flag to avoid rebinding the module symbol
        self._use_dateentry = False
        if DateEntry is not None:
            try:
                # DateEntry supports date_pattern; keep format dd-MM-yyyy for consistency
                self.date_entry = DateEntry(opts_row, date_pattern='dd-MM-yyyy')
                self.date_entry.pack(side=tk.LEFT)
                self._use_dateentry = True
            except Exception:
                # If DateEntry instantiation fails for any reason, fall back to spinboxes
                self._use_dateentry = False

        if DateEntry is None:
            # Fallback: spinboxes
            import datetime
            today = datetime.date.today()
            self.day_var = tk.StringVar(value=f"{today.day:02d}")
            self.month_var = tk.StringVar(value=f"{today.month:02d}")
            self.year_var = tk.StringVar(value=str(today.year))
            tk.Spinbox(opts_row, from_=1, to=31, width=3, textvariable=self.day_var, format="%02.0f").pack(side=tk.LEFT)
            tk.Label(opts_row, text="/").pack(side=tk.LEFT)
            tk.Spinbox(opts_row, from_=1, to=12, width=3, textvariable=self.month_var, format="%02.0f").pack(side=tk.LEFT)
            tk.Label(opts_row, text="/").pack(side=tk.LEFT)
            current_year = today.year
            tk.Spinbox(opts_row, from_=current_year-5, to=current_year+5, width=6, textvariable=self.year_var).pack(side=tk.LEFT)

        tk.Label(opts_row, text="Session:", width=8).pack(side=tk.LEFT, padx=(6,0))
        self.session_var = tk.StringVar(value="FN")
        tk.OptionMenu(opts_row, self.session_var, "FN", "AN").pack(side=tk.LEFT)

        # Buttons
        btn_row = tk.Frame(frame)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        tk.Button(btn_row, text="Run Allocation", command=self.run_allocation_threaded, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=(0,6))
        tk.Button(btn_row, text="Preview (print to console)", command=self.preview_allocation).pack(side=tk.LEFT)

        # Log area
        log_label = tk.Label(frame, text="Log:")
        log_label.pack(anchor=tk.W, pady=(8, 0))
        self.log = tk.Text(frame, height=10)
        self.log.pack(fill=tk.BOTH, expand=True)

    def browse_single(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            initialdir=os.getcwd(),
            filetypes=[("Excel files", ("*.xlsx", "*.xls")), ("All files", "*")]
        )
        if path:
            self.single_path_var.set(path)

    def browse_students(self):
        path = filedialog.askopenfilename(
            title="Select Students file",
            initialdir=os.getcwd(),
            filetypes=[("Excel files", ("*.xlsx", "*.xls")), ("All files", "*")]
        )
        if path:
            self.students_path_var.set(path)

    def browse_halls(self):
        path = filedialog.askopenfilename(
            title="Select Halls file",
            initialdir=os.getcwd(),
            filetypes=[("Excel files", ("*.xlsx", "*.xls")), ("All files", "*")]
        )
        if path:
            self.halls_path_var.set(path)

    def browse_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.out_dir_var.set(path)


    def _preview_excel(self, path: str):
        """Open a small window showing sheet names and a small head() preview of the selected sheet."""
        if not path or not os.path.exists(path):
            messagebox.showwarning("Preview", "File not found or not specified.")
            return
        try:
            # lazy import pandas here to avoid GUI-only dependency at import time
            import pandas as _pd
            xls = _pd.read_excel(path, sheet_name=None)
        except Exception as e:
            messagebox.showerror("Preview Error", f"Could not read Excel file:\n{e}")
            return

        preview_win = tk.Toplevel(self.root)
        preview_win.title(f"Preview: {os.path.basename(path)}")
        preview_win.geometry("800x500")

        sheets = list(xls.keys())
        sheet_var = tk.StringVar(value=sheets[0] if sheets else "")

        top_row = tk.Frame(preview_win)
        top_row.pack(fill=tk.X, padx=6, pady=6)
        tk.Label(top_row, text="Sheet:").pack(side=tk.LEFT)
        sheet_menu = tk.OptionMenu(top_row, sheet_var, *sheets)
        sheet_menu.pack(side=tk.LEFT, padx=(4, 8))

        txt = tk.Text(preview_win)
        txt.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0,6))

        def show_sheet(*_):
            s = sheet_var.get()
            txt.delete("1.0", tk.END)
            try:
                df = xls[s]
                txt.insert(tk.END, df.head(50).to_string(index=False))
            except Exception as e:
                txt.insert(tk.END, f"Could not render sheet '{s}': {e}")

        # refresh when selection changes
        sheet_var.trace_add('write', lambda *_: show_sheet())

        # initial render
        if sheets:
            show_sheet()

        btn_row = tk.Frame(preview_win)
        btn_row.pack(fill=tk.X, padx=6, pady=6)
        tk.Button(btn_row, text="Close", command=preview_win.destroy).pack(side=tk.RIGHT)


    def preview_single(self):
        path = self.single_path_var.get()
        self._preview_excel(path)

    def preview_students(self):
        path = self.students_path_var.get()
        self._preview_excel(path)

    def preview_halls(self):
        path = self.halls_path_var.get()
        self._preview_excel(path)

    def _append_log(self, *lines):
        for l in lines:
            self.log.insert(tk.END, f"{l}\n")
        self.log.see(tk.END)

    def preview_allocation(self):
        # Run allocation but only print seating plan to console (no PDFs)
        try:
            Exam = _safe_import_allocator()
            allocator = self._make_allocator(Exam)
            allocator.allocate_seats()

            # Propagate selected date/session into allocations (if allocation code doesn't already set it)
            # Build selected_date string in DD-MM-YYYY format from DateEntry (or spinboxes fallback)
            try:
                if getattr(self, '_use_dateentry', False) and hasattr(self, 'date_entry'):
                    # DateEntry provides a datetime.date via get_date()
                    selected_date = self.date_entry.get_date().strftime('%d-%m-%Y')
                else:
                    d = int(self.day_var.get())
                    m = int(self.month_var.get())
                    y = int(self.year_var.get())
                    selected_date = f"{d:02d}-{m:02d}-{y}"
            except Exception:
                selected_date = ""
            selected_session = self.session_var.get().strip()
            try:
                for reg_no, alloc in allocator.allocations.items():
                    if selected_date:
                        alloc['date'] = selected_date
                    alloc['session'] = selected_session
                # also set attributes for allocators that might read them
                setattr(allocator, 'exam_date', selected_date)
                setattr(allocator, 'session', selected_session)
            except Exception:
                # If allocations isn't populated yet or isn't a dict-like, ignore
                pass
            allocator.print_seating_plan()
            messagebox.showinfo("Preview", "Printed seating plan to console.")
        except Exception as e:
            tb = traceback.format_exc()
            self._append_log(tb)
            messagebox.showerror("Error", str(e))

    def run_allocation_threaded(self):
        t = threading.Thread(target=self.run_allocation, daemon=True)
        t.start()

    def run_allocation(self):
        try:
            Exam = _safe_import_allocator()
            allocator = self._make_allocator(Exam)
            self._append_log("Starting allocation...")
            allocator.allocate_seats()

            # Propagate selected date/session into allocations (so exports can use them)
            # Build selected_date string in DD-MM-YYYY format from DateEntry (or spinboxes fallback)
            try:
                if getattr(self, '_use_dateentry', False) and hasattr(self, 'date_entry'):
                    selected_date = self.date_entry.get_date().strftime('%d-%m-%Y')
                else:
                    d = int(self.day_var.get())
                    m = int(self.month_var.get())
                    y = int(self.year_var.get())
                    selected_date = f"{d:02d}-{m:02d}-{y}"
            except Exception:
                selected_date = ""
            selected_session = self.session_var.get().strip()
            try:
                for reg_no, alloc in allocator.allocations.items():
                    if selected_date:
                        alloc['date'] = selected_date
                    alloc['session'] = selected_session
                setattr(allocator, 'exam_date', selected_date)
                setattr(allocator, 'session', selected_session)
            except Exception:
                pass

            out_dir = self.out_dir_var.get() or "pdf_reports"
            os.makedirs(out_dir, exist_ok=True)

            # Build sanitized filename prefix from date and session
            def _sanitize(s: str) -> str:
                # Keep common safe characters, replace others with '_'
                import re
                if not s:
                    return ""
                s = s.strip()
                # replace spaces and slashes with '-'
                s = s.replace(' ', '_').replace('/', '-').replace('\\', '-')
                # replace any character not alnum, dash, or underscore with '_'
                return re.sub(r"[^A-Za-z0-9_\-]", '_', s)

            date_part = _sanitize(selected_date)
            session_part = _sanitize(selected_session)
            prefix = ""
            if date_part and session_part:
                prefix = f"{date_part}_{session_part}"
            elif date_part:
                prefix = date_part
            elif session_part:
                prefix = session_part

            seating_fname = f"{prefix + '_' if prefix else ''}Seating_Plans.pdf"
            master_fname = f"{prefix + '_' if prefix else ''}Master_Seating_Plan.pdf"

            # Export using the allocator methods; try common signatures with fallbacks
            try:
                allocator.export_pdf_seating_plan(output_dir=out_dir, filename=seating_fname)
            except TypeError:
                # older signature might accept filename only or no args
                try:
                    allocator.export_pdf_seating_plan(filename=seating_fname)
                except TypeError:
                    try:
                        allocator.export_pdf_seating_plan(seating_fname)
                    except Exception:
                        # Last resort: call without filename
                        allocator.export_pdf_seating_plan()

            # Master export: try output_dir+output_file, then output_file only, then positional
            try:
                allocator.export_master_seating_plan(output_dir=out_dir, output_file=master_fname)
            except TypeError:
                try:
                    allocator.export_master_seating_plan(output_file=master_fname)
                except TypeError:
                    try:
                        allocator.export_master_seating_plan(master_fname)
                    except Exception:
                        try:
                            allocator.export_master_seating_plan()
                        except Exception:
                            pass

            self._append_log("Allocation and exports completed.")
            messagebox.showinfo("Done", f"Allocation complete. PDFs (if any) written to: {out_dir}")
        except Exception as e:
            tb = traceback.format_exc()
            self._append_log(tb)
            messagebox.showerror("Error", str(e))

    def _make_allocator(self, Exam):
        """Construct an ExamSeatAllocator instance using the available inputs.

        We try the single-excel constructor first and fall back to the
        two-file constructor if it fails.
        """
        mode = self.mode_var.get()
        if mode == "single":
            excel = self.single_path_var.get()
            if not excel:
                raise ValueError("Please select an Excel file.")
            try:
                return Exam(excel_file=excel)
            except TypeError:
                # fallback: maybe constructor expects a single positional path
                return Exam(excel)
        else:
            studs = self.students_path_var.get()
            halls = self.halls_path_var.get()
            if not studs or not halls:
                raise ValueError("Please select both students and halls files.")
            try:
                return Exam(students_file=studs, halls_file=halls)
            except TypeError:
                # maybe constructor accepts two positional args
                return Exam(studs, halls)


def main():
    root = tk.Tk()
    app = ExamGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
