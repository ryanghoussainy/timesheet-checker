__version__ = "1.0.0" # Major.Minor.Patch

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import tkinterdnd2 as tkdnd
import os
import threading
import sys
import re
from PIL import Image, ImageTk
import json

from check_timesheets import check_timesheets
from amindefy import amindefy_timesheets
from colours import *

RATES_FILE = "rates.json"

class OutputCapture:
    """Context manager to capture print statements and user input"""
    def __init__(self, output_widget, input_callback=None):
        self.output_widget = output_widget
        self.input_callback = input_callback
        self.old_stdout = None
        self.old_stdin = None
        
    def __enter__(self):
        self.old_stdout = sys.stdout
        self.old_stdin = sys.stdin
        sys.stdout = self
        if self.input_callback:
            sys.stdin = self
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        sys.stdout = self.old_stdout
        sys.stdin = self.old_stdin
        
    def write(self, text):
        # Write all text, including newlines
        if self.output_widget.winfo_exists():
            try:
                self.output_widget.after_idle(lambda: self._write_to_widget(text))
            except tk.TclError:
                pass  # Widget destroyed
        return len(text)
        
    def _write_to_widget(self, text):
        try:
            if not self.output_widget.winfo_exists():
                return
                
            self.output_widget.config(state='normal')
            self._insert_coloured_text(text)
            self.output_widget.see(tk.END)
            self.output_widget.config(state='disabled')
        except tk.TclError:
            pass  # Widget destroyed or not accessible
    
    def _insert_coloured_text(self, text):
        """Insert text with colour support by parsing ANSI escape codes"""
        # ANSI colour code pattern
        ansi_pattern = r'\033\[(\d+)m'
        
        # Split text by ANSI codes
        parts = re.split(ansi_pattern, text)
        
        current_colour = None
        for i, part in enumerate(parts):
            if i % 2 == 0:  # Text part
                if current_colour:
                    self.output_widget.insert(tk.END, part, current_colour)
                else:
                    self.output_widget.insert(tk.END, part)
            else:  # colour code part
                colour_code = int(part)
                if colour_code == 91:  # Red
                    current_colour = "red"
                elif colour_code == 93:  # Yellow
                    current_colour = "yellow"
                elif colour_code == 92:  # Green
                    current_colour = "green"
                elif colour_code == 0:  # Reset
                    current_colour = None
        
    def flush(self):
        pass
        
    def readline(self):
        if self.input_callback:
            return self.input_callback()
        return "\n"

class TimesheetCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Timesheet Checker")
        self.root.geometry("1200x800")
        self.root.configure(bg=APP_BACKGROUND)

        # Setup modern styles
        self.setup_modern_styles()
        
        # File paths storage
        self.file_paths = {
            'folder_path': None,
            'amindefied_excel': None,
            'sign_in_sheet': None
        }
        
        self.setup_ui()

    def resource_path(self, relative_path):
        """Get absolute path to resource, works for development and PyInstaller"""
        try:
            base_path = sys._MEIPASS  # For PyInstaller
        except AttributeError:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)
        
    def setup_ui(self):
        # Main title with logo
        title_frame = tk.Frame(self.root, bg=APP_BACKGROUND)
        title_frame.pack(pady=(10, 0))

        # Try to load and display logo
        try:
            # Load logo
            logo_path = self.resource_path("images/esc-logo.png")
            if os.path.exists(logo_path):
                logo = Image.open(logo_path)
                # Resize logo
                logo = logo.resize((48, 48), Image.Resampling.LANCZOS)
                self.title_image = ImageTk.PhotoImage(logo)

                # Image label
                image_label = tk.Label(
                    title_frame,
                    image=self.title_image,
                    bg=APP_BACKGROUND
                )
                image_label.pack(side=tk.LEFT, padx=(0, 15))
        except Exception as e:
            print(f"Could not load logo: {e}")

        # Title label
        title_label = tk.Label(
            title_frame, 
            text="Timesheet Checker", 
            font=("Segoe UI", 18, "bold"),
            bg=APP_BACKGROUND,
            fg=APP_TITLE,
        )
        title_label.pack(side=tk.LEFT)

        # Add version label
        version_label = tk.Label(
            title_frame,
            text=f"Version {__version__}",
            font=("Segoe UI", 10),
            bg=APP_BACKGROUND,
            fg=APP_TITLE,
        )
        version_label.pack(side=tk.RIGHT, padx=(0, 20))
        
        # Create main container with paned window
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL, style="Modern.TFrame")
        main_paned.pack(expand=True, fill='both', padx=30, pady=20)
        
        # Left side - Controls
        left_frame = tk.Frame(main_paned, bg=FRAME_BACKGROUND)
        main_paned.add(left_frame, weight=1)
        
        # Right side - Output
        right_frame = tk.Frame(main_paned, bg=FRAME_BACKGROUND)
        main_paned.add(right_frame, weight=1)
        
        # Create notebook for tabs on left side
        self.notebook = ttk.Notebook(left_frame, style="Modern.TNotebook")
        self.notebook.pack(expand=True, fill='both', padx=10, pady=20)
        
        # Tab 1: Folder Processing
        self.create_amindefy_tab()
        self.create_rates_tab()  # <-- Add this line for the new tab
        # Tab 2: Excel Comparison
        self.create_check_timesheets_tab()
        
        # Output panel on right side
        self.create_output_panel(right_frame)

    def setup_modern_styles(self):
        # Configure modern ttk styles
        style = ttk.Style()
        
        # Configure modern button style
        style.configure(
            'Modern.TButton',
            background=CONTAINER_BACKGROUND,
            font=('Segoe UI', 12),
        )
        
        # Configure modern frame style
        style.configure(
            'Modern.TFrame',
            background=FRAME_BACKGROUND,
            relief='flat'
        )
        
        # Configure modern notebook style
        style.configure(
            'Modern.TNotebook',
            background=NOTEBOOK_BACKGROUND,
            borderwidth=0
        )
        style.configure(
            'Modern.TNotebook.Tab',
            padding=(20, 10),
            font=('Segoe UI', 12)
        )
    
    def create_output_panel(self, parent):
        # Output text area with modern styling
        self.output_text = scrolledtext.ScrolledText(
            parent,
            wrap=tk.WORD,
            width=50,
            height=20,
            font=("JetBrains Mono", 12),
            bg=OUTPUT_BACKGROUND,
            fg=OUTPUT_FOREGROUND,
            state='disabled',
            borderwidth=0,
            highlightthickness=1,
            highlightcolor=OUTPUT_HIGHLIGHT,
            insertbackground=OUTPUT_CURSOR,
            relief='flat'
        )
        self.output_text.pack(expand=True, fill='both', pady=(0, 20), padx=10)
        
        # Configure modern colour tags
        self.output_text.tag_configure("red", foreground=TEXT_RED)
        self.output_text.tag_configure("yellow", foreground=TEXT_YELLOW)
        self.output_text.tag_configure("green", foreground=TEXT_GREEN)
    
    def clear_output(self):
        self.output_text.config(state='normal')
        self.output_text.delete(1.0, tk.END)
        self.output_text.config(state='disabled')
    
    def wait_for_enter(self):
        """Wait until the user presses Enter in the output terminal."""
        # Create a variable to signal when Enter is pressed
        self._enter_pressed = False

        def on_key(event):
            if event.keysym == "Return":
                self._enter_pressed = True
                # Remove the binding after Enter is pressed
                self.output_text.unbind("<Key>")

        # Enable the output_text widget to receive key events
        self.output_text.config(state='normal')
        self.output_text.focus_set()
        self.output_text.bind("<Key>", on_key)

        # Wait until Enter is pressed
        while not self._enter_pressed:
            self.output_text.update()
        
        # Restore output_text to disabled state
        self.output_text.config(state='disabled')
        """Wait until the user presses Enter in a popup dialog."""
        result = {"pressed": False}

    def get_user_input(self, message=None):
        self.wait_for_enter()
        return "\n"
    
    def _write_to_output(self, text):
        self.output_text.config(state='normal')
        self.output_text.insert(tk.END, text)
        self.output_text.see(tk.END)
        self.output_text.config(state='disabled')
    
    def create_amindefy_tab(self):
        frame = tk.Frame(self.notebook, bg=NOTEBOOK_TAB_BACKGROUND)
        self.notebook.add(frame, text="1. Amindefy Timesheets")
        
        # Instructions
        instructions = tk.Label(
            frame,
            text="Combine all timesheets in a folder into a single Excel file.",
            font=("Segoe UI", 12),
            bg=LABEL_BACKGROUND,
            fg=LABEL_FOREGROUND,
            wraplength=400
        )
        instructions.pack(pady=20)
        
        # Folder input area
        self.create_folder_input(frame, "Timesheets Folder", 'folder_path')

        # Output file selection
        self.create_output_file_input(frame, "Output Excel File", 'output_file', [('Excel files', '*.xlsx')])
        
        # Process button
        process_btn = ttk.Button(
            frame,
            text="Create Amindefied Excel File",
            command=self.run_amindefy,
            style="Modern.TButton",
        )
        process_btn.pack(pady=30)
    
    def create_rates_tab(self):
        frame = tk.Frame(self.notebook, bg=NOTEBOOK_TAB_BACKGROUND)
        self.notebook.add(frame, text="2. Rates")

        instructions = tk.Label(
            frame,
            text="Review and edit rates for each level. If any are incorrect, please change them and click Save.",
            font=("Segoe UI", 12),
            bg=LABEL_BACKGROUND,
            fg=LABEL_FOREGROUND,
            wraplength=400
        )
        instructions.pack(pady=20)

        # Rate change checkbox and date (UI only)
        self.rate_change_var = tk.BooleanVar(value=False)
        self.rate_change_date_var = tk.StringVar(value="")

        rate_change_frame = tk.Frame(frame, bg=CONTAINER_BACKGROUND)
        rate_change_frame.pack(padx=10, pady=5, anchor="w")

        rate_change_check = tk.Checkbutton(
            rate_change_frame,
            text="Rate change",
            variable=self.rate_change_var,
            bg=CONTAINER_BACKGROUND,
            command=self.toggle_rate_change,
        )
        rate_change_check.pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(rate_change_frame, text="Change date (DD/MM/YYYY):", bg=CONTAINER_BACKGROUND).pack(side=tk.LEFT)
        self.rate_change_date_entry = tk.Entry(rate_change_frame, textvariable=self.rate_change_date_var, state="disabled")
        self.rate_change_date_entry.pack(side=tk.LEFT, padx=(5, 0))

        # Load nested rates format (rates, rates_after, date)
        rates, rates_after, rate_change_date = self.load_rates()
        self.rates = rates
        # if rates_after is None, keep internal copy for editing but hide the table UI
        self.rates_after = {k: float(v) for k, v in rates.items()} if rates_after is None else {k: float(v) for k, v in rates_after.items()}

        # set date UI if present
        if rate_change_date:
            self.rate_change_date_var.set(rate_change_date)
            self.rate_change_var.set(True)
            self.rate_change_date_entry.config(state='normal')
        else:
            self.rate_change_var.set(False)
            self.rate_change_date_entry.config(state='disabled')

        self.rate_vars = {}
        self.rate_vars_after = {}

        # Parent container holding two table frames side-by-side
        tables_container = tk.Frame(frame, bg=CONTAINER_BACKGROUND)
        tables_container.pack(padx=10, pady=10, fill=tk.X)

        # Left table (current rates)
        self.table_frame = tk.Frame(tables_container, bg=CONTAINER_BACKGROUND)
        self.table_frame.pack(side=tk.LEFT, padx=(0, 20))

        tk.Label(self.table_frame, text="Level", font=("Arial", 11, "bold"), bg=CONTAINER_BACKGROUND).grid(row=0, column=0, padx=10, pady=5)
        tk.Label(self.table_frame, text="Rate (£/hr)", font=("Arial", 11, "bold"), bg=CONTAINER_BACKGROUND).grid(row=0, column=1, padx=10, pady=5)

        for i, (level, rate) in enumerate(self.rates.items(), start=1):
            tk.Label(self.table_frame, text=level, font=("Arial", 11), bg=CONTAINER_BACKGROUND).grid(row=i, column=0, padx=10, pady=5, sticky="w")
            var = tk.StringVar(value=f"{float(rate):.2f}")
            entry = tk.Entry(self.table_frame, textvariable=var, width=10, font=("Arial", 11))
            entry.grid(row=i, column=1, padx=10, pady=5)
            self.rate_vars[level] = var

        # Right table (rates after change) — create but may be hidden depending on rates_after
        self.table_frame_after = tk.Frame(tables_container, bg=CONTAINER_BACKGROUND)
        tk.Label(self.table_frame_after, text="Level", font=("Arial", 11, "bold"), bg=CONTAINER_BACKGROUND).grid(row=0, column=0, padx=10, pady=5)
        tk.Label(self.table_frame_after, text="Rate (£/hr) (after)", font=("Arial", 11, "bold"), bg=CONTAINER_BACKGROUND).grid(row=0, column=1, padx=10, pady=5)

        for i, level in enumerate(self.rates.keys(), start=1):
            tk.Label(self.table_frame_after, text=level, font=("Arial", 11), bg=CONTAINER_BACKGROUND).grid(row=i, column=0, padx=10, pady=5, sticky="w")
            after_value = self.rates_after.get(level, 0.0)
            var_after = tk.StringVar(value=f"{float(after_value):.2f}")
            entry_after = tk.Entry(self.table_frame_after, textvariable=var_after, width=10, font=("Arial", 11))
            entry_after.grid(row=i, column=1, padx=10, pady=5)
            self.rate_vars_after[level] = var_after

        # Show after-table only if rates_after was present in file (not None) or checkbox is set
        if rates_after is not None:
            self.table_frame_after.pack(side=tk.LEFT)

        save_btn = ttk.Button(
            frame,
            text="Save",
            command=self.on_save_rates,
            style="Modern.TButton",
        )
        save_btn.pack(pady=20)
    
    def toggle_rate_change(self):
        """Show/hide the after-table. If turning on and rates_after was None, copy current rates into it."""
        if self.rate_change_var.get():
            # ensure internal rates_after exists and populate from current rates if it was None
            if getattr(self, "rates_after", None) is None:
                self.rates_after = {k: float(v) for k, v in self.rates.items()}
                # update UI entries to match
                for lvl, var in self.rate_vars_after.items():
                    var.set(f"{self.rates_after.get(lvl, 0.0):.2f}")
            self.table_frame_after.pack(side=tk.LEFT, padx=(0, 0))
            self.rate_change_date_entry.config(state='normal')
        else:
            # Hide it and mark internal as None (so save writes null)
            self.table_frame_after.pack_forget()
            self.rate_change_date_entry.config(state='disabled')
            self.rates_after = None

    def load_rates(self):
        """
        Load nested rates JSON:
        { "rate_change_date": "DD/MM/YYYY" | null,
            "rates": {...},
            "rates_after": {...} | null
        }
        Returns (rates_dict, rates_after_dict_or_None, rate_change_date_or_None)
        """
        try:
            with open(RATES_FILE, "r") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                raise ValueError("rates.json must contain an object")
            rates = {str(k): float(v) for k, v in data.get("rates", {}).items()}
            rates_after_raw = data.get("rates_after", None)
            rates_after = None if rates_after_raw is None else {str(k): float(v) for k, v in rates_after_raw.items()}
            rate_change_date = data.get("rate_change_date", None)
            # If file contained only a flat dict (older format), treat that as rates (back-compat)
            if not rates:
                # check if top-level keys look like rate levels (flat mapping)
                flat_candidate = {str(k): float(v) for k, v in data.items() if k not in ("rates_after", "rate_change_date")}
                if flat_candidate:
                    return flat_candidate, None, None
            return rates, rates_after, rate_change_date
        except Exception:
            levels = [
                "L1", "L2", "NQL2", "Enhanced L2", "Lower Enhanced L2",
                "Safeguarding", "Admin", "Gala Full Day", "Gala Half Day"
            ]
            return {level: 0.0 for level in levels}, None, None

    def save_rates(self):
        """
        Save nested structure:
        { "rate_change_date": <str or null>, "rates": {...}, "rates_after": {...} or null }
        If rate_change checkbox is unchecked, rates_after will be saved as null.
        """
        try:
            # choose which rates_after to persist: either dict or None
            rates_to_save = {k: float(v) for k, v in self.rates.items()}
            if getattr(self, "rate_change_var", None) and self.rate_change_var.get() and getattr(self, "rates_after", None) is not None:
                rates_after_to_save = {k: float(v) for k, v in self.rates_after.items()}
            else:
                rates_after_to_save = None

            data = {
                "rate_change_date": self.rate_change_date_var.get() if (getattr(self, "rate_change_var", None) and self.rate_change_var.get()) else None,
                "rates": rates_to_save,
                "rates_after": rates_after_to_save,
            }

            with open(RATES_FILE, "w") as f:
                json.dump(data, f, indent=2)

            messagebox.showinfo("Saved", "Rates saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save rates: {e}")

    def on_save_rates(self):
        """Validate both tables and then persist nested JSON (rates + optional rates_after + date)."""
        new_rates = {}
        new_rates_after = None

        # Validate left table (current rates)
        for level, var in self.rate_vars.items():
            try:
                value = float(var.get())
                new_rates[level] = value
            except ValueError:
                messagebox.showerror("Error", f"Invalid rate for {level}: {var.get()}")
                return

        # Validate right table only if visible (rate change checked)
        if getattr(self, 'rate_change_var', None) and self.rate_change_var.get():

            if not self.rate_change_date_var.get():
                messagebox.showerror("Error", "Please enter a valid rate change date (DD/MM/YYYY).")
                return

            new_rates_after = {}
            for level, var in self.rate_vars_after.items():
                try:
                    value = float(var.get())
                    new_rates_after[level] = value
                except ValueError:
                    messagebox.showerror("Error", f"Invalid rate for {level} (after change): {var.get()}")
                    return

        # assign to instance
        self.rates = new_rates
        self.rates_after = new_rates_after

        # Save nested structure (save_rates writes rates_after as null if None)
        self.save_rates()

    def create_check_timesheets_tab(self):
        frame = tk.Frame(self.notebook, bg=NOTEBOOK_TAB_BACKGROUND)
        self.notebook.add(frame, text="2. Check Timesheets")
        
        # Month dropdown menu
        months = [
            "September", "October", "November", "December", "January",
            "February", "March", "April", "May", "June", "July"
        ]
        self.month_var = tk.StringVar(value=months[0])
        self.month = months[0]  # Default month

        def on_month_change(*args):
            self.month = self.month_var.get()

        self.month_var.trace_add("write", on_month_change)

        month_label = tk.Label(
            frame,
            text="Select Month:",
            font=("Segoe UI", 11, "bold"),
            bg=LABEL_BACKGROUND,
            fg=LABEL_FOREGROUND
        )
        month_label.pack(pady=(15, 5))

        month_dropdown = ttk.Combobox(
            frame,
            textvariable=self.month_var,
            values=months,
            state="readonly",
            font=("Segoe UI", 11)
        )
        month_dropdown.pack(pady=(0, 15))

        # Instructions
        instructions = tk.Label(
            frame,
            text="Compare the timesheets with the sign in sheet.",
            font=("Segoe UI", 12),
            bg=LABEL_BACKGROUND,
            fg=LABEL_FOREGROUND,
            wraplength=400
        )
        instructions.pack(pady=20)
        
        # File input areas
        self.create_file_input(frame, "Timesheets Excel File", 'amindefied_excel', [('Excel files', '*.xls *.xlsx')])
        self.create_file_input(frame, "Sign In Sheet", 'sign_in_sheet', [('Excel files', '*.xls *.xlsx')])
        
        # Process button
        process_btn = ttk.Button(
            frame,
            text="Check Timesheets",
            command=self.run_check_timesheets,
            style="Modern.TButton",
        )
        process_btn.pack(pady=30)
    
    def create_folder_input(self, parent, label_text, key):
        # Container frame
        container = tk.Frame(parent, bg=CONTAINER_BACKGROUND, relief=tk.RAISED, bd=1)
        container.pack(fill=tk.X, padx=10, pady=10)
        
        # Label
        label = tk.Label(container, text=label_text, font=("Arial", 10, "bold"), bg=CONTAINER_BACKGROUND)
        label.pack(anchor=tk.W, padx=10, pady=(10, 5))
        
        # Drop area
        drop_frame = tk.Frame(container, bg=DROP_AREA_DEFAULT_BG, height=60, relief=tk.SUNKEN, bd=2)
        drop_frame.pack(fill=tk.X, padx=10, pady=5)
        drop_frame.pack_propagate(False)
        
        drop_label = tk.Label(
            drop_frame,
            text="Drag & Drop folder here or click to browse",
            bg=DROP_AREA_DEFAULT_BG,
            fg=DROP_AREA_DEFAULT_FG,
            font=("Arial", 11, "italic"),
        )
        drop_label.pack(expand=True)
        
        # File path display
        path_var = tk.StringVar()
        path_label = tk.Label(container, textvariable=path_var, bg=CONTAINER_BACKGROUND, fg=LABEL_TEXT_DARK, wraplength=400)
        path_label.pack(anchor=tk.W, padx=10, pady=(0, 10))
        
        # Store references
        setattr(self, f'{key}_var', path_var)
        setattr(self, f'{key}_frame', drop_frame)
        
        # Bind click event
        drop_frame.bind("<Button-1>", lambda e: self.browse_folder(key))
        drop_label.bind("<Button-1>", lambda e: self.browse_folder(key))
        
        # Enable drag and drop
        drop_frame.drop_target_register(tkdnd.DND_FILES)
        drop_frame.dnd_bind('<<Drop>>', lambda e: self.handle_folder_drop(e, key))

    def create_file_input(self, parent, label_text, key, filetypes):
        # Container frame
        container = tk.Frame(parent, bg=CONTAINER_BACKGROUND, relief=tk.RAISED, bd=1)
        container.pack(fill=tk.X, padx=10, pady=10)
        
        # Label
        label = tk.Label(container, text=label_text, font=("Arial", 10, "bold"), bg=CONTAINER_BACKGROUND)
        label.pack(anchor=tk.W, padx=10, pady=(10, 5))
        
        # Drop area
        drop_frame = tk.Frame(container, bg=DROP_AREA_DEFAULT_BG, height=60, relief=tk.SUNKEN, bd=2)
        drop_frame.pack(fill=tk.X, padx=10, pady=5)
        drop_frame.pack_propagate(False)
        
        drop_label = tk.Label(
            drop_frame,
            text="Drag & Drop file here or click to browse",
            bg=DROP_AREA_DEFAULT_BG,
            fg=DROP_AREA_DEFAULT_FG,
            font=("Arial", 11, "italic"),
        )
        drop_label.pack(expand=True)
        
        # File path display
        path_var = tk.StringVar()
        path_label = tk.Label(container, textvariable=path_var, bg=CONTAINER_BACKGROUND, fg=LABEL_TEXT_DARK, wraplength=400)
        path_label.pack(anchor=tk.W, padx=10, pady=(0, 10))
        
        # Store references
        setattr(self, f'{key}_var', path_var)
        setattr(self, f'{key}_frame', drop_frame)
        
        # Bind click event
        drop_frame.bind("<Button-1>", lambda e: self.browse_file(key, filetypes))
        drop_label.bind("<Button-1>", lambda e: self.browse_file(key, filetypes))
        
        # Enable drag and drop
        drop_frame.drop_target_register(tkdnd.DND_FILES)
        
    def create_output_file_input(self, parent, label_text, key, filetypes):
        # Container frame
        container = tk.Frame(parent, bg=CONTAINER_BACKGROUND, relief=tk.RAISED, bd=1)
        container.pack(fill=tk.X, padx=10, pady=10)
        
        # Label
        label = tk.Label(container, text=label_text, font=("Arial", 10, "bold"), bg=CONTAINER_BACKGROUND)
        label.pack(anchor=tk.W, padx=10, pady=(10, 5))
        
        # Browse button and path display in same row
        button_frame = tk.Frame(container, bg=CONTAINER_BACKGROUND)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        browse_btn = ttk.Button(
            button_frame,
            text="Choose Location",
            command=lambda: self.browse_output_file(key, filetypes),
            style="Modern.TButton",
        )
        browse_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # File path display
        path_var = tk.StringVar()
        path_var.set("No location selected (will use default: all_timesheets.xlsx)")
        path_label = tk.Label(button_frame, textvariable=path_var, bg=CONTAINER_BACKGROUND, fg=LABEL_TEXT_DARK, wraplength=300)
        path_label.pack(side=tk.LEFT, anchor=tk.W, pady=(0, 10))
        
        # Store references
        setattr(self, f'{key}_var', path_var)

    def browse_output_file(self, key, filetypes):
        filename = filedialog.asksaveasfilename(
            title=f"Choose output file location",
            defaultextension=".xlsx",
            filetypes=filetypes + [('All files', '*.*')]
        )
        if filename:
            self.file_paths[key] = filename
            path_var = getattr(self, f'{key}_var')
            path_var.set(f"Will save to: {os.path.basename(filename)}")

    def browse_folder(self, key):
        folder = filedialog.askdirectory(title="Select folder")
        if folder:
            self.set_file_path(key, folder)

    def browse_file(self, key, filetypes):
        filename = filedialog.askopenfilename(
            title=f"Select file for {key}",
            filetypes=filetypes + [('All files', '*.*')]
        )
        if filename:
            self.set_file_path(key, filename)
    
    def handle_folder_drop(self, event, key):
        files = self.root.tk.splitlist(event.data)
        if files:
            path = files[0]
            if os.path.isdir(path):
                self.set_file_path(key, path)

    def handle_drop(self, event, key):
        files = self.root.tk.splitlist(event.data)
        if files:
            self.set_file_path(key, files[0])
    
    def set_file_path(self, key, path):
        self.file_paths[key] = path
        path_var = getattr(self, f'{key}_var')
        path_var.set(f"Selected: {os.path.basename(path)}")
        
        # Update drop area appearance
        drop_frame = getattr(self, f'{key}_frame')
        drop_frame.configure(bg=DROP_AREA_SUCCESS_BG)
        for child in drop_frame.winfo_children():
            if isinstance(child, tk.Label):
                child.configure(bg=DROP_AREA_SUCCESS_BG, fg=DROP_AREA_SUCCESS_FG, text="✓ File loaded")
    
    def run_amindefy(self):
        if not self.file_paths['folder_path']:
            messagebox.showerror("Error", "Please select a folder")
            return
        
        def process():
            try:
                self.clear_output()
                
                with OutputCapture(self.output_text, self.get_user_input):
                    print("Processing folder...")
                    print(f"Folder: {self.file_paths['folder_path']}")
                    amindefy_timesheets(self.file_paths['folder_path'], self.file_paths.get('output_file', 'all_timesheets.xlsx'))
                
                self._write_to_output(f"\n✅ TIMESHEETS PROCESSED SUCCESSFULLY!\n")
            except Exception as e:
                self._write_to_output(f"\n❌ ERROR: {str(e)}\n")
        
        # Run in separate thread to prevent GUI freezing
        threading.Thread(target=process, daemon=True).start()
    
    def run_check_timesheets(self):
        if not self.file_paths['amindefied_excel'] or not self.file_paths['sign_in_sheet']:
            messagebox.showerror("Error", "Please select both Excel files")
            return

        def process():
            try:
                self.clear_output()
                with OutputCapture(self.output_text, self.get_user_input):
                    # load nested rates triple and pass all three to check_timesheets
                    rates, rates_after, rate_change_date = self.load_rates()
                    check_timesheets(
                        self.file_paths['amindefied_excel'],
                        self.file_paths['sign_in_sheet'],
                        rates,
                        rates_after,
                        rate_change_date,
                        self.month
                    )
                self._write_to_output(f"\n✅ TIMESHEET CHECK COMPLETED!\n")
            except Exception as e:
                self._write_to_output(f"\n❌ ERROR: {str(e)}\n")

        threading.Thread(target=process, daemon=True).start()


def main():
    root = tkdnd.TkinterDnD.Tk()
    app = TimesheetCheckerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()