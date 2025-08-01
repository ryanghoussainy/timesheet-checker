__version__ = "1.0.0" # Major.Minor.Patch

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import tkinterdnd2 as tkdnd
import os
import threading
import sys
import re
from PIL import Image, ImageTk


from check_timesheets import check_timesheets
from amindefy import amindefy_timesheets
from colours import *

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
    
    def get_user_input(self, message=None):
        # If message is provided directly, use it
        if message:
            result = self.show_match_confirmation_dialog(message)
            return result if result else "exit"
        
        # Extract the prompt from the last few lines written to output
        all_text = self.output_text.get(1.0, tk.END).strip()
        lines = all_text.split('\n')
        
        # Look for match confirmation prompts
        prompt = ""
        swimmer_info = ""
        
        for line in reversed(lines):
            clean_line = re.sub(r'\033\[\d+m', '', line).strip()
            if "Is this the right match?" in clean_line:
                prompt = clean_line
                break
            elif "->" in clean_line and "similarity score" in clean_line:
                swimmer_info = clean_line
                break
        
        # Show popup dialog for match confirmation
        if prompt or swimmer_info:
            result = self.show_match_confirmation_dialog(swimmer_info or prompt)
            return result if result else "exit"
        
        # Fallback for any other input (shouldn't happen with current implementation)
        return "n"

    def show_match_confirmation_dialog(self, message):
        """Show a custom dialog for match confirmation"""
        # Extract info from the message for a cleaner dialog
        if "->" in message:
            # Parse the match info
            match_info = re.sub(r'\033\[\d+m', '', message).strip()
            dialog_message = f"Confirm match:\n\n{match_info}"
        else:
            dialog_message = message
    
        # Create custom dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Confirm Match")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()  # Make dialog modal
        
        # Dialog dimensions
        dialog_width = 400
        dialog_height = 200
        
        # Force the parent window to update its geometry info
        self.root.update_idletasks()
        
        # Now get accurate parent window information
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        
        # Calculate centered position
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        # Ensure dialog stays on screen
        x = max(0, x)
        y = max(0, y)
        
        # Set the position explicitly
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
    
        # Message label
        msg_label = tk.Label(
            dialog,
            text=dialog_message,
            font=("Segoe UI", 11),
            wraplength=350,
            justify=tk.LEFT,
            padx=20,
            pady=20
        )
        msg_label.pack(expand=True)
    
        # Button frame
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
    
        result = {"value": "exit"}
    
        def accept():
            result["value"] = "y"
            dialog.destroy()
    
        def deny():
            result["value"] = "n"
            dialog.destroy()

        def on_close():
            result["value"] = "exit"
            dialog.destroy()

        # Bind the close event
        dialog.protocol("WM_DELETE_WINDOW", on_close)
    
        # Accept button
        accept_btn = tk.Button(
            button_frame,
            text="Accept Match",
            command=accept,
            bg=BUTTON_ACCEPT_BG,
            fg=BUTTON_ACCEPT_FG,
            font=("Segoe UI", 10, "bold"),
            padx=20,
            pady=8
        )
        accept_btn.pack(side=tk.LEFT, padx=10)
    
        # Deny button
        deny_btn = tk.Button(
            button_frame,
            text="Deny Match",
            command=deny,
            bg=BUTTON_DENY_BG,
            fg=BUTTON_DENY_FG,
            font=("Segoe UI", 10, "bold"),
            padx=20,
            pady=8
        )
        deny_btn.pack(side=tk.LEFT, padx=10)
        
        # Ensure dialog is visible and focused
        dialog.deiconify()
        dialog.lift()
        dialog.focus_force()
    
        # Wait for dialog to close
        dialog.wait_window()
    
        return result["value"]
    
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
    
    def create_check_timesheets_tab(self):
        frame = tk.Frame(self.notebook, bg=NOTEBOOK_TAB_BACKGROUND)
        self.notebook.add(frame, text="2. Check Timesheets")
        
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
        self.create_file_input(frame, "First Excel File", 'amindefied_excel', [('Excel files', '*.xls *.xlsx')])
        self.create_file_input(frame, "Second Excel File", 'sign_in_sheet', [('Excel files', '*.xls *.xlsx')])
        
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
        path_var.set("No location selected (will use default: timesheets_excel.xlsx)")
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
                    amindefy_timesheets(self.file_paths['folder_path'], self.file_paths.get('output_file', 'timesheets_excel.xlsx'))
                
                self._write_to_output(f"\n✅ FOLDER PROCESSED SUCCESSFULLY!\n")
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
                    check_timesheets(self.file_paths['amindefied_excel'], self.file_paths['sign_in_sheet'])
                
                self._write_to_output(f"\n✅ TIMESHEET CHECK COMPLETED!\n")
            except Exception as e:
                self._write_to_output(f"\n❌ ERROR: {str(e)}\n")
        
        # Run in separate thread to prevent GUI freezing
        threading.Thread(target=process, daemon=True).start()

def main():
    root = tkdnd.TkinterDnD.Tk()
    app = TimesheetCheckerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()