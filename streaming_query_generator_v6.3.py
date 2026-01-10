"""
SQL Query Generator for Streaming Data Analysis
Created by: Fadzli Abdullah

KEY FEATURES:
1. AUTO-LOAD: Automatically loads MAPPING.csv from the same directory
2. eNodeB NAME CONVERTER (DEFAULT): Converts eNodeB Names to full 7-digit ECI using database lookup
   - Format: eNodeB_Name → looks up eNodeB_ID → converts to hex (5) + Cell_Number (2 hex) = 7-digit ECI
   - Example: SNOTM → eNodeB 260398 (3F92E) + Cells 1-3 (01, 02, 03) = 3F92E01, 3F92E02, 3F92E03
   - AUTOMATICALLY ADDS converted values to ECI Selection (no manual copy/paste needed!)
   - Supports eNodeB Names (e.g., SNOTM, MEBUM, AKOIM)
3. SECTOR ID CONVERTER: Converts Sector IDs (e.g., SNOTM_2, MEBUM_3) to full 7-digit ECI
   - Sector IDs must contain underscore (format: XXXXX_Y, e.g., SNOTM_2, MEBUM_3)
4. TRIPLE CONVERTER: Supports eNodeB Name-to-Hex, Decimal-to-Hex, and Sector ID-to-Hex conversion
5. FLEXIBLE HEX: Supports 5-8 digit hexadecimal values

USAGE:
1. Place MAPPING.csv in the same directory as this script
2. Run the application
3. In "eNodeB Name to Hex" mode (DEFAULT):
   - Paste your eNodeB Names (e.g., SNOTM, MEBUM, AKOIM)
   - Click "Convert & Add All"
   - Converted 7-digit ECIs are AUTOMATICALLY added to ECI Selection!
4. Or switch to "Sector ID to Hex" mode for specific sectors (e.g., SNOTM_2, MEBUM_3)
5. Or switch to "Decimal to Hex" mode for decimal eNodeB IDs
6. Select dates, RAT, and applications
7. Generate your SQL query

ECI Structure: 
- Full ECI = eNodeB_ID (5 hex digits) + Cell_Number (2 hex digits)
- Example: 3F92E02 = eNodeB 260398 (3F92E) + Cell 2 (02)
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
from datetime import datetime, timedelta, date
from tkcalendar import DateEntry
import pyperclip
import re
import os
import sys

class ColoredButton(tk.Canvas):
    """Custom button widget that matches ttk button appearance but with custom colors"""
    def __init__(self, parent, text, command, bg_color='#006400', fg_color='white', **kwargs):
        # Create canvas with appropriate size
        super().__init__(parent, height=26, highlightthickness=0, **kwargs)
        
        self.bg_color = bg_color
        self.hover_color = '#004d00'  # Darker green for hover
        self.fg_color = fg_color
        self.command = command
        self.text = text
        
        # Calculate button width based on text
        self.button_width = len(text) * 8 + 20
        self.config(width=self.button_width)
        
        # Draw the button
        self.draw_button(self.bg_color)
        
        # Bind events
        self.bind('<Button-1>', self.on_click)
        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)
        
    def draw_button(self, color):
        """Draw rounded rectangle button"""
        self.delete('all')
        # Draw rounded rectangle background
        self.create_rectangle(1, 1, self.button_width-1, 25, 
                            fill=color, outline='#808080', width=1,
                            tags='button_bg')
        # Draw text
        self.create_text(self.button_width//2, 13, text=self.text, 
                        fill=self.fg_color, font=('TkDefaultFont', 9),
                        tags='button_text')
        
    def on_click(self, event):
        """Handle button click"""
        if self.command:
            self.command()
    
    def on_enter(self, event):
        """Handle mouse enter"""
        self.draw_button(self.hover_color)
        self.config(cursor='hand2')
    
    def on_leave(self, event):
        """Handle mouse leave"""
        self.draw_button(self.bg_color)
        self.config(cursor='')

class StreamingQueryGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("SQL Query Generator - Fadzli Abdullah")
        
        # Set window size - optimized for 14" laptop display
        window_width = 650
        window_height = 720  # Increased to better accommodate converter + query area

        # Apply geometry and allow resizing so the window height fits content better
        try:
            # Center the window if possible
            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()
            center_x = int((screen_width - window_width) / 2)
            center_y = int((screen_height - window_height) / 2)
            self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        except Exception:
            # Fallback to a reasonable default size
            self.root.geometry(f"{window_width}x{window_height}")

        self.root.minsize(window_width, 600)
        self.root.resizable(True, True)

        # Store selected ECIs
        self.selected_ecis = []
        
        # Cell name to eNodeB_ID mapping dictionary
        self.cell_mapping = {}  # Format: {'AKOIM_1': 110345, 'AKOIM_2': 110345, ...}
        
        # eNodeB name to ID mapping dictionary (derived from cell_mapping)
        self.enodeb_mapping = {}  # Format: {'AKOIM': 110345, 'SNOTM': 260398, ...}
        
        # Application data
        self.apps = {
            '342': 'YouTube',
            '829': 'Facebook',
            '1181': 'Instagram',
            '4860': 'TikTok'
        }
        
        # App selection variables
        self.app_vars = {}
        self.select_all_var = tk.BooleanVar(value=False)
        
        # Create main container with reduced padding
        main_frame = ttk.Frame(root, padding="5")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)  # Changed to make column 0 expandable
        main_frame.columnconfigure(1, weight=1)  # Make column 1 expandable as well
        main_frame.rowconfigure(7, weight=1)  # Updated row for query text (was 6, now 7)
               
        # ECI Selection Section with reduced padding
        eci_frame = ttk.LabelFrame(main_frame, text="ECI Selection", padding="5")
        eci_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=3)
        eci_frame.columnconfigure(5, weight=1)  # Make the text area column expandable
        
        # ECI Input
        ttk.Label(eci_frame, text="Enter ECI (7-digit hex):").grid(row=0, column=0, sticky=tk.W)
        self.eci_entry = ttk.Entry(eci_frame, width=15)
        self.eci_entry.grid(row=0, column=1, padx=3)
        self.eci_entry.bind('<Return>', lambda e: self.add_eci())
        
        ttk.Button(eci_frame, text="Add ECI", command=self.add_eci).grid(row=0, column=2, padx=3)
        
        # Dark green Paste Bulk button
        self.paste_bulk_btn = ColoredButton(eci_frame, "Paste Clipboard", self.paste_bulk_eci)
        self.paste_bulk_btn.grid(row=0, column=3, padx=3)
        
        ttk.Button(eci_frame, text="Clear All", command=self.clear_ecis).grid(row=0, column=4, padx=3)
        
        # ECI Text Display (simple text, no checkboxes) - EXTENDED WIDTH
        text_frame = ttk.Frame(eci_frame)
        text_frame.grid(row=1, column=0, columnspan=6, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)  # Make text widget expandable
        
        # Create text widget with scrollbar for ECI display
        self.eci_text = tk.Text(text_frame, height=3, wrap=tk.WORD, 
                                font=('Courier', 9), relief=tk.SUNKEN, borderwidth=1)
        eci_scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.eci_text.yview)
        self.eci_text.configure(yscrollcommand=eci_scrollbar.set)
        
        self.eci_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        eci_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # NEW: Triple-Mode Converter Section (eNodeB Name, Decimal & Sector ID)
        converter_frame = ttk.LabelFrame(main_frame, text="Converter: eNodeB Name → Hex | Decimal ↔ Hex | Sector ID → Hex", padding="5")
        converter_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=3)
        converter_frame.columnconfigure(0, weight=1)
        converter_frame.columnconfigure(1, weight=1)
        
        # Mode selection
        mode_frame = ttk.Frame(converter_frame)
        mode_frame.grid(row=0, column=0, columnspan=2, pady=(0, 5))
        
        self.converter_mode = tk.StringVar(value="enodebname")
        ttk.Radiobutton(mode_frame, text="eNodeB Name to Hex", variable=self.converter_mode, 
                       value="enodebname", command=self.switch_converter_mode).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="Decimal to Hex", variable=self.converter_mode, 
                       value="decimal", command=self.switch_converter_mode).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="Sector ID to Hex", variable=self.converter_mode, 
                       value="sectorid", command=self.switch_converter_mode).pack(side=tk.LEFT, padx=5)
        
        # Load mapping file button (for eNodeB name and sector ID modes)
        self.load_mapping_btn = ttk.Button(mode_frame, text="Reload Mapping", 
                                          command=self.load_cell_mapping, state='normal')
        self.load_mapping_btn.pack(side=tk.LEFT, padx=5)
        
        self.mapping_status = ttk.Label(mode_frame, text="No mapping loaded", 
                                       font=('Ubuntu', 8), foreground='red')
        self.mapping_status.pack(side=tk.LEFT, padx=5)
        
        # Create two-column layout for input and output
        input_frame = ttk.Frame(converter_frame)
        input_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 3))
        input_frame.columnconfigure(0, weight=1)
        input_frame.rowconfigure(1, weight=1)
        
        output_frame = ttk.Frame(converter_frame)
        output_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(3, 0))
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(1, weight=1)
        
        # Input side
        self.input_label = ttk.Label(input_frame, text="eNodeB Names (one per line or comma-separated):")
        self.input_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 2))
        
        self.converter_input_text = tk.Text(input_frame, height=4, width=25, wrap=tk.WORD,
                                    font=('Courier', 9), relief=tk.SUNKEN, borderwidth=1)
        input_scroll = ttk.Scrollbar(input_frame, orient=tk.VERTICAL, command=self.converter_input_text.yview)
        self.converter_input_text.configure(yscrollcommand=input_scroll.set)
        self.converter_input_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        input_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        # Output side
        ttk.Label(output_frame, text="Hexadecimal Results (5-8 digit):").grid(row=0, column=0, sticky=tk.W, pady=(0, 2))
        
        self.hex_result_text = tk.Text(output_frame, height=4, width=25, wrap=tk.WORD,
                                       font=('Courier', 9), relief=tk.SUNKEN, borderwidth=1,
                                       state='disabled')
        hex_scroll = ttk.Scrollbar(output_frame, orient=tk.VERTICAL, command=self.hex_result_text.yview)
        self.hex_result_text.configure(yscrollcommand=hex_scroll.set)
        self.hex_result_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        hex_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        # Button frame
        button_frame = ttk.Frame(converter_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(5, 0))
        
        ttk.Button(button_frame, text="Convert & Add All", command=self.convert_and_add_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Clear Converter", command=self.clear_converter).pack(side=tk.LEFT, padx=2)
        
        # Dark green Paste button for converter
        self.paste_converter_btn = ColoredButton(button_frame, "Paste Clipboard", self.paste_converter_values)
        self.paste_converter_btn.pack(side=tk.LEFT, padx=2)
        
        # Info label for converter
        self.converter_info = ttk.Label(converter_frame, 
                                   text="eNodeB Name mode (DEFAULT): Converts to 5-digit hex - includes ALL cells! Use Sector ID mode for specific cells.",
                                   font=('Ubuntu', 8), foreground='gray')
        self.converter_info.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=3, pady=(3, 0))
        
        # Date Selection Section with reduced padding and Quick Selection
        date_frame = ttk.LabelFrame(main_frame, text="Date Selection & Partition", padding="5")
        date_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)
        
        # Quick Selection Buttons
        quick_frame = ttk.Frame(date_frame)
        quick_frame.grid(row=0, column=0, columnspan=4, pady=3)
        
        ttk.Label(quick_frame, text="Quick Selection:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(quick_frame, text="3 Days", command=lambda: self.quick_select_days(3)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_frame, text="7 Days", command=lambda: self.quick_select_days(7)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_frame, text="14 Days", command=lambda: self.quick_select_days(14)).pack(side=tk.LEFT, padx=2)
        
        # Reference date (today)
        ttk.Label(date_frame, text="Reference Date (Today):").grid(row=1, column=0, sticky=tk.W, padx=3)
        self.reference_date = DateEntry(date_frame, width=12, background='darkblue',
                                        foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.reference_date.set_date(date.today())
        self.reference_date.grid(row=1, column=1, padx=3)
        
        # Info label
        info_label = ttk.Label(date_frame, text="(p0=T (today), p1=T-1, p2=T-2, p3=T-3, etc)",
                              font=('Ubuntu', 8), foreground='gray')
        info_label.grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=3)
        
        # Date range selection
        ttk.Label(date_frame, text="Start Date:").grid(row=2, column=0, sticky=tk.W, padx=3, pady=3)
        self.start_date = DateEntry(date_frame, width=12, background='darkblue',
                                    foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.start_date.grid(row=2, column=1, padx=3)
        self.start_date.bind("<<DateEntrySelected>>", self.calculate_partitions)
        
        ttk.Label(date_frame, text="End Date:").grid(row=2, column=2, sticky=tk.W, padx=3)
        self.end_date = DateEntry(date_frame, width=12, background='darkblue',
                                  foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.end_date.grid(row=2, column=3, padx=3)
        self.end_date.bind("<<DateEntrySelected>>", self.calculate_partitions)
        
        # Partition display
        ttk.Label(date_frame, text="Partitions:").grid(row=3, column=0, sticky=tk.W, padx=3, pady=3)
        self.partition_var = tk.StringVar()
        partition_entry = ttk.Entry(date_frame, textvariable=self.partition_var, 
                                   width=45, state='readonly')
        partition_entry.grid(row=3, column=1, columnspan=3, sticky=(tk.W, tk.E), padx=3)
        
        # RAT Selection
        rat_frame = ttk.LabelFrame(main_frame, text="RAT Selection", padding="5")
        rat_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)
        
        self.rat_var = tk.StringVar(value="6")
        ttk.Radiobutton(rat_frame, text="LTE (6)", variable=self.rat_var, 
                       value="6").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(rat_frame, text="5G NR (9)", variable=self.rat_var, 
                       value="9").grid(row=0, column=1, padx=5)
        ttk.Radiobutton(rat_frame, text="Both (6,9)", variable=self.rat_var,
                       value="6,9").grid(row=0, column=2, padx=5)
        
        # Add checkbox for optional Resolution column
        self.include_resolution_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(rat_frame, text="Include Resolution (uses video_data_rate column)", 
                       variable=self.include_resolution_var).grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=5, pady=(5,0))
        
        # Application Selection with reduced padding
        app_frame = ttk.LabelFrame(main_frame, text="Application Selection", padding="5")
        # Place the Application Selection in row 5 spanning both columns to align with layout
        app_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)
        
        # Select All checkbox (left) and individual app checkboxes on a single horizontal row
        select_all_check = ttk.Checkbutton(app_frame, text="Select All", 
                                          variable=self.select_all_var,
                                          command=self.toggle_all_apps)
        select_all_check.grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)

        # Place all app checkboxes on the same row (row 0), starting from column 1
        col = 1
        for app_id, app_name in self.apps.items():
            var = tk.BooleanVar(value=False)
            self.app_vars[app_id] = var
            var.trace('w', self.check_all_apps_selected)

            check = ttk.Checkbutton(app_frame, text=f"{app_name} ({app_id})", 
                                   variable=var)
            check.grid(row=0, column=col, padx=5, pady=2, sticky=tk.W)
            col += 1

        # Make app_frame columns non-stretch by default; layout will remain horizontal
        for i in range(col):
            app_frame.columnconfigure(i, weight=0)
        
        # Generate Query Button with reduced padding
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=5)
        
        generate_btn = ttk.Button(button_frame, text="Generate Query", 
                                 command=self.generate_query)
        generate_btn.pack(side=tk.LEFT, padx=5)
        
        copy_btn = ttk.Button(button_frame, text="Copy to Clipboard", 
                             command=self.copy_to_clipboard)
        copy_btn.pack(side=tk.LEFT, padx=5)
        
        save_btn = ttk.Button(button_frame, text="Save to File", 
                             command=self.save_to_file)
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # Query Display with reduced padding
        query_frame = ttk.LabelFrame(main_frame, text="Generated Query", padding="5")
        query_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=3)
        query_frame.columnconfigure(0, weight=1)
        query_frame.rowconfigure(0, weight=1)
        
        # Compact height for Generated Query so overall window is not too tall
        self.query_text = scrolledtext.ScrolledText(query_frame, wrap=tk.WORD, 
                    width=70, height=18,
                    font=('Courier', 9))
        self.query_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Status Bar with reduced padding
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)
        status_frame.columnconfigure(0, weight=1)
        
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                relief=tk.SUNKEN, anchor=tk.W)
        status_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Footer with version and author information
        footer_frame = ttk.Frame(main_frame)
        footer_frame.grid(row=9, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 3))
        footer_frame.columnconfigure(0, weight=1)
        
        footer_label = ttk.Label(footer_frame, 
                                text="V6.3.3925. Written in Python by Fadzli Abdullah. Huawei Technologies.",
                                font=('Ubuntu', 8), foreground='gray', anchor=tk.CENTER)
        footer_label.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # --- Auto-size window to content (compact) ---
        try:
            # Ensure all geometry is calculated
            self.root.update_idletasks()

            # Requested size of the window (content)
            req_w = self.root.winfo_reqwidth() + 20
            req_h = self.root.winfo_reqheight() + 20

            # Cap to screen size with small margin so it fits on 14" displays
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            max_w = screen_w - 80
            max_h = screen_h - 120

            new_w = min(req_w, max_w)
            new_h = min(req_h, max_h)

            # Don't make it too tight — enforce comfortable minimums
            new_w = max(new_w, 620)
            new_h = max(new_h, 520)

            # Center the resized window
            pos_x = int((screen_w - new_w) / 2)
            pos_y = int((screen_h - new_h) / 2)
            self.root.geometry(f"{new_w}x{new_h}+{pos_x}+{pos_y}")

            # Update sensible minimums so user can resize but not shrink too small
            self.root.minsize(620, 500)
        except Exception:
            pass
        
        # Set custom icon (if available) - MUST be called AFTER mapping_status is created
        self.set_custom_icon()
        
        # Auto-load MAPPING.csv from the same directory as the script
        self.auto_load_mapping()
    
    def set_custom_icon(self):
        """Set custom icon for the application window (title bar and taskbar on Windows)"""
        try:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            # Try app_icon.ico first, then fall back to icon.ico
            icon_path = os.path.join(base_path, 'app_icon.ico')
            if not os.path.exists(icon_path):
                icon_path = os.path.join(base_path, 'icon.ico')
            
            if os.path.exists(icon_path):
                # Use iconbitmap for Windows - applies to title bar and taskbar
                try:
                    self.root.iconbitmap(icon_path)
                except Exception as e:
                    print(f"Error setting icon: {e}")
            else:
                # Icon file not found - continue without custom icon
                pass
        except Exception as e:
            # Silently fail if icon loading has any issues
            pass
    
    def auto_load_mapping(self):
        """Automatically load MAPPING.csv from the same directory as the script"""
        try:
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                base_path = os.path.dirname(sys.executable)
            else:
                # Running as script
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            mapping_file = os.path.join(base_path, 'MAPPING.csv')
            
            if os.path.exists(mapping_file):
                self.load_mapping_from_file(mapping_file)
            else:
                # Try alternate location for development
                mapping_file_alt = os.path.join(os.getcwd(), 'MAPPING.csv')
                if os.path.exists(mapping_file_alt):
                    self.load_mapping_from_file(mapping_file_alt)
        except Exception as e:
            print(f"Could not auto-load MAPPING.csv: {str(e)}")
    
    def switch_converter_mode(self):
        """Switch between eNodeB name, decimal and sector ID converter modes"""
        mode = self.converter_mode.get()
        
        if mode == "enodebname":
            self.input_label.config(text="eNodeB Names (one per line or comma-separated):")
            self.converter_info.config(text="eNodeB Name mode (DEFAULT): Converts to 5-digit hex - includes ALL cells! Use Sector ID mode for specific cells.")
            self.load_mapping_btn.config(state='normal')
        elif mode == "decimal":
            self.input_label.config(text="Decimal Values (one per line or comma-separated):")
            self.converter_info.config(text="Decimal mode: Paste eNodeB_IDs (e.g., 518169, 520001) to convert to hex")
            self.load_mapping_btn.config(state='disabled')
        else:  # sectorid mode
            self.input_label.config(text="Sector IDs (one per line or comma-separated):")
            self.converter_info.config(text="Sector ID mode: AUTO-ADDS to ECI list! Format: eNodeB_hex(5) + Cell_hex(2) = 7-digit ECI")
            self.load_mapping_btn.config(state='normal')
        
        # Clear both input and output
        self.clear_converter()
    
    def sector_to_number(self, sector_str):
        """
        Convert sector identifier to number.
        - Numeric sectors (1, 2, 3, 31, 32): return as integer
        - Alphanumeric sectors (A, B, C, D, L, M): map to numbers (A=10, B=11, C=12, etc.)
        """
        sector_str = sector_str.strip().upper()
        
        # Try to convert directly to integer
        try:
            return int(sector_str)
        except ValueError:
            pass
        
        # For single letter, use alphabet position + 9 (A=10, B=11, C=12, etc.)
        if len(sector_str) == 1 and sector_str.isalpha():
            return ord(sector_str) - ord('A') + 10
        
        # For multiple characters, try to parse
        # Could be combinations like "1A", "2B", etc.
        # Extract numeric and alpha parts
        numeric_part = ''.join(c for c in sector_str if c.isdigit())
        alpha_part = ''.join(c for c in sector_str if c.isalpha())
        
        if numeric_part:
            return int(numeric_part)
        elif alpha_part and len(alpha_part) == 1:
            return ord(alpha_part) - ord('A') + 10
        
        # Default fallback
        return 0
    
    def load_cell_mapping(self):
        """Load cell name to eNodeB_ID mapping from file"""
        filename = filedialog.askopenfilename(
            title="Select Cell Mapping File",
            filetypes=[
                ("CSV files", "*.csv"),
                ("Text files", "*.txt"),
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        
        if not filename:
            return
        
        self.load_mapping_from_file(filename)
    
    def load_mapping_from_file(self, filename):
        """Load mapping from specified file"""
        try:
            import csv
            self.cell_mapping.clear()
            loaded_count = 0
            
            # Determine file type and load accordingly
            if filename.endswith('.csv') or filename.endswith('.txt'):
                with open(filename, 'r', encoding='utf-8-sig') as f:
                    # Try to detect delimiter
                    sample = f.read(1024)
                    f.seek(0)
                    
                    # Check for common delimiters
                    if '\t' in sample:
                        delimiter = '\t'
                    elif ',' in sample:
                        delimiter = ','
                    elif ';' in sample:
                        delimiter = ';'
                    else:
                        delimiter = ','
                    
                    reader = csv.reader(f, delimiter=delimiter)
                    
                    # Skip header if it looks like a header
                    first_row = next(reader, None)
                    if first_row:
                        # Check if first row is header
                        is_header = any(keyword in str(cell).lower() for cell in first_row 
                                      for keyword in ['sector', 'enodeb', 'cell', 'name', 'id', '4lrd', '5lrd', '6lrd'])
                        
                        if not is_header:
                            # Process first row if it's not a header
                            try:
                                loaded_count += self._process_mapping_row(first_row)
                            except (ValueError, IndexError):
                                pass
                    
                    # Process remaining rows
                    for row in reader:
                        if len(row) >= 2:
                            try:
                                loaded_count += self._process_mapping_row(row)
                            except (ValueError, IndexError):
                                continue
            
            elif filename.endswith('.xlsx'):
                # For Excel files, try to use pandas or openpyxl
                try:
                    import pandas as pd
                    df = pd.read_excel(filename)
                    
                    for _, row in df.iterrows():
                        try:
                            row_list = [str(cell).strip() for cell in row.tolist()]
                            loaded_count += self._process_mapping_row(row_list)
                        except (ValueError, IndexError):
                            continue
                except ImportError:
                    self.status_var.set("Error: pandas library not installed. Please use CSV format.")
                    return
            
            if loaded_count > 0:
                # Build eNodeB name mapping from cell_mapping
                self.build_enodeb_mapping()
                self.mapping_status.config(text=f"{loaded_count} mappings loaded", foreground='green')
                self.status_var.set(f"Successfully loaded {loaded_count} sector ID mappings ({len(self.enodeb_mapping)} unique eNodeBs) from {os.path.basename(filename)}")
            else:
                self.mapping_status.config(text="No valid mappings found", foreground='red')
                self.status_var.set("No valid mappings found in file")
        
        except Exception as e:
            self.mapping_status.config(text="Error loading file", foreground='red')
            self.status_var.set(f"Error loading mapping file: {str(e)}")
    
    def _process_mapping_row(self, row):
        """Process a single mapping row and add entries to cell_mapping and enodeb_mapping
        
        Supports multiple formats:
        - 5-column format: 4LRD, 5LRD, eNodeB Name, Sector ID, eNodeB ID
          - Column 2 (eNodeB Name) maps to eNodeB ID for eNodeB Name lookup
          - Column 3 (Sector ID) maps to eNodeB ID for Sector ID lookup
        - 2-column format: Sector ID, eNodeB ID
        
        Maps both Sector ID and eNodeB Name to eNodeB ID
        """
        count = 0
        
        try:
            if len(row) >= 5:
                # 5-column format: col0=4LRD, col1=5LRD, col2=eNodeB Name, col3=Sector ID, col4=eNodeB ID
                enodeb_name = row[2].strip().upper()
                sector_id = row[3].strip().upper()
                enodeb_id = int(row[4])
                
                # Add Sector ID mapping
                if sector_id and sector_id != 'NAN' and enodeb_id >= 0:
                    self.cell_mapping[sector_id] = enodeb_id
                    count += 1
                
                # Add eNodeB Name mapping (Column C)
                if enodeb_name and enodeb_name != 'NAN' and enodeb_id >= 0:
                    # Store in enodeb_mapping directly
                    if enodeb_name not in self.enodeb_mapping:
                        self.enodeb_mapping[enodeb_name] = enodeb_id
                        
            elif len(row) >= 2:
                # 2-column format: col0=Sector ID, col1=eNodeB ID
                sector_id = row[0].strip().upper()
                enodeb_id = int(row[1])
                
                if sector_id and sector_id != 'NAN' and enodeb_id >= 0:
                    self.cell_mapping[sector_id] = enodeb_id
                    count += 1
        
        except (ValueError, IndexError):
            pass
        
        return count
    
    def paste_converter_values(self):
        """Paste values from clipboard into the converter input area"""
        try:
            clipboard_text = self.root.clipboard_get()
            self.converter_input_text.delete(1.0, tk.END)
            self.converter_input_text.insert(1.0, clipboard_text)
            self.status_var.set("Values pasted from clipboard")
        except tk.TclError:
            self.status_var.set("Clipboard is empty or contains no text")
    
    def clear_converter(self):
        """Clear both converter input and hex result text areas"""
        self.converter_input_text.delete(1.0, tk.END)
        self.hex_result_text.config(state='normal')
        self.hex_result_text.delete(1.0, tk.END)
        self.hex_result_text.config(state='disabled')
        self.status_var.set("Converter cleared")
    
    def build_enodeb_mapping(self):
        """Build eNodeB name to ID mapping from cell_mapping.
        This is called after loading the mapping file to ensure we have both:
        1. Direct mappings from Column C (eNodeB Name) - already loaded in _process_mapping_row
        2. Fallback mappings from Sector ID prefix (for backward compatibility)"""
        
        # Add fallback mappings from Sector ID prefixes if eNodeB Name wasn't in the file
        for sector_id, enodeb_id in self.cell_mapping.items():
            # Extract eNodeB name (part before underscore) as fallback
            if '_' in sector_id:
                enodeb_name = sector_id.split('_')[0].strip().upper()
                if enodeb_name:
                    # Only add if not already in mapping (Column C takes precedence)
                    if enodeb_name not in self.enodeb_mapping:
                        self.enodeb_mapping[enodeb_name] = enodeb_id
    
    def convert_and_add_all(self):
        """Convert values based on current mode (eNodeB name, decimal or sector ID) and add to ECI list"""
        mode = self.converter_mode.get()
        
        if mode == "enodebname":
            self.convert_enodebname_bulk()
        elif mode == "decimal":
            self.convert_decimal_bulk()
        else:
            self.convert_sectorid_bulk()
    
    def convert_decimal_bulk(self):
        """Convert multiple decimal values to hexadecimal and add to ECI list"""
        input_text = self.converter_input_text.get(1.0, tk.END).strip()
        
        if not input_text:
            self.status_var.set("Please enter decimal values to convert")
            return
        
        # Split by common delimiters (comma, space, newline, tab)
        raw_values = re.split(r'[,\s\n\r\t]+', input_text)
        
        added = 0
        skipped = 0
        invalid = 0
        hex_results = []
        
        for value_str in raw_values:
            value_str = value_str.strip()
            if not value_str:
                continue
            
            try:
                # Convert to integer
                decimal_value = int(value_str)
                
                # Check valid range for 8-digit hex (0 to 268435455 = 0xFFFFFFF, 28-bit ECI)
                if decimal_value < 0 or decimal_value > 268435455:
                    invalid += 1
                    hex_results.append(f"{value_str} -> OUT OF RANGE (max: 268435455)")
                    continue
                
                # Convert to hexadecimal (5-8 digits, uppercase, no '0x' prefix)
                if decimal_value <= 1048575:  # 5 digits (0xFFFFF)
                    hex_value = format(decimal_value, '05X')
                elif decimal_value <= 16777215:  # 6 digits (0xFFFFFF)
                    hex_value = format(decimal_value, '06X')
                elif decimal_value <= 268435455:  # 7-8 digits (0xFFFFFFF)
                    hex_value = format(decimal_value, '07X')
                else:
                    hex_value = format(decimal_value, '08X')
                
                hex_results.append(f"{value_str} -> {hex_value}")
                
                # Check if already in list
                if hex_value in self.selected_ecis:
                    skipped += 1
                    continue
                
                # Add to ECI list
                self.selected_ecis.append(hex_value)
                added += 1
                
            except ValueError:
                invalid += 1
                hex_results.append(f"{value_str} -> INVALID")
        
        # Display results
        self.hex_result_text.config(state='normal')
        self.hex_result_text.delete(1.0, tk.END)
        self.hex_result_text.insert(1.0, "\n".join(hex_results))
        self.hex_result_text.config(state='disabled')
        
        # Update ECI display if any were added
        if added > 0:
            self.update_eci_display()
        
        # Build status message
        status_parts = []
        if added > 0:
            status_parts.append(f"Added {added} ECIs")
        if skipped > 0:
            status_parts.append(f"{skipped} duplicates")
        if invalid > 0:
            status_parts.append(f"{invalid} invalid")
        
        self.status_var.set(", ".join(status_parts) if status_parts else "No valid values found")
    
    def convert_sectorid_bulk(self):
        """Convert multiple Sector IDs to hexadecimal ECI and AUTO-ADD to ECI Selection.
        
        ECI Format: eNodeB_ID (5 hex digits) + Cell_Number (2 hex digits) = 7 hex digits
        Example: SNOTM_2 → eNodeB 260398 (3F92E) + Cell 2 (02) = 3F92E02
        
        This function automatically adds all successfully converted hexadecimal
        values to the ECI Selection list - no manual copying required!
        """
        if not self.cell_mapping:
            self.status_var.set("Please load MAPPING.csv first! Place it in the same directory as this script.")
            return
        
        input_text = self.converter_input_text.get(1.0, tk.END).strip()
        
        if not input_text:
            self.status_var.set("Please enter Sector IDs to convert")
            return
        
        # Split by common delimiters
        raw_values = re.split(r'[,\s\n\r\t]+', input_text)
        
        added = 0
        skipped = 0
        not_found = 0
        invalid_format = 0
        hex_results = []
        
        for sector_id in raw_values:
            sector_id = sector_id.strip().upper()
            if not sector_id:
                continue
            
            # Validate Sector ID format (must contain underscore)
            if '_' not in sector_id:
                invalid_format += 1
                hex_results.append(f"{sector_id} -> INVALID FORMAT (must contain '_', e.g., MEBUM_3)")
                continue
            
            # Look up Sector ID in mapping to get eNodeB ID
            if sector_id in self.cell_mapping:
                enodeb_id = self.cell_mapping[sector_id]
                
                # Extract sector number from Sector ID (format: XXXXX_Y where Y is sector number)
                # Examples: MEBUM_3, SNAVM_1, AKOIM_1
                parts = sector_id.split('_')
                if len(parts) >= 2:
                    sector_number_str = parts[-1]  # Get last part after underscore
                    sector_number = self.sector_to_number(sector_number_str)
                else:
                    # Should not reach here due to earlier validation
                    sector_number = 0
                
                # Validate eNodeB_ID range for 5-digit hex (0 to 1048575 = 0xFFFFF)
                if enodeb_id < 0 or enodeb_id > 1048575:
                    hex_results.append(f"{sector_id} -> eNodeB OUT OF RANGE (eNB:{enodeb_id}, max:1048575)")
                    skipped += 1
                    continue
                
                # Validate sector number for 2-digit hex (0 to 255 = 0xFF)
                if sector_number < 0 or sector_number > 255:
                    hex_results.append(f"{sector_id} -> CELL OUT OF RANGE (Cell:{sector_number}, max:255)")
                    skipped += 1
                    continue
                
                # Convert to proper ECI format: eNodeB_hex (5 digits) + Cell_hex (2 digits)
                enodeb_hex = format(enodeb_id, '05X')  # 5-digit hex for eNodeB ID
                cell_hex = format(sector_number, '02X')  # 2-digit hex for cell number
                hex_value = enodeb_hex + cell_hex  # 7-digit ECI
                
                hex_results.append(f"{sector_id} -> {hex_value} (eNB:{enodeb_id}={enodeb_hex}, Cell:{sector_number}={cell_hex})")
                
                # Check if already in list
                if hex_value in self.selected_ecis:
                    skipped += 1
                    continue
                
                # Automatically add to ECI list
                self.selected_ecis.append(hex_value)
                added += 1
            else:
                not_found += 1
                hex_results.append(f"{sector_id} -> NOT FOUND IN MAPPING")
        
        # Display results
        self.hex_result_text.config(state='normal')
        self.hex_result_text.delete(1.0, tk.END)
        self.hex_result_text.insert(1.0, "\n".join(hex_results))
        self.hex_result_text.config(state='disabled')
        
        # Automatically update ECI display if any were added
        if added > 0:
            self.update_eci_display()
        
        # Build status message with clear auto-add notification
        status_parts = []
        if added > 0:
            status_parts.append(f"✓ Auto-added {added} ECIs to selection")
        if skipped > 0:
            status_parts.append(f"{skipped} duplicates/out-of-range")
        if not_found > 0:
            status_parts.append(f"{not_found} not found in mapping")
        if invalid_format > 0:
            status_parts.append(f"{invalid_format} invalid format")
        
        self.status_var.set(", ".join(status_parts) if status_parts else "No valid Sector IDs found")
    
    def convert_enodebname_bulk(self):
        """Convert multiple eNodeB Names to hexadecimal and AUTO-ADD to ECI Selection.
        
        This function looks up eNodeB Names in the database, retrieves their eNodeB IDs,
        and converts them to 5-digit hex. This includes ALL cells under that eNodeB.
        
        ECI Format: eNodeB_ID (5 hex digits) - includes all cells
        Example: MC13ML → eNodeB 13 (0000D) - matches all cells (0000D01, 0000D02, 0000D03, etc.)
        
        For specific cells, use Sector ID mode instead.
        
        This function automatically adds all successfully converted hexadecimal
        values to the ECI Selection list - no manual copying required!
        """
        if not self.enodeb_mapping:
            self.status_var.set("Please load MAPPING.csv first! Place it in the same directory as this script.")
            return
        
        input_text = self.converter_input_text.get(1.0, tk.END).strip()
        
        if not input_text:
            self.status_var.set("Please enter eNodeB Names to convert")
            return
        
        # Split by common delimiters
        raw_values = re.split(r'[,\s\n\r\t]+', input_text)
        
        added = 0
        skipped = 0
        not_found = 0
        hex_results = []
        
        for enodeb_name in raw_values:
            enodeb_name = enodeb_name.strip().upper()
            if not enodeb_name:
                continue
            
            # Remove underscore and anything after it if present (in case user pastes sector IDs)
            if '_' in enodeb_name:
                enodeb_name = enodeb_name.split('_')[0]
            
            # Look up eNodeB Name in mapping to get eNodeB ID
            if enodeb_name in self.enodeb_mapping:
                enodeb_id = self.enodeb_mapping[enodeb_name]
                
                # Validate eNodeB_ID range for 5-digit hex (0 to 1048575 = 0xFFFFF)
                if enodeb_id < 0 or enodeb_id > 1048575:
                    hex_results.append(f"{enodeb_name} -> eNodeB OUT OF RANGE (eNB:{enodeb_id}, max:1048575)")
                    skipped += 1
                    continue
                
                # Convert to 5-digit hex for eNodeB ID (includes all cells)
                enodeb_hex = format(enodeb_id, '05X')
                
                # Check if already in list
                if enodeb_hex in self.selected_ecis:
                    hex_results.append(f"{enodeb_name} -> {enodeb_hex} (eNB:{enodeb_id}) [Already in list - includes ALL cells]")
                    skipped += 1
                    continue
                
                # Add only the 5-digit eNodeB hex (this includes all cells)
                self.selected_ecis.append(enodeb_hex)
                added += 1
                
                hex_results.append(f"{enodeb_name} -> {enodeb_hex} (eNB:{enodeb_id}) [Includes ALL cells under this eNodeB]")
            else:
                not_found += 1
                hex_results.append(f"{enodeb_name} -> NOT FOUND IN MAPPING")
        
        # Display results
        self.hex_result_text.config(state='normal')
        self.hex_result_text.delete(1.0, tk.END)
        self.hex_result_text.insert(1.0, "\n".join(hex_results))
        self.hex_result_text.config(state='disabled')
        
        # Automatically update ECI display if any were added
        if added > 0:
            self.update_eci_display()
        
        # Build status message with clear auto-add notification
        status_parts = []
        if added > 0:
            status_parts.append(f"✓ Auto-added {added} eNodeBs to selection (includes all cells)")
        if skipped > 0:
            status_parts.append(f"{skipped} duplicates/out-of-range")
        if not_found > 0:
            status_parts.append(f"{not_found} not found in mapping")
        
        self.status_var.set(", ".join(status_parts) if status_parts else "No valid eNodeB Names found")
    
    def add_eci(self):
        eci = self.eci_entry.get().strip().upper()
        
        # Validate ECI format (7-digit hexadecimal is standard, 5-8 supported)
        if not re.match(r'^[0-9A-F]{5,8}$', eci):
            self.status_var.set("Invalid ECI format. Must be 5-8 digit hexadecimal (standard: 7-digit, e.g., 3F92E02)")
            return
        
        if eci in self.selected_ecis:
            self.status_var.set(f"ECI {eci} already in list")
            return
        
        self.selected_ecis.append(eci)
        self.update_eci_display()
        self.eci_entry.delete(0, tk.END)
        self.status_var.set(f"Added ECI {eci}")
    
    def paste_bulk_eci(self):
        """Paste and process multiple ECIs from clipboard"""
        try:
            clipboard_text = self.root.clipboard_get()
            
            # Split by common delimiters and clean
            raw_ecis = re.split(r'[,\s\n\r\t]+', clipboard_text)
            
            added = 0
            skipped = 0
            invalid = 0
            
            for eci in raw_ecis:
                eci = eci.strip().upper()
                if not eci:
                    continue
                
                # Validate format (5-8 digit hexadecimal)
                if not re.match(r'^[0-9A-F]{5,8}$', eci):
                    invalid += 1
                    continue
                
                if eci in self.selected_ecis:
                    skipped += 1
                    continue
                
                self.selected_ecis.append(eci)
                added += 1
            
            self.update_eci_display()
            
            status_parts = []
            if added > 0:
                status_parts.append(f"Added {added} ECIs")
            if skipped > 0:
                status_parts.append(f"{skipped} duplicates skipped")
            if invalid > 0:
                status_parts.append(f"{invalid} invalid entries")
            
            self.status_var.set(", ".join(status_parts) if status_parts else "No valid ECIs found in clipboard")
            
        except tk.TclError:
            self.status_var.set("Clipboard is empty or contains no text")
    
    def clear_ecis(self):
        self.selected_ecis.clear()
        self.update_eci_display()
        self.status_var.set("All ECIs cleared")
    
    def update_eci_display(self):
        """Update the ECI text display with current ECIs"""
        self.eci_text.delete(1.0, tk.END)
        
        if self.selected_ecis:
            # Display ECIs in a clean, comma-separated format with wrapping
            display_text = ", ".join(self.selected_ecis)
            self.eci_text.insert(1.0, display_text)
        else:
            self.eci_text.insert(1.0, "No ECIs selected")
    
    def quick_select_days(self, days):
        """Quickly select date range based on number of days (excludes P0/today)"""
        reference = self.reference_date.get_date()
        end = reference - timedelta(days=1)  # Yesterday (P1)
        start = reference - timedelta(days=days)  # For 3 days: P3, 7 days: P7, etc.
        
        self.start_date.set_date(start)
        self.end_date.set_date(end)
        
        self.calculate_partitions(None)
    
    def calculate_partitions(self, event):
        """Calculate partition numbers based on selected dates"""
        try:
            reference = self.reference_date.get_date()
            start = self.start_date.get_date()
            end = self.end_date.get_date()
            
            if start > end:
                self.partition_var.set("Error: Start date must be before end date")
                return
            
            # Calculate partition numbers
            partitions = []
            current = start
            while current <= end:
                days_diff = (reference - current).days
                partitions.append(f"p{days_diff}")
                current += timedelta(days=1)
            
            self.partition_var.set(", ".join(partitions))
            
        except Exception as e:
            self.partition_var.set(f"Error: {str(e)}")
    
    def toggle_all_apps(self):
        """Toggle all app selections"""
        select_all = self.select_all_var.get()
        for var in self.app_vars.values():
            var.set(select_all)
    
    def check_all_apps_selected(self, *args):
        """Check if all apps are selected and update Select All checkbox"""
        all_selected = all(var.get() for var in self.app_vars.values())
        self.select_all_var.set(all_selected)
    
    def generate_query(self):
        if not self.selected_ecis:
            self.status_var.set("Please add at least one ECI")
            return
        
        # Check if at least one app is selected
        selected_apps = [app_id for app_id, var in self.app_vars.items() if var.get()]
        if not selected_apps:
            self.status_var.set("Please select at least one application")
            return
        
        partitions = self.partition_var.get()
        if not partitions or partitions.startswith("Error"):
            self.status_var.set("Please select valid dates")
            return
        
        rat = self.rat_var.get()
        app_ids = ", ".join(selected_apps)
        eci_list = "', '".join(self.selected_ecis)
        
        start_date = self.start_date.get_date().strftime('%Y-%m-%d')
        end_date = self.end_date.get_date().strftime('%Y-%m-%d')
        
        # Check if Resolution column should be included
        include_resolution = self.include_resolution_var.get()
        
        # Build the query based on whether resolution is included
        if include_resolution:
            # Query WITH max_video_data_rate and Resolution
            query = f"""-- Streaming Data Query (WITH Resolution)
-- Date Range: {start_date} to {end_date}
-- ECIs: {len(self.selected_ecis)} selected
-- Apps: {', '.join([self.apps[app_id] for app_id in selected_apps])}
-- Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

WITH lvl0 AS (
  SELECT
    from_unixtime(a.begin_time, 'yyyy-MM-dd') AS date,
    a.imsi,
    a.eci,
    substr(a.eci, 1, 5) AS eci_prefix,
    a.app_id,
    MAX(COALESCE(a.video_data_rate, 0)) AS max_video_data_rate,
    SUM(CASE
          WHEN (((a.PLAY_STATE = 0) OR (a.PLAY_STATE = 1)) OR a.ENCRYPTED_MODEL_FLAG = 1)
               AND a.STREAMING_DW_PACKETS >= 600
          THEN a.STREAMING_DW_PACKETS ELSE 0 END) AS Video_Streaming_Download_Throughput_nom,
    SUM(CASE
          WHEN (((a.PLAY_STATE = 0) OR (a.PLAY_STATE = 1)) OR a.ENCRYPTED_MODEL_FLAG = 1)
               AND a.STREAMING_DW_PACKETS >= 600
          THEN a.STREAMING_DOWNLOAD_DELAY ELSE 0 END) AS Video_Streaming_Throughput_denom,
    SUM(CASE
          WHEN a.VIDEO_START_FLAG = 0 AND a.VIDEO_START_IDLE_DELAY IS NOT NULL
            THEN (a.VIDEO_START_DELAY - a.VIDEO_START_IDLE_DELAY)
          WHEN a.VIDEO_START_FLAG = 0 AND a.VIDEO_START_DELAY IS NOT NULL
            THEN a.VIDEO_START_DELAY
          ELSE 0 END) AS video_xkb_start_delay_nom,
    SUM(CASE WHEN a.VIDEO_START_FLAG = 0 THEN 1 ELSE 0 END) AS video_xkb_start_delay_denom,
    SUM(CASE
          WHEN (a.play_duration > 0 AND a.imsi <> '' AND a.IPMOS_FLAG IN (0, 2)
                AND a.SERVICE_VALID_FLAG = 1 AND a.VIDEO_START_FLAG = 0)
          THEN a.stall_duration ELSE 0 END) AS stall_duration_ms,
    SUM(CASE
          WHEN (a.play_duration > 0 AND a.imsi <> '' AND a.IPMOS_FLAG IN (0, 2)
                AND a.SERVICE_VALID_FLAG = 1 AND a.VIDEO_START_FLAG = 0)
          THEN a.play_duration ELSE 0 END) AS play_duration_ms,
    SUM(a.L4_UL_THROUGHPUT) AS ul_thruput_byte,
    SUM(a.L4_DW_THROUGHPUT) AS dl_thruput_byte,
    SUM(a.L4_DW_THROUGHPUT) AS dl_throughput_num,
    SUM(a.DATATRANS_DW_TOTAL_DURATION) AS dl_throughput_denom
  FROM xdr.detail_ufdr_streaming PARTITION ({partitions}) a
  WHERE a.rat IN ({rat})
    AND a.app_id IN ({app_ids})
    AND a.eci IN ('{eci_list}')
  GROUP BY from_unixtime(a.begin_time, 'yyyy-MM-dd'), a.imsi, a.eci, a.app_id
),

lvl1 AS (
  SELECT
    a.date,
    a.imsi,
    a.eci,
    a.eci_prefix,
    a.app_id,
    MAX(a.max_video_data_rate) AS max_video_data_rate,
    ((SUM(a.ul_thruput_byte) + SUM(a.dl_thruput_byte)) / 1024.0) AS totalTraffic_kb,
    ((SUM(a.Video_Streaming_Download_Throughput_nom) * 8.0)
      / NULLIF(SUM(a.Video_Streaming_Throughput_denom), 0)) / 1024.0 AS vid_stream_dwld_thru_kbps,
    (SUM(a.video_xkb_start_delay_nom)
      / NULLIF(SUM(a.video_xkb_start_delay_denom), 0)) AS video_xkb_start_delay_ms,
    SUM(a.stall_duration_ms) AS stall_duration_ms,
    SUM(a.play_duration_ms) AS play_duration_ms,
    ((SUM(a.dl_throughput_num) * 8.0) / NULLIF(SUM(a.dl_throughput_denom), 0)) AS dl_throughput_kbps,
    SUM(a.dl_throughput_num) AS dl_throughput_num,
    SUM(a.dl_throughput_denom) AS dl_throughput_denom,
    SUM(a.video_xkb_start_delay_nom) AS video_start_delay_num,
    SUM(a.video_xkb_start_delay_denom) AS video_start_delay_denom
  FROM lvl0 a
  GROUP BY a.date, a.imsi, a.eci, a.eci_prefix, a.app_id
),

lvl2 AS (
  SELECT
    a.date,
    a.imsi,
    a.eci,
    a.eci_prefix,
    a.app_id,
    MAX(a.max_video_data_rate) AS max_video_data_rate,
    SUM(a.totalTraffic_kb) AS totalTraffic_kb,
    AVG(a.vid_stream_dwld_thru_kbps) AS vid_stream_dwld_thru_kbps,
    AVG(a.video_xkb_start_delay_ms) AS video_xkb_start_delay_ms,
    SUM(a.stall_duration_ms) AS stall_duration_ms,
    SUM(a.play_duration_ms) AS play_duration_ms,
    AVG(a.dl_throughput_kbps) AS dl_throughput_kbps,
    SUM(a.dl_throughput_num) AS dl_throughput_num,
    SUM(a.dl_throughput_denom) AS dl_throughput_denom,
    SUM(a.video_start_delay_num) AS video_start_delay_num,
    SUM(a.video_start_delay_denom) AS video_start_delay_denom
  FROM lvl1 a
  GROUP BY a.date, a.imsi, a.eci, a.eci_prefix, a.app_id
),

final_calc AS (
  SELECT
    x.*,
    /* Convert HEX prefix (5 digits) to DECIMAL eNodeB_ID */
    (
      (ascii(upper(substr(x.eci_prefix,1,1))) - CASE WHEN upper(substr(x.eci_prefix,1,1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END) * 65536 +
      (ascii(upper(substr(x.eci_prefix,2,1))) - CASE WHEN upper(substr(x.eci_prefix,2,1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END) * 4096 +
      (ascii(upper(substr(x.eci_prefix,3,1))) - CASE WHEN upper(substr(x.eci_prefix,3,1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END) * 256 +
      (ascii(upper(substr(x.eci_prefix,4,1))) - CASE WHEN upper(substr(x.eci_prefix,4,1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END) * 16 +
      (ascii(upper(substr(x.eci_prefix,5,1))) - CASE WHEN upper(substr(x.eci_prefix,5,1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END)
    ) AS eNodeB_ID,

    /* Convert HEX last 2 digits to DECIMAL Cell_Dec */
    (
      (ascii(upper(substr(x.eci, length(x.eci)-1, 1))) - CASE WHEN upper(substr(x.eci, length(x.eci)-1, 1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END) * 16 +
      (ascii(upper(substr(x.eci, length(x.eci), 1))) - CASE WHEN upper(substr(x.eci, length(x.eci), 1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END)
    ) AS Cell_Dec,
    
    /* Calculate Video Resolution based on max_video_data_rate */
    CASE 
      WHEN x.max_video_data_rate >= 0 AND x.max_video_data_rate < 300 THEN '240P'
      WHEN x.max_video_data_rate >= 300 AND x.max_video_data_rate < 500 THEN '360P'
      WHEN x.max_video_data_rate >= 500 AND x.max_video_data_rate < 1024 THEN '480P'
      WHEN x.max_video_data_rate >= 1024 AND x.max_video_data_rate < 2048 THEN '720P'
      WHEN x.max_video_data_rate >= 2048 AND x.max_video_data_rate < 4096 THEN '1080P'
      WHEN x.max_video_data_rate >= 4096 AND x.max_video_data_rate < 9000 THEN '2K'
      WHEN x.max_video_data_rate >= 9000 THEN '4K'
      ELSE 'UNKNOWN' 
    END AS Resolution,
    
    /* App name lookup */
    CASE x.app_id
      WHEN 342  THEN 'YouTube'
      WHEN 829  THEN 'Facebook'
      WHEN 1181 THEN 'Instagram'
      WHEN 4860 THEN 'TikTok'
      ELSE 'Unknown'
    END AS App_Name
  FROM lvl2 x
)

SELECT
  date,
  imsi,
  eci,
  eci_prefix,
  eNodeB_ID,
  concat(cast(eNodeB_ID AS string), '_', cast(Cell_Dec AS string)) AS Cell_ID,
  app_id,
  App_Name,
  totalTraffic_kb,
  vid_stream_dwld_thru_kbps,
  video_xkb_start_delay_ms,
  stall_duration_ms,
  play_duration_ms,
  dl_throughput_kbps,
  dl_throughput_num,
  dl_throughput_denom,
  video_start_delay_num,
  video_start_delay_denom,
  max_video_data_rate,
  Resolution
FROM final_calc;"""
        else:
            # Query WITHOUT video_data_rate and Resolution
            query = f"""-- Streaming Data Query (WITHOUT Resolution)
-- Date Range: {start_date} to {end_date}
-- ECIs: {len(self.selected_ecis)} selected
-- Apps: {', '.join([self.apps[app_id] for app_id in selected_apps])}
-- Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

WITH lvl0 AS (
  SELECT
    from_unixtime(a.begin_time, 'yyyy-MM-dd') AS date,
    a.imsi,
    a.eci,
    substr(a.eci, 1, 5) AS eci_prefix,
    a.app_id,
    SUM(CASE
          WHEN (((a.PLAY_STATE = 0) OR (a.PLAY_STATE = 1)) OR a.ENCRYPTED_MODEL_FLAG = 1)
               AND a.STREAMING_DW_PACKETS >= 600
          THEN a.STREAMING_DW_PACKETS ELSE 0 END) AS Video_Streaming_Download_Throughput_nom,
    SUM(CASE
          WHEN (((a.PLAY_STATE = 0) OR (a.PLAY_STATE = 1)) OR a.ENCRYPTED_MODEL_FLAG = 1)
               AND a.STREAMING_DW_PACKETS >= 600
          THEN a.STREAMING_DOWNLOAD_DELAY ELSE 0 END) AS Video_Streaming_Throughput_denom,
    SUM(CASE
          WHEN a.VIDEO_START_FLAG = 0 AND a.VIDEO_START_IDLE_DELAY IS NOT NULL
            THEN (a.VIDEO_START_DELAY - a.VIDEO_START_IDLE_DELAY)
          WHEN a.VIDEO_START_FLAG = 0 AND a.VIDEO_START_DELAY IS NOT NULL
            THEN a.VIDEO_START_DELAY
          ELSE 0 END) AS video_xkb_start_delay_nom,
    SUM(CASE WHEN a.VIDEO_START_FLAG = 0 THEN 1 ELSE 0 END) AS video_xkb_start_delay_denom,
    SUM(CASE
          WHEN (a.play_duration > 0 AND a.imsi <> '' AND a.IPMOS_FLAG IN (0, 2)
                AND a.SERVICE_VALID_FLAG = 1 AND a.VIDEO_START_FLAG = 0)
          THEN a.stall_duration ELSE 0 END) AS stall_duration_ms,
    SUM(CASE
          WHEN (a.play_duration > 0 AND a.imsi <> '' AND a.IPMOS_FLAG IN (0, 2)
                AND a.SERVICE_VALID_FLAG = 1 AND a.VIDEO_START_FLAG = 0)
          THEN a.play_duration ELSE 0 END) AS play_duration_ms,
    SUM(a.L4_UL_THROUGHPUT) AS ul_thruput_byte,
    SUM(a.L4_DW_THROUGHPUT) AS dl_thruput_byte,
    SUM(a.L4_DW_THROUGHPUT) AS dl_throughput_num,
    SUM(a.DATATRANS_DW_TOTAL_DURATION) AS dl_throughput_denom
  FROM xdr.detail_ufdr_streaming PARTITION ({partitions}) a
  WHERE a.rat IN ({rat})
    AND a.app_id IN ({app_ids})
    AND a.eci IN ('{eci_list}')
  GROUP BY from_unixtime(a.begin_time, 'yyyy-MM-dd'), a.imsi, a.eci, a.app_id
),

lvl1 AS (
  SELECT
    a.date,
    a.imsi,
    a.eci,
    a.eci_prefix,
    a.app_id,
    ((SUM(a.ul_thruput_byte) + SUM(a.dl_thruput_byte)) / 1024.0) AS totalTraffic_kb,
    ((SUM(a.Video_Streaming_Download_Throughput_nom) * 8.0)
      / NULLIF(SUM(a.Video_Streaming_Throughput_denom), 0)) / 1024.0 AS vid_stream_dwld_thru_kbps,
    (SUM(a.video_xkb_start_delay_nom)
      / NULLIF(SUM(a.video_xkb_start_delay_denom), 0)) AS video_xkb_start_delay_ms,
    SUM(a.stall_duration_ms) AS stall_duration_ms,
    SUM(a.play_duration_ms) AS play_duration_ms,
    ((SUM(a.dl_throughput_num) * 8.0) / NULLIF(SUM(a.dl_throughput_denom), 0)) AS dl_throughput_kbps,
    SUM(a.dl_throughput_num) AS dl_throughput_num,
    SUM(a.dl_throughput_denom) AS dl_throughput_denom,
    SUM(a.video_xkb_start_delay_nom) AS video_start_delay_num,
    SUM(a.video_xkb_start_delay_denom) AS video_start_delay_denom
  FROM lvl0 a
  GROUP BY a.date, a.imsi, a.eci, a.eci_prefix, a.app_id
),

lvl2 AS (
  SELECT
    a.date,
    a.imsi,
    a.eci,
    a.eci_prefix,
    a.app_id,
    SUM(a.totalTraffic_kb) AS totalTraffic_kb,
    AVG(a.vid_stream_dwld_thru_kbps) AS vid_stream_dwld_thru_kbps,
    AVG(a.video_xkb_start_delay_ms) AS video_xkb_start_delay_ms,
    SUM(a.stall_duration_ms) AS stall_duration_ms,
    SUM(a.play_duration_ms) AS play_duration_ms,
    AVG(a.dl_throughput_kbps) AS dl_throughput_kbps,
    SUM(a.dl_throughput_num) AS dl_throughput_num,
    SUM(a.dl_throughput_denom) AS dl_throughput_denom,
    SUM(a.video_start_delay_num) AS video_start_delay_num,
    SUM(a.video_start_delay_denom) AS video_start_delay_denom
  FROM lvl1 a
  GROUP BY a.date, a.imsi, a.eci, a.eci_prefix, a.app_id
),

final_calc AS (
  SELECT
    x.*,
    /* Convert HEX prefix (5 digits) to DECIMAL eNodeB_ID */
    (
      (ascii(upper(substr(x.eci_prefix,1,1))) - CASE WHEN upper(substr(x.eci_prefix,1,1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END) * 65536 +
      (ascii(upper(substr(x.eci_prefix,2,1))) - CASE WHEN upper(substr(x.eci_prefix,2,1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END) * 4096 +
      (ascii(upper(substr(x.eci_prefix,3,1))) - CASE WHEN upper(substr(x.eci_prefix,3,1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END) * 256 +
      (ascii(upper(substr(x.eci_prefix,4,1))) - CASE WHEN upper(substr(x.eci_prefix,4,1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END) * 16 +
      (ascii(upper(substr(x.eci_prefix,5,1))) - CASE WHEN upper(substr(x.eci_prefix,5,1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END)
    ) AS eNodeB_ID,

    /* Convert HEX last 2 digits to DECIMAL Cell_Dec */
    (
      (ascii(upper(substr(x.eci, length(x.eci)-1, 1))) - CASE WHEN upper(substr(x.eci, length(x.eci)-1, 1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END) * 16 +
      (ascii(upper(substr(x.eci, length(x.eci), 1))) - CASE WHEN upper(substr(x.eci, length(x.eci), 1)) BETWEEN 'A' AND 'F' THEN 55 ELSE 48 END)
    ) AS Cell_Dec,
    
    /* App name lookup */
    CASE x.app_id
      WHEN 342  THEN 'YouTube'
      WHEN 829  THEN 'Facebook'
      WHEN 1181 THEN 'Instagram'
      WHEN 4860 THEN 'TikTok'
      ELSE 'Unknown'
    END AS App_Name
  FROM lvl2 x
)

SELECT
  date,
  imsi,
  eci,
  eci_prefix,
  eNodeB_ID,
  concat(cast(eNodeB_ID AS string), '_', cast(Cell_Dec AS string)) AS Cell_ID,
  app_id,
  App_Name,
  totalTraffic_kb,
  vid_stream_dwld_thru_kbps,
  video_xkb_start_delay_ms,
  stall_duration_ms,
  play_duration_ms,
  dl_throughput_kbps,
  dl_throughput_num,
  dl_throughput_denom,
  video_start_delay_num,
  video_start_delay_denom
FROM final_calc;"""
        
        self.query_text.delete(1.0, tk.END)
        self.query_text.insert(1.0, query)
        
        # Get selected app names for status
        selected_apps = [self.apps[app_id] for app_id, var in self.app_vars.items() if var.get()]
        apps_str = ', '.join(selected_apps)
        
        resolution_status = " (with Resolution)" if include_resolution else " (without Resolution)"
        self.status_var.set(f"Query generated{resolution_status} for {len(self.selected_ecis)} ECIs, {len(selected_apps)} app(s) ({apps_str}), dates {start_date} to {end_date}")
    
    def copy_to_clipboard(self):
        query = self.query_text.get(1.0, tk.END).strip()
        if not query:
            self.status_var.set("Please generate a query first")
            return
        
        try:
            pyperclip.copy(query)
            self.status_var.set("Query copied to clipboard successfully")
        except Exception as e:
            self.status_var.set(f"Failed to copy to clipboard: {str(e)}")
    
    def save_to_file(self):
        query = self.query_text.get(1.0, tk.END).strip()
        if not query:
            self.status_var.set("Please generate a query first")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".sql",
            filetypes=[("SQL files", "*.sql"), ("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"streaming_query_{datetime.now().strftime('%Y%m%d_%H%M%S')}.sql"
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    f.write(query)
                self.status_var.set(f"Query saved to {filename}")
            except Exception as e:
                self.status_var.set(f"Failed to save file: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = StreamingQueryGenerator(root)
    root.mainloop()