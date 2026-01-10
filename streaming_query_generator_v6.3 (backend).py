"""
SQL Query Generator for Streaming Data Analysis - BACKEND VERSION
Created by: Fadzli Abdullah

KEY FEATURES:
1. AUTO-LOAD: Automatically loads MAPPING.csv from the same directory
2. eNodeB NAME CONVERTER (DEFAULT): Converts eNodeB Names to full 7-digit ECI using database lookup
3. SECTOR ID CONVERTER: Converts Sector IDs to full 7-digit ECI
4. TRIPLE CONVERTER: Supports eNodeB Name-to-Hex, Decimal-to-Hex, and Sector ID-to-Hex conversion
5. FLEXIBLE HEX: Supports 5-8 digit hexadecimal values

BACKEND-SPECIFIC:
- Generates queries using UNION ALL instead of PARTITION clause
- Uses ps.detail_ufdr_streaming_XXXXX table format
- Auto-calculates partition numbers from date range
- Partition formula: 20395 = Nov 3, 2025 (base date)

Version: 6.3.3925 (Backend)
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
from datetime import datetime, timedelta, date
from tkcalendar import DateEntry
import pyperclip
import re
import os
import sys

class ColoredButton(tk.Canvas):
    """Custom button widget with custom colors"""
    def __init__(self, parent, text, command, bg_color='#006400', fg_color='white', **kwargs):
        super().__init__(parent, height=26, highlightthickness=0, **kwargs)
        
        self.bg_color = bg_color
        self.hover_color = '#004d00'
        self.fg_color = fg_color
        self.command = command
        self.text = text
        
        self.button_width = len(text) * 8 + 20
        self.config(width=self.button_width)
        
        self.draw_button(self.bg_color)
        
        self.bind('<Button-1>', self.on_click)
        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)
        
    def draw_button(self, color):
        self.delete('all')
        self.create_rectangle(1, 1, self.button_width-1, 25, 
                            fill=color, outline='#808080', width=1,
                            tags='button_bg')
        self.create_text(self.button_width//2, 13, text=self.text, 
                        fill=self.fg_color, font=('TkDefaultFont', 9),
                        tags='button_text')
        
    def on_click(self, event):
        if self.command:
            self.command()
    
    def on_enter(self, event):
        self.draw_button(self.hover_color)
        self.config(cursor='hand2')
    
    def on_leave(self, event):
        self.draw_button(self.bg_color)
        self.config(cursor='')

class StreamingQueryGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("SQL Query Generator (BACKEND) - Fadzli Abdullah")
        
        # Set window size
        window_width = 600
        window_height = 750

        try:
            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()
            center_x = int((screen_width - window_width) / 2)
            center_y = int((screen_height - window_height) / 2)
            self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        except Exception:
            self.root.geometry(f"{window_width}x{window_height}")

        self.root.minsize(window_width, 700)
        self.root.resizable(True, True)

        # Store selected ECIs
        self.selected_ecis = []
        
        # Mapping dictionaries
        self.cell_mapping = {}
        self.enodeb_mapping = {}
        
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
        
        # Partition display variable
        self.partition_var = tk.StringVar(value="")
        
        # Create main container with minimal padding
        main_frame = ttk.Frame(root, padding="3")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)  # Query text area
               
        # ECI Selection Section - COMPACT
        eci_frame = ttk.LabelFrame(main_frame, text="ECI Selection", padding="3")
        eci_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=2)
        eci_frame.columnconfigure(5, weight=1)
        
        ttk.Label(eci_frame, text="Enter ECI (7-digit hex):").grid(row=0, column=0, sticky=tk.W, padx=2)
        self.eci_entry = ttk.Entry(eci_frame, width=15)
        self.eci_entry.grid(row=0, column=1, padx=2)
        self.eci_entry.bind('<Return>', lambda e: self.add_eci())
        
        ttk.Button(eci_frame, text="Add ECI", command=self.add_eci).grid(row=0, column=2, padx=2)
        self.paste_bulk_btn = ColoredButton(eci_frame, "Paste Clipboard", self.paste_bulk_eci)
        self.paste_bulk_btn.grid(row=0, column=3, padx=2)
        ttk.Button(eci_frame, text="Clear All", command=self.clear_ecis).grid(row=0, column=4, padx=2)
        
        # ECI Text Display - increased height for better visibility
        self.eci_text = tk.Text(eci_frame, height=6, wrap=tk.WORD, font=('Courier', 9))
        eci_scrollbar = ttk.Scrollbar(eci_frame, orient=tk.VERTICAL, command=self.eci_text.yview)
        self.eci_text.configure(yscrollcommand=eci_scrollbar.set)
        self.eci_text.grid(row=1, column=0, columnspan=5, sticky=(tk.W, tk.E), pady=2)
        eci_scrollbar.grid(row=1, column=5, sticky=(tk.N, tk.S))
        
        # Converter Section - TWO COLUMNS LAYOUT
        converter_frame = ttk.LabelFrame(main_frame, text="Converter: eNodeB Name → Hex | Decimal ↔ Hex | Sector ID → Hex", padding="3")
        converter_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=2)
        converter_frame.columnconfigure(0, weight=1)
        converter_frame.columnconfigure(1, weight=1)
        
        # Mode selection - horizontal
        mode_frame = ttk.Frame(converter_frame)
        mode_frame.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        self.converter_mode = tk.StringVar(value="name")
        ttk.Radiobutton(mode_frame, text="eNodeB Name to Hex", variable=self.converter_mode, 
                       value="name", command=self.on_converter_mode_change).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="Decimal to Hex", variable=self.converter_mode, 
                       value="decimal", command=self.on_converter_mode_change).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="Sector ID to Hex", variable=self.converter_mode, 
                       value="sector", command=self.on_converter_mode_change).pack(side=tk.LEFT, padx=5)
        ttk.Button(mode_frame, text="Reload Mapping", command=self.load_mapping_file).pack(side=tk.LEFT, padx=5)
        
        self.mapping_status_label = ttk.Label(mode_frame, text="", foreground='green')
        self.mapping_status_label.pack(side=tk.LEFT, padx=5)
        
        # LEFT COLUMN - Input
        left_frame = ttk.Frame(converter_frame)
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=2)
        
        self.converter_label = ttk.Label(left_frame, text="eNodeB Names (one per line or comma-separated):")
        self.converter_label.pack(anchor=tk.W)
        
        self.converter_input = scrolledtext.ScrolledText(left_frame, height=4, wrap=tk.WORD)
        self.converter_input.pack(fill=tk.BOTH, expand=True)
        
        # RIGHT COLUMN - Results
        right_frame = ttk.Frame(converter_frame)
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=2)
        
        ttk.Label(right_frame, text="Hexadecimal Results (5-8 digit):").pack(anchor=tk.W)
        
        self.converter_result = scrolledtext.ScrolledText(right_frame, height=4, wrap=tk.WORD)
        self.converter_result.pack(fill=tk.BOTH, expand=True)
        self.converter_result.config(state=tk.DISABLED)
        
        # Buttons below - horizontal
        button_frame = ttk.Frame(converter_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=2)
        
        self.convert_add_btn = ColoredButton(button_frame, "Convert & Add All", self.convert_and_add_all)
        self.convert_add_btn.pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Clear Converter", command=self.clear_converter_input).pack(side=tk.LEFT, padx=2)
        self.paste_converter_btn = ColoredButton(button_frame, "Paste Clipboard", self.paste_to_converter)
        self.paste_converter_btn.pack(side=tk.LEFT, padx=2)
        
        # Help text
        help_label = ttk.Label(converter_frame, 
                              text="eNodeB Name mode (DEFAULT): Converts to 5-digit hex - includes ALL cells! Use Sector ID mode for specific cells.",
                              foreground='gray', font=('TkDefaultFont', 8))
        help_label.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # Date Selection & Partition - COMPACT HORIZONTAL
        date_frame = ttk.LabelFrame(main_frame, text="Date Selection & Partition", padding="3")
        date_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=2)
        
        # Row 0: Quick selection buttons
        quick_frame = ttk.Frame(date_frame)
        quick_frame.grid(row=0, column=0, columnspan=6, sticky=tk.W, pady=2)
        
        ttk.Label(quick_frame, text="Quick Selection:").pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_frame, text="3 Days", width=8, command=lambda: self.quick_date_select(3)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_frame, text="7 Days", width=8, command=lambda: self.quick_date_select(7)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_frame, text="14 Days", width=8, command=lambda: self.quick_date_select(14)).pack(side=tk.LEFT, padx=2)
        
        # Row 1: Reference date, Start date, End date - all horizontal
        ttk.Label(date_frame, text="Reference Date (Today):").grid(row=1, column=0, sticky=tk.W, padx=2)
        self.reference_date = DateEntry(date_frame, width=10, background='darkblue',
                                       foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.reference_date.grid(row=1, column=1, padx=2)
        self.reference_date.bind('<<DateEntrySelected>>', lambda e: self.update_partition_display())
        
        ref_help = ttk.Label(date_frame, text="(p0=T (today), p1=T-1, p2=T-2, p3=T-3, etc)", 
                            foreground='gray', font=('TkDefaultFont', 8))
        ref_help.grid(row=1, column=2, sticky=tk.W, padx=2)
        
        ttk.Label(date_frame, text="Start Date:").grid(row=2, column=0, sticky=tk.W, padx=2)
        self.start_date = DateEntry(date_frame, width=10, background='darkblue',
                                    foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.start_date.grid(row=2, column=1, padx=2)
        self.start_date.bind('<<DateEntrySelected>>', lambda e: self.update_partition_display())
        
        ttk.Label(date_frame, text="End Date:").grid(row=2, column=2, sticky=tk.W, padx=2)
        self.end_date = DateEntry(date_frame, width=10, background='darkblue',
                                  foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.end_date.grid(row=2, column=3, padx=2)
        self.end_date.bind('<<DateEntrySelected>>', lambda e: self.update_partition_display())
        
        # Partitions display
        ttk.Label(date_frame, text="Partitions:").grid(row=3, column=0, sticky=tk.W, padx=2)
        partition_entry = ttk.Entry(date_frame, textvariable=self.partition_var, width=60, state='readonly')
        partition_entry.grid(row=3, column=1, columnspan=5, sticky=(tk.W, tk.E), padx=2)
        
        # Set default dates
        today = date.today()
        self.reference_date.set_date(today)
        self.end_date.set_date(today)
        self.start_date.set_date(today - timedelta(days=6))
        self.update_partition_display()
        
        # RAT Selection & Resolution - HORIZONTAL COMPACT
        rat_res_frame = ttk.LabelFrame(main_frame, text="RAT Selection", padding="3")
        rat_res_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=2)
        
        self.rat_var = tk.StringVar(value="6")
        ttk.Radiobutton(rat_res_frame, text="LTE (6)", variable=self.rat_var, value="6").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(rat_res_frame, text="5G NR (9)", variable=self.rat_var, value="9").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(rat_res_frame, text="Both (6,9)", variable=self.rat_var, value="6,9").pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(rat_res_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        self.resolution_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(rat_res_frame, text="Include Resolution (uses video_data_rate column)", 
                       variable=self.resolution_var).pack(side=tk.LEFT, padx=5)
        
        # App Selection - HORIZONTAL SINGLE ROW
        app_frame = ttk.LabelFrame(main_frame, text="Application Selection", padding="3")
        app_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=2)
        
        ttk.Checkbutton(app_frame, text="Select All", variable=self.select_all_var, 
                       command=self.toggle_select_all).pack(side=tk.LEFT, padx=5)
        
        for app_id, app_name in self.apps.items():
            var = tk.BooleanVar(value=False)
            self.app_vars[app_id] = var
            ttk.Checkbutton(app_frame, text=f"{app_name} ({app_id})", 
                          variable=var, command=self.update_select_all).pack(side=tk.LEFT, padx=5)
        
        # Query Generation - compact 3 lines
        query_frame = ttk.LabelFrame(main_frame, text="Generated Query", padding="3")
        query_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=2)
        query_frame.columnconfigure(0, weight=1)
        query_frame.rowconfigure(0, weight=1)
        
        self.query_text = scrolledtext.ScrolledText(query_frame, height=3, wrap=tk.WORD)
        self.query_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Buttons - HORIZONTAL
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, pady=2)
        
        ttk.Button(button_frame, text="Generate Query", command=self.generate_query).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Copy to Clipboard", command=self.copy_to_clipboard).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Save to File", command=self.save_to_file).pack(side=tk.LEFT, padx=2)
        
        # Status bar
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=7, column=0, sticky=(tk.W, tk.E))
        
        self.status_var = tk.StringVar(value="Ready. Select ECIs and apps to generate query.")
        status_bar = ttk.Label(status_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Footer - centered at the bottom
        footer_frame = ttk.Frame(main_frame)
        footer_frame.grid(row=8, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        version_label = ttk.Label(footer_frame, text="Fadzli Abdullah. Huawei Technologies.",
                                 font=('TkDefaultFont', 8), foreground='gray')
        version_label.pack(anchor='center')
        
        # Load mapping file
        self.load_mapping_file()
    
    def quick_date_select(self, days):
        """Quick date selection buttons"""
        ref_date = self.reference_date.get_date()
        self.end_date.set_date(ref_date)
        self.start_date.set_date(ref_date - timedelta(days=days-1))
        self.update_partition_display()
    
    def update_partition_display(self):
        """Update partition number display"""
        try:
            start_date = self.start_date.get_date()
            end_date = self.end_date.get_date()
            partitions = self.calculate_partition_numbers(start_date, end_date)
            
            if len(partitions) <= 10:
                partition_str = ', '.join(map(str, partitions))
            else:
                partition_str = f"{partitions[0]} to {partitions[-1]} ({len(partitions)} partitions)"
            
            self.partition_var.set(partition_str)
        except Exception:
            self.partition_var.set("")
    
    def load_mapping_file(self):
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
                else:
                    self.mapping_status_label.config(text="MAPPING.csv not found", foreground='red')
                    self.status_var.set("Warning: MAPPING.csv not found. Place it in the same directory as the script.")
        except Exception as e:
            self.mapping_status_label.config(text=f"Error: {str(e)}", foreground='red')
            self.status_var.set(f"Error loading MAPPING.csv: {str(e)}")
    
    def load_mapping_from_file(self, filename):
        """Load mapping from specified file with robust CSV handling"""
        try:
            import csv
            self.cell_mapping.clear()
            self.enodeb_mapping.clear()
            loaded_count = 0
            
            # Read file with proper encoding
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
            
            if loaded_count > 0:
                # Build eNodeB name mapping from cell_mapping
                self.build_enodeb_mapping()
                unique_enodebs = len(self.enodeb_mapping)
                self.mapping_status_label.config(
                    text=f"{loaded_count} ({unique_enodebs})", 
                    foreground='green'
                )
                self.status_var.set(f"Successfully loaded {loaded_count} sector ID mappings ({unique_enodebs} unique eNodeBs) from {os.path.basename(filename)}")
            else:
                self.mapping_status_label.config(text="No valid mappings found", foreground='red')
                self.status_var.set("No valid mappings found in file")
        
        except Exception as e:
            self.mapping_status_label.config(text="Error loading file", foreground='red')
            self.status_var.set(f"Error loading mapping file: {str(e)}")
    
    def _process_mapping_row(self, row):
        """Process a single mapping row and add entries to cell_mapping and enodeb_mapping
        
        Supports multiple formats:
        - 5-column format: 4LRD, 5LRD, eNodeB Name, Sector ID, eNodeB ID
        - 2-column format: Sector ID, eNodeB ID
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
                
                # Add eNodeB Name mapping
                if enodeb_name and enodeb_name != 'NAN' and enodeb_id >= 0:
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
    
    def build_enodeb_mapping(self):
        """Build eNodeB name to ID mapping from cell_mapping.
        Adds fallback mappings from Sector ID prefix for backward compatibility"""
        
        # Add fallback mappings from Sector ID prefixes if eNodeB Name wasn't in the file
        for sector_id, enodeb_id in self.cell_mapping.items():
            # Extract eNodeB name (part before underscore) as fallback
            if '_' in sector_id:
                enodeb_name = sector_id.split('_')[0].strip().upper()
                if enodeb_name:
                    # Only add if not already in mapping (Column C takes precedence)
                    if enodeb_name not in self.enodeb_mapping:
                        self.enodeb_mapping[enodeb_name] = enodeb_id
    
    def paste_to_converter(self):
        """Paste clipboard content to converter input"""
        try:
            clipboard_content = self.root.clipboard_get()
            current_text = self.converter_input.get(1.0, tk.END).strip()
            if current_text:
                self.converter_input.insert(tk.END, '\n' + clipboard_content)
            else:
                self.converter_input.insert(1.0, clipboard_content)
        except Exception as e:
            self.status_var.set(f"Failed to paste: {str(e)}")
    
    def on_converter_mode_change(self):
        """Update converter label based on selected mode"""
        mode = self.converter_mode.get()
        if mode == "name":
            self.converter_label.config(text="eNodeB Names (one per line or comma-separated):")
        elif mode == "decimal":
            self.converter_label.config(text="Decimal IDs (one per line or comma-separated):")
        else:
            self.converter_label.config(text="Sector IDs (one per line or comma-separated):")
    
    def convert_and_add_all(self):
        """Convert input and automatically add to ECI selection"""
        input_text = self.converter_input.get(1.0, tk.END).strip()
        if not input_text:
            self.status_var.set("Please enter values to convert")
            return
        
        mode = self.converter_mode.get()
        results = []
        added_count = 0
        
        lines = [line.strip() for line in input_text.split('\n') if line.strip()]
        
        for line in lines:
            items = [item.strip() for item in re.split(r'[,\s]+', line) if item.strip()]
            
            for item in items:
                if mode == "name":
                    result = self.convert_name_to_hex(item)
                elif mode == "decimal":
                    result = self.convert_decimal_to_hex(item)
                else:
                    result = self.convert_sector_to_hex(item)
                
                if result and "Error" not in result:
                    ecis = [eci.strip() for eci in result.split(',')]
                    for eci in ecis:
                        if eci and self.validate_eci(eci):
                            if eci not in self.selected_ecis:
                                self.selected_ecis.append(eci)
                                added_count += 1
                    results.append(result)
                elif result:
                    results.append(result)
        
        self.converter_result.config(state=tk.NORMAL)
        self.converter_result.delete(1.0, tk.END)
        self.converter_result.insert(1.0, '\n'.join(results))
        self.converter_result.config(state=tk.DISABLED)
        
        self.update_eci_display()
        self.status_var.set(f"Converted and added {added_count} ECI(s) to selection")
    
    def convert_name_to_hex(self, enodeb_name):
        """Convert eNodeB name to hex ECIs"""
        enodeb_name = enodeb_name.strip().upper()
        
        if enodeb_name not in self.enodeb_mapping:
            return f"{enodeb_name}: Error - Not found in mapping"
        
        enodeb_id = self.enodeb_mapping[enodeb_name]
        hex_prefix = format(enodeb_id, '05X')
        
        cells = []
        for cell_name, cell_enodeb_id in self.cell_mapping.items():
            if cell_enodeb_id == enodeb_id:
                if '_' in cell_name:
                    base, cell_num = cell_name.rsplit('_', 1)
                    if base.upper() == enodeb_name:
                        try:
                            cell_hex = format(int(cell_num), '02X')
                            cells.append(f"{hex_prefix}{cell_hex}")
                        except ValueError:
                            continue
        
        if not cells:
            for i in range(1, 4):
                cells.append(f"{hex_prefix}{format(i, '02X')}")
        
        cells.sort()
        return ', '.join(cells)
    
    def convert_decimal_to_hex(self, decimal_str):
        """Convert decimal eNodeB ID to hex"""
        try:
            decimal_id = int(decimal_str.strip())
            if decimal_id < 0 or decimal_id > 1048575:
                return f"{decimal_str}: Error - Out of range (0-1048575)"
            
            hex_value = format(decimal_id, '05X')
            ecis = [f"{hex_value}{format(i, '02X')}" for i in range(1, 4)]
            return ', '.join(ecis)
        except ValueError:
            return f"{decimal_str}: Error - Invalid decimal number"
    
    def convert_sector_to_hex(self, sector_id):
        """Convert sector ID to hex ECI"""
        sector_id = sector_id.strip().upper()
        
        if '_' not in sector_id:
            return f"{sector_id}: Error - Invalid format (use XXXXX_Y)"
        
        if sector_id not in self.cell_mapping:
            return f"{sector_id}: Error - Not found in mapping"
        
        enodeb_id = self.cell_mapping[sector_id]
        hex_prefix = format(enodeb_id, '05X')
        
        _, cell_num = sector_id.rsplit('_', 1)
        try:
            cell_hex = format(int(cell_num), '02X')
            return f"{hex_prefix}{cell_hex}"
        except ValueError:
            return f"{sector_id}: Error - Invalid cell number"
    
    def clear_converter_input(self):
        """Clear converter input and results"""
        self.converter_input.delete(1.0, tk.END)
        self.converter_result.config(state=tk.NORMAL)
        self.converter_result.delete(1.0, tk.END)
        self.converter_result.config(state=tk.DISABLED)
    
    def validate_eci(self, eci):
        """Validate ECI format"""
        eci = eci.strip().upper()
        if not re.match(r'^[0-9A-F]{5,8}$', eci):
            return False
        return True
    
    def add_eci(self):
        """Add ECI to selection"""
        eci = self.eci_entry.get().strip().upper()
        
        if not eci:
            self.status_var.set("Please enter an ECI")
            return
        
        if not self.validate_eci(eci):
            self.status_var.set("Invalid ECI format. Use 5-8 digit hexadecimal.")
            return
        
        if eci in self.selected_ecis:
            self.status_var.set(f"ECI {eci} already added")
            return
        
        self.selected_ecis.append(eci)
        self.update_eci_display()
        self.eci_entry.delete(0, tk.END)
        self.status_var.set(f"Added ECI: {eci}")
    
    def paste_bulk_eci(self):
        """Paste bulk ECIs from clipboard"""
        try:
            clipboard_content = self.root.clipboard_get()
            ecis = re.findall(r'[0-9A-Fa-f]{5,8}', clipboard_content)
            
            added_count = 0
            for eci in ecis:
                eci = eci.upper()
                if self.validate_eci(eci) and eci not in self.selected_ecis:
                    self.selected_ecis.append(eci)
                    added_count += 1
            
            self.update_eci_display()
            self.status_var.set(f"Added {added_count} ECI(s) from clipboard")
        except Exception as e:
            self.status_var.set(f"Failed to paste: {str(e)}")
    
    def clear_ecis(self):
        """Clear all selected ECIs"""
        self.selected_ecis.clear()
        self.update_eci_display()
        self.status_var.set("Cleared all ECIs")
    
    def update_eci_display(self):
        """Update the ECI text display"""
        self.eci_text.delete(1.0, tk.END)
        if self.selected_ecis:
            self.eci_text.insert(1.0, ', '.join(self.selected_ecis))
    
    def toggle_select_all(self):
        """Toggle all app selections"""
        select_all = self.select_all_var.get()
        for var in self.app_vars.values():
            var.set(select_all)
    
    def update_select_all(self):
        """Update Select All checkbox based on individual selections"""
        all_selected = all(var.get() for var in self.app_vars.values())
        self.select_all_var.set(all_selected)
    
    def calculate_partition_numbers(self, start_date, end_date):
        """Calculate partition numbers from date range"""
        base_date = datetime(2025, 11, 3).date()
        
        if isinstance(start_date, str):
            start_date = datetime.strptime(start_date, '%Y-%m-%d').date()
        if isinstance(end_date, str):
            end_date = datetime.strptime(end_date, '%Y-%m-%d').date()
        
        start_offset = (start_date - base_date).days
        end_offset = (end_date - base_date).days
        
        start_partition = 20395 + start_offset
        end_partition = 20395 + end_offset
        
        partitions = list(range(start_partition, end_partition + 1))
        return partitions
    
    def generate_query(self):
        """Generate the SQL query"""
        if not self.selected_ecis:
            self.status_var.set("Please select at least one ECI")
            return
        
        selected_apps = [app_id for app_id, var in self.app_vars.items() if var.get()]
        if not selected_apps:
            self.status_var.set("Please select at least one application")
            return
        
        include_resolution = self.resolution_var.get()
        
        if include_resolution:
            self.generate_query_with_resolution()
        else:
            self.generate_query_without_resolution()
    
    def generate_query_with_resolution(self):
        """Generate query with resolution analysis"""
        start_date = self.start_date.get_date().strftime('%Y-%m-%d')
        end_date = self.end_date.get_date().strftime('%Y-%m-%d')
        
        partitions = self.calculate_partition_numbers(start_date, end_date)
        
        rat = self.rat_var.get()
        selected_apps = [app_id for app_id, var in self.app_vars.items() if var.get()]
        app_ids = ', '.join(selected_apps)
        eci_list = "', '".join(self.selected_ecis)
        
        # Generate UNION ALL statements
        union_statements = []
        for partition in partitions:
            union_statements.append(
                f"    SELECT * FROM ps.detail_ufdr_streaming_{partition} "
                f"WHERE rat IN ({rat}) AND app_id IN ({app_ids}) AND eci IN ('{eci_list}')"
            )
        
        union_clause = "\n    UNION ALL ".join(union_statements)
        
        query = f"""-- Streaming Data Query (Backend - WITH Resolution)
-- Date Range: {start_date} to {end_date}
-- Partitions: {partitions[0]} to {partitions[-1]} ({len(partitions)} days)
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
  FROM (
{union_clause}
  ) a
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
        
        self.query_text.delete(1.0, tk.END)
        self.query_text.insert(1.0, query)
        
        selected_apps = [self.apps[app_id] for app_id, var in self.app_vars.items() if var.get()]
        apps_str = ', '.join(selected_apps)
        
        self.status_var.set(f"Backend query generated (with Resolution) for {len(self.selected_ecis)} ECIs, {len(selected_apps)} app(s) ({apps_str}), {len(partitions)} day(s)")
    
    def generate_query_without_resolution(self):
        """Generate query without resolution analysis"""
        start_date = self.start_date.get_date().strftime('%Y-%m-%d')
        end_date = self.end_date.get_date().strftime('%Y-%m-%d')
        
        partitions = self.calculate_partition_numbers(start_date, end_date)
        
        rat = self.rat_var.get()
        selected_apps = [app_id for app_id, var in self.app_vars.items() if var.get()]
        app_ids = ', '.join(selected_apps)
        eci_list = "', '".join(self.selected_ecis)
        
        # Generate UNION ALL statements
        union_statements = []
        for partition in partitions:
            union_statements.append(
                f"    SELECT * FROM ps.detail_ufdr_streaming_{partition} "
                f"WHERE rat IN ({rat}) AND app_id IN ({app_ids}) AND eci IN ('{eci_list}')"
            )
        
        union_clause = "\n    UNION ALL ".join(union_statements)
        
        query = f"""-- Streaming Data Query (Backend - WITHOUT Resolution)
-- Date Range: {start_date} to {end_date}
-- Partitions: {partitions[0]} to {partitions[-1]} ({len(partitions)} days)
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
  FROM (
{union_clause}
  ) a
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
        
        selected_apps = [self.apps[app_id] for app_id, var in self.app_vars.items() if var.get()]
        apps_str = ', '.join(selected_apps)
        
        self.status_var.set(f"Backend query generated (without Resolution) for {len(self.selected_ecis)} ECIs, {len(selected_apps)} app(s) ({apps_str}), {len(partitions)} day(s)")
    
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
            initialfile=f"streaming_query_backend_{datetime.now().strftime('%Y%m%d_%H%M%S')}.sql"
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