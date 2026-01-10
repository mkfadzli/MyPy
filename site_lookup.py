import sys
import os
import sqlite3
import pandas as pd
from datetime import datetime
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QPushButton, QLineEdit, QLabel, 
                              QFileDialog, QMessageBox, QTextEdit, QProgressBar,
                              QFrame, QGroupBox, QMenu)
from PyQt6.QtCore import Qt, QCoreApplication, QTimer
from PyQt6.QtGui import QFont, QKeySequence, QShortcut, QAction, QIcon

class CachedSiteLookup(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
        self.log_visible = False
        self.database_path = None
        self.cache_db_path = None
        self.search_count = 0
        
        # Set up cache directory in user's home folder
        self.cache_dir = Path.home() / ".site_lookup_cache"
        self.cache_dir.mkdir(exist_ok=True)
        self.cache_db_path = self.cache_dir / "site_database.db"
        
        self.init_ui()
        self.setup_shortcuts()
        self.add_log("Application initialized", "INFO")
        
        # Auto-load from cache if available
        QTimer.singleShot(100, self.auto_load_cache)
        
    def init_ui(self):
        self.setWindowTitle("Site Lookup")
        self.setMinimumSize(500, 260)
        self.resize(500, 260)
        
        # Set application icon
        icon_path = Path(__file__).parent / "app_icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        
        # Enhanced Windows Classic style
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ECE9D8;
            }
            
            /* Enhanced Classic Buttons */
            QPushButton {
                background-color: #D4D0C8;
                border-top: 2px solid #FFFFFF;
                border-left: 2px solid #FFFFFF;
                border-right: 2px solid #808080;
                border-bottom: 2px solid #808080;
                border-radius: 0px;
                padding: 4px 10px;
                font-family: 'Ubuntu', 'Segoe UI', 'Tahoma', sans-serif;
                font-size: 9pt;
                font-weight: 600;
                color: #000000;
                min-height: 20px;
            }
            QPushButton:hover {
                background-color: #E8E5DD;
            }
            QPushButton:pressed {
                border-top: 2px solid #808080;
                border-left: 2px solid #808080;
                border-right: 2px solid #FFFFFF;
                border-bottom: 2px solid #FFFFFF;
                background-color: #C0BEB0;
                padding: 5px 9px 3px 11px;
            }
            QPushButton:disabled {
                color: #999999;
                background-color: #D4D0C8;
            }
            QPushButton#primaryBtn {
                background-color: #5B8BB9;
                font-weight: bold;
                color: #FFFFFF;
            }
            QPushButton#primaryBtn:hover {
                background-color: #6A9BC9;
                color: #FFFFFF;
            }
            QPushButton#primaryBtn:pressed {
                background-color: #4A7AA8;
                color: #FFFFFF;
            }
            QPushButton#compactBtn {
                padding: 2px 6px;
                font-size: 8pt;
                min-height: 16px;
            }
            
            QLineEdit {
                background-color: #FFFFFF;
                border-top: 2px solid #7F9DB9;
                border-left: 2px solid #7F9DB9;
                border-right: 2px solid #E3E3E3;
                border-bottom: 2px solid #E3E3E3;
                padding: 3px 4px;
                font-family: 'Ubuntu', 'Segoe UI', 'Tahoma', sans-serif;
                font-size: 9pt;
                color: #000000;
                selection-background-color: #316AC5;
                selection-color: #FFFFFF;
            }
            QLineEdit:disabled {
                background-color: #E8E5DD;
                color: #999999;
                border-top: 2px solid #999999;
                border-left: 2px solid #999999;
            }
            QLineEdit:focus {
                background-color: #FFFFEE;
                border-top: 2px solid #5B8BB9;
                border-left: 2px solid #5B8BB9;
            }
            
            QLabel {
                background-color: transparent;
                font-family: 'Ubuntu', 'Segoe UI', 'Tahoma', sans-serif;
                font-size: 9pt;
                color: #000000;
            }
            
            QGroupBox {
                background-color: #ECE9D8;
                border-top: 2px solid #FFFFFF;
                border-left: 2px solid #FFFFFF;
                border-right: 2px solid #808080;
                border-bottom: 2px solid #808080;
                border-radius: 0px;
                margin-top: 8px;
                padding-top: 8px;
                font-family: 'Ubuntu', 'Segoe UI', 'Tahoma', sans-serif;
                font-size: 9pt;
                font-weight: bold;
                color: #000066;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 8px;
                padding: 0 3px;
                background-color: #ECE9D8;
            }
            
            QFrame#resultPanel {
                background-color: #FFFFFF;
                border-top: 2px solid #808080;
                border-left: 2px solid #808080;
                border-right: 2px solid #E3E3E3;
                border-bottom: 2px solid #E3E3E3;
            }
            
            QFrame#statusBar {
                background-color: #D4D0C8;
                border-top: 2px solid #FFFFFF;
                border-bottom: 1px solid #808080;
            }
            
            QProgressBar {
                border-top: 2px solid #808080;
                border-left: 2px solid #808080;
                border-right: 2px solid #E3E3E3;
                border-bottom: 2px solid #E3E3E3;
                background-color: #FFFFFF;
                text-align: center;
                font-family: 'Ubuntu', 'Segoe UI', 'Tahoma', sans-serif;
                font-size: 8pt;
                color: #000000;
                height: 18px;
            }
            QProgressBar::chunk {
                background-color: #316AC5;
                border: 1px solid #5B8BB9;
            }
            
            QTextEdit {
                background-color: #FFFFFF;
                border-top: 2px solid #808080;
                border-left: 2px solid #808080;
                border-right: 2px solid #E3E3E3;
                border-bottom: 2px solid #E3E3E3;
                font-family: 'Ubuntu Mono', 'Consolas', 'Courier New', monospace;
                font-size: 8pt;
                color: #000000;
                padding: 2px;
            }
            
            QMenu {
                background-color: #FFFFFF;
                border: 2px solid #808080;
                padding: 2px;
                font-family: 'Ubuntu', 'Segoe UI', 'Tahoma', sans-serif;
                font-size: 9pt;
            }
            QMenu::item {
                padding: 3px 20px;
                background-color: transparent;
                color: #000000;
            }
            QMenu::item:selected {
                background-color: #316AC5;
                color: #FFFFFF;
            }
        """)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)
        
        # === TOP BUTTON ROW ===
        button_row = QHBoxLayout()
        button_row.setSpacing(4)
        
        # Load button with dropdown menu
        self.load_btn = QPushButton("Load Database ▼")
        self.load_btn.setObjectName("primaryBtn")
        self.load_btn.clicked.connect(self.show_load_menu)
        button_row.addWidget(self.load_btn)
        
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self.clear_results)
        self.clear_btn.setEnabled(False)
        button_row.addWidget(self.clear_btn)
        
        button_row.addStretch()
        
        layout.addLayout(button_row)
        
        # === PROGRESS BAR ===
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(18)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # === SEARCH GROUP ===
        search_group = QGroupBox("Search Criteria")
        search_layout = QHBoxLayout(search_group)
        search_layout.setContentsMargins(6, 10, 6, 6)
        search_layout.setSpacing(4)
        
        site_label = QLabel("Site Name:")
        site_label.setMinimumWidth(65)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter site (single or multiple): 4LRD or 6LRD")
        self.search_input.returnPressed.connect(self.handle_search)
        self.search_input.textChanged.connect(self.on_text_changed)
        self.search_input.setEnabled(False)
        
        search_layout.addWidget(site_label)
        search_layout.addWidget(self.search_input)
        
        layout.addWidget(search_group)
        
        # === RESULTS PANEL ===
        results_frame = QFrame()
        results_frame.setObjectName("resultPanel")
        results_layout = QVBoxLayout(results_frame)
        results_layout.setContentsMargins(8, 6, 8, 6)
        results_layout.setSpacing(4)
        
        results_header = QLabel("Site Information")
        results_header.setStyleSheet("""
            font-weight: bold;
            color: #000066;
            font-size: 9pt;
            border-bottom: 1px solid #D4D0C8;
            padding-bottom: 2px;
        """)
        results_layout.addWidget(results_header)
        
        # eNodeBID row
        enodeb_container = QHBoxLayout()
        enodeb_container.setSpacing(8)
        
        enodeb_icon = QLabel("●")
        enodeb_icon.setStyleSheet("color: #316AC5; font-size: 10pt;")
        
        enodeb_label = QLabel("eNodeB ID:")
        enodeb_label.setMinimumWidth(70)
        enodeb_label.setStyleSheet("color: #000000; font-weight: 600;")
        
        self.enodeb_value = QLineEdit("")
        self.enodeb_value.setReadOnly(True)
        self.enodeb_value.setStyleSheet("""
            color: #000080;
            font-weight: bold;
            font-size: 10pt;
            font-family: 'Consolas', 'Courier New', monospace;
            background-color: #FFFFFF;
            border: 1px solid #D4D0C8;
            padding: 2px 4px;
        """)
        
        enodeb_container.addWidget(enodeb_icon)
        enodeb_container.addWidget(enodeb_label)
        enodeb_container.addWidget(self.enodeb_value, 1)
        results_layout.addLayout(enodeb_container)
        
        # Region row
        region_container = QHBoxLayout()
        region_container.setSpacing(8)
        
        region_icon = QLabel("●")
        region_icon.setStyleSheet("color: #CC3333; font-size: 10pt;")
        
        region_label = QLabel("Region:")
        region_label.setMinimumWidth(70)
        region_label.setStyleSheet("color: #000000; font-weight: 600;")
        
        self.region_value = QLineEdit("")
        self.region_value.setReadOnly(True)
        self.region_value.setStyleSheet("""
            color: #800000;
            font-weight: bold;
            font-size: 10pt;
            font-family: 'Consolas', 'Courier New', monospace;
            background-color: #FFFFFF;
            border: 1px solid #D4D0C8;
            padding: 2px 4px;
        """)
        
        region_container.addWidget(region_icon)
        region_container.addWidget(region_label)
        region_container.addWidget(self.region_value, 1)
        results_layout.addLayout(region_container)
        
        layout.addWidget(results_frame)
        
        # === STATUS BAR ===
        status_frame = QFrame()
        status_frame.setObjectName("statusBar")
        status_layout = QHBoxLayout(status_frame)
        status_layout.setContentsMargins(6, 3, 6, 3)
        status_layout.setSpacing(6)
        
        # LED indicator
        self.status_led = QLabel("●")
        self.status_led.setStyleSheet("color: #CC3333; font-size: 9pt;")
        status_layout.addWidget(self.status_led)
        
        self.status_message = QLabel("Loading cache...")
        self.status_message.setStyleSheet("""
            color: #000000;
            font-size: 8pt;
            background: transparent;
            border: none;
        """)
        status_layout.addWidget(self.status_message)
        
        status_layout.addStretch()
        
        # Cache indicator
        self.cache_indicator = QLabel("💾")
        self.cache_indicator.setStyleSheet("font-size: 9pt;")
        self.cache_indicator.setToolTip("Using cached database")
        self.cache_indicator.setVisible(False)
        status_layout.addWidget(self.cache_indicator)
        
        # Search counter
        counter_label = QLabel("Searches:")
        counter_label.setStyleSheet("color: #666666; font-size: 8pt;")
        status_layout.addWidget(counter_label)
        
        self.search_count_label = QLabel("0")
        self.search_count_label.setStyleSheet("""
            color: #000000;
            font-size: 8pt;
            font-weight: bold;
            background: transparent;
            border: none;
        """)
        status_layout.addWidget(self.search_count_label)
        
        # Log toggle
        self.toggle_log_btn = QPushButton("▼ Log")
        self.toggle_log_btn.setObjectName("compactBtn")
        self.toggle_log_btn.setMaximumWidth(55)
        self.toggle_log_btn.clicked.connect(self.toggle_log)
        status_layout.addWidget(self.toggle_log_btn)
        
        layout.addWidget(status_frame)
        
        # === FOOTER MESSAGE ===
        footer_frame = QFrame()
        footer_frame.setObjectName("footerBar")
        footer_frame.setStyleSheet("""
            QFrame#footerBar {
                background-color: #D4D0C8;
                border-top: 1px solid #FFFFFF;
            }
        """)
        footer_layout = QHBoxLayout(footer_frame)
        footer_layout.setContentsMargins(6, 2, 6, 2)
        
        footer_label = QLabel("V1.0.4025. Fadzli Abdullah. Huawei Technologies")
        footer_label.setStyleSheet("""
            color: #666666;
            font-size: 7pt;
            font-weight: italic;
            font-family: 'Ubuntu', 'Segoe UI', 'Tahoma', sans-serif;
        """)
        footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer_layout.addWidget(footer_label)
        
        layout.addWidget(footer_frame)
        
        # === LOG PANEL ===
        log_container = QWidget()
        log_layout = QVBoxLayout(log_container)
        log_layout.setContentsMargins(0, 4, 0, 0)
        log_layout.setSpacing(2)
        
        log_header_row = QHBoxLayout()
        log_header = QLabel("Processing Log")
        log_header.setStyleSheet("""
            font-weight: bold;
            font-size: 8pt;
            color: #000066;
        """)
        log_header_row.addWidget(log_header)
        
        log_clear_btn = QPushButton("Clear Log")
        log_clear_btn.setObjectName("compactBtn")
        log_clear_btn.setMaximumWidth(70)
        log_clear_btn.clicked.connect(lambda: self.log_text.clear())
        log_header_row.addWidget(log_clear_btn)
        log_header_row.addStretch()
        
        log_layout.addLayout(log_header_row)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(130)
        log_layout.addWidget(self.log_text)
        
        layout.addWidget(log_container)
        self.log_container = log_container
        log_container.hide()
        
        self.search_btn = None
    
    def show_load_menu(self):
        """Show load options menu"""
        menu = QMenu(self)
        
        # Load from file (import new)
        load_action = QAction("Load from File... (Ctrl+L)", self)
        load_action.triggered.connect(self.load_database_from_file)
        menu.addAction(load_action)
        
        # Refresh from last source
        if self.database_path and os.path.exists(self.database_path):
            refresh_action = QAction("Refresh from Last Source (Ctrl+R)", self)
            refresh_action.triggered.connect(self.refresh_from_source)
            menu.addAction(refresh_action)
        
        menu.addSeparator()
        
        # Cache info
        cache_action = QAction("Cache Information...", self)
        cache_action.triggered.connect(self.show_cache_info)
        menu.addAction(cache_action)
        
        # Clear cache
        clear_cache_action = QAction("Clear Cache", self)
        clear_cache_action.triggered.connect(self.clear_cache)
        menu.addAction(clear_cache_action)
        
        # Show menu at button position
        menu.exec(self.load_btn.mapToGlobal(self.load_btn.rect().bottomLeft()))
    
    def setup_shortcuts(self):
        """Setup keyboard shortcuts"""
        load_shortcut = QShortcut(QKeySequence("Ctrl+L"), self)
        load_shortcut.activated.connect(self.load_database_from_file)
        
        refresh_shortcut = QShortcut(QKeySequence("Ctrl+R"), self)
        refresh_shortcut.activated.connect(self.refresh_from_source)
        
        focus_shortcut = QShortcut(QKeySequence("Ctrl+F"), self)
        focus_shortcut.activated.connect(self.focus_search)
        
        clear_shortcut = QShortcut(QKeySequence("Escape"), self)
        clear_shortcut.activated.connect(self.clear_search_input)
    
    def focus_search(self):
        """Focus search input"""
        if self.search_input.isEnabled():
            self.search_input.setFocus()
            self.search_input.selectAll()
    
    def on_text_changed(self):
        """React to text changes"""
        pass
    
    def add_log(self, message, level="INFO"):
        """Add log entry"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        color_map = {
            "INFO": "#0066CC",
            "SUCCESS": "#008800",
            "WARNING": "#CC6600",
            "ERROR": "#CC0000",
            "DEBUG": "#666666"
        }
        color = color_map.get(level, "#000000")
        
        log_entry = f'<span style="color: #666666;">[{timestamp}]</span> <span style="color: {color}; font-weight: bold;">[{level}]</span> {message}'
        self.log_text.append(log_entry)
        
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def update_progress(self, value, text=""):
        """Update progress bar"""
        self.progress_bar.setValue(value)
        if text:
            self.progress_bar.setFormat(f"{value}% - {text}")
        else:
            self.progress_bar.setFormat(f"{value}%")
        QCoreApplication.processEvents()
    
    def update_status_led(self, color="red", message=""):
        """Update LED indicator"""
        color_map = {
            "red": "#CC3333",
            "green": "#33CC33",
            "yellow": "#CCCC33",
            "blue": "#3366CC"
        }
        self.status_led.setStyleSheet(f"color: {color_map.get(color, '#CC3333')}; font-size: 9pt;")
        if message:
            self.status_message.setText(message)
    
    def toggle_log(self):
        """Toggle log panel"""
        self.log_visible = not self.log_visible
        if self.log_visible:
            self.log_container.show()
            self.toggle_log_btn.setText("▲ Log")
            self.resize(380, 450)
            self.add_log("Log panel opened", "DEBUG")
        else:
            self.log_container.hide()
            self.toggle_log_btn.setText("▼ Log")
            self.resize(380, 260)
    
    def clear_results(self):
        """Clear search results"""
        self.enodeb_value.setText("-")
        self.region_value.setText("-")
        self.status_message.setText(f"Ready: {len(self.df):,} records" if self.df is not None else "No database loaded")
        self.search_input.clear()
        self.search_input.setFocus()
        self.clear_btn.setEnabled(False)
        self.add_log("Results cleared", "DEBUG")
    
    def clear_search_input(self):
        """Clear search input"""
        if self.search_input.isEnabled():
            self.search_input.clear()
            self.search_input.setFocus()
    
    def get_cache_metadata(self):
        """Get metadata from cache"""
        try:
            conn = sqlite3.connect(self.cache_db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT source_file, import_date, record_count FROM metadata")
            result = cursor.fetchone()
            conn.close()
            return result
        except:
            return None
    
    def auto_load_cache(self):
        """Automatically load from cache on startup"""
        if not self.cache_db_path.exists():
            self.add_log("No cache found - please load database", "INFO")
            self.update_status_led("red", "No database loaded")
            return
        
        try:
            self.add_log("Loading from cache...", "INFO")
            self.update_status_led("yellow", "Loading cache...")
            
            # Load from SQLite
            conn = sqlite3.connect(self.cache_db_path)
            self.df = pd.read_sql_query("SELECT * FROM sites", conn)
            
            # Get metadata
            metadata = self.get_cache_metadata()
            if metadata:
                source_file, import_date, record_count = metadata
                self.database_path = source_file
                self.add_log(f"Cache loaded: {record_count:,} records", "SUCCESS")
                self.add_log(f"Source: {Path(source_file).name}", "DEBUG")
                self.add_log(f"Cached: {import_date}", "DEBUG")
                
                # Check if source file still exists and is newer
                if os.path.exists(source_file):
                    source_mtime = os.path.getmtime(source_file)
                    cache_mtime = os.path.getmtime(self.cache_db_path)
                    if source_mtime > cache_mtime:
                        self.add_log("⚠ Source file has been updated!", "WARNING")
                        self.update_status_led("yellow", f"Cache outdated - {record_count:,} records")
                        self.cache_indicator.setVisible(True)
                        self.cache_indicator.setToolTip("Cache is outdated! Use 'Refresh from Last Source' to update.")
                    else:
                        self.update_status_led("green", f"Ready: {record_count:,} records (cached)")
                        self.cache_indicator.setVisible(True)
                        self.cache_indicator.setToolTip(f"Using cache from {import_date}")
                else:
                    self.add_log("Original source file not found", "WARNING")
                    self.update_status_led("green", f"Ready: {record_count:,} records (cached)")
                    self.cache_indicator.setVisible(True)
                    self.cache_indicator.setToolTip("Using cached data (source file moved)")
            
            conn.close()
            
            self.search_input.setEnabled(True)
            self.search_count = 0
            self.setWindowTitle(f" Site Lookup - {Path(source_file).name if metadata else 'Cached'}")
            
        except Exception as e:
            self.add_log(f"Failed to load cache: {str(e)}", "ERROR")
            self.update_status_led("red", "Cache load failed")
            self.cache_indicator.setVisible(False)
    
    def save_to_cache(self, source_file):
        """Save dataframe to SQLite cache"""
        try:
            self.add_log("Saving to cache...", "INFO")
            
            # Create SQLite database
            conn = sqlite3.connect(self.cache_db_path)
            
            # Save data
            self.df.to_sql('sites', conn, if_exists='replace', index=False)
            
            # Save metadata
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS metadata (
                    source_file TEXT,
                    import_date TEXT,
                    record_count INTEGER
                )
            """)
            cursor.execute("DELETE FROM metadata")
            cursor.execute("""
                INSERT INTO metadata VALUES (?, ?, ?)
            """, (source_file, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), len(self.df)))
            
            # Create index for faster searches
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_enodeb_name ON sites ([eNodeB Name])")
            
            conn.commit()
            conn.close()
            
            self.add_log(f"Cache saved: {len(self.df):,} records", "SUCCESS")
            self.cache_indicator.setVisible(True)
            self.cache_indicator.setToolTip(f"Cache created: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
        except Exception as e:
            self.add_log(f"Failed to save cache: {str(e)}", "ERROR")
    
    def load_database_from_file(self):
        """Load database from file and cache it"""
        self.add_log("Opening file dialog", "INFO")
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Database File",
            "",
            "Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*)"
        )
        
        if not file_path:
            self.add_log("File selection cancelled", "INFO")
            return
        
        self.load_and_cache_database(file_path)
    
    def refresh_from_source(self):
        """Refresh database from last loaded source"""
        if not self.database_path or not os.path.exists(self.database_path):
            QMessageBox.warning(self, "Source Not Found", 
                              "Original source file not found. Please load a new database.")
            return
        
        self.add_log(f"Refreshing from: {Path(self.database_path).name}", "INFO")
        self.load_and_cache_database(self.database_path)
    
    def load_and_cache_database(self, file_path):
        """Load database from file and save to cache"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.load_btn.setEnabled(False)
        self.search_input.setEnabled(False)
        self.update_status_led("yellow", "Loading...")
        
        try:
            filename = file_path.split('/')[-1].split('\\')[-1]
            self.update_progress(10, "Reading file")
            self.add_log(f"Loading: {filename}", "INFO")
            
            if file_path.endswith('.csv'):
                self.update_progress(25, "Parsing CSV")
                self.df = pd.read_csv(file_path, encoding='utf-8', on_bad_lines='skip')
            else:
                self.update_progress(25, "Parsing Excel")
                self.df = pd.read_excel(file_path, engine='openpyxl')
            
            self.update_progress(50, "Processing data")
            self.add_log(f"Loaded {len(self.df):,} rows", "INFO")
            
            # Verify columns
            self.update_progress(60, "Validating")
            required_cols = ['eNodeB Name', 'eNodeBID', 'Sub Region']
            missing_cols = [col for col in required_cols if col not in self.df.columns]
            
            if missing_cols:
                self.add_log(f"Missing columns: {', '.join(missing_cols)}", "ERROR")
                self.update_progress(0, "Failed")
                QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))
                self.load_btn.setEnabled(True)
                self.update_status_led("red", "Validation failed")
                QMessageBox.warning(self, "Missing Columns",
                                  f"Required columns not found:\n{', '.join(missing_cols)}")
                self.df = None
                return
            
            self.update_progress(70, "Cleaning data")
            initial_count = len(self.df)
            self.df = self.df.dropna(subset=['eNodeB Name'])
            removed = initial_count - len(self.df)
            if removed > 0:
                self.add_log(f"Removed {removed} null records", "WARNING")
            
            self.df['eNodeB Name'] = self.df['eNodeB Name'].astype(str).str.strip().str.upper()
            self.df['eNodeBID'] = self.df['eNodeBID'].astype(str).str.strip()
            self.df['Sub Region'] = self.df['Sub Region'].astype(str).str.strip()
            
            # Save to cache
            self.update_progress(85, "Caching")
            self.save_to_cache(file_path)
            
            # Complete
            self.database_path = file_path
            self.search_count = 0
            self.update_progress(100, "Complete")
            self.add_log(f"Database ready: {len(self.df):,} records", "SUCCESS")
            
            self.search_input.setEnabled(True)
            self.update_status_led("green", f"Ready: {len(self.df):,} records (cached)")
            self.setWindowTitle(f"Site Lookup - {filename}")
            
            self.clear_results()
            QTimer.singleShot(100, self.focus_search)
            QTimer.singleShot(1500, lambda: self.progress_bar.setVisible(False))
            self.load_btn.setEnabled(True)
            
        except Exception as e:
            self.add_log(f"Load error: {str(e)}", "ERROR")
            self.update_progress(0, "Error")
            QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))
            self.load_btn.setEnabled(True)
            self.search_input.setEnabled(False)
            self.update_status_led("red", "Load failed")
            QMessageBox.critical(self, "Error", f"Failed to load database:\n\n{str(e)}")
            self.df = None
            self.setWindowTitle("Site Lookup")
    
    def show_cache_info(self):
        """Show cache information"""
        if not self.cache_db_path.exists():
            QMessageBox.information(self, "Cache Info", "No cache exists yet.\n\nLoad a database to create cache.")
            return
        
        try:
            metadata = self.get_cache_metadata()
            if metadata:
                source_file, import_date, record_count = metadata
                cache_size = self.cache_db_path.stat().st_size / (1024 * 1024)  # MB
                
                info = f"""Cache Information:

Source File: {Path(source_file).name}
Full Path: {source_file}

Cached Date: {import_date}
Record Count: {record_count:,}
Cache Size: {cache_size:.2f} MB

Cache Location:
{self.cache_db_path}"""
                
                QMessageBox.information(self, "Cache Information", info)
            else:
                QMessageBox.warning(self, "Cache Info", "Cache exists but metadata is missing.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read cache info:\n\n{str(e)}")
    
    def clear_cache(self):
        """Clear the cache"""
        if not self.cache_db_path.exists():
            QMessageBox.information(self, "Clear Cache", "No cache to clear.")
            return
        
        reply = QMessageBox.question(
            self,
            "Clear Cache",
            "Are you sure you want to clear the cache?\n\nYou will need to reload the database on next startup.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                os.remove(self.cache_db_path)
                self.add_log("Cache cleared", "INFO")
                self.cache_indicator.setVisible(False)
                QMessageBox.information(self, "Cache Cleared", "Cache has been cleared successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to clear cache:\n\n{str(e)}")
    
    def handle_search(self):
        """Handle search - supports multiple space-separated site names"""
        if self.df is None:
            self.add_log("No database loaded", "ERROR")
            QMessageBox.warning(self, "No Database", "Please load a database first.")
            return
        
        search_text = self.search_input.text().strip().upper()
        if not search_text:
            return
        
        # Split input by spaces to support multiple sites
        search_terms = [term.strip() for term in search_text.split() if term.strip()]
        
        if not search_terms:
            return
        
        self.search_count += 1
        self.search_count_label.setText(str(self.search_count))
        
        if len(search_terms) == 1:
            self.add_log(f"Search #{self.search_count}: '{search_terms[0]}'", "INFO")
        else:
            self.add_log(f"Search #{self.search_count}: {len(search_terms)} sites: {' '.join(search_terms)}", "INFO")
        
        enodeb_results = []
        region_results = []
        found_count = 0
        not_found_count = 0
        
        # Process each search term in order
        for search_term in search_terms:
            # Convert 4LRD to 6LRD
            original_term = search_term
            if len(search_term) >= 5 and search_term[0] == '4':
                search_term = '6' + search_term[1:]
                self.add_log(f"Converted: {original_term} → {search_term}", "INFO")
            
            try:
                # Exact match first
                result = self.df[self.df['eNodeB Name'].str.upper() == search_term]
                match_type = "exact"
                
                # Partial match if no exact match
                if result.empty:
                    result = self.df[self.df['eNodeB Name'].str.upper().str.contains(search_term, na=False)]
                    match_type = "partial"
                
                if not result.empty:
                    enodeb_id = result.iloc[0]['eNodeBID']
                    region = result.iloc[0]['Sub Region']
                    
                    enodeb_results.append(str(enodeb_id))
                    region_results.append(str(region))
                    
                    found_count += 1
                    
                    if len(result) > 1:
                        self.add_log(f"✓ {original_term}: {enodeb_id} | {region} ({len(result)} matches, showing first)", "SUCCESS")
                    else:
                        self.add_log(f"✓ {original_term}: {enodeb_id} | {region} ({match_type})", "SUCCESS")
                    
                else:
                    enodeb_results.append("Not Found")
                    region_results.append("Not Found")
                    not_found_count += 1
                    self.add_log(f"✗ {original_term}: No match found", "WARNING")
                
            except Exception as e:
                enodeb_results.append("Error")
                region_results.append("Error")
                self.add_log(f"Search error for '{original_term}': {str(e)}", "ERROR")
        
        # Format results with aligned pipelines for visual pairing
        # Calculate the width needed for each pair (max of ID and Region length)
        pair_widths = []
        for i in range(len(enodeb_results)):
            max_width = max(len(str(enodeb_results[i])), len(str(region_results[i])))
            pair_widths.append(max_width)
        
        # Pad each result to its pair width for alignment
        padded_enodeb = [str(enodeb_results[i]).ljust(pair_widths[i]) for i in range(len(enodeb_results))]
        padded_region = [str(region_results[i]).ljust(pair_widths[i]) for i in range(len(region_results))]
        
        # Display with pipeline separator between pairs
        self.enodeb_value.setText(" | ".join(padded_enodeb))
        self.region_value.setText(" | ".join(padded_region))
        
        # Update status
        if found_count > 0 and not_found_count == 0:
            self.update_status_led("green", f"Found all {found_count} site(s)")
        elif found_count > 0 and not_found_count > 0:
            self.update_status_led("yellow", f"Found {found_count}/{len(search_terms)} site(s)")
        else:
            self.update_status_led("red", "No matches found")
        
        self.clear_btn.setEnabled(True)
        self.search_input.selectAll()

def main():
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    
    app = QApplication(sys.argv)
    app.setApplicationName("Site Lookup")
    app.setOrganizationName("Huawei Technologies")
    app.setApplicationVersion("1.0")
    
    # Set application icon globally
    icon_path = Path(__file__).parent / "app_icon.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    
    window = CachedSiteLookup()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()