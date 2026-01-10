import sys
import os
import math
import csv
import simplekml
import pandas as pd
from lxml import etree
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLineEdit, QPushButton, QTextEdit, QFileDialog, QComboBox, QLabel,
                             QFrame, QScrollArea, QMessageBox, QSizePolicy, QListWidget, QTabWidget)
from PyQt6.QtGui import QFont, QIcon, QColor
from PyQt6.QtCore import Qt, QSize, QTimer, QDateTime
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
from timezonefinder import TimezoneFinder
import pytz
from datetime import datetime
import re

class MergedTextUtility(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Fadzli Multi-Utility Tool")
        self.setGeometry(100, 100, 1100, 730)
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #1e1e1e;
                color: #ffffff;
                font-family: 'Andale Mono';
            }
            QLabel {
                font-size: 12px;
                color: #e0e0e0;
            }
            QLineEdit, QTextEdit, QComboBox, QListWidget {
                background-color: #2d2d2d;
                border: 1px solid #3a3a3a;
                border-radius: 0px;
                padding: 8px;
                font-size: 12px;
                color: #ffffff;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 5px 10px;
                text-align: center;
                text-decoration: none;
                font-size: 12px;
                margin: 4px 2px;
                border-radius: 0px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QScrollArea {
                border: none;
            }
            QTabWidget::pane {
                border: 1px solid #3a3a3a;
                background-color: #252525;
            }
            QTabBar::tab {
                background-color: #2d2d2d;
                color: #ffffff;
                padding: 8px 16px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #0078d4;
            }
        """)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(20)

        self.original_text = ""
        self.selected_text_files = []
        self.selected_conversion_files = []
        self.setup_ui()
        
        self.setWindowIcon(QIcon("myicon.ico"))
        
        self.start_time = QDateTime.currentDateTime()
        self.uptime_timer = QTimer(self)
        self.uptime_timer.timeout.connect(self.update_uptime)
        self.uptime_timer.start(1000)

        self.datetime_timer = QTimer(self)
        self.datetime_timer.timeout.connect(self.update_datetime)
        self.datetime_timer.start(1000)

    def setup_ui(self):
        self.tab_widget = QTabWidget()
        self.layout.addWidget(self.tab_widget)

        # Text Processing Tab
        text_processing_widget = self.create_text_processing()
        self.tab_widget.addTab(text_processing_widget, "Textcase Processing")

        # Distance and Coordinate Tab
        distance_coord_widget = self.create_distance_and_coordinate()
        self.tab_widget.addTab(distance_coord_widget, "Distance Coordinates")

        # File Converter Tab
        file_converter_widget = self.create_file_converter()
        self.tab_widget.addTab(file_converter_widget, "CSV/KML Converter")

        # Combined Lookup Tab
        combined_lookup_widget = self.create_combined_lookup()
        self.tab_widget.addTab(combined_lookup_widget, "Address Geocoder-Time Zone Finder-Elevation")

        # Footer
        self.create_footer()   
    
    def create_text_processing(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Text Case Converter Section
        case_converter_layout = QVBoxLayout()
        case_converter_layout.addWidget(QLabel("Text Case Converter"))

        self.input_text = QTextEdit()
        self.input_text.setPlaceholderText("Enter text to convert...")
        self.input_text.setStyleSheet("min-height: 100px;")
        self.input_text.textChanged.connect(self.update_original_text)
        case_converter_layout.addWidget(self.input_text)

        button_scroll = QScrollArea()
        button_scroll.setWidgetResizable(True)
        button_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        button_widget = QWidget()
        button_layout = QHBoxLayout(button_widget)
        button_layout.setSpacing(5)
        
        self.convert_buttons = [
            QPushButton("Lowercase"),
            QPushButton("Uppercase"),
            QPushButton("Proper Case"),
            QPushButton("Subscript"),
            QPushButton("Superscript"),
            QPushButton("Strike-through"),
            QPushButton("Reset")
        ]
        for button in self.convert_buttons:
            button.clicked.connect(self.convert_text)
            button.setMinimumWidth(120)
            button_layout.addWidget(button)
        
        button_scroll.setWidget(button_widget)
        case_converter_layout.addWidget(button_scroll)

        layout.addLayout(case_converter_layout)

        # Separator
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(separator)

        # Batch Text Processing Section
        batch_layout = QVBoxLayout()
        batch_layout.addWidget(QLabel("Batch Text Processing"))
        
        file_layout = QHBoxLayout()
        self.file_list = QListWidget()
        file_layout.addWidget(self.file_list)
        
        file_buttons_layout = QVBoxLayout()
        self.select_files_btn = QPushButton("Select Files")
        self.select_files_btn.clicked.connect(self.select_text_files)
        file_buttons_layout.addWidget(self.select_files_btn)
        
        self.clear_files_btn = QPushButton("Clear Files")
        self.clear_files_btn.clicked.connect(self.clear_text_files)
        file_buttons_layout.addWidget(self.clear_files_btn)
        
        file_layout.addLayout(file_buttons_layout)
        batch_layout.addLayout(file_layout)

        process_layout = QHBoxLayout()
        self.process_combo = QComboBox()
        self.process_combo.addItems(["Lowercase", "Uppercase", "Proper Case"])
        process_layout.addWidget(self.process_combo)
        
        button_layout = QHBoxLayout()
        self.process_btn = QPushButton("Process Files")
        self.process_btn.clicked.connect(self.process_files)
        button_layout.addWidget(self.process_btn)
        
        self.output_folder_btn = QPushButton("Output Folder")
        self.output_folder_btn.clicked.connect(self.open_output_folder)
        button_layout.addWidget(self.output_folder_btn)
        
        process_layout.addLayout(button_layout)
        batch_layout.addLayout(process_layout)

        layout.addLayout(batch_layout)

        return widget
    
    def update_original_text(self):
        current_text = self.input_text.toPlainText()
        if current_text != self.original_text:
            self.original_text = current_text

    def convert_text(self):
        text = self.input_text.toPlainText()
        sender = self.sender().text()
        
        if sender == "Lowercase":
            result = text.lower()
        elif sender == "Uppercase":
            result = text.upper()
        elif sender == "Proper Case":
            result = ' '.join(word.capitalize() for word in text.split())
        elif sender == "Subscript":
            result = self.to_subscript(text)
        elif sender == "Superscript":
            result = self.to_superscript(text)
        elif sender == "Strike-through":
            result = '\u0336'.join(text) + '\u0336' if text else ''
        elif sender == "Reset":
            result = self.original_text
        else:
            return  # If the sender is not recognized, do nothing
        
        self.input_text.setPlainText(result)
        if sender != "Reset":
            self.original_text = text  # Update original text after conversion

    def to_subscript(self, text):
        subscript_map = str.maketrans("0123456789aehijklmnoprstuvx", "₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ")
        return text.translate(subscript_map)

    def to_superscript(self, text):
        superscript_map = str.maketrans("0123456789abcdefghijklmnoprstuvwxyz", "⁰¹²³⁴⁵⁶⁷⁸⁹ᵃᵇᶜᵈᵉᶠᵍʰⁱʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ")
        return text.translate(superscript_map)

    def select_text_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Text Files", "", "Text Files (*.txt)")
        self.selected_text_files.extend(files)
        self.file_list.clear()
        self.file_list.addItems([os.path.basename(f) for f in files])
    
    def open_output_folder(self):
        output_folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if output_folder:
            QDesktopServices.openUrl(QUrl.fromLocalFile(output_folder))

    def create_distance_and_coordinate(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Distance Calculator Section
        distance_layout = QVBoxLayout()
        distance_layout.addWidget(QLabel("Distance Calculator"))

        for i in range(1, 6):
            point_layout = QHBoxLayout()
            point_layout.addWidget(QLabel(f"Point {i}"))
            lat_input = QLineEdit()
            lat_input.setPlaceholderText("Latitude")
            lat_input.setObjectName(f"lat_input_{i-1}")
            point_layout.addWidget(lat_input)
            lon_input = QLineEdit()
            lon_input.setPlaceholderText("Longitude")
            lon_input.setObjectName(f"lon_input_{i-1}")
            point_layout.addWidget(lon_input)
            distance_layout.addLayout(point_layout)

        calc_button = QPushButton("Calculate Distance")
        calc_button.clicked.connect(self.calculate_distance)
        distance_layout.addWidget(calc_button)

        self.distance_output = QTextEdit()
        self.distance_output.setReadOnly(True)
        distance_layout.addWidget(self.distance_output)

        layout.addLayout(distance_layout)

        # Separator
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(separator)

        # Coordinate Converter Section
        coord_layout = QVBoxLayout()
        coord_layout.addWidget(QLabel("Coordinate Converter"))

        input_layout = QHBoxLayout()
        self.coord_input = QLineEdit()
        self.coord_input.setPlaceholderText("Enter coordinates (e.g., 40.7128, -74.0060)")
        input_layout.addWidget(self.coord_input)

        self.coord_format = QComboBox()
        self.coord_format.addItems(["Decimal Degrees", "Degrees Minutes Seconds"])
        input_layout.addWidget(self.coord_format)

        coord_layout.addLayout(input_layout)

        convert_button = QPushButton("Convert Coordinates")
        convert_button.clicked.connect(self.convert_coordinates)
        coord_layout.addWidget(convert_button)

        self.coord_output = QTextEdit()
        self.coord_output.setReadOnly(True)
        coord_layout.addWidget(self.coord_output)

        layout.addLayout(coord_layout)

        return widget

    def create_file_converter(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        self.conversion_file_list = QListWidget()
        layout.addWidget(self.conversion_file_list)  

        file_buttons_layout = QHBoxLayout()
        select_files_btn = QPushButton("Select Files")
        select_files_btn.clicked.connect(self.select_conversion_files)
        file_buttons_layout.addWidget(select_files_btn)

        clear_files_btn = QPushButton("Clear Files")
        clear_files_btn.clicked.connect(self.clear_conversion_files)
        file_buttons_layout.addWidget(clear_files_btn)

        layout.addLayout(file_buttons_layout)

        self.conversion_type = QComboBox()
        self.conversion_type.addItems(["CSV to KML", "KML to CSV"])
        layout.addWidget(self.conversion_type)

        convert_btn = QPushButton("Convert")
        convert_btn.clicked.connect(self.convert_files)
        layout.addWidget(convert_btn)

        return widget    

    def create_combined_lookup(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Shared input for coordinates
        input_layout = QHBoxLayout()
        self.combined_input = QLineEdit()
        self.combined_input.setPlaceholderText("Enter latitude, longitude (you may paste the coordinate here)")
        input_layout.addWidget(self.combined_input)

        lookup_button = QPushButton("Lookup")
        lookup_button.clicked.connect(self.perform_combined_lookup)
        input_layout.addWidget(lookup_button)

        layout.addLayout(input_layout)

        # Address Geocoder output
        layout.addWidget(QLabel("Address Geocoder:"))
        self.geocode_output = QTextEdit()
        self.geocode_output.setReadOnly(True)
        layout.addWidget(self.geocode_output)

        # Time Zone Finder output
        layout.addWidget(QLabel("Time Zone Finder:"))
        self.timezone_output = QTextEdit()
        self.timezone_output.setReadOnly(True)
        layout.addWidget(self.timezone_output)

        # Elevation Lookup output
        layout.addWidget(QLabel("Elevation Lookup:"))
        self.elevation_output = QTextEdit()
        self.elevation_output.setReadOnly(True)
        layout.addWidget(self.elevation_output)

        return widget

    def select_conversion_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Files",
            "",
            "Supported Files (*.csv *.xls *.xlsx *.kml);;CSV Files (*.csv);;Excel Files (*.xls *.xlsx);;KML Files (*.kml);;All Files (*.*)"
        )
        self.selected_conversion_files.extend(files)
        self.conversion_file_list.clear()
        self.conversion_file_list.addItems([os.path.basename(f) for f in files])

    def clear_text_files(self):
        self.selected_text_files.clear()
        self.file_list.clear()

    def clear_conversion_files(self):
        self.selected_conversion_files.clear()
        self.conversion_file_list.clear()

    def process_files(self):
        if not self.selected_text_files:
            QMessageBox.warning(self, "No Files Selected", "Please select files to process.")
            return

        operation = self.process_combo.currentText()
        processed_count = 0

        for file_path in self.selected_text_files:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()

                if operation == "Lowercase":
                    processed_content = content.lower()
                elif operation == "Uppercase":
                    processed_content = content.upper()
                elif operation == "Proper Case":
                    processed_content = ' '.join(word.capitalize() for word in content.split())

                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(processed_content)

                processed_count += 1

            except Exception as e:
                QMessageBox.warning(self, "Processing Error", f"Error processing {os.path.basename(file_path)}: {str(e)}")

        QMessageBox.information(self, "Batch Processing Complete", f"Successfully processed {processed_count} out of {len(self.selected_text_files)} files.")

    def calculate_distance(self):
        coordinates = []
        for i in range(5):  # We have 5 point inputs
            lat_input = self.findChild(QLineEdit, f"lat_input_{i}")
            lon_input = self.findChild(QLineEdit, f"lon_input_{i}")
            if lat_input and lon_input and lat_input.text() and lon_input.text():
                try:
                    lat = float(lat_input.text())
                    lon = float(lon_input.text())
                    coordinates.append((lat, lon))
                except ValueError:
                    self.distance_output.setPlainText("Invalid input. Please enter valid numbers for latitude and longitude.")
                    return

        if len(coordinates) < 2:
            self.distance_output.setPlainText("Please enter at least two points to calculate distance.")
            return

        total_distance = 0
        path_description = "Path:\n"
        for i in range(len(coordinates) - 1):
            point1 = coordinates[i]
            point2 = coordinates[i + 1]
            distance = self.haversine_distance(point1[0], point1[1], point2[0], point2[1])
            total_distance += distance
            path_description += f"Point {i+1} to Point {i+2}: {distance:.2f} km\n"

        result = f"{path_description}\nTotal Distance: {total_distance:.2f} km"
        self.distance_output.setPlainText(result)

    def haversine_distance(self, lat1, lon1, lat2, lon2):
        lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
        dlat = lat2 - lat1 
        dlon = lon2 - lon1 
        a = math.sin(dlat/2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2)**2
        c = 2 * math.asin(math.sqrt(a)) 
        r = 6371 # Radius of earth in kilometers
        return c * r

    def convert_files(self):
        conversion_type = self.conversion_type.currentText()
        input_files = [self.conversion_file_list.item(i).text() for i in range(self.conversion_file_list.count())]
        
        if not input_files:
            QMessageBox.warning(self, "No Files", "Please select files to convert.")
            return

        output_folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not output_folder:
            return

        for file in input_files:
            input_path = os.path.join(os.path.dirname(self.selected_conversion_files[0]), file)
            file_extension = os.path.splitext(file)[1].lower()

            if conversion_type == "To KML":
                if file_extension in ['.csv', '.xls', '.xlsx']:
                    coordinates = self.parse_spreadsheet(input_path)
                    output_file = os.path.join(output_folder, f"{os.path.splitext(file)[0]}.kml")
                    self.create_kml(coordinates, output_file)
                else:
                    QMessageBox.warning(self, "Invalid File", f"{file} is not a CSV, XLS, or XLSX file.")
            elif conversion_type == "From KML":
                if file_extension == '.kml':
                    coordinates = self.parse_kml(input_path)
                    csv_output = os.path.join(output_folder, f"{os.path.splitext(file)[0]}.csv")
                    xlsx_output = os.path.join(output_folder, f"{os.path.splitext(file)[0]}.xlsx")
                    self.create_csv(coordinates, csv_output)
                    self.create_xlsx(coordinates, xlsx_output)
                else:
                    QMessageBox.warning(self, "Invalid File", f"{file} is not a KML file.")
            else:
                QMessageBox.warning(self, "Invalid Conversion", "Please select a valid conversion type.")

        QMessageBox.information(self, "Conversion Complete", "File conversion completed.")

    def parse_spreadsheet(self, file_path):
        if file_path.lower().endswith('.csv'):
            df = pd.read_csv(file_path)
        else:  # .xls or .xlsx
            df = pd.read_excel(file_path)
        
        coordinates = []
        for _, row in df.iterrows():
            lat, lon = float(row[0]), float(row[1])
            coordinates.append((lon, lat))
        return coordinates

    def parse_kml(self, file_path):
        coordinates = []
        tree = etree.parse(file_path)
        root = tree.getroot()
        
        # KML namespace
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        
        # Find all Placemark elements
        for placemark in root.findall('.//kml:Placemark', namespaces=ns):
            coord_elem = placemark.find('.//kml:coordinates', namespaces=ns)
            if coord_elem is not None:
                coord_text = coord_elem.text.strip().split(',')
                lon, lat = float(coord_text[0]), float(coord_text[1])
                coordinates.append((lat, lon))
        
        return coordinates

    def create_kml(self, coordinates, output_file):
        kml = simplekml.Kml()
        for lon, lat in coordinates:
            kml.newpoint(coords=[(lon, lat)])
        kml.save(output_file)

    def create_csv(self, coordinates, output_file):
        with open(output_file, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Latitude', 'Longitude'])
            for lat, lon in coordinates:
                writer.writerow([lat, lon])

    def create_xlsx(self, coordinates, output_file):
        df = pd.DataFrame(coordinates, columns=['Latitude', 'Longitude'])
        df.to_excel(output_file, index=False)

    def convert_coordinates(self):
        input_coords = self.coord_input.text().strip()
        output_format = self.coord_format.currentText()
        
        try:
            lat, lon = self.parse_input_coordinates(input_coords)
            
            if output_format == "Decimal Degrees":
                self.coord_output.setPlainText(f"Latitude: {lat:.6f}\nLongitude: {lon:.6f}")
            else:  # Degrees Minutes Seconds
                dms_lat = self.decimal_to_dms(lat)
                dms_lon = self.decimal_to_dms(lon)
                self.coord_output.setPlainText(f"Latitude: {dms_lat}\nLongitude: {dms_lon}")
        except ValueError as e:
            self.coord_output.setPlainText(str(e))

    def parse_input_coordinates(self, input_coords):
        # Check if input is in DMS format
        dms_pattern = r'(\d+)°\s*(\d+)\'?\s*(\d+(\.\d+)?)"?\s*([NSEW])'
        dms_matches = re.findall(dms_pattern, input_coords, re.IGNORECASE)
        
        if dms_matches and len(dms_matches) == 2:
            lat = self.dms_to_decimal(dms_matches[0])
            lon = self.dms_to_decimal(dms_matches[1])
        else:
            # Assume decimal degrees
            parts = re.split(r'[,\s]+', input_coords)
            if len(parts) != 2:
                raise ValueError("Invalid input. Please enter coordinates as 'latitude, longitude' or in DMS format.")
            lat, lon = map(float, parts)
        
        # Validate latitude and longitude ranges
        if not -90 <= lat <= 90:
            raise ValueError("Invalid latitude. Must be between -90 and 90 degrees.")
        if not -180 <= lon <= 180:
            raise ValueError("Invalid longitude. Must be between -180 and 180 degrees.")
        
        return lat, lon

    def decimal_to_dms(self, decimal):
        is_positive = decimal >= 0
        decimal = abs(decimal)
        degrees = int(decimal)
        minutes = int((decimal - degrees) * 60)
        seconds = ((decimal - degrees) * 60 - minutes) * 60
        direction = 'N' if is_positive else 'S' if degrees != 0 else ''  # For latitude
        if not direction:
            direction = 'E' if is_positive else 'W' if degrees != 0 else ''  # For longitude
        return f"{degrees}° {minutes}' {seconds:.2f}\" {direction}"

    def dms_to_decimal(self, dms):
        degrees, minutes, seconds, _, direction = dms
        decimal = float(degrees) + float(minutes)/60 + float(seconds)/3600
        if direction.upper() in ['S', 'W']:
            decimal = -decimal
        return decimal
    
    def perform_combined_lookup(self):
        coords = self.combined_input.text().split(',')
        if len(coords) == 2:
            try:
                lat, lon = float(coords[0]), float(coords[1])
                self.geocode_coordinates(lat, lon)
                self.find_timezone(lat, lon)
                self.lookup_elevation(lat, lon)
            except ValueError:
                self.show_error("Invalid coordinates. Please enter as latitude,longitude.")
        else:
            self.show_error("Invalid input. Please enter coordinates as latitude,longitude.")

    def geocode_coordinates(self, lat, lon):
        geolocator = Nominatim(user_agent="MergedTextUtility")
        try:
            location = geolocator.reverse(f"{lat}, {lon}")
            if location:
                self.geocode_output.setPlainText(f"Address: {location.address}")
            else:
                self.geocode_output.setPlainText("Unable to find address for the given coordinates.")
        except Exception as e:
            self.geocode_output.setPlainText(f"An error occurred: {str(e)}")

    def find_timezone(self, lat, lon):
        tf = TimezoneFinder()
        timezone_str = tf.timezone_at(lat=lat, lng=lon)
        if timezone_str:
            timezone = pytz.timezone(timezone_str)
            current_time = datetime.now(timezone)
            self.timezone_output.setPlainText(f"Time Zone: {timezone_str}\nCurrent Time: {current_time.strftime('%Y-%m-%d %H:%M:%S %Z')}")
        else:
            self.timezone_output.setPlainText("Unable to determine timezone for the given coordinates.")

    def lookup_elevation(self, lat, lon):
        # Note: This is still a placeholder. In a real application, you would use an elevation API.
        self.elevation_output.setPlainText(f"Elevation for {lat}, {lon}: 100 meters (placeholder)")

    def show_error(self, message):
        self.geocode_output.setPlainText(message)
        self.timezone_output.setPlainText(message)
        self.elevation_output.setPlainText(message)

    def create_footer(self):
        footer_frame = QFrame()
        footer_frame.setStyleSheet("QFrame { background-color: #252525; border-radius: 0px; }")
        footer_layout = QHBoxLayout(footer_frame)

        self.uptime_label = QLabel("Uptime: 00:00:00")
        footer_layout.addWidget(self.uptime_label, alignment=Qt.AlignmentFlag.AlignLeft)

        footer_message = QLabel("Written in Python by Fadzli Abdullah")
        footer_layout.addWidget(footer_message, alignment=Qt.AlignmentFlag.AlignCenter)

        self.datetime_label = QLabel()
        footer_layout.addWidget(self.datetime_label, alignment=Qt.AlignmentFlag.AlignRight)

        self.layout.addWidget(footer_frame)

    def update_uptime(self):
        current_time = QDateTime.currentDateTime()
        uptime = self.start_time.secsTo(current_time)
        hours, remainder = divmod(uptime, 3600)
        minutes, seconds = divmod(remainder, 60)
        self.uptime_label.setText(f"Uptime: {hours:02d}:{minutes:02d}:{seconds:02d}")

    def update_datetime(self):
        current_datetime = QDateTime.currentDateTime()
        formatted_datetime = current_datetime.toString("yyyy-MM-dd hh:mm:ss")
        self.datetime_label.setText(formatted_datetime)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MergedTextUtility()
    window.show()
    sys.exit(app.exec())