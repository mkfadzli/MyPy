import sys
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QGridLayout, QLabel, QLineEdit, QPushButton, QFrame, 
                             QMessageBox, QFileDialog, QSpacerItem, QSizePolicy, QDialog,
                             QComboBox, QSpinBox, QGroupBox, QTextEdit, QProgressBar, QTabWidget)
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QFont, QGuiApplication
from geopy.distance import geodesic
import simplekml
from datetime import datetime
import psutil
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os
import subprocess

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.layout = QVBoxLayout(self)

        # System Font Settings
        self.system_group = QGroupBox("System Font")
        system_layout = QVBoxLayout()
        self.system_font_style = QComboBox(self)
        self.system_font_style.addItems(["Andale Mono", "Aptos Mono", "Roboto", "Eras ITC", "Fixedsys","Britannic Bold"])
        system_layout.addWidget(QLabel("Font Style:"))
        system_layout.addWidget(self.system_font_style)
        self.system_font_size = QSpinBox(self)
        self.system_font_size.setRange(8, 24)
        system_layout.addWidget(QLabel("Font Size:"))
        system_layout.addWidget(self.system_font_size)
        self.system_group.setLayout(system_layout)
        self.layout.addWidget(self.system_group)

        # Content Font Settings
        self.content_group = QGroupBox("Content Font")
        content_layout = QVBoxLayout()
        self.content_font_style = QComboBox(self)
        self.content_font_style.addItems(["Andale Mono", "Aptos Mono", "Roboto", "Eras ITC", "Fixedsys", "Britannic Bold"])
        content_layout.addWidget(QLabel("Font Style:"))
        content_layout.addWidget(self.content_font_style)
        self.content_font_size = QSpinBox(self)
        self.content_font_size.setRange(8, 24)
        content_layout.addWidget(QLabel("Font Size:"))
        content_layout.addWidget(self.content_font_size)
        self.content_group.setLayout(content_layout)
        self.layout.addWidget(self.content_group)

        self.save_button = QPushButton("Save", self)
        self.save_button.clicked.connect(self.accept)
        self.layout.addWidget(self.save_button)

class ConversionThread(QThread):
    update_progress = pyqtSignal(int)
    update_log = pyqtSignal(str)
    conversion_done = pyqtSignal(bool, str)

    def __init__(self, input_file, output_format):
        QThread.__init__(self)
        self.input_file = input_file
        self.output_format = output_format

    def run(self):
        try:
            input_extension = os.path.splitext(self.input_file)[1].lower()
            
            if input_extension in ['.csv', '.xlsx']:
                self.convert_to_kml()
            elif input_extension == '.kml':
                self.convert_from_kml()
            else:
                raise ValueError("Unsupported file type. Please use CSV, XLSX, or KML files.")
            
        except Exception as e:
            self.conversion_done.emit(False, f"An error occurred: {str(e)}")

    def convert_to_kml(self):
        kml_file = os.path.splitext(self.input_file)[0] + '.kml'
        file_extension = os.path.splitext(self.input_file)[1].lower()
        
        if file_extension == '.csv':
            df = pd.read_csv(self.input_file)
        elif file_extension == '.xlsx':
            df = pd.read_excel(self.input_file)

        self.update_log.emit(f"Columns in the file: {df.columns.tolist()}")

        required_columns = ['latitude', 'longitude']
        df.columns = df.columns.str.lower()
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise KeyError(f"Missing required columns: {', '.join(missing_columns)}")

        df = df.rename(columns={'latitude': 'latitude', 'longitude': 'longitude'})

        self.dataframe_to_kml(df, kml_file)
        self.conversion_done.emit(True, f"KML file '{kml_file}' has been created successfully.")

    def convert_from_kml(self):
        output_file = os.path.splitext(self.input_file)[0] + f'.{self.output_format}'
        
        tree = ET.parse(self.input_file)
        root = tree.getroot()

        data = []
        total_placemarks = len(root.findall('.//{http://www.opengis.net/kml/2.2}Placemark'))
        
        for i, placemark in enumerate(root.findall('.//{http://www.opengis.net/kml/2.2}Placemark')):
            point_data = {}
            name = placemark.find('{http://www.opengis.net/kml/2.2}name')
            if name is not None:
                point_data['name'] = name.text
            
            coordinates = placemark.find('.//{http://www.opengis.net/kml/2.2}coordinates')
            if coordinates is not None:
                lon, lat, _ = coordinates.text.split(',')
                point_data['longitude'] = float(lon)
                point_data['latitude'] = float(lat)
            
            description = placemark.find('{http://www.opengis.net/kml/2.2}description')
            if description is not None:
                for item in description.text.split(','):
                    key, value = item.strip().split(':')
                    point_data[key.strip()] = value.strip()
            
            data.append(point_data)
            
            progress = int((i + 1) / total_placemarks * 100)
            self.update_progress.emit(progress)

        df = pd.DataFrame(data)
        
        if self.output_format == 'csv':
            df.to_csv(output_file, index=False)
        elif self.output_format == 'xlsx':
            df.to_excel(output_file, index=False)
        
        self.conversion_done.emit(True, f"{self.output_format.upper()} file '{output_file}' has been created successfully.")

    def dataframe_to_kml(self, df, kml_file):
        kml = ET.Element('kml', xmlns="http://www.opengis.net/kml/2.2")
        document = ET.SubElement(kml, 'Document')

        style = ET.SubElement(document, 'Style', id="customStyle")
        icon_style = ET.SubElement(style, 'IconStyle')
        icon = ET.SubElement(icon_style, 'Icon')
        href = ET.SubElement(icon, 'href')
        href.text = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"

        total_rows = len(df)
        for index, row in df.iterrows():
            placemark = ET.SubElement(document, 'Placemark')
            
            name = ET.SubElement(placemark, 'name')
            name.text = f"{row['latitude']}, {row['longitude']}"
            
            description = ET.SubElement(placemark, 'description')
            description.text = ', '.join(f"{k}: {v}" for k, v in row.items() if k not in ['latitude', 'longitude'])
            
            style_url = ET.SubElement(placemark, 'styleUrl')
            style_url.text = "#customStyle"
            
            point = ET.SubElement(placemark, 'Point')
            coordinates = ET.SubElement(point, 'coordinates')
            coordinates.text = f"{row['longitude']},{row['latitude']},0"

            progress = int((index + 1) / total_rows * 100)
            self.update_progress.emit(progress)

        kml_string = ET.tostring(kml, encoding='unicode')
        pretty_kml = minidom.parseString(kml_string).toprettyxml(indent="  ")

        with open(kml_file, 'w', encoding='utf-8') as f:
            f.write(pretty_kml)

class MergedApplication(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Distance Calculator & CSV/KML Converter")
        self.setFixedSize(860, 600)
        self.setWindowIcon(QIcon("myicon.ico"))
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        
        self.load_settings()
        self.setup_ui()
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        self.update_time()

    def load_settings(self):
        try:
            with open('settings.json', 'r') as f:
                settings = json.load(f)
            self.system_font_style = settings.get('system_font_style', 'Andale Mono')
            self.system_font_size = settings.get('system_font_size', 12)
            self.content_font_style = settings.get('content_font_style', 'Andale Mono')
            self.content_font_size = settings.get('content_font_size', 12)
        except FileNotFoundError:
            self.system_font_style = 'Andale Mono'
            self.system_font_size = 12
            self.content_font_style = 'Andale Mono'
            self.content_font_size = 12

    def save_settings(self):
        settings = {
            'system_font_style': self.system_font_style,
            'system_font_size': self.system_font_size,
            'content_font_style': self.content_font_style,
            'content_font_size': self.content_font_size
        }
        with open('settings.json', 'w') as f:
            json.dump(settings, f)

    def setup_ui(self):
        # Header
        self.header = QLabel("Distance Calculator & CSV/KML Converter")
        self.header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.header.setStyleSheet("background-color: #2c3e50; color: white; padding: 10px; font-weight: bold; font-size: 24px; font-family: 'Cooper Black';")
        self.layout.addWidget(self.header)
        
        # Tab Widget
        self.tab_widget = QTabWidget()
        self.layout.addWidget(self.tab_widget)
        
        # Distance Calculator Tab
        self.distance_tab = QWidget()
        self.distance_layout = QVBoxLayout(self.distance_tab)
        self.setup_distance_calculator()
        self.tab_widget.addTab(self.distance_tab, "Distance Calculator")
        
        # File Converter Tab
        self.converter_tab = QWidget()
        self.converter_layout = QVBoxLayout(self.converter_tab)
        self.setup_file_converter()
        self.tab_widget.addTab(self.converter_tab, "File Converter")
        
        # Footer
        self.setup_footer()

    def setup_distance_calculator(self):
        # Input Grid
        self.input_frame = QFrame()
        self.input_frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.input_frame.setStyleSheet("background-color: white; color: black;")
        self.input_layout = QGridLayout(self.input_frame)
        self.input_layout.setVerticalSpacing(10)  # Increased vertical spacing
        self.input_layout.setHorizontalSpacing(10)  # Increased horizontal spacing
        self.input_layout.setContentsMargins(20, 20, 20, 20)  # Increased margins
        
        # Headers
        self.lat_label = QLabel("Latitude")
        self.lat_label.setFont(QFont(self.content_font_style, self.content_font_size))
        self.lat_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.input_layout.addWidget(self.lat_label, 0, 1)
        
        self.lon_label = QLabel("Longitude")
        self.lon_label.setFont(QFont(self.content_font_style, self.content_font_size))
        self.lon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.input_layout.addWidget(self.lon_label, 0, 2)
        
        self.coord_inputs = []
        self.point_labels = []
        for i in range(5):
            point_label = QLabel(f"Point {i+1}")
            point_label.setFont(QFont(self.content_font_style, self.content_font_size - 1))
            point_label.setFixedWidth(60)  # Set a fixed width for point labels
            self.input_layout.addWidget(point_label, i+1, 0)
            self.point_labels.append(point_label)
            
            for j in range(2):
                line_edit = QLineEdit()
                line_edit.setFont(QFont(self.system_font_style, self.system_font_size))
                line_edit.setMinimumWidth(200)  # Set a minimum width for input fields
                self.input_layout.addWidget(line_edit, i+1, j+1)
                self.coord_inputs.append(line_edit)
        
        # Set column stretch to make input fields expand
        self.input_layout.setColumnStretch(1, 1)
        self.input_layout.setColumnStretch(2, 1)
        
        self.info_label = QLabel("* Point 1 and Point 2 are mandatory")
        self.info_label.setStyleSheet("color: #4CAF50; font-size: 10px; font-style: italic;")
        self.input_layout.addWidget(self.info_label, 6, 0, 1, 3)
        
        self.distance_layout.addWidget(self.input_frame)
        
        # Calculate Button
        self.calc_button = QPushButton("Calculate")
        self.calc_button.setStyleSheet("background-color: #008CBA; color: white; padding: 5px; font-weight: bold;")
        self.calc_button.setFont(QFont(self.system_font_style, self.system_font_size - 1))
        self.calc_button.clicked.connect(self.calculate_distance)
        self.distance_layout.addWidget(self.calc_button)
        
        # Result Display
        self.result_frame = QFrame()
        self.result_frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.result_frame.setStyleSheet("background-color: #2C3E50; color: white; padding: 2px;")  # Reduced padding
        self.result_layout = QHBoxLayout(self.result_frame)
        self.result_layout.setSpacing(5)  # Reduced spacing
        self.result_layout.setContentsMargins(5, 2, 5, 2)  # Reduced margins
        
        self.result_labels = {}
        self.unit_labels = {}
        for unit in ["meter", "kilometer", "mile", "nautical mile"]:
            label_layout = QVBoxLayout()
            label_layout.setSpacing(0)  # Minimize spacing between unit and value
            unit_label = QLabel(f"{unit}:")
            unit_label.setFont(QFont(self.content_font_style, self.content_font_size - 3))  # Smaller font
            label_layout.addWidget(unit_label)
            self.unit_labels[unit] = unit_label
            self.result_labels[unit] = QLabel("0.00")
            self.result_labels[unit].setStyleSheet("font-weight: bold;")
            self.result_labels[unit].setFont(QFont(self.system_font_style, self.system_font_size - 2))  # Smaller font
            label_layout.addWidget(self.result_labels[unit])
            self.result_layout.addLayout(label_layout)
        
        # Set a fixed height for the result frame to reduce its size by approximately 5%
        self.result_frame.setFixedHeight(int(self.result_frame.sizeHint().height() * 0.95))
        
        self.distance_layout.addWidget(self.result_frame)
        
        # Export to KML and Settings Buttons
        button_layout = QHBoxLayout()
        self.export_button = QPushButton("Export to KML")
        self.export_button.setStyleSheet("background-color: #008CBA; color: white; padding: 5px; font-weight: bold;")
        self.export_button.setFont(QFont(self.system_font_style, self.system_font_size - 1))
        self.export_button.clicked.connect(self.export_to_kml)
        button_layout.addWidget(self.export_button)
        
        self.settings_button = QPushButton("Settings")
        self.settings_button.setStyleSheet("background-color: #008CBA; color: white; padding: 5px; font-weight: bold;")
        self.settings_button.setFont(QFont(self.system_font_style, self.system_font_size - 1))
        self.settings_button.clicked.connect(self.open_settings)
        button_layout.addWidget(self.settings_button)
        
        self.distance_layout.addLayout(button_layout)

    def setup_file_converter(self):
        file_layout = QHBoxLayout()
        self.select_file_button = QPushButton("Select File")
        self.select_file_button.clicked.connect(self.select_file)
        file_layout.addWidget(self.select_file_button)

        self.file_label = QLabel("No file selected")
        file_layout.addWidget(self.file_label)
        self.converter_layout.addLayout(file_layout)

        output_format_layout = QHBoxLayout()
        output_format_label = QLabel("Output Format:")
        self.output_format_combo = QComboBox()
        self.output_format_combo.addItems([".kml", ".csv", ".xlsx"])
        output_format_layout.addWidget(output_format_label)
        output_format_layout.addWidget(self.output_format_combo)
        self.converter_layout.addLayout(output_format_layout)

        self.progress_bar = QProgressBar()
        self.converter_layout.addWidget(self.progress_bar)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.converter_layout.addWidget(self.log_output)

        self.open_folder_button = QPushButton("Open Output Folder")
        self.open_folder_button.clicked.connect(self.open_output_folder)
        self.open_folder_button.setEnabled(False)
        self.converter_layout.addWidget(self.open_folder_button)

    def setup_footer(self):
        footer_frame = QFrame()
        footer_frame.setStyleSheet("background-color: #2c3e50; color: white;")
        footer_layout = QGridLayout(footer_frame)
        footer_layout.setContentsMargins(10, 1, 10, 1)
        footer_layout.setSpacing(0)

        self.cpu_ram_label = QLabel()
        self.cpu_ram_label.setStyleSheet("font-size: 10px;")
        self.cpu_ram_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        footer_layout.addWidget(self.cpu_ram_label, 0, 0)

        self.footer_label = QLabel("Written in Python by Fadzli Abdullah")
        self.footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        self.footer_label.setStyleSheet("font-size: 10px;")
        footer_layout.addWidget(self.footer_label, 0, 1)

        self.time_label = QLabel()
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.time_label.setStyleSheet("font-size: 10px;")
        footer_layout.addWidget(self.time_label, 0, 2)

        footer_layout.setColumnStretch(0, 1)
        footer_layout.setColumnStretch(1, 1)
        footer_layout.setColumnStretch(2, 1)

        footer_frame.setFixedHeight(20)
        self.layout.addWidget(footer_frame)

    def calculate_distance(self):
        try:
            points = []
            for i in range(0, len(self.coord_inputs), 2):
                lat = self.coord_inputs[i].text().strip()
                lon = self.coord_inputs[i+1].text().strip()
                if i < 4:  # Points 1 and 2 are mandatory
                    if not lat or not lon:
                        raise ValueError(f"Point {i//2 + 1} is mandatory and must be complete.")
                    points.append((float(lat), float(lon)))
                elif lat and lon:  # Optional points
                    points.append((float(lat), float(lon)))
                elif lat or lon:  # Incomplete optional point
                    raise ValueError(f"Point {i//2 + 1} is incomplete. Please provide both latitude and longitude.")
            
            if len(points) < 2:
                raise ValueError("At least two points (Point 1 and Point 2) are required.")
            
            total_distance_m = sum(geodesic(points[i], points[i+1]).meters for i in range(len(points)-1))
            
            # Convert to different units and update result labels
            self.result_labels["meter"].setText(f"{total_distance_m:.2f}")
            self.result_labels["kilometer"].setText(f"{total_distance_m / 1000:.2f}")
            self.result_labels["mile"].setText(f"{total_distance_m / 1609.344:.2f}")
            self.result_labels["nautical mile"].setText(f"{total_distance_m / 1852:.2f}")
        except ValueError as e:
            QMessageBox.warning(self, "Error", str(e))

    def export_to_kml(self):
        try:
            points = []
            for i in range(0, len(self.coord_inputs), 2):
                lat = self.coord_inputs[i].text().strip()
                lon = self.coord_inputs[i+1].text().strip()
                if lat and lon:
                    points.append((float(lat), float(lon)))
            
            if len(points) < 2:
                raise ValueError("At least two points are required to create a KML file.")
            
            kml = simplekml.Kml()
            for i, point in enumerate(points, 1):
                kml.newpoint(name=f"Point {i}", coords=[point[::-1]])  # KML uses (lon, lat) order
            
            file_path, _ = QFileDialog.getSaveFileName(self, "Save KML File", "", "KML files (*.kml)")
            if file_path:
                kml.save(file_path)
                QMessageBox.information(self, "Success", f"KML file saved to {file_path}")
        except ValueError as e:
            QMessageBox.warning(self, "Error", str(e))

    def update_time(self):
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        self.time_label.setText(current_time)

        # Update CPU and RAM usage
        cpu_usage = psutil.cpu_percent()
        ram_usage = psutil.virtual_memory().percent
        self.cpu_ram_label.setText(f"CPU: {cpu_usage:.2f}% | RAM: {ram_usage:.2f}%")

    def open_settings(self):
        dialog = SettingsDialog(self)
        dialog.system_font_style.setCurrentText(self.system_font_style)
        dialog.system_font_size.setValue(self.system_font_size)
        dialog.content_font_style.setCurrentText(self.content_font_style)
        dialog.content_font_size.setValue(self.content_font_size)
        
        if dialog.exec():
            self.system_font_style = dialog.system_font_style.currentText()
            self.system_font_size = dialog.system_font_size.value()
            self.content_font_style = dialog.content_font_style.currentText()
            self.content_font_size = dialog.content_font_size.value()
            self.save_settings()
            self.apply_settings()

    def apply_settings(self):
        system_font = QFont(self.system_font_style, self.system_font_size)
        content_font = QFont(self.content_font_style, self.content_font_size)

        # Update system fonts
        for input_field in self.coord_inputs:
            input_field.setFont(system_font)
        self.calc_button.setFont(system_font)
        self.export_button.setFont(system_font)
        self.settings_button.setFont(system_font)
        for label in self.result_labels.values():
            label.setFont(system_font)

        # Update content fonts
        self.header.setFont(QFont(self.content_font_style, 14))
        self.lat_label.setFont(content_font)
        self.lon_label.setFont(content_font)
        for point_label in self.point_labels:
            point_label.setFont(content_font)
        for unit_label in self.unit_labels.values():
            unit_label.setFont(content_font)

        # Refresh the layout
        self.central_widget.setLayout(self.layout)

    def select_file(self):
        file_dialog = QFileDialog()
        input_file, _ = file_dialog.getOpenFileName(self, "Select CSV, XLSX, or KML file", "", "All Supported Files (*.csv *.xlsx *.kml);;CSV files (*.csv);;Excel files (*.xlsx);;KML files (*.kml)")
        
        if input_file:
            self.file_label.setText(os.path.basename(input_file))
            self.log_output.clear()
            self.log_output.append(f"Selected file: {input_file}")
            self.output_folder = os.path.dirname(input_file)
            
            file_extension = os.path.splitext(input_file)[1].lower()
            if file_extension in ['.csv', '.xlsx']:
                self.output_format_combo.setCurrentText(".kml")
                self.output_format_combo.setEnabled(False)
            elif file_extension == '.kml':
                self.output_format_combo.setEnabled(True)
                self.output_format_combo.setCurrentIndex(1)  # Set to CSV by default
            
            self.start_conversion(input_file)

    def start_conversion(self, input_file):
        output_format = self.output_format_combo.currentText().lower().strip('.')
        self.conversion_thread = ConversionThread(input_file, output_format)
        self.conversion_thread.update_progress.connect(self.update_progress)
        self.conversion_thread.update_log.connect(self.update_log)
        self.conversion_thread.conversion_done.connect(self.conversion_done)
        self.conversion_thread.start()
        self.select_file_button.setEnabled(False)
        self.open_folder_button.setEnabled(False)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_log(self, message):
        self.log_output.append(message)

    def conversion_done(self, success, message):
        self.update_log(message)
        self.select_file_button.setEnabled(True)
        self.open_folder_button.setEnabled(True)
        self.progress_bar.setValue(100 if success else 0)

    def open_output_folder(self):
        if self.output_folder:
            if sys.platform == 'win32':
                os.startfile(self.output_folder)
            elif sys.platform == 'darwin':  # macOS
                subprocess.run(['open', self.output_folder])
            else:  # Linux and other Unix-like
                subprocess.run(['xdg-open', self.output_folder])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    icon_path = os.path.join(os.path.dirname(__file__), "myicon.ico")
    app.setWindowIcon(QIcon(icon_path))
    
    window = MergedApplication()
    window.show()
    sys.exit(app.exec())
