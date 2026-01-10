import sys
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QGridLayout, QLabel, QLineEdit, QPushButton, QFrame, 
                             QMessageBox, QFileDialog, QSpacerItem, QSizePolicy, QDialog,
                             QComboBox, QSpinBox, QGroupBox)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QIcon, QFont
from geopy.distance import geodesic
import simplekml
from datetime import datetime
import psutil

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.layout = QVBoxLayout(self)

        # System Font Settings
        self.system_group = QGroupBox("System Font")
        system_layout = QVBoxLayout()
        self.system_font_style = QComboBox(self)
        self.system_font_style.addItems(["Andale Mono", "Helvetica", "Roboto", "Eras ITC", "Fixedsys"])
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
        self.content_font_style.addItems(["Andale Mono", "Helvetica", "Roboto", "Eras ITC", "Fixedsys"])
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

class DistanceCalculator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Distance Calculator")
        self.setFixedSize(650, 450)
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
        self.header = QLabel("Calculate The Distance Based on Coordinates Latlong")
        self.header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.header.setStyleSheet("background-color: #2c3e50; color: white; padding: 5px; font-weight: bold;")
        self.header.setFont(QFont("Arial", 12))  # Fixed font style to Arial
        self.layout.addWidget(self.header)
        
        # Input Grid
        self.input_frame = QFrame()
        self.input_frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.input_frame.setStyleSheet("background-color: white; color: black;")
        self.input_layout = QGridLayout(self.input_frame)
        self.input_layout.setVerticalSpacing(5)
        self.input_layout.setHorizontalSpacing(5)
        self.input_layout.setContentsMargins(5, 5, 5, 5)
        
        # Headers
        self.lat_label = QLabel("Latitude")
        self.lat_label.setFont(QFont(self.content_font_style, self.content_font_size - 1))
        self.input_layout.addWidget(self.lat_label, 0, 1, alignment=Qt.AlignmentFlag.AlignCenter)
        self.lon_label = QLabel("Longitude")
        self.lon_label.setFont(QFont(self.content_font_style, self.content_font_size - 1))
        self.input_layout.addWidget(self.lon_label, 0, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        
        self.coord_inputs = []
        self.point_labels = []
        for i in range(5):
            point_label = QLabel(f"Point {i+1}")
            point_label.setFont(QFont(self.content_font_style, self.content_font_size - 1))
            point_label.setFixedWidth(50)  # Set a fixed width for point labels
            self.input_layout.addWidget(point_label, i+1, 0)
            self.point_labels.append(point_label)
            for j in range(2):
                line_edit = QLineEdit()
                line_edit.setFixedWidth(100)  # Further reduced width
                line_edit.setFont(QFont(self.system_font_style, self.system_font_size - 1))
                self.input_layout.addWidget(line_edit, i+1, j+1)
                self.coord_inputs.append(line_edit)
        
        self.info_label = QLabel("* Point 1 and Point 2 are mandatory")
        self.info_label.setStyleSheet("color: #4CAF50; font-size: 8px; font-style: italic;")
        self.input_layout.addWidget(self.info_label, 6, 0, 1, 3)
        
        self.layout.addWidget(self.input_frame)
        
        # Calculate Button
        self.calc_button = QPushButton("Calculate")
        self.calc_button.setStyleSheet("background-color: #008CBA; color: white; padding: 5px; font-weight: bold;")
        self.calc_button.setFont(QFont(self.system_font_style, self.system_font_size - 1))
        self.calc_button.clicked.connect(self.calculate_distance)
        self.layout.addWidget(self.calc_button)
        
        # Result Display
        self.result_frame = QFrame()
        self.result_frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.result_frame.setStyleSheet("background-color: #2C3E50; color: white; padding: 5px;")
        self.result_layout = QHBoxLayout(self.result_frame)
        self.result_layout.setSpacing(10)
        
        self.result_labels = {}
        self.unit_labels = {}
        for unit in ["meter", "kilometer", "mile", "nautical mile"]:
            label_layout = QVBoxLayout()
            unit_label = QLabel(f"{unit}:")
            unit_label.setFont(QFont(self.content_font_style, self.content_font_size - 2))
            label_layout.addWidget(unit_label)
            self.unit_labels[unit] = unit_label
            self.result_labels[unit] = QLabel("0.00")
            self.result_labels[unit].setStyleSheet("font-weight: bold;")
            self.result_labels[unit].setFont(QFont(self.system_font_style, self.system_font_size - 1))
            label_layout.addWidget(self.result_labels[unit])
            self.result_layout.addLayout(label_layout)
        
        self.layout.addWidget(self.result_frame)
        
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
        
        self.layout.addLayout(button_layout)
        
        # Footer
        footer_frame = QFrame()
        footer_frame.setStyleSheet("background-color: #2c3e50; color: white;")
        footer_layout = QGridLayout(footer_frame)
        footer_layout.setContentsMargins(10, 1, 10, 1)  # Reduced vertical padding
        footer_layout.setSpacing(0)

        # CPU and RAM usage
        self.cpu_ram_label = QLabel()
        self.cpu_ram_label.setStyleSheet("font-size: 10px;")
        self.cpu_ram_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        footer_layout.addWidget(self.cpu_ram_label, 0, 0)

        # Footer message
        self.footer_label = QLabel("Written in Python by Fadzli Abdullah")
        self.footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        self.footer_label.setStyleSheet("font-size: 10px;")
        footer_layout.addWidget(self.footer_label, 0, 1)

        # Time label
        self.time_label = QLabel()
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.time_label.setStyleSheet("font-size: 10px;")
        footer_layout.addWidget(self.time_label, 0, 2)

        # Set column stretches to ensure proper alignment
        footer_layout.setColumnStretch(0, 1)  # Left element (CPU/RAM)
        footer_layout.setColumnStretch(1, 1)  # Center element (footer message)
        footer_layout.setColumnStretch(2, 1)  # Right element (time)

        # Set a fixed height for the footer frame
        footer_frame.setFixedHeight(20)  # Adjust this value as needed

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
            
            # Convert to different units and update result labels with units
            self.result_labels["m"].setText(f"{total_distance_m:.2f} m")
            self.result_labels["km"].setText(f"{total_distance_m / 1000:.2f} km")
            self.result_labels["mile"].setText(f"{total_distance_m / 1609.344:.2f} mi")
            self.result_labels["nautical mile"].setText(f"{total_distance_m / 1852:.2f} nm")
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

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # For a more modern look across platforms
    window = DistanceCalculator()
    window.show()
    sys.exit(app.exec())