import sys
import os
import zipfile
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                              QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                              QTextEdit, QGroupBox, QSpinBox, QDoubleSpinBox,
                              QCheckBox, QMessageBox, QSplitter, QLineEdit)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont

import geopandas as gpd
from shapely.geometry import Polygon, MultiPolygon, Point
from shapely.ops import unary_union
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


class ConversionThread(QThread):
    """Thread for handling file conversion operations"""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)
    
    def __init__(self, input_file, output_dir, formats, simplify, tolerance, max_points, buffer_width=40, output_name=None):
        super().__init__()
        self.input_file = input_file
        self.output_dir = output_dir
        self.formats = formats
        self.simplify = simplify
        self.tolerance = tolerance
        self.max_points = max_points
        self.buffer_width = buffer_width  # Now in meters
        self.output_name = output_name
        
    def run(self):
        try:
            # Load KMZ/KML
            self.progress.emit("Reading KMZ/KML file...")
            gdf = self.load_kmz_kml(self.input_file, self.buffer_width)
            
            if gdf is None or gdf.empty:
                self.finished.emit(False, "No valid geometries found in file")
                return
            
            # Check point count and simplify if needed
            original_points = sum([len(geom.exterior.coords) if isinstance(geom, Polygon) else 
                                 sum([len(p.exterior.coords) for p in geom.geoms]) 
                                 for geom in gdf.geometry])
            
            self.progress.emit(f"Original polygon has {original_points} points")
            
            if self.simplify or original_points > self.max_points:
                self.progress.emit(f"Simplifying polygon (max points: {self.max_points})...")
                gdf = self.simplify_geometries(gdf, self.tolerance, self.max_points)
                
                new_points = sum([len(geom.exterior.coords) if isinstance(geom, Polygon) else 
                                sum([len(p.exterior.coords) for p in geom.geoms]) 
                                for geom in gdf.geometry])
                
                self.progress.emit(f"Simplified to {new_points} points")
            
            # Export to selected formats
            # Determine output base name
            if self.output_name and self.output_name.strip():
                base_name = self.output_name.strip()
                # Remove any extensions if user added them
                base_name = base_name.replace('.shp', '').replace('.geojson', '').replace('.tab', '')
                self.progress.emit(f"Using custom filename: {base_name}")
            else:
                base_name = Path(self.input_file).stem
                self.progress.emit(f"Using original filename: {base_name}")
            
            output_files = []
            
            if 'shapefile' in self.formats:
                self.progress.emit("Converting to Shapefile...")
                shp_path = os.path.join(self.output_dir, f"{base_name}.shp")
                gdf.to_file(shp_path, driver='ESRI Shapefile')
                output_files.append(shp_path)
            
            if 'geojson' in self.formats:
                self.progress.emit("Converting to GeoJSON...")
                geojson_path = os.path.join(self.output_dir, f"{base_name}.geojson")
                gdf.to_file(geojson_path, driver='GeoJSON')
                output_files.append(geojson_path)
            
            if 'tab' in self.formats:
                self.progress.emit("Converting to MapInfo TAB...")
                tab_path = os.path.join(self.output_dir, f"{base_name}.tab")
                try:
                    gdf.to_file(tab_path, driver='MapInfo File')
                    output_files.append(tab_path)
                except Exception as e:
                    self.progress.emit(f"Warning: TAB export failed - {str(e)}")
            
            success_msg = f"Conversion successful!\n\nFiles created:\n" + "\n".join(output_files)
            self.finished.emit(True, success_msg)
            
        except Exception as e:
            self.finished.emit(False, f"Conversion failed: {str(e)}")
    
    def load_kmz_kml(self, file_path, buffer_width=40):
        """Load KMZ or KML file and convert to polygons
        
        Args:
            file_path: Path to KMZ or KML file
            buffer_width: Buffer width in METERS for LineString conversion
        """
        try:
            # Handle KMZ (zipped KML)
            if file_path.lower().endswith('.kmz'):
                with tempfile.TemporaryDirectory() as temp_dir:
                    with zipfile.ZipFile(file_path, 'r') as kmz:
                        kmz.extractall(temp_dir)
                        # Find KML file in extracted contents
                        kml_files = list(Path(temp_dir).rglob('*.kml'))
                        if not kml_files:
                            raise Exception("No KML file found inside KMZ")
                        kml_path = str(kml_files[0])
                        
                        # Parse KML inside the temp directory context
                        return self._parse_kml_file(kml_path, buffer_width)
            else:
                # Direct KML file
                return self._parse_kml_file(file_path, buffer_width)
            
        except Exception as e:
            raise Exception(f"Failed to load KMZ/KML: {str(e)}")
    
    def _parse_kml_file(self, kml_file_path, buffer_width=40):
        """Parse KML file and return GeoDataFrame
        
        Args:
            kml_file_path: Path to KML file
            buffer_width: Buffer width in METERS for LineString conversion
        """
        try:
            # Parse KML
            tree = ET.parse(kml_file_path)
            root = tree.getroot()
            
            # Handle KML namespace
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}
            if not root.tag.startswith('{'):
                ns = {}
            
            geometries = []
            names = []
            geom_types = []
            
            # Extract all geometry types
            placemarks = root.findall('.//kml:Placemark', ns)
            if not placemarks:
                placemarks = root.findall('.//Placemark')
            
            for placemark in placemarks:
                # Get name
                name_elem = placemark.find('.//kml:name', ns)
                if name_elem is None:
                    name_elem = placemark.find('.//name')
                name = name_elem.text if name_elem is not None else "Unnamed"
                
                # Look for Polygon
                polygon_elem = placemark.find('.//kml:Polygon', ns)
                if polygon_elem is None:
                    polygon_elem = placemark.find('.//Polygon')
                    
                if polygon_elem is not None:
                    coords_elem = polygon_elem.find('.//kml:coordinates', ns)
                    if coords_elem is None:
                        coords_elem = polygon_elem.find('.//coordinates')
                        
                    if coords_elem is not None:
                        coords_text = coords_elem.text.strip()
                        coords = self.parse_coordinates(coords_text)
                        if len(coords) >= 3:
                            poly = Polygon(coords)
                            geometries.append(poly)
                            names.append(name)
                            geom_types.append('Polygon')
                            continue
                
                # Look for LineString (route paths)
                linestring_elem = placemark.find('.//kml:LineString', ns)
                if linestring_elem is None:
                    linestring_elem = placemark.find('.//LineString')
                    
                if linestring_elem is not None:
                    coords_elem = linestring_elem.find('.//kml:coordinates', ns)
                    if coords_elem is None:
                        coords_elem = linestring_elem.find('.//coordinates')
                        
                    if coords_elem is not None:
                        coords_text = coords_elem.text.strip()
                        coords = self.parse_coordinates(coords_text)
                        if len(coords) >= 2:
                            from shapely.geometry import LineString
                            import pyproj
                            from pyproj import Transformer
                            
                            line = LineString(coords)
                            
                            # buffer_width is now in meters
                            buffer_meters = buffer_width
                            
                            # Get the centroid to determine appropriate UTM zone
                            centroid = line.centroid
                            lon, lat = centroid.x, centroid.y
                            
                            # Calculate UTM zone
                            utm_zone = int((lon + 180) / 6) + 1
                            utm_crs = f"EPSG:326{utm_zone:02d}" if lat >= 0 else f"EPSG:327{utm_zone:02d}"
                            
                            try:
                                # Transform to UTM for accurate meter-based buffering
                                transformer_to_utm = Transformer.from_crs("EPSG:4326", utm_crs, always_xy=True)
                                transformer_to_wgs = Transformer.from_crs(utm_crs, "EPSG:4326", always_xy=True)
                                
                                # Transform line to UTM
                                from shapely.ops import transform
                                line_utm = transform(transformer_to_utm.transform, line)
                                
                                # Buffer in meters
                                self.progress.emit(f"  Buffering LineString by {buffer_meters}m using {utm_crs}")
                                buffered_utm = line_utm.buffer(buffer_meters)
                                
                                # Transform back to WGS84
                                buffered = transform(transformer_to_wgs.transform, buffered_utm)
                                
                            except Exception as e:
                                # Fallback to degree-based buffering if UTM fails
                                # Convert meters to approximate degrees (at equator)
                                buffer_degrees = buffer_meters / 111320
                                self.progress.emit(f"Warning: UTM transformation failed, using degree-based buffer ({buffer_degrees:.6f}°)")
                                buffered = line.buffer(buffer_degrees)
                            
                            geometries.append(buffered)
                            names.append(name)
                            geom_types.append('LineString (converted)')
                            continue
                
                # Look for Point
                point_elem = placemark.find('.//kml:Point', ns)
                if point_elem is None:
                    point_elem = placemark.find('.//Point')
                    
                if point_elem is not None:
                    geom_types.append('Point (skipped)')
            
            # Diagnostic information
            if not geometries:
                found_types = ', '.join(set(geom_types)) if geom_types else 'None'
                raise Exception(
                    f"No valid geometries found.\n\n"
                    f"Found geometry types: {found_types}\n\n"
                    f"Tip: If file has Points only, you need area/route data.\n"
                    f"If file has LineStrings, this version now supports them!"
                )
            
            # Log what was found
            self.progress.emit(f"Found {len(geometries)} geometries: {', '.join(geom_types)}")
            
            # Create GeoDataFrame
            gdf = gpd.GeoDataFrame(
                {'name': names, 'type': geom_types},
                geometry=geometries,
                crs='EPSG:4326'
            )
            
            return gdf
            
        except Exception as e:
            raise Exception(f"Failed to parse KML: {str(e)}")
    
    def parse_coordinates(self, coords_text):
        """Parse KML coordinates text"""
        coords = []
        for line in coords_text.split():
            parts = line.split(',')
            if len(parts) >= 2:
                lon, lat = float(parts[0]), float(parts[1])
                coords.append((lon, lat))
        return coords
    
    def simplify_geometries(self, gdf, tolerance, max_points):
        """Simplify geometries to reduce point count"""
        simplified = []
        
        for geom in gdf.geometry:
            current_tolerance = tolerance
            simplified_geom = geom.simplify(current_tolerance, preserve_topology=True)
            
            # Increase tolerance if still too many points
            while self.count_points(simplified_geom) > max_points and current_tolerance < 0.1:
                current_tolerance *= 1.5
                simplified_geom = geom.simplify(current_tolerance, preserve_topology=True)
            
            simplified.append(simplified_geom)
        
        gdf_simplified = gdf.copy()
        gdf_simplified.geometry = simplified
        return gdf_simplified
    
    def count_points(self, geom):
        """Count total points in geometry"""
        if isinstance(geom, Polygon):
            return len(geom.exterior.coords)
        elif isinstance(geom, MultiPolygon):
            return sum([len(p.exterior.coords) for p in geom.geoms])
        return 0


class MapCanvas(FigureCanvas):
    """Canvas for displaying polygon"""
    def __init__(self, parent=None):
        self.figure = Figure(figsize=(8, 6), facecolor='white')
        super().__init__(self.figure)
        self.setParent(parent)
        self.ax = self.figure.add_subplot(111)
        self.clear_plot()
    
    def clear_plot(self):
        """Clear the plot"""
        self.ax.clear()
        self.ax.set_title("Polygon Preview")
        self.ax.set_xlabel("Longitude")
        self.ax.set_ylabel("Latitude")
        self.ax.grid(True, alpha=0.3)
        self.draw()
    
    def plot_geometry(self, gdf):
        """Plot GeoDataFrame geometry"""
        self.ax.clear()
        
        try:
            gdf.plot(ax=self.ax, facecolor='lightblue', edgecolor='blue', 
                    alpha=0.5, linewidth=2)
            
            # Add labels
            for idx, row in gdf.iterrows():
                centroid = row.geometry.centroid
                self.ax.annotate(row['name'], xy=(centroid.x, centroid.y),
                               xytext=(3, 3), textcoords="offset points",
                               fontsize=9, color='darkblue')
            
            # Get bounds and add some padding
            minx, miny, maxx, maxy = gdf.total_bounds
            x_pad = (maxx - minx) * 0.1
            y_pad = (maxy - miny) * 0.1
            
            self.ax.set_xlim(minx - x_pad, maxx + x_pad)
            self.ax.set_ylim(miny - y_pad, maxy + y_pad)
            
            # Count total points
            total_points = sum([len(geom.exterior.coords) if isinstance(geom, Polygon) else 
                              sum([len(p.exterior.coords) for p in geom.geoms]) 
                              for geom in gdf.geometry])
            
            self.ax.set_title(f"Polygon Preview ({len(gdf)} features, {total_points} points)")
            self.ax.set_xlabel("Longitude")
            self.ax.set_ylabel("Latitude")
            self.ax.grid(True, alpha=0.3)
            
            self.draw()
            
        except Exception as e:
            self.ax.text(0.5, 0.5, f"Error plotting: {str(e)}", 
                        ha='center', va='center', transform=self.ax.transAxes)
            self.draw()


class PolygonConverterApp(QMainWindow):
    """Main application window"""
    def __init__(self):
        super().__init__()
        self.input_file = None
        self.gdf = None
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Route Polygon Converter - KMZ/KML to TAB/SHP/GeoJSON")
        self.setGeometry(100, 100, 1200, 800)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Title
        title = QLabel("Route Polygon Converter")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title)
        
        # Splitter for left/right panels
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left panel - Controls
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        # Input file section
        input_group = QGroupBox("Input File")
        input_layout = QVBoxLayout()
        
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        self.file_label.setWordWrap(True)
        file_layout.addWidget(self.file_label)
        
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_input_file)
        file_layout.addWidget(browse_btn)
        input_layout.addLayout(file_layout)
        
        load_btn = QPushButton("Load and Preview")
        load_btn.clicked.connect(self.load_and_preview)
        input_layout.addWidget(load_btn)
        
        input_group.setLayout(input_layout)
        left_layout.addWidget(input_group)
        
        # Simplification settings
        simplify_group = QGroupBox("Simplification Settings")
        simplify_layout = QVBoxLayout()
        
        self.auto_simplify_cb = QCheckBox("Auto-simplify if exceeds limit")
        self.auto_simplify_cb.setChecked(True)
        simplify_layout.addWidget(self.auto_simplify_cb)
        
        max_points_layout = QHBoxLayout()
        max_points_layout.addWidget(QLabel("Max Points:"))
        self.max_points_spin = QSpinBox()
        self.max_points_spin.setRange(100, 10000)
        self.max_points_spin.setValue(1000)
        self.max_points_spin.setSingleStep(100)
        max_points_layout.addWidget(self.max_points_spin)
        simplify_layout.addLayout(max_points_layout)
        
        tolerance_layout = QHBoxLayout()
        tolerance_layout.addWidget(QLabel("Tolerance:"))
        self.tolerance_spin = QDoubleSpinBox()
        self.tolerance_spin.setRange(0.0001, 1.0)
        self.tolerance_spin.setValue(0.0001)
        self.tolerance_spin.setSingleStep(0.0001)
        self.tolerance_spin.setDecimals(4)
        tolerance_layout.addWidget(self.tolerance_spin)
        simplify_layout.addLayout(tolerance_layout)
        
        # Buffer width for LineString conversion
        simplify_layout.addWidget(QLabel("LineString Buffer (for route lines):"))
        buffer_layout = QHBoxLayout()
        buffer_layout.addWidget(QLabel("Buffer Width (meters):"))
        self.buffer_spin = QDoubleSpinBox()
        self.buffer_spin.setRange(1, 500)
        self.buffer_spin.setValue(40)  # 40 meters default
        self.buffer_spin.setSingleStep(5)
        self.buffer_spin.setDecimals(1)
        self.buffer_spin.setToolTip("Width in meters to convert line routes to polygons\n40m = typical highway width\n100m = wider coverage area")
        buffer_layout.addWidget(self.buffer_spin)
        simplify_layout.addLayout(buffer_layout)
        
        simplify_group.setLayout(simplify_layout)
        left_layout.addWidget(simplify_group)
        
        # Output format selection
        format_group = QGroupBox("Output Formats")
        format_layout = QVBoxLayout()
        
        self.shp_cb = QCheckBox("Shapefile (.shp)")
        self.shp_cb.setChecked(False)
        format_layout.addWidget(self.shp_cb)
        
        self.geojson_cb = QCheckBox("GeoJSON (.geojson)")
        self.geojson_cb.setChecked(False)
        format_layout.addWidget(self.geojson_cb)
        
        self.tab_cb = QCheckBox("MapInfo TAB (.tab)")
        self.tab_cb.setChecked(True)
        format_layout.addWidget(self.tab_cb)
        
        format_group.setLayout(format_layout)
        left_layout.addWidget(format_group)
        
        # Output filename customization
        output_name_group = QGroupBox("Output Filename")
        output_name_layout = QVBoxLayout()
        
        output_name_layout.addWidget(QLabel("Custom filename (optional):"))
        filename_layout = QHBoxLayout()
        self.output_name_edit = QLineEdit()
        self.output_name_edit.setPlaceholderText("Leave empty to use original filename")
        self.output_name_edit.setToolTip("Enter a custom filename without extension\nExample: 'Highway_Route_Converted'")
        filename_layout.addWidget(self.output_name_edit)
        
        # Add button to use input filename
        use_original_btn = QPushButton("Use Original")
        use_original_btn.setMaximumWidth(100)
        use_original_btn.clicked.connect(self.use_original_filename)
        use_original_btn.setToolTip("Clear custom name and use original filename")
        filename_layout.addWidget(use_original_btn)
        
        output_name_layout.addLayout(filename_layout)
        
        # Preview of output filename
        self.output_preview_label = QLabel("")
        self.output_preview_label.setStyleSheet("color: #666; font-style: italic;")
        self.output_preview_label.setWordWrap(True)
        output_name_layout.addWidget(self.output_preview_label)
        
        # Connect text change to update preview
        self.output_name_edit.textChanged.connect(self.update_output_preview)
        
        output_name_group.setLayout(output_name_layout)
        left_layout.addWidget(output_name_group)
        
        # Convert button
        self.convert_btn = QPushButton("Convert and Export")
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self.convert_file)
        self.convert_btn.setStyleSheet("QPushButton { padding: 10px; font-size: 14px; }")
        left_layout.addWidget(self.convert_btn)
        
        # Log area
        log_group = QGroupBox("Log")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        left_layout.addWidget(log_group)
        
        left_layout.addStretch()
        
        # Right panel - Map preview
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        preview_label = QLabel("Polygon Preview")
        preview_font = QFont()
        preview_font.setPointSize(12)
        preview_font.setBold(True)
        preview_label.setFont(preview_font)
        right_layout.addWidget(preview_label)
        
        self.map_canvas = MapCanvas(self)
        right_layout.addWidget(self.map_canvas)
        
        # Add panels to splitter
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        
        main_layout.addWidget(splitter)
        
        # Footer message
        footer_label = QLabel("v1.0. Written in Python ❤️ Fadzli Abdullah")
        footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer_label.setStyleSheet("color: #666; font-size: 10px; padding: 5px;")
        main_layout.addWidget(footer_label)
        
        # Initial log message
        self.log("Polygon Converter Ready")
        self.log("Supported formats: KMZ, KML")
        self.log("Max point limit: 1000 (adjustable)")
        self.log("Buffer width: Meters (accurate projection-based)")
    
    def log(self, message):
        """Add message to log"""
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def use_original_filename(self):
        """Clear custom filename to use original"""
        self.output_name_edit.clear()
        self.log("Will use original filename")
    
    def update_output_preview(self):
        """Update the preview of output filenames"""
        custom_name = self.output_name_edit.text().strip()
        
        if not custom_name:
            if self.input_file:
                base_name = Path(self.input_file).stem
                self.output_preview_label.setText(f"Output: {base_name}.[shp/geojson/tab]")
            else:
                self.output_preview_label.setText("Output: [original_name].[shp/geojson/tab]")
        else:
            # Remove any file extensions if user added them
            custom_name = custom_name.replace('.shp', '').replace('.geojson', '').replace('.tab', '')
            self.output_preview_label.setText(f"Output: {custom_name}.[shp/geojson/tab]")
    
    def browse_input_file(self):
        """Browse for input file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select KMZ/KML File",
            "",
            "KML/KMZ Files (*.kml *.kmz);;All Files (*)"
        )
        
        if file_path:
            self.input_file = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.log(f"Selected: {file_path}")
            self.convert_btn.setEnabled(False)
            # Update output preview
            self.update_output_preview()
    
    def load_and_preview(self):
        """Load file and preview polygon"""
        if not self.input_file:
            QMessageBox.warning(self, "Warning", "Please select an input file first")
            return
        
        try:
            self.log("Loading file...")
            
            # Load using conversion thread logic
            conv = ConversionThread(
                self.input_file, "", [], False, 0.0001, 1000, 
                40, None  # 40 meters default
            )
            self.gdf = conv.load_kmz_kml(self.input_file, self.buffer_spin.value())
            
            if self.gdf is None or self.gdf.empty:
                QMessageBox.warning(self, "Warning", "No valid geometries found in file")
                return
            
            # Get statistics
            total_points = sum([len(geom.exterior.coords) if isinstance(geom, Polygon) else 
                              sum([len(p.exterior.coords) for p in geom.geoms]) 
                              for geom in self.gdf.geometry])
            
            self.log(f"Loaded {len(self.gdf)} feature(s)")
            self.log(f"Total points: {total_points}")
            
            if total_points > self.max_points_spin.value():
                self.log(f"⚠️ Warning: Point count exceeds limit ({self.max_points_spin.value()})")
                self.log("   Auto-simplification will be applied during conversion")
            
            # Plot
            self.map_canvas.plot_geometry(self.gdf)
            self.log("Preview displayed")
            
            self.convert_btn.setEnabled(True)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load file:\n{str(e)}")
            self.log(f"Error: {str(e)}")
    
    def convert_file(self):
        """Convert and export file"""
        if not self.input_file:
            QMessageBox.warning(self, "Warning", "Please select and load an input file first")
            return
        
        # Get selected formats
        formats = []
        if self.shp_cb.isChecked():
            formats.append('shapefile')
        if self.geojson_cb.isChecked():
            formats.append('geojson')
        if self.tab_cb.isChecked():
            formats.append('tab')
        
        if not formats:
            QMessageBox.warning(self, "Warning", "Please select at least one output format")
            return
        
        # Get output directory
        output_dir = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory",
            os.path.dirname(self.input_file)
        )
        
        if not output_dir:
            return
        
        # Disable button during conversion
        self.convert_btn.setEnabled(False)
        self.log("\n--- Starting Conversion ---")
        
        # Get custom output name
        custom_name = self.output_name_edit.text().strip()
        if custom_name:
            # Remove any extensions if user added them
            custom_name = custom_name.replace('.shp', '').replace('.geojson', '').replace('.tab', '')
            self.log(f"Custom output name: {custom_name}")
        
        # Create and start conversion thread
        self.conversion_thread = ConversionThread(
            self.input_file,
            output_dir,
            formats,
            self.auto_simplify_cb.isChecked(),
            self.tolerance_spin.value(),
            self.max_points_spin.value(),
            self.buffer_spin.value(),
            custom_name if custom_name else None
        )
        
        self.conversion_thread.progress.connect(self.log)
        self.conversion_thread.finished.connect(self.conversion_finished)
        self.conversion_thread.start()
    
    def conversion_finished(self, success, message):
        """Handle conversion completion"""
        self.convert_btn.setEnabled(True)
        
        if success:
            self.log("✓ " + message)
            QMessageBox.information(self, "Success", message)
        else:
            self.log("✗ " + message)
            QMessageBox.critical(self, "Error", message)


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Set default font to Ubuntu with appropriate size
    default_font = QFont("Ubuntu", 9)
    app.setFont(default_font)
    
    window = PolygonConverterApp()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()