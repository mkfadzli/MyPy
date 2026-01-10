import sys
import os
from pathlib import Path
import zipfile
import shutil
import warnings
import subprocess
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QTextEdit, QLabel, QFileDialog,
    QMessageBox, QGroupBox, QSplitter, QProgressBar, QLineEdit,
    QTabWidget, QDialog, QFormLayout, QDialogButtonBox, QListWidgetItem,
    QCheckBox, QDoubleSpinBox, QSpinBox, QComboBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon
import traceback

warnings.filterwarnings('ignore')

try:
    from osgeo import ogr, osr
    GDAL_AVAILABLE = True
except ImportError:
    GDAL_AVAILABLE = False

try:
    import geopandas as gpd
    import pandas as pd
    from shapely.ops import unary_union
    from shapely.geometry import Polygon, MultiPolygon
    GEOPANDAS_AVAILABLE = True
except ImportError:
    GEOPANDAS_AVAILABLE = False

try:
    import matplotlib
    matplotlib.use('Qt5Agg')
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.figure import Figure
    from matplotlib.patches import Polygon as MplPolygon
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False


class MapCanvas(FigureCanvas):
    def __init__(self, parent=None, width=8, height=6, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super().__init__(self.fig)
        self.setParent(parent)
        self.polygon_data = []  # Store polygon info
        self.parent_gui = None
        
    def clear_map(self):
        self.axes.clear()
        self.axes.set_xlabel('X Coordinate')
        self.axes.set_ylabel('Y Coordinate')
        self.axes.set_title('Map Preview')
        self.axes.grid(True, alpha=0.3)
        self.polygon_data = []
        self.draw()
        
    def plot_geometries(self, all_coords, colors_per_file):
        self.axes.clear()
        self.polygon_data = []
        
        if not all_coords:
            self.axes.text(0.5, 0.5, 'No geometries found', 
                          ha='center', va='center', transform=self.axes.transAxes)
            self.draw()
            return
        
        all_x = []
        all_y = []
        
        for file_idx, (filename, polys) in enumerate(all_coords):
            color = colors_per_file[file_idx % len(colors_per_file)]
            
            for poly_idx, poly_coords in enumerate(polys):
                if poly_coords and len(poly_coords) >= 3:
                    x_coords = [c[0] for c in poly_coords]
                    y_coords = [c[1] for c in poly_coords]
                    
                    # Store polygon data
                    self.polygon_data.append({
                        'coords': poly_coords,
                        'filename': filename,
                        'color': color,
                        'index': len(self.polygon_data)
                    })
                    
                    polygon = MplPolygon(list(zip(x_coords, y_coords)), 
                                        facecolor=color, edgecolor='black', 
                                        linewidth=1, alpha=0.6, picker=5)
                    self.axes.add_patch(polygon)
                    
                    all_x.extend(x_coords)
                    all_y.extend(y_coords)
        
        if all_x and all_y:
            x_margin = (max(all_x) - min(all_x)) * 0.1 or 1
            y_margin = (max(all_y) - min(all_y)) * 0.1 or 1
            self.axes.set_xlim(min(all_x) - x_margin, max(all_x) + x_margin)
            self.axes.set_ylim(min(all_y) - y_margin, max(all_y) + y_margin)
        
        self.axes.set_xlabel('X Coordinate (Easting)')
        self.axes.set_ylabel('Y Coordinate (Northing)')
        self.axes.set_title(f'Map Preview - {len(self.polygon_data)} Polygons (Click to select)')
        self.axes.grid(True, alpha=0.3)
        self.axes.set_aspect('equal', adjustable='datalim')
        
        # Connect click event
        self.fig.canvas.mpl_connect('button_press_event', self.on_click)
        
        self.draw()
    
    def on_click(self, event):
        """Handle polygon click for selection"""
        if event.inaxes != self.axes:
            return
        
        x, y = event.xdata, event.ydata
        
        # Find which polygon was clicked
        for idx, poly_info in enumerate(self.polygon_data):
            coords = poly_info['coords']
            # Simple point-in-polygon check
            if self.point_in_polygon(x, y, coords):
                if self.parent_gui:
                    self.parent_gui.select_polygon_from_map(idx)
                break
    
    def point_in_polygon(self, x, y, poly_coords):
        """Check if point is inside polygon using ray casting"""
        n = len(poly_coords)
        inside = False
        
        p1x, p1y = poly_coords[0]
        for i in range(n + 1):
            p2x, p2y = poly_coords[i % n]
            if y > min(p1y, p2y):
                if y <= max(p1y, p2y):
                    if x <= max(p1x, p2x):
                        if p1y != p2y:
                            xinters = (y - p1y) * (p2x - p1x) / (p2y - p1y) + p1x
                        if p1x == p2x or x <= xinters:
                            inside = not inside
            p1x, p1y = p2x, p2y
        
        return inside
    
    def highlight_selected(self, selected_indices):
        """Highlight selected polygons"""
        self.axes.clear()
        all_x = []
        all_y = []
        
        for idx, poly_info in enumerate(self.polygon_data):
            coords = poly_info['coords']
            x_coords = [c[0] for c in coords]
            y_coords = [c[1] for c in coords]
            
            if idx in selected_indices:
                # Highlight selected
                polygon = MplPolygon(list(zip(x_coords, y_coords)), 
                                    facecolor='yellow', edgecolor='red', 
                                    linewidth=3, alpha=0.8, picker=5)
            else:
                # Normal display
                polygon = MplPolygon(list(zip(x_coords, y_coords)), 
                                    facecolor=poly_info['color'], edgecolor='black', 
                                    linewidth=1, alpha=0.3, picker=5)
            
            self.axes.add_patch(polygon)
            all_x.extend(x_coords)
            all_y.extend(y_coords)
        
        if all_x and all_y:
            x_margin = (max(all_x) - min(all_x)) * 0.1 or 1
            y_margin = (max(all_y) - min(all_y)) * 0.1 or 1
            self.axes.set_xlim(min(all_x) - x_margin, max(all_x) + x_margin)
            self.axes.set_ylim(min(all_y) - y_margin, max(all_y) + y_margin)
        
        self.axes.set_xlabel('X Coordinate (Easting)')
        self.axes.set_ylabel('Y Coordinate (Northing)')
        title = f'{len(selected_indices)} Selected' if selected_indices else 'Click polygons to select'
        self.axes.set_title(f'Map Preview - {title}')
        self.axes.grid(True, alpha=0.3)
        self.axes.set_aspect('equal', adjustable='datalim')
        
        self.draw()


class AttributeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Set Attribute Information")
        self.setModal(True)
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        form_layout = QFormLayout()
        
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("e.g., HIGHWAY_KARAK")
        form_layout.addRow("Name:", self.name_input)
        
        self.group_input = QLineEdit()
        self.group_input.setPlaceholderText("e.g., HUAWEI")
        form_layout.addRow("Group:", self.group_input)
        
        layout.addLayout(form_layout)
        
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
    def get_attributes(self):
        return {
            'Name': self.name_input.text() or 'Merged',
            'Group': self.group_input.text() or ''
        }


class MergeWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str, list)
    preview_ready = pyqtSignal(list, dict, list)
    log_message = pyqtSignal(str)  # Add log signal
    
    def __init__(self, input_files, output_file, merge_attributes=None, selected_polygon_indices=None):
        super().__init__()
        self.input_files = input_files
        self.output_file = output_file
        self.preview_mode = False
        self.merge_attributes = merge_attributes or {}
        self.selected_polygon_indices = selected_polygon_indices or []
        
    def set_preview_mode(self, preview=True):
        self.preview_mode = preview
        
    def run(self):
        try:
            if self.preview_mode:
                self.generate_preview()
            else:
                self.merge_files()
        except Exception as e:
            error_msg = f"Error: {str(e)}\n{traceback.format_exc()}"
            self.finished.emit(False, error_msg, [])
    
    def remove_holes(self, geometry):
        """Remove interior rings (holes) from polygon"""
        if not geometry:
            return None
        
        geom_name = geometry.GetGeometryName()
        
        if geom_name == 'POLYGON':
            exterior_ring = geometry.GetGeometryRef(0)
            if exterior_ring:
                new_poly = ogr.Geometry(ogr.wkbPolygon)
                new_poly.AddGeometry(exterior_ring)
                return new_poly
            return None
            
        elif geom_name == 'MULTIPOLYGON':
            multi_poly = ogr.Geometry(ogr.wkbMultiPolygon)
            for i in range(geometry.GetGeometryCount()):
                poly = geometry.GetGeometryRef(i)
                exterior_ring = poly.GetGeometryRef(0)
                if exterior_ring:
                    new_poly = ogr.Geometry(ogr.wkbPolygon)
                    new_poly.AddGeometry(exterior_ring)
                    multi_poly.AddGeometry(new_poly)
            return multi_poly if multi_poly.GetGeometryCount() > 0 else None
        
        return geometry.Clone()
    
    def remove_duplicate_points(self, coords, tolerance=1e-8):
        """Remove duplicate consecutive points from coordinate list"""
        if not coords or len(coords) < 2:
            return coords
        
        cleaned = [coords[0]]
        
        for i in range(1, len(coords)):
            x1, y1 = cleaned[-1]
            x2, y2 = coords[i]
            
            # Check if points are different (with tolerance)
            if abs(x1 - x2) > tolerance or abs(y1 - y2) > tolerance:
                cleaned.append(coords[i])
        
        # Ensure first and last points are same (closed polygon)
        if len(cleaned) > 2:
            x1, y1 = cleaned[0]
            x2, y2 = cleaned[-1]
            if abs(x1 - x2) > tolerance or abs(y1 - y2) > tolerance:
                cleaned.append(cleaned[0])
        
        return cleaned
            
    def generate_preview(self):
        preview_data = []
        stats = {
            'total_features': 0,
            'total_polygons': 0,
            'geometry_types': set(),
            'srs_list': [],
            'field_names': set()
        }
        all_coords = []
        
        for idx, filepath in enumerate(self.input_files):
            self.progress.emit(int((idx / len(self.input_files)) * 100), 
                             f"Reading {Path(filepath).name}...")
            
            ds = ogr.Open(filepath)
            if not ds:
                continue
                
            layer = ds.GetLayer(0)
            feature_count = layer.GetFeatureCount()
            stats['total_features'] += feature_count
            
            geom_type = ogr.GeometryTypeToName(layer.GetGeomType())
            stats['geometry_types'].add(geom_type)
            
            srs = layer.GetSpatialRef()
            srs_name = srs.GetName() if srs else "Unknown"
            stats['srs_list'].append(srs_name)
            
            layer_defn = layer.GetLayerDefn()
            fields = [layer_defn.GetFieldDefn(i).GetName() for i in range(layer_defn.GetFieldCount())]
            stats['field_names'].update(fields)
            
            file_polys = []
            layer.ResetReading()
            for feature in layer:
                geom = feature.GetGeometryRef()
                if geom and geom.GetGeometryName() in ['POLYGON', 'MULTIPOLYGON']:
                    stats['total_polygons'] += 1
                    
                    if geom.GetGeometryName() == 'POLYGON':
                        ring = geom.GetGeometryRef(0)
                        if ring:
                            points = ring.GetPoints()
                            coords = [(p[0], p[1]) for p in points]
                            file_polys.append(coords)
                    elif geom.GetGeometryName() == 'MULTIPOLYGON':
                        for i in range(geom.GetGeometryCount()):
                            poly = geom.GetGeometryRef(i)
                            ring = poly.GetGeometryRef(0)
                            if ring:
                                points = ring.GetPoints()
                                coords = [(p[0], p[1]) for p in points]
                                file_polys.append(coords)
            
            if file_polys:
                all_coords.append((Path(filepath).name, file_polys))
            
            ds = None
            
        self.progress.emit(100, "Preview ready")
        self.preview_ready.emit(preview_data, stats, all_coords)
        
    def merge_files(self):
        """Merge selected polygons or all polygons into ONE - supports multiple output formats"""
        # Determine output driver based on file extension
        output_ext = os.path.splitext(self.output_file)[1].lower()
        
        driver_map = {
            '.tab': 'MapInfo File',
            '.shp': 'ESRI Shapefile',
            '.geojson': 'GeoJSON',
            '.gpkg': 'GPKG',
            '.gml': 'GML',
            '.json': 'GeoJSON'
        }
        
        driver_name = driver_map.get(output_ext, 'MapInfo File')
        driver = ogr.GetDriverByName(driver_name)
        
        if not driver:
            self.finished.emit(False, f"{driver_name} driver not available", [])
            return
            
        if os.path.exists(self.output_file):
            try:
                driver.DeleteDataSource(self.output_file)
            except:
                pass  # Some drivers don't support deletion
            
        self.progress.emit(0, "Creating output file...")
        out_ds = driver.CreateDataSource(self.output_file)
        if not out_ds:
            self.finished.emit(False, f"Could not create output file", [])
            return
            
        first_ds = ogr.Open(self.input_files[0])
        if not first_ds:
            self.finished.emit(False, f"Could not open first file", [])
            return
            
        first_layer = first_ds.GetLayer(0)
        srs = first_layer.GetSpatialRef()
        first_ds = None
        
        out_layer = out_ds.CreateLayer('merged', srs, ogr.wkbPolygon)
        if not out_layer:
            self.finished.emit(False, "Could not create output layer", [])
            return
        
        name_field = ogr.FieldDefn('Name', ogr.OFTString)
        name_field.SetWidth(254)
        out_layer.CreateField(name_field)
        
        group_field = ogr.FieldDefn('Group', ogr.OFTString)
        group_field.SetWidth(254)
        out_layer.CreateField(group_field)
        
        self.progress.emit(10, "Collecting polygons...")
        union_geom = None
        total_polys = 0
        file_stats = []
        current_poly_idx = 0
        
        for idx, filepath in enumerate(self.input_files):
            self.progress.emit(int(10 + (idx / len(self.input_files)) * 70),
                             f"Reading {Path(filepath).name}...")
            
            ds = ogr.Open(filepath)
            if not ds:
                file_stats.append(f"⚠ Could not open: {Path(filepath).name}")
                continue
                
            layer = ds.GetLayer(0)
            count = 0
            
            layer.ResetReading()
            for feature in layer:
                geom = feature.GetGeometryRef()
                if geom:
                    # Check if this polygon is selected (if selection mode)
                    if self.selected_polygon_indices:
                        if current_poly_idx not in self.selected_polygon_indices:
                            current_poly_idx += 1
                            continue
                    
                    current_poly_idx += 1
                    clean_geom = self.remove_holes(geom)
                    
                    if clean_geom:
                        # Extract coordinates and remove duplicates
                        if clean_geom.GetGeometryName() == 'POLYGON':
                            ring = clean_geom.GetGeometryRef(0)
                            if ring:
                                points = ring.GetPoints()
                                coords = [(p[0], p[1]) for p in points]
                                coords = self.remove_duplicate_points(coords)
                                
                                # Rebuild geometry with cleaned coords
                                new_ring = ogr.Geometry(ogr.wkbLinearRing)
                                for x, y in coords:
                                    new_ring.AddPoint(x, y)
                                clean_geom = ogr.Geometry(ogr.wkbPolygon)
                                clean_geom.AddGeometry(new_ring)
                        
                        if union_geom is None:
                            union_geom = clean_geom
                        else:
                            union_geom = union_geom.Union(clean_geom)
                        count += 1
            
            total_polys += count
            ds = None
            file_stats.append(f"✓ {Path(filepath).name}: {count} polygons")
        
        if union_geom is None:
            self.finished.emit(False, "No polygons found to merge", [])
            return
        
        self.progress.emit(90, "Creating merged feature...")
        
        union_geom = self.remove_holes(union_geom)
        
        if union_geom.GetGeometryName() == 'MULTIPOLYGON':
            union_geom = union_geom.Buffer(0)
            union_geom = self.remove_holes(union_geom)
        
        # CRITICAL FIX: Remove duplicate points from merged geometry
        self.log_message.emit("Cleaning merged geometry for Discovery compatibility...")
        if union_geom.GetGeometryName() == 'POLYGON':
            ring = union_geom.GetGeometryRef(0)
            if ring:
                points = ring.GetPoints()
                original_count = len(points)
                coords = [(p[0], p[1]) for p in points]
                cleaned_coords = self.remove_duplicate_points(coords, tolerance=0.00001)
                
                # Rebuild geometry with cleaned points
                new_ring = ogr.Geometry(ogr.wkbLinearRing)
                for x, y in cleaned_coords:
                    new_ring.AddPoint(x, y)
                union_geom = ogr.Geometry(ogr.wkbPolygon)
                union_geom.AddGeometry(new_ring)
                
                removed_count = original_count - len(cleaned_coords)
                if removed_count > 0:
                    self.log_message.emit(f"  Removed {removed_count} duplicate points ({original_count} → {len(cleaned_coords)} points)")
        
        out_feature = ogr.Feature(out_layer.GetLayerDefn())
        out_feature.SetGeometry(union_geom)
        out_feature.SetField('Name', self.merge_attributes.get('Name', 'Merged'))
        out_feature.SetField('Group', self.merge_attributes.get('Group', ''))
        
        out_layer.CreateFeature(out_feature)
        
        out_ds = None
        
        self.progress.emit(100, "Merge complete!")
        self.finished.emit(True, 
                          f"Successfully merged {total_polys} polygons into 1 dissolved polygon\nOutput format: {driver_name}", 
                          file_stats)


class ProcessingThread(QThread):
    """Thread for running geospatial processing operations"""
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, operation, params):
        super().__init__()
        self.operation = operation
        self.params = params
        
    def run(self):
        try:
            if self.operation == "dissolve":
                self.run_dissolve()
            elif self.operation == "merge":
                self.run_merge()
            elif self.operation == "buffer":
                self.run_buffer()
            elif self.operation == "simplify":
                self.run_simplify()
            elif self.operation == "clip":
                self.run_clip()
            elif self.operation == "intersection":
                self.run_intersection()
            elif self.operation == "difference":
                self.run_difference()
            elif self.operation == "union":
                self.run_union()
            elif self.operation == "convert_to_shapefile_zip":
                self.run_convert_to_shapefile_zip()
            elif self.operation == "convert":
                self.run_convert()
        except Exception as e:
            self.finished.emit(False, f"Error: {str(e)}")
    
    def run_dissolve(self):
        """Dissolve multipolygons into single continuous polygon"""
        self.status.emit("Loading input files...")
        self.progress.emit(10)
        
        input_files = self.params['input_files']
        output_file = self.params['output_file']
        
        # Load all files with better error handling
        gdfs = []
        for idx, file_path in enumerate(input_files):
            try:
                self.status.emit(f"Reading file {idx+1}/{len(input_files)}: {os.path.basename(file_path)}")
                # Read file - for shapefiles, explicitly specify layer=0
                if file_path.lower().endswith('.shp'):
                    gdf = gpd.read_file(file_path, layer=0)
                else:
                    gdf = gpd.read_file(file_path)
                
                if len(gdf) == 0:
                    raise ValueError(f"File is empty: {os.path.basename(file_path)}")
                
                gdfs.append(gdf)
            except Exception as e:
                raise Exception(f"Error reading {os.path.basename(file_path)}: {str(e)}")
        
        self.status.emit(f"Loaded {len(gdfs)} file(s). Merging...")
        self.progress.emit(30)
        
        # Merge all GeoDataFrames
        if len(gdfs) > 1:
            merged_gdf = gpd.GeoDataFrame(pd.concat(gdfs, ignore_index=True))
        else:
            merged_gdf = gdfs[0]
        
        # Ensure same CRS
        if merged_gdf.crs is None:
            merged_gdf.set_crs(epsg=4326, inplace=True)
        
        self.status.emit("Dissolving geometries into single polygon...")
        self.progress.emit(60)
        
        # Get the merge strategy from params (default to union)
        merge_strategy = self.params.get('merge_strategy', 'union')
        custom_name = self.params.get('custom_name', f"DISSOLVED_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        
        # Dissolve all geometries into one
        dissolved_geom = unary_union(merged_gdf.geometry)
        
        from shapely.geometry import Polygon, MultiPolygon
        
        original_type = type(dissolved_geom).__name__
        part_count = 1
        
        if isinstance(dissolved_geom, MultiPolygon):
            part_count = len(dissolved_geom.geoms)
            self.status.emit(f"Found {part_count} separate parts. Creating continuous polygon...")
            
            if merge_strategy == 'convex_hull':
                # Create convex hull - creates truly continuous single polygon
                self.status.emit("Using Convex Hull method to create continuous polygon...")
                dissolved_geom = dissolved_geom.convex_hull
            elif merge_strategy == 'buffer':
                # Buffer to connect separate parts, then remove buffer
                self.status.emit("Using Buffer method to connect separate parts...")
                buffer_distance = 0.0001  # Small buffer to connect parts
                dissolved_geom = dissolved_geom.buffer(buffer_distance).buffer(-buffer_distance)
            else:  # 'largest' - keep only largest polygon
                self.status.emit(f"Keeping largest polygon, removing {part_count-1} smaller parts...")
                largest_polygon = max(dissolved_geom.geoms, key=lambda p: p.area)
                dissolved_geom = largest_polygon
            
            # Check if still MultiPolygon after processing
            if isinstance(dissolved_geom, MultiPolygon):
                self.status.emit("Still MultiPolygon after processing, taking largest part...")
                dissolved_geom = max(dissolved_geom.geoms, key=lambda p: p.area)
        
        # Double-check: Ensure it's absolutely a single Polygon
        if not isinstance(dissolved_geom, Polygon):
            raise ValueError(f"ERROR: Result is {type(dissolved_geom).__name__}, not a single Polygon!")
        
        # Triple-check: Verify geometry type
        if dissolved_geom.geom_type != 'Polygon':
            raise ValueError(f"ERROR: Geometry type is {dissolved_geom.geom_type}, must be 'Polygon'!")
        
        # Create new GeoDataFrame with single feature using custom name
        dissolved_gdf = gpd.GeoDataFrame(
            {'id': [1], 'name': [custom_name]}, 
            geometry=[dissolved_geom],
            crs=merged_gdf.crs
        )
        
        self.status.emit("Saving output file...")
        self.progress.emit(80)
        
        # Determine driver based on file extension
        ext = os.path.splitext(output_file)[1].lower()
        if ext == '.tab':
            driver = 'MapInfo File'
        elif ext == '.shp':
            driver = 'ESRI Shapefile'
        elif ext == '.geojson':
            driver = 'GeoJSON'
        elif ext == '.gpkg':
            driver = 'GPKG'
        else:
            driver = 'GeoJSON'
        
        # Save the dissolved geometry
        # First, remove existing file if it exists to prevent layer conflicts
        if os.path.exists(output_file):
            try:
                # For shapefiles, remove all associated files
                if ext == '.shp':
                    base = os.path.splitext(output_file)[0]
                    for suffix in ['.shp', '.shx', '.dbf', '.prj', '.cpg', '.sbn', '.sbx', '.shp.xml']:
                        file_to_remove = base + suffix
                        if os.path.exists(file_to_remove):
                            os.remove(file_to_remove)
                else:
                    os.remove(output_file)
                self.status.emit("Removed existing file...")
            except PermissionError:
                raise PermissionError(f"Cannot overwrite {output_file}. Please close any programs using this file (QGIS, ArcGIS, etc.) and try again.")
        
        dissolved_gdf.to_file(output_file, driver=driver)
        
        # Verify the saved file
        self.status.emit("Verifying saved file...")
        self.progress.emit(95)
        
        verify_gdf = gpd.read_file(output_file)
        verify_geom = verify_gdf.geometry.iloc[0]
        
        if verify_geom.geom_type != 'Polygon':
            raise ValueError(f"VERIFICATION FAILED: Saved file has {verify_geom.geom_type}, not Polygon!")
        
        self.progress.emit(100)
        
        self.finished.emit(True, 
            f"✅ SUCCESS! Created TRUE SINGLE-PART Polygon!\n\n"
            f"Polygon Name: {custom_name}\n"
            f"Original: {original_type} with {part_count} part(s)\n"
            f"Result: Polygon (1 part only)\n"
            f"Merge Strategy: {merge_strategy}\n"
            f"Geometry Type: {dissolved_geom.geom_type}\n"
            f"Input features: {len(merged_gdf)}\n"
            f"Output features: 1\n\n"
            f"✅ READY FOR HUAWEI DISCOVERY!\n"
            f"Output: {output_file}")
    
    def run_merge(self):
        """Merge multiple files without dissolving"""
        self.status.emit("Loading files...")
        self.progress.emit(20)
        
        input_files = self.params['input_files']
        output_file = self.params['output_file']
        
        # Load all files
        gdfs = []
        for file_path in input_files:
            gdf = gpd.read_file(file_path)
            gdfs.append(gdf)
        
        self.progress.emit(50)
        self.status.emit("Merging files...")
        
        # Merge all GeoDataFrames
        merged_gdf = gpd.GeoDataFrame(pd.concat(gdfs, ignore_index=True))
        
        # Ensure CRS is set
        if merged_gdf.crs is None:
            merged_gdf.set_crs(epsg=4326, inplace=True)
        
        self.status.emit("Saving merged file...")
        self.progress.emit(80)
        
        # Determine driver
        ext = os.path.splitext(output_file)[1].lower()
        if ext == '.tab':
            driver = 'MapInfo File'
        elif ext == '.shp':
            driver = 'ESRI Shapefile'
        elif ext == '.geojson':
            driver = 'GeoJSON'
        elif ext == '.gpkg':
            driver = 'GPKG'
        else:
            driver = 'GeoJSON'
        
        merged_gdf.to_file(output_file, driver=driver)
        
        self.progress.emit(100)
        self.finished.emit(True, 
            f"✓ Merged {len(gdfs)} files successfully!\n"
            f"Total features: {len(merged_gdf)}\n"
            f"Output: {output_file}")
    
    def run_buffer(self):
        """Create buffer around geometries"""
        self.status.emit("Loading input file...")
        self.progress.emit(20)
        
        input_file = self.params['input_layer']
        output_file = self.params['output_file']
        distance = self.params['distance']
        dissolve = self.params.get('dissolve', False)
        
        gdf = gpd.read_file(input_file)
        
        self.status.emit(f"Creating buffer ({distance}m)...")
        self.progress.emit(50)
        
        # Reproject to metric CRS if needed (for accurate distance)
        original_crs = gdf.crs
        if gdf.crs and gdf.crs.is_geographic:
            # Reproject to Web Mercator for buffering
            gdf = gdf.to_crs(epsg=3857)
        
        # Create buffer
        gdf['geometry'] = gdf.geometry.buffer(distance)
        
        # Dissolve if requested
        if dissolve:
            self.status.emit("Dissolving overlapping buffers...")
            self.progress.emit(70)
            dissolved_geom = unary_union(gdf.geometry)
            gdf = gpd.GeoDataFrame({'id': [1]}, geometry=[dissolved_geom], crs=gdf.crs)
        
        # Reproject back to original CRS
        if original_crs:
            gdf = gdf.to_crs(original_crs)
        
        self.status.emit("Saving buffer...")
        self.progress.emit(90)
        
        ext = os.path.splitext(output_file)[1].lower()
        driver = {'.tab': 'MapInfo File', '.shp': 'ESRI Shapefile', 
                 '.geojson': 'GeoJSON', '.gpkg': 'GPKG'}.get(ext, 'GeoJSON')
        
        gdf.to_file(output_file, driver=driver)
        
        self.progress.emit(100)
        self.finished.emit(True, 
            f"✓ Buffer created successfully!\n"
            f"Distance: {distance}m\n"
            f"Features: {len(gdf)}\n"
            f"Output: {output_file}")
    
    def run_simplify(self):
        """Simplify geometry"""
        self.status.emit("Loading input file...")
        self.progress.emit(20)
        
        input_file = self.params['input_layer']
        output_file = self.params['output_file']
        tolerance = self.params['tolerance']
        
        gdf = gpd.read_file(input_file)
        
        self.status.emit(f"Simplifying geometry (tolerance: {tolerance})...")
        self.progress.emit(50)
        
        # Simplify geometries
        gdf['geometry'] = gdf.geometry.simplify(tolerance, preserve_topology=True)
        
        self.status.emit("Saving simplified file...")
        self.progress.emit(90)
        
        ext = os.path.splitext(output_file)[1].lower()
        driver = {'.tab': 'MapInfo File', '.shp': 'ESRI Shapefile', 
                 '.geojson': 'GeoJSON', '.gpkg': 'GPKG'}.get(ext, 'GeoJSON')
        
        gdf.to_file(output_file, driver=driver)
        
        self.progress.emit(100)
        self.finished.emit(True, 
            f"✓ Geometry simplified!\n"
            f"Tolerance: {tolerance}\n"
            f"Output: {output_file}")
    
    def run_clip(self):
        """Clip geometry - keep only area within clip polygon"""
        self.status.emit("Loading files...")
        self.progress.emit(20)
        
        input_file = self.params['input_layer']
        clip_file = self.params['clip_layer']
        output_file = self.params['output_file']
        
        gdf_input = gpd.read_file(input_file)
        gdf_clip = gpd.read_file(clip_file)
        
        # Ensure same CRS
        if gdf_input.crs != gdf_clip.crs:
            gdf_clip = gdf_clip.to_crs(gdf_input.crs)
        
        self.status.emit("Clipping geometry...")
        self.progress.emit(60)
        
        # Clip operation
        clipped = gpd.clip(gdf_input, gdf_clip)
        
        self.status.emit("Saving result...")
        self.progress.emit(90)
        
        ext = os.path.splitext(output_file)[1].lower()
        driver = {'.tab': 'MapInfo File', '.shp': 'ESRI Shapefile', 
                 '.geojson': 'GeoJSON', '.gpkg': 'GPKG'}.get(ext, 'GeoJSON')
        
        clipped.to_file(output_file, driver=driver)
        
        self.progress.emit(100)
        self.finished.emit(True, 
            f"✓ Clipping completed!\n"
            f"Input features: {len(gdf_input)}\n"
            f"Output features: {len(clipped)}\n"
            f"Output: {output_file}")
    
    def run_intersection(self):
        """Intersection of two layers"""
        self.status.emit("Loading files...")
        self.progress.emit(20)
        
        input_file = self.params['input_layer']
        overlay_file = self.params['overlay_layer']
        output_file = self.params['output_file']
        
        gdf1 = gpd.read_file(input_file)
        gdf2 = gpd.read_file(overlay_file)
        
        # Ensure same CRS
        if gdf1.crs != gdf2.crs:
            gdf2 = gdf2.to_crs(gdf1.crs)
        
        self.status.emit("Computing intersection...")
        self.progress.emit(60)
        
        # Intersection
        result = gpd.overlay(gdf1, gdf2, how='intersection')
        
        self.status.emit("Saving result...")
        self.progress.emit(90)
        
        ext = os.path.splitext(output_file)[1].lower()
        driver = {'.tab': 'MapInfo File', '.shp': 'ESRI Shapefile', 
                 '.geojson': 'GeoJSON', '.gpkg': 'GPKG'}.get(ext, 'GeoJSON')
        
        result.to_file(output_file, driver=driver)
        
        self.progress.emit(100)
        self.finished.emit(True, 
            f"✓ Intersection completed!\n"
            f"Output features: {len(result)}\n"
            f"Output: {output_file}")
    
    def run_difference(self):
        """Difference operation - remove overlay area from input"""
        self.status.emit("Loading files...")
        self.progress.emit(20)
        
        input_file = self.params['input_layer']
        overlay_file = self.params['overlay_layer']
        output_file = self.params['output_file']
        
        gdf_input = gpd.read_file(input_file)
        gdf_overlay = gpd.read_file(overlay_file)
        
        # Ensure same CRS
        if gdf_input.crs != gdf_overlay.crs:
            gdf_overlay = gdf_overlay.to_crs(gdf_input.crs)
        
        self.status.emit("Removing selected area...")
        self.progress.emit(60)
        
        # Difference operation
        result = gpd.overlay(gdf_input, gdf_overlay, how='difference')
        
        self.status.emit("Saving result...")
        self.progress.emit(90)
        
        ext = os.path.splitext(output_file)[1].lower()
        driver = {'.tab': 'MapInfo File', '.shp': 'ESRI Shapefile', 
                 '.geojson': 'GeoJSON', '.gpkg': 'GPKG'}.get(ext, 'GeoJSON')
        
        result.to_file(output_file, driver=driver)
        
        self.progress.emit(100)
        self.finished.emit(True, 
            f"✓ Parts removed successfully!\n"
            f"Input features: {len(gdf_input)}\n"
            f"Output features: {len(result)}\n"
            f"Output: {output_file}")
    
    def run_union(self):
        """Union of two layers"""
        self.status.emit("Loading files...")
        self.progress.emit(20)
        
        input_file = self.params['input_layer']
        overlay_file = self.params['overlay_layer']
        output_file = self.params['output_file']
        
        gdf1 = gpd.read_file(input_file)
        gdf2 = gpd.read_file(overlay_file)
        
        # Ensure same CRS
        if gdf1.crs != gdf2.crs:
            gdf2 = gdf2.to_crs(gdf1.crs)
        
        self.status.emit("Computing union...")
        self.progress.emit(60)
        
        # Union
        result = gpd.overlay(gdf1, gdf2, how='union')
        
        self.status.emit("Saving result...")
        self.progress.emit(90)
        
        ext = os.path.splitext(output_file)[1].lower()
        driver = {'.tab': 'MapInfo File', '.shp': 'ESRI Shapefile', 
                 '.geojson': 'GeoJSON', '.gpkg': 'GPKG'}.get(ext, 'GeoJSON')
        
        result.to_file(output_file, driver=driver)
        
        self.progress.emit(100)
        self.finished.emit(True, 
            f"✓ Union completed!\n"
            f"Output features: {len(result)}\n"
            f"Output: {output_file}")

    def run_convert(self):
        """Generic conversion supporting single-file and batch conversions.
        Supports shapefile field name truncation/renaming via params['shapefile_options'].
        """
        self.status.emit("Preparing conversion...")
        self.progress.emit(5)

        # Batch mode
        input_files = self.params.get('input_files')
        output_folder = self.params.get('output_folder')
        output_ext = self.params.get('output_ext')

        # Single mode
        input_file = self.params.get('input_file')
        output_file = self.params.get('output_file')

        shp_opts = self.params.get('shapefile_options', {}) or {}
        truncate = bool(shp_opts.get('truncate', False))
        truncate_len = int(shp_opts.get('truncate_len', 10)) if shp_opts.get('truncate_len') else 10
        prefix = shp_opts.get('prefix', '') or ''

        tasks = []
        if input_files and output_folder:
            # Build tasks list for batch
            for in_f in input_files:
                base = os.path.splitext(os.path.basename(in_f))[0]
                out_path = os.path.join(output_folder, base + (output_ext or os.path.splitext(in_f)[1]))
                tasks.append((in_f, out_path))
        elif input_file and output_file:
            tasks = [(input_file, output_file)]
        else:
            self.finished.emit(False, "Invalid conversion parameters")
            return

        results = []
        total = len(tasks)
        for idx, (src, dst) in enumerate(tasks, start=1):
            try:
                self.status.emit(f"Reading {os.path.basename(src)} ({idx}/{total})")
                self.progress.emit(int(5 + (idx - 1) / max(1, total) * 60))
                gdf = gpd.read_file(src)
            except Exception as e:
                results.append((src, False, f"Read failed: {str(e)}"))
                continue

            # If saving to shapefile and truncation/rename requested
            dst_ext = os.path.splitext(dst)[1].lower()
            if dst_ext == '.shp' and truncate:
                try:
                    # Build mapping for column names (exclude geometry column)
                    cols = [c for c in gdf.columns if c != gdf.geometry.name]
                    mapping = {}
                    seen = set()
                    for col in cols:
                        # make upper, remove spaces, and prefix
                        new_col = prefix + str(col).upper()
                        # truncate
                        new_col = new_col[:truncate_len]
                        # ensure uniqueness
                        base_name = new_col
                        i = 1
                        while new_col in seen or new_col == 'GEOMETRY':
                            # append number to make unique within allowed length
                            suffix = str(i)
                            allowed = truncate_len - len(suffix)
                            new_col = (base_name[:allowed] if allowed>0 else base_name) + suffix
                            i += 1
                        seen.add(new_col)
                        mapping[col] = new_col

                    # Apply rename
                    if mapping:
                        gdf = gdf.rename(columns=mapping)
                except Exception as e:
                    results.append((src, False, f"Field rename failed: {str(e)}"))
                    continue

            # Save file
            try:
                self.status.emit(f"Saving {os.path.basename(dst)} ({idx}/{total})")
                self.progress.emit(int(65 + (idx / max(1, total)) * 30))

                driver_map = {
                    '.tab': 'MapInfo File',
                    '.shp': 'ESRI Shapefile',
                    '.geojson': 'GeoJSON',
                    '.gpkg': 'GPKG',
                    '.gml': 'GML'
                }
                drv = driver_map.get(dst_ext)

                # Remove existing outputs to avoid conflicts
                if os.path.exists(dst):
                    try:
                        os.remove(dst)
                    except Exception:
                        # for shapefile, remove sidecar files
                        base = os.path.splitext(dst)[0]
                        for suf in ['.shp', '.shx', '.dbf', '.prj', '.cpg']:
                            p = base + suf
                            if os.path.exists(p):
                                try:
                                    os.remove(p)
                                except Exception:
                                    pass

                if drv:
                    gdf.to_file(dst, driver=drv)
                else:
                    gdf.to_file(dst)

                results.append((src, True, f"Saved to {dst}"))
            except Exception as e:
                results.append((src, False, f"Save failed: {str(e)}"))

        # Build summary
        success_count = sum(1 for r in results if r[1])
        fail_count = len(results) - success_count
        msg_lines = [f"Conversion completed: {success_count}/{len(results)} succeeded"]
        for src, ok, note in results:
            status = 'OK' if ok else 'ERROR'
            msg_lines.append(f"{status}: {os.path.basename(src)} -> {note}")

        self.progress.emit(100)
        if fail_count == 0:
            self.finished.emit(True, "\n".join(msg_lines))
        else:
            self.finished.emit(False, "\n".join(msg_lines))
    
    def run_convert_to_shapefile_zip(self):
        """Convert any geospatial file to zipped shapefile"""
        self.status.emit("Loading input file...")
        self.progress.emit(20)
        
        input_file = self.params['input_file']
        output_zip = self.params['output_zip']
        
        # Read the input file
        gdf = gpd.read_file(input_file)
        
        self.status.emit("Creating shapefile...")
        self.progress.emit(50)
        
        # Create temporary directory for shapefile components
        import tempfile
        temp_dir = tempfile.mkdtemp()
        
        # Base name for shapefile (without extension)
        base_name = os.path.splitext(os.path.basename(output_zip))[0]
        shp_path = os.path.join(temp_dir, f"{base_name}.shp")
        
        # Save as shapefile
        gdf.to_file(shp_path, driver='ESRI Shapefile')
        
        self.status.emit("Creating ZIP archive...")
        self.progress.emit(80)
        
        # Create ZIP file containing all shapefile components
        with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add all files in temp directory (shp, shx, dbf, prj, etc.)
            for file in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, file)
                zipf.write(file_path, file)
        
        # Clean up temp directory
        shutil.rmtree(temp_dir)
        
        self.progress.emit(100)
        self.finished.emit(True, 
            f"✓ Shapefile ZIP created successfully!\n"
            f"Input features: {len(gdf)}\n"
            f"Ready for Huawei Discovery upload!\n"
            f"Output: {output_zip}")


class KMZConversionWorker(QThread):
    """Worker thread for KMZ/KML file conversion to prevent GUI freezing"""
    progress = pyqtSignal(str)
    file_started = pyqtSignal(str)
    file_finished = pyqtSignal(str, bool)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, input_files, output_dir, formats, create_zip=False):
        super().__init__()
        self.input_files = input_files
        self.output_dir = output_dir
        self.formats = formats
        self.create_zip = create_zip
        self.output_files = []
        self.ogr2ogr_cmd = self._find_ogr2ogr()
    
    def _find_ogr2ogr(self):
        """Find ogr2ogr executable - simplified version that trusts conda environment"""
        import sys
        
        # First try direct command - this should work in conda environment
        try:
            result = subprocess.run(['ogr2ogr', '--version'], 
                                  capture_output=True, timeout=5)
            if result.returncode == 0:
                return 'ogr2ogr'
        except:
            pass
        
        # If direct command fails, try to find full path in conda environment
        conda_prefix = os.environ.get('CONDA_PREFIX')
        if conda_prefix:
            # Windows paths
            possible_paths = [
                os.path.join(conda_prefix, 'Library', 'bin', 'ogr2ogr.exe'),
                os.path.join(conda_prefix, 'Scripts', 'ogr2ogr.exe'),
                os.path.join(conda_prefix, 'bin', 'ogr2ogr.exe'),
                # Linux/Mac paths
                os.path.join(conda_prefix, 'bin', 'ogr2ogr'),
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    return path
        
        # Fallback: just return 'ogr2ogr' and let it fail with a clear error if not found
        return 'ogr2ogr'
        
    def run(self):
        try:
            total_files = len(self.input_files)
            successful_files = 0
            failed_files = 0
            
            for file_index, input_file in enumerate(self.input_files, 1):
                base_name = Path(input_file).stem
                self.file_started.emit(os.path.basename(input_file))
                self.progress.emit(f"\n{'='*60}")
                self.progress.emit(f"Processing file {file_index}/{total_files}: {os.path.basename(input_file)}")
                self.progress.emit(f"{'='*60}")
                
                try:
                    success_count = 0
                    total_formats = len(self.formats)
                    
                    # Convert KMZ to KML first if needed
                    if input_file.lower().endswith('.kmz'):
                        kml_file = os.path.join(self.output_dir, f"{base_name}.kml")
                        self.progress.emit(f"Extracting KML from KMZ...")
                        
                        try:
                            with zipfile.ZipFile(input_file, 'r') as zip_ref:
                                kml_files = [f for f in zip_ref.namelist() if f.endswith('.kml')]
                                self.progress.emit(f"Found {len(kml_files)} KML file(s) in archive")
                                
                                if kml_files:
                                    self.progress.emit(f"Extracting: {kml_files[0]}")
                                    kml_content = zip_ref.read(kml_files[0])
                                    with open(kml_file, 'wb') as f:
                                        f.write(kml_content)
                                    self.progress.emit(f"✓ KML extracted successfully to: {kml_file}")
                                    if 'kml' in self.formats:
                                        success_count += 1
                                        self.output_files.append(kml_file)
                                else:
                                    raise Exception("No KML file found in KMZ archive")
                        except zipfile.BadZipFile as e:
                            self.progress.emit(f"✗ Invalid KMZ file (not a valid ZIP): {str(e)}")
                            self.file_finished.emit(os.path.basename(input_file), False)
                            failed_files += 1
                            continue
                        except Exception as e:
                            self.progress.emit(f"✗ KML extraction failed: {str(e)}")
                            import traceback
                            self.progress.emit(f"Traceback: {traceback.format_exc()}")
                            self.file_finished.emit(os.path.basename(input_file), False)
                            failed_files += 1
                            continue
                        
                        source_file = kml_file
                    else:
                        source_file = input_file
                        if 'kml' in self.formats:
                            if os.path.dirname(source_file) != self.output_dir:
                                output_kml = os.path.join(self.output_dir, os.path.basename(source_file))
                                shutil.copy2(source_file, output_kml)
                                self.output_files.append(output_kml)
                            else:
                                self.output_files.append(source_file)
                            success_count += 1
                    
                    # Convert to SHP
                    if 'shp' in self.formats:
                        output_shp = os.path.join(self.output_dir, f"{base_name}.shp")
                        self.progress.emit(f"Converting to Shapefile...")
                        self.progress.emit(f"  Source: {source_file}")
                        self.progress.emit(f"  Target: {output_shp}")
                        self.progress.emit(f"  Command: {self.ogr2ogr_cmd}")
                        
                        try:
                            result = subprocess.run([
                                self.ogr2ogr_cmd, '-f', 'ESRI Shapefile',
                                output_shp, source_file
                            ], capture_output=True, text=True, timeout=30)
                            
                            if result.returncode == 0:
                                self.progress.emit(f"✓ Shapefile created successfully")
                                success_count += 1
                                shp_dir = os.path.dirname(output_shp)
                                shp_base = os.path.splitext(output_shp)[0]
                                for ext in ['.shp', '.shx', '.dbf', '.prj', '.cpg']:
                                    shp_file = shp_base + ext
                                    if os.path.exists(shp_file):
                                        self.output_files.append(shp_file)
                                        self.progress.emit(f"  Created: {os.path.basename(shp_file)}")
                            else:
                                self.progress.emit(f"✗ Shapefile conversion failed (return code: {result.returncode})")
                                if result.stderr:
                                    self.progress.emit(f"  Error: {result.stderr}")
                                if result.stdout:
                                    self.progress.emit(f"  Output: {result.stdout}")
                        except subprocess.TimeoutExpired:
                            self.progress.emit(f"✗ Shapefile conversion timed out (>30 seconds)")
                        except Exception as e:
                            self.progress.emit(f"✗ Shapefile conversion error: {str(e)}")
                            import traceback
                            self.progress.emit(f"Traceback: {traceback.format_exc()}")
                    
                    # Convert to TAB
                    if 'tab' in self.formats:
                        output_tab = os.path.join(self.output_dir, f"{base_name}.tab")
                        self.progress.emit(f"Converting to MapInfo TAB...")
                        self.progress.emit(f"  Source: {source_file}")
                        self.progress.emit(f"  Target: {output_tab}")
                        
                        try:
                            result = subprocess.run([
                                self.ogr2ogr_cmd, '-f', 'MapInfo File',
                                output_tab, source_file
                            ], capture_output=True, text=True, timeout=30)
                            
                            if result.returncode == 0:
                                self.progress.emit(f"✓ MapInfo TAB created successfully")
                                success_count += 1
                                tab_dir = os.path.dirname(output_tab)
                                tab_base = os.path.splitext(output_tab)[0]
                                for ext in ['.tab', '.dat', '.id', '.map', '.ind']:
                                    tab_file = tab_base + ext
                                    if os.path.exists(tab_file):
                                        self.output_files.append(tab_file)
                                        self.progress.emit(f"  Created: {os.path.basename(tab_file)}")
                            else:
                                self.progress.emit(f"✗ TAB conversion failed (return code: {result.returncode})")
                                if result.stderr:
                                    self.progress.emit(f"  Error: {result.stderr}")
                                if result.stdout:
                                    self.progress.emit(f"  Output: {result.stdout}")
                        except subprocess.TimeoutExpired:
                            self.progress.emit(f"✗ TAB conversion timed out (>30 seconds)")
                        except Exception as e:
                            self.progress.emit(f"✗ TAB conversion error: {str(e)}")
                            import traceback
                            self.progress.emit(f"Traceback: {traceback.format_exc()}")
                    
                    # Check if this file succeeded
                    if success_count == total_formats:
                        self.progress.emit(f"✓ All formats converted for {os.path.basename(input_file)}")
                        self.file_finished.emit(os.path.basename(input_file), True)
                        successful_files += 1
                    elif success_count > 0:
                        self.progress.emit(f"⚠ Partial success: {success_count}/{total_formats} formats")
                        self.file_finished.emit(os.path.basename(input_file), True)
                        successful_files += 1
                    else:
                        self.progress.emit(f"✗ All conversions failed for {os.path.basename(input_file)}")
                        self.file_finished.emit(os.path.basename(input_file), False)
                        failed_files += 1
                        
                except Exception as e:
                    self.progress.emit(f"✗ Error processing {os.path.basename(input_file)}: {str(e)}")
                    self.file_finished.emit(os.path.basename(input_file), False)
                    failed_files += 1
            
            # Create ZIP file if requested
            if self.create_zip and self.output_files:
                self.progress.emit(f"\n{'='*60}")
                self.progress.emit(f"Creating ZIP archive for Huawei Discovery...")
                self.progress.emit(f"{'='*60}")
                
                try:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    zip_filename = f"converted_files_{timestamp}.zip"
                    zip_path = os.path.join(self.output_dir, zip_filename)
                    
                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for file_path in self.output_files:
                            if os.path.exists(file_path):
                                arcname = os.path.basename(file_path)
                                zipf.write(file_path, arcname)
                                self.progress.emit(f"  Added: {arcname}")
                    
                    self.progress.emit(f"✓ ZIP archive created: {zip_filename}")
                    self.progress.emit(f"✓ Ready for Huawei Discovery upload!")
                    
                except Exception as e:
                    self.progress.emit(f"✗ ZIP creation failed: {str(e)}")
            
            # Final summary
            self.progress.emit(f"\n{'='*60}")
            self.progress.emit(f"CONVERSION SUMMARY")
            self.progress.emit(f"{'='*60}")
            self.progress.emit(f"Total files: {total_files}")
            self.progress.emit(f"Successful: {successful_files}")
            self.progress.emit(f"Failed: {failed_files}")
            
            if self.create_zip and self.output_files:
                self.progress.emit(f"ZIP archive: {zip_filename}")
            
            summary = f"✓ Batch conversion completed!\n\nSuccessful: {successful_files}/{total_files}"
            if self.create_zip and self.output_files:
                summary += f"\n\nZIP archive created and ready for Huawei Discovery!"
            
            self.finished.emit(successful_files > 0, summary)
            
        except Exception as e:
            self.finished.emit(False, f"Error: {str(e)}")


class UnifiedGeospatialTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_files = []
        self.output_file = ""
        self.merge_attributes = {}
        self.all_coords_data = []
        self.selected_polygon_indices = []
        
        # Variables for additional tabs
        self.loaded_layers = {}
        self.last_dissolved_file = None
        
        self.init_ui()
        
    def init_ui(self):
        """Initialize the UI with tabs"""
        self.setWindowTitle("Geospatial Tool (Fadzli Edition)")
        self.setGeometry(100, 100, 1400, 900)
        
        # Set only minimum size, allow maximizing
        self.setMinimumSize(1250, 800)
        
        # Create central widget with layout
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Create main tab widget
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Add footer with attribution
        footer = QLabel("V2.0.3525 | Written in Python with ❤️ | By Fadzli Abdullah | Huawei Technologies.")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setStyleSheet("""
            QLabel {
                background-color: #2a2a2a;
                color: #888888;
                padding: 8px;
                font-size: 11px;
                font-family: Ubuntu, sans-serif;
                font-weight: italic;
                border-top: 1px solid #444444;
            }
        """)
        main_layout.addWidget(footer)
        
        self.setCentralWidget(central_widget)
        
        # Create Tab 1: TAB Merger (original interface)
        self.create_tab_merger()
        
        # Create additional tabs
        if GEOPANDAS_AVAILABLE:
            self.create_tab_dissolve()
            self.create_tab_buffer()
            self.create_tab_simplify()
            self.create_tab_overlay()
            self.create_tab_converter()
            self.create_tab_edit()  # New edit tab for cutting/deleting geometry
        
        # Apply dark theme
        self.apply_dark_theme()
        
    def create_tab_merger(self):
        """Create the TAB Merger tab with original interface"""
        tab = QWidget()
        self.tabs.addTab(tab, "Polygon Merger")
        
        # Create layout for this tab
        tab_layout = QVBoxLayout(tab)
        
        title = QLabel("Multi-Format Polygon Merger")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tab_layout.addWidget(title)
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left panel
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        input_group = QGroupBox("Input Files")
        input_layout = QVBoxLayout()
        
        btn_layout = QHBoxLayout()
        self.btn_add_files = QPushButton("Add Files")
        self.btn_add_files.clicked.connect(self.add_files)
        self.btn_remove = QPushButton("Remove Selected")
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear = QPushButton("Clear All")
        self.btn_clear.clicked.connect(self.clear_all)
        
        btn_layout.addWidget(self.btn_add_files)
        btn_layout.addWidget(self.btn_remove)
        btn_layout.addWidget(self.btn_clear)
        input_layout.addLayout(btn_layout)
        
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        input_layout.addWidget(self.file_list)
        
        input_group.setLayout(input_layout)
        left_layout.addWidget(input_group)
        
        output_group = QGroupBox("Output File")
        output_layout = QVBoxLayout()
        
        output_btn_layout = QHBoxLayout()
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("Select output file location...")
        self.btn_browse_output = QPushButton("Browse...")
        self.btn_browse_output.clicked.connect(self.browse_output)
        
        output_btn_layout.addWidget(self.output_path)
        output_btn_layout.addWidget(self.btn_browse_output)
        output_layout.addLayout(output_btn_layout)
        
        output_group.setLayout(output_layout)
        left_layout.addWidget(output_group)
        
        action_layout = QHBoxLayout()
        self.btn_preview = QPushButton("Generate Preview")
        self.btn_preview.clicked.connect(self.generate_preview)
        self.btn_merge = QPushButton("Merge Selected Polygons")
        self.btn_merge.clicked.connect(self.merge_files)
        self.btn_merge.setEnabled(False)
        self.btn_export_shp = QPushButton("Export to Shapefile (ZIP)")
        self.btn_export_shp.clicked.connect(self.export_to_shapefile)
        self.btn_export_shp.setEnabled(False)
        
        action_layout.addWidget(self.btn_preview)
        action_layout.addWidget(self.btn_merge)
        action_layout.addWidget(self.btn_export_shp)
        left_layout.addLayout(action_layout)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        left_layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel("Ready")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(self.status_label)
        
        log_group = QGroupBox("Log")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        left_layout.addWidget(log_group)
        
        # Right panel
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout()
        
        self.preview_tabs = QTabWidget()
        
        # Map tab with polygon list
        map_tab_widget = QWidget()
        map_tab_layout = QVBoxLayout(map_tab_widget)
        
        # Polygon selection list
        poly_select_group = QGroupBox("Polygons (Click map or check boxes)")
        poly_select_layout = QVBoxLayout()
        
        poly_btn_layout = QHBoxLayout()
        self.btn_select_all = QPushButton("Select All")
        self.btn_select_all.clicked.connect(self.select_all_polygons)
        self.btn_deselect_all = QPushButton("Deselect All")
        self.btn_deselect_all.clicked.connect(self.deselect_all_polygons)
        self.btn_auto_save_all = QPushButton("Auto-Save All Separately")
        self.btn_auto_save_all.clicked.connect(self.auto_save_all_polygons)
        self.btn_auto_save_all.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.btn_save_selected = QPushButton("Save Selected Separately")
        self.btn_save_selected.clicked.connect(self.save_selected_polygons)
        self.btn_delete_selected = QPushButton("Delete Selected")
        self.btn_delete_selected.clicked.connect(self.delete_selected_polygons)
        
        poly_btn_layout.addWidget(self.btn_select_all)
        poly_btn_layout.addWidget(self.btn_deselect_all)
        poly_btn_layout.addWidget(self.btn_auto_save_all)
        poly_btn_layout.addWidget(self.btn_save_selected)
        poly_btn_layout.addWidget(self.btn_delete_selected)
        poly_select_layout.addLayout(poly_btn_layout)
        
        self.polygon_list = QListWidget()
        self.polygon_list.itemClicked.connect(self.polygon_list_clicked)
        poly_select_layout.addWidget(self.polygon_list)
        
        poly_select_group.setLayout(poly_select_layout)
        poly_select_group.setMaximumHeight(200)
        map_tab_layout.addWidget(poly_select_group)
        
        if MATPLOTLIB_AVAILABLE:
            self.map_canvas = MapCanvas(map_tab_widget, width=8, height=4)
            self.map_canvas.parent_gui = self
            map_tab_layout.addWidget(self.map_canvas)
            self.map_canvas.clear_map()
        else:
            map_tab_layout.addWidget(QLabel("Matplotlib not installed"))
            
        self.preview_tabs.addTab(map_tab_widget, "Map View")
        
        stats_tab = QWidget()
        stats_layout = QVBoxLayout(stats_tab)
        self.stats_text = QTextEdit()
        self.stats_text.setReadOnly(True)
        stats_layout.addWidget(self.stats_text)
        self.preview_tabs.addTab(stats_tab, "Statistics")
        
        preview_layout.addWidget(self.preview_tabs)
        preview_group.setLayout(preview_layout)
        right_layout.addWidget(preview_group)
        
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        
        tab_layout.addWidget(splitter)
        
        self.log("TAB Polygon Merger initialized")
        self.log("Click polygons on map to select/deselect for merging")
        
    def log(self, message):
        self.log_text.append(message)
    
    def remove_duplicate_points(self, coords, tolerance=1e-8):
        """Remove duplicate consecutive points from coordinate list"""
        if not coords or len(coords) < 2:
            return coords
        
        cleaned = [coords[0]]
        
        for i in range(1, len(coords)):
            x1, y1 = cleaned[-1]
            x2, y2 = coords[i]
            
            # Check if points are different (with tolerance)
            if abs(x1 - x2) > tolerance or abs(y1 - y2) > tolerance:
                cleaned.append(coords[i])
        
        # Ensure first and last points are same (closed polygon)
        if len(cleaned) > 2:
            x1, y1 = cleaned[0]
            x2, y2 = cleaned[-1]
            if abs(x1 - x2) > tolerance or abs(y1 - y2) > tolerance:
                cleaned.append(cleaned[0])
        
        return cleaned
        

    def apply_dark_theme(self):
        """Apply professional dark theme without blue frames"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
            }
            
            QWidget {
                background-color: #2b2b2b;
                color: #e0e0e0;
            }
            
            QGroupBox {
                font-weight: bold;
                border: 1px solid #555555;
                border-radius: 2px;
                margin-top: 12px;
                padding-top: 15px;
                background-color: #333333;
                color: #e0e0e0;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #b0b0b0;
            }
            
            QPushButton {
                background-color: #3d3d3d;
                color: #ffffff;
                border: 1px solid #555555;
                padding: 8px 16px;
                border-radius: 2px;
                font-weight: normal;
                min-height: 28px;
            }
            
            QPushButton:hover {
                background-color: #4a4a4a;
                border: 1px solid #666666;
            }
            
            QPushButton:pressed {
                background-color: #2a2a2a;
            }
            
            QPushButton:disabled {
                background-color: #2d2d2d;
                color: #666666;
                border: 1px solid #3d3d3d;
            }
            
            QListWidget {
                border: 1px solid #555555;
                border-radius: 2px;
                background-color: #1e1e1e;
                color: #e0e0e0;
                padding: 5px;
                selection-background-color: #0d47a1;
                selection-color: #ffffff;
            }
            
            QListWidget::item {
                color: #e0e0e0;
                padding: 4px;
            }
            
            QListWidget::item:selected {
                background-color: #0d47a1;
                color: #ffffff;
            }
            
            QTextEdit {
                border: 1px solid #555555;
                border-radius: 2px;
                background-color: #1e1e1e;
                color: #e0e0e0;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 9pt;
                padding: 5px;
            }
            
            QLineEdit, QDoubleSpinBox, QSpinBox, QComboBox {
                border: 1px solid #555555;
                border-radius: 2px;
                padding: 6px;
                background-color: #1e1e1e;
                color: #e0e0e0;
            }
            
            QLineEdit:focus, QDoubleSpinBox:focus, QSpinBox:focus {
                border: 1px solid #888888;
            }
            
            QComboBox::drop-down {
                border: none;
            }
            
            QComboBox QAbstractItemView {
                background-color: #1e1e1e;
                color: #e0e0e0;
                selection-background-color: #0d47a1;
            }
            
            QTabWidget::pane {
                border: 1px solid #555555;
                border-radius: 0px;
                background-color: #2b2b2b;
                top: -1px;
            }
            
            QTabBar::tab {
                background-color: #3d3d3d;
                color: #b0b0b0;
                padding: 10px 20px;
                margin-right: 2px;
                border: 1px solid #555555;
                border-bottom: none;
                border-top-left-radius: 0px;
                border-top-right-radius: 0px;
            }
            
            QTabBar::tab:selected {
                background-color: #2b2b2b;
                color: #ffffff;
                border-bottom: 1px solid #2b2b2b;
            }
            
            QTabBar::tab:!selected {
                margin-top: 2px;
            }
            
            QTabBar::tab:hover:!selected {
                background-color: #4a4a4a;
            }
            
            QProgressBar {
                border: 1px solid #555555;
                border-radius: 2px;
                text-align: center;
                background-color: #1e1e1e;
                color: #e0e0e0;
                height: 22px;
            }
            
            QProgressBar::chunk {
                background-color: #4caf50;
                border-radius: 1px;
            }
            
            QLabel {
                color: #e0e0e0;
                background-color: transparent;
            }
            
            QCheckBox {
                color: #e0e0e0;
                spacing: 8px;
            }
            
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 1px solid #555555;
                border-radius: 2px;
                background-color: #1e1e1e;
            }
            
            QCheckBox::indicator:checked {
                background-color: #0d47a1;
            }
            
            QCheckBox::indicator:hover {
                border: 1px solid #888888;
            }
            
            QSplitter::handle {
                background-color: #3d3d3d;
            }
            
            QSplitter::handle:horizontal {
                width: 2px;
            }
            
            QSplitter::handle:vertical {
                height: 2px;
            }
        """)
    
    def create_tab_dissolve(self):
        """Create dissolve tab with full functionality"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(10)
        
        # Info label
        info = QLabel("Dissolve: Convert multiple polygons into ONE single polygon (Huawei Discovery compatible)")
        info.setWordWrap(True)
        info.setStyleSheet("background-color: #4a4a4a; padding: 10px; border-radius: 2px;")
        layout.addWidget(info)
        
        # Input files
        input_group = QGroupBox("Input Files")
        input_layout = QVBoxLayout()
        
        self.dissolve_list = QListWidget()
        self.dissolve_list.setMinimumHeight(200)
        input_layout.addWidget(self.dissolve_list)
        
        btn_layout = QHBoxLayout()
        btn_add = QPushButton("Add Files")
        btn_add.clicked.connect(lambda: self.add_geop_files(self.dissolve_list))
        btn_clear = QPushButton("Clear")
        btn_clear.clicked.connect(lambda: self.dissolve_list.clear())
        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_clear)
        input_layout.addLayout(btn_layout)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        # Dissolve button
        btn_dissolve = QPushButton("Dissolve to Single Polygon")
        btn_dissolve.clicked.connect(self.run_dissolve_operation)
        btn_dissolve.setMinimumHeight(40)
        layout.addWidget(btn_dissolve)
        
        # Export button
        btn_export = QPushButton("Export to Shapefile ZIP")
        btn_export.clicked.connect(self.export_dissolved_to_shp)
        layout.addWidget(btn_export)
        
        layout.addStretch()
        self.tabs.addTab(tab, "Dissolve")
    
    def create_tab_buffer(self):
        """Create buffer tab with full functionality"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        info = QLabel("Buffer: Create zones around geometries at specified distance")
        info.setWordWrap(True)
        info.setStyleSheet("background-color: #4a4a4a; padding: 10px;")
        layout.addWidget(info)
        
        input_group = QGroupBox("Input Layer")
        input_layout = QVBoxLayout()
        self.buffer_file_label = QLabel("No file loaded")
        input_layout.addWidget(self.buffer_file_label)
        btn = QPushButton("Load File")
        btn.clicked.connect(lambda: self.load_single_file('buffer', self.buffer_file_label))
        input_layout.addWidget(btn)
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        param_group = QGroupBox("Parameters")
        param_layout = QFormLayout()
        self.buffer_distance = QDoubleSpinBox()
        self.buffer_distance.setRange(-10000, 10000)
        self.buffer_distance.setValue(10)
        param_layout.addRow("Distance (m):", self.buffer_distance)
        param_group.setLayout(param_layout)
        layout.addWidget(param_group)
        
        btn_run = QPushButton("Create Buffer")
        btn_run.setMinimumHeight(40)
        btn_run.clicked.connect(self.run_buffer_operation)
        layout.addWidget(btn_run)
        layout.addStretch()
        self.tabs.addTab(tab, "Buffer")
    
    def create_tab_simplify(self):
        """Create simplify tab with full functionality"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        info = QLabel("Simplify: Reduce vertex count while preserving shape")
        info.setStyleSheet("background-color: #4a4a4a; padding: 10px;")
        layout.addWidget(info)
        
        input_group = QGroupBox("Input Layer")
        input_layout = QVBoxLayout()
        self.simplify_file_label = QLabel("No file loaded")
        input_layout.addWidget(self.simplify_file_label)
        btn = QPushButton("Load File")
        btn.clicked.connect(lambda: self.load_single_file('simplify', self.simplify_file_label))
        input_layout.addWidget(btn)
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        param_group = QGroupBox("Parameters")
        param_layout = QFormLayout()
        self.simplify_tolerance = QDoubleSpinBox()
        self.simplify_tolerance.setRange(0.0001, 1000)
        self.simplify_tolerance.setValue(0.0001)  # Changed default to 0.0001 for more detail
        self.simplify_tolerance.setDecimals(4)
        param_layout.addRow("Tolerance:", self.simplify_tolerance)
        param_group.setLayout(param_layout)
        layout.addWidget(param_group)
        
        btn_run = QPushButton("Simplify Geometry")
        btn_run.setMinimumHeight(40)
        btn_run.clicked.connect(self.run_simplify_operation)
        layout.addWidget(btn_run)
        
        # Add Polygon Cleanup button
        cleanup_group = QGroupBox("Polygon Cleanup")
        cleanup_layout = QVBoxLayout()
        cleanup_info = QLabel("Remove holes/inner rings from polygons")
        cleanup_layout.addWidget(cleanup_info)
        btn_cleanup = QPushButton("Remove Polygon Holes")
        btn_cleanup.setMinimumHeight(40)
        btn_cleanup.clicked.connect(self.remove_polygon_holes)
        cleanup_layout.addWidget(btn_cleanup)
        cleanup_group.setLayout(cleanup_layout)
        layout.addWidget(cleanup_group)
        
        layout.addStretch()
        self.tabs.addTab(tab, "Simplify")
    
    def create_tab_overlay(self):
        """Create overlay tab with full functionality"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        info = QLabel("Overlay: Clip, Intersection, Difference, Union operations")
        info.setStyleSheet("background-color: #4a4a4a; padding: 10px;")
        layout.addWidget(info)
        
        op_group = QGroupBox("Select Operation")
        op_layout = QVBoxLayout()
        self.overlay_operation = QComboBox()
        self.overlay_operation.addItems([
            "Clip", "Intersection", "Difference", "Union"
        ])
        op_layout.addWidget(self.overlay_operation)
        op_group.setLayout(op_layout)
        layout.addWidget(op_group)
        
        layers_layout = QHBoxLayout()
        layer1_group = QGroupBox("Layer 1")
        layer1_layout = QVBoxLayout()
        self.overlay_layer1_label = QLabel("No file")
        layer1_layout.addWidget(self.overlay_layer1_label)
        btn1 = QPushButton("Load")
        btn1.clicked.connect(lambda: self.load_single_file('overlay1', self.overlay_layer1_label))
        layer1_layout.addWidget(btn1)
        layer1_group.setLayout(layer1_layout)
        layers_layout.addWidget(layer1_group)
        
        layer2_group = QGroupBox("Layer 2")
        layer2_layout = QVBoxLayout()
        self.overlay_layer2_label = QLabel("No file")
        layer2_layout.addWidget(self.overlay_layer2_label)
        btn2 = QPushButton("Load")
        btn2.clicked.connect(lambda: self.load_single_file('overlay2', self.overlay_layer2_label))
        layer2_layout.addWidget(btn2)
        layer2_group.setLayout(layer2_layout)
        layers_layout.addWidget(layer2_group)
        layout.addLayout(layers_layout)
        
        btn_run = QPushButton("Run Overlay Operation")
        btn_run.setMinimumHeight(40)
        btn_run.clicked.connect(self.run_overlay_operation)
        layout.addWidget(btn_run)
        layout.addStretch()
        self.tabs.addTab(tab, "Overlay")
    


    def create_tab_converter(self):
        """Create enhanced Converter tab with KMZ/KML support and professional GUI"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Header info
        info = QLabel("🔄 Converter: Convert KMZ/KML and other geospatial files to multiple formats")
        info.setWordWrap(True)
        info.setStyleSheet("background-color: #4a4a4a; padding: 12px; font-weight: bold; font-size: 11pt;")
        layout.addWidget(info)

        # Input files section
        input_group = QGroupBox("Input Files (KMZ, KML, SHP, TAB, GeoJSON)")
        input_layout = QVBoxLayout(input_group)
        input_layout.setSpacing(8)
        
        # File list
        self.converter_file_list = QListWidget()
        self.converter_file_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.converter_file_list.setMinimumHeight(120)
        input_layout.addWidget(self.converter_file_list)
        
        # File count label
        self.converter_file_count_label = QLabel("No files selected")
        self.converter_file_count_label.setStyleSheet("color: #b0b0b0; font-size: 10pt;")
        input_layout.addWidget(self.converter_file_count_label)
        
        # Buttons row
        btn_row = QHBoxLayout()
        btn_add_files = QPushButton("➕ Add Files")
        btn_add_files.clicked.connect(self.converter_add_files)
        btn_remove_selected = QPushButton("➖ Remove Selected")
        btn_remove_selected.clicked.connect(self.converter_remove_selected)
        self.converter_remove_btn = btn_remove_selected
        self.converter_remove_btn.setEnabled(False)
        btn_clear_all = QPushButton("🗑 Clear All")
        btn_clear_all.clicked.connect(self.converter_clear_all)
        self.converter_clear_btn = btn_clear_all
        self.converter_clear_btn.setEnabled(False)
        
        btn_row.addWidget(btn_add_files)
        btn_row.addWidget(btn_remove_selected)
        btn_row.addWidget(btn_clear_all)
        input_layout.addLayout(btn_row)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)

        # Output formats section
        format_group = QGroupBox("Output Formats (Select one or more)")
        format_layout = QVBoxLayout(format_group)
        format_layout.setSpacing(6)
        
        format_info = QLabel("Note: KML extraction from KMZ is automatic. SHP and TAB require GDAL/ogr2ogr.")
        format_info.setStyleSheet("color: #888; font-size: 9pt; font-style: italic;")
        format_info.setWordWrap(True)
        format_layout.addWidget(format_info)
        
        checkbox_row = QHBoxLayout()
        self.converter_kml_check = QCheckBox("KML (Keyhole Markup)")
        self.converter_kml_check.setChecked(True)
        self.converter_shp_check = QCheckBox("Shapefile (*.shp)")
        self.converter_shp_check.setChecked(True)
        self.converter_tab_check = QCheckBox("MapInfo TAB (*.tab)")
        
        checkbox_row.addWidget(self.converter_kml_check)
        checkbox_row.addWidget(self.converter_shp_check)
        checkbox_row.addWidget(self.converter_tab_check)
        checkbox_row.addStretch()
        format_layout.addLayout(checkbox_row)
        
        format_group.setLayout(format_layout)
        layout.addWidget(format_group)

        # Output directory section
        output_group = QGroupBox("Output Directory")
        output_layout = QHBoxLayout(output_group)
        output_layout.setSpacing(8)
        
        self.converter_output_label = QLabel("Not selected")
        self.converter_output_label.setStyleSheet("color: #b0b0b0; font-size: 10pt;")
        self.converter_output_label.setWordWrap(True)
        output_layout.addWidget(self.converter_output_label, 1)
        
        btn_browse_output = QPushButton("📁 Browse...")
        btn_browse_output.clicked.connect(self.converter_browse_output)
        btn_browse_output.setFixedWidth(100)
        output_layout.addWidget(btn_browse_output)
        
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)

        # Auto-ZIP option for Huawei Discovery
        zip_layout = QHBoxLayout()
        self.converter_auto_zip_check = QCheckBox("📦 Auto-create ZIP archive (for Huawei Discovery upload)")
        self.converter_auto_zip_check.setChecked(True)
        self.converter_auto_zip_check.setStyleSheet("font-weight: bold;")
        zip_layout.addWidget(self.converter_auto_zip_check)
        zip_layout.addStretch()
        layout.addLayout(zip_layout)

        # Progress bar
        self.converter_progress = QProgressBar()
        self.converter_progress.setVisible(False)
        self.converter_progress.setTextVisible(True)
        self.converter_progress.setFixedHeight(25)  # Fixed height to prevent resize
        layout.addWidget(self.converter_progress)

        # Convert button with Copy Log button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.converter_convert_btn = QPushButton("🔄 Start Conversion")
        self.converter_convert_btn.clicked.connect(self.converter_start_conversion)
        self.converter_convert_btn.setEnabled(False)
        self.converter_convert_btn.setMinimumHeight(40)
        self.converter_convert_btn.setStyleSheet("""
            QPushButton {
                background-color: #4caf50;
                color: white;
                font-weight: bold;
                font-size: 11pt;
                border-radius: 3px;
                padding: 8px 24px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QPushButton:disabled {
                background-color: #2d2d2d;
                color: #666666;
            }
        """)
        btn_layout.addWidget(self.converter_convert_btn)
        
        # Copy Log button next to Start Conversion
        self.converter_copy_log_btn = QPushButton("📋 Copy Log")
        self.converter_copy_log_btn.clicked.connect(self.converter_copy_log)
        self.converter_copy_log_btn.setMinimumHeight(40)
        self.converter_copy_log_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                font-size: 11pt;
                border-radius: 3px;
                padding: 8px 24px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
        """)
        btn_layout.addWidget(self.converter_copy_log_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Log section
        log_group = QGroupBox("Conversion Log")
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(5, 10, 5, 5)
        
        self.converter_log = QTextEdit()
        self.converter_log.setReadOnly(True)
        self.converter_log.setMinimumHeight(180)
        self.converter_log.setMaximumHeight(250)  # Allow reasonable height for scrolling
        self.converter_log.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.converter_log.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.converter_log.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)  # Better for log viewing
        self.converter_log.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #e0e0e0;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 9pt;
                border: 1px solid #555555;
            }
        """)
        log_layout.addWidget(self.converter_log)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group, 0)  # No stretch, fixed size

        # Initialize
        self.converter_input_files = []
        self.converter_output_dir = None
        self.converter_worker = None
        
        # Connect file list selection changed
        self.converter_file_list.itemSelectionChanged.connect(self.converter_on_selection_changed)
        
        self.tabs.addTab(tab, "Converter")
        self.log("Converter tab initialized - supports KMZ, KML, SHP, TAB, GeoJSON")

    def converter_add_files(self):
        """Add files to converter input list"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Input Files", "",
            "All Supported (*.kmz *.kml *.tab *.shp *.geojson *.gpkg *.gml);;KMZ Files (*.kmz);;KML Files (*.kml);;All Files (*.*)"
        )
        
        if files:
            for f in files:
                if f not in self.converter_input_files:
                    self.converter_input_files.append(f)
            
            self.converter_update_file_list()
            self.converter_log_message(f"Added {len(files)} file(s)")
    
    def converter_remove_selected(self):
        """Remove selected files from the list"""
        selected_items = self.converter_file_list.selectedItems()
        if not selected_items:
            return
        
        for item in selected_items:
            file_path = item.data(Qt.ItemDataRole.UserRole)
            if file_path in self.converter_input_files:
                self.converter_input_files.remove(file_path)
        
        self.converter_update_file_list()
        self.converter_log_message(f"Removed {len(selected_items)} file(s)")
    
    def converter_clear_all(self):
        """Clear all files"""
        if not self.converter_input_files:
            return
        
        reply = QMessageBox.question(
            self, "Clear All Files",
            f"Remove all {len(self.converter_input_files)} file(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            count = len(self.converter_input_files)
            self.converter_input_files.clear()
            self.converter_update_file_list()
            self.converter_log_message(f"Cleared {count} file(s)")
    
    def converter_update_file_list(self):
        """Update the file list widget"""
        self.converter_file_list.clear()
        
        for file_path in self.converter_input_files:
            item = QListWidgetItem(os.path.basename(file_path))
            item.setData(Qt.ItemDataRole.UserRole, file_path)
            item.setToolTip(file_path)
            self.converter_file_list.addItem(item)
        
        # Update file count label
        count = len(self.converter_input_files)
        if count == 0:
            self.converter_file_count_label.setText("No files selected")
            self.converter_convert_btn.setEnabled(False)
            self.converter_clear_btn.setEnabled(False)
        elif count == 1:
            self.converter_file_count_label.setText("1 file selected")
            self.converter_convert_btn.setEnabled(True)
            self.converter_clear_btn.setEnabled(True)
        else:
            self.converter_file_count_label.setText(f"{count} files selected")
            self.converter_convert_btn.setEnabled(True)
            self.converter_clear_btn.setEnabled(True)
    
    def converter_on_selection_changed(self):
        """Enable/disable remove button based on selection"""
        has_selection = len(self.converter_file_list.selectedItems()) > 0
        self.converter_remove_btn.setEnabled(has_selection and len(self.converter_input_files) > 0)
    
    def converter_browse_output(self):
        """Browse for output directory"""
        dir_path = QFileDialog.getExistingDirectory(
            self, "Select Output Directory"
        )
        if dir_path:
            self.converter_output_dir = dir_path
            self.converter_output_label.setText(dir_path)
            self.converter_output_label.setStyleSheet("color: #e0e0e0; font-size: 10pt;")
            self.converter_log_message(f"Output directory: {dir_path}")
    
    def converter_copy_log(self):
        """Copy conversion log to clipboard"""
        log_text = self.converter_log.toPlainText()
        clipboard = QApplication.clipboard()
        clipboard.setText(log_text)
        self.status_label.setText("Log copied to clipboard!")
        QMessageBox.information(self, "Log Copied", 
                               "Conversion log has been copied to clipboard.\n\n"
                               "You can paste it into a text file or share it for debugging.")
    
    def converter_log_message(self, message):
        """Add message to converter log without forcing scroll"""
        # Get current scroll position
        scrollbar = self.converter_log.verticalScrollBar()
        was_at_bottom = scrollbar.value() == scrollbar.maximum()
        
        # Add message
        self.converter_log.append(message)
        
        # Only auto-scroll if user was already at bottom
        if was_at_bottom:
            scrollbar.setValue(scrollbar.maximum())
        
        # Process events to keep UI responsive
        QApplication.processEvents()
    
    def converter_start_conversion(self):
        """Start the KMZ/KML conversion process"""
        if not self.converter_input_files:
            QMessageBox.warning(self, "No Input Files", 
                               "Please add at least one input file.")
            return
        
        # Get selected formats
        formats = []
        if self.converter_kml_check.isChecked():
            formats.append('kml')
        if self.converter_shp_check.isChecked():
            formats.append('shp')
        if self.converter_tab_check.isChecked():
            formats.append('tab')
        
        if not formats:
            QMessageBox.warning(self, "No Format Selected", 
                               "Please select at least one output format.")
            return
        
        # Determine output directory
        if not self.converter_output_dir:
            # Use the directory of the first file
            self.converter_output_dir = os.path.dirname(self.converter_input_files[0])
            self.converter_output_label.setText(self.converter_output_dir)
            self.converter_output_label.setStyleSheet("color: #e0e0e0; font-size: 10pt;")
        
        # Check for ogr2ogr if SHP or TAB is selected
        if 'shp' in formats or 'tab' in formats:
            ogr2ogr_available = False
            
            # If GDAL Python bindings are available in conda, trust that ogr2ogr is too
            if GDAL_AVAILABLE:
                self.converter_log_message("✓ GDAL Python bindings detected in conda environment")
                self.converter_log_message("  Assuming ogr2ogr is also available...")
                ogr2ogr_available = True
            else:
                # No GDAL Python, try to detect ogr2ogr directly
                try:
                    result = subprocess.run(['ogr2ogr', '--version'], 
                                          capture_output=True, timeout=5)
                    if result.returncode == 0:
                        ogr2ogr_available = True
                except:
                    pass
            
            if not ogr2ogr_available:
                msg = QMessageBox(self)
                msg.setIcon(QMessageBox.Icon.Critical)
                msg.setWindowTitle("GDAL/OGR Not Found")
                msg.setText("ogr2ogr (GDAL) is required for SHP and TAB conversion.")
                
                info_text = "Please install GDAL:\n\n"
                info_text += "• Conda: conda install -c conda-forge gdal\n"
                info_text += "• Linux: sudo apt-get install gdal-bin\n"
                info_text += "• Windows: Download from https://gdal.org\n"
                info_text += "• Mac: brew install gdal\n\n"
                info_text += "Or uncheck SHP and TAB formats to convert KML only."
                
                msg.setInformativeText(info_text)
                msg.exec()
                return
        
        # Disable controls during conversion
        self.converter_convert_btn.setEnabled(False)
        self.converter_remove_btn.setEnabled(False)
        self.converter_clear_btn.setEnabled(False)
        self.converter_progress.setVisible(True)
        self.converter_progress.setRange(0, len(self.converter_input_files))
        self.converter_progress.setValue(0)
        
        self.converter_log.clear()
        self.converter_log_message(f"🔄 Starting batch conversion...")
        # Scroll to top so user can see from the beginning
        self.converter_log.verticalScrollBar().setValue(0)
        self.converter_log_message(f"Total files: {len(self.converter_input_files)}")
        self.converter_log_message(f"Output directory: {self.converter_output_dir}")
        self.converter_log_message(f"Formats: {', '.join(formats).upper()}")
        if self.converter_auto_zip_check.isChecked():
            self.converter_log_message(f"📦 ZIP archive: Will be created after conversion")
        
        # Show environment info
        if GDAL_AVAILABLE:
            try:
                from osgeo import gdal
                self.converter_log_message(f"GDAL Version: {gdal.__version__}")
            except:
                pass
        
        self.converter_log_message("-" * 60)
        
        # Start worker thread
        self.converter_worker = KMZConversionWorker(
            self.converter_input_files, 
            self.converter_output_dir, 
            formats,
            create_zip=self.converter_auto_zip_check.isChecked()
        )
        self.converter_worker.progress.connect(self.converter_log_message)
        self.converter_worker.file_started.connect(self.converter_on_file_started)
        self.converter_worker.file_finished.connect(self.converter_on_file_finished)
        self.converter_worker.finished.connect(self.converter_conversion_finished)
        self.converter_worker.start()
        
        self.status_label.setText("Converting files...")
    
    def converter_on_file_started(self, filename):
        """Called when a new file starts processing"""
        self.status_label.setText(f"Processing: {filename}")
    
    def converter_on_file_finished(self, filename, success):
        """Called when a file finishes processing"""
        current = self.converter_progress.value()
        self.converter_progress.setValue(current + 1)
    
    def converter_conversion_finished(self, success, message):
        """Called when conversion is complete"""
        self.converter_progress.setVisible(False)
        self.converter_convert_btn.setEnabled(True)
        self.converter_clear_btn.setEnabled(len(self.converter_input_files) > 0)
        self.converter_on_selection_changed()
        
        self.converter_log_message("-" * 60)
        self.converter_log_message(message)
        
        if success:
            self.status_label.setText("Batch conversion completed!")
            QMessageBox.information(self, "Success", 
                                   f"{message}\n\nOutput directory:\n{self.converter_output_dir}")
        else:
            self.status_label.setText("Conversion failed")
            QMessageBox.warning(self, "Conversion Failed", message)
    
    def _browse_converter_input(self):
        """Legacy method for compatibility - redirects to new implementation"""
        self.converter_add_files()

    def _browse_converter_output(self):
        """Legacy method for compatibility - redirects to new implementation"""
        self.converter_browse_output()

    def start_converter(self):
        """Legacy method for compatibility - redirects to new implementation"""
        self.converter_start_conversion()
    
    def create_tab_edit(self):
        """Create edit tab for cutting and deleting geometry parts with interactive map"""
        tab = QWidget()
        main_layout = QHBoxLayout(tab)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(10)
        
        # Left panel - controls
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)
        left_panel.setMaximumWidth(280)
        left_panel.setMinimumWidth(250)
        
        info = QLabel("Edit: Cut and remove portions interactively")
        info.setStyleSheet("background-color: #4a4a4a; padding: 5px; border-radius: 2px;")
        info.setWordWrap(True)
        left_layout.addWidget(info)
        
        # Input file section
        input_group = QGroupBox("Input File")
        input_layout = QVBoxLayout(input_group)
        input_layout.setContentsMargins(5, 10, 5, 5)
        input_layout.setSpacing(3)
        
        self.edit_file_label = QLabel("No file loaded")
        self.edit_file_label.setStyleSheet("font-size: 10px;")
        input_layout.addWidget(self.edit_file_label)
        
        btn_load = QPushButton("Load File to Edit")
        btn_load.setFixedHeight(28)
        btn_load.clicked.connect(self.load_file_for_editing)
        input_layout.addWidget(btn_load)
        
        left_layout.addWidget(input_group, 0)
        
        # Feature list
        features_group = QGroupBox("Features")
        features_layout = QVBoxLayout(features_group)
        features_layout.setContentsMargins(5, 10, 5, 5)
        features_layout.setSpacing(3)
        
        self.edit_feature_list = QListWidget()
        self.edit_feature_list.setFixedHeight(60)
        self.edit_feature_list.itemClicked.connect(self.preview_edit_feature)
        features_layout.addWidget(self.edit_feature_list)
        
        left_layout.addWidget(features_group, 0)
        
        # Drawing tools
        draw_group = QGroupBox("Interactive Drawing")
        draw_layout = QVBoxLayout(draw_group)
        draw_layout.setContentsMargins(5, 10, 5, 5)
        draw_layout.setSpacing(3)
        
        draw_info = QLabel("Click map to draw polygon")
        draw_info.setWordWrap(True)
        draw_info.setStyleSheet("font-size: 9px;")
        draw_layout.addWidget(draw_info)
        
        btn_start_draw = QPushButton("Start Drawing")
        btn_start_draw.setFixedHeight(28)
        btn_start_draw.clicked.connect(self.start_polygon_drawing)
        draw_layout.addWidget(btn_start_draw)
        
        btn_clear_draw = QPushButton("Clear Drawing")
        btn_clear_draw.setFixedHeight(28)
        btn_clear_draw.clicked.connect(self.clear_polygon_drawing)
        draw_layout.addWidget(btn_clear_draw)
        
        self.draw_points_label = QLabel("Points: 0")
        self.draw_points_label.setStyleSheet("font-size: 9px;")
        draw_layout.addWidget(self.draw_points_label)
        
        left_layout.addWidget(draw_group, 0)
        
        # Cut operations
        cut_group = QGroupBox("Cut Operations")
        cut_layout = QVBoxLayout(cut_group)
        cut_layout.setContentsMargins(5, 10, 5, 5)
        cut_layout.setSpacing(3)
        
        cut_info = QLabel("Draw polygon, then cut to remove area")
        cut_info.setWordWrap(True)
        cut_info.setStyleSheet("font-size: 9px;")
        cut_layout.addWidget(cut_info)
        
        btn_cut_polygon = QPushButton("Cut Drawn Polygon")
        btn_cut_polygon.setFixedHeight(32)
        btn_cut_polygon.clicked.connect(self.cut_geometry_polygon)
        cut_layout.addWidget(btn_cut_polygon)
        
        left_layout.addWidget(cut_group, 0)
        
        # Save buttons
        save_group = QGroupBox("Save Options")
        save_layout = QVBoxLayout(save_group)
        save_layout.setContentsMargins(5, 10, 5, 5)
        save_layout.setSpacing(3)
        
        btn_save_multi = QPushButton("Choose Format")
        btn_save_multi.setFixedHeight(28)
        btn_save_multi.clicked.connect(self.save_edited_file_multi_format)
        save_layout.addWidget(btn_save_multi)
        
        btn_save_shp_zip = QPushButton("Save as Shapefile & ZIP")
        btn_save_shp_zip.setFixedHeight(28)
        btn_save_shp_zip.clicked.connect(self.save_edited_file_shapefile_zip)
        save_layout.addWidget(btn_save_shp_zip)
        
        left_layout.addWidget(save_group, 0)
        
        # Add stretch at the bottom
        left_layout.addStretch(1)
        
        # Right panel - Map preview
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)
        
        map_label = QLabel("Map Preview")
        map_label.setStyleSheet("background-color: #4a4a4a; padding: 5px; font-weight: bold;")
        map_label.setFixedHeight(25)
        right_layout.addWidget(map_label)
        
        # Create matplotlib canvas for map
        if MATPLOTLIB_AVAILABLE:
            self.edit_map_canvas = MapCanvas(self, width=8, height=7)
            self.edit_map_canvas.mpl_connect('button_press_event', self.on_map_click)
            right_layout.addWidget(self.edit_map_canvas, 1)
        else:
            no_map_label = QLabel("Map preview unavailable\nInstall matplotlib")
            no_map_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            right_layout.addWidget(no_map_label)
        
        # Add panels to main layout
        main_layout.addWidget(left_panel, 0)
        main_layout.addWidget(right_panel, 1)
        
        # Initialize drawing state
        self.edit_drawing_active = False
        self.edit_drawn_points = []
        self.edit_selected_feature_idx = None
        
        self.tabs.addTab(tab, "Edit")


    # ====================================================================
    # GeoPandas Operations Support Methods
    # ====================================================================
    
    def add_geop_files(self, list_widget):
        """Add files to a list widget for GeoPandas operations"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Files", "",
            "All Supported (*.tab *.shp *.geojson *.gpkg);;TAB Files (*.tab);;Shapefile (*.shp);;GeoJSON (*.geojson);;GeoPackage (*.gpkg);;All Files (*)"
        )
        
        if files:
            for file in files:
                list_widget.addItem(file)
            self.log(f"Added {len(files)} file(s)")
    
    def load_single_file(self, operation_type, label_widget):
        """Load a single file for operations like simplify, buffer, etc."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select File", "",
            "All Supported (*.tab *.shp *.geojson *.gpkg *.kml *.kmz);;TAB Files (*.tab);;Shapefile (*.shp);;GeoJSON (*.geojson);;KML (*.kml *.kmz);;GeoPackage (*.gpkg);;All Files (*)"
        )
        
        if file_path:
            # Store the file path based on operation type
            if operation_type == 'simplify':
                self.simplify_file_path = file_path
            elif operation_type == 'buffer':
                self.buffer_file_path = file_path
            elif operation_type == 'overlay1':
                self.overlay1_file_path = file_path
            elif operation_type == 'overlay2':
                self.overlay2_file_path = file_path
            elif operation_type == 'overlay':
                # generic overlay
                self.overlay_file_path = file_path
            
            # Update the label
            label_widget.setText(os.path.basename(file_path))
            self.log(f"Loaded: {os.path.basename(file_path)}")
    
    def run_dissolve_operation(self):
        """Run dissolve operation with custom name and merge strategy"""
        if self.dissolve_list.count() == 0:
            QMessageBox.warning(self, "Warning", "Please add files to dissolve")
            return
        
        # Show dialog to get custom name and merge strategy
        dialog = QDialog(self)
        dialog.setWindowTitle("Dissolve Options")
        dialog.setModal(True)
        layout = QVBoxLayout(dialog)
        
        # Polygon name input
        name_label = QLabel("Polygon Name (for Huawei Discovery):")
        layout.addWidget(name_label)
        
        name_input = QLineEdit()
        name_input.setPlaceholderText("e.g., COVERAGE_AREA_EAST")
        from datetime import datetime
        default_name = f"DISSOLVED_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        name_input.setText(default_name)
        layout.addWidget(name_input)
        
        # Merge strategy selection
        strategy_label = QLabel("\nMerge Strategy:")
        layout.addWidget(strategy_label)
        
        strategy_combo = QComboBox()
        strategy_combo.addItem("Convex Hull (creates bounding polygon)", "convex_hull")
        strategy_combo.addItem("Buffer Method (connects nearby parts)", "buffer")
        strategy_combo.addItem("Keep Largest Only (discards smaller parts)", "largest")
        layout.addWidget(strategy_combo)
        
        # Info labels
        info_label = QLabel(
            "\n• Convex Hull: Creates one continuous polygon covering all areas\n"
            "• Buffer: Attempts to connect separate parts with small buffer\n"
            "• Keep Largest: Keeps only the biggest polygon, removes others"
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        
        # Get user inputs
        custom_name = name_input.text().strip()
        if not custom_name:
            QMessageBox.warning(self, "Warning", "Please enter a polygon name")
            return
        
        merge_strategy = strategy_combo.currentData()
        
        # Get input files
        input_files = [self.dissolve_list.item(i).text() 
                      for i in range(self.dissolve_list.count())]
        
        # Get output file
        output_file, _ = QFileDialog.getSaveFileName(
            self, "Save Dissolved Result", "",
            "TAB File (*.tab);;Shapefile (*.shp);;GeoJSON (*.geojson);;GeoPackage (*.gpkg);;All Files (*)"
        )
        
        if output_file and GEOPANDAS_AVAILABLE:
            self.last_dissolved_file = output_file
            params = {
                'input_files': input_files,
                'output_file': output_file,
                'custom_name': custom_name,
                'merge_strategy': merge_strategy
            }
            self.start_geop_processing("dissolve", params)
        elif not GEOPANDAS_AVAILABLE:
            QMessageBox.critical(self, "Error", "GeoPandas not available. Install: conda install -c conda-forge geopandas")
    
    def export_dissolved_to_shp(self):
        """Export last dissolved file to shapefile ZIP"""
        if not self.last_dissolved_file or not os.path.exists(self.last_dissolved_file):
            QMessageBox.warning(self, "Warning", "No dissolved file. Run Dissolve first.")
            return
        
        zip_path, _ = QFileDialog.getSaveFileName(
            self, "Save Shapefile ZIP",
            str(Path(self.last_dissolved_file).with_suffix('.zip')),
            "ZIP Files (*.zip)"
        )
        
        if zip_path and GEOPANDAS_AVAILABLE:
            if not zip_path.lower().endswith('.zip'):
                zip_path += '.zip'
            
            params = {
                'input_file': self.last_dissolved_file,
                'output_zip': zip_path
            }
            self.start_geop_processing("convert_to_shapefile_zip", params)
    
    def start_geop_processing(self, operation, params):
        """Start GeoPandas processing thread"""
        self.processing_thread = ProcessingThread(operation, params)
        self.processing_thread.progress.connect(self.update_geop_progress)
        self.processing_thread.status.connect(self.update_geop_status)
        self.processing_thread.finished.connect(self.geop_processing_finished)
        
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.log(f"Starting {operation} operation...")
        self.processing_thread.start()
    
    def update_geop_progress(self, value):
        """Update progress bar for GeoPandas operations"""
        self.progress_bar.setValue(value)
    
    def update_geop_status(self, message):
        """Update status for GeoPandas operations"""
        self.status_label.setText(message)
        self.log(message)
    
    def geop_processing_finished(self, success, message):
        """Handle GeoPandas processing completion"""
        self.progress_bar.setVisible(False)
        
        if success:
            QMessageBox.information(self, "Success", message)
            self.log("✓ " + message.split('\n')[0])
            self.progress_bar.setValue(100)
        else:
            QMessageBox.critical(self, "Error", message)
            self.log("✗ " + message)
            self.progress_bar.setValue(0)
        
        self.status_label.setText("Ready")

    def run_buffer_operation(self):
        """Collect parameters from Buffer tab and start buffer operation"""
        if not hasattr(self, 'buffer_file_path') or not self.buffer_file_path:
            QMessageBox.warning(self, "Warning", "Please load a file first")
            return

        # Ask for output file
        default_name = os.path.splitext(os.path.basename(self.buffer_file_path))[0] + "_buffered" + os.path.splitext(self.buffer_file_path)[1]
        output_file, _ = QFileDialog.getSaveFileName(
            self, "Save Buffer File", default_name,
            "TAB File (*.tab);;Shapefile (*.shp);;GeoJSON (*.geojson);;GeoPackage (*.gpkg);;All Files (*)"
        )

        if not output_file:
            return

        params = {
            'input_layer': self.buffer_file_path,
            'output_file': output_file,
            'distance': float(self.buffer_distance.value()),
            'dissolve': False
        }

        self.start_geop_processing('buffer', params)

    def run_overlay_operation(self):
        """Start overlay operation (clip/intersection/difference/union) based on UI selection"""
        # Ensure layers loaded
        if not hasattr(self, 'overlay_layer1_label') or not hasattr(self, 'overlay_layer2_label'):
            QMessageBox.warning(self, "Warning", "Please load both layers first")
            return

        # We stored file paths in load_single_file into attributes like overlay1/overlay2
        file1 = getattr(self, 'overlay1_file_path', None)
        file2 = getattr(self, 'overlay2_file_path', None)

        if not file1 or not file2:
            QMessageBox.warning(self, "Warning", "Please select both Layer 1 and Layer 2 files")
            return

        # Select output file
        default_name = os.path.splitext(os.path.basename(file1))[0] + f"_{self.overlay_operation.currentText().lower()}" + os.path.splitext(file1)[1]
        output_file, _ = QFileDialog.getSaveFileName(
            self, "Save Overlay Result", default_name,
            "TAB File (*.tab);;Shapefile (*.shp);;GeoJSON (*.geojson);;GeoPackage (*.gpkg);;All Files (*)"
        )

        if not output_file:
            return

        op_map = {
            'Clip': 'clip',
            'Intersection': 'intersection',
            'Difference': 'difference',
            'Union': 'union'
        }

        selected_ui = self.overlay_operation.currentText()
        operation_key = op_map.get(selected_ui, 'clip')

        params = {
            'input_layer': file1,
            'overlay_layer': file2,
            'output_file': output_file
        }

        self.start_geop_processing(operation_key, params)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Geospatial Files", "",
            "All Supported (*.tab *.shp *.geojson *.gpkg *.gml *.json);;MapInfo TAB (*.tab);;ESRI Shapefile (*.shp);;GeoJSON (*.geojson);;GeoPackage (*.gpkg);;GML (*.gml);;JSON (*.json);;All Files (*.*)"
        )
        
        if files:
            for filepath in files:
                if filepath not in self.input_files:
                    self.input_files.append(filepath)
                    self.file_list.addItem(Path(filepath).name)
            self.log(f"Added {len(files)} file(s)")
            self.update_button_states()
            
    def remove_selected(self):
        for item in self.file_list.selectedItems():
            row = self.file_list.row(item)
            self.file_list.takeItem(row)
            self.input_files.pop(row)
        self.update_button_states()
        
    def clear_all(self):
        self.file_list.clear()
        self.input_files.clear()
        self.polygon_list.clear()
        self.selected_polygon_indices = []
        self.update_button_states()
        
    def browse_output(self):
        filepath, filter_selected = QFileDialog.getSaveFileName(
            self, "Save Merged File", "",
            "MapInfo TAB (*.tab);;ESRI Shapefile (*.shp);;GeoJSON (*.geojson);;GeoPackage (*.gpkg);;GML (*.gml)"
        )
        
        if filepath:
            # Ensure file has proper extension based on selection
            ext_map = {
                'MapInfo TAB': '.tab',
                'ESRI Shapefile': '.shp',
                'GeoJSON': '.geojson',
                'GeoPackage': '.gpkg',
                'GML': '.gml'
            }
            
            # Extract extension from filter if available
            for filter_name, ext in ext_map.items():
                if filter_name in filter_selected:
                    if not filepath.lower().endswith(ext):
                        filepath = filepath + ext
                    break
            else:
                # Default to .tab if no filter match
                if not any(filepath.lower().endswith(ext) for ext in ext_map.values()):
                    filepath += '.tab'
            
            self.output_file = filepath
            self.output_path.setText(filepath)
            self.log(f"Output: {Path(filepath).name}")
            self.update_button_states()
            
    def update_button_states(self):
        has_files = len(self.input_files) > 0
        has_output = bool(self.output_file)
        self.btn_preview.setEnabled(has_files)
        self.btn_merge.setEnabled(has_files and has_output)
        
        if has_output and os.path.exists(self.output_file):
            self.btn_export_shp.setEnabled(True)
        else:
            self.btn_export_shp.setEnabled(False)
            
    def generate_preview(self):
        if not self.input_files:
            return
            
        self.log("Generating preview...")
        self.progress_bar.setVisible(True)
        self.status_label.setText("Generating preview...")
        
        self.worker = MergeWorker(self.input_files, "")
        self.worker.set_preview_mode(True)
        self.worker.progress.connect(self.update_progress)
        self.worker.preview_ready.connect(self.show_preview)
        self.worker.log_message.connect(self.log)  # Connect log signal
        self.worker.start()
        
    def show_preview(self, preview_data, stats, all_coords):
        self.all_coords_data = all_coords
        colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#FFA07A', '#98D8C8', 
                  '#F7DC6F', '#BB8FCE']
        
        stats_text = f"Total Input Features: {stats['total_features']}\n"
        stats_text += f"Total Polygons: {stats['total_polygons']}\n"
        stats_text += f"Geometry Types: {', '.join(stats['geometry_types'])}\n"
        stats_text += f"Coordinate Systems: {', '.join(set(stats['srs_list']))}\n\n"
        stats_text += f"Click polygons on map to select which ones to merge.\n"
        stats_text += f"Or use 'Select All' to merge all polygons."
        
        self.stats_text.setText(stats_text)
        
        # Populate polygon list
        self.polygon_list.clear()
        polygon_count = 0
        for filename, polys in all_coords:
            for i in range(len(polys)):
                item = QListWidgetItem(f"Polygon {polygon_count + 1} ({filename})")
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Unchecked)
                self.polygon_list.addItem(item)
                polygon_count += 1
        
        if MATPLOTLIB_AVAILABLE and self.map_canvas and all_coords:
            self.map_canvas.plot_geometries(all_coords, colors)
            self.log(f"Map shows {stats['total_polygons']} polygons")
            self.log("Click polygons on map or check boxes to select")
        
        self.progress_bar.setVisible(False)
        self.status_label.setText("Preview ready - Select polygons to merge")
        self.log("Preview complete")
        self.preview_tabs.setCurrentIndex(0)
        
    def select_polygon_from_map(self, idx):
        """Called when polygon is clicked on map"""
        if idx >= self.polygon_list.count():
            return
        
        item = self.polygon_list.item(idx)
        if item.checkState() == Qt.CheckState.Checked:
            item.setCheckState(Qt.CheckState.Unchecked)
            if idx in self.selected_polygon_indices:
                self.selected_polygon_indices.remove(idx)
        else:
            item.setCheckState(Qt.CheckState.Checked)
            if idx not in self.selected_polygon_indices:
                self.selected_polygon_indices.append(idx)
        
        self.update_map_highlighting()
        self.log(f"Polygon {idx + 1}: {'Selected' if idx in self.selected_polygon_indices else 'Deselected'}")
        
    def polygon_list_clicked(self, item):
        """Handle polygon list checkbox click"""
        idx = self.polygon_list.row(item)
        
        if item.checkState() == Qt.CheckState.Checked:
            if idx not in self.selected_polygon_indices:
                self.selected_polygon_indices.append(idx)
        else:
            if idx in self.selected_polygon_indices:
                self.selected_polygon_indices.remove(idx)
        
        self.update_map_highlighting()
        
    def select_all_polygons(self):
        """Select all polygons"""
        self.selected_polygon_indices = list(range(self.polygon_list.count()))
        for i in range(self.polygon_list.count()):
            self.polygon_list.item(i).setCheckState(Qt.CheckState.Checked)
        self.update_map_highlighting()
        self.log(f"Selected all {len(self.selected_polygon_indices)} polygons")
        
    def deselect_all_polygons(self):
        """Deselect all polygons"""
        self.selected_polygon_indices = []
        for i in range(self.polygon_list.count()):
            self.polygon_list.item(i).setCheckState(Qt.CheckState.Unchecked)
        self.update_map_highlighting()
        self.log("Deselected all polygons")
        
    def update_map_highlighting(self):
        """Update map to show selected polygons"""
        if MATPLOTLIB_AVAILABLE and self.map_canvas:
            self.map_canvas.highlight_selected(self.selected_polygon_indices)
    
    def auto_save_all_polygons(self):
        """Automatically save ALL polygons separately with sequential naming"""
        if not self.map_canvas.polygon_data:
            QMessageBox.warning(self, "No Polygons", "Generate preview first.")
            return
        
        # Ask for output folder
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder for All Polygons")
        if not folder:
            return
        
        # Ask for base name
        dialog = AttributeDialog(self)
        dialog.setWindowTitle("Set Base Name")
        dialog.name_input.setPlaceholderText("e.g., E2_PLUS")
        
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        
        attrs = dialog.get_attributes()
        base_name = attrs['Name']
        group = attrs['Group']
        
        try:
            total_polygons = len(self.map_canvas.polygon_data)
            self.log(f"Auto-saving ALL {total_polygons} polygons separately...")
            self.status_label.setText(f"Saving {total_polygons} polygons...")
            
            first_ds = ogr.Open(self.input_files[0])
            if not first_ds:
                raise Exception("Could not open first file")
            
            first_layer = first_ds.GetLayer(0)
            srs = first_layer.GetSpatialRef()
            first_ds = None
            
            driver = ogr.GetDriverByName('MapInfo File')
            saved_files = []
            
            for idx in range(total_polygons):
                poly_info = self.map_canvas.polygon_data[idx]
                coords = poly_info['coords']
                
                # Remove duplicate points - aggressive cleaning
                cleaned_coords = []
                if coords:
                    cleaned_coords.append(coords[0])
                    removed = 0
                    
                    for i in range(1, len(coords)):
                        x1, y1 = cleaned_coords[-1]
                        x2, y2 = coords[i]
                        
                        dist = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
                        dx = abs(x2 - x1)
                        dy = abs(y2 - y1)
                        
                        # Point must be different by at least 0.00001 degrees
                        if dist > 0.00001 or (dx > 0.00001 or dy > 0.00001):
                            cleaned_coords.append(coords[i])
                        else:
                            removed += 1
                    
                    # Ensure closed
                    if len(cleaned_coords) > 2:
                        x1, y1 = cleaned_coords[0]
                        x2, y2 = cleaned_coords[-1]
                        dist = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
                        
                        if dist > 0.00001:
                            cleaned_coords.append(cleaned_coords[0])
                        elif len(cleaned_coords) > 3 and dist > 0:
                            cleaned_coords[-1] = cleaned_coords[0]
                    
                    if removed > 0:
                        self.log(f"  Polygon {idx+1}: Removed {removed} duplicate points")
                
                coords = cleaned_coords
                
                # Validate point count
                unique_points = len(coords) - 1 if coords and coords[0] == coords[-1] else len(coords)
                
                if len(coords) < 4 or unique_points < 3:
                    self.log(f"⚠ Polygon {idx + 1} has illegal point count ({len(coords)} points, {unique_points} unique), skipping")
                    continue
                
                output_name = f"{base_name}_{idx+1}.tab"
                polygon_name = f"{base_name}_{idx+1}"
                output_path = os.path.join(folder, output_name)
                
                if os.path.exists(output_path):
                    driver.DeleteDataSource(output_path)
                
                out_ds = driver.CreateDataSource(output_path)
                out_layer = out_ds.CreateLayer('polygon', srs, ogr.wkbPolygon)
                
                name_field = ogr.FieldDefn('Name', ogr.OFTString)
                name_field.SetWidth(254)
                out_layer.CreateField(name_field)
                
                group_field = ogr.FieldDefn('Group', ogr.OFTString)
                group_field.SetWidth(254)
                out_layer.CreateField(group_field)
                
                ring = ogr.Geometry(ogr.wkbLinearRing)
                for x, y in coords:
                    ring.AddPoint(x, y)
                
                polygon = ogr.Geometry(ogr.wkbPolygon)
                polygon.AddGeometry(ring)
                
                out_feature = ogr.Feature(out_layer.GetLayerDefn())
                out_feature.SetGeometry(polygon)
                out_feature.SetField('Name', polygon_name)
                out_feature.SetField('Group', group)
                
                out_layer.CreateFeature(out_feature)
                out_ds = None
                
                saved_files.append(output_name)
                self.log(f"✓ Saved: {output_name}")
            
            self.status_label.setText("Auto-save complete!")
            QMessageBox.information(self, "Success", 
                                   f"Auto-saved ALL {len(saved_files)} polygons to:\n{folder}")
            
        except Exception as e:
            error_msg = f"Auto-save failed: {str(e)}\n{traceback.format_exc()}"
            self.log(error_msg)
            self.status_label.setText("Auto-save failed")
            QMessageBox.critical(self, "Error", error_msg)
    
    def save_selected_polygons(self):
        """Save selected polygons as separate TAB files"""
        if not self.selected_polygon_indices:
            QMessageBox.warning(self, "No Selection", "Please select polygons to save first.")
            return
        
        # Ask for output folder
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not folder:
            return
        
        # Ask for base name
        dialog = AttributeDialog(self)
        dialog.setWindowTitle("Set Base Name for Saved Polygons")
        dialog.name_input.setPlaceholderText("e.g., E2_PLUS")
        
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        
        attrs = dialog.get_attributes()
        base_name = attrs['Name']
        group = attrs['Group']
        
        try:
            self.log(f"Saving {len(self.selected_polygon_indices)} selected polygon(s) separately...")
            self.status_label.setText("Saving polygons...")
            
            # Get SRS from first input file
            first_ds = ogr.Open(self.input_files[0])
            if not first_ds:
                raise Exception("Could not open first file")
            
            first_layer = first_ds.GetLayer(0)
            srs = first_layer.GetSpatialRef()
            first_ds = None
            
            driver = ogr.GetDriverByName('MapInfo File')
            saved_files = []
            
            # Save each selected polygon from map canvas data
            for i, poly_idx in enumerate(sorted(self.selected_polygon_indices)):
                if poly_idx >= len(self.map_canvas.polygon_data):
                    continue
                
                poly_info = self.map_canvas.polygon_data[poly_idx]
                coords = poly_info['coords']
                
                self.log(f"Processing Polygon {poly_idx + 1}: {len(coords)} points")
                
                # Remove duplicate points - more aggressive
                cleaned_coords = []
                if coords:
                    cleaned_coords.append(coords[0])
                    removed = 0
                    
                    for i in range(1, len(coords)):
                        x1, y1 = cleaned_coords[-1]
                        x2, y2 = coords[i]
                        
                        # Use both distance and coordinate difference
                        dist = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
                        dx = abs(x2 - x1)
                        dy = abs(y2 - y1)
                        
                        # Point must be different by at least 0.00001 degrees (~1 meter)
                        if dist > 0.00001 or (dx > 0.00001 or dy > 0.00001):
                            cleaned_coords.append(coords[i])
                        else:
                            removed += 1
                    
                    # Ensure closed polygon - first and last must be same
                    if len(cleaned_coords) > 2:
                        x1, y1 = cleaned_coords[0]
                        x2, y2 = cleaned_coords[-1]
                        dist = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
                        
                        if dist > 0.00001:
                            # Not closed, add first point
                            cleaned_coords.append(cleaned_coords[0])
                        elif len(cleaned_coords) > 3 and dist > 0:
                            # Very close but not exact - force exact match
                            cleaned_coords[-1] = cleaned_coords[0]
                    
                    if removed > 0:
                        self.log(f"  Removed {removed} duplicate points")
                
                coords = cleaned_coords
                
                # Validate point count for Discovery
                # Discovery requires at least 3 unique points + closing point = 4 total
                unique_points = len(coords) - 1 if coords and coords[0] == coords[-1] else len(coords)
                
                if len(coords) < 4:
                    self.log(f"⚠ Polygon {poly_idx + 1} has only {len(coords)} points (need at least 4), skipping")
                    continue
                
                if unique_points < 3:
                    self.log(f"⚠ Polygon {poly_idx + 1} has only {unique_points} unique points (need at least 3), skipping")
                    continue
                
                self.log(f"  Final: {len(coords)} points ({unique_points} unique)")
                
                # Create output filename
                if len(self.selected_polygon_indices) > 1:
                    output_name = f"{base_name}_{i+1}.tab"
                    polygon_name = f"{base_name}_{i+1}"
                else:
                    output_name = f"{base_name}.tab"
                    polygon_name = base_name
                    
                output_path = os.path.join(folder, output_name)
                
                # Remove if exists
                if os.path.exists(output_path):
                    driver.DeleteDataSource(output_path)
                
                # Create output datasource
                out_ds = driver.CreateDataSource(output_path)
                out_layer = out_ds.CreateLayer('polygon', srs, ogr.wkbPolygon)
                
                # Add fields
                name_field = ogr.FieldDefn('Name', ogr.OFTString)
                name_field.SetWidth(254)
                out_layer.CreateField(name_field)
                
                group_field = ogr.FieldDefn('Group', ogr.OFTString)
                group_field.SetWidth(254)
                out_layer.CreateField(group_field)
                
                # Create polygon geometry from coordinates
                ring = ogr.Geometry(ogr.wkbLinearRing)
                for x, y in coords:
                    ring.AddPoint(x, y)
                
                polygon = ogr.Geometry(ogr.wkbPolygon)
                polygon.AddGeometry(ring)
                
                # Create feature
                out_feature = ogr.Feature(out_layer.GetLayerDefn())
                out_feature.SetGeometry(polygon)
                out_feature.SetField('Name', polygon_name)
                out_feature.SetField('Group', group)
                
                out_layer.CreateFeature(out_feature)
                out_ds = None
                
                saved_files.append(output_name)
                self.log(f"✓ Saved: {output_name}")
            
            self.status_label.setText("Save complete!")
            QMessageBox.information(self, "Success", 
                                   f"Saved {len(saved_files)} polygon(s) to:\n{folder}\n\n" + 
                                   "\n".join(saved_files))
            
        except Exception as e:
            error_msg = f"Save failed: {str(e)}\n{traceback.format_exc()}"
            self.log(error_msg)
            self.status_label.setText("Save failed")
            QMessageBox.critical(self, "Error", error_msg)
    
    def delete_selected_polygons(self):
        """Remove selected polygons from the preview"""
        if not self.selected_polygon_indices:
            QMessageBox.warning(self, "No Selection", "Please select polygons to delete first.")
            return
        
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Delete {len(self.selected_polygon_indices)} selected polygon(s) from preview?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply != QMessageBox.StandardButton.Yes:
            return
        
        try:
            # Remove from highest index to lowest to avoid index shifting
            for idx in sorted(self.selected_polygon_indices, reverse=True):
                if idx < self.polygon_list.count():
                    self.polygon_list.takeItem(idx)
                    
                # Remove from map canvas polygon data
                if MATPLOTLIB_AVAILABLE and self.map_canvas:
                    if idx < len(self.map_canvas.polygon_data):
                        del self.map_canvas.polygon_data[idx]
            
            self.log(f"Deleted {len(self.selected_polygon_indices)} polygon(s) from preview")
            
            # Clear selection
            self.selected_polygon_indices = []
            
            # Redraw map
            if MATPLOTLIB_AVAILABLE and self.map_canvas:
                colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#FFA07A', '#98D8C8', 
                          '#F7DC6F', '#BB8FCE']
                
                # Rebuild display from remaining polygon data
                self.map_canvas.axes.clear()
                all_x = []
                all_y = []
                
                for poly_info in self.map_canvas.polygon_data:
                    coords = poly_info['coords']
                    x_coords = [c[0] for c in coords]
                    y_coords = [c[1] for c in coords]
                    
                    polygon = MplPolygon(list(zip(x_coords, y_coords)), 
                                        facecolor=poly_info['color'], 
                                        edgecolor='black', 
                                        linewidth=1, alpha=0.6, picker=5)
                    self.map_canvas.axes.add_patch(polygon)
                    
                    all_x.extend(x_coords)
                    all_y.extend(y_coords)
                
                if all_x and all_y:
                    x_margin = (max(all_x) - min(all_x)) * 0.1 or 1
                    y_margin = (max(all_y) - min(all_y)) * 0.1 or 1
                    self.map_canvas.axes.set_xlim(min(all_x) - x_margin, max(all_x) + x_margin)
                    self.map_canvas.axes.set_ylim(min(all_y) - y_margin, max(all_y) + y_margin)
                
                self.map_canvas.axes.set_xlabel('X Coordinate (Easting)')
                self.map_canvas.axes.set_ylabel('Y Coordinate (Northing)')
                self.map_canvas.axes.set_title(f'Map Preview - {len(self.map_canvas.polygon_data)} Polygons')
                self.map_canvas.axes.grid(True, alpha=0.3)
                self.map_canvas.axes.set_aspect('equal', adjustable='datalim')
                self.map_canvas.draw()
            
            self.status_label.setText("Polygons deleted from preview")
            
        except Exception as e:
            error_msg = f"Delete failed: {str(e)}"
            self.log(error_msg)
            QMessageBox.critical(self, "Error", error_msg)
        
    def merge_files(self):
        if not self.input_files or not self.output_file:
            return
        
        # Check if any polygons selected
        if not self.selected_polygon_indices:
            reply = QMessageBox.question(
                self,
                "No Selection",
                "No polygons selected. Merge ALL polygons?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.select_all_polygons()
            else:
                return
        
        dialog = AttributeDialog(self)
        dialog.name_input.setText(Path(self.output_file).stem)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.merge_attributes = dialog.get_attributes()
            
            self.log(f"Merging {len(self.selected_polygon_indices)} selected polygon(s)...")
            self.log(f"Name: {self.merge_attributes['Name']}")
            self.log(f"Group: {self.merge_attributes['Group']}")
            self.progress_bar.setVisible(True)
            self.status_label.setText("Merging...")
            
            self.btn_merge.setEnabled(False)
            self.btn_preview.setEnabled(False)
            
            self.worker = MergeWorker(self.input_files, self.output_file, 
                                     self.merge_attributes, self.selected_polygon_indices)
            self.worker.set_preview_mode(False)
            self.worker.progress.connect(self.update_progress)
            self.worker.finished.connect(self.merge_finished)
            self.worker.log_message.connect(self.log)  # Connect log signal
            self.worker.start()
    
    def export_to_shapefile(self):
        """Convert TAB file to Shapefile and create ZIP"""
        if not self.output_file or not os.path.exists(self.output_file):
            QMessageBox.warning(self, "Warning", "No TAB file to export. Merge files first.")
            return
        
        zip_path, _ = QFileDialog.getSaveFileName(
            self, "Save Shapefile ZIP", 
            str(Path(self.output_file).with_suffix('.zip')),
            "ZIP Files (*.zip)"
        )
        
        if not zip_path:
            return
        
        try:
            self.log("Converting TAB to Shapefile...")
            self.status_label.setText("Converting to Shapefile...")
            
            shp_base = Path(self.output_file).stem
            temp_dir = Path(self.output_file).parent / "temp_shp"
            temp_dir.mkdir(exist_ok=True)
            
            shp_path = temp_dir / f"{shp_base}.shp"
            
            src_ds = ogr.Open(self.output_file)
            if not src_ds:
                raise Exception("Could not open TAB file")
            
            src_layer = src_ds.GetLayer(0)
            
            driver = ogr.GetDriverByName('ESRI Shapefile')
            if os.path.exists(str(shp_path)):
                driver.DeleteDataSource(str(shp_path))
            
            dst_ds = driver.CreateDataSource(str(shp_path))
            dst_layer = dst_ds.CreateLayer(shp_base, src_layer.GetSpatialRef(), 
                                          src_layer.GetGeomType())
            
            src_layer_defn = src_layer.GetLayerDefn()
            for i in range(src_layer_defn.GetFieldCount()):
                field_defn = src_layer_defn.GetFieldDefn(i)
                dst_layer.CreateField(field_defn)
            
            src_layer.ResetReading()
            for feature in src_layer:
                dst_layer.CreateFeature(feature)
            
            src_ds = None
            dst_ds = None
            
            self.log("Creating ZIP archive...")
            
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for ext in ['.shp', '.shx', '.dbf', '.prj', '.cpg']:
                    file_path = temp_dir / f"{shp_base}{ext}"
                    if file_path.exists():
                        zipf.write(file_path, file_path.name)
                        self.log(f"Added {file_path.name} to ZIP")
            
            import shutil
            shutil.rmtree(temp_dir)
            
            self.log(f"Shapefile exported to: {zip_path}")
            self.status_label.setText("Export complete!")
            
            QMessageBox.information(self, "Success", 
                                   f"Shapefile exported and zipped:\n{zip_path}")
            
        except Exception as e:
            error_msg = f"Export failed: {str(e)}"
            self.log(error_msg)
            self.status_label.setText("Export failed")
            QMessageBox.critical(self, "Error", error_msg)
    
    def update_progress(self, value, message):
        self.progress_bar.setValue(value)
        self.status_label.setText(message)
        
    def merge_finished(self, success, message, file_stats):
        self.progress_bar.setVisible(False)
        
        if success:
            self.status_label.setText("Merge completed!")
            self.log("=" * 50)
            self.log("MERGE COMPLETED")
            self.log(message)
            for stat in file_stats:
                self.log(stat)
            self.log(f"Output: {self.output_file}")
            self.log("=" * 50)
            QMessageBox.information(self, "Success", 
                                   f"{message}\n\nAttribute Information:\nName: {self.merge_attributes['Name']}\nGroup: {self.merge_attributes['Group']}")
        else:
            self.status_label.setText("Merge failed")
            self.log("ERROR: " + message)
            QMessageBox.critical(self, "Error", message)
            
        self.update_button_states()
    
    
    # ====================================================================
    # Tools Tab Methods
    # ====================================================================
    
    def convert_file_to_shp(self):
        """Convert any geospatial file to zipped shapefile"""
        # Use existing ProcessingThread/GeoPandas pipeline instead of a separate undefined worker
        if not GEOPANDAS_AVAILABLE:
            QMessageBox.warning(self, "GeoPandas Required",
                                "GeoPandas is required for format conversion.\n\n"
                                "Install: conda install -c conda-forge geopandas")
            return

        # Select input file
        input_file, _ = QFileDialog.getOpenFileName(
            self, "Select Geospatial File", "",
            "All Supported (*.geojson *.json *.kml *.kmz *.gpkg *.shp *.gml);;GeoJSON (*.geojson *.json);;KML/KMZ (*.kml *.kmz);;GeoPackage (*.gpkg);;Shapefile (*.shp);;GML (*.gml);;All Files (*.*)"
        )

        if not input_file:
            return

        # Select output location
        default_name = os.path.splitext(os.path.basename(input_file))[0] + "_converted.zip"
        output_zip, _ = QFileDialog.getSaveFileName(
            self, "Save Shapefile ZIP", default_name,
            "ZIP files (*.zip)"
        )

        if not output_zip:
            return

        # Ensure .zip extension
        if not output_zip.lower().endswith('.zip'):
            output_zip += '.zip'

        # Start processing using the existing ProcessingThread
        params = {
            'input_file': input_file,
            'output_zip': output_zip
        }

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Converting to shapefile...")

        self.log(f"Converting: {os.path.basename(input_file)}")
        self.log(f"Output: {os.path.basename(output_zip)}")

        self.start_geop_processing('convert_to_shapefile_zip', params)
    
    def check_geom_type(self):
        """Check geometry type of file for Huawei Discovery compatibility"""
        if not GEOPANDAS_AVAILABLE:
            QMessageBox.warning(self, "GeoPandas Required",
                              "GeoPandas is required for geometry checking.\n\n"
                              "Install: conda install -c conda-forge geopandas")
            return
        
        # Select file to check
        input_file, _ = QFileDialog.getOpenFileName(
            self, "Select File to Check", "",
            "All Supported (*.geojson *.json *.kml *.kmz *.gpkg *.shp *.gml *.zip);;All Files (*.*)"
        )
        
        if not input_file:
            return
        
        try:
            self.log("Checking geometry types...")
            self.status_label.setText("Analyzing file...")
            
            # Read the file
            gdf = gpd.read_file(input_file)
            
            # Get geometry types
            geom_types = gdf.geometry.geom_type.value_counts()
            total_features = len(gdf)
            
            # Check for compatibility
            compatible_types = {'Polygon', 'MultiPolygon'}
            incompatible_types = set(geom_types.index) - compatible_types
            
            # Build report
            report = []
            report.append("=" * 60)
            report.append("GEOMETRY TYPE CHECK REPORT")
            report.append("=" * 60)
            report.append(f"File: {os.path.basename(input_file)}")
            report.append(f"Total Features: {total_features}")
            report.append("")
            report.append("Geometry Types Found:")
            for geom_type, count in geom_types.items():
                percentage = (count / total_features) * 100
                report.append(f"  • {geom_type}: {count} ({percentage:.1f}%)")
            
            report.append("")
            
            if incompatible_types:
                report.append("⚠ COMPATIBILITY ISSUES DETECTED:")
                report.append("")
                report.append("Huawei Discovery requires POLYGON geometries only.")
                report.append("The following incompatible types were found:")
                for geom_type in incompatible_types:
                    count = geom_types[geom_type]
                    report.append(f"  ✗ {geom_type}: {count} features")
                report.append("")
                report.append("RECOMMENDATION:")
                report.append("  1. Filter to keep only Polygon/MultiPolygon features")
                report.append("  2. Convert other geometries to polygons if applicable")
                report.append("  3. Remove incompatible features before upload")
                status = "❌ NOT COMPATIBLE"
                is_compatible = False
            else:
                report.append("✓ COMPATIBILITY CHECK PASSED")
                report.append("")
                report.append("All geometries are Polygon or MultiPolygon types.")
                report.append("This file is ready for Huawei Discovery upload!")
                status = "✓ COMPATIBLE"
                is_compatible = True
            
            report.append("=" * 60)
            
            # Display in log
            for line in report:
                self.log(line)
            
            # Show message box
            if is_compatible:
                QMessageBox.information(self, "Geometry Check - Compatible", 
                                      f"{status}\n\n"
                                      f"Total Features: {total_features}\n"
                                      f"All geometries are Polygon/MultiPolygon.\n\n"
                                      f"✓ Ready for Huawei Discovery!")
            else:
                msg = f"{status}\n\n"
                msg += f"Total Features: {total_features}\n\n"
                msg += "Incompatible geometry types found:\n"
                for geom_type in incompatible_types:
                    count = geom_types[geom_type]
                    msg += f"  • {geom_type}: {count} features\n"
                msg += "\nPlease clean the data before uploading to Huawei Discovery."
                QMessageBox.warning(self, "Geometry Check - Issues Found", msg)
            
            self.status_label.setText("Geometry check completed")
            
        except Exception as e:
            error_msg = f"Error checking geometry types: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            self.log(traceback.format_exc())
            QMessageBox.critical(self, "Error", error_msg)
            self.status_label.setText("Geometry check failed")
    
    
    def run_simplify_operation(self):
        """Run simplify operation to reduce vertex count"""
        if not hasattr(self, 'simplify_file_path') or not self.simplify_file_path:
            QMessageBox.warning(self, "Warning", "Please load a file first")
            return
        
        if not GEOPANDAS_AVAILABLE:
            QMessageBox.warning(self, "GeoPandas Required",
                              "GeoPandas is required for simplification.\n\n"
                              "Install: conda install -c conda-forge geopandas")
            return
        
        # Get output file
        default_name = os.path.splitext(os.path.basename(self.simplify_file_path))[0] + "_simplified" + os.path.splitext(self.simplify_file_path)[1]
        
        output_file, _ = QFileDialog.getSaveFileName(
            self, "Save Simplified File", default_name,
            "TAB File (*.tab);;Shapefile (*.shp);;GeoJSON (*.geojson);;GeoPackage (*.gpkg);;All Files (*)"
        )
        
        if not output_file:
            return
        
        try:
            from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, Point, MultiPoint
            
            self.log("Starting simplification...")
            self.status_label.setText("Simplifying geometries...")
            
            # Read file
            gdf = gpd.read_file(self.simplify_file_path)
            tolerance = self.simplify_tolerance.value()
            
            # Function to count points in any geometry type
            def count_points(geom):
                if geom is None or geom.is_empty:
                    return 0
                if isinstance(geom, Point):
                    return 1
                elif isinstance(geom, (LineString, Polygon)):
                    try:
                        if isinstance(geom, Polygon):
                            # Count exterior + all interior rings
                            count = len(geom.exterior.coords)
                            for interior in geom.interiors:
                                count += len(interior.coords)
                            return count
                        else:
                            return len(geom.coords)
                    except:
                        return 0
                elif isinstance(geom, (MultiPolygon, MultiLineString, MultiPoint)):
                    return sum(count_points(g) for g in geom.geoms)
                return 0
            
            # Count original points
            original_points = sum(count_points(geom) for geom in gdf.geometry)
            
            self.log(f"Original geometry has {original_points} points")
            self.log(f"Applying tolerance: {tolerance}")
            
            # Simplify geometries with error handling
            def safe_simplify(geom):
                try:
                    if geom is None or geom.is_empty:
                        return geom
                    
                    # Make sure geometry is valid first
                    if not geom.is_valid:
                        geom = geom.buffer(0)  # Fix invalid geometries
                    
                    # Simplify
                    simplified = geom.simplify(tolerance, preserve_topology=True)
                    
                    # Make sure result is valid
                    if not simplified.is_valid:
                        simplified = simplified.buffer(0)
                    
                    return simplified
                except Exception as e:
                    self.log(f"Warning: Could not simplify one geometry: {str(e)}")
                    return geom  # Return original if simplification fails
            
            gdf['geometry'] = gdf['geometry'].apply(safe_simplify)
            
            # Count simplified points
            simplified_points = sum(count_points(geom) for geom in gdf.geometry)
            
            # Save simplified file
            gdf.to_file(output_file)
            
            reduction_pct = ((original_points - simplified_points) / original_points * 100) if original_points > 0 else 0
            
            self.log(f"✓ Simplification completed!")
            self.log(f"Simplified geometry has {simplified_points} points")
            self.log(f"Reduction: {reduction_pct:.1f}%")
            self.log(f"Output: {output_file}")
            
            # Check if still too many points for Discovery
            max_points_per_feature = max(count_points(geom) for geom in gdf.geometry) if len(gdf) > 0 else 0
            warning_msg = ""
            if max_points_per_feature > 1000:
                warning_msg = f"\n\n⚠ Warning: Largest feature still has {max_points_per_feature} points.\nHuawei Discovery limit is ~500-1000 points.\nTry higher tolerance (e.g., {tolerance * 5:.4f})"
            
            QMessageBox.information(self, "Simplification Complete",
                                  f"Geometry simplified successfully!\n\n"
                                  f"Original points: {original_points}\n"
                                  f"Simplified points: {simplified_points}\n"
                                  f"Reduction: {reduction_pct:.1f}%\n"
                                  f"Max points in single feature: {max_points_per_feature}\n\n"
                                  f"Saved to: {os.path.basename(output_file)}{warning_msg}")
            
            self.status_label.setText("Simplification completed")
            
        except Exception as e:
            error_msg = f"Error simplifying geometry: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            self.log(traceback.format_exc())
            QMessageBox.critical(self, "Error", error_msg)
            self.status_label.setText("Simplification failed")
    
    def remove_polygon_holes(self):
        """Remove holes/inner rings from polygons"""
        if not GEOPANDAS_AVAILABLE:
            QMessageBox.warning(self, "GeoPandas Required",
                              "GeoPandas is required for polygon cleanup.\n\n"
                              "Install: conda install -c conda-forge geopandas")
            return
        
        # Select input file
        input_file, _ = QFileDialog.getOpenFileName(
            self, "Select File to Clean", "",
            "All Supported (*.geojson *.json *.kml *.kmz *.gpkg *.shp *.tab *.gml);;All Files (*.*)"
        )
        
        if not input_file:
            return
        
        # Select output location
        default_name = os.path.splitext(os.path.basename(input_file))[0] + "_cleaned" + os.path.splitext(input_file)[1]
        output_file, _ = QFileDialog.getSaveFileName(
            self, "Save Cleaned File", default_name,
            "TAB File (*.tab);;Shapefile (*.shp);;GeoJSON (*.geojson);;GeoPackage (*.gpkg);;All Files (*.*)"
        )
        
        if not output_file:
            return
        
        try:
            self.log("Removing polygon holes...")
            self.status_label.setText("Cleaning polygons...")
            
            # Read the file
            gdf = gpd.read_file(input_file)
            
            from shapely.geometry import Polygon, MultiPolygon
            
            holes_removed = 0
            
            # Function to remove holes from a single polygon
            def remove_holes(geom):
                nonlocal holes_removed
                if isinstance(geom, Polygon):
                    if len(geom.interiors) > 0:
                        holes_removed += len(geom.interiors)
                        # Return polygon with only exterior ring
                        return Polygon(geom.exterior)
                    return geom
                elif isinstance(geom, MultiPolygon):
                    # Process each polygon in the MultiPolygon
                    cleaned_polys = []
                    for poly in geom.geoms:
                        if len(poly.interiors) > 0:
                            holes_removed += len(poly.interiors)
                            cleaned_polys.append(Polygon(poly.exterior))
                        else:
                            cleaned_polys.append(poly)
                    return MultiPolygon(cleaned_polys)
                return geom
            
            # Apply hole removal to all geometries
            gdf['geometry'] = gdf['geometry'].apply(remove_holes)
            
            # Save cleaned file
            gdf.to_file(output_file)
            
            self.log(f"✓ Polygon cleanup completed!")
            self.log(f"Holes removed: {holes_removed}")
            self.log(f"Total features: {len(gdf)}")
            self.log(f"Output: {output_file}")
            
            QMessageBox.information(self, "Cleanup Complete",
                                  f"Polygon cleanup completed!\n\n"
                                  f"Holes removed: {holes_removed}\n"
                                  f"Total features: {len(gdf)}\n\n"
                                  f"Saved to: {os.path.basename(output_file)}")
            
            self.status_label.setText("Polygon cleanup completed")
            
        except Exception as e:
            error_msg = f"Error cleaning polygons: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            self.log(traceback.format_exc())
            QMessageBox.critical(self, "Error", error_msg)
            self.status_label.setText("Polygon cleanup failed")
    
    # ====================================================================
    # Edit Tab Methods (Cut/Delete Geometry)
    # ====================================================================
    
    def load_file_for_editing(self):
        """Load a file for geometry editing with map preview"""
        if not GEOPANDAS_AVAILABLE:
            QMessageBox.warning(self, "GeoPandas Required",
                              "GeoPandas is required for editing.\n\n"
                              "Install: conda install -c conda-forge geopandas")
            return
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select File to Edit", "",
            "All Supported (*.tab *.shp *.geojson *.gpkg *.kml);;TAB Files (*.tab);;Shapefile (*.shp);;GeoJSON (*.geojson);;KML (*.kml);;GeoPackage (*.gpkg);;All Files (*.*)"
        )
        
        if not file_path:
            return
        
        try:
            self.edit_gdf = gpd.read_file(file_path)
            self.edit_file_path = file_path
            self.edit_file_label.setText(os.path.basename(file_path))
            
            # Populate feature list
            self.edit_feature_list.clear()
            for idx, row in self.edit_gdf.iterrows():
                name = row.get('name', row.get('Name', f'Feature {idx}'))
                item = QListWidgetItem(f"{idx}: {name}")
                item.setData(Qt.ItemDataRole.UserRole, idx)
                self.edit_feature_list.addItem(item)
            
            # Draw all features on map
            self.draw_edit_map()
            
            self.log(f"Loaded {len(self.edit_gdf)} features from {os.path.basename(file_path)}")
            self.status_label.setText(f"Loaded {len(self.edit_gdf)} features")
            
        except Exception as e:
            error_msg = f"Error loading file: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            QMessageBox.critical(self, "Error", error_msg)
    
    def draw_edit_map(self):
        """Draw all geometries on the edit map"""
        if not hasattr(self, 'edit_map_canvas') or not hasattr(self, 'edit_gdf'):
            return
        
        try:
            from shapely.geometry import Polygon, MultiPolygon, LineString, MultiLineString, Point
            
            self.edit_map_canvas.axes.clear()
            
            # Plot all geometries
            for idx, row in self.edit_gdf.iterrows():
                geom = row.geometry
                
                # Determine color based on selection
                if idx == self.edit_selected_feature_idx:
                    color = '#ff6b6b'  # Red for selected
                    alpha = 0.7
                    linewidth = 2
                else:
                    color = '#4ecdc4'  # Cyan for unselected
                    alpha = 0.5
                    linewidth = 1
                
                # Plot different geometry types
                if isinstance(geom, (Polygon, MultiPolygon)):
                    if isinstance(geom, Polygon):
                        polys = [geom]
                    else:
                        polys = list(geom.geoms)
                    
                    for poly in polys:
                        x, y = poly.exterior.xy
                        self.edit_map_canvas.axes.plot(x, y, color='black', linewidth=linewidth)
                        self.edit_map_canvas.axes.fill(x, y, color=color, alpha=alpha)
                
                elif isinstance(geom, (LineString, MultiLineString)):
                    if isinstance(geom, LineString):
                        lines = [geom]
                    else:
                        lines = list(geom.geoms)
                    
                    for line in lines:
                        x, y = line.xy
                        self.edit_map_canvas.axes.plot(x, y, color=color, linewidth=linewidth+1, alpha=alpha)
                
                elif isinstance(geom, Point):
                    self.edit_map_canvas.axes.plot(geom.x, geom.y, 'o', color=color, markersize=8, alpha=alpha)
            
            # Draw polygon being drawn
            if self.edit_drawn_points:
                points = self.edit_drawn_points + [self.edit_drawn_points[0]]  # Close the polygon
                x_coords = [p[0] for p in points]
                y_coords = [p[1] for p in points]
                self.edit_map_canvas.axes.plot(x_coords, y_coords, 'r--', linewidth=2, label='Drawing')
                self.edit_map_canvas.axes.fill(x_coords, y_coords, color='red', alpha=0.3)
                
                # Draw points
                for point in self.edit_drawn_points:
                    self.edit_map_canvas.axes.plot(point[0], point[1], 'ro', markersize=8)
            
            self.edit_map_canvas.axes.set_xlabel('Longitude')
            self.edit_map_canvas.axes.set_ylabel('Latitude')
            self.edit_map_canvas.axes.set_title('Edit Map - Click to draw cutting polygon')
            self.edit_map_canvas.axes.grid(True, alpha=0.3)
            self.edit_map_canvas.axes.legend()
            self.edit_map_canvas.draw()
            
        except Exception as e:
            self.log(f"ERROR drawing map: {str(e)}")
    
    def preview_edit_feature(self, item):
        """Preview selected feature on map"""
        if not hasattr(self, 'edit_gdf'):
            return
        
        try:
            idx = item.data(Qt.ItemDataRole.UserRole)
            self.edit_selected_feature_idx = idx
            
            # Redraw map to highlight selected feature
            self.draw_edit_map()
            
            self.log(f"Feature {idx} selected")
            self.status_label.setText(f"Selected feature {idx}")
            
        except Exception as e:
            self.log(f"ERROR previewing feature: {str(e)}")
    
    def start_polygon_drawing(self):
        """Start interactive polygon drawing mode"""
        self.edit_drawing_active = True
        self.edit_drawn_points = []
        self.draw_points_label.setText("Points: 0 - Click on map to draw")
        self.log("Drawing mode activated - click on map to add points")
        self.status_label.setText("Drawing mode: Click on map to add points")
    
    def clear_polygon_drawing(self):
        """Clear the drawn polygon"""
        self.edit_drawn_points = []
        self.edit_drawing_active = False
        self.draw_points_label.setText("Points: 0")
        if hasattr(self, 'edit_gdf'):
            self.draw_edit_map()
        self.log("Drawing cleared")
        self.status_label.setText("Drawing cleared")
    
    def on_map_click(self, event):
        """Handle mouse clicks on the map for drawing"""
        if not self.edit_drawing_active or not event.inaxes:
            return
        
        # Add point to drawn polygon
        self.edit_drawn_points.append((event.xdata, event.ydata))
        self.draw_points_label.setText(f"Points: {len(self.edit_drawn_points)}")
        
        # Redraw map with new point
        self.draw_edit_map()
        
        self.log(f"Added point: ({event.xdata:.6f}, {event.ydata:.6f})")
    
    def cut_geometry_polygon(self):
        """Cut using the drawn polygon"""
        if not hasattr(self, 'edit_gdf'):
            QMessageBox.warning(self, "Warning", "Please load a file first")
            return
        
        if len(self.edit_drawn_points) < 3:
            QMessageBox.warning(self, "Warning", "Please draw a polygon with at least 3 points\n\n1. Click 'Start Drawing Polygon'\n2. Click on map 3+ times\n3. Then click 'Cut Drawn Polygon'")
            return
        
        # Get selected feature or auto-select first one
        selected_items = self.edit_feature_list.selectedItems()
        
        if not selected_items:
            # Auto-select first feature if none selected
            if self.edit_feature_list.count() > 0:
                self.edit_feature_list.setCurrentRow(0)
                selected_items = self.edit_feature_list.selectedItems()
                self.log("Auto-selected first feature")
            else:
                QMessageBox.warning(self, "Warning", "No features available to cut")
                return
        
        try:
            from shapely.geometry import Polygon
            
            # Get selected feature index
            idx = selected_items[0].data(Qt.ItemDataRole.UserRole)
            
            # Store point count before clearing
            point_count = len(self.edit_drawn_points)
            
            # Create cutting polygon from drawn points
            cut_polygon = Polygon(self.edit_drawn_points)
            
            # Get original geometry
            original_geom = self.edit_gdf.iloc[idx].geometry
            
            # Cut: remove the polygon area from the geometry (difference operation)
            result_geom = original_geom.difference(cut_polygon)
            
            # Update the geometry
            self.edit_gdf.at[idx, 'geometry'] = result_geom
            
            # Clear drawing and redraw map
            self.edit_drawn_points = []
            self.edit_drawing_active = False
            self.draw_points_label.setText("Points: 0")
            self.draw_edit_map()
            
            self.log(f"✓ Cut polygon operation completed on feature {idx}")
            self.log(f"Removed polygon with {point_count} vertices")
            
            QMessageBox.information(self, "Cut Complete",
                                  f"Successfully cut drawn polygon from feature {idx}\n\n"
                                  f"Polygon had {point_count} points\n"
                                  f"Remember to save the edited file!")
            
            self.status_label.setText(f"Feature {idx} edited - not saved yet")
            
        except Exception as e:
            error_msg = f"Error cutting with polygon: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            self.log(traceback.format_exc())
            QMessageBox.critical(self, "Error", error_msg)
    
    def save_edited_file_multi_format(self):
        """Save the edited file in multiple format options"""
        if not hasattr(self, 'edit_gdf'):
            QMessageBox.warning(self, "Warning", "No file loaded to save")
            return
        
        # Get output file path with format options
        default_name = os.path.splitext(os.path.basename(self.edit_file_path))[0] + "_edited"
        
        output_file, _ = QFileDialog.getSaveFileName(
            self, "Save Edited File (Multi Format)", default_name,
            "MapInfo TAB (*.tab);;ESRI Shapefile (*.shp);;GeoJSON (*.geojson);;GeoPackage (*.gpkg);;GML (*.gml);;All Files (*.*)"
        )
        
        if not output_file:
            return
        
        try:
            # Determine driver based on file extension
            ext = os.path.splitext(output_file)[1].lower()
            
            if ext == '.tab':
                driver = 'MapInfo File'
            elif ext == '.shp':
                driver = 'ESRI Shapefile'
            elif ext == '.geojson':
                driver = 'GeoJSON'
            elif ext == '.gpkg':
                driver = 'GPKG'
            elif ext == '.gml':
                driver = 'GML'
            else:
                driver = 'GeoJSON'  # Default to GeoJSON if unknown extension
            
            # Save the edited GeoDataFrame
            self.edit_gdf.to_file(output_file, driver=driver)
            
            self.log(f"✓ Edited file saved: {output_file}")
            self.log(f"  Format: {driver}")
            self.log(f"  Features: {len(self.edit_gdf)}")
            
            QMessageBox.information(self, "Save Complete",
                                  f"Edited file saved successfully!\n\n"
                                  f"File: {os.path.basename(output_file)}\n"
                                  f"Format: {driver}\n"
                                  f"Features: {len(self.edit_gdf)}")
            
            self.status_label.setText("Edited file saved (Multi Format)")
            
        except Exception as e:
            error_msg = f"Error saving file: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            self.log(traceback.format_exc())
            QMessageBox.critical(self, "Error", error_msg)
    
    def save_edited_file_shapefile_zip(self):
        """Save the edited file as Shapefile and automatically ZIP it"""
        if not hasattr(self, 'edit_gdf'):
            QMessageBox.warning(self, "Warning", "No file loaded to save")
            return
        
        # Get output ZIP file path
        default_name = os.path.splitext(os.path.basename(self.edit_file_path))[0] + "_edited.zip"
        
        output_zip, _ = QFileDialog.getSaveFileName(
            self, "Save as Shapefile ZIP", default_name,
            "ZIP Archive (*.zip);;All Files (*.*)"
        )
        
        if not output_zip:
            return
        
        try:
            self.status_label.setText("Creating Shapefile ZIP...")
            
            # Create temporary directory for shapefile components
            import tempfile
            import zipfile
            temp_dir = tempfile.mkdtemp()
            
            # Base name for shapefile (without extension)
            base_name = os.path.splitext(os.path.basename(output_zip))[0]
            shp_path = os.path.join(temp_dir, f"{base_name}.shp")
            
            # Save as shapefile
            self.log(f"Creating temporary shapefile: {shp_path}")
            self.edit_gdf.to_file(shp_path, driver='ESRI Shapefile')
            
            # Create ZIP file containing all shapefile components
            self.log(f"Zipping shapefile components...")
            with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file in os.listdir(temp_dir):
                    file_path = os.path.join(temp_dir, file)
                    zipf.write(file_path, arcname=file)
            
            # Clean up temp directory
            shutil.rmtree(temp_dir)
            
            self.log(f"✓ Shapefile ZIP created: {output_zip}")
            self.log(f"  Format: ESRI Shapefile (zipped)")
            self.log(f"  Features: {len(self.edit_gdf)}")
            self.log(f"  Ready for Huawei Discovery upload!")
            
            QMessageBox.information(self, "ZIP Creation Complete",
                                  f"Shapefile ZIP created successfully!\n\n"
                                  f"File: {os.path.basename(output_zip)}\n"
                                  f"Format: ESRI Shapefile (zipped)\n"
                                  f"Features: {len(self.edit_gdf)}\n\n"
                                  f"Ready for Huawei Discovery upload!")
            
            self.status_label.setText("Edited file saved as Shapefile ZIP")
            
        except Exception as e:
            error_msg = f"Error creating Shapefile ZIP: {str(e)}"
            self.log(f"ERROR: {error_msg}")
            self.log(traceback.format_exc())
            QMessageBox.critical(self, "Error", error_msg)
    
    def save_edited_file(self):
        """Legacy save function - calls multi-format save"""
        self.save_edited_file_multi_format()



def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    if not GDAL_AVAILABLE:
        QMessageBox.critical(None, "GDAL Not Available",
            "GDAL/OGR not installed.\n\nInstall: conda install -c conda-forge gdal")
        sys.exit(1)
    
    window = UnifiedGeospatialTool()
    
    # Set custom icon
    icon_path = os.path.join(os.path.dirname(__file__), 'app_icon.ico')
    if os.path.exists(icon_path):
        window.setWindowIcon(QIcon(icon_path))
    
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()