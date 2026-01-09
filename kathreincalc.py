import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import numpy as np
import json
import math


# ==============================================
# DATA MODELS
# ==============================================

class Antenna:
    """Antenna data model"""

    def __init__(self, name, manufacturer, model, frequency_range, gain,
                 horizontal_bw, vertical_bw, front_to_back, tilt_range,
                 electrical_tilt_steps, patterns, connector_type="N-Type",
                 polarization="Dual"):
        self.name = name
        self.manufacturer = manufacturer
        self.model = model
        self.frequency_range = frequency_range  # MHz
        self.gain = gain  # dBi
        self.horizontal_bw = horizontal_bw  # degrees
        self.vertical_bw = vertical_bw  # degrees
        self.front_to_back = front_to_back  # dB
        self.tilt_range = tilt_range  # degrees
        self.electrical_tilt_steps = electrical_tilt_steps  # degrees per step
        self.patterns = patterns  # horizontal and vertical patterns
        self.connector_type = connector_type
        self.polarization = polarization


class SiteConfiguration:
    """Site configuration model"""

    def __init__(self, site_id="SITE001", latitude=0.0, longitude=0.0,
                 antenna_height=30.0, azimuth=0.0, mechanical_tilt=0.0,
                 electrical_tilt=0.0, antenna_type="Kathrein 80010638",
                 frequency=900.0, power=43.0):
        self.site_id = site_id
        self.latitude = latitude
        self.longitude = longitude
        self.antenna_height = antenna_height  # meters
        self.azimuth = azimuth  # degrees
        self.mechanical_tilt = mechanical_tilt  # degrees
        self.electrical_tilt = electrical_tilt  # degrees
        self.antenna_type = antenna_type
        self.frequency = frequency  # MHz
        self.power = power  # dBm


# ==============================================
# RF CALCULATIONS ENGINE
# ==============================================

class RFCalculator:
    """Core RF calculations engine"""

    @staticmethod
    def calculate_total_tilt(mechanical: float, electrical: float) -> float:
        """Calculate total tilt from mechanical and electrical components"""
        return mechanical + electrical

    @staticmethod
    def dbm_to_watts(dbm: float) -> float:
        """Convert dBm to Watts"""
        return 10 ** ((dbm - 30) / 10)

    @staticmethod
    def watts_to_dbm(watts: float) -> float:
        """Convert Watts to dBm"""
        return 10 * math.log10(watts) + 30

    @staticmethod
    def calculate_eirp(power_dbm: float, gain_dbi: float, losses_db: float = 2) -> float:
        """Calculate Equivalent Isotropically Radiated Power"""
        return power_dbm + gain_dbi - losses_db

    @staticmethod
    def free_space_path_loss(distance_km: float, frequency_mhz: float) -> float:
        """Calculate free space path loss in dB"""
        return 20 * math.log10(distance_km) + 20 * math.log10(frequency_mhz) + 32.44

    @staticmethod
    def cost231_hata_path_loss(distance_km: float, frequency_mhz: float,
                               tx_height_m: float, rx_height_m: float,
                               environment: str = "urban") -> float:
        """
        COST-231 Hata model for path loss prediction
        Environments: urban, suburban, rural
        """
        if frequency_mhz < 150 or frequency_mhz > 2000:
            raise ValueError("Frequency must be between 150 and 2000 MHz")

        a_hr = (1.1 * math.log10(frequency_mhz) - 0.7) * rx_height_m - (1.56 * math.log10(frequency_mhz) - 0.8)

        loss = 69.55 + 26.16 * math.log10(frequency_mhz) - 13.82 * math.log10(tx_height_m) - a_hr
        loss += (44.9 - 6.55 * math.log10(tx_height_m)) * math.log10(distance_km)

        if environment == "suburban":
            loss -= 2 * (math.log10(frequency_mhz / 28)) ** 2 - 5.4
        elif environment == "rural":
            loss -= 4.78 * (math.log10(frequency_mhz)) ** 2 + 18.33 * math.log10(frequency_mhz) - 40.94

        return loss

    @staticmethod
    def calculate_coverage_radius(eirp_dbm: float, rx_sensitivity_dbm: float,
                                  frequency_mhz: float, environment: str = "urban") -> float:
        """Calculate approximate coverage radius"""
        max_path_loss = eirp_dbm - rx_sensitivity_dbm

        if environment == "urban":
            factor = 0.03
        elif environment == "suburban":
            factor = 0.02
        else:  # rural
            factor = 0.01

        return max_path_loss * factor  # km


# ==============================================
# ANTENNA DATABASE
# ==============================================

class AntennaDatabase:
    """Antenna database manager"""

    def __init__(self):
        self.antennas = self._load_default_antennas()
        self.categories = {
            "Kathrein": [],
            "Huawei": [],
            "Ericsson": [],
            "Nokia": [],
            "CommScope": []
        }
        self._categorize_antennas()

    def _load_default_antennas(self):
        """Load default antenna patterns"""
        antennas = []

        # Kathrein antenna
        patterns_h = list(np.concatenate([
            np.linspace(-40, 0, 60),
            np.linspace(0, -40, 61)[1:]
        ]))
        patterns_v = list(np.concatenate([
            np.linspace(-20, 0, 30),
            np.linspace(0, -20, 31)[1:]
        ]))

        antennas.append(Antenna(
            name="Kathrein 80010638",
            manufacturer="Kathrein",
            model="80010638",
            frequency_range=(790, 960),
            gain=18.5,
            horizontal_bw=65,
            vertical_bw=6.5,
            front_to_back=30,
            tilt_range=(-10, 10),
            electrical_tilt_steps=1,
            patterns={
                "horizontal": patterns_h,
                "vertical": patterns_v
            }
        ))

        # Huawei antenna
        antennas.append(Antenna(
            name="Huawei APM3030",
            manufacturer="Huawei",
            model="APM3030",
            frequency_range=(1710, 2690),
            gain=18.0,
            horizontal_bw=65,
            vertical_bw=6.0,
            front_to_back=28,
            tilt_range=(-15, 15),
            electrical_tilt_steps=0.5,
            patterns={
                "horizontal": list(np.concatenate([
                    np.linspace(-35, 0, 70),
                    np.linspace(0, -35, 71)[1:]
                ])),
                "vertical": list(np.concatenate([
                    np.linspace(-25, 0, 50),
                    np.linspace(0, -25, 51)[1:]
                ]))
            }
        ))

        # Ericsson antenna
        antennas.append(Antenna(
            name="Ericsson AIR 6488",
            manufacturer="Ericsson",
            model="AIR 6488",
            frequency_range=(694, 960),
            gain=17.8,
            horizontal_bw=65,
            vertical_bw=7.0,
            front_to_back=25,
            tilt_range=(-12, 12),
            electrical_tilt_steps=1.0,
            patterns={
                "horizontal": list(np.concatenate([
                    np.linspace(-38, 0, 76),
                    np.linspace(0, -38, 77)[1:]
                ])),
                "vertical": list(np.concatenate([
                    np.linspace(-22, 0, 44),
                    np.linspace(0, -22, 45)[1:]
                ]))
            }
        ))

        return antennas

    def _categorize_antennas(self):
        """Categorize antennas by manufacturer"""
        for antenna in self.antennas:
            if antenna.manufacturer in self.categories:
                self.categories[antenna.manufacturer].append(antenna.name)

    def get_antenna_by_name(self, name: str):
        """Get antenna by name"""
        for antenna in self.antennas:
            if antenna.name == name:
                return antenna
        return None

    def get_antennas_by_manufacturer(self, manufacturer: str):
        """Get antenna names by manufacturer"""
        return self.categories.get(manufacturer, [])

    def search_antennas(self, frequency: float = None,
                        min_gain: float = None,
                        max_hbw: float = None):
        """Search antennas based on criteria"""
        results = []
        for antenna in self.antennas:
            match = True

            if frequency:
                if not (antenna.frequency_range[0] <= frequency <= antenna.frequency_range[1]):
                    match = False

            if min_gain:
                if antenna.gain < min_gain:
                    match = False

            if max_hbw:
                if antenna.horizontal_bw > max_hbw:
                    match = False

            if match:
                results.append(antenna)

        return results


# ==============================================
# MAIN APPLICATION
# ==============================================

class KathreinCalculatorApp:
    """Main application window"""

    def __init__(self, root):
        self.root = root
        self.root.title("RF Antenna Calculator - Fadzli Edition")
        self.root.geometry("1400x800")

        # Initialize components
        self.rf_calc = RFCalculator()
        self.antenna_db = AntennaDatabase()
        self.current_site = SiteConfiguration()

        # Style configuration
        self.bg_color = "#f0f0f0"
        self.root.configure(bg=self.bg_color)

        # Initialize widgets dictionary
        self.widgets = {}

        # Build UI
        self.setup_ui()

        # Initialize plots
        self.update_plots()

    def setup_ui(self):
        """Setup the main user interface"""
        # Footer element (pack first so it stays at the very bottom)
        footer_frame = tk.Frame(self.root, bg=self.bg_color)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        footer_label = tk.Label(footer_frame, text="Written in Python | Fadzli Abdullah | Huawei Technologies", font=('Ubuntu', 8), bg=self.bg_color)
        footer_label.pack(pady=2)

        # Create main container with paned window
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left panel - Controls
        left_frame = ttk.Frame(main_paned, padding=10)
        main_paned.add(left_frame, weight=1)

        # Right panel - Visualizations
        right_frame = ttk.Frame(main_paned, padding=10)
        main_paned.add(right_frame, weight=2)

        # Build left panel
        self.build_control_panel(left_frame)

        # Build right panel
        self.build_visualization_panel(right_frame)

        # Status bar (pack after footer so it appears above)
        self.setup_status_bar()

    def build_control_panel(self, parent):
        """Build the control panel with input fields"""
        # Notebook for tabs
        control_notebook = ttk.Notebook(parent)
        control_notebook.pack(fill=tk.BOTH, expand=True)

        # Site Configuration Tab
        site_frame = ttk.Frame(control_notebook, padding=10)
        control_notebook.add(site_frame, text="Site Config")
        self.build_site_config_tab(site_frame)

        # Antenna Selection Tab
        antenna_frame = ttk.Frame(control_notebook, padding=10)
        control_notebook.add(antenna_frame, text="Antenna")
        self.build_antenna_tab(antenna_frame)

        # Calculations Tab
        calc_frame = ttk.Frame(control_notebook, padding=10)
        control_notebook.add(calc_frame, text="Calculations")
        self.build_calculations_tab(calc_frame)

        # Tools Tab
        tools_frame = ttk.Frame(control_notebook, padding=10)
        control_notebook.add(tools_frame, text="Tools")
        self.build_tools_tab(tools_frame)

    def build_site_config_tab(self, parent):
        """Build site configuration tab"""
        row = 0

        # Site ID
        ttk.Label(parent, text="Site ID:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        self.site_id_var = tk.StringVar(value=self.current_site.site_id)
        ttk.Entry(parent, textvariable=self.site_id_var, width=20).grid(row=row, column=1, pady=5, padx=5, columnspan=2)
        row += 1

        # Coordinates
        ttk.Label(parent, text="Latitude:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.lat_var = tk.DoubleVar(value=self.current_site.latitude)
        ttk.Entry(parent, textvariable=self.lat_var, width=15).grid(row=row, column=1, pady=5, padx=5)
        row += 1

        ttk.Label(parent, text="Longitude:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.lon_var = tk.DoubleVar(value=self.current_site.longitude)
        ttk.Entry(parent, textvariable=self.lon_var, width=15).grid(row=row, column=1, pady=5, padx=5)
        row += 1

        # Antenna Height
        ttk.Label(parent, text="Antenna Height (m):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.height_var = tk.DoubleVar(value=self.current_site.antenna_height)
        height_spin = ttk.Spinbox(parent, from_=5, to=300, textvariable=self.height_var,
                                  width=15, command=self.update_calculations)
        height_spin.grid(row=row, column=1, pady=5, padx=5)
        row += 1

        # Azimuth
        ttk.Label(parent, text="Azimuth (°):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.azimuth_var = tk.DoubleVar(value=self.current_site.azimuth)
        azimuth_scale = ttk.Scale(parent, from_=0, to=360, variable=self.azimuth_var,
                                  command=lambda v: self.update_azimuth_display())
        azimuth_scale.grid(row=row, column=1, pady=5, padx=5, sticky=tk.EW)

        self.azimuth_display = ttk.Label(parent, text=f"{self.current_site.azimuth:.1f}°")
        self.azimuth_display.grid(row=row, column=2, pady=5, padx=5)
        row += 1

        # Mechanical Tilt
        ttk.Label(parent, text="Mechanical Tilt (°):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.mech_tilt_var = tk.DoubleVar(value=self.current_site.mechanical_tilt)
        mech_tilt_scale = ttk.Scale(parent, from_=-20, to=20, variable=self.mech_tilt_var,
                                    command=lambda v: self.update_tilt_display())
        mech_tilt_scale.grid(row=row, column=1, pady=5, padx=5, sticky=tk.EW)
        row += 1

        # Electrical Tilt
        ttk.Label(parent, text="Electrical Tilt (°):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.elec_tilt_var = tk.DoubleVar(value=self.current_site.electrical_tilt)
        elec_tilt_scale = ttk.Scale(parent, from_=-10, to=10, variable=self.elec_tilt_var,
                                    command=lambda v: self.update_tilt_display())
        elec_tilt_scale.grid(row=row, column=1, pady=5, padx=5, sticky=tk.EW)

        self.tilt_display = ttk.Label(parent, text="Total: 0.0°")
        self.tilt_display.grid(row=row, column=2, pady=5, padx=5)
        row += 1

        # Frequency
        ttk.Label(parent, text="Frequency (MHz):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.freq_var = tk.DoubleVar(value=self.current_site.frequency)
        freq_combo = ttk.Combobox(parent, textvariable=self.freq_var,
                                  values=[700, 800, 900, 1800, 2100, 2600, 3500], width=15)
        freq_combo.grid(row=row, column=1, pady=5, padx=5)
        freq_combo.bind('<<ComboboxSelected>>', lambda e: self.update_calculations())
        row += 1

        # Transmit Power
        ttk.Label(parent, text="TX Power (dBm):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.power_var = tk.DoubleVar(value=self.current_site.power)
        power_spin = ttk.Spinbox(parent, from_=0, to=50, textvariable=self.power_var,
                                 width=15, command=self.update_calculations)
        power_spin.grid(row=row, column=1, pady=5, padx=5)
        row += 1

        # Update button
        ttk.Button(parent, text="Update Configuration",
                   command=self.update_configuration).grid(row=row, column=0, columnspan=3, pady=20)

    def build_antenna_tab(self, parent):
        """Build antenna selection tab"""
        # Manufacturer filter
        ttk.Label(parent, text="Manufacturer:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.manufacturer_var = tk.StringVar(value="Kathrein")
        manu_combo = ttk.Combobox(parent, textvariable=self.manufacturer_var,
                                  values=list(self.antenna_db.categories.keys()), width=20)
        manu_combo.grid(row=0, column=1, pady=5, padx=5, columnspan=2)
        manu_combo.bind('<<ComboboxSelected>>', lambda e: self.update_antenna_list())

        # Antenna list
        ttk.Label(parent, text="Available Antennas:").grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)

        # Create a frame for the listbox with scrollbar
        list_frame = ttk.Frame(parent)
        list_frame.grid(row=2, column=0, columnspan=3, pady=5, sticky=tk.NSEW)

        # Configure grid weights for resizing
        parent.grid_rowconfigure(2, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_columnconfigure(1, weight=1)
        parent.grid_columnconfigure(2, weight=1)

        self.antenna_listbox = tk.Listbox(list_frame, height=10, selectmode=tk.SINGLE)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.antenna_listbox.yview)
        self.antenna_listbox.configure(yscrollcommand=scrollbar.set)

        self.antenna_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Initialize antenna list
        self.update_antenna_list()
        self.antenna_listbox.bind('<<ListboxSelect>>', self.on_antenna_select)

        # Antenna details
        details_frame = ttk.LabelFrame(parent, text="Antenna Details", padding=10)
        details_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky=tk.EW)

        self.antenna_details = tk.Text(details_frame, height=8, width=40, wrap=tk.WORD)
        self.antenna_details.pack(fill=tk.BOTH, expand=True)

        # Search criteria
        ttk.Label(parent, text="Search by Frequency (MHz):").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.search_freq_var = tk.DoubleVar(value=900)
        ttk.Entry(parent, textvariable=self.search_freq_var, width=15).grid(row=4, column=1, pady=5, padx=5)

        ttk.Button(parent, text="Search Antennas",
                   command=self.search_antennas).grid(row=4, column=2, pady=5, padx=5)

    def build_calculations_tab(self, parent):
        """Build calculations tab"""
        # Path Loss Calculator
        pl_frame = ttk.LabelFrame(parent, text="Path Loss Calculator", padding=10)
        pl_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky=tk.EW, padx=5)

        ttk.Label(pl_frame, text="Distance (km):").grid(row=0, column=0, sticky=tk.W)
        self.distance_var = tk.DoubleVar(value=1.0)
        ttk.Entry(pl_frame, textvariable=self.distance_var, width=10).grid(row=0, column=1, padx=5)

        ttk.Label(pl_frame, text="RX Height (m):").grid(row=1, column=0, sticky=tk.W)
        self.rx_height_var = tk.DoubleVar(value=1.5)
        ttk.Entry(pl_frame, textvariable=self.rx_height_var, width=10).grid(row=1, column=1, padx=5)

        ttk.Label(pl_frame, text="Environment:").grid(row=2, column=0, sticky=tk.W)
        self.env_var = tk.StringVar(value="urban")
        env_combo = ttk.Combobox(pl_frame, textvariable=self.env_var,
                                 values=["urban", "suburban", "rural"], width=10)
        env_combo.grid(row=2, column=1, padx=5)

        ttk.Button(pl_frame, text="Calculate Path Loss",
                   command=self.calculate_path_loss).grid(row=3, column=0, columnspan=2, pady=10)

        self.path_loss_result = ttk.Label(pl_frame, text="", font=('Arial', 10, 'bold'))
        self.path_loss_result.grid(row=4, column=0, columnspan=2)

        # EIRP Calculator
        eirp_frame = ttk.LabelFrame(parent, text="EIRP Calculator", padding=10)
        eirp_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky=tk.EW, padx=5)

        ttk.Label(eirp_frame, text="TX Power (dBm):").grid(row=0, column=0, sticky=tk.W)
        self.eirp_power_var = tk.DoubleVar(value=43.0)
        ttk.Entry(eirp_frame, textvariable=self.eirp_power_var, width=10).grid(row=0, column=1, padx=5)

        ttk.Label(eirp_frame, text="Antenna Gain (dBi):").grid(row=1, column=0, sticky=tk.W)
        self.eirp_gain_var = tk.DoubleVar(value=18.0)
        ttk.Entry(eirp_frame, textvariable=self.eirp_gain_var, width=10).grid(row=1, column=1, padx=5)

        ttk.Label(eirp_frame, text="Feeder Loss (dB):").grid(row=2, column=0, sticky=tk.W)
        self.feeder_loss_var = tk.DoubleVar(value=2.0)
        ttk.Entry(eirp_frame, textvariable=self.feeder_loss_var, width=10).grid(row=2, column=1, padx=5)

        ttk.Button(eirp_frame, text="Calculate EIRP",
                   command=self.calculate_eirp).grid(row=3, column=0, columnspan=2, pady=10)

        self.eirp_result = ttk.Label(eirp_frame, text="", font=('Arial', 10, 'bold'))
        self.eirp_result.grid(row=4, column=0, columnspan=2)

        # Coverage Radius
        cov_frame = ttk.LabelFrame(parent, text="Coverage Estimation", padding=10)
        cov_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky=tk.EW, padx=5)

        ttk.Label(cov_frame, text="RX Sensitivity (dBm):").grid(row=0, column=0, sticky=tk.W)
        self.rx_sens_var = tk.DoubleVar(value=-95.0)
        ttk.Entry(cov_frame, textvariable=self.rx_sens_var, width=10).grid(row=0, column=1, padx=5)

        ttk.Button(cov_frame, text="Estimate Coverage",
                   command=self.estimate_coverage).grid(row=1, column=0, columnspan=2, pady=10)

        self.coverage_result = ttk.Label(cov_frame, text="", font=('Arial', 10, 'bold'))
        self.coverage_result.grid(row=2, column=0, columnspan=2)

    def build_tools_tab(self, parent):
        """Build tools tab"""
        # Unit Converter
        conv_frame = ttk.LabelFrame(parent, text="Unit Converter", padding=10)
        conv_frame.grid(row=0, column=0, pady=5, padx=5, sticky=tk.EW)

        ttk.Label(conv_frame, text="dBm to Watts:").grid(row=0, column=0, sticky=tk.W)
        self.dbm_var = tk.DoubleVar(value=30.0)
        ttk.Entry(conv_frame, textvariable=self.dbm_var, width=10).grid(row=0, column=1, padx=5)
        ttk.Button(conv_frame, text="Convert",
                   command=lambda: self.convert_units('dbm_to_watts')).grid(row=0, column=2, padx=5)
        self.watts_result = ttk.Label(conv_frame, text="")
        self.watts_result.grid(row=0, column=3, padx=10)

        ttk.Label(conv_frame, text="Watts to dBm:").grid(row=1, column=0, sticky=tk.W)
        self.watts_in_var = tk.DoubleVar(value=1.0)
        ttk.Entry(conv_frame, textvariable=self.watts_in_var, width=10).grid(row=1, column=1, padx=5)
        ttk.Button(conv_frame, text="Convert",
                   command=lambda: self.convert_units('watts_to_dbm')).grid(row=1, column=2, padx=5)
        self.dbm_result = ttk.Label(conv_frame, text="")
        self.dbm_result.grid(row=1, column=3, padx=10)

        # Export/Import
        io_frame = ttk.LabelFrame(parent, text="Import/Export", padding=10)
        io_frame.grid(row=1, column=0, pady=10, padx=5, sticky=tk.EW)

        ttk.Button(io_frame, text="Export Site Configuration",
                   command=self.export_config).pack(pady=5)
        ttk.Button(io_frame, text="Import Configuration",
                   command=self.import_config).pack(pady=5)
        ttk.Button(io_frame, text="Save Antenna Pattern",
                   command=self.save_pattern).pack(pady=5)

        # Quick Calculations
        quick_frame = ttk.LabelFrame(parent, text="Quick Calculations", padding=10)
        quick_frame.grid(row=2, column=0, pady=10, padx=5, sticky=tk.EW)

        ttk.Button(quick_frame, text="Calculate Total Tilt",
                   command=self.calculate_total_tilt).pack(pady=2)
        ttk.Button(quick_frame, text="Free Space Path Loss",
                   command=self.calculate_free_space_loss).pack(pady=2)
        ttk.Button(quick_frame, text="Beamwidth Calculator",
                   command=self.calculate_beamwidth).pack(pady=2)

    def build_visualization_panel(self, parent):
        """Build visualization panel with plots"""
        # Create notebook for different visualizations
        viz_notebook = ttk.Notebook(parent)
        viz_notebook.pack(fill=tk.BOTH, expand=True)

        # Antenna Pattern Tab
        pattern_frame = ttk.Frame(viz_notebook)
        viz_notebook.add(pattern_frame, text="Antenna Pattern")

        # Create matplotlib figure for antenna pattern
        self.fig_pattern = Figure(figsize=(8, 6), dpi=100)
        self.ax_pattern = self.fig_pattern.add_subplot(111, projection='polar')

        self.canvas_pattern = FigureCanvasTkAgg(self.fig_pattern, pattern_frame)
        self.canvas_pattern.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Add toolbar
        toolbar_frame = ttk.Frame(pattern_frame)
        toolbar_frame.pack(fill=tk.X)
        NavigationToolbar2Tk(self.canvas_pattern, toolbar_frame)

        # Site View Tab
        site_view_frame = ttk.Frame(viz_notebook)
        viz_notebook.add(site_view_frame, text="Site View")

        self.fig_site = Figure(figsize=(8, 6), dpi=100)
        self.ax_site = self.fig_site.add_subplot(111)

        self.canvas_site = FigureCanvasTkAgg(self.fig_site, site_view_frame)
        self.canvas_site.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Coverage Map Tab
        coverage_frame = ttk.Frame(viz_notebook)
        viz_notebook.add(coverage_frame, text="Coverage Map")

        self.fig_coverage = Figure(figsize=(8, 6), dpi=100)
        self.ax_coverage = self.fig_coverage.add_subplot(111)

        self.canvas_coverage = FigureCanvasTkAgg(self.fig_coverage, coverage_frame)
        self.canvas_coverage.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Results Display
        results_frame = ttk.LabelFrame(parent, text="Calculation Results", padding=10)
        results_frame.pack(fill=tk.X, pady=10)

        self.results_text = tk.Text(results_frame, height=6, width=80, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)

        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def setup_status_bar(self):
        """Setup status bar at bottom"""
        status_bar = ttk.Frame(self.root, relief=tk.SUNKEN)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.status_label = ttk.Label(status_bar, text="Ready")
        self.status_label.pack(side=tk.LEFT, padx=5)

        version_label = ttk.Label(status_bar, text="RF Antenna Calculator v1.0 - Fadzli Edition")
        version_label.pack(side=tk.RIGHT, padx=5)

    # ==============================================
    # EVENT HANDLERS
    # ==============================================

    def update_azimuth_display(self):
        """Update azimuth display"""
        self.azimuth_display.config(text=f"{self.azimuth_var.get():.1f}°")
        self.update_plots()

    def update_tilt_display(self):
        """Update tilt display"""
        total_tilt = self.rf_calc.calculate_total_tilt(
            self.mech_tilt_var.get(),
            self.elec_tilt_var.get()
        )
        self.tilt_display.config(text=f"Total: {total_tilt:.1f}°")
        self.update_plots()

    def update_antenna_list(self):
        """Update antenna list based on manufacturer"""
        manufacturer = self.manufacturer_var.get()
        antennas = self.antenna_db.get_antennas_by_manufacturer(manufacturer)

        self.antenna_listbox.delete(0, tk.END)
        for antenna in antennas:
            self.antenna_listbox.insert(tk.END, antenna)

        # Select first antenna if available
        if antennas:
            self.antenna_listbox.selection_set(0)
            # Don't call on_antenna_select here - let the user click or we'll do it after UI is fully built

    def on_antenna_select(self, event):
        """Handle antenna selection"""
        selection = self.antenna_listbox.curselection()
        if selection:
            antenna_name = self.antenna_listbox.get(selection[0])
            antenna = self.antenna_db.get_antenna_by_name(antenna_name)

            if antenna:
                # Update current site
                self.current_site.antenna_type = antenna_name

                # Display antenna details
                details = f"""Manufacturer: {antenna.manufacturer}
Model: {antenna.model}
Frequency Range: {antenna.frequency_range[0]}-{antenna.frequency_range[1]} MHz
Gain: {antenna.gain} dBi
Horizontal Beamwidth: {antenna.horizontal_bw}°
Vertical Beamwidth: {antenna.vertical_bw}°
Front-to-Back Ratio: {antenna.front_to_back} dB
Electrical Tilt Range: {antenna.tilt_range[0]} to {antenna.tilt_range[1]}°
Connector: {antenna.connector_type}
Polarization: {antenna.polarization}
"""
                self.antenna_details.delete(1.0, tk.END)
                self.antenna_details.insert(1.0, details.strip())

                # Update plots
                self.update_plots()

    def update_configuration(self):
        """Update site configuration from UI"""
        try:
            self.current_site.site_id = self.site_id_var.get()
            self.current_site.latitude = self.lat_var.get()
            self.current_site.longitude = self.lon_var.get()
            self.current_site.antenna_height = self.height_var.get()
            self.current_site.azimuth = self.azimuth_var.get()
            self.current_site.mechanical_tilt = self.mech_tilt_var.get()
            self.current_site.electrical_tilt = self.elec_tilt_var.get()
            self.current_site.frequency = self.freq_var.get()
            self.current_site.power = self.power_var.get()

            messagebox.showinfo("Success", "Configuration updated successfully!")
            self.update_plots()

        except ValueError as e:
            messagebox.showerror("Error", f"Invalid input: {str(e)}")

    def update_calculations(self):
        """Update all calculations based on current configuration"""
        self.update_plots()

    def update_plots(self):
        """Update all plots"""
        try:
            # Get current antenna
            antenna = self.antenna_db.get_antenna_by_name(self.current_site.antenna_type)
            if not antenna:
                return

            # Update antenna pattern plot
            self.update_pattern_plot(antenna)

            # Update site view plot
            self.update_site_view()

            # Update coverage plot
            self.update_coverage_plot(antenna)

        except Exception as e:
            print(f"Error updating plots: {e}")

    def update_pattern_plot(self, antenna):
        """Update antenna pattern visualization"""
        self.ax_pattern.clear()

        if antenna and 'horizontal' in antenna.patterns:
            pattern = antenna.patterns['horizontal']
            theta = np.linspace(0, 2 * np.pi, len(pattern))

            # Convert pattern to radians and normalize
            pattern_db = np.array(pattern)
            pattern_linear = 10 ** (pattern_db / 20)  # Convert dB to linear

            # Plot horizontal pattern
            self.ax_pattern.plot(theta, pattern_linear, 'b-', linewidth=2, label='Horizontal')

            # Apply tilt if any
            total_tilt = self.rf_calc.calculate_total_tilt(
                self.current_site.mechanical_tilt,
                self.current_site.electrical_tilt
            )

            # Rotate pattern by azimuth
            rotation = np.radians(self.current_site.azimuth)
            self.ax_pattern.set_theta_offset(rotation)

            self.ax_pattern.set_title(f"Antenna Pattern: {antenna.name}\n"
                                      f"Azimuth: {self.current_site.azimuth:.1f}°, "
                                      f"Total Tilt: {total_tilt:.1f}°")
            self.ax_pattern.grid(True)
            self.ax_pattern.set_ylim(0, 1)
            self.ax_pattern.legend()

        self.canvas_pattern.draw()

    def update_site_view(self):
        """Update site visualization"""
        self.ax_site.clear()

        # Create a simple site diagram
        # Antenna tower
        tower_height = self.current_site.antenna_height
        self.ax_site.plot([0, 0], [0, tower_height], 'k-', linewidth=3, label='Tower')

        # Antenna
        tilt_rad = np.radians(self.current_site.mechanical_tilt)
        ant_length = tower_height * 0.1
        ant_x = ant_length * np.sin(tilt_rad)
        ant_y = tower_height + ant_length * np.cos(tilt_rad)

        self.ax_site.plot([0, ant_x], [tower_height, ant_y], 'r-', linewidth=2, label='Antenna')
        self.ax_site.plot(ant_x, ant_y, 'ro', markersize=10)

        # Beam pattern (simplified)
        beam_angle = np.radians(10)  # Simplified beamwidth
        for angle in [tilt_rad - beam_angle / 2, tilt_rad + beam_angle / 2]:
            beam_x = tower_height * 1.5 * np.sin(angle)
            beam_y = tower_height + tower_height * 1.5 * np.cos(angle)
            self.ax_site.plot([0, beam_x], [tower_height, beam_y], 'g--', alpha=0.5)

        self.ax_site.set_xlim(-tower_height * 2, tower_height * 2)
        self.ax_site.set_ylim(0, tower_height * 2)
        self.ax_site.set_aspect('equal')
        self.ax_site.set_title(f"Site View: {self.current_site.site_id}\n"
                               f"Height: {tower_height}m, Azimuth: {self.current_site.azimuth:.1f}°")
        self.ax_site.grid(True, alpha=0.3)
        self.ax_site.legend()

        self.canvas_site.draw()

    def update_coverage_plot(self, antenna):
        """Update coverage estimation plot"""
        self.ax_coverage.clear()

        # Calculate coverage radius
        eirp = self.rf_calc.calculate_eirp(
            self.current_site.power,
            antenna.gain if antenna else 18.0
        )

        radius = self.rf_calc.calculate_coverage_radius(
            eirp,
            -95,  # Default sensitivity
            self.current_site.frequency,
            "urban"
        )

        # Create coverage circle
        theta = np.linspace(0, 2 * np.pi, 100)
        x = radius * np.cos(theta)
        y = radius * np.sin(theta)

        # Rotate by azimuth
        azimuth_rad = np.radians(self.current_site.azimuth)
        x_rot = x * np.cos(azimuth_rad) - y * np.sin(azimuth_rad)
        y_rot = x * np.sin(azimuth_rad) + y * np.cos(azimuth_rad)

        self.ax_coverage.plot(x_rot, y_rot, 'b-', alpha=0.5, label='Coverage')
        self.ax_coverage.fill(x_rot, y_rot, 'b', alpha=0.1)

        # Site location
        self.ax_coverage.plot(0, 0, 'r^', markersize=15, label='Site')

        # Sector lines
        sector_width = np.radians(65)  # Typical sector width
        for angle in [azimuth_rad - sector_width / 2, azimuth_rad + sector_width / 2]:
            line_x = radius * 1.2 * np.cos(angle)
            line_y = radius * 1.2 * np.sin(angle)
            self.ax_coverage.plot([0, line_x], [0, line_y], 'r--', alpha=0.7)

        self.ax_coverage.set_xlim(-radius * 1.5, radius * 1.5)
        self.ax_coverage.set_ylim(-radius * 1.5, radius * 1.5)
        self.ax_coverage.set_aspect('equal')
        self.ax_coverage.set_title(f"Coverage Estimation\nRadius: ~{radius:.2f} km")
        self.ax_coverage.grid(True, alpha=0.3)
        self.ax_coverage.legend()

        self.canvas_coverage.draw()

    # ==============================================
    # CALCULATION METHODS
    # ==============================================

    def calculate_path_loss(self):
        """Calculate path loss using COST-231 Hata model"""
        try:
            distance = self.distance_var.get()
            rx_height = self.rx_height_var.get()
            environment = self.env_var.get()

            pl = self.rf_calc.cost231_hata_path_loss(
                distance,
                self.current_site.frequency,
                self.current_site.antenna_height,
                rx_height,
                environment
            )

            # Calculate received power
            antenna = self.antenna_db.get_antenna_by_name(self.current_site.antenna_type)
            eirp = self.rf_calc.calculate_eirp(
                self.current_site.power,
                antenna.gain if antenna else 18.0
            )
            rx_power = eirp - pl

            result = (f"Path Loss ({environment}): {pl:.2f} dB\n"
                      f"Distance: {distance} km\n"
                      f"EIRP: {eirp:.2f} dBm\n"
                      f"Estimated RX Power: {rx_power:.2f} dBm")

            self.path_loss_result.config(text=result)
            self.add_to_results("Path Loss Calculation", result)

        except Exception as e:
            messagebox.showerror("Error", f"Calculation failed: {str(e)}")

    def calculate_eirp(self):
        """Calculate EIRP"""
        try:
            eirp = self.rf_calc.calculate_eirp(
                self.eirp_power_var.get(),
                self.eirp_gain_var.get(),
                self.feeder_loss_var.get()
            )

            result = f"EIRP: {eirp:.2f} dBm"
            self.eirp_result.config(text=result)
            self.add_to_results("EIRP Calculation", result)

        except Exception as e:
            messagebox.showerror("Error", f"Calculation failed: {str(e)}")

    def estimate_coverage(self):
        """Estimate coverage radius"""
        try:
            antenna = self.antenna_db.get_antenna_by_name(self.current_site.antenna_type)
            eirp = self.rf_calc.calculate_eirp(
                self.current_site.power,
                antenna.gain if antenna else 18.0
            )

            radius = self.rf_calc.calculate_coverage_radius(
                eirp,
                self.rx_sens_var.get(),
                self.current_site.frequency,
                "urban"
            )

            result = f"Estimated Coverage Radius: {radius:.2f} km"
            self.coverage_result.config(text=result)
            self.add_to_results("Coverage Estimation", result)

        except Exception as e:
            messagebox.showerror("Error", f"Calculation failed: {str(e)}")

    def convert_units(self, conversion_type):
        """Handle unit conversions"""
        try:
            if conversion_type == 'dbm_to_watts':
                watts = self.rf_calc.dbm_to_watts(self.dbm_var.get())
                self.watts_result.config(text=f"{watts:.6f} W")
                self.add_to_results("Unit Conversion",
                                    f"{self.dbm_var.get():.2f} dBm = {watts:.6f} W")

            elif conversion_type == 'watts_to_dbm':
                dbm = self.rf_calc.watts_to_dbm(self.watts_in_var.get())
                self.dbm_result.config(text=f"{dbm:.2f} dBm")
                self.add_to_results("Unit Conversion",
                                    f"{self.watts_in_var.get():.6f} W = {dbm:.2f} dBm")

        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

    def calculate_total_tilt(self):
        """Calculate total tilt"""
        total = self.rf_calc.calculate_total_tilt(
            self.mech_tilt_var.get(),
            self.elec_tilt_var.get()
        )
        result = f"Total Tilt: {total:.2f}°"
        messagebox.showinfo("Total Tilt", result)
        self.add_to_results("Tilt Calculation", result)

    def calculate_free_space_loss(self):
        """Calculate free space path loss"""
        loss = self.rf_calc.free_space_path_loss(1.0, 900.0)
        result = f"Free Space Path Loss (1km @ 900MHz): {loss:.2f} dB"
        messagebox.showinfo("Free Space Loss", result)
        self.add_to_results("Free Space Loss", result)

    def calculate_beamwidth(self):
        """Calculate beamwidth-related parameters"""
        antenna = self.antenna_db.get_antenna_by_name(self.current_site.antenna_type)
        if antenna:
            result = (f"Antenna: {antenna.name}\n"
                      f"Horizontal Beamwidth: {antenna.horizontal_bw}°\n"
                      f"Vertical Beamwidth: {antenna.vertical_bw}°\n"
                      f"Gain: {antenna.gain} dBi")
            messagebox.showinfo("Beamwidth Calculator", result)
            self.add_to_results("Beamwidth Calculator", result)

    def search_antennas(self):
        """Search antennas based on criteria"""
        try:
            frequency = self.search_freq_var.get()
            results = self.antenna_db.search_antennas(frequency=frequency)

            if results:
                result_text = f"Found {len(results)} antennas at {frequency} MHz:\n\n"
                for ant in results:
                    result_text += f"- {ant.name} ({ant.manufacturer}): "
                    result_text += f"{ant.gain} dBi, {ant.horizontal_bw}° HBW\n"

                messagebox.showinfo("Antenna Search Results", result_text)
                self.add_to_results("Antenna Search", result_text)
            else:
                messagebox.showinfo("Antenna Search", "No antennas found matching criteria")

        except Exception as e:
            messagebox.showerror("Error", f"Search failed: {str(e)}")

    # ==============================================
    # FILE OPERATIONS
    # ==============================================

    def export_config(self):
        """Export site configuration to file"""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )

            if filename:
                config = {
                    'site_id': self.current_site.site_id,
                    'latitude': self.current_site.latitude,
                    'longitude': self.current_site.longitude,
                    'antenna_height': self.current_site.antenna_height,
                    'azimuth': self.current_site.azimuth,
                    'mechanical_tilt': self.current_site.mechanical_tilt,
                    'electrical_tilt': self.current_site.electrical_tilt,
                    'antenna_type': self.current_site.antenna_type,
                    'frequency': self.current_site.frequency,
                    'power': self.current_site.power
                }

                with open(filename, 'w') as f:
                    json.dump(config, f, indent=4)

                messagebox.showinfo("Success", f"Configuration exported to {filename}")
                self.add_to_results("Export", f"Exported configuration to {filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")

    def import_config(self):
        """Import site configuration from file"""
        try:
            filename = filedialog.askopenfilename(
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )

            if filename:
                with open(filename, 'r') as f:
                    config = json.load(f)

                # Update UI variables
                self.site_id_var.set(config.get('site_id', 'SITE001'))
                self.lat_var.set(config.get('latitude', 0.0))
                self.lon_var.set(config.get('longitude', 0.0))
                self.height_var.set(config.get('antenna_height', 30.0))
                self.azimuth_var.set(config.get('azimuth', 0.0))
                self.mech_tilt_var.set(config.get('mechanical_tilt', 0.0))
                self.elec_tilt_var.set(config.get('electrical_tilt', 0.0))
                self.freq_var.set(config.get('frequency', 900.0))
                self.power_var.set(config.get('power', 43.0))

                # Update antenna if specified
                antenna_type = config.get('antenna_type', '')
                if antenna_type:
                    self.current_site.antenna_type = antenna_type

                # Update configuration
                self.update_configuration()

                messagebox.showinfo("Success", f"Configuration imported from {filename}")
                self.add_to_results("Import", f"Imported configuration from {filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Import failed: {str(e)}")

    def save_pattern(self):
        """Save antenna pattern to file"""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("All files", "*.*")]
            )

            if filename:
                self.fig_pattern.savefig(filename, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Success", f"Pattern saved to {filename}")
                self.add_to_results("Export", f"Saved antenna pattern to {filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Save failed: {str(e)}")

    def add_to_results(self, title, content):
        """Add calculation results to results display"""
        self.results_text.insert(tk.END, f"\n{'=' * 60}\n")
        self.results_text.insert(tk.END, f"{title}\n")
        self.results_text.insert(tk.END, f"{'=' * 60}\n")
        self.results_text.insert(tk.END, f"{content}\n")
        self.results_text.see(tk.END)


# ==============================================
# MAIN ENTRY POINT
# ==============================================

def main():
    """Main entry point"""
    root = tk.Tk()
    app = KathreinCalculatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()