import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from sklearn.cluster import AgglomerativeClustering
from scipy.spatial import ConvexHull
from shapely.geometry import Polygon, Point
import threading
import os
import sys

# Import display preferences
try:
    from display_preferences import (
        initialize as init_display_prefs,
        get_show_names,
        set_show_names,
        register_callback
    )
    DISPLAY_PREFS_AVAILABLE = True
except ImportError as e:
    DISPLAY_PREFS_AVAILABLE = False
    # Create stub functions so the code doesn't crash
    def init_display_prefs(dir): pass
    def get_show_names(): return False
    def set_show_names(val): pass
    def register_callback(func): pass
except Exception as e:
    DISPLAY_PREFS_AVAILABLE = False
    # Create stub functions so the code doesn't crash
    def init_display_prefs(dir): pass
    def get_show_names(): return False
    def set_show_names(val): pass
    def register_callback(func): pass

# Outlook Category Colors Enumeration (OlCategoryColor)
# All 25 available colors in Outlook
OUTLOOK_COLORS = {
    0: "None",
    1: "Red",
    2: "Orange", 
    3: "Peach",
    4: "Yellow",
    5: "Green",
    6: "Teal",
    7: "Olive",
    8: "Blue",
    9: "Purple",
    10: "Maroon",
    11: "Steel",
    12: "DarkSteel",
    13: "Gray",
    14: "DarkGray",
    15: "Black",
    16: "DarkRed",
    17: "DarkOrange",
    18: "DarkPeach",
    19: "DarkYellow",
    20: "DarkGreen",
    21: "DarkTeal",
    22: "DarkOlive",
    23: "DarkBlue",
    24: "DarkPurple"
}


class TSPClusteringApp:
    def __init__(self, root, project_dir=None):
        self.root = root
        self.root.title("TSP Regional Clustering Optimizer")
        self.root.geometry("1400x900")
        
        # Project directory from command line
        self.project_dir = project_dir
        
        # Variables
        self.locations_file = None
        self.distances_file = None
        self.output_dir = None
        self.customers = None
        self.distance_matrix = None
        self.depot_location = None
        self.canvas = None
        self.toolbar = None
        self.log_window = None
        
        # Store clustering results for saving
        self.clustered_results = None
        self.summary_results = None
        self.has_results = False
        
        # Configuration variables
        self.num_regions_var = tk.StringVar(value="6")
        self.depot_postcode_var = tk.StringVar(value="")
        self.service_time_var = tk.StringVar(value="1.0")
        self.work_hours_var = tk.StringVar(value="8")
        
        # Store available postcodes for depot selection
        self.available_postcodes = []
        
        # Store custom region names
        self.region_names = {}  # {region_number: custom_name}
        
        # Store region colors (Outlook color codes)
        self.region_colors = {}  # {region_number: color_index}
        
        # Initialize display preferences
        if DISPLAY_PREFS_AVAILABLE:
            try:
                init_display_prefs(self.project_dir if self.project_dir else os.getcwd())
                register_callback(self.on_display_preference_changed)
            except Exception as e:
                print(f"Warning: Could not initialize display preferences: {e}")
        
        self.setup_ui()
        
        # Auto-load project files if project directory provided
        if self.project_dir:
            self.auto_load_project_files()
        
    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="5")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Button bar at top for quick menu access
        button_bar = ttk.Frame(main_frame)
        button_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.config_btn = ttk.Button(button_bar, text="Create Regions", command=self.show_config_menu, width=18)
        self.config_btn.pack(side=tk.LEFT, padx=2)
        self.save_btn = ttk.Button(button_bar, text="Save Results", command=self.save_results, width=12, state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT, padx=2)
        self.edit_btn = ttk.Button(button_bar, text="Edit Regions", command=self.show_edit_regions_dialog, width=12, state=tk.DISABLED)
        self.edit_btn.pack(side=tk.LEFT, padx=2)
        self.rename_color_btn = ttk.Button(button_bar, text="Rename/Recolor", command=self.show_rename_recolor_dialog, width=15, state=tk.DISABLED)
        self.rename_color_btn.pack(side=tk.LEFT, padx=2)
        
        # Add toggle button on the right
        self.toggle_btn = ttk.Button(button_bar, text="Show Postcodes", 
                                    command=self.toggle_display_preference, width=18)
        self.toggle_btn.pack(side=tk.RIGHT, padx=(10, 0))
        self.update_toggle_button_text()
        self.rename_color_btn.pack(side=tk.LEFT, padx=2)
        self.view_btn = ttk.Button(button_bar, text="Analytics", command=self.show_log_window, width=12)
        self.view_btn.pack(side=tk.LEFT, padx=2)
        
        # Visualization frame (main content area)
        self.viz_frame = ttk.LabelFrame(main_frame, text="Visualization", padding="5")
        self.viz_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.viz_frame.columnconfigure(0, weight=1)
        self.viz_frame.rowconfigure(0, weight=1)
        self.viz_canvas_container = None
        
        # Welcome message in viz area
        welcome_frame = ttk.Frame(self.viz_frame)
        welcome_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Label(welcome_frame, text="TSP Regional Clustering Optimizer", 
                 font=('Arial', 20, 'bold')).pack(pady=50)
        ttk.Label(welcome_frame, text="Click the File button to load your data files", 
                 font=('Arial', 12)).pack(pady=10)
        ttk.Label(welcome_frame, text="Configure clustering parameters with the Configure and Run button", 
                 font=('Arial', 12)).pack(pady=10)
        ttk.Label(welcome_frame, text="Start clustering analysis from Configure and Run or Run", 
                 font=('Arial', 12)).pack(pady=10)
        
        # Progress Section at bottom
        progress_frame = ttk.Frame(main_frame, relief=tk.SUNKEN, borderwidth=1)
        progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=3)
        
        self.status_label = ttk.Label(progress_frame, text="Ready", foreground="green", width=30)
        self.status_label.pack(side=tk.RIGHT, padx=5)
        
        # Hidden log text (for internal use)
        self.log_text = scrolledtext.ScrolledText(self.root, height=20, width=80, 
                                                  font=('Consolas', 9))
        # Not displayed in main window
        
        # Try to set accent button style
        try:
            style = ttk.Style()
            style.configure('Accent.TButton', font=('Arial', 10, 'bold'))
        except:
            pass
        
        # Set close protocol to exit immediately
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def on_closing(self):
        """Handle window close event"""
        if self.log_window:
            self.log_window.destroy()
        self.root.destroy()

    def _assign_location_instance_ids(self, df):
        """Add stable per-postcode instance identifiers to distinguish duplicate postcodes."""
        if 'postcode' not in df.columns:
            return df

        out = df.copy()
        out['postcode'] = out['postcode'].astype(str)

        if 'location_instance' not in out.columns:
            out['location_instance'] = out.groupby('postcode').cumcount() + 1

        if 'location_id' not in out.columns:
            out['location_id'] = out['postcode'] + '#' + out['location_instance'].astype(int).astype(str)

        return out
    
    def toggle_display_preference(self):
        """Toggle between showing names and postcodes"""
        print("[DEBUG] Toggle button clicked!")
        try:
            current = get_show_names()
            print(f"[DEBUG] Current preference: show_names = {current}")
            new_value = not current
            set_show_names(new_value)
            print(f"[DEBUG] New preference set to: {new_value}")
            self.update_toggle_button_text()
            print("[DEBUG] Toggle button text updated")
        except Exception as e:
            print(f"[DEBUG] Error in toggle_display_preference: {e}")
            import traceback
            traceback.print_exc()
    
    def update_toggle_button_text(self):
        """Update toggle button text based on current preference"""
        print("[DEBUG] update_toggle_button_text called")
        if hasattr(self, 'toggle_btn'):
            try:
                show_names = get_show_names()
                print(f"[DEBUG] get_show_names() returned: {show_names}")
                if show_names:
                    self.toggle_btn.config(text="Show Postcodes")
                    print("[DEBUG] Button set to 'Show Postcodes'")
                else:
                    self.toggle_btn.config(text="Show Names")
                    print("[DEBUG] Button set to 'Show Names'")
            except Exception as e:
                print(f"[DEBUG] Error updating button text: {e}")
                self.toggle_btn.config(text="Display Mode")
        else:
            print("[DEBUG] toggle_btn attribute not found!")
    
    def on_display_preference_changed(self, show_names):
        """Callback when display preference changes from another app"""
        self.update_toggle_button_text()
        # Redraw visualization if we have results loaded
        if self.has_results and hasattr(self, 'coords'):
            try:
                customer_names = getattr(self, 'customer_names', [None] * len(self.customer_postcodes))
                self.create_visualization(
                    self.coords, 
                    self.labels, 
                    self.depot, 
                    self.n_clusters,
                    self.customer_postcodes,
                    customer_names,
                    self.depot_postcode
                )
            except Exception as e:
                print(f"Error redrawing visualization: {e}")
    
    def auto_load_project_files(self):
        """Auto-load files from project directory"""
        if not self.project_dir or not os.path.exists(self.project_dir):
            return
        
        project_name = os.path.basename(self.project_dir)
        self.root.title(f"TSP Regional Clustering Optimizer - Project: {project_name}")
        
        # Set output directory to project directory
        self.output_dir = self.project_dir
        
        # Load locations.csv
        locations_path = os.path.join(self.project_dir, "locations.csv")
        if os.path.exists(locations_path):
            self.locations_file = locations_path
            self.log(f"✓ Auto-loaded: {locations_path}")
        else:
            self.log(f"⚠ locations.csv not found in project directory")
        
        # Load distances.csv
        distances_path = os.path.join(self.project_dir, "distances.csv")
        if os.path.exists(distances_path):
            self.distances_file = distances_path
            self.log(f"✓ Auto-loaded: {distances_path}")
        else:
            self.log(f"⚠ distances.csv not found in project directory")
        
        # Check if previous clustering exists and load it
        clustered_file = os.path.join(self.project_dir, "clustered_regions.csv")
        if os.path.exists(clustered_file):
            self.log(f"\n✓ Found previous clustering results - loading...")
            self.load_previous_clustering()
        elif self.locations_file and self.distances_file:
            # Load initial visualization if no clustering exists
            self.load_and_display_initial_visualization()
            self.log(f"\n✓ Project '{project_name}' loaded successfully")
        else:
            self.log(f"\n⚠ Project '{project_name}' loaded with missing files")
    
    def show_config_menu(self):
        """Show Configure dialog - directly open clustering parameters"""
        self.show_config_dialog()
    
    def show_config_dialog(self):
        """Show configuration dialog window"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Clustering Parameters")
        dialog.geometry("450x350")
        dialog.transient(self.root)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Number of regions
        ttk.Label(frame, text="Desired Number of Regions:", 
                 font=('Arial', 10)).grid(row=0, column=0, sticky=tk.W, pady=10)
        ttk.Spinbox(frame, from_=2, to=20, textvariable=self.num_regions_var, 
                   width=15).grid(row=0, column=1, sticky=tk.W, pady=10, padx=10)
        
        # Depot postcode
        ttk.Label(frame, text="Home Base Postcode:", 
                 font=('Arial', 10)).grid(row=1, column=0, sticky=tk.W, pady=10)
        self.depot_combo = ttk.Combobox(frame, textvariable=self.depot_postcode_var, 
                 width=15, state='readonly')
        self.depot_combo.grid(row=1, column=1, sticky=tk.W, pady=10, padx=10)
        self.depot_combo['values'] = self.available_postcodes
        ttk.Label(frame, text="(Required - select from list)", 
                 foreground="red", font=('Arial', 8)).grid(row=2, column=1, sticky=tk.W, padx=10)
        
        # Service time
        ttk.Label(frame, text="Service Time per Customer (hours):", 
                 font=('Arial', 10)).grid(row=3, column=0, sticky=tk.W, pady=10)
        ttk.Entry(frame, textvariable=self.service_time_var, 
                 width=15).grid(row=3, column=1, sticky=tk.W, pady=10, padx=10)
        
        # Work hours
        ttk.Label(frame, text="Work Hours per Day:", 
                 font=('Arial', 10)).grid(row=4, column=0, sticky=tk.W, pady=10)
        ttk.Spinbox(frame, from_=4, to=12, textvariable=self.work_hours_var, 
                   width=15).grid(row=4, column=1, sticky=tk.W, pady=10, padx=10)
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=20)
        
        def run_and_close():
            dialog.destroy()
            self.start_clustering()

        ttk.Button(btn_frame, text="Create Regions", command=run_and_close, 
              width=14).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy, 
              width=10).pack(side=tk.LEFT, padx=5)
    
    def show_log_window(self):
        """Show log window"""
        if self.log_window and tk.Toplevel.winfo_exists(self.log_window):
            self.log_window.lift()
            return
        
        self.log_window = tk.Toplevel(self.root)
        self.log_window.title("Analysis Log")
        self.log_window.geometry("900x600")
        
        # Create new log text widget for the window
        log_frame = ttk.Frame(self.log_window, padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        log_display = scrolledtext.ScrolledText(log_frame, height=30, width=100, 
                                               font=('Consolas', 9))
        log_display.pack(fill=tk.BOTH, expand=True)
        
        # Copy existing log content
        log_display.insert(tk.END, self.log_text.get("1.0", tk.END))
        log_display.config(state=tk.DISABLED)
        
        # Store reference to update it
        self.log_display_window = log_display
    
    def log(self, message):
        """Add message to log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Update log window if it's open
        if self.log_window and tk.Toplevel.winfo_exists(self.log_window):
            if hasattr(self, 'log_display_window'):
                self.log_display_window.config(state=tk.NORMAL)
                self.log_display_window.insert(tk.END, f"{message}\n")
                self.log_display_window.see(tk.END)
                self.log_display_window.config(state=tk.DISABLED)
        
        self.root.update_idletasks()
        
    def update_status(self, message, color="black"):
        """Update status label"""
        self.status_label.config(text=message, foreground=color)
        self.root.update_idletasks()
        

    def load_previous_clustering(self):
        """Load previously saved clustering results for editing"""
        # First ensure we have output directory and basic files
        if not self.output_dir:
            messagebox.showwarning("No Output Directory", 
                                  "Please set output directory first.\n\n"
                                  "Use 'Set Output Directory' to select the project folder.")
            return
        
        # Check for clustered_regions.csv
        clustered_file = os.path.join(self.output_dir, "clustered_regions.csv")
        if not os.path.exists(clustered_file):
            messagebox.showwarning("No Previous Results", 
                                  f"No clustered_regions.csv found in:\n{self.output_dir}\n\n"
                                  f"Run clustering analysis first to create results.")
            return
        
        try:
            self.log("\n" + "="*80)
            self.log("LOADING PREVIOUS CLUSTERING RESULTS")
            self.log("="*80)
            self.update_status("Loading previous results...", "blue")
            
            # Load clustered regions
            results_df = pd.read_csv(clustered_file)
            self.log(f"✓ Loaded {len(results_df)} locations from clustered_regions.csv")
            
            # Extract depot (region 0) and customers
            depot_row = results_df[results_df['region'] == 0]
            if depot_row.empty:
                # Fallback: use region -1 or first row
                depot_row = results_df[results_df['region'] == -1]
                if depot_row.empty:
                    depot_row = results_df.iloc[[0]]
                    self.log(f"⚠ No depot found (region 0), using first location")
            
            depot_postcode = depot_row.iloc[0]['postcode']
            depot_lat = depot_row.iloc[0]['latitude']
            depot_lon = depot_row.iloc[0]['longitude']
            depot = np.array([[depot_lat, depot_lon]])
            
            self.log(f"✓ Depot: {depot_postcode} at ({depot_lat:.4f}, {depot_lon:.4f})")
            self.depot_postcode_var.set(depot_postcode)
            
            # Get customers (exclude depot, but include excluded locations with region -1)
            customers_df = results_df[results_df['region'] != 0].copy()
            customers_df = self._assign_location_instance_ids(customers_df)
            coords = customers_df[['latitude', 'longitude']].values
            customer_postcodes = customers_df['postcode'].tolist()
            customer_location_ids = customers_df['location_id'].tolist()
            customer_location_instances = customers_df['location_instance'].astype(int).tolist()
            
            # Extract cluster labels (convert from 1-indexed to 0-indexed)
            # Region -1 stays as -1 (excluded), others convert from 1-indexed to 0-indexed
            labels = customers_df['region'].values.copy()
            labels = np.where(labels == -1, -1, labels - 1).astype(int)
            
            # Calculate n_clusters (excluding -1 which is "excluded")
            active_regions = labels[labels >= 0]
            n_clusters = int(active_regions.max() + 1) if len(active_regions) > 0 else 0
            
            excluded_count = np.sum(labels == -1)
            if excluded_count > 0:
                self.log(f"✓ Loaded {len(coords)} customers in {n_clusters} regions ({excluded_count} excluded)")
            else:
                self.log(f"✓ Loaded {len(coords)} customers in {n_clusters} regions")
            
            # Update configuration
            self.num_regions_var.set(str(n_clusters))
            
            # Load driving time matrix for minimum days calculation
            try:
                distances_file = os.path.join(self.output_dir, "distances.csv")
                if os.path.exists(distances_file):
                    distances_df = pd.read_csv(distances_file)
                    
                    # Build postcode list
                    all_postcodes = sorted(set(list(distances_df['origin'].unique()) + list(distances_df['destination'].unique())))
                    postcode_to_idx = {pc: i for i, pc in enumerate(all_postcodes)}
                    n = len(all_postcodes)
                    
                    # Initialize matrix
                    driving_time_matrix = np.full((n, n), np.inf)
                    np.fill_diagonal(driving_time_matrix, 0)
                    
                    # Fill in driving times
                    for _, row in distances_df.iterrows():
                        if row['origin'] in postcode_to_idx and row['destination'] in postcode_to_idx:
                            i = postcode_to_idx[row['origin']]
                            j = postcode_to_idx[row['destination']]
                            driving_time_matrix[i, j] = row['driving_time_minutes']
                            driving_time_matrix[j, i] = row['driving_time_minutes']
                    
                    # Store for minimum days calculation
                    self.driving_time_matrix = driving_time_matrix
                    self.customer_postcode_to_idx = {pc: postcode_to_idx[pc] for pc in customer_postcodes if pc in postcode_to_idx}
                    self.depot_postcode_idx = postcode_to_idx[depot_postcode] if depot_postcode in postcode_to_idx else 0
                    
                    self.log(f"✓ Loaded driving time matrix for minimum days calculation")
            except Exception as e:
                self.log(f"⚠ Could not load driving time matrix: {e}")
            
            # Store clustering data for editing
            self.coords = coords
            self.labels = labels
            self.depot = depot
            self.n_clusters = n_clusters
            self.customer_postcodes = customer_postcodes
            self.customer_location_ids = customer_location_ids
            self.customer_location_instances = customer_location_instances
            self.customer_names = customers_df['client_name'].tolist() if 'client_name' in customers_df.columns else [None] * len(customer_postcodes)
            self.depot_postcode = depot_postcode
            
            # Prepare results for potential saving
            self.clustered_results = self._assign_location_instance_ids(results_df)
            
            # Load summary if available
            summary_file = os.path.join(self.output_dir, "region_summary.csv")
            if os.path.exists(summary_file):
                self.summary_results = pd.read_csv(summary_file)
                self.log(f"✓ Loaded region_summary.csv")
            else:
                # Recreate summary
                summary = []
                for i in range(n_clusters):
                    region_postcodes = customers_df[customers_df['region'] == i+1]['postcode'].tolist()
                    summary.append({
                        'region': i+1,
                        'customer_count': len(region_postcodes),
                        'postcodes': ', '.join(region_postcodes)
                    })
                
                # Add excluded locations if any
                excluded_postcodes = customers_df[customers_df['region'] == -1]['postcode'].tolist()
                if excluded_postcodes:
                    summary.append({
                        'region': 'Excluded',
                        'customer_count': len(excluded_postcodes),
                        'postcodes': ', '.join(excluded_postcodes)
                    })
                
                self.summary_results = pd.DataFrame(summary)
                self.log(f"✓ Recreated region summary")
            
            self.has_results = True
            
            # Enable edit and save buttons
            self.edit_btn.config(state=tk.NORMAL)
            self.save_btn.config(state=tk.NORMAL)
            self.rename_color_btn.config(state=tk.NORMAL)
            
            # Load region names if available
            self.load_region_names()
            
            # Auto-assign default colors if not already set (starting from 1: Red)
            self.auto_assign_default_colors()
            
            # Create visualization
            self.log("\nGenerating visualization...")
            customer_names = self.customer_names if hasattr(self, 'customer_names') else [None] * len(customer_postcodes)
            self.create_visualization(coords, labels, depot, n_clusters, customer_postcodes, customer_names, depot_postcode)
            
            self.update_status("Previous clustering loaded", "green")
            self.log("="*80)
            self.log(f"PREVIOUS CLUSTERING LOADED SUCCESSFULLY")
            self.log("="*80)
            self.log(f"You can now edit regions or re-run clustering with different parameters.")
            self.log("="*80)
            
        except Exception as e:
            self.log(f"\n✗ ERROR loading previous clustering: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.update_status("Error loading results", "red")
            messagebox.showerror("Load Error", f"Error loading previous clustering:\n{e}")
    
    def load_and_display_initial_visualization(self):
        """Load data and display initial visualization after configuration"""
        if not self.locations_file or not self.distances_file or not self.output_dir:
            return
        
        try:
            self.log("\nLoading data for initial visualization...")
            self.update_status("Loading data...", "blue")
            
            # Load coordinates from distance_matrix.csv
            distance_matrix_file = os.path.join(self.output_dir, "distance_matrix.csv")
            if not os.path.exists(distance_matrix_file):
                self.log(f"⚠ distance_matrix.csv not found - run Postcode Distance Calculator first")
                self.update_status("Missing distance_matrix.csv", "orange")
                return
            
            locations_df = pd.read_csv(distance_matrix_file)
            self.log(f"✓ Loaded {len(locations_df)} locations with coordinates")
            
            # Populate available postcodes (alphabetically sorted)
            self.available_postcodes = sorted(locations_df['postcode'].unique())
            
            # Update depot combobox if it exists
            if hasattr(self, 'depot_combo'):
                self.depot_combo['values'] = self.available_postcodes
            
            # Get depot location if one is selected
            depot_postcode = self.depot_postcode_var.get().strip().upper()
            if not depot_postcode:
                self.log(f"⚠ No home base postcode selected - using first location for visualization")
                depot_row = locations_df.iloc[[0]]
            else:
                depot_row = locations_df[locations_df['postcode'].str.upper() == depot_postcode]
                
                if depot_row.empty:
                    self.log(f"⚠ Home base postcode '{depot_postcode}' not found, using first location")
                    depot_row = locations_df.iloc[[0]]
            
            depot_lat = depot_row.iloc[0]['latitude']
            depot_lon = depot_row.iloc[0]['longitude']
            depot = np.array([[depot_lat, depot_lon]])
            
            # Get all locations
            coords = locations_df[['latitude', 'longitude']].values
            postcodes = locations_df['postcode'].tolist()
            
            # Create initial visualization (no clustering yet)
            self.create_initial_visualization(coords, depot, postcodes, depot_postcode if depot_postcode else "TBD")
            self.update_status("Ready to cluster", "green")
            self.log("✓ Initial visualization displayed - ready to run clustering")
            
        except Exception as e:
            self.log(f"✗ Error loading data: {e}")
            self.update_status("Error loading data", "red")
            
    def reset_clustering(self):
        """Reset the clustering to start fresh without restarting the program"""
        response = messagebox.askyesno("Reset Clustering", 
                                       "This will clear the current clustering results.\n\n"
                                       "Files will remain loaded, but you can reconfigure and re-run.\n\n"
                                       "Continue?")
        if response:
            self.log("\n" + "="*80)
            self.log("RESET CLUSTERING")
            self.log("="*80)
            
            # Clear results
            self.clustered_results = None
            self.summary_results = None
            self.has_results = False
            self.region_names = {}
            
            # Disable edit and save buttons
            self.edit_btn.config(state=tk.DISABLED)
            self.save_btn.config(state=tk.DISABLED)
            self.rename_color_btn.config(state=tk.DISABLED)
            
            # Reset progress
            self.progress_bar['value'] = 0
            self.update_status("Ready", "green")
            
            # Reload initial visualization
            if self.locations_file and self.distances_file and self.output_dir:
                self.load_and_display_initial_visualization()
                self.log("✓ Reset complete - ready for new clustering configuration")
            else:
                # Clear visualization area
                for widget in self.viz_frame.winfo_children():
                    widget.destroy()
                self.log("✓ Reset complete - please reload data files")
    
    def save_results(self):
        """Save clustering results to CSV files"""
        if not self.has_results or self.clustered_results is None:
            messagebox.showwarning("No Results", 
                                  "No clustering results to save.\n\n"
                                  "Run clustering analysis first.")
            return
        
        if not self.output_dir:
            messagebox.showwarning("No Output Directory", 
                                  "Please set output directory first.")
            return
        
        try:
            self.log("\n" + "="*80)
            self.log("SAVING RESULTS")
            self.log("="*80)
            
            # Calculate minimum days for each region
            self.log("\nCalculating minimum days required for each region...")
            
            # Check if we have necessary data
            if not hasattr(self, 'driving_time_matrix'):
                self.log("⚠ Warning: No driving time matrix available - cannot calculate minimum days")
                self.log("  Minimum days will be set to 0. Run clustering first to calculate properly.")
            
            # Always recalculate minimum_days (don't skip if column exists)
            minimum_days_list = []
            
            for _, row in self.summary_results.iterrows():
                region_num = row['region']
                
                # Skip if region is 'Excluded' or not a valid number
                if region_num == 'Excluded' or not isinstance(region_num, (int, float)):
                    minimum_days_list.append(0)
                    continue
                
                # Ensure region_num is an integer
                region_num = int(region_num)
                
                min_days = self.calculate_minimum_days_for_region(region_num)
                minimum_days_list.append(min_days)
                
                region_name = self.get_region_display_name(region_num)
                self.log(f"  {region_name}: {row['customer_count']} customers → {min_days} days minimum")
            
            self.summary_results['minimum_days'] = minimum_days_list
            self.log("✓ Minimum days calculated for all regions")
            
            # Save clustered regions
            output_file = os.path.join(self.output_dir, "clustered_regions.csv")
            self.clustered_results.to_csv(output_file, index=False)
            self.log(f"\n✓ Saved: {output_file}")
            
            # Save summary with minimum days
            summary_file = os.path.join(self.output_dir, "region_summary.csv")
            self.summary_results.to_csv(summary_file, index=False)
            self.log(f"✓ Saved: {summary_file}")
            
            self.log("="*80)
            
            messagebox.showinfo("Success", 
                              f"Results saved successfully!\n\n"
                              f"Files saved to:\n{self.output_dir}\n\n"
                              f"• clustered_regions.csv\n"
                              f"• region_summary.csv\n\n"
                              f"Minimum days calculated for scheduling.")
            
        except Exception as e:
            self.log(f"\n✗ ERROR saving results: {e}")
            messagebox.showerror("Save Error", f"Error saving results:\n{e}")
    
    def start_clustering(self):
        """Start the clustering process in a separate thread"""
        # Check if files are loaded
        if not self.locations_file or not self.distances_file or not self.output_dir:
            messagebox.showwarning("Missing Data", 
                                  "Please load locations CSV, distances CSV, and set output directory first.\n\n"
                                  "Use File menu to load data files.")
            return
        
        # Validate depot postcode is selected
        if not self.depot_postcode_var.get().strip():
            messagebox.showwarning("Missing Home Base", 
                                  "Please select a Home Base Postcode first.\n\n"
                                  "Go to Configure menu and select a postcode from the dropdown.")
            return
        
        self.update_status("Processing...", "orange")
        self.progress_bar['value'] = 0
        
        # Run in separate thread to keep UI responsive
        thread = threading.Thread(target=self.run_clustering)
        thread.daemon = True
        thread.start()
        
    def run_clustering(self):
        """Run the TSP clustering analysis"""
        try:
            # Clear old region data to prevent stale data from previous clustering runs
            self.region_names = {}
            self.region_colors = {}
            
            # Get parameters
            desired_regions = int(self.num_regions_var.get())
            depot_postcode = self.depot_postcode_var.get().strip().upper()
            service_time = float(self.service_time_var.get())
            work_hours = float(self.work_hours_var.get())
            
            self.log("\n" + "="*80)
            self.log("TSP REGIONAL CLUSTERING ANALYSIS")
            self.log("="*80)
            self.log(f"\nConfiguration:")
            self.log(f"  Home base postcode: {depot_postcode}")
            self.log(f"  Desired regions: {desired_regions}")
            self.log(f"  Service time: {service_time} hours per customer")
            self.log(f"  Work hours: {work_hours} hours per day")
            self.log(f"  Using driving times from CSV file")
            
            # Load data
            self.log("\nLoading data...")
            self.progress_bar['value'] = 10
            
            # Load coordinates from distance_matrix.csv
            distance_matrix_file = os.path.join(self.output_dir, "distance_matrix.csv")
            locations_df = pd.read_csv(distance_matrix_file)
            self.log(f"✓ Loaded {len(locations_df)} locations with coordinates")
            
            # Load distances
            distances_df = pd.read_csv(self.distances_file)
            self.log(f"✓ Loaded {len(distances_df)} distance records")
            
            self.progress_bar['value'] = 20
            
            # Build driving time matrix
            self.log("\nBuilding driving time matrix...")
            postcodes = sorted(locations_df['postcode'].unique())
            n = len(postcodes)
            postcode_to_idx = {pc: i for i, pc in enumerate(postcodes)}
            
            # Initialize with infinity
            driving_time_matrix = np.full((n, n), np.inf)
            np.fill_diagonal(driving_time_matrix, 0)
            
            # Fill in known driving times
            for _, row in distances_df.iterrows():
                if row['origin'] in postcode_to_idx and row['destination'] in postcode_to_idx:
                    i = postcode_to_idx[row['origin']]
                    j = postcode_to_idx[row['destination']]
                    driving_time_matrix[i, j] = row['driving_time_minutes']
                    driving_time_matrix[j, i] = row['driving_time_minutes']  # Symmetric
            
            self.log(f"✓ Built {n}x{n} driving time matrix")
            self.progress_bar['value'] = 30
            
            # Find depot postcode in locations
            depot_row = locations_df[locations_df['postcode'].str.upper() == depot_postcode]
            
            if depot_row.empty:
                error_msg = f"Home base postcode '{depot_postcode}' not found in locations CSV!"
                self.log(f"\n✗ ERROR: {error_msg}")
                self.update_status("Error!", "red")
                messagebox.showerror("Invalid Home Base", error_msg)
                return
            
            depot_lat = depot_row.iloc[0]['latitude']
            depot_lon = depot_row.iloc[0]['longitude']
            depot = np.array([[depot_lat, depot_lon]])
            
            self.log(f"✓ Home base location: {depot_postcode} at ({depot_lat:.4f}, {depot_lon:.4f})")
            
            # Extract customer coordinates (excluding depot)
            customers_df = locations_df[locations_df['postcode'].str.upper() != depot_postcode].copy()
            customers_df = self._assign_location_instance_ids(customers_df)
            coords = customers_df[['latitude', 'longitude']].values
            self.log(f"✓ Clustering {len(coords)} customers (depot excluded)")
            
            # Use desired regions directly
            actual_regions = desired_regions
            
            self.log(f"\nUsing {actual_regions} regions")
            self.progress_bar['value'] = 35
            
            # Customer metadata (keep exact row order aligned with coords/labels)
            customer_postcodes = customers_df['postcode'].tolist()
            customer_location_ids = customers_df['location_id'].tolist()
            customer_location_instances = customers_df['location_instance'].astype(int).tolist()
            # Map customer postcode to driving-matrix index (duplicates intentionally share index)
            customer_postcode_to_idx = {pc: postcode_to_idx[pc] for pc in set(customer_postcodes) if pc in postcode_to_idx}
            
            # Run clustering
            self.log("\nPerforming clustering optimization...")
            labels, cluster_metrics = self.balance_clusters(
                coords, depot, driving_time_matrix, actual_regions
            )
            
            self.log(f"✓ Clustering complete")
            self.progress_bar['value'] = 60
            
            # Analyze clusters
            self.log("\nCluster Statistics:")
            self.log("="*60)
            for i in range(actual_regions):
                count = np.sum(labels == i)
                metric = cluster_metrics[i]
                self.log(f"Region {i+1}: {count} customers, Metric: {metric:.2f}")
            
            self.log(f"\nMean: {np.mean(cluster_metrics):.2f}")
            self.log(f"Std Dev: {np.std(cluster_metrics):.2f}")
            self.log(f"Balance ratio (max/min): {np.max(cluster_metrics)/np.min(cluster_metrics):.2f}")
            
            self.progress_bar['value'] = 70
            
            # Save results
            self.log("\nSaving results...")
            
            # Add cluster assignments to customer locations (depot separate)
            results_df = customers_df.copy()
            results_df['region'] = labels + 1
            results_df = results_df.sort_values('region')
            
            # Add depot as a separate entry with region 0
            depot_row_copy = depot_row.copy()
            depot_row_copy['region'] = 0
            results_df = pd.concat([depot_row_copy, results_df], ignore_index=True)
            
            # Create summary
            summary = []
            for i in range(actual_regions):
                region_postcodes = results_df[results_df['region'] == i+1]['postcode'].tolist()
                summary.append({
                    'region': i+1,
                    'customer_count': len(region_postcodes),
                    'postcodes': ', '.join(region_postcodes)
                })
            
            summary_df = pd.DataFrame(summary)
            
            # Store results for manual saving
            self.clustered_results = results_df
            self.summary_results = summary_df
            self.has_results = True
            
            # Store clustering data for editing and day calculation
            self.coords = coords
            self.labels = labels
            self.depot = depot
            self.n_clusters = actual_regions
            self.customer_postcodes = customer_postcodes
            self.customer_location_ids = customer_location_ids
            self.customer_location_instances = customer_location_instances
            self.customer_names = customers_df['client_name'].tolist() if 'client_name' in customers_df.columns else [None] * len(customer_postcodes)
            self.depot_postcode = depot_postcode
            self.driving_time_matrix = driving_time_matrix
            self.customer_postcode_to_idx = customer_postcode_to_idx
            self.depot_postcode_idx = postcode_to_idx[depot_postcode]
            
            # Enable edit and save buttons
            self.root.after(0, lambda: self.edit_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.save_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.rename_color_btn.config(state=tk.NORMAL))
            
            # Auto-assign default colors if not already set (starting from 1: Red)
            self.auto_assign_default_colors()
            
            self.log("\n✓ Results ready to save (use Run > Save Results to CSV)")
            
            self.progress_bar['value'] = 85
            
            # Create visualization
            self.log("\nGenerating visualization...")
            customer_names = self.customer_names
            self.root.after(0, lambda: self.create_visualization(coords, labels, depot, actual_regions, customer_postcodes, customer_names, depot_postcode))
            
            self.log(f"✓ Visualization displayed in GUI")
            
            self.progress_bar['value'] = 100
            self.update_status("Complete!", "green")
            
            self.log("\n" + "="*80)
            self.log("ANALYSIS COMPLETE")
            self.log("="*80)
            self.log(f"Total customers: {n}")
            self.log(f"Regions created: {actual_regions}")
            self.log(f"\nResults ready - use 'Run > Save Results to CSV' to save")
            self.log("="*80)
            
        except Exception as e:
            self.log(f"\n✗ ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.update_status("Error!", "red")
            messagebox.showerror("Error", f"An error occurred:\n{e}")
    
    def check_convex_hulls_overlap(self, coords, labels, n_clusters):
        """Check if any convex hulls of clusters overlap"""
        polygons = []
        
        for cluster_id in range(n_clusters):
            cluster_mask = labels == cluster_id
            cluster_points = coords[cluster_mask]
            
            if len(cluster_points) < 3:
                # Need at least 3 points for a polygon, use bounding circle approximation
                if len(cluster_points) == 1:
                    point = cluster_points[0]
                    radius = 1.0
                    circle_points = np.array([
                        [point[0] + radius * np.cos(theta), point[1] + radius * np.sin(theta)]
                        for theta in np.linspace(0, 2*np.pi, 8, endpoint=False)
                    ])
                    polygons.append(Polygon(circle_points))
                elif len(cluster_points) == 2:
                    # Create a thin rectangle around the two points
                    p1, p2 = cluster_points[0], cluster_points[1]
                    vec = p2 - p1
                    perp = np.array([-vec[1], vec[0]])
                    perp = perp / np.linalg.norm(perp) * 0.5
                    rect_points = np.array([p1 + perp, p2 + perp, p2 - perp, p1 - perp])
                    polygons.append(Polygon(rect_points))
            else:
                try:
                    hull = ConvexHull(cluster_points)
                    hull_points = cluster_points[hull.vertices]
                    polygons.append(Polygon(hull_points))
                except:
                    # If convex hull fails, use all points
                    polygons.append(Polygon(cluster_points))
        
        # Check all pairs for overlaps
        for i in range(len(polygons)):
            for j in range(i + 1, len(polygons)):
                if polygons[i].intersects(polygons[j]) and not polygons[i].touches(polygons[j]):
                    return True
        
        return False

    def _get_cluster_polygon(self, cluster_points):
        """Build a convex hull polygon for a cluster, or None if not possible."""
        if len(cluster_points) < 3:
            return None
        try:
            hull = ConvexHull(cluster_points)
            return Polygon(cluster_points[hull.vertices])
        except Exception:
            return None

    def _enforce_geographic_constraints(self, coords, labels, n_clusters, depot, min_size=3, max_iterations=200):
        """
        Post-process labels to reduce hull overlaps and keep depot outside region hulls.
        This is a lightweight local search and preserves minimum cluster size.
        """
        depot_point = Point(depot[0, 0], depot[0, 1])

        for _ in range(max_iterations):
            changed = False

            # Recompute centroids each pass
            centroids = {}
            for cid in range(n_clusters):
                mask = labels == cid
                if np.sum(mask) > 0:
                    centroids[cid] = coords[mask].mean(axis=0)
                else:
                    centroids[cid] = None

            # 1) Keep depot outside all cluster hulls
            for cid in range(n_clusters):
                indices = np.where(labels == cid)[0]
                if len(indices) <= min_size:
                    continue

                poly = self._get_cluster_polygon(coords[indices])
                if poly is None or not poly.contains(depot_point):
                    continue

                # Move one boundary-point candidate to nearest alternative cluster centroid
                best_move = None
                for idx in indices:
                    if len(indices) - 1 < min_size:
                        continue
                    for target in range(n_clusters):
                        if target == cid or centroids[target] is None:
                            continue

                        # Prefer moves that improve centroid fit
                        dist_to_own = np.linalg.norm(coords[idx] - centroids[cid])
                        dist_to_target = np.linalg.norm(coords[idx] - centroids[target])
                        score = dist_to_target - dist_to_own

                        if best_move is None or score < best_move[0]:
                            best_move = (score, idx, target)

                if best_move is not None:
                    _, move_idx, target_cluster = best_move
                    labels[move_idx] = target_cluster
                    changed = True
                    break

            if changed:
                continue

            # 2) Reduce overlapping hulls by moving closest boundary point from larger cluster
            cluster_polys = {}
            cluster_sizes = {}
            for cid in range(n_clusters):
                indices = np.where(labels == cid)[0]
                cluster_sizes[cid] = len(indices)
                cluster_polys[cid] = self._get_cluster_polygon(coords[indices])

            for i in range(n_clusters):
                for j in range(i + 1, n_clusters):
                    poly_i = cluster_polys[i]
                    poly_j = cluster_polys[j]

                    if poly_i is None or poly_j is None:
                        continue

                    if not (poly_i.intersects(poly_j) and not poly_i.touches(poly_j)):
                        continue

                    # Move from larger cluster to smaller cluster
                    src = i if cluster_sizes[i] >= cluster_sizes[j] else j
                    dst = j if src == i else i

                    src_indices = np.where(labels == src)[0]
                    if len(src_indices) <= min_size or centroids[dst] is None:
                        continue

                    # Candidate is point in source cluster closest to destination centroid
                    dists = [np.linalg.norm(coords[idx] - centroids[dst]) for idx in src_indices]
                    move_idx = src_indices[int(np.argmin(dists))]
                    labels[move_idx] = dst
                    changed = True
                    break
                if changed:
                    break

            if not changed:
                break

        return labels

    def _depot_inside_any_hull(self, coords, labels, n_clusters, depot):
        """Return True if depot lies strictly inside any cluster convex hull."""
        depot_point = Point(depot[0, 0], depot[0, 1])
        for cid in range(n_clusters):
            cluster_points = coords[labels == cid]
            poly = self._get_cluster_polygon(cluster_points)
            if poly is not None and poly.contains(depot_point):
                return True
        return False

    def _recompute_centroids(self, coords, labels, n_clusters):
        """Compute centroids for each cluster; fallback to global mean for empty clusters."""
        global_centroid = coords.mean(axis=0)
        centroids = np.zeros((n_clusters, coords.shape[1]))
        for cid in range(n_clusters):
            cluster_points = coords[labels == cid]
            if len(cluster_points) > 0:
                centroids[cid] = cluster_points.mean(axis=0)
            else:
                centroids[cid] = global_centroid
        return centroids

    def _repair_min_cluster_sizes(self, coords, labels, centroids, min_size, max_moves=5000):
        """Greedy repair so every cluster has at least min_size points."""
        moves = 0
        labels = labels.copy()

        while moves < max_moves:
            cluster_sizes = np.array([np.sum(labels == cid) for cid in range(len(centroids))])
            small_clusters = np.where(cluster_sizes < min_size)[0]
            if len(small_clusters) == 0:
                break

            moved_any = False
            for target in small_clusters:
                cluster_sizes = np.array([np.sum(labels == cid) for cid in range(len(centroids))])
                donor_clusters = np.where(cluster_sizes > min_size)[0]
                if len(donor_clusters) == 0:
                    continue

                best_candidate = None
                for donor in donor_clusters:
                    donor_indices = np.where(labels == donor)[0]
                    for idx in donor_indices:
                        own_dist = np.linalg.norm(coords[idx] - centroids[donor])
                        target_dist = np.linalg.norm(coords[idx] - centroids[target])
                        penalty = target_dist - own_dist
                        if best_candidate is None or penalty < best_candidate[0]:
                            best_candidate = (penalty, idx, donor, target)

                if best_candidate is not None:
                    _, idx, donor, target = best_candidate
                    labels[idx] = target
                    moves += 1
                    moved_any = True

            if not moved_any:
                break

        return labels

    def _strict_non_overlap_repartition(self, coords, labels, n_clusters, min_size, max_iterations=40):
        """
        Hard geometric fallback:
        1) Voronoi-style reassignment to nearest centroid (creates disjoint spatial partition)
        2) Minimum-size repair
        Repeated until stable.
        """
        labels = labels.copy()

        for _ in range(max_iterations):
            centroids = self._recompute_centroids(coords, labels, n_clusters)

            # Voronoi-style assignment (nearest centroid)
            dists = np.linalg.norm(coords[:, np.newaxis, :] - centroids[np.newaxis, :, :], axis=2)
            new_labels = np.argmin(dists, axis=1)

            # Enforce minimum cluster size
            new_labels = self._repair_min_cluster_sizes(coords, new_labels, centroids, min_size)

            if np.array_equal(new_labels, labels):
                break
            labels = new_labels

        return labels

    def _emergency_geometry_sanitize(self, coords, labels, n_clusters, depot, max_iterations=1000):
        """
        Last-resort sanitizer to guarantee geometric constraints without aborting run.
        Strategy:
        - If depot is inside a cluster hull, move one point out of that cluster.
        - If two hulls overlap, move one point from the smaller overlap-driving cluster.
        This may create small (1-2 point) clusters, which is acceptable for geometric validity.
        """
        labels = labels.copy()
        depot_point = Point(depot[0, 0], depot[0, 1])

        for _ in range(max_iterations):
            changed = False

            # Recompute centroids every pass
            centroids = self._recompute_centroids(coords, labels, n_clusters)

            # A) Remove depot from inside any hull
            for cid in range(n_clusters):
                cluster_indices = np.where(labels == cid)[0]
                if len(cluster_indices) < 3:
                    continue

                poly = self._get_cluster_polygon(coords[cluster_indices])
                if poly is None or not poly.contains(depot_point):
                    continue

                # Move point closest to another cluster centroid
                best_move = None
                for idx in cluster_indices:
                    for target in range(n_clusters):
                        if target == cid:
                            continue
                        dist_target = np.linalg.norm(coords[idx] - centroids[target])
                        dist_own = np.linalg.norm(coords[idx] - centroids[cid])
                        score = dist_target - dist_own
                        if best_move is None or score < best_move[0]:
                            best_move = (score, idx, target)

                if best_move is not None:
                    _, idx, target = best_move
                    labels[idx] = target
                    changed = True
                    break

            if changed:
                continue

            # B) Remove hull overlaps
            polys = {}
            sizes = {}
            for cid in range(n_clusters):
                cluster_indices = np.where(labels == cid)[0]
                sizes[cid] = len(cluster_indices)
                polys[cid] = self._get_cluster_polygon(coords[cluster_indices])

            for i in range(n_clusters):
                for j in range(i + 1, n_clusters):
                    pi = polys[i]
                    pj = polys[j]
                    if pi is None or pj is None:
                        continue
                    if not (pi.intersects(pj) and not pi.touches(pj)):
                        continue

                    # Move one point from smaller cluster (or i when equal) to the other
                    src = i if sizes[i] <= sizes[j] else j
                    dst = j if src == i else i

                    src_indices = np.where(labels == src)[0]
                    if len(src_indices) == 0:
                        continue

                    dists = [np.linalg.norm(coords[idx] - centroids[dst]) for idx in src_indices]
                    move_idx = src_indices[int(np.argmin(dists))]
                    labels[move_idx] = dst
                    changed = True
                    break
                if changed:
                    break

            if not changed:
                break

        return labels
    
    def balance_clusters(self, coords, depot, driving_time_matrix, n_clusters):
        """
        Spatial efficiency clustering without depot considerations
        Objective: Create compact, non-overlapping clusters that minimize total travel distance
        """
        min_size = 3  # Hard-coded minimum cluster size
        self.log("  Using spatial efficiency clustering (depot-independent)...")
        self.log(f"  Minimum cluster size: {min_size} (hard-coded)")
        
        # Calculate proximity threshold - customers closer than this MUST be in same cluster
        all_distances = []
        for i in range(len(coords)):
            for j in range(i+1, len(coords)):
                dist = np.linalg.norm(coords[i] - coords[j])
                all_distances.append(dist)
        
        proximity_threshold = np.percentile(all_distances, 10)  # Bottom 10% of distances
        self.log(f"  Proximity threshold: {proximity_threshold:.4f} (keeping nearest neighbors together)")
        
        # Use hierarchical clustering with Ward linkage (minimizes variance = compactness)
        # This creates the most spatially efficient clusters regardless of depot location
        self.log("  Running hierarchical clustering...")
        clustering = AgglomerativeClustering(
            n_clusters=n_clusters, 
            linkage='ward',  # Minimizes within-cluster variance
            metric='euclidean'
        )
        labels = clustering.fit_predict(coords)
        
        # Enforce proximity constraint: nearby points must be in same cluster
        self.log("  Enforcing proximity constraints...")
        max_proximity_iterations = 100
        for prox_iter in range(max_proximity_iterations):
            violations_fixed = 0
            
            for i in range(len(coords)):
                for j in range(i+1, len(coords)):
                    if labels[i] != labels[j]:
                        dist = np.linalg.norm(coords[i] - coords[j])
                        
                        if dist < proximity_threshold:
                            # These points are too close but in different clusters - merge them
                            cluster_i_size = np.sum(labels == labels[i])
                            cluster_j_size = np.sum(labels == labels[j])
                            
                            # Move from larger cluster to smaller (or merge smaller into larger)
                            if cluster_i_size > cluster_j_size and cluster_i_size > min_size:
                                labels[i] = labels[j]
                                violations_fixed += 1
                            elif cluster_j_size >= min_size:
                                labels[j] = labels[i]
                                violations_fixed += 1
            
            if violations_fixed == 0:
                break
        
        self.log(f"  ✓ Proximity constraints enforced after {prox_iter + 1} iterations")
        
        # Ensure minimum cluster sizes
        self.log("  Ensuring minimum cluster sizes...")
        for cluster_id in range(n_clusters):
            while np.sum(labels == cluster_id) < min_size:
                cluster_sizes = [np.sum(labels == i) for i in range(n_clusters)]
                largest_cluster = np.argmax(cluster_sizes)
                
                if cluster_sizes[largest_cluster] <= min_size:
                    break
                
                # Find customer in largest cluster closest to this cluster
                largest_mask = labels == largest_cluster
                cluster_mask = labels == cluster_id
                
                if np.sum(cluster_mask) > 0:
                    cluster_points = coords[cluster_mask]
                    cluster_centroid = cluster_points.mean(axis=0)
                else:
                    cluster_centroid = coords[np.random.choice(np.where(largest_mask)[0])]
                
                largest_indices = np.where(largest_mask)[0]
                distances_to_cluster = [np.linalg.norm(coords[idx] - cluster_centroid) for idx in largest_indices]
                closest_idx = largest_indices[np.argmin(distances_to_cluster)]
                
                labels[closest_idx] = cluster_id
        
        # Check for and fix overlaps while maintaining spatial compactness
        self.log("  Checking for region overlaps...")
        max_overlap_iterations = 100
        for overlap_iter in range(max_overlap_iterations):
            if overlap_iter % 25 == 0 and overlap_iter > 0:
                self.log(f"    Overlap check iteration {overlap_iter}/100...")
            
            if not self.check_convex_hulls_overlap(coords, labels, n_clusters):
                break
            
            # Find overlapping regions and move boundary points
            for i in range(n_clusters):
                for j in range(i + 1, n_clusters):
                    mask_i = labels == i
                    mask_j = labels == j
                    
                    if np.sum(mask_i) <= min_size or np.sum(mask_j) <= min_size:
                        continue
                    
                    points_i = coords[mask_i]
                    points_j = coords[mask_j]
                    
                    if len(points_i) >= 3 and len(points_j) >= 3:
                        try:
                            hull_i = ConvexHull(points_i)
                            hull_j = ConvexHull(points_j)
                            poly_i = Polygon(points_i[hull_i.vertices])
                            poly_j = Polygon(points_j[hull_j.vertices])
                            
                            if poly_i.intersects(poly_j) and not poly_i.touches(poly_j):
                                # Find boundary point in cluster i furthest from its centroid
                                centroid_i = points_i.mean(axis=0)
                                indices_i = np.where(mask_i)[0]
                                
                                # Get points on convex hull (boundary points)
                                hull_indices_i = indices_i[hull_i.vertices]
                                
                                # Find hull point closest to cluster j
                                centroid_j = points_j.mean(axis=0)
                                distances = [np.linalg.norm(coords[idx] - centroid_j) for idx in hull_indices_i]
                                closest_hull_idx = hull_indices_i[np.argmin(distances)]
                                
                                # Move it to cluster j
                                labels[closest_hull_idx] = j
                                break
                        except:
                            pass
        
        if overlap_iter < max_overlap_iterations - 1:
            self.log(f"  ✓ Overlaps resolved after {overlap_iter + 1} iterations")
        else:
            self.log(f"  ⚠ Some overlaps may remain after {max_overlap_iterations} iterations")
        
        # Additional compactness optimization - reduce cluster sprawl
        self.log("  Optimizing cluster compactness...")
        max_compactness_iterations = 50
        for compact_iter in range(max_compactness_iterations):
            improved = False
            
            for cluster_id in range(n_clusters):
                mask = labels == cluster_id
                cluster_size = np.sum(mask)
                
                if cluster_size <= min_size:
                    continue
                
                cluster_points = coords[mask]
                centroid = cluster_points.mean(axis=0)
                indices = np.where(mask)[0]
                
                # Find outlier points (furthest from centroid)
                distances = [np.linalg.norm(coords[idx] - centroid) for idx in indices]
                if len(distances) == 0:
                    continue
                    
                # Check if outlier would fit better in another cluster
                outlier_idx = indices[np.argmax(distances)]
                outlier_dist_from_own = max(distances)
                
                # Find nearest other cluster
                best_cluster = None
                min_dist_to_other = float('inf')
                
                for other_id in range(n_clusters):
                    if other_id == cluster_id:
                        continue
                    other_mask = labels == other_id
                    other_points = coords[other_mask]
                    if len(other_points) > 0:
                        other_centroid = other_points.mean(axis=0)
                        dist = np.linalg.norm(coords[outlier_idx] - other_centroid)
                        
                        # Move outlier if it's significantly closer to another cluster
                        if dist < min_dist_to_other and dist < outlier_dist_from_own * 0.8:
                            min_dist_to_other = dist
                            best_cluster = other_id
                
                if best_cluster is not None:
                    labels[outlier_idx] = best_cluster
                    improved = True
            
            if not improved:
                break
        
        self.log(f"  ✓ Compactness optimized after {compact_iter + 1} iterations")

        # Final geographic clean-up: reduce overlaps and keep depot outside region hulls
        self.log("  Enforcing non-overlap/depot-outside constraints...")
        labels = self._enforce_geographic_constraints(
            coords,
            labels,
            n_clusters,
            depot,
            min_size=min_size,
            max_iterations=200
        )

        # Hard fallback if constraints still violated
        has_overlap = self.check_convex_hulls_overlap(coords, labels, n_clusters)
        depot_inside = self._depot_inside_any_hull(coords, labels, n_clusters, depot)
        if has_overlap or depot_inside:
            self.log("  Constraints still violated - running strict geometric repartition...")
            labels = self._strict_non_overlap_repartition(
                coords,
                labels,
                n_clusters,
                min_size=min_size,
                max_iterations=40
            )
            labels = self._enforce_geographic_constraints(
                coords,
                labels,
                n_clusters,
                depot,
                min_size=min_size,
                max_iterations=400
            )

        has_overlap = self.check_convex_hulls_overlap(coords, labels, n_clusters)
        depot_inside = self._depot_inside_any_hull(coords, labels, n_clusters, depot)
        if has_overlap or depot_inside:
            self.log("  Running emergency geometry sanitizer...")
            labels = self._emergency_geometry_sanitize(coords, labels, n_clusters, depot, max_iterations=1000)
            has_overlap = self.check_convex_hulls_overlap(coords, labels, n_clusters)
            depot_inside = self._depot_inside_any_hull(coords, labels, n_clusters, depot)

            # Keep app running even in pathological geometry; this is now best-effort automatic recovery
            if has_overlap or depot_inside:
                self.log("  ⚠ Geometry sanitizer reached limit; constraints may be partially unresolved")
            else:
                self.log("  ✓ Emergency sanitizer resolved all geometric constraints")

        if not has_overlap and not depot_inside:
            self.log("  ✓ Final geometric constraints satisfied (no overlap, depot outside)")
        else:
            self.log("  ⚠ Final geometric constraints not fully satisfied")
        
        # Calculate final metrics - sum of intra-cluster distances
        self.log("  Calculating final metrics...")
        metrics = []
        total_intra_distance = 0
        
        for cluster_id in range(n_clusters):
            cluster_mask = labels == cluster_id
            cluster_indices = np.where(cluster_mask)[0]
            
            if len(cluster_indices) == 0:
                metrics.append(0)
                continue
            
            # Calculate sum of all pairwise distances within cluster
            cluster_distance_sum = 0
            for i in cluster_indices:
                for j in cluster_indices:
                    if i < j:
                        cluster_distance_sum += driving_time_matrix[i, j]
            
            metrics.append(cluster_distance_sum)
            total_intra_distance += cluster_distance_sum
        
        self.log(f"  Total intra-cluster distance: {total_intra_distance:.2f}")
        self.log(f"  Cluster sizes: {[np.sum(labels == i) for i in range(n_clusters)]}")
        
        return labels, metrics
    
    def calculate_minimum_days_for_region(self, region_num):
        """Calculate minimum days needed to service all customers in a region
        Returns technical minimum + 1 day buffer"""
        if not hasattr(self, 'driving_time_matrix') or not hasattr(self, 'customer_postcode_to_idx'):
            # Return fallback - 1 day per 5 customers as rough estimate
            if hasattr(self, 'labels'):
                region_mask = self.labels == (region_num - 1)
                customer_count = np.sum(region_mask)
                return max(1, int(np.ceil(customer_count / 5.0)))
            return 1
        
        # Get configuration parameters
        try:
            service_time_hours = float(self.service_time_var.get())
            work_hours = float(self.work_hours_var.get())
        except:
            service_time_hours = 1.0
            work_hours = 8.0
        
        # Get customers in this region
        region_mask = self.labels == (region_num - 1)  # labels are 0-indexed
        region_customer_indices = np.where(region_mask)[0]
        
        if len(region_customer_indices) == 0:
            return 1
        
        # Map to driving matrix indices
        matrix_indices = []
        for customer_idx in region_customer_indices:
            postcode = self.customer_postcodes[customer_idx]
            if postcode in self.customer_postcode_to_idx:
                matrix_indices.append(self.customer_postcode_to_idx[postcode])
        
        if len(matrix_indices) == 0:
            return 1
        
        # Use nearest neighbor to get approximate tour through all customers
        # Start from depot, visit all customers, return to depot
        tour_time_minutes = 0
        
        # Travel from depot to nearest customer
        min_depot_distance = np.inf
        for idx in matrix_indices:
            dist = self.driving_time_matrix[self.depot_postcode_idx, idx]
            if not np.isinf(dist) and dist < min_depot_distance:
                min_depot_distance = dist
        
        # If no valid distances found, use a default estimate
        if np.isinf(min_depot_distance):
            # Estimate based on number of customers (rough fallback)
            customers_count = len(matrix_indices)
            service_time_minutes = service_time_hours * 60 * customers_count
            total_time_hours = service_time_minutes / 60
            technical_minimum = np.ceil(total_time_hours / work_hours)
            return int(technical_minimum + 1)
        
        tour_time_minutes += min_depot_distance
        
        # Travel between customers using nearest neighbor approximation
        # For a more accurate estimate of total tour time
        if len(matrix_indices) > 1:
            # Calculate average pairwise distance within region (excluding infinities)
            total_intra_distance = 0
            count = 0
            for i in matrix_indices:
                for j in matrix_indices:
                    if i != j:
                        dist = self.driving_time_matrix[i, j]
                        if not np.isinf(dist):
                            total_intra_distance += dist
                            count += 1
            
            if count > 0:
                avg_distance = total_intra_distance / count
                # Estimate tour time as number of hops * average distance
                # (n-1 hops between n customers)
                tour_time_minutes += avg_distance * (len(matrix_indices) - 1)
            else:
                # Fallback: no valid inter-customer distances
                tour_time_minutes += min_depot_distance * len(matrix_indices)
        
        # Travel from last customer back to depot (approximate same as first leg)
        tour_time_minutes += min_depot_distance
        
        # Add service time for all customers
        service_time_minutes = service_time_hours * 60 * len(matrix_indices)
        
        # Total time needed
        total_time_minutes = tour_time_minutes + service_time_minutes
        total_time_hours = total_time_minutes / 60
        
        # Calculate minimum days (round up)
        technical_minimum = np.ceil(total_time_hours / work_hours)
        
        # Add 1 day buffer as requested
        minimum_days = int(technical_minimum + 1)
        
        return minimum_days
    
    def _add_region_labels_with_overlap_prevention(self, ax, coords, customer_postcodes, depot):
        """Add region labels with simple overlap prevention"""
        # Collect all existing label positions (customers + depot)
        occupied_positions = []
        
        # Add customer positions
        for coord in coords:
            occupied_positions.append((coord[1], coord[0]))
        
        # Add depot position
        occupied_positions.append((depot[0, 1], depot[0, 0]))        # Convert to numpy array for easier distance calculations
        occupied_positions = np.array(occupied_positions)
        
        # Get plot limits to constrain label positions
        xlim = ax.get_xlim()
        ylim = ax.get_ylim()
        
        # Process each region label
        for region_info in self._region_labels_to_draw:
            best_lon = region_info['lon']
            best_lat = region_info['lat']
            
            # Check if centroid position overlaps with existing labels
            centroid_pos = np.array([best_lon, best_lat])
            distances = np.sqrt(np.sum((occupied_positions - centroid_pos)**2, axis=1))
            
            # If too close to any existing label, try to find a better position
            min_distance_threshold = 0.005  # Adjust based on your coordinate scale
            if np.min(distances) < min_distance_threshold:
                # Try offsetting in various directions
                offsets = [
                    (0.01, 0.01), (-0.01, 0.01), (0.01, -0.01), (-0.01, -0.01),
                    (0.015, 0), (-0.015, 0), (0, 0.015), (0, -0.015),
                    (0.02, 0.01), (-0.02, -0.01), (0.01, 0.02), (-0.01, -0.02)
                ]
                
                best_distance = np.min(distances)
                for offset_lon, offset_lat in offsets:
                    test_lon = region_info['lon'] + offset_lon
                    test_lat = region_info['lat'] + offset_lat
                    
                    # Check if within plot bounds
                    if xlim[0] <= test_lon <= xlim[1] and ylim[0] <= test_lat <= ylim[1]:
                        test_pos = np.array([test_lon, test_lat])
                        test_distances = np.sqrt(np.sum((occupied_positions - test_pos)**2, axis=1))
                        min_test_dist = np.min(test_distances)
                        
                        if min_test_dist > best_distance:
                            best_distance = min_test_dist
                            best_lon = test_lon
                            best_lat = test_lat
            
            # Draw the region label at the best position found
            ax.annotate(region_info['text'],
                       xy=(best_lon, best_lat),
                       fontsize=7,  # Same size as location labels
                       fontweight='bold',
                       color='white',
                       ha='center',
                       va='center',
                       bbox=dict(boxstyle='round,pad=0.4', 
                               facecolor=region_info['color'], 
                               edgecolor='black', 
                               alpha=0.9,
                               linewidth=2),
                       zorder=20)
            
            # Add this position to occupied positions for next iteration
            occupied_positions = np.vstack([occupied_positions, [best_lon, best_lat]])
        
        # Clear the stored labels
        self._region_labels_to_draw = []
            
    def create_visualization(self, coords, labels, depot, n_clusters, customer_postcodes, customer_names, depot_postcode):
        """Create visualization of clusters with postcode labels embedded in GUI"""
        # Initialize region labels list
        self._region_labels_to_draw = []
        
        # Clear any existing content in viz frame
        for widget in self.viz_frame.winfo_children():
            widget.destroy()
        
        # Ensure viz frame is visible
        self.viz_frame.grid_rowconfigure(0, weight=1)
        self.viz_frame.grid_columnconfigure(0, weight=1)
        
        # Create container for canvas
        self.viz_canvas_container = ttk.Frame(self.viz_frame)
        self.viz_canvas_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create figure (constrained layout prevents axis labels from being clipped)
        fig = Figure(figsize=(12, 8), dpi=100, constrained_layout=True)
        ax = fig.add_subplot(111)
        
        # Build color list from region colors (use Outlook colors if available)
        colors = []
        for i in range(n_clusters):
            region_num = i + 1
            color_code = self.region_colors.get(region_num, 1)  # Default to Red
            matplotlib_color = self.outlook_color_to_matplotlib(color_code)
            colors.append(matplotlib_color)
        
        # Plot clusters
        for i in range(n_clusters):
            cluster_mask = labels == i
            cluster_coords = coords[cluster_mask]
            
            if len(cluster_coords) > 0:
                # Get custom name if available
                region_name = self.get_region_display_name(i + 1)
                
                ax.scatter(cluster_coords[:, 1], cluster_coords[:, 0], 
                          c=colors[i], s=100, alpha=0.6, 
                          edgecolors='black', linewidth=1,
                          label=f'{region_name} ({np.sum(cluster_mask)} locations)')
                
                # Draw convex hull if possible
                if len(cluster_coords) >= 3:
                    try:
                        hull = ConvexHull(cluster_coords)
                        for simplex in hull.simplices:
                            ax.plot(cluster_coords[simplex, 1], cluster_coords[simplex, 0], 
                                   colors[i], linewidth=2, alpha=0.5)
                    except:
                        pass
                
                # Add region name label at centroid (store for overlap prevention)
                centroid_lon = cluster_coords[:, 1].mean()
                centroid_lat = cluster_coords[:, 0].mean()
                
                # Store region label info for later (after all customer labels)
                self._region_labels_to_draw.append({
                    'text': region_name,
                    'lon': centroid_lon,
                    'lat': centroid_lat,
                    'color': colors[i]
                })
        
        # Plot excluded locations (region = -1)
        excluded_mask = labels == -1
        excluded_coords = coords[excluded_mask]
        if len(excluded_coords) > 0:
            ax.scatter(excluded_coords[:, 1], excluded_coords[:, 0], 
                      c='lightgray', s=150, alpha=0.6, 
                      edgecolors='red', linewidth=2,
                      marker='D',
                      label=f'Excluded ({np.sum(excluded_mask)} locations)')
        
        # Add postcode labels for customer locations
        postcode_total_counts = {}
        postcode_seen_counts = {}
        for pc in customer_postcodes:
            postcode_total_counts[pc] = postcode_total_counts.get(pc, 0) + 1

        # For points with identical coordinates, spread label offsets radially
        coord_keys = [(round(float(c[0]), 6), round(float(c[1]), 6)) for c in coords]
        coord_total_counts = {}
        coord_seen_counts = {}
        for key in coord_keys:
            coord_total_counts[key] = coord_total_counts.get(key, 0) + 1

        show_names = get_show_names()
        for idx, (coord, postcode) in enumerate(zip(coords, customer_postcodes)):
            postcode_seen_counts[postcode] = postcode_seen_counts.get(postcode, 0) + 1
            postcode_instance = postcode_seen_counts[postcode]
            postcode_total = postcode_total_counts[postcode]
            postcode_display = postcode if postcode_total == 1 else f"{postcode} ({postcode_instance}/{postcode_total})"

            # Determine what to display
            customer_name = customer_names[idx] if idx < len(customer_names) else None
            if show_names and customer_name:
                display_text = customer_name if postcode_total == 1 else f"{customer_name} [{postcode_display}]"
            else:
                display_text = postcode_display

            # Dynamic annotation offset for exact coordinate duplicates
            key = coord_keys[idx]
            coord_seen_counts[key] = coord_seen_counts.get(key, 0) + 1
            coord_instance = coord_seen_counts[key]
            coord_total = coord_total_counts[key]
            if coord_total > 1:
                angle = 2 * np.pi * (coord_instance - 1) / coord_total
                radius = 8
                offset_x = int(round(radius * np.cos(angle)))
                offset_y = int(round(radius * np.sin(angle)))
            else:
                offset_x, offset_y = 3, 3
            
            # Different styling for excluded postcodes
            if labels[idx] == -1:
                bbox_style = dict(boxstyle='round,pad=0.3', facecolor='lightgray', 
                                edgecolor='red', alpha=0.7, linestyle='--', linewidth=1.5)
            else:
                bbox_style = dict(boxstyle='round,pad=0.3', facecolor='white', 
                                edgecolor='gray', alpha=0.7)
            
            ax.annotate(display_text, 
                       xy=(coord[1], coord[0]),
                       xytext=(offset_x, offset_y),
                       textcoords='offset points',
                       fontsize=7,
                       fontweight='bold',
                       bbox=bbox_style,
                       zorder=10)
        
        # Plot depot
        ax.scatter(depot[0, 1], depot[0, 0], c='gold', s=500, marker='*', 
                  edgecolors='black', linewidth=2, label='Home Base (Depot)', zorder=5)
        
        # Add depot label
        ax.annotate(depot_postcode, 
                   xy=(depot[0, 1], depot[0, 0]),
                   xytext=(5, 5),  # Offset by 5 points
                   textcoords='offset points',
                   fontsize=9,
                   fontweight='bold',
                   color='darkgoldenrod',
                   bbox=dict(boxstyle='round,pad=0.4', facecolor='yellow', 
                            edgecolor='black', alpha=0.8, linewidth=2),
                   zorder=15)
        
        # Now add region labels with overlap prevention
        self._add_region_labels_with_overlap_prevention(ax, coords, customer_postcodes, depot)
        
        ax.set_xlabel('Longitude', fontsize=12, fontweight='bold')
        ax.set_ylabel('Latitude', fontsize=12, fontweight='bold')
        ax.set_title(f'Regional Clustering: {n_clusters} Regions, {len(coords)} Locations\nHome Base at center', 
                    fontsize=14, fontweight='bold')
        ax.legend(loc='best', fontsize=9)
        ax.grid(True, alpha=0.3)
        
        # Embed in tkinter with dedicated frames so toolbar stays visible on resize
        self.viz_canvas_container.grid_rowconfigure(0, weight=1)
        self.viz_canvas_container.grid_rowconfigure(1, weight=0)
        self.viz_canvas_container.grid_columnconfigure(0, weight=1)

        plot_frame = ttk.Frame(self.viz_canvas_container)
        plot_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        toolbar_frame = ttk.Frame(self.viz_canvas_container)
        toolbar_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))

        canvas = FigureCanvasTkAgg(fig, master=plot_frame)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(fill=tk.BOTH, expand=True)
        canvas.draw()

        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame, pack_toolbar=False)
        toolbar.update()
        toolbar.pack(fill=tk.X)
        
        self.canvas = canvas
        self.toolbar = toolbar
    
    def create_initial_visualization(self, coords, depot, postcodes, depot_postcode):
        """Create initial visualization showing all locations before clustering"""
        # Clear any existing content in viz frame
        for widget in self.viz_frame.winfo_children():
            widget.destroy()
        
        # Ensure viz frame is visible
        self.viz_frame.grid_rowconfigure(0, weight=1)
        self.viz_frame.grid_columnconfigure(0, weight=1)
        
        # Create container for canvas
        self.viz_canvas_container = ttk.Frame(self.viz_frame)
        self.viz_canvas_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create figure (constrained layout prevents axis labels from being clipped)
        fig = Figure(figsize=(12, 8), dpi=100, constrained_layout=True)
        ax = fig.add_subplot(111)
        
        # Plot all locations (not yet clustered)
        ax.scatter(coords[:, 1], coords[:, 0], 
                  c='lightblue', s=100, alpha=0.6, 
                  edgecolors='black', linewidth=1,
                  label=f'{len(coords)} Locations (Unclustered)')
        
        # Add postcode labels for all locations
        postcode_total_counts = {}
        postcode_seen_counts = {}
        for pc in postcodes:
            postcode_total_counts[pc] = postcode_total_counts.get(pc, 0) + 1

        coord_keys = [(round(float(c[0]), 6), round(float(c[1]), 6)) for c in coords]
        coord_total_counts = {}
        coord_seen_counts = {}
        for key in coord_keys:
            coord_total_counts[key] = coord_total_counts.get(key, 0) + 1

        for idx, (coord, postcode) in enumerate(zip(coords, postcodes)):
            # Skip depot in this loop (will be added separately)
            if postcode.upper() == depot_postcode.upper():
                continue

            postcode_seen_counts[postcode] = postcode_seen_counts.get(postcode, 0) + 1
            instance = postcode_seen_counts[postcode]
            total = postcode_total_counts[postcode]
            label_text = postcode if total == 1 else f"{postcode} ({instance}/{total})"

            key = coord_keys[idx]
            coord_seen_counts[key] = coord_seen_counts.get(key, 0) + 1
            coord_instance = coord_seen_counts[key]
            coord_total = coord_total_counts[key]
            if coord_total > 1:
                angle = 2 * np.pi * (coord_instance - 1) / coord_total
                radius = 8
                offset_x = int(round(radius * np.cos(angle)))
                offset_y = int(round(radius * np.sin(angle)))
            else:
                offset_x, offset_y = 3, 3

            ax.annotate(label_text, 
                       xy=(coord[1], coord[0]),
                       xytext=(offset_x, offset_y),
                       textcoords='offset points',
                       fontsize=7,
                       fontweight='bold',
                       bbox=dict(boxstyle='round,pad=0.3', facecolor='white', 
                                edgecolor='gray', alpha=0.7),
                       zorder=10)
        
        # Plot depot
        ax.scatter(depot[0, 1], depot[0, 0], c='gold', s=500, marker='*', 
                  edgecolors='black', linewidth=2, label='Home Base (Depot)', zorder=5)
        
        # Add depot label
        ax.annotate(depot_postcode, 
                   xy=(depot[0, 1], depot[0, 0]),
                   xytext=(5, 5),
                   textcoords='offset points',
                   fontsize=9,
                   fontweight='bold',
                   color='darkgoldenrod',
                   bbox=dict(boxstyle='round,pad=0.4', facecolor='yellow', 
                            edgecolor='black', alpha=0.8, linewidth=2),
                   zorder=15)
        
        ax.set_xlabel('Longitude', fontsize=12, fontweight='bold')
        ax.set_ylabel('Latitude', fontsize=12, fontweight='bold')
        ax.set_title(f'All Locations: {len(coords)} Points\\nReady for Clustering - Configure parameters and Run', 
                    fontsize=14, fontweight='bold')
        ax.legend(loc='best', fontsize=10)
        ax.grid(True, alpha=0.3)
        
        # Embed in tkinter with dedicated frames so toolbar stays visible on resize
        self.viz_canvas_container.grid_rowconfigure(0, weight=1)
        self.viz_canvas_container.grid_rowconfigure(1, weight=0)
        self.viz_canvas_container.grid_columnconfigure(0, weight=1)

        plot_frame = ttk.Frame(self.viz_canvas_container)
        plot_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        toolbar_frame = ttk.Frame(self.viz_canvas_container)
        toolbar_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))

        canvas = FigureCanvasTkAgg(fig, master=plot_frame)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(fill=tk.BOTH, expand=True)
        canvas.draw()

        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame, pack_toolbar=False)
        toolbar.update()
        toolbar.pack(fill=tk.X)
        
        self.canvas = canvas
        self.toolbar = toolbar
    
    def show_edit_regions_dialog(self):
        """Show dialog with table + region dropdowns for all locations"""
        if not self.has_results:
            messagebox.showwarning("No Results", 
                                  "No clustering results available.\n\n"
                                  "Run clustering analysis first.")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Location Regions")
        dialog.geometry("1100x720")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Edit Location Regions", font=('Arial', 14, 'bold')).pack(anchor=tk.W, pady=(0, 6))
        instructions = ttk.Label(
            main_frame,
            text="Use the Region dropdown on each row. Choose 'Create New Region...' to add any region number.",
            font=('Arial', 9),
            foreground='gray'
        )
        instructions.pack(anchor=tk.W, pady=(0, 10))
        
        # Ensure unique location IDs exist
        self.clustered_results = self._assign_location_instance_ids(self.clustered_results)
        editable_df = self.clustered_results[self.clustered_results['region'] != 0].copy()
        editable_df = editable_df.sort_values(['region', 'postcode'], kind='mergesort')
        
        # Build dynamic region list from current data + known cluster count
        existing_regions = set()
        for region_val in editable_df['region'].tolist():
            if pd.notna(region_val):
                region_int = int(region_val)
                if region_int > 0:
                    existing_regions.add(region_int)
        for i in range(1, int(self.n_clusters) + 1):
            existing_regions.add(i)
        
        def to_region_label(region_value):
            return "Excluded" if int(region_value) == -1 else f"Region {int(region_value)}"
        
        def region_options():
            ordered = [f"Region {r}" for r in sorted(existing_regions)]
            return ordered + ["Excluded", "Create New Region..."]
        
        # Header row
        header = ttk.Frame(main_frame)
        header.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(header, text="Postcode", width=16, anchor='w', font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(header, text="Client Name", width=34, anchor='w', font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(header, text="Region", width=22, anchor='w', font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(header, text="Color", width=18, anchor='w', font=('Arial', 9, 'bold')).pack(side=tk.LEFT)
        
        # Scrollable rows area
        table_container = ttk.Frame(main_frame)
        table_container.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(table_container, highlightthickness=0)
        y_scroll = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=canvas.yview)
        rows_frame = ttk.Frame(canvas)
        
        rows_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        window_id = canvas.create_window((0, 0), window=rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=y_scroll.set)
        
        def resize_rows_frame(event):
            canvas.itemconfigure(window_id, width=event.width)
        
        canvas.bind("<Configure>", resize_rows_frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        row_widgets = {}
        region_vars = {}
        region_combos = []
        color_widgets = {}

        def get_region_color_info(region_num):
            """Return (hex_color, color_name) for a region number."""
            if region_num <= 0:
                return '#D3D3D3', 'Excluded'
            color_code = int(self.region_colors.get(region_num, ((region_num - 1) % 24) + 1))
            if region_num not in self.region_colors:
                self.region_colors[region_num] = color_code
            hex_color = self.outlook_color_to_matplotlib(color_code)
            color_name = OUTLOOK_COLORS.get(color_code, 'Red')
            return hex_color, color_name

        def parse_region_label(label):
            if label == "Excluded":
                return -1
            if label.startswith("Region "):
                try:
                    return int(label.split(" ", 1)[1])
                except (ValueError, IndexError):
                    return None
            return None

        def update_row_color(location_id):
            selected = region_vars[location_id].get().strip()
            region_num = parse_region_label(selected)
            if region_num is None:
                return
            hex_color, color_name = get_region_color_info(region_num)
            swatch = color_widgets[location_id]['swatch']
            color_label = color_widgets[location_id]['label']
            swatch.configure(bg=hex_color)
            color_label.configure(text=color_name)
        
        def refresh_region_dropdowns():
            opts = region_options()
            for combo in region_combos:
                combo['values'] = opts
        
        def handle_region_pick(location_id, event=None):
            picked = region_vars[location_id].get().strip()
            if picked != "Create New Region...":
                update_row_color(location_id)
                return
            
            suggested = max(existing_regions) + 1 if existing_regions else max(1, int(self.n_clusters) + 1)
            new_region = simpledialog.askinteger(
                "Create New Region",
                "Enter new region number (any positive integer):",
                parent=dialog,
                minvalue=1,
                initialvalue=suggested
            )
            
            if new_region is None:
                current_region = int(self.clustered_results[self.clustered_results['location_id'] == location_id]['region'].iloc[0])
                region_vars[location_id].set(to_region_label(current_region))
                update_row_color(location_id)
                return
            
            existing_regions.add(int(new_region))
            if int(new_region) > int(self.n_clusters):
                self.n_clusters = int(new_region)
            if int(new_region) not in self.region_colors:
                self.region_colors[int(new_region)] = ((int(new_region) - 1) % 24) + 1
            refresh_region_dropdowns()
            region_vars[location_id].set(f"Region {int(new_region)}")
            update_row_color(location_id)
            self.log(f"\n✨ Created new Region {int(new_region)}")
        
        # Build rows
        for _, row in editable_df.iterrows():
            location_id = str(row['location_id'])
            postcode = str(row['postcode'])
            client_name = str(row['client_name']).strip() if 'client_name' in row and pd.notna(row['client_name']) else ""
            region = int(row['region'])
            
            row_frame = ttk.Frame(rows_frame)
            row_frame.pack(fill=tk.X, pady=1)
            
            ttk.Label(row_frame, text=postcode, width=16, anchor='w').pack(side=tk.LEFT, padx=(0, 6))
            ttk.Label(row_frame, text=client_name, width=34, anchor='w').pack(side=tk.LEFT, padx=(0, 6))
            
            region_var = tk.StringVar(value=to_region_label(region))
            region_vars[location_id] = region_var
            region_combo = ttk.Combobox(
                row_frame,
                textvariable=region_var,
                values=region_options(),
                state='readonly',
                width=20
            )
            region_combo.pack(side=tk.LEFT, padx=(0, 6))
            region_combo.bind('<<ComboboxSelected>>', lambda e, lid=location_id: handle_region_pick(lid, e))
            region_combos.append(region_combo)
            
            color_frame = ttk.Frame(row_frame)
            color_frame.pack(side=tk.LEFT, padx=(0, 6))

            swatch = tk.Label(color_frame, text="", width=3, height=1, relief=tk.SOLID, bd=1)
            swatch.pack(side=tk.LEFT, padx=(0, 6), pady=1)
            color_name_label = ttk.Label(color_frame, text="", width=12, anchor='w')
            color_name_label.pack(side=tk.LEFT)

            color_widgets[location_id] = {
                'swatch': swatch,
                'label': color_name_label,
            }
            update_row_color(location_id)
            
            row_widgets[location_id] = {
                'postcode': postcode,
                'client_name': client_name,
            }
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        def apply_changes():
            changes = 0
            for location_id, region_var in region_vars.items():
                selected = region_var.get().strip()
                if not selected:
                    continue
                
                if selected == "Excluded":
                    new_region = -1
                elif selected.startswith("Region "):
                    try:
                        new_region = int(selected.split(" ", 1)[1])
                    except (ValueError, IndexError):
                        continue
                else:
                    continue
                
                old_region = int(self.clustered_results[
                    self.clustered_results['location_id'] == location_id
                ]['region'].iloc[0])
                
                if old_region == new_region:
                    continue
                
                self.clustered_results.loc[
                    self.clustered_results['location_id'] == location_id, 'region'
                ] = new_region
                
                if hasattr(self, 'customer_location_ids'):
                    try:
                        location_idx = [str(x) for x in self.customer_location_ids].index(str(location_id))
                        self.labels[location_idx] = -1 if new_region == -1 else (new_region - 1)
                    except ValueError:
                        pass
                
                row_data = row_widgets.get(location_id, {})
                location_display = row_data.get('postcode', '')
                if row_data.get('client_name'):
                    location_display = f"{location_display} - {row_data['client_name']}"
                
                old_display = "Excluded" if old_region == -1 else f"Region {old_region}"
                new_display = "Excluded" if new_region == -1 else f"Region {new_region}"
                self.log(f"\n✏️ Manual Edit: {location_display} moved from {old_display} to {new_display}")
                changes += 1
            
            self.update_summary_results()
            self.refresh_visualization()
            message = "No region changes were detected." if changes == 0 else f"Applied {changes} region change(s)."
            messagebox.showinfo("Update Complete", message)
        
        ttk.Button(button_frame, text="Apply Changes & Refresh", command=apply_changes, width=25).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Close", command=dialog.destroy, width=15).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(
            main_frame,
            text="Tip: use each row's dropdown for region selection. Choose 'Create New Region...' to add any number.",
            font=('Arial', 8),
            foreground='gray'
        ).pack(anchor=tk.W, pady=(8, 0))
    
    def update_summary_results(self):
        """Update the summary results after manual edits"""
        summary = []
        for i in range(self.n_clusters):
            region_postcodes = self.clustered_results[
                self.clustered_results['region'] == i+1
            ]['postcode'].tolist()
            summary.append({
                'region': i+1,
                'customer_count': len(region_postcodes),
                'postcodes': ', '.join(region_postcodes)
            })
        
        # Add excluded locations if any
        excluded_postcodes = self.clustered_results[
            self.clustered_results['region'] == -1
        ]['postcode'].tolist()
        
        if excluded_postcodes:
            summary.append({
                'region': 'Excluded',
                'customer_count': len(excluded_postcodes),
                'postcodes': ', '.join(excluded_postcodes)
            })
        
        self.summary_results = pd.DataFrame(summary)
    
    def refresh_visualization(self):
        """Refresh the visualization after manual edits"""
        # Create updated visualization with modified labels
        customer_names = getattr(self, 'customer_names', [None] * len(self.customer_postcodes))
        self.create_visualization(
            self.coords, 
            self.labels, 
            self.depot, 
            self.n_clusters, 
            self.customer_postcodes, 
            customer_names,
            self.depot_postcode
        )
        self.log("✓ Visualization refreshed with manual edits")
    
    def show_rename_recolor_dialog(self):
        """Show combined dialog for renaming and recoloring regions"""
        if not self.has_results:
            messagebox.showwarning("No Results", 
                                  "No clustering results available.\n\n"
                                  "Run clustering analysis first.")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Rename and Recolor Regions")
        dialog.geometry("750x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Rename and Recolor Regions", font=('Arial', 14, 'bold')).pack(pady=(0, 10))
        
        ttk.Label(frame, text="Give your regions custom names and assign colors for calendar scheduling",
                 font=('Arial', 9), foreground='gray').pack(pady=(0, 20))
        
        # Create scrollable frame for region entries
        content_frame = ttk.Frame(frame)
        content_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(content_frame, height=400)
        scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Store entry widgets
        region_entries = {}
        color_combos = {}
        
        # Create color options list
        color_options = [f"{idx}: {name}" for idx, name in OUTLOOK_COLORS.items() if idx > 0]
        
        # Add entry for each region
        for i in range(self.n_clusters):
            region_num = i + 1
            region_frame = ttk.Frame(scrollable_frame)
            region_frame.pack(fill=tk.X, pady=8, padx=10)
            
            # Get current name and color
            current_name = self.region_names.get(region_num, f"Region {region_num}")
            current_color = self.region_colors.get(region_num, 1)  # Default to Red
            
            # Get customer count
            customer_count = len(self.clustered_results[self.clustered_results['region'] == region_num])
            
            # Region number label
            ttk.Label(region_frame, text=f"Region {region_num}:", 
                     font=('Arial', 10, 'bold'), width=10).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
            
            # Name entry
            ttk.Label(region_frame, text="Name:", font=('Arial', 9)).grid(row=0, column=1, sticky=tk.W, padx=(0, 5))
            entry_var = tk.StringVar(value=current_name)
            entry = ttk.Entry(region_frame, textvariable=entry_var, width=25)
            entry.grid(row=0, column=2, sticky=tk.W, padx=(0, 20))
            region_entries[region_num] = entry_var
            
            # Color dropdown
            ttk.Label(region_frame, text="Color:", font=('Arial', 9)).grid(row=0, column=3, sticky=tk.W, padx=(0, 5))
            color_var = tk.StringVar(value=f"{current_color}: {OUTLOOK_COLORS[current_color]}")
            combo = ttk.Combobox(region_frame, textvariable=color_var, 
                               values=color_options, state='readonly', width=20)
            combo.grid(row=0, column=4, sticky=tk.W, padx=(0, 10))
            color_combos[region_num] = color_var
            
            # Customer count
            ttk.Label(region_frame, text=f"({customer_count} customers)", 
                     font=('Arial', 8), foreground='gray').grid(row=0, column=5, sticky=tk.W)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Buttons
        def apply_changes():
            # Apply names
            for region_num, entry_var in region_entries.items():
                new_name = entry_var.get().strip()
                if new_name:
                    self.region_names[region_num] = new_name
                else:
                    # Revert to default if empty
                    if region_num in self.region_names:
                        del self.region_names[region_num]
            
            # Apply colors
            for region_num, color_var in color_combos.items():
                color_str = color_var.get()
                # Parse color index from "1: Red" format
                color_index = int(color_str.split(':')[0])
                self.region_colors[region_num] = color_index
            
            self.save_region_names()
            self.refresh_visualization()
            self.log(f"\n✓ Region names and colors updated")
            messagebox.showinfo("Success", 
                              f"Region names and colors have been updated!\n\n"
                              f"These settings will be used in the Calendar Organizer.")
            dialog.destroy()
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=15)
        
        ttk.Button(btn_frame, text="Apply Changes", command=apply_changes, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=15).pack(side=tk.LEFT, padx=5)
    
    def save_region_names(self):
        """Save region names and colors to CSV"""
        # Use the unified save_region_colors method which saves both names and colors
        self.save_region_colors()
    
    def load_region_names(self):
        """Load region names and colors from CSV"""
        if not self.output_dir:
            return
        
        names_file = os.path.join(self.output_dir, "region_names.csv")
        if not os.path.exists(names_file):
            return
        
        try:
            df = pd.read_csv(names_file)
            self.region_names = {}
            self.region_colors = {}
            
            for _, row in df.iterrows():
                region_num = int(row['region'])
                self.region_names[region_num] = row['name']
                
                # Load color code if available
                if 'color_code' in df.columns:
                    self.region_colors[region_num] = int(row['color_code'])
            
            self.log(f"✓ Loaded {len(self.region_names)} region names and colors")
        except Exception as e:
            self.log(f"⚠ Failed to load region names: {e}")
    
    def get_region_display_name(self, region_num):
        """Get display name for a region (custom name or default)"""
        return self.region_names.get(region_num, f"Region {region_num}")
    
    def outlook_color_to_matplotlib(self, color_code):
        """Convert Outlook color code to matplotlib RGB color"""
        # Approximate mapping of Outlook colors to RGB hex values
        color_map = {
            1: '#DC143C',   # Red
            2: '#FF8C00',   # Orange
            3: '#FFB6C1',   # Peach
            4: '#FFD700',   # Yellow
            5: '#32CD32',   # Green
            6: '#008B8B',   # Teal
            7: '#808000',   # Olive
            8: '#4169E1',   # Blue
            9: '#9370DB',   # Purple
            10: '#800000',  # Maroon
            11: '#4682B4',  # Steel
            12: '#36454F',  # DarkSteel
            13: '#808080',  # Gray
            14: '#696969',  # DarkGray
            15: '#000000',  # Black
            16: '#8B0000',  # DarkRed
            17: '#FF4500',  # DarkOrange
            18: '#CD5C5C',  # DarkPeach
            19: '#DAA520',  # DarkYellow
            20: '#006400',  # DarkGreen
            21: '#008080',  # DarkTeal
            22: '#556B2F',  # DarkOlive
            23: '#00008B',  # DarkBlue
            24: '#483D8B',  # DarkPurple
        }
        return color_map.get(color_code, '#DC143C')  # Default to Red
    
    def auto_assign_default_colors(self):
        """Auto-assign default Outlook colors to regions (starting from 1: Red)"""
        if not self.n_clusters:
            return
        
        # Only assign colors that haven't been set yet
        for i in range(self.n_clusters):
            region_num = i + 1
            if region_num not in self.region_colors:
                # Cycle through colors starting from 1 (Red)
                # Skip 0 (None) and use colors 1-24
                color_index = ((i % 24) + 1)
                self.region_colors[region_num] = color_index
        
        # Save the colors
        self.save_region_colors()
        self.log(f"✓ Auto-assigned default colors to {self.n_clusters} regions")
    
    def save_region_colors(self):
        """Save region colors to CSV along with names"""
        if not self.output_dir:
            return
        
        try:
            names_file = os.path.join(self.output_dir, "region_names.csv")
            data = []
            
            # Only use regions from current clustering (1 to n_clusters)
            # This prevents stale data from previous runs with more regions
            if self.n_clusters:
                all_regions = set(range(1, self.n_clusters + 1))
            else:
                # Fallback to dictionary keys if n_clusters not set
                all_regions = set(self.region_names.keys()) | set(self.region_colors.keys())
            
            for region in sorted(all_regions):
                name = self.region_names.get(region, f"Region {region}")
                color = self.region_colors.get(region, 1)  # Default to Red (1)
                data.append({
                    'region': region,
                    'name': name,
                    'color_code': color
                })
            
            if data:
                df = pd.DataFrame(data)
                df.to_csv(names_file, index=False)
                self.log(f"✓ Saved region names and colors to region_names.csv")
            elif os.path.exists(names_file):
                # Remove file if no data
                os.remove(names_file)
        except Exception as e:
            self.log(f"⚠ Failed to save region colors: {e}")
    
    def load_region_colors(self):
        """Load region colors from CSV"""
        if not self.output_dir:
            return
        
        names_file = os.path.join(self.output_dir, "region_names.csv")
        if not os.path.exists(names_file):
            return
        
        try:
            df = pd.read_csv(names_file)
            self.region_colors = {}
            
            # Check if color_code column exists
            if 'color_code' in df.columns:
                for _, row in df.iterrows():
                    region_num = int(row['region'])
                    color_code = int(row['color_code'])
                    self.region_colors[region_num] = color_code
                self.log(f"✓ Loaded color codes for {len(self.region_colors)} regions")
            else:
                self.log("⚠ No color codes found in region_names.csv")
        except Exception as e:
            self.log(f"⚠ Failed to load region colors: {e}")
    
def main():
    # Check for project directory argument
    project_dir = None
    if len(sys.argv) > 1:
        project_dir = sys.argv[1]
        if not os.path.exists(project_dir):
            print(f"Warning: Project directory not found: {project_dir}")
            project_dir = None
    
    root = tk.Tk()
    app = TSPClusteringApp(root, project_dir=project_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
