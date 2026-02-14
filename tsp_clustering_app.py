import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from sklearn.cluster import AgglomerativeClustering
from scipy.spatial import ConvexHull
from shapely.geometry import Polygon
import threading
import os
import sys

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
        main_frame.rowconfigure(1, weight=1)
        
        # Button bar at top for quick menu access
        button_bar = ttk.Frame(main_frame)
        button_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.reload_btn = ttk.Button(button_bar, text="Reload Files", command=self.load_previous_clustering, width=12)
        self.reload_btn.pack(side=tk.LEFT, padx=2)
        self.config_btn = ttk.Button(button_bar, text="Configure", command=self.show_config_menu, width=12)
        self.config_btn.pack(side=tk.LEFT, padx=2)
        self.run_btn = ttk.Button(button_bar, text="Run", command=self.show_run_menu, width=12)
        self.run_btn.pack(side=tk.LEFT, padx=2)
        self.save_btn = ttk.Button(button_bar, text="Save Results", command=self.save_results, width=12, state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT, padx=2)
        self.edit_btn = ttk.Button(button_bar, text="Edit Regions", command=self.show_edit_regions_dialog, width=12, state=tk.DISABLED)
        self.edit_btn.pack(side=tk.LEFT, padx=2)
        self.rename_color_btn = ttk.Button(button_bar, text="Rename/Recolor", command=self.show_rename_recolor_dialog, width=15, state=tk.DISABLED)
        self.rename_color_btn.pack(side=tk.LEFT, padx=2)
        self.view_btn = ttk.Button(button_bar, text="Analytics", command=self.show_log_window, width=12)
        self.view_btn.pack(side=tk.LEFT, padx=2)
        
        # Status bar showing current configuration
        status_bar_frame = ttk.Frame(main_frame, relief=tk.SUNKEN, borderwidth=1)
        status_bar_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.locations_status = ttk.Label(status_bar_frame, text="Locations: Not loaded", 
                                         font=('Arial', 9))
        self.locations_status.pack(side=tk.LEFT, padx=5)
        
        self.distances_status = ttk.Label(status_bar_frame, text="Distances: Not loaded", 
                                         font=('Arial', 9))
        self.distances_status.pack(side=tk.LEFT, padx=5)
        
        self.output_status = ttk.Label(status_bar_frame, text="Output: Not set", 
                                      font=('Arial', 9))
        self.output_status.pack(side=tk.LEFT, padx=5)
        
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
        ttk.Label(welcome_frame, text="Configure clustering parameters with the Configure button", 
                 font=('Arial', 12)).pack(pady=10)
        ttk.Label(welcome_frame, text="Click Run button to start clustering analysis", 
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
    
    def auto_load_project_files(self):
        """Auto-load files from project directory"""
        if not self.project_dir or not os.path.exists(self.project_dir):
            return
        
        project_name = os.path.basename(self.project_dir)
        self.root.title(f"TSP Regional Clustering Optimizer - Project: {project_name}")
        
        # Set output directory to project directory
        self.output_dir = self.project_dir
        self.output_status.config(text=f"Output: {project_name}", foreground="green")
        
        # Load locations.csv
        locations_path = os.path.join(self.project_dir, "locations.csv")
        if os.path.exists(locations_path):
            self.locations_file = locations_path
            self.locations_status.config(text=f"Locations: locations.csv", foreground="green")
            self.log(f"✓ Auto-loaded: {locations_path}")
        else:
            self.log(f"⚠ locations.csv not found in project directory")
        
        # Load distances.csv
        distances_path = os.path.join(self.project_dir, "distances.csv")
        if os.path.exists(distances_path):
            self.distances_file = distances_path
            self.distances_status.config(text=f"Distances: distances.csv", foreground="green")
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
    
    def show_run_menu(self):
        """Show Run dialog window"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Run Options")
        dialog.geometry("350x180")
        dialog.transient(self.root)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Run Operations", font=('Arial', 14, 'bold')).pack(pady=(0, 20))
        
        ttk.Button(frame, text="Run Clustering Analysis", command=self.start_clustering,
                  width=30).pack(pady=5)
        ttk.Button(frame, text="Reset Clustering", command=self.reset_clustering,
                  width=30).pack(pady=5)
    
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
        
        ttk.Button(btn_frame, text="OK", command=dialog.destroy, 
                  width=10).pack(side=tk.LEFT, padx=5)
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
            customers_df = results_df[results_df['region'] != 0]
            coords = customers_df[['latitude', 'longitude']].values
            customer_postcodes = customers_df['postcode'].tolist()
            
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
            self.depot_postcode = depot_postcode
            
            # Prepare results for potential saving
            self.clustered_results = results_df
            
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
            self.create_visualization(coords, labels, depot, n_clusters, customer_postcodes, depot_postcode)
            
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
            customers_df = locations_df[locations_df['postcode'].str.upper() != depot_postcode]
            coords = customers_df[['latitude', 'longitude']].values
            self.log(f"✓ Clustering {len(coords)} customers (depot excluded)")
            
            # Use desired regions directly
            actual_regions = desired_regions
            
            self.log(f"\nUsing {actual_regions} regions")
            self.progress_bar['value'] = 35
            
            # Update postcode_to_idx for customers only (excluding depot)
            customer_postcodes = sorted(customers_df['postcode'])
            customer_postcode_to_idx = {pc: i for i, pc in enumerate(customer_postcodes)}
            
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
            results_df['region'] = [labels[customer_postcode_to_idx[pc]] + 1 for pc in results_df['postcode']]
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
            self.root.after(0, lambda: self.create_visualization(coords, labels, depot, actual_regions, customer_postcodes, depot_postcode))
            
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
            
    def create_visualization(self, coords, labels, depot, n_clusters, customer_postcodes, depot_postcode):
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
        
        # Create figure
        fig = Figure(figsize=(12, 8), dpi=100)
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
        for idx, (coord, postcode) in enumerate(zip(coords, customer_postcodes)):
            # Different styling for excluded postcodes
            if labels[idx] == -1:
                bbox_style = dict(boxstyle='round,pad=0.3', facecolor='lightgray', 
                                edgecolor='red', alpha=0.7, linestyle='--', linewidth=1.5)
            else:
                bbox_style = dict(boxstyle='round,pad=0.3', facecolor='white', 
                                edgecolor='gray', alpha=0.7)
            
            ax.annotate(postcode, 
                       xy=(coord[1], coord[0]),
                       xytext=(3, 3),  # Offset by 3 points
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
        
        fig.tight_layout()
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.viz_canvas_container)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # Add toolbar
        toolbar = NavigationToolbar2Tk(canvas, self.viz_canvas_container)
        toolbar.update()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
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
        
        # Create figure
        fig = Figure(figsize=(12, 8), dpi=100)
        ax = fig.add_subplot(111)
        
        # Plot all locations (not yet clustered)
        ax.scatter(coords[:, 1], coords[:, 0], 
                  c='lightblue', s=100, alpha=0.6, 
                  edgecolors='black', linewidth=1,
                  label=f'{len(coords)} Locations (Unclustered)')
        
        # Add postcode labels for all locations
        for idx, (coord, postcode) in enumerate(zip(coords, postcodes)):
            # Skip depot in this loop (will be added separately)
            if postcode.upper() == depot_postcode.upper():
                continue
            ax.annotate(postcode, 
                       xy=(coord[1], coord[0]),
                       xytext=(3, 3),
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
        
        fig.tight_layout()
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.viz_canvas_container)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # Add toolbar
        toolbar = NavigationToolbar2Tk(canvas, self.viz_canvas_container)
        toolbar.update()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        self.canvas = canvas
        self.toolbar = toolbar
    
    def show_edit_regions_dialog(self):
        """Show dialog for editing location regions manually"""
        if not self.has_results:
            messagebox.showwarning("No Results", 
                                  "No clustering results available.\n\n"
                                  "Run clustering analysis first.")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Location's Region")
        dialog.geometry("550x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Edit Location's Region", font=('Arial', 14, 'bold')).pack(pady=(0, 20))
        
        # Instructions
        instructions = ttk.Label(frame, text="Select a postcode and assign it to a different region or exclude it.",
                                font=('Arial', 9), foreground='gray')
        instructions.pack(pady=(0, 15))
        
        # Postcode selection
        postcode_frame = ttk.Frame(frame)
        postcode_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(postcode_frame, text="Postcode:", font=('Arial', 10)).pack(side=tk.LEFT, padx=(0, 10))
        
        # Get all postcodes except depot, sorted alphabetically
        all_postcodes = sorted(self.clustered_results[self.clustered_results['region'] != 0]['postcode'].tolist())
        postcode_var = tk.StringVar()
        postcode_combo = ttk.Combobox(postcode_frame, textvariable=postcode_var, 
                                     values=all_postcodes, state='readonly', width=20)
        postcode_combo.pack(side=tk.LEFT, padx=(0, 10))
        
        # Current region display
        current_region_var = tk.StringVar(value="Select a postcode")
        ttk.Label(postcode_frame, text="Current Region:", font=('Arial', 9)).pack(side=tk.LEFT, padx=(10, 5))
        current_region_label = ttk.Label(postcode_frame, textvariable=current_region_var, 
                                        font=('Arial', 9, 'bold'), foreground='blue')
        current_region_label.pack(side=tk.LEFT)
        
        def on_postcode_selected(event):
            selected = postcode_var.get()
            if selected:
                current_region = self.clustered_results[
                    self.clustered_results['postcode'] == selected
                ]['region'].iloc[0]
                if current_region == -1:
                    current_region_var.set("Excluded")
                else:
                    current_region_var.set(f"Region {current_region}")
        
        postcode_combo.bind('<<ComboboxSelected>>', on_postcode_selected)
        
        # New region selection
        region_frame = ttk.Frame(frame)
        region_frame.pack(fill=tk.X, pady=20)
        
        ttk.Label(region_frame, text="New Region:", font=('Arial', 10)).pack(side=tk.LEFT, padx=(0, 10))
        
        # Region options: 1 to n_clusters, plus "Create New Region" and "Exclude"
        region_options = [f"Region {i+1}" for i in range(self.n_clusters)] + [f"Create New Region {self.n_clusters + 1}"] + ["Exclude from all regions"]
        new_region_var = tk.StringVar()
        region_combo = ttk.Combobox(region_frame, textvariable=new_region_var, 
                                   values=region_options, state='readonly', width=30)
        region_combo.pack(side=tk.LEFT)
        
        # Apply button
        def apply_edit():
            selected_postcode = postcode_var.get()
            new_region_str = new_region_var.get()
            
            if not selected_postcode:
                messagebox.showwarning("No Postcode", "Please select a postcode.")
                return
            
            if not new_region_str:
                messagebox.showwarning("No Region", "Please select a new region.")
                return
            
            # Parse new region
            if new_region_str == "Exclude from all regions":
                new_region = -1
                region_display = "Excluded"
            elif new_region_str.startswith("Create New Region"):
                # Create a new region
                new_region = self.n_clusters + 1
                region_display = f"Region {new_region}"
                # Increment the cluster count
                self.n_clusters += 1
                self.log(f"\n✨ Created new {region_display}")
            else:
                new_region = int(new_region_str.split()[-1])
                region_display = new_region_str
            
            # Get current region
            current_region = self.clustered_results[
                self.clustered_results['postcode'] == selected_postcode
            ]['region'].iloc[0]
            
            if current_region == new_region:
                messagebox.showinfo("No Change", f"{selected_postcode} is already in {region_display}.")
                return
            
            # Update the region in clustered_results
            self.clustered_results.loc[
                self.clustered_results['postcode'] == selected_postcode, 'region'
            ] = new_region
            
            # Update the labels array for visualization
            postcode_idx = self.customer_postcodes.index(selected_postcode)
            if new_region == -1:
                # For excluded, we'll use a special value
                self.labels[postcode_idx] = -1
            else:
                self.labels[postcode_idx] = new_region - 1  # labels are 0-indexed
            
            # Update summary
            self.update_summary_results()
            
            # Log the change
            old_display = "Excluded" if current_region == -1 else f"Region {current_region}"
            self.log(f"\n✏️ Manual Edit: {selected_postcode} moved from {old_display} to {region_display}")
            
            # Refresh visualization
            self.refresh_visualization()
            
            messagebox.showinfo("Success", 
                              f"{selected_postcode} has been moved to {region_display}.\n\n"
                              f"The visualization has been updated.")
            
            # Update current region display
            current_region_var.set(region_display)
            
            # Update the region dropdown to include any newly created regions
            region_options = [f"Region {i+1}" for i in range(self.n_clusters)] + [f"Create New Region {self.n_clusters + 1}"] + ["Exclude from all regions"]
            region_combo.config(values=region_options)
        
        apply_frame = ttk.Frame(frame)
        apply_frame.pack(pady=20)
        
        ttk.Button(apply_frame, text="Apply Change", command=apply_edit, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(apply_frame, text="Close", command=dialog.destroy, width=15).pack(side=tk.LEFT, padx=5)
        
        # Info text
        info_text = ttk.Label(frame, 
                             text="Note: Changes are applied immediately to the visualization.\n"
                                  "Use 'Run > Save Results to CSV' to save your edits.",
                             font=('Arial', 8), foreground='gray', justify=tk.CENTER)
        info_text.pack(pady=(20, 0))
    
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
        self.create_visualization(
            self.coords, 
            self.labels, 
            self.depot, 
            self.n_clusters, 
            self.customer_postcodes, 
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
        canvas = tk.Canvas(frame, height=400)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
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
        btn_frame.pack(pady=15)
        
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
