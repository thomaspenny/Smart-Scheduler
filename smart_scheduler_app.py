import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import win32com.client
import time

# Outlook Category Colors Enumeration (OlCategoryColor)
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


class SmartSchedulerApp:
    def __init__(self, root, project_dir=None):
        self.root = root
        self.root.title("Smart Scheduler")
        self.root.state('zoomed')  # Fullscreen
        
        # Project directory from command line
        self.project_dir = project_dir
        
        # Data
        self.regions_df = None
        self.schedule_df = None
        self.distances_df = None
        self.region_names_df = None
        self.clustered_regions_df = None
        self.home_postcode = None  # Home base postcode
        
        # Current selection
        self.selected_region = None
        self.selected_dates = []
        self.region_postcodes = []  # Postcodes in selected region
        self.appointments = {}  # {(date, time_slot): 'postcode'} - temporary/visual only
        self.pending_appointment = None  # Staged appointment: (date, time, postcode, duration) before submit
        self.confirmed_appointments = {}  # Confirmed appointments: {postcode: (date, time, duration)} from CSV
        self.travel_segments = []  # List of (date, start_minutes, end_minutes, info_dict)
        self.conflicting_segments = set()  # Set of (date, start_minutes, end_minutes) tuples for conflicts
        
        # Timetable configuration
        self.start_hour = 8
        self.end_hour = 19
        self.appointment_duration = 60  # Appointment duration in minutes (default 1 hour)
        self.max_appointments_per_day = 4
        self.route_efficiency_threshold = 1.3  # Routes can be max 130% of optimal
        
        # Time slots (30-minute intervals from start to end hour)
        self.generate_time_slots()
        
        # Initialize appointments CSV path
        if self.project_dir:
            self.appointments_csv = Path(self.project_dir) / 'confirmed_appointments.csv'
        else:
            self.appointments_csv = None
        
        self.setup_ui()
        
        # Load project data if available
        if self.project_dir:
            self.load_project_data()
            self.load_confirmed_appointments()
    
    def generate_time_slots(self):
        """Generate time slots based on start and end hours"""
        self.time_slots = []
        start_time = self.start_hour * 60
        end_time = self.end_hour * 60
        for minutes in range(start_time, end_time, 30):
            hours = minutes // 60
            mins = minutes % 60
            self.time_slots.append(f"{hours}:{mins:02d}")
    
    def setup_ui(self):
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=10)
        main_frame.rowconfigure(4, weight=1)
        
        # Title and Analysis Button Row
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        title_frame.columnconfigure(0, weight=1)
        
        title_label = ttk.Label(title_frame, text="Smart Scheduler", 
                               font=('Arial', 18, 'bold'))
        title_label.pack(side=tk.LEFT)
        
        # Selection frame
        selection_frame = ttk.LabelFrame(main_frame, text="Selection", padding="10")
        selection_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Configure columns to not expand (keep them compact)
        for col in range(8):
            selection_frame.columnconfigure(col, weight=0, minsize=0)
        
        # Row 0: Timetable configuration
        ttk.Label(selection_frame, text="Timetable Start:", font=('Arial', 10)).grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        # Timetable start in a frame to keep controls close together
        start_frame = ttk.Frame(selection_frame)
        start_frame.grid(row=0, column=1, sticky=tk.W)
        self.start_hour_var = tk.StringVar(value=str(self.start_hour))
        start_spinbox = ttk.Spinbox(start_frame, from_=0, to=23, textvariable=self.start_hour_var, 
                                   width=3, command=self.on_time_config_changed)
        start_spinbox.pack(side=tk.LEFT)
        ttk.Label(start_frame, text=":00", font=('Arial', 10)).pack(side=tk.LEFT, padx=2)
        
        ttk.Label(selection_frame, text="End:", font=('Arial', 10)).grid(row=0, column=2, sticky=tk.W, padx=(20, 5))
        
        # Timetable end in a frame to keep controls close together
        end_frame = ttk.Frame(selection_frame)
        end_frame.grid(row=0, column=3, sticky=tk.W)
        self.end_hour_var = tk.StringVar(value=str(self.end_hour))
        end_spinbox = ttk.Spinbox(end_frame, from_=1, to=24, textvariable=self.end_hour_var, 
                                 width=3, command=self.on_time_config_changed)
        end_spinbox.pack(side=tk.LEFT)
        ttk.Label(end_frame, text=":00", font=('Arial', 10)).pack(side=tk.LEFT, padx=2)
        
        # Region and Postcode selection
        ttk.Label(selection_frame, text="Region:", font=('Arial', 10)).grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(10, 0))
        self.region_var = tk.StringVar()
        self.region_combo = ttk.Combobox(selection_frame, textvariable=self.region_var, 
                                        state='readonly', width=30)
        self.region_combo.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=(10, 0))
        self.region_combo.bind('<<ComboboxSelected>>', self.on_region_selected)
        
        ttk.Label(selection_frame, text="Postcode:", font=('Arial', 10)).grid(row=1, column=3, sticky=tk.W, padx=(20, 5), pady=(10, 0))
        self.postcode_var = tk.StringVar()
        self.postcode_combo = ttk.Combobox(selection_frame, textvariable=self.postcode_var, 
                                          state='readonly', width=12)
        self.postcode_combo.grid(row=1, column=4, sticky=tk.W, pady=(10, 0))
        self.postcode_combo.bind('<<ComboboxSelected>>', self.on_postcode_selected)
        
        ttk.Label(selection_frame, text="Appt Duration:", font=('Arial', 10)).grid(row=1, column=5, sticky=tk.W, padx=(20, 5), pady=(10, 0))
        self.appointment_duration_var = tk.StringVar(value=str(self.appointment_duration))
        ttk.Spinbox(selection_frame, from_=30, to=180, textvariable=self.appointment_duration_var, 
                   width=4, increment=30).grid(row=1, column=6, sticky=tk.W, pady=(10, 0))
        ttk.Label(selection_frame, text="min", font=('Arial', 10)).grid(row=1, column=7, sticky=tk.W, padx=(2, 0), pady=(10, 0))
        
        ttk.Label(selection_frame, text="Home Base:", font=('Arial', 10)).grid(row=1, column=8, sticky=tk.W, padx=(20, 5), pady=(10, 0))
        self.home_label = ttk.Label(selection_frame, text="-", font=('Arial', 10, 'bold'), foreground='blue')
        self.home_label.grid(row=1, column=9, sticky=tk.W, pady=(10, 0))
        
        # Timetable frame with scrollbars
        timetable_frame = ttk.LabelFrame(main_frame, text="Timetable", padding="10")
        timetable_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        main_frame.grid_rowconfigure(2, minsize=300)
        timetable_frame.columnconfigure(0, weight=1)
        timetable_frame.rowconfigure(0, weight=1)
        
        # Create canvas with scrollbars
        canvas = tk.Canvas(timetable_frame, bg='white')
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        v_scrollbar = ttk.Scrollbar(timetable_frame, orient=tk.VERTICAL, command=canvas.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        h_scrollbar = ttk.Scrollbar(timetable_frame, orient=tk.HORIZONTAL, command=canvas.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Frame inside canvas for timetable
        self.timetable_inner_frame = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=self.timetable_inner_frame, anchor='nw')
        
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
        
        self.timetable_inner_frame.bind('<Configure>', configure_scroll_region)
        
        self.canvas = canvas
        
        # Status bar and legend
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.status_label = ttk.Label(status_frame, text="Ready", 
                                     font=('Arial', 9), foreground='green')
        self.status_label.pack(side=tk.LEFT)
        
        # Pending appointment label
        self.pending_label = ttk.Label(status_frame, text="", 
                                      font=('Arial', 9, 'bold'), foreground='orange')
        self.pending_label.pack(side=tk.LEFT, padx=(20, 0))
        
        # Submit and Clear buttons
        ttk.Button(status_frame, text="Submit Appointment", 
                  command=self.submit_appointment).pack(side=tk.RIGHT, padx=(0, 10))
        ttk.Button(status_frame, text="Sync to Outlook", 
                  command=self.sync_to_outlook).pack(side=tk.RIGHT, padx=(0, 10))
        ttk.Button(status_frame, text="Clear Schedule", 
                  command=self.clear_schedule).pack(side=tk.RIGHT, padx=(0, 10))
        
        # Legend
        legend_frame = ttk.Frame(status_frame)
        legend_frame.pack(side=tk.RIGHT, padx=20)
        
        ttk.Label(legend_frame, text="Legend:", font=('Arial', 9, 'bold')).pack(side=tk.LEFT, padx=(0, 10))
        
        # Confirmed Appointment color
        appt_canvas = tk.Canvas(legend_frame, width=20, height=15, bg='#90EE90', highlightthickness=1, highlightbackground='black')
        appt_canvas.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(legend_frame, text="Confirmed", font=('Arial', 8)).pack(side=tk.LEFT, padx=(0, 15))
        
        # Pending Appointment color
        pending_canvas = tk.Canvas(legend_frame, width=20, height=15, bg='#228B22', highlightthickness=1, highlightbackground='black')
        pending_canvas.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(legend_frame, text="Pending", font=('Arial', 8)).pack(side=tk.LEFT, padx=(0, 15))
        
        # Travel to appointment color
        travel_appt_canvas = tk.Canvas(legend_frame, width=20, height=15, bg='#FFD700', highlightthickness=1, highlightbackground='black')
        travel_appt_canvas.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(legend_frame, text="Travel (to appt)", font=('Arial', 8)).pack(side=tk.LEFT, padx=(0, 15))
        
        # Travel from home color
        travel_from_home_canvas = tk.Canvas(legend_frame, width=20, height=15, bg='#87CEEB', highlightthickness=1, highlightbackground='black')
        travel_from_home_canvas.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(legend_frame, text="Travel (from home)", font=('Arial', 8)).pack(side=tk.LEFT, padx=(0, 15))
        
        # Travel home color
        travel_home_canvas = tk.Canvas(legend_frame, width=20, height=15, bg='#FFA500', highlightthickness=1, highlightbackground='black')
        travel_home_canvas.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(legend_frame, text="Travel (to home)", font=('Arial', 8)).pack(side=tk.LEFT, padx=(0, 15))
        
        # Conflict color
        conflict_canvas = tk.Canvas(legend_frame, width=20, height=15, bg='#FF0000', highlightthickness=1, highlightbackground='black')
        conflict_canvas.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(legend_frame, text="Conflict", font=('Arial', 8)).pack(side=tk.LEFT)
        
        # Bottom area - split into map (left) and suggestions (right)
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)
        bottom_frame.rowconfigure(0, weight=1)
        
        # Left side - Visualization minimap
        viz_frame = ttk.LabelFrame(bottom_frame, text="Region Map", padding="10")
        viz_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        viz_frame.columnconfigure(0, weight=1)
        viz_frame.rowconfigure(0, weight=1)
        
        # Create larger matplotlib figure for map
        self.fig = Figure(figsize=(8, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.viz_canvas = FigureCanvasTkAgg(self.fig, master=viz_frame)
        self.viz_canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Right side - Travel Times
        suggestions_frame = ttk.LabelFrame(bottom_frame, text="Travel Times from Selected Postcode", padding="10")
        suggestions_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        suggestions_frame.columnconfigure(0, weight=1)
        suggestions_frame.rowconfigure(0, weight=1)
        
        # Scrollable text widget for travel times
        self.suggestions_text = tk.Text(suggestions_frame, height=10, width=40, wrap=tk.WORD, 
                                       font=('Consolas', 10), state='disabled')
        suggestions_scrollbar = ttk.Scrollbar(suggestions_frame, orient=tk.VERTICAL, 
                                             command=self.suggestions_text.yview)
        self.suggestions_text.config(yscrollcommand=suggestions_scrollbar.set)
        self.suggestions_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        suggestions_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    
    def load_project_data(self):
        """Load project data files"""
        try:
            # Load region schedule
            schedule_path = os.path.join(self.project_dir, "region_schedule.csv")
            if os.path.exists(schedule_path):
                self.schedule_df = pd.read_csv(schedule_path)
                self.schedule_df['date'] = pd.to_datetime(self.schedule_df['date'])
            
            # Load region names
            names_path = os.path.join(self.project_dir, "region_names.csv")
            if os.path.exists(names_path):
                self.region_names_df = pd.read_csv(names_path)
            
            # Load clustered regions
            clustered_path = os.path.join(self.project_dir, "clustered_regions.csv")
            if os.path.exists(clustered_path):
                self.clustered_regions_df = pd.read_csv(clustered_path)
                
                # Get home base from region 0 (depot)
                depot_region = self.clustered_regions_df[self.clustered_regions_df['region'] == 0]
                if len(depot_region) > 0:
                    self.home_postcode = depot_region['postcode'].iloc[0].strip().upper()
                    self.home_label.config(text=self.home_postcode)
            
            # Load distances
            distances_path = os.path.join(self.project_dir, "distances.csv")
            if os.path.exists(distances_path):
                self.distances_df = pd.read_csv(distances_path)
            
            # Populate region dropdown
            if self.region_names_df is not None and self.schedule_df is not None:
                region_options = []
                for _, row in self.region_names_df.iterrows():
                    region_id = row['region']
                    region_name = row['name']
                    # Count dates for this region
                    date_count = len(self.schedule_df[self.schedule_df['region'] == region_id])
                    region_options.append(f"Region {region_id}: {region_name} ({date_count} dates)")
                
                self.region_combo['values'] = region_options
                if region_options:
                    self.region_combo.current(0)
                    self.on_region_selected(None)
            
            self.status_label.config(text="Project data loaded successfully", foreground='green')
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load project data:\n{e}")
            self.status_label.config(text="Error loading data", foreground='red')
    
    def on_time_config_changed(self):
        """Handle timetable start/end time changes"""
        try:
            new_start = int(self.start_hour_var.get())
            new_end = int(self.end_hour_var.get())
            
            if new_start >= new_end:
                messagebox.showwarning("Invalid Times", "Start time must be before end time.")
                self.start_hour_var.set(str(self.start_hour))
                self.end_hour_var.set(str(self.end_hour))
                return
            
            self.start_hour = new_start
            self.end_hour = new_end
            self.generate_time_slots()
            
            # Rebuild appointments and travel segments from confirmed appointments
            self.appointments.clear()
            self.travel_segments.clear()
            
            # Repopulate appointments from confirmed appointments
            for postcode, (date, time, duration, in_outlook) in self.confirmed_appointments.items():
                cell_key = (date, time)
                self.appointments[cell_key] = postcode
            
            # Also add pending appointment if exists
            if self.pending_appointment:
                pending_date, pending_time, pending_postcode, pending_duration = self.pending_appointment
                cell_key = (pending_date, pending_time)
                self.appointments[cell_key] = pending_postcode
            
            # Recalculate travel times for all dates with appointments
            dates_with_appointments = set([date for (date, time) in self.appointments.keys()])
            for date in dates_with_appointments:
                self.recalculate_travel_times(date)
            
            self.update_timetable()
            self.status_label.config(text=f"Timetable updated: {self.start_hour}:00 - {self.end_hour}:00", foreground='blue')
        except ValueError:
            pass
    
    def on_region_selected(self, event):
        """Handle region selection"""
        selection = self.region_var.get()
        if not selection:
            return
        
        # Extract region ID from selection
        region_id = int(selection.split(':')[0].replace('Region ', ''))
        self.selected_region = region_id
        
        # Get dates for this region
        if self.schedule_df is not None:
            region_schedule = self.schedule_df[self.schedule_df['region'] == region_id]
            self.selected_dates = sorted(region_schedule['date'].dt.date.tolist())
        
        # Get postcodes for this region
        self.region_postcodes = []
        if self.clustered_regions_df is not None:
            region_data = self.clustered_regions_df[self.clustered_regions_df['region'] == region_id]
            self.region_postcodes = sorted(region_data['postcode'].unique().tolist())
            self.postcode_combo['values'] = self.region_postcodes
            if self.region_postcodes:
                self.postcode_combo.current(0)
        
        # Calculate optimal days needed
        optimal_days = self.calculate_optimal_days()
        
        # Update timetable
        self.update_timetable()
        self.update_region_visualization()
        
        # Update travel times display for the first postcode
        if self.region_postcodes:
            self.display_travel_times(self.region_postcodes[0])
        
        self.status_label.config(text=f"Region {region_id}: {len(self.region_postcodes)} postcodes, {len(self.selected_dates)} dates available, {optimal_days} optimal days", 
                                foreground='blue')
    
    def calculate_optimal_days(self):
        """Calculate optimal number of days needed for region based on max appointments per day"""
        if not self.region_postcodes:
            return 0
        
        num_postcodes = len(self.region_postcodes)
        
        # Calculate minimum days needed
        import math
        optimal_days = math.ceil(num_postcodes / self.max_appointments_per_day)
        
        return optimal_days
    
    def update_region_visualization(self):
        """Update the map visualization for the selected region"""
        self.ax.clear()
        
        if self.selected_region is None or self.clustered_regions_df is None:
            self.ax.text(0.5, 0.5, 'No region selected', 
                        horizontalalignment='center', verticalalignment='center',
                        transform=self.ax.transAxes, fontsize=12)
            self.viz_canvas.draw()
            return
        
        # Get locations for this region
        region_data = self.clustered_regions_df[self.clustered_regions_df['region'] == self.selected_region]
        
        if len(region_data) == 0:
            self.ax.text(0.5, 0.5, 'No locations in this region', 
                        horizontalalignment='center', verticalalignment='center',
                        transform=self.ax.transAxes, fontsize=12)
            self.viz_canvas.draw()
            return
        
        # Draw links between appointments (confirmed and pending, grouped by date)
        appointments_by_date = {}
        
        # Add confirmed appointments
        for postcode, (date, time, duration, in_outlook) in self.confirmed_appointments.items():
            if date not in appointments_by_date:
                appointments_by_date[date] = []
            appointments_by_date[date].append((time, postcode, True))  # True = confirmed
        
        # Add pending appointment if it exists
        if self.pending_appointment:
            pending_date, pending_time, pending_postcode, pending_duration = self.pending_appointment
            if pending_date not in appointments_by_date:
                appointments_by_date[pending_date] = []
            appointments_by_date[pending_date].append((pending_time, pending_postcode, False))  # False = pending
        
        # Define colors for different dates
        date_colors = ['#0066CC', '#CC0066', '#00CC66', '#CC6600', '#6600CC', '#CCCC00']
        sorted_dates = sorted(appointments_by_date.keys())
        
        # Get home base coordinates
        home_coords = None
        if self.home_postcode and self.clustered_regions_df is not None:
            home_data = self.clustered_regions_df[self.clustered_regions_df['postcode'] == self.home_postcode]
            if len(home_data) > 0:
                home_row = home_data.iloc[0]
                home_coords = (home_row['longitude'], home_row['latitude'])
        
        # For each date, draw lines connecting appointments in time order
        for date_idx, date in enumerate(sorted_dates):
            appointments = appointments_by_date[date]
            # Sort by time - convert time strings to minutes for proper sorting
            def time_to_minutes(time_str):
                parts = time_str.split(':')
                return int(parts[0]) * 60 + int(parts[1])
            
            appointments.sort(key=lambda x: time_to_minutes(x[0]))
            postcodes_ordered = [pc for _, pc, _ in appointments]
            
            # Get color for this date
            color = date_colors[date_idx % len(date_colors)]
            
            # Draw line from home to first appointment
            if home_coords and len(postcodes_ordered) > 0:
                first_pc = postcodes_ordered[0]
                first_loc = region_data[region_data['postcode'] == first_pc]
                if len(first_loc) > 0:
                    x1, y1 = home_coords
                    x2, y2 = first_loc.iloc[0]['longitude'], first_loc.iloc[0]['latitude']
                    self.ax.plot([x1, x2], [y1, y2], color=color, linewidth=2, alpha=0.5, linestyle='--', zorder=2)
            
            # Draw lines between consecutive appointments
            for i in range(len(postcodes_ordered) - 1):
                pc1, pc2 = postcodes_ordered[i], postcodes_ordered[i+1]
                
                # Get coordinates
                loc1 = region_data[region_data['postcode'] == pc1]
                loc2 = region_data[region_data['postcode'] == pc2]
                
                if len(loc1) > 0 and len(loc2) > 0:
                    x1, y1 = loc1.iloc[0]['longitude'], loc1.iloc[0]['latitude']
                    x2, y2 = loc2.iloc[0]['longitude'], loc2.iloc[0]['latitude']
                    self.ax.plot([x1, x2], [y1, y2], color=color, linewidth=2, alpha=0.7, zorder=2, label=date)
            
            # Draw line from last appointment back to home
            if home_coords and len(postcodes_ordered) > 0:
                last_pc = postcodes_ordered[-1]
                last_loc = region_data[region_data['postcode'] == last_pc]
                if len(last_loc) > 0:
                    x1, y1 = last_loc.iloc[0]['longitude'], last_loc.iloc[0]['latitude']
                    x2, y2 = home_coords
                    self.ax.plot([x1, x2], [y1, y2], color=color, linewidth=2, alpha=0.5, linestyle='--', zorder=2)
        
        # Plot locations - highlight differently for scheduled vs unscheduled
        scheduled_postcodes = set(self.confirmed_appointments.keys())
        selected_postcode = self.postcode_var.get()
        
        for _, row in region_data.iterrows():
            pc = row['postcode']
            if pc in scheduled_postcodes:
                # Scheduled - green
                color = '#228B22'  # Forest green
                size = 150
            elif pc == selected_postcode:
                # Currently selected - orange
                color = '#FFA500'
                size = 150
            else:
                # Unscheduled - light green
                color = '#90EE90'
                size = 100
            
            self.ax.scatter(row['longitude'], row['latitude'], 
                           c=color, s=size, alpha=0.8, edgecolors='black', linewidth=1.5, zorder=3)
        
        # Add postcode labels
        for _, row in region_data.iterrows():
            self.ax.annotate(row['postcode'], 
                           (row['longitude'], row['latitude']),
                           xytext=(5, 5), textcoords='offset points',
                           fontsize=8, fontweight='bold')
        
        # Get home base location from region 0
        if self.home_postcode and self.clustered_regions_df is not None:
            home_data = self.clustered_regions_df[self.clustered_regions_df['postcode'] == self.home_postcode]
            if len(home_data) > 0:
                home_row = home_data.iloc[0]
                self.ax.scatter(home_row['longitude'], home_row['latitude'], 
                              c='red', s=200, marker='*', edgecolors='black', 
                              linewidth=2, zorder=5)
        
        # Add legend for route dates if there are any appointments
        if appointments_by_date:
            # Get unique labels from the plot (removes duplicates)
            handles, labels = self.ax.get_legend_handles_labels()
            by_label = dict(zip(labels, handles))
            self.ax.legend(by_label.values(), by_label.keys(), loc='upper right', 
                          title='Route Dates', fontsize=8, title_fontsize=9)
        
        # Format plot - remove labels and title to maximize graph area
        self.ax.set_xticks([])
        self.ax.set_yticks([])
        self.ax.grid(True, alpha=0.3)
        self.ax.set_aspect('equal', adjustable='datalim')
        
        self.fig.tight_layout(pad=0.1)
        self.viz_canvas.draw()
    
    def update_timetable(self):
        """Create/update the timetable grid"""
        # Clear existing timetable
        for widget in self.timetable_inner_frame.winfo_children():
            widget.destroy()
        
        if not self.selected_dates:
            ttk.Label(self.timetable_inner_frame, text="No dates available for selected region",
                     font=('Arial', 12)).grid(row=0, column=0, padx=20, pady=20)
            return
        
        # Create header row
        # Date column header
        date_header = tk.Label(self.timetable_inner_frame, text="Date", bg='#2C5F8D', fg='white',
                              font=('Arial', 10, 'bold'), width=15, height=2, relief=tk.RIDGE, bd=1)
        date_header.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Time slot headers
        for col, time_slot in enumerate(self.time_slots, start=1):
            time_label = tk.Label(self.timetable_inner_frame, text=time_slot, bg='#2C5F8D', fg='white',
                                 font=('Arial', 9, 'bold'), width=8, height=2, relief=tk.RIDGE, bd=1)
            time_label.grid(row=0, column=col, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create row for each date
        for row_idx, date in enumerate(self.selected_dates, start=1):
            # Date label
            date_str = date.strftime('%d-%b-%y')
            date_label = tk.Label(self.timetable_inner_frame, text=date_str, bg='#E8E8E8',
                                 font=('Arial', 9, 'bold'), width=15, height=3, relief=tk.RIDGE, bd=1)
            date_label.grid(row=row_idx, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            # Time slot cells
            for col_idx, time_slot in enumerate(self.time_slots, start=1):
                cell_key = (date_str, time_slot)
                
                # Convert time slot to minutes from midnight
                slot_start_minutes = self.time_to_minutes(time_slot)
                slot_end_minutes = slot_start_minutes + 30
                
                # Check if this cell is covered by a previous appointment (skip rendering it)
                is_covered = False
                
                # Check previous slots to see if any appointment covers this slot
                for check_col in range(1, col_idx):
                    check_time_slot = self.time_slots[check_col - 1]  # -1 because col_idx starts at 1
                    check_cell_key = (date_str, check_time_slot)
                    
                    if check_cell_key in self.appointments:
                        check_postcode = self.appointments[check_cell_key]
                        # Get the actual duration of this appointment
                        if check_postcode in self.confirmed_appointments:
                            _, _, check_duration, _ = self.confirmed_appointments[check_postcode]
                        else:
                            check_duration = int(self.appointment_duration_var.get())
                        
                        # Check if this appointment extends to cover the current slot
                        appt_start_minutes = self.time_to_minutes(check_time_slot)
                        appt_end_minutes = appt_start_minutes + check_duration
                        
                        if slot_start_minutes < appt_end_minutes:
                            is_covered = True
                            break
                
                if is_covered:
                    # Skip this cell as it's covered by a previous appointment's columnspan
                    continue
                
                # Check if there's an appointment starting at this time
                if cell_key in self.appointments:
                    # Appointment cell - check if confirmed or pending
                    postcode = self.appointments[cell_key]
                    
                    # Get duration - use stored duration for confirmed appointments, current setting for pending
                    if postcode in self.confirmed_appointments:
                        bg_color = '#90EE90'  # Light green for confirmed
                        # Get stored duration from confirmed appointments
                        _, _, duration_minutes, in_outlook = self.confirmed_appointments[postcode]
                        # Add email indicator if synced to Outlook
                        display_text = f"{postcode} ✉" if in_outlook else postcode
                    else:
                        bg_color = '#228B22'  # Forest green for pending (darker)
                        # Use current duration setting for pending appointments
                        duration_minutes = int(self.appointment_duration_var.get())
                        display_text = postcode
                    
                    # Calculate columnspan based on appointment duration (30-minute slots)
                    columnspan = duration_minutes // 30  # Each column is 30 minutes
                    
                    # Use larger font size if Outlook indicator is present for better visibility
                    font_size = 9 if (postcode in self.confirmed_appointments and self.confirmed_appointments[postcode][3]) else 8
                    
                    cell = tk.Label(self.timetable_inner_frame, text=display_text, bg=bg_color,
                                   font=('Arial', font_size, 'bold'), width=8, height=3, relief=tk.RIDGE, bd=1,
                                   cursor='hand2', anchor='center', justify='center', wraplength=60)
                    cell.grid(row=row_idx, column=col_idx, columnspan=columnspan, sticky=(tk.W, tk.E, tk.N, tk.S))
                    cell.bind('<Button-1>', lambda e, d=date_str, t=time_slot: self.on_cell_click(d, t))
                    
                else:
                    # Check if any travel segments overlap this time slot
                    overlapping_segments = []
                    for seg_date, seg_start, seg_end, seg_info in self.travel_segments:
                        if seg_date == date_str and seg_start < slot_end_minutes and seg_end > slot_start_minutes:
                            overlapping_segments.append((seg_start, seg_end, seg_info))
                    
                    if overlapping_segments:
                        # Create a canvas for custom drawing
                        cell_canvas = tk.Canvas(self.timetable_inner_frame, width=60, height=45, 
                                               relief=tk.RIDGE, bd=1, highlightthickness=0)
                        cell_canvas.grid(row=row_idx, column=col_idx, sticky=(tk.W, tk.E, tk.N, tk.S))
                        
                        # Draw white background
                        cell_canvas.create_rectangle(0, 0, 60, 45, fill='white', outline='')
                        
                        # Draw each overlapping segment
                        for seg_start, seg_end, seg_info in overlapping_segments:
                            # Calculate overlap within this slot
                            overlap_start = max(seg_start, slot_start_minutes)
                            overlap_end = min(seg_end, slot_end_minutes)
                            
                            # Calculate pixel positions (0-60 for the 30-minute slot)
                            start_pixel = int(((overlap_start - slot_start_minutes) / 30.0) * 60)
                            end_pixel = int(((overlap_end - slot_start_minutes) / 30.0) * 60)
                            
                            # Determine color - red if conflicting, otherwise normal colors
                            is_conflicting = (date_str, seg_start, seg_end) in self.conflicting_segments
                            
                            if is_conflicting:
                                travel_color = '#FF0000'  # Red for conflicts
                            elif seg_info['to_home']:
                                travel_color = '#FFA500'  # Orange
                            elif seg_info.get('from_home', False):
                                travel_color = '#87CEEB'  # Sky blue
                            else:
                                travel_color = '#FFD700'  # Gold
                            
                            # Draw colored rectangle
                            cell_canvas.create_rectangle(start_pixel, 0, end_pixel, 45, 
                                                        fill=travel_color, outline='')
                            
                            # Add text in the slot immediately adjacent to the appointment
                            total_minutes = seg_end - seg_start
                            # For travel FROM home: show in the last slot before segment ends (left of appointment)
                            # For travel TO next/home: show in the first slot where segment starts (right of appointment)
                            show_text = False
                            if seg_info.get('from_home', False):
                                # Show text if this is the last slot before the segment ends
                                if seg_end > slot_start_minutes and seg_end <= slot_end_minutes:
                                    show_text = True
                            else:
                                # Show text if this is the first slot where the segment starts
                                if seg_start >= slot_start_minutes and seg_start < slot_end_minutes:
                                    show_text = True
                            
                            if show_text:
                                cell_canvas.create_text(30, 22, text=f"Travel\n{total_minutes} min", 
                                                      font=('Arial', 8), justify='center')
                        
                        # Bind click event
                        cell_canvas.bind('<Button-1>', lambda e, d=date_str, t=time_slot: self.on_cell_click(d, t))
                        cell_canvas.config(cursor='hand2')
                    else:
                        # Empty cell
                        cell = tk.Label(self.timetable_inner_frame, text="", bg='white',
                                       font=('Arial', 8), width=8, height=3, relief=tk.RIDGE, bd=1,
                                       cursor='hand2', anchor='center', justify='center')
                        cell.grid(row=row_idx, column=col_idx, sticky=(tk.W, tk.E, tk.N, tk.S))
                        cell.bind('<Button-1>', lambda e, d=date_str, t=time_slot: self.on_cell_click(d, t))
        
        # Update scroll region
        self.timetable_inner_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))
    
    def on_cell_click(self, date_str, time_slot):
        """Handle cell click to stage appointment (not confirmed until submit)"""
        cell_key = (date_str, time_slot)
        
        # If cell already has appointment, remove it
        if cell_key in self.appointments:
            postcode = self.appointments[cell_key]
            
            # Check if it's a confirmed appointment
            if postcode in self.confirmed_appointments:
                if messagebox.askyesno("Remove Confirmed Appointment", 
                                       f"This is a confirmed appointment for {postcode}.\nAre you sure you want to remove it?"):
                    # Remove from confirmed appointments
                    del self.confirmed_appointments[postcode]
                    # Remove from CSV
                    df = pd.read_csv(self.appointments_csv)
                    df = df[df['postcode'] != postcode]
                    df.to_csv(self.appointments_csv, index=False)
                    
                    del self.appointments[cell_key]
                    self.recalculate_travel_times(date_str)
                    self.update_timetable()
                    self.update_region_visualization()
                    self.status_label.config(text=f"Removed confirmed appointment: {postcode}", foreground='orange')
                    
                    # Update travel times display
                    if self.postcode_var.get():
                        self.display_travel_times(self.postcode_var.get())
                return
            else:
                # Remove pending appointment
                del self.appointments[cell_key]
                self.pending_appointment = None
                self.pending_label.config(text="")
                self.recalculate_travel_times(date_str)
                self.update_timetable()
                self.status_label.config(text=f"Removed pending appointment: {postcode}", foreground='orange')
                return
        
        # Check if there's already a pending appointment
        if self.pending_appointment:
            pending_date, pending_time, pending_postcode, pending_duration = self.pending_appointment
            response = messagebox.askyesno(
                "Replace Pending Appointment?",
                f"You already have a pending appointment:\n{pending_postcode} on {pending_date} at {pending_time} ({pending_duration} min)\n\nDo you want to replace it with a new selection?\n\n(Submit the current appointment first to keep it)"
            )
            if response:
                # Remove old pending appointment
                old_key = (pending_date, pending_time)
                if old_key in self.appointments:
                    del self.appointments[old_key]
                self.pending_appointment = None
                self.pending_label.config(text="")
                self.recalculate_travel_times(pending_date)
                self.update_timetable()
            else:
                # User chose not to replace, do nothing
                return
        
        # Get selected postcode
        postcode = self.postcode_var.get().strip().upper()
        
        if not postcode:
            messagebox.showwarning("No Postcode Selected", "Please select a postcode first.")
            return
        
        # VALIDATION: Check if this postcode already has a confirmed appointment
        if postcode in self.confirmed_appointments:
            existing_date, existing_time, _, _ = self.confirmed_appointments[postcode]
            messagebox.showerror(
                "Duplicate Location",
                f"Location {postcode} already has a confirmed appointment on {existing_date} at {existing_time}.\n\nOnly 1 appointment per location is allowed.\n\nPlease remove the existing appointment first if you need to reschedule."
            )
            return
        
        # Temporarily add appointment to check for conflicts
        self.appointments[cell_key] = postcode
        self.recalculate_travel_times(date_str)
        
        # Check for conflicts
        conflicts = self.check_travel_conflicts(date_str)
        
        if conflicts:
            conflict_msg = "Note: This appointment creates travel time conflicts:\n\n"
            for conflict in conflicts:
                conflict_msg += f"• {conflict}\n"
            conflict_msg += "\nConflicting travel times are marked in red."
            
            messagebox.showinfo("Travel Time Conflict", conflict_msg, icon='info')
        
        # Stage as pending appointment with current duration setting
        current_duration = int(self.appointment_duration_var.get())
        self.pending_appointment = (date_str, time_slot, postcode, current_duration)
        self.pending_label.config(text=f"Pending: {postcode} on {date_str} at {time_slot} ({current_duration} min)")
        
        # Update display
        self.update_timetable()
        self.update_region_visualization()
        status_msg = f"Staged appointment: {postcode} on {date_str} at {time_slot} (click Submit to confirm)"
        if conflicts:
            status_msg += " (has conflicts)"
        self.status_label.config(text=status_msg, foreground='orange')
    
    def time_to_minutes(self, time_str):
        """Convert time string (HH:MM) to minutes from midnight"""
        hours, mins = map(int, time_str.split(':'))
        return hours * 60 + mins
    
    def check_travel_conflicts(self, date_str):
        """Check for conflicts between travel segments and appointments"""
        conflicts = []
        self.conflicting_segments = set()
        
        # Get all appointments for this date with their time ranges
        appt_ranges = []
        for (d, t), postcode in self.appointments.items():
            if d == date_str:
                start_min = self.time_to_minutes(t)
                # Get actual duration for this appointment
                if postcode in self.confirmed_appointments:
                    _, _, duration, _ = self.confirmed_appointments[postcode]
                else:
                    duration = int(self.appointment_duration_var.get())
                end_min = start_min + duration
                appt_ranges.append((start_min, end_min, t))
        
        # Check each travel segment for conflicts with appointments
        for seg_date, seg_start, seg_end, seg_info in self.travel_segments:
            if seg_date != date_str:
                continue
            
            # Check if travel overlaps with any appointment
            for appt_start, appt_end, appt_time in appt_ranges:
                # Check for overlap: travel and appointment overlap if one starts before the other ends
                if seg_start < appt_end and seg_end > appt_start:
                    # Conflict detected
                    travel_type = "from home" if seg_info.get('from_home') else ("to home" if seg_info['to_home'] else "between appointments")
                    conflicts.append(f"Travel {travel_type} ({seg_info['minutes']} min) overlaps with appointment at {appt_time}")
                    self.conflicting_segments.add((seg_date, seg_start, seg_end))
        
        return conflicts
    
    def recalculate_travel_times(self, date_str):
        """Recalculate travel times for a specific date"""
        # Remove existing travel segments for this date
        self.travel_segments = [seg for seg in self.travel_segments if seg[0] != date_str]
        
        # Get all appointments for this date, sorted by time
        date_appointments = [(k, v) for k, v in self.appointments.items() if k[0] == date_str]
        if not date_appointments:
            return
        
        # Sort by time slot
        date_appointments.sort(key=lambda x: self.time_slots.index(x[0][1]))
        
        # Calculate travel TO first appointment from home
        first_appt = date_appointments[0]
        first_time_minutes = self.time_to_minutes(first_appt[0][1])
        
        if self.home_postcode:
            travel_to_first = self.get_travel_time(self.home_postcode, first_appt[1])
            # Travel starts before the appointment and ends at appointment time
            travel_start = first_time_minutes - travel_to_first
            if travel_start >= self.start_hour * 60:  # Only if it starts after configured start time
                self.travel_segments.append((date_str, travel_start, first_time_minutes, {
                    'minutes': travel_to_first,
                    'to_home': False,
                    'from_home': True
                }))
        
        # Calculate travel between appointments
        for i in range(len(date_appointments) - 1):
            current_appt = date_appointments[i]
            next_appt = date_appointments[i + 1]
            
            current_postcode = current_appt[1]
            # Get actual duration for current appointment
            if current_postcode in self.confirmed_appointments:
                _, _, current_duration, _ = self.confirmed_appointments[current_postcode]
            else:
                current_duration = int(self.appointment_duration_var.get())
            
            current_end_minutes = self.time_to_minutes(current_appt[0][1]) + current_duration
            next_start_minutes = self.time_to_minutes(next_appt[0][1])
            
            # Get travel time
            travel_minutes = self.get_travel_time(current_appt[1], next_appt[1])
            
            # Travel starts after current appointment ends
            travel_end = current_end_minutes + travel_minutes
            
            self.travel_segments.append((date_str, current_end_minutes, travel_end, {
                'minutes': travel_minutes,
                'to_home': False,
                'from_home': False
            }))
        
        # Add travel home after last appointment
        last_appt = date_appointments[-1]
        last_postcode = last_appt[1]
        # Get actual duration for last appointment
        if last_postcode in self.confirmed_appointments:
            _, _, last_duration, _ = self.confirmed_appointments[last_postcode]
        else:
            last_duration = int(self.appointment_duration_var.get())
        
        last_end_minutes = self.time_to_minutes(last_appt[0][1]) + last_duration
        
        # Get actual travel home time
        if self.home_postcode:
            travel_home_minutes = self.get_travel_time(last_appt[1], self.home_postcode)
        else:
            travel_home_minutes = 30  # Default if no home set
        
        travel_home_end = last_end_minutes + travel_home_minutes
        
        if travel_home_end <= self.end_hour * 60:  # Only if it ends before configured end time
            self.travel_segments.append((date_str, last_end_minutes, travel_home_end, {
                'minutes': travel_home_minutes,
                'to_home': True,
                'from_home': False
            }))
    
    def get_travel_time(self, origin, destination):
        """Get travel time between two postcodes"""
        if not origin or not destination or self.distances_df is None:
            return 30  # Default 30 minutes
        
        # Normalize postcodes
        origin = origin.strip().upper()
        destination = destination.strip().upper()
        
        if origin == destination:
            return 0  # No travel time if same location
        
        try:
            # Look up in distances dataframe (check both directions)
            match = self.distances_df[
                ((self.distances_df['origin'] == origin) & (self.distances_df['destination'] == destination)) |
                ((self.distances_df['origin'] == destination) & (self.distances_df['destination'] == origin))
            ]
            
            if not match.empty:
                travel_time = match.iloc[0]['driving_time_minutes']
                # Round up to nearest multiple of 30 for slot allocation
                return max(int(travel_time), 1) if travel_time > 0 else 30
            else:
                print(f"Warning: No distance found for {origin} -> {destination}, using default 30 minutes")
                return 30  # Default if not found
        except Exception as e:
            print(f"Error looking up travel time between {origin} and {destination}: {e}")
            return 30
    
    def display_travel_times(self, postcode):
        """Display travel times from selected postcode to all other postcodes in region"""
        self.suggestions_text.config(state='normal')
        self.suggestions_text.delete('1.0', tk.END)
        
        # Configure text tags for red highlighting
        self.suggestions_text.tag_configure('scheduled', foreground='red', font=('Consolas', 10, 'bold'))
        self.suggestions_text.tag_configure('normal', foreground='black', font=('Consolas', 10))
        self.suggestions_text.tag_configure('header', foreground='blue', font=('Consolas', 10, 'bold'))
        
        if not self.region_postcodes:
            self.suggestions_text.insert('1.0', "No postcodes in selected region.")
            self.suggestions_text.config(state='disabled')
            return
        
        # Get all postcodes except the selected one
        other_postcodes = [pc for pc in self.region_postcodes if pc != postcode]
        
        if not other_postcodes:
            self.suggestions_text.insert('1.0', f"{postcode} is the only postcode in this region.")
            self.suggestions_text.config(state='disabled')
            return
        
        # Calculate travel times and sort by duration
        travel_info = []
        for other_pc in other_postcodes:
            travel_time = self.get_travel_time(postcode, other_pc)
            is_scheduled = other_pc in self.confirmed_appointments
            travel_info.append((travel_time, other_pc, is_scheduled))
        
        # Sort by travel time (ascending)
        travel_info.sort()
        
        # Display header for travel times to other postcodes
        self.suggestions_text.insert(tk.END, f"Travel times from {postcode}:\n", 'header')
        self.suggestions_text.insert(tk.END, f"{'Postcode':<12}{'Time (min)':<12}\n", 'normal')
        self.suggestions_text.insert(tk.END, "-" * 40 + "\n", 'normal')
        
        # Display each postcode with travel time
        for travel_time, other_pc, is_scheduled in travel_info:
            line = f"{other_pc:<12}{travel_time:<12}\n"
            
            if is_scheduled:
                # Highlight in red if already scheduled
                self.suggestions_text.insert(tk.END, line, 'scheduled')
            else:
                self.suggestions_text.insert(tk.END, line, 'normal')
        
        # Add section for travel times to home base
        if self.home_postcode:
            self.suggestions_text.insert(tk.END, f"\nTravel times to {self.home_postcode} (Home):\n", 'header')
            self.suggestions_text.insert(tk.END, f"{'Postcode':<12}{'Time (min)':<12}\n", 'normal')
            self.suggestions_text.insert(tk.END, "-" * 40 + "\n", 'normal')
            
            # Calculate travel times to home for all postcodes
            home_travel_info = []
            for pc in self.region_postcodes:
                travel_time = self.get_travel_time(pc, self.home_postcode)
                is_scheduled = pc in self.confirmed_appointments
                home_travel_info.append((travel_time, pc, is_scheduled))
            
            # Sort by travel time
            home_travel_info.sort()
            
            # Display each postcode with travel time to home
            for travel_time, pc, is_scheduled in home_travel_info:
                line = f"{pc:<12}{travel_time:<12}\n"
                
                if is_scheduled:
                    self.suggestions_text.insert(tk.END, line, 'scheduled')
                else:
                    self.suggestions_text.insert(tk.END, line, 'normal')
        
        self.suggestions_text.config(state='disabled')
    
    def clear_schedule(self):
        """Clear appointments for the currently selected region"""
        if not self.selected_region:
            messagebox.showinfo("No Region Selected", "Please select a region first.")
            return
        
        # Get postcodes in the current region
        region_postcodes_set = set(self.region_postcodes)
        
        # Check if there are any appointments in this region
        region_appointments = {pc: data for pc, data in self.confirmed_appointments.items() if pc in region_postcodes_set}
        region_pending = self.pending_appointment and self.pending_appointment[2] in region_postcodes_set
        
        if not region_appointments and not region_pending:
            messagebox.showinfo("Empty Schedule", "No appointments in this region.")
            return
        
        response = messagebox.askyesno("Clear Region Schedule", 
                                      f"Are you sure you want to clear all appointments for Region {self.selected_region}?")
        if response:
            # Clear appointments for postcodes in this region (appointments dict has (date, time) keys)
            for cell_key in list(self.appointments.keys()):
                if self.appointments[cell_key] in region_postcodes_set:
                    del self.appointments[cell_key]
            
            # Clear confirmed appointments for postcodes in this region
            for postcode in list(region_appointments.keys()):
                del self.confirmed_appointments[postcode]
            
            # Clear pending if it's in this region
            if region_pending:
                self.pending_appointment = None
                self.pending_label.config(text="")
            
            # Clear travel segments for dates in this region
            region_dates = [d.strftime('%d-%b-%y') for d in self.selected_dates]
            self.travel_segments = [seg for seg in self.travel_segments if seg[0] not in region_dates]
            self.conflicting_segments.clear()
            
            # Update CSV - remove appointments for postcodes in this region
            if self.appointments_csv and self.appointments_csv.exists():
                df = pd.read_csv(self.appointments_csv)
                df = df[~df['postcode'].isin(region_postcodes_set)]
                df.to_csv(self.appointments_csv, index=False)
            
            # Update display
            self.update_timetable()
            self.update_region_visualization()
            
            # Update travel times display
            if self.postcode_var.get():
                self.display_travel_times(self.postcode_var.get())
            
            self.status_label.config(text=f"Cleared schedule for Region {self.selected_region}", foreground='orange')
    
    def on_postcode_selected(self, event=None):
        """Handle postcode selection - update travel times display"""
        postcode = self.postcode_var.get()
        if postcode:
            self.display_travel_times(postcode)
            # Also update the map to highlight the selected postcode
            self.update_region_visualization()
    
    def get_region_color_for_postcode(self, postcode):
        """Get the region color code for a given postcode"""
        if self.clustered_regions_df is None:
            return 1  # Default to Red
        
        # Find which region this postcode belongs to
        region_row = self.clustered_regions_df[self.clustered_regions_df['postcode'] == postcode]
        if region_row.empty:
            return 1  # Default to Red
        
        region_num = int(region_row.iloc[0]['region'])
        
        # Get color from region_summary
        if self.regions_df is not None:
            region_data = self.regions_df[self.regions_df['region'] == region_num]
            if not region_data.empty and 'color_code' in region_data.columns:
                return int(region_data.iloc[0]['color_code'])
        
        return 1  # Default to Red
    
    def create_or_update_category(self, outlook, category_name, color_index):
        """Create or update an Outlook category with a specific color"""
        try:
            namespace = outlook.GetNamespace("MAPI")
            categories = namespace.Categories
            
            # Try to get existing category
            try:
                category = categories.Item(category_name)
                category.Color = color_index
            except:
                # Category doesn't exist, create it
                category = categories.Add(category_name, color_index)
        except Exception as e:
            print(f"Error managing category '{category_name}': {e}")
    
    def create_outlook_appointment(self, outlook, postcode, date_str, time_str, duration_minutes, category_name, color_index):
        """Create an Outlook appointment for a confirmed appointment"""
        try:
            # Ensure category exists with correct color
            self.create_or_update_category(outlook, category_name, color_index)
            
            # Parse date and time
            date_obj = datetime.strptime(date_str, "%d-%b-%y")
            time_parts = time_str.split(':')
            hours = int(time_parts[0])
            minutes = int(time_parts[1])
            
            start_datetime = datetime(date_obj.year, date_obj.month, date_obj.day, hours, minutes)
            end_datetime = start_datetime + timedelta(minutes=duration_minutes)
            
            # Create appointment (1 = olAppointmentItem)
            appointment = outlook.CreateItem(1)
            appointment.Subject = postcode  # Use postcode as title
            appointment.Start = start_datetime
            appointment.End = end_datetime
            appointment.AllDayEvent = False
            appointment.BusyStatus = 2  # 2 = olBusy (busy status)
            appointment.Categories = category_name
            appointment.ReminderSet = True
            appointment.ReminderMinutesBeforeStart = 30  # 30 minute reminder
            
            # Add useful info in the body
            appointment.Body = f"Appointment at {postcode}\nDate: {date_str}\nTime: {time_str}\nDuration: {duration_minutes} minutes"
            
            appointment.Save()
            return True
            
        except Exception as e:
            # Re-raise the exception so it can be caught by the caller with full details
            raise Exception(f"Error creating appointment for {postcode}: {str(e)}")
    
    def sync_to_outlook(self):
        """Sync all appointments that aren't yet in Outlook"""
        if not self.confirmed_appointments:
            messagebox.showinfo("No Appointments", "No confirmed appointments to sync.")
            return
        
        # Count how many need syncing
        to_sync = [(pc, data) for pc, data in self.confirmed_appointments.items() if not data[3]]  # data[3] is in_outlook
        
        if not to_sync:
            messagebox.showinfo("Already Synced", "All appointments are already in Outlook!")
            return
        
        # Confirm with user
        response = messagebox.askyesno(
            "Sync to Outlook",
            f"Found {len(to_sync)} appointment(s) not yet in Outlook.\n\nDo you want to create Outlook events for these appointments?"
        )
        
        if not response:
            return
        
        try:
            # Connect to Outlook - try active object first, then dispatch
            try:
                outlook = win32com.client.GetActiveObject("Outlook.Application")
            except:
                outlook = win32com.client.Dispatch("Outlook.Application")
                time.sleep(1)
            
            created_count = 0
            failed = []
            
            for postcode, (date, time, duration, in_outlook) in to_sync:
                try:
                    # Get region color for this postcode
                    color_code = self.get_region_color_for_postcode(postcode)
                    color_name = OUTLOOK_COLORS.get(color_code, "Red")
                    category_name = f"Appointment - {color_name}"
                    
                    # Create Outlook appointment
                    if self.create_outlook_appointment(outlook, postcode, date, time, duration, category_name, color_code):
                        created_count += 1
                        # Update in memory
                        self.confirmed_appointments[postcode] = (date, time, duration, True)
                    else:
                        failed.append(postcode)
                except Exception as e:
                    failed.append(f"{postcode} ({str(e)})")
                    print(f"Error syncing {postcode}: {e}")
            
            # Update CSV with in_outlook flag
            df = pd.read_csv(self.appointments_csv)
            for postcode in [pc for pc, _ in to_sync if pc not in failed]:
                df.loc[df['postcode'] == postcode, 'in_outlook'] = True
            df.to_csv(self.appointments_csv, index=False)
            
            # Show results
            if created_count > 0:
                msg = f"Successfully synced {created_count} appointment(s) to Outlook!"
                if failed:
                    msg += f"\n\nFailed to sync {len(failed)} appointment(s):\n" + "\n".join(failed)
                messagebox.showinfo("Sync Complete", msg)
            else:
                error_details = "\n".join(failed) if failed else "Unknown error"
                messagebox.showerror("Sync Failed", f"Failed to sync appointments to Outlook.\n\nDetails:\n{error_details}")
                
        except Exception as e:
            import traceback
            error_trace = traceback.format_exc()
            messagebox.showerror("Outlook Error", f"Failed to connect to Outlook:\n\n{e}\n\nDetails:\n{error_trace}")
    
    def load_confirmed_appointments(self):
        """Load confirmed appointments from CSV"""
        if not self.appointments_csv.exists():
            # Create empty CSV with headers
            df = pd.DataFrame(columns=['postcode', 'date', 'time', 'duration', 'in_outlook'])
            df.to_csv(self.appointments_csv, index=False)
            return
        
        df = pd.read_csv(self.appointments_csv)
        self.confirmed_appointments = {}
        
        for _, row in df.iterrows():
            postcode = row['postcode']
            date = row['date']
            time = row['time']
            # Default to 60 minutes if duration column doesn't exist (backward compatibility)
            duration = int(row['duration']) if 'duration' in row and pd.notna(row['duration']) else 60
            # Track if appointment is in Outlook (default to False for backward compatibility)
            in_outlook = bool(row['in_outlook']) if 'in_outlook' in row and pd.notna(row['in_outlook']) else False
            self.confirmed_appointments[postcode] = (date, time, duration, in_outlook)
        
        # Also add to visual appointments dict and recalculate travel
        for postcode, (date, time, duration, in_outlook) in self.confirmed_appointments.items():
            self.appointments[(date, time)] = postcode
            self.recalculate_travel_times(date)
        
        # Update timetable display if we have selected dates
        if self.selected_dates:
            self.update_timetable()
        
        # Update map visualization to show routes
        if self.selected_region is not None:
            self.update_region_visualization()
    
    def submit_appointment(self):
        """Submit the pending appointment after validation"""
        if not self.pending_appointment:
            messagebox.showinfo("No Appointment", "No appointment selected to submit.")
            return
        
        date, time, postcode, duration = self.pending_appointment
        
        # Validation: Check if this postcode already has a confirmed appointment
        if postcode in self.confirmed_appointments:
            existing_date, existing_time, existing_duration, _ = self.confirmed_appointments[postcode]
            messagebox.showerror(
                "Duplicate Location", 
                f"Location {postcode} already has a confirmed appointment on {existing_date} at {existing_time}.\\n\\nOnly 1 appointment per location is allowed."
            )
            return
        
        # Show custom dialog with Outlook checkbox
        add_to_outlook = self.show_submit_dialog(postcode, date, time, duration)
        
        if add_to_outlook is None:
            # User cancelled
            return
        
        # Create Outlook appointment if requested
        outlook_success = False
        if add_to_outlook:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                color_code = self.get_region_color_for_postcode(postcode)
                color_name = OUTLOOK_COLORS.get(color_code, "Red")
                category_name = f"Appointment - {color_name}"
                outlook_success = self.create_outlook_appointment(outlook, postcode, date, time, duration, category_name, color_code)
            except Exception as e:
                messagebox.showerror("Outlook Error", f"Failed to create Outlook appointment:\\n{e}")
                outlook_success = False
        
        # Save to confirmed appointments (with outlook status)
        self.confirmed_appointments[postcode] = (date, time, duration, outlook_success if add_to_outlook else False)
        
        # Add to CSV
        df = pd.read_csv(self.appointments_csv)
        new_row = pd.DataFrame([{
            'postcode': postcode, 
            'date': date, 
            'time': time, 
            'duration': duration,
            'in_outlook': outlook_success if add_to_outlook else False
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_csv(self.appointments_csv, index=False)
        
        # Clear pending
        self.pending_appointment = None
        self.pending_label.config(text="")
        
        # Update displays
        self.update_timetable()
        self.update_region_visualization()
        if self.postcode_var.get():
            self.display_travel_times(self.postcode_var.get())
        
        # Update status
        outlook_msg = " (added to Outlook)" if outlook_success else " (Outlook sync skipped)" if not add_to_outlook else " (Outlook failed)"
        self.status_label.config(text=f"Appointment confirmed: {postcode} on {date} at {time} ({duration} min){outlook_msg}", 
                                foreground='green')
    
    def show_submit_dialog(self, postcode, date, time, duration):
        """Show custom dialog for appointment submission with Outlook checkbox
        Returns: True if add to Outlook, False if don't add, None if cancelled"""
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Confirm Appointment")
        dialog.geometry("450x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        result = [None]  # Use list to allow modification in nested function
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(main_frame, text="Confirm Appointment", font=('Arial', 14, 'bold')).pack(pady=(0, 15))
        
        # Appointment details
        details_frame = ttk.LabelFrame(main_frame, text="Appointment Details", padding="10")
        details_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(details_frame, text=f"Location: {postcode}", font=('Arial', 10)).pack(anchor=tk.W, pady=2)
        ttk.Label(details_frame, text=f"Date: {date}", font=('Arial', 10)).pack(anchor=tk.W, pady=2)
        ttk.Label(details_frame, text=f"Time: {time}", font=('Arial', 10)).pack(anchor=tk.W, pady=2)
        ttk.Label(details_frame, text=f"Duration: {duration} minutes", font=('Arial', 10)).pack(anchor=tk.W, pady=2)
        
        # Outlook checkbox
        outlook_var = tk.BooleanVar(value=True)  # Default to checked
        outlook_check = ttk.Checkbutton(
            main_frame, 
            text="Add appointment to Outlook Calendar",
            variable=outlook_var
        )
        outlook_check.pack(pady=(0, 15))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        def on_confirm():
            result[0] = outlook_var.get()
            dialog.destroy()
        
        def on_cancel():
            result[0] = None
            dialog.destroy()
        
        ttk.Button(button_frame, text="Confirm", command=on_confirm, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=on_cancel, width=15).pack(side=tk.LEFT, padx=5)
        
        # Wait for dialog to close
        dialog.wait_window()
        
        return result[0]


def main():
    # Get project directory from command line argument if provided
    project_dir = sys.argv[1] if len(sys.argv) > 1 else None
    
    root = tk.Tk()
    app = SmartSchedulerApp(root, project_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
