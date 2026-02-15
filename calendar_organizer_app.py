import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import pandas as pd
import calendar
from datetime import datetime, timedelta
import win32com.client
import time

# Import display preferences
try:
    from display_preferences import (
        initialize as init_display_prefs,
        get_show_names,
        set_show_names,
        register_callback
    )
    DISPLAY_PREFS_AVAILABLE = True
except ImportError:
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


class CalendarOrganizerApp:
    def __init__(self, root, project_dir=None):
        self.root = root
        self.root.title("Calendar Organizer")
        self.root.geometry("1400x900")
        
        # Project directory from command line
        self.project_dir = project_dir
        
        # Data variables
        self.regions = []
        self.region_data = {}  # {region_number: {'name': str, 'postcodes': list, 'count': int}}
        self.region_assignments = {}  # {date_str: region_number}
        self.selected_region = None
        self.schedule_saved = False  # Track if schedule has been saved
        
        # Calendar variables
        self.current_month = datetime.now().month
        self.current_year = datetime.now().year
        
        # Initialize display preferences
        if DISPLAY_PREFS_AVAILABLE:
            try:
                init_display_prefs(self.project_dir if self.project_dir else os.getcwd())
                register_callback(self.on_display_preference_changed)
            except Exception as e:
                print(f"Warning: Could not initialize display preferences: {e}")
        
        self.setup_ui()
        
        # Auto-load project if provided
        if self.project_dir:
            self.auto_load_project()
        
        # Set close protocol
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def toggle_display_preference(self):
        """Toggle between showing names and postcodes"""
        try:
            current = get_show_names()
            set_show_names(not current)
            self.update_toggle_button_text()
            self.refresh_postcodes_display()
        except:
            pass
    
    def update_toggle_button_text(self):
        """Update toggle button text based on current preference"""
        if hasattr(self, 'toggle_btn'):
            try:
                if get_show_names():
                    self.toggle_btn.config(text="Show Postcodes")
                else:
                    self.toggle_btn.config(text="Show Names")
            except:
                self.toggle_btn.config(text="Display Mode")
    
    def on_display_preference_changed(self, show_names):
        """Callback when display preference changes from another app"""
        self.update_toggle_button_text()
        self.refresh_postcodes_display()
    
    def refresh_postcodes_display(self):
        """Refresh the postcodes display with current preference"""
        if self.selected_region and self.selected_region in self.region_data:
            self.postcodes_text.config(state=tk.NORMAL)
            self.postcodes_text.delete('1.0', tk.END)
            
            region_info = self.region_data[self.selected_region]
            
            # Get display format
            if DISPLAY_PREFS_AVAILABLE and get_show_names() and 'client_names' in region_info:
                # Show names if available
                display_items = []
                for i, postcode in enumerate(region_info['postcodes']):
                    if i < len(region_info['client_names']) and region_info['client_names'][i]:
                        display_items.append(region_info['client_names'][i])
                    else:
                        display_items.append(postcode)
                display_str = ', '.join(display_items)
            else:
                # Show postcodes
                display_str = ', '.join(region_info['postcodes'])
            
            self.postcodes_text.insert('1.0', display_str)
            self.postcodes_text.config(state=tk.DISABLED)
    
    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="5")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Button bar at top
        button_bar = ttk.Frame(main_frame)
        button_bar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        
        ttk.Button(button_bar, text="File", command=self.show_file_menu, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_bar, text="Save Schedule", command=self.save_schedule, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_bar, text="Reload Schedule", command=self.load_schedule, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_bar, text="Export to Outlook", command=self.export_to_outlook, width=18).pack(side=tk.LEFT, padx=2)
        
        # Add toggle button on the right
        self.toggle_btn = ttk.Button(button_bar, text="Show Postcodes", 
                                    command=self.toggle_display_preference, width=18)
        self.toggle_btn.pack(side=tk.RIGHT, padx=(10, 0))
        self.update_toggle_button_text()
        
        # Left panel - Region selection
        left_panel = ttk.LabelFrame(main_frame, text="Regions", padding="10")
        left_panel.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        left_panel.rowconfigure(1, weight=1)
        
        ttk.Label(left_panel, text="Select region to assign to dates:", 
                 font=('Arial', 10)).grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Region listbox
        self.region_listbox = tk.Listbox(left_panel, font=('Arial', 10), height=20)
        self.region_listbox.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.region_listbox.bind('<<ListboxSelect>>', self.on_region_selected)
        
        scrollbar = ttk.Scrollbar(left_panel, orient=tk.VERTICAL, command=self.region_listbox.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.region_listbox.config(yscrollcommand=scrollbar.set)
        
        # Selected region info
        self.selected_region_label = ttk.Label(left_panel, text="No region selected", 
                                               font=('Arial', 9, 'italic'),
                                               foreground='gray')
        self.selected_region_label.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        # Postcodes display
        ttk.Label(left_panel, text="Locations in selected region:", 
                 font=('Arial', 9)).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(5, 5))
        
        self.postcodes_text = tk.Text(left_panel, height=8, width=30, font=('Consolas', 8),
                                      wrap=tk.WORD, state=tk.DISABLED)
        self.postcodes_text.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 5))
        left_panel.rowconfigure(4, weight=1)
        
        postcodes_scroll = ttk.Scrollbar(left_panel, orient=tk.VERTICAL, command=self.postcodes_text.yview)
        self.postcodes_text.config(yscrollcommand=postcodes_scroll.set)
        
        # Right panel - Calendar
        right_panel = ttk.LabelFrame(main_frame, text="Calendar", padding="10")
        right_panel.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(1, weight=1)
        
        # Calendar controls
        controls_frame = ttk.Frame(right_panel)
        controls_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(controls_frame, text="<", command=self.prev_month, width=3).pack(side=tk.LEFT, padx=2)
        
        self.month_var = tk.StringVar()
        self.month_combo = ttk.Combobox(controls_frame, textvariable=self.month_var, 
                                       state='readonly', width=12)
        self.month_combo['values'] = [calendar.month_name[i] for i in range(1, 13)]
        self.month_combo.pack(side=tk.LEFT, padx=5)
        self.month_combo.bind('<<ComboboxSelected>>', self.on_month_changed)
        
        self.year_var = tk.StringVar()
        self.year_spinbox = ttk.Spinbox(controls_frame, from_=2020, to=2030, 
                                       textvariable=self.year_var, width=8)
        self.year_spinbox.pack(side=tk.LEFT, padx=5)
        self.year_spinbox.bind('<Return>', self.on_year_changed)
        self.year_spinbox.bind('<FocusOut>', self.on_year_changed)
        
        ttk.Button(controls_frame, text=">", command=self.next_month, width=3).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(controls_frame, text="Today", command=self.go_to_today, width=8).pack(side=tk.LEFT, padx=10)
        
        # Calendar grid
        self.calendar_frame = ttk.Frame(right_panel)
        self.calendar_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Initialize calendar
        self.update_calendar_display()
        
        # Progress bar at bottom
        progress_frame = ttk.Frame(main_frame, relief=tk.SUNKEN, borderwidth=1)
        progress_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        
        self.status_label = ttk.Label(progress_frame, text="Ready", 
                                     foreground="green", width=30)
        self.status_label.pack(side=tk.RIGHT, padx=5, pady=3)
    
    def auto_load_project(self):
        """Auto-load project information"""
        if not self.project_dir or not os.path.exists(self.project_dir):
            return
        
        project_name = os.path.basename(self.project_dir)
        self.root.title(f"Calendar Organizer - Project: {project_name}")
        
        # Load regions from clustered_regions.csv
        self.load_regions()
        
        # Load existing schedule if available
        self.load_schedule()
    
    def load_regions(self):
        """Load regions from clustered_regions.csv"""
        if not self.project_dir:
            return
        
        clustered_file = os.path.join(self.project_dir, "clustered_regions.csv")
        if not os.path.exists(clustered_file):
            messagebox.showwarning("No Regions", 
                                  "clustered_regions.csv not found.\n\n"
                                  "Please run TSP Clustering Optimizer first.")
            return
        
        try:
            df = pd.read_csv(clustered_file)
            # Get unique regions (excluding depot which is region 0 or -1)
            regions_df = df[(df['region'] > 0) & (df['region'] != -1)]
            unique_regions = sorted(regions_df['region'].unique())
            
            # Load custom region names if available
            region_names = {}
            region_colors = {}  # Store color codes for calendar appointments
            names_file = os.path.join(self.project_dir, "region_names.csv")
            if os.path.exists(names_file):
                names_df = pd.read_csv(names_file)
                for _, row in names_df.iterrows():
                    region_num = int(row['region'])
                    region_names[region_num] = row['name']
                    # Load color code if available
                    if 'color_code' in names_df.columns:
                        region_colors[region_num] = int(row['color_code'])
            
            # Load minimum days from region_summary.csv
            region_min_days = {}
            summary_file = os.path.join(self.project_dir, "region_summary.csv")
            if os.path.exists(summary_file):
                summary_df = pd.read_csv(summary_file)
                if 'minimum_days' in summary_df.columns:
                    for _, row in summary_df.iterrows():
                        if row['region'] != 'Excluded':
                            try:
                                region_num = int(row['region'])
                                region_min_days[region_num] = int(row['minimum_days'])
                            except:
                                pass
            
            self.regions = []
            self.region_data = {}
            self.region_listbox.delete(0, tk.END)
            
            for region_num in unique_regions:
                region_customers = regions_df[regions_df['region'] == region_num]
                customer_count = len(region_customers)
                postcodes = sorted(region_customers['postcode'].tolist())
                
                # Get client names if available
                client_names = []
                if 'client_name' in region_customers.columns:
                    for pc in postcodes:
                        pc_row = region_customers[region_customers['postcode'] == pc]
                        if len(pc_row) > 0:
                            client_name = pc_row.iloc[0]['client_name']
                            if client_name and pd.notna(client_name):
                                client_names.append(str(client_name).strip())
                            else:
                                client_names.append(None)
                        else:
                            client_names.append(None)
                
                # Get custom name or default
                region_name = region_names.get(region_num, f"Region {region_num}")
                region_color = region_colors.get(region_num, 1)  # Default to Red (1)
                min_days = region_min_days.get(region_num, 0)  # Get minimum days
                
                self.regions.append(region_num)
                self.region_data[region_num] = {
                    'name': region_name,
                    'postcodes': postcodes,
                    'client_names': client_names if client_names else [None] * len(postcodes),
                    'count': customer_count,
                    'color_code': region_color,  # Store color code for calendar appointments
                    'minimum_days': min_days  # Store minimum days
                }
                
                # Display with minimum days info if available
                if min_days > 0:
                    self.region_listbox.insert(tk.END, f"{region_name} ({customer_count}) - Min: {min_days} days")
                else:
                    self.region_listbox.insert(tk.END, f"{region_name} ({customer_count})")
            
            self.status_label.config(text=f"Loaded {len(self.regions)} regions", foreground="green")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load regions:\n{e}")
    
    def on_region_selected(self, event):
        """Handle region selection"""
        selection = self.region_listbox.curselection()
        if selection:
            index = selection[0]
            self.selected_region = self.regions[index]
            
            # Get region data
            region_info = self.region_data[self.selected_region]
            region_name = region_info['name']
            minimum_days = region_info.get('minimum_days', 0)
            
            # Count currently assigned days for this region
            assigned_days = sum(1 for r in self.region_assignments.values() if r == self.selected_region)
            
            # Show selection with day count info
            if minimum_days > 0:
                if assigned_days < minimum_days:
                    self.selected_region_label.config(
                        text=f"Selected: {region_name} (Click dates to assign) - ⚠️ {assigned_days}/{minimum_days} days",
                        foreground="orange", font=('Arial', 9, 'bold')
                    )
                else:
                    self.selected_region_label.config(
                        text=f"Selected: {region_name} (Click dates to assign) - ✓ {assigned_days}/{minimum_days} days",
                        foreground="green", font=('Arial', 9, 'bold')
                    )
            else:
                self.selected_region_label.config(
                    text=f"Selected: {region_name} (Click dates to assign)",
                    foreground="blue", font=('Arial', 9, 'bold')
                )
            
            # Use the refresh method to display postcodes
            self.refresh_postcodes_display()
    
    def update_calendar_display(self):
        """Update the calendar display for current month/year"""
        # Clear existing calendar
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        
        # Set month/year controls
        self.month_var.set(calendar.month_name[self.current_month])
        self.year_var.set(str(self.current_year))
        
        # Get calendar for the month
        cal = calendar.monthcalendar(self.current_year, self.current_month)
        
        # Day headers
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        for col, day in enumerate(days):
            label = ttk.Label(self.calendar_frame, text=day, font=('Arial', 10, 'bold'),
                            anchor=tk.CENTER)
            label.grid(row=0, column=col, sticky=(tk.W, tk.E, tk.N, tk.S), padx=1, pady=1)
        
        # Calendar days
        for week_num, week in enumerate(cal):
            for day_num, day in enumerate(week):
                if day == 0:
                    # Empty cell
                    frame = ttk.Frame(self.calendar_frame, relief=tk.FLAT)
                else:
                    # Create date button
                    date_str = f"{self.current_year}-{self.current_month:02d}-{day:02d}"
                    
                    # Check if date has assignment
                    bg_color = 'white'
                    text_color = 'black'
                    assignment_text = ''
                    
                    if date_str in self.region_assignments:
                        region = self.region_assignments[date_str]
                        assignment_text = f"\nR{region}"
                        bg_color = self.get_region_color(region)
                        text_color = 'white'
                    
                    # Create button
                    btn = tk.Button(self.calendar_frame, text=f"{day}{assignment_text}",
                                  font=('Arial', 9),
                                  width=8, height=3,
                                  bg=bg_color, fg=text_color,
                                  relief=tk.RAISED,
                                  command=lambda d=date_str: self.on_date_clicked(d))
                    frame = btn
                
                frame.grid(row=week_num+1, column=day_num, sticky=(tk.W, tk.E, tk.N, tk.S), 
                          padx=1, pady=1)
        
        # Configure grid weights
        for i in range(7):
            self.calendar_frame.columnconfigure(i, weight=1)
        for i in range(len(cal)+1):
            self.calendar_frame.rowconfigure(i, weight=1)
    
    def get_region_color(self, region):
        """Get color for a region based on Outlook color code"""
        # Get color code from region data
        if region in self.region_data:
            color_code = self.region_data[region].get('color_code', 1)
            return self.outlook_color_to_matplotlib(color_code)
        # Default to Red if region not found
        return '#DC143C'
    
    def outlook_color_to_matplotlib(self, color_code):
        """Convert Outlook color code to RGB hex color for matplotlib/tkinter"""
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
    
    def on_date_clicked(self, date_str):
        """Handle date click"""
        if self.selected_region is None:
            messagebox.showinfo("No Region Selected", 
                              "Please select a region from the left panel first.")
            return
        
        # Toggle assignment
        if date_str in self.region_assignments:
            if self.region_assignments[date_str] == self.selected_region:
                # Unassign if clicking same region
                del self.region_assignments[date_str]
                self.status_label.config(text=f"Removed assignment from {date_str}", 
                                       foreground="orange")
            else:
                # Reassign to new region
                old_region = self.region_assignments[date_str]
                self.region_assignments[date_str] = self.selected_region
                self.status_label.config(
                    text=f"Reassigned {date_str} from Region {old_region} to Region {self.selected_region}", 
                    foreground="blue")
        else:
            # Assign to selected region
            self.region_assignments[date_str] = self.selected_region
            self.status_label.config(text=f"Assigned {date_str} to Region {self.selected_region}", 
                                   foreground="green")
        
        # Mark schedule as modified (needs to be saved before export)
        self.schedule_saved = False
        
        # Refresh calendar
        self.update_calendar_display()
        
        # Update region selection display to show new day count
        if self.selected_region:
            # Trigger region selection update to refresh day counts
            current_selection = self.region_listbox.curselection()
            if current_selection:
                self.on_region_selected(None)
    
    def prev_month(self):
        """Go to previous month"""
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.update_calendar_display()
    
    def next_month(self):
        """Go to next month"""
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.update_calendar_display()
    
    def go_to_today(self):
        """Go to current month"""
        today = datetime.now()
        self.current_month = today.month
        self.current_year = today.year
        self.update_calendar_display()
    
    def on_month_changed(self, event):
        """Handle month selection"""
        month_name = self.month_var.get()
        self.current_month = list(calendar.month_name).index(month_name)
        self.update_calendar_display()
    
    def on_year_changed(self, event):
        """Handle year change"""
        try:
            self.current_year = int(self.year_var.get())
            self.update_calendar_display()
        except ValueError:
            pass
    
    def check_minimum_days_constraint(self):
        """Check if any regions have fewer days assigned than minimum
        Returns list of (region_num, assigned_days, minimum_days) for regions below minimum"""
        warnings = []
        
        # Count days assigned to each region
        region_day_counts = {}
        for date_str, region in self.region_assignments.items():
            region_day_counts[region] = region_day_counts.get(region, 0) + 1
        
        # Check each region against its minimum
        for region_num in self.regions:
            if region_num in self.region_data:
                minimum_days = self.region_data[region_num].get('minimum_days', 0)
                if minimum_days > 0:
                    assigned_days = region_day_counts.get(region_num, 0)
                    if assigned_days < minimum_days:
                        warnings.append((region_num, assigned_days, minimum_days))
        
        return warnings
    
    def save_schedule(self):
        """Save schedule to CSV"""
        if not self.project_dir:
            messagebox.showwarning("No Project", "No project loaded.")
            return
        
        if not self.region_assignments:
            messagebox.showwarning("No Assignments", "No region assignments to save.")
            return
        
        # Check if any regions have fewer days than minimum
        warnings = self.check_minimum_days_constraint()
        
        if warnings:
            warning_message = "⚠️ The following regions have fewer days assigned than recommended:\n\n"
            for region_num, assigned, minimum in warnings:
                region_name = self.region_data[region_num]['name']
                warning_message += f"• {region_name}: {assigned} day(s) assigned (recommended: {minimum})\n"
            
            warning_message += "\n📌 You can still save, but consider adding more days for better coverage.\n\nDo you want to save anyway?"
            
            response = messagebox.askyesno("Minimum Days Warning", warning_message, icon='warning')
            if not response:
                return
        
        try:
            # Create DataFrame
            data = []
            for date_str, region in sorted(self.region_assignments.items()):
                data.append({'date': date_str, 'region': region})
            
            df = pd.DataFrame(data)
            
            # Save to CSV
            schedule_file = os.path.join(self.project_dir, "region_schedule.csv")
            df.to_csv(schedule_file, index=False)
            
            # Mark that schedule has been saved
            self.schedule_saved = True
            
            self.status_label.config(text=f"Schedule saved to {os.path.basename(schedule_file)}", 
                                   foreground="green")
            messagebox.showinfo("Success", 
                              f"Schedule saved successfully!\n\n"
                              f"File: region_schedule.csv\n"
                              f"Assignments: {len(self.region_assignments)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save schedule:\n{e}")
    
    def load_schedule(self):
        """Load schedule from CSV"""
        if not self.project_dir:
            return
        
        schedule_file = os.path.join(self.project_dir, "region_schedule.csv")
        if not os.path.exists(schedule_file):
            return
        
        try:
            df = pd.read_csv(schedule_file)
            self.region_assignments = {}
            
            for _, row in df.iterrows():
                self.region_assignments[row['date']] = int(row['region'])
            
            # Mark schedule as saved since we loaded it from file
            self.schedule_saved = True
            
            self.status_label.config(text=f"Loaded {len(self.region_assignments)} assignments", 
                                   foreground="green")
            self.update_calendar_display()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load schedule:\n{e}")
    
    def show_file_menu(self):
        """Show file menu dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("File Options")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="File Operations", font=('Arial', 14, 'bold')).pack(pady=(0, 20))
        
        ttk.Button(frame, text="Reload Regions from CSV", command=self.load_regions,
                  width=30).pack(pady=5)
        ttk.Button(frame, text="Clear All Assignments", command=self.clear_assignments,
                  width=30).pack(pady=5)
    
    def clear_assignments(self):
        """Clear all region assignments"""
        if not self.region_assignments:
            messagebox.showinfo("No Assignments", "No assignments to clear.")
            return
        
        response = messagebox.askyesno("Clear Assignments", 
                                       f"Clear all {len(self.region_assignments)} assignments?\n\n"
                                       "This cannot be undone unless you reload the saved schedule.")
        if response:
            self.region_assignments = {}
            self.update_calendar_display()
            self.status_label.config(text="All assignments cleared", foreground="orange")
    
    def get_region_color_info(self, region_num):
        """Get color code and name for a region"""
        if region_num in self.region_data:
            color_code = self.region_data[region_num].get('color_code', 1)
            color_name = OUTLOOK_COLORS.get(color_code, "Red")
            return color_code, color_name
        return 1, "Red"  # Default to Red
    
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
    
    def create_appointment(self, outlook, subject, start_time, category_name, color_index):
        """Create an Outlook appointment for a region assignment"""
        try:
            # Ensure category exists with correct color
            self.create_or_update_category(outlook, category_name, color_index)
            
            # Create appointment (1 = olAppointmentItem)
            appointment = outlook.CreateItem(1)
            appointment.Subject = subject
            appointment.Start = start_time
            appointment.AllDayEvent = True
            appointment.BusyStatus = 0  # 0 = Free
            appointment.Categories = category_name
            appointment.ReminderSet = False  # No reminder
            appointment.Save()
            
            return appointment
        except Exception as e:
            print(f"Error creating appointment: {e}")
            return None
    
    def export_to_outlook(self):
        """Export all scheduled region assignments to Outlook calendar"""
        if not self.region_assignments:
            messagebox.showinfo("No Assignments", 
                              "No region assignments to export.\n\n"
                              "Please assign regions to dates first.")
            return
        
        # Check if schedule has been saved
        if not self.schedule_saved:
            messagebox.showwarning("Schedule Not Saved", 
                                  "⚠️ You must save the schedule before exporting to Outlook!\n\n"
                                  "Click 'Save Schedule' first to save your assignments.")
            return
        
        try:
            # Connect to Outlook
            try:
                outlook = win32com.client.GetActiveObject("Outlook.Application")
            except:
                outlook = win32com.client.Dispatch("Outlook.Application")
                time.sleep(1)
            
            created_count = 0
            failed_count = 0
            
            # Create appointments for each assignment
            for date_str, region_num in self.region_assignments.items():
                # Get region info
                region_name = self.region_data[region_num]['name']
                color_code, color_name = self.get_region_color_info(region_num)
                
                # Create appointment
                appointment = self.create_appointment(
                    outlook=outlook,
                    subject=region_name,
                    start_time=date_str,
                    category_name=region_name,
                    color_index=color_code
                )
                
                if appointment:
                    created_count += 1
                else:
                    failed_count += 1
            
            # Show result
            if failed_count == 0:
                messagebox.showinfo("Success", 
                                  f"Successfully created {created_count} appointments in Outlook!\n\n"
                                  f"All appointments marked as Free (time available).")
                self.status_label.config(text=f"Exported {created_count} appointments to Outlook", 
                                       foreground="green")
            else:
                messagebox.showwarning("Partially Complete", 
                                      f"Created {created_count} appointments\n"
                                      f"Failed: {failed_count} appointments")
                self.status_label.config(text=f"Exported {created_count} appointments ({failed_count} failed)", 
                                       foreground="orange")
        
        except Exception as e:
            messagebox.showerror("Error", 
                               f"Failed to export to Outlook:\n{e}\n\n"
                               f"Make sure Outlook is installed and accessible.")
            self.status_label.config(text="Outlook export failed", foreground="red")
    
    def on_closing(self):
        """Handle window close event"""
        self.root.destroy()


def main():
    # Get project directory from command line if provided
    project_dir = None
    if len(sys.argv) > 1:
        project_dir = sys.argv[1]
        if not os.path.exists(project_dir):
            print(f"Warning: Project directory does not exist: {project_dir}")
            project_dir = None
    
    root = tk.Tk()
    app = CalendarOrganizerApp(root, project_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
