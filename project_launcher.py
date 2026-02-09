import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import json
import subprocess
import sys
import shutil
import time
import stat
import pandas as pd
import threading

# Import the other apps for single-EXE compatibility
APPS_IMPORTED = False
PostcodeDistanceApp = None
TSPClusteringApp = None
CalendarOrganizerApp = None
SmartSchedulerApp = None

try:
    print("Attempting to import postcode_distance_app...")
    from postcode_distance_app import PostcodeDistanceApp
    print("✓ postcode_distance_app imported")
except Exception as e:
    print(f"✗ Failed to import postcode_distance_app: {e}")

try:
    print("Attempting to import tsp_clustering_app...")
    from tsp_clustering_app import TSPClusteringApp
    print("✓ tsp_clustering_app imported")
except Exception as e:
    print(f"✗ Failed to import tsp_clustering_app: {e}")

try:
    print("Attempting to import calendar_organizer_app...")
    from calendar_organizer_app import CalendarOrganizerApp
    print("✓ calendar_organizer_app imported")
except Exception as e:
    print(f"✗ Failed to import calendar_organizer_app: {e}")

try:
    print("Attempting to import smart_scheduler_app...")
    from smart_scheduler_app import SmartSchedulerApp
    print("✓ smart_scheduler_app imported")
except Exception as e:
    print(f"✗ Failed to import smart_scheduler_app: {e}")

# Check if all imports succeeded
if all([PostcodeDistanceApp, TSPClusteringApp, CalendarOrganizerApp, SmartSchedulerApp]):
    APPS_IMPORTED = True
    print("SUCCESS: All apps imported successfully")
else:
    print(f"WARNING: Some apps failed to import. APPS_IMPORTED=False")


class ProjectLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("TSP Project Launcher")
        
        # Set to fullscreen
        self.root.state('zoomed')  # Windows fullscreen (maximized)
        
        # Get the directory where this script/exe is located
        # Use sys.executable for frozen (EXE) mode, __file__ for script mode
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            self.app_directory = os.path.dirname(sys.executable)
        else:
            # Running as script
            self.app_directory = os.path.dirname(os.path.abspath(__file__))
        
        self.config_file = os.path.join(self.app_directory, "launcher_config.txt")
        
        # Load configuration
        self.config = self.load_config()
        
        # Ensure projects directory exists
        if not os.path.exists(self.config['projects_directory']):
            os.makedirs(self.config['projects_directory'])
        
        self.setup_ui()
        self.refresh_projects_list()
        
    def load_config(self):
        """Load configuration from file or create default"""
        default_config = {
            'projects_directory': os.path.join(self.app_directory, 'Projects'),
            'active_project': None,
            'recent_projects': []
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    # Merge with defaults to ensure all keys exist
                    for key in default_config:
                        if key not in config:
                            config[key] = default_config[key]
                    return config
            except Exception as e:
                print(f"Error loading config: {e}")
                return default_config
        else:
            return default_config
    
    def save_config(self):
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration:\n{e}")
    
    def show_launching_notification(self, app_name):
        """Show a temporary 'Launching...' notification that auto-closes"""
        notification = tk.Toplevel(self.root)
        notification.title("Launching")
        notification.geometry("350x100")
        notification.resizable(False, False)
        
        # Center the notification
        notification.update_idletasks()
        x = (notification.winfo_screenwidth() // 2) - (350 // 2)
        y = (notification.winfo_screenheight() // 2) - (100 // 2)
        notification.geometry(f"350x100+{x}+{y}")
        
        # Make it stay on top
        notification.attributes('-topmost', True)
        
        # Message
        frame = ttk.Frame(notification, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=f"Launching {app_name}...", 
                 font=('Arial', 12)).pack(pady=10)
        
        progress = ttk.Progressbar(frame, mode='indeterminate', length=300)
        progress.pack(pady=10)
        progress.start(10)
        
        # Auto-close after 1 second
        notification.after(1000, notification.destroy)
        
        return notification
    
    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="TSP Project Launcher", 
                               font=('Arial', 18, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Left column container
        left_frame = ttk.Frame(main_frame)
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # Projects directory display
        dir_frame = ttk.LabelFrame(left_frame, text="Projects Directory", padding="10")
        dir_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.dir_label = ttk.Label(dir_frame, text=self.config['projects_directory'], 
                                   foreground='blue', font=('Arial', 9))
        self.dir_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        ttk.Button(dir_frame, text="Change Directory", 
                  command=self.change_projects_directory).grid(row=0, column=1)
        
        # Active project display
        active_frame = ttk.LabelFrame(left_frame, text="Active Project", padding="10")
        active_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        active_project_name = self.config['active_project'] if self.config['active_project'] else "None"
        self.active_project_label = ttk.Label(active_frame, 
                                              text=active_project_name, 
                                              font=('Arial', 12, 'bold'),
                                              foreground='green')
        self.active_project_label.grid(row=0, column=0, sticky=tk.W)
        
        # Project management buttons
        project_frame = ttk.LabelFrame(left_frame, text="Project Management", padding="15")
        project_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        btn_frame = ttk.Frame(project_frame)
        btn_frame.grid(row=0, column=0, columnspan=2)
        
        ttk.Button(btn_frame, text="New Plan", command=self.new_project, 
                  width=20).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(btn_frame, text="Open Existing Plan", command=self.open_project, 
                  width=20).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(btn_frame, text="Delete Project", command=self.delete_project, 
                  width=20).grid(row=0, column=2, padx=5, pady=5)
        
        # Recent projects dropdown
        ttk.Label(project_frame, text="Recent Projects:", 
                 font=('Arial', 10)).grid(row=1, column=0, sticky=tk.W, pady=(15, 5))
        
        self.projects_var = tk.StringVar()
        self.projects_combo = ttk.Combobox(project_frame, textvariable=self.projects_var, 
                                          state='readonly', width=40)
        self.projects_combo.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        self.projects_combo.bind('<<ComboboxSelected>>', self.on_project_selected)
        
        # Launch buttons
        launch_frame = ttk.LabelFrame(left_frame, text="Launch Applications", padding="15")
        launch_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.distance_btn = ttk.Button(launch_frame, text="Launch Postcode Distance Calculator", 
                                      command=self.launch_distance_app, width=40)
        self.distance_btn.grid(row=0, column=0, pady=5)
        
        self.clustering_btn = ttk.Button(launch_frame, text="Launch TSP Clustering Optimizer",
                                        command=self.launch_clustering_app, width=40)
        self.clustering_btn.grid(row=1, column=0, pady=5)
        
        self.scheduler_btn = ttk.Button(launch_frame, text="Launch Calendar Organizer", 
                                       command=self.launch_scheduler_app, width=40)
        self.scheduler_btn.grid(row=2, column=0, pady=5)
        
        self.smart_scheduler_btn = ttk.Button(launch_frame, text="Launch Smart Scheduler", 
                                             command=self.launch_smart_scheduler_app, width=40)
        self.smart_scheduler_btn.grid(row=3, column=0, pady=5)
        
        # Right column - Project Status
        info_frame = ttk.LabelFrame(main_frame, text="Project Status & Workflow", padding="10")
        info_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        info_frame.rowconfigure(1, weight=1)
        info_frame.columnconfigure(0, weight=1)
        
        # Add refresh button at top of info frame
        refresh_btn_frame = ttk.Frame(info_frame)
        refresh_btn_frame.grid(row=0, column=0, sticky=tk.E, pady=(0, 5))
        ttk.Button(refresh_btn_frame, text="Refresh", command=self.update_project_info, 
                  width=12).pack(side=tk.RIGHT)
        
        self.info_text = tk.Text(info_frame, height=20, width=50, font=('Consolas', 9),
                                state=tk.DISABLED, wrap=tk.WORD)
        self.info_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Update button states
        self.update_button_states()
        
    def refresh_projects_list(self):
        """Refresh the list of available projects"""
        try:
            projects_dir = self.config['projects_directory']
            if os.path.exists(projects_dir):
                projects = [d for d in os.listdir(projects_dir) 
                           if os.path.isdir(os.path.join(projects_dir, d))]
                projects.sort()
                self.projects_combo['values'] = projects
                
                # Set current selection if active project exists
                if self.config['active_project'] and self.config['active_project'] in projects:
                    self.projects_var.set(self.config['active_project'])
                    self.update_project_info()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh projects list:\n{e}")
    
    def on_project_selected(self, event):
        """Handle project selection from dropdown"""
        selected = self.projects_var.get()
        if selected:
            self.config['active_project'] = selected
            self.active_project_label.config(text=selected)
            self.add_to_recent_projects(selected)
            self.save_config()
            self.update_button_states()
            self.update_project_info()
    
    def add_to_recent_projects(self, project_name):
        """Add project to recent projects list"""
        if project_name in self.config['recent_projects']:
            self.config['recent_projects'].remove(project_name)
        self.config['recent_projects'].insert(0, project_name)
        # Keep only last 10
        self.config['recent_projects'] = self.config['recent_projects'][:10]
    
    def update_button_states(self):
        """Enable/disable launch buttons based on active project"""
        if self.config['active_project']:
            self.distance_btn.config(state=tk.NORMAL)
            self.clustering_btn.config(state=tk.NORMAL)
            self.scheduler_btn.config(state=tk.NORMAL)
            self.smart_scheduler_btn.config(state=tk.NORMAL)
        else:
            self.distance_btn.config(state=tk.DISABLED)
            self.clustering_btn.config(state=tk.DISABLED)
            self.scheduler_btn.config(state=tk.DISABLED)
            self.smart_scheduler_btn.config(state=tk.DISABLED)
    
    def update_project_info(self):
        """Update the project files info display"""
        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete('1.0', tk.END)
        
        if not self.config['active_project']:
            self.info_text.insert('1.0', "No active project selected.")
            self.info_text.config(state=tk.DISABLED)
            return
        
        project_path = os.path.join(self.config['projects_directory'], 
                                   self.config['active_project'])
        
        if not os.path.exists(project_path):
            self.info_text.insert('1.0', "Project directory not found.")
            self.info_text.config(state=tk.DISABLED)
            return
        
        info = f"Project: {self.config['active_project']}\n"
        info += f"Location: {project_path}\n"
        info += "\n" + "=" * 60 + "\n"
        info += "PROJECT WORKFLOW STATUS\n"
        info += "=" * 60 + "\n\n"
        
        # Task 1: Initial Locations Setup
        locations_file = os.path.join(project_path, 'locations.csv')
        if os.path.exists(locations_file):
            try:
                df = pd.read_csv(locations_file)
                num_locations = len(df)
                info += f"✓ Task 1: Initial Locations Setup - COMPLETE\n"
                info += f"  └─ {num_locations} location(s) loaded\n\n"
            except:
                info += f"✓ Task 1: Initial Locations Setup - COMPLETE\n"
                info += f"  └─ locations.csv exists\n\n"
        else:
            info += f"✗ Task 1: Initial Locations Setup - PENDING\n"
            info += f"  └─ Need to add locations.csv file\n\n"
        
        # Task 2: Distance Calculation
        distance_matrix_file = os.path.join(project_path, 'distance_matrix.csv')
        distances_file = os.path.join(project_path, 'distances.csv')
        if os.path.exists(distance_matrix_file) and os.path.exists(distances_file):
            try:
                dist_df = pd.read_csv(distances_file)
                num_distances = len(dist_df)
                info += f"✓ Task 2: Distance Calculation - COMPLETE\n"
                info += f"  └─ {num_distances} distance pair(s) calculated\n\n"
            except:
                info += f"✓ Task 2: Distance Calculation - COMPLETE\n"
                info += f"  └─ Distance files exist\n\n"
        else:
            info += f"✗ Task 2: Distance Calculation - PENDING\n"
            info += f"  └─ Run Postcode Distance Calculator\n\n"
        
        # Task 3: Region Clustering
        clustered_file = os.path.join(project_path, 'clustered_regions.csv')
        region_summary_file = os.path.join(project_path, 'region_summary.csv')
        if os.path.exists(clustered_file) and os.path.exists(region_summary_file):
            try:
                cluster_df = pd.read_csv(clustered_file)
                summary_df = pd.read_csv(region_summary_file)
                num_regions = len(summary_df)
                num_clustered_locs = len(cluster_df)
                info += f"✓ Task 3: Region Clustering - COMPLETE\n"
                info += f"  ├─ {num_clustered_locs} location(s) assigned to regions\n"
                info += f"  └─ {num_regions} region(s) created\n\n"
            except:
                info += f"✓ Task 3: Region Clustering - COMPLETE\n"
                info += f"  └─ Clustering files exist\n\n"
        else:
            info += f"✗ Task 3: Region Clustering - PENDING\n"
            info += f"  └─ Run TSP Clustering Optimizer\n\n"
        
        # Task 4: Calendar Organization
        schedule_file = os.path.join(project_path, 'region_schedule.csv')
        region_names_file = os.path.join(project_path, 'region_names.csv')
        if os.path.exists(schedule_file):
            try:
                schedule_df = pd.read_csv(schedule_file)
                num_scheduled_days = len(schedule_df)
                info += f"✓ Task 4: Calendar Organization - COMPLETE\n"
                info += f"  ├─ {num_scheduled_days} day(s) scheduled\n"
                if os.path.exists(region_names_file):
                    names_df = pd.read_csv(region_names_file)
                    named_regions = len(names_df)
                    info += f"  └─ {named_regions} region(s) customized\n\n"
                else:
                    info += f"  └─ Regions not yet customized\n\n"
            except:
                info += f"✓ Task 4: Calendar Organization - COMPLETE\n"
                info += f"  └─ Schedule file exists\n\n"
        else:
            info += f"✗ Task 4: Calendar Organization - PENDING\n"
            info += f"  └─ Run Calendar Organizer\n\n"
        
        # Task 5: Smart Scheduling
        confirmed_appointments_file = os.path.join(project_path, 'confirmed_appointments.csv')
        if os.path.exists(confirmed_appointments_file):
            try:
                appt_df = pd.read_csv(confirmed_appointments_file)
                num_appointments = len(appt_df)
                
                # Check how many are in Outlook
                outlook_synced = 0
                if 'in_outlook' in appt_df.columns:
                    outlook_synced = appt_df['in_outlook'].sum()
                
                info += f"✓ Task 5: Smart Scheduling - IN PROGRESS\n"
                info += f"  ├─ {num_appointments} appointment(s) scheduled\n"
                
                if outlook_synced > 0:
                    info += f"  ├─ {outlook_synced} appointment(s) synced to Outlook\n"
                
                # Calculate statistics
                if num_appointments > 0:
                    scheduled_locations = set(appt_df['postcode'].values)
                    
                    # Try to get total locations from clustered_regions
                    if os.path.exists(clustered_file):
                        cluster_df = pd.read_csv(clustered_file)
                        total_locations = len(cluster_df)
                        coverage = (len(scheduled_locations) / total_locations) * 100
                        info += f"  └─ {len(scheduled_locations)}/{total_locations} locations scheduled ({coverage:.1f}%)\n\n"
                    else:
                        info += f"  └─ {len(scheduled_locations)} unique location(s)\n\n"
                else:
                    info += f"  └─ No appointments scheduled yet\n\n"
            except Exception as e:
                info += f"✓ Task 5: Smart Scheduling - IN PROGRESS\n"
                info += f"  └─ Appointments file exists\n\n"
        else:
            info += f"✗ Task 5: Smart Scheduling - PENDING\n"
            info += f"  └─ Run Smart Scheduler\n\n"
        
        info += "=" * 60 + "\n"
        
        self.info_text.insert('1.0', info)
        self.info_text.config(state=tk.DISABLED)
    
    def change_projects_directory(self):
        """Change the default projects directory"""
        new_dir = filedialog.askdirectory(title="Select Projects Directory",
                                         initialdir=self.config['projects_directory'])
        if new_dir:
            self.config['projects_directory'] = new_dir
            self.dir_label.config(text=new_dir)
            self.save_config()
            
            # Ensure directory exists
            if not os.path.exists(new_dir):
                os.makedirs(new_dir)
            
            self.refresh_projects_list()
            messagebox.showinfo("Success", f"Projects directory updated to:\n{new_dir}")
    
    def new_project(self):
        """Create a new project"""
        # Ask for project name
        dialog = tk.Toplevel(self.root)
        dialog.title("New Project")
        dialog.geometry("400x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Enter project name:", font=('Arial', 11)).pack(pady=(0, 10))
        
        name_var = tk.StringVar()
        name_entry = ttk.Entry(frame, textvariable=name_var, width=40)
        name_entry.pack(pady=(0, 20))
        name_entry.focus()
        
        result = {'confirmed': False}
        
        def confirm():
            if name_var.get().strip():
                result['confirmed'] = True
                dialog.destroy()
            else:
                messagebox.showwarning("Invalid Name", "Please enter a project name.")
        
        def cancel():
            dialog.destroy()
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        
        ttk.Button(btn_frame, text="Create", command=confirm, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=cancel, width=12).pack(side=tk.LEFT, padx=5)
        
        # Bind Enter key
        name_entry.bind('<Return>', lambda e: confirm())
        
        self.root.wait_window(dialog)
        
        if not result['confirmed']:
            return
        
        project_name = name_var.get().strip()
        project_path = os.path.join(self.config['projects_directory'], project_name)
        
        # Check if project already exists
        if os.path.exists(project_path):
            messagebox.showerror("Error", f"Project '{project_name}' already exists.")
            return
        
        # Create project directory
        try:
            os.makedirs(project_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create project directory:\n{e}")
            return
        
        # Ask for initial locations CSV file
        locations_file = filedialog.askopenfilename(
            title="Select Initial Locations CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not locations_file:
            # User cancelled, but we already created the directory
            response = messagebox.askyesno("No File Selected", 
                                          "No locations file selected. Keep empty project?")
            if not response:
                # Remove the directory
                try:
                    os.rmdir(project_path)
                except:
                    pass
                return
        else:
            # Copy file to project directory as locations.csv
            try:
                dest_path = os.path.join(project_path, "locations.csv")
                shutil.copy2(locations_file, dest_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy locations file:\n{e}")
                return
        
        # Set as active project
        self.config['active_project'] = project_name
        self.active_project_label.config(text=project_name)
        self.add_to_recent_projects(project_name)
        self.save_config()
        
        # Refresh UI
        self.refresh_projects_list()
        self.projects_var.set(project_name)
        self.update_button_states()
        self.update_project_info()
        
        messagebox.showinfo("Success", 
                          f"Project '{project_name}' created successfully!\n\n"
                          f"Location: {project_path}")
    
    def open_project(self):
        """Open an existing project by browsing"""
        project_dir = filedialog.askdirectory(
            title="Select Project Directory",
            initialdir=self.config['projects_directory']
        )
        
        if not project_dir:
            return
        
        project_name = os.path.basename(project_dir)
        
        # Check if it's within the projects directory
        if not project_dir.startswith(self.config['projects_directory']):
            response = messagebox.askyesno("Different Location", 
                                          f"The selected project is outside the configured projects directory.\n\n"
                                          f"Do you want to update the projects directory to:\n"
                                          f"{os.path.dirname(project_dir)}?")
            if response:
                self.config['projects_directory'] = os.path.dirname(project_dir)
                self.dir_label.config(text=self.config['projects_directory'])
        
        # Set as active project
        self.config['active_project'] = project_name
        self.active_project_label.config(text=project_name)
        self.add_to_recent_projects(project_name)
        self.save_config()
        
        # Refresh UI
        self.refresh_projects_list()
        self.projects_var.set(project_name)
        self.update_button_states()
        self.update_project_info()
    
    def delete_project(self):
        """Delete an existing project"""
        # Get list of projects
        try:
            projects_dir = self.config['projects_directory']
            if not os.path.exists(projects_dir):
                messagebox.showwarning("No Projects", "Projects directory does not exist.")
                return
            
            projects = [d for d in os.listdir(projects_dir) 
                       if os.path.isdir(os.path.join(projects_dir, d))]
            
            if not projects:
                messagebox.showwarning("No Projects", "No projects found to delete.")
                return
            
            projects.sort()
            
            # Create selection dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Delete Project")
            dialog.geometry("400x300")
            dialog.transient(self.root)
            dialog.grab_set()
            
            frame = ttk.Frame(dialog, padding="20")
            frame.pack(fill=tk.BOTH, expand=True)
            
            ttk.Label(frame, text="Select Project to Delete", 
                     font=('Arial', 12, 'bold'), foreground='red').pack(pady=(0, 10))
            
            ttk.Label(frame, text="⚠️ Warning: This action cannot be undone!", 
                     font=('Arial', 9), foreground='darkred').pack(pady=(0, 20))
            
            # Project listbox
            listbox_frame = ttk.Frame(frame)
            listbox_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
            
            project_listbox = tk.Listbox(listbox_frame, height=8, font=('Arial', 10))
            project_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=project_listbox.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            project_listbox.config(yscrollcommand=scrollbar.set)
            
            for project in projects:
                project_listbox.insert(tk.END, project)
            
            result = {'confirmed': False, 'project': None}
            
            def confirm_delete():
                selection = project_listbox.curselection()
                if not selection:
                    messagebox.showwarning("No Selection", "Please select a project to delete.")
                    return
                
                project_name = project_listbox.get(selection[0])
                
                # Final confirmation
                response = messagebox.askyesno(
                    "Confirm Deletion",
                    f"Are you sure you want to delete project '{project_name}'?\n\n"
                    f"This will permanently delete:\n"
                    f"• The project folder\n"
                    f"• All CSV files\n"
                    f"• All analysis results\n\n"
                    f"This action CANNOT be undone!",
                    icon='warning'
                )
                
                if response:
                    result['confirmed'] = True
                    result['project'] = project_name
                    dialog.destroy()
            
            def cancel():
                dialog.destroy()
            
            btn_frame = ttk.Frame(frame)
            btn_frame.pack()
            
            ttk.Button(btn_frame, text="Delete", command=confirm_delete, width=12).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="Cancel", command=cancel, width=12).pack(side=tk.LEFT, padx=5)
            
            self.root.wait_window(dialog)
            
            if not result['confirmed']:
                return
            
            project_to_delete = result['project']
            project_path = os.path.join(projects_dir, project_to_delete)
            
            # Delete the project directory with Windows-specific handling
            try:
                # Helper function to handle read-only files on Windows
                def remove_readonly(func, path, excinfo):
                    """Error handler for Windows readonly files"""
                    os.chmod(path, stat.S_IWRITE)
                    func(path)
                
                # Try to delete with error handler
                shutil.rmtree(project_path, onerror=remove_readonly)
                
                # Give Windows a moment to release file handles
                time.sleep(0.1)
                
                # Verify deletion
                if os.path.exists(project_path):
                    # If still exists, try one more time with force
                    for root, dirs, files in os.walk(project_path, topdown=False):
                        for name in files:
                            filepath = os.path.join(root, name)
                            try:
                                os.chmod(filepath, stat.S_IWRITE)
                                os.remove(filepath)
                            except:
                                pass
                        for name in dirs:
                            try:
                                os.rmdir(os.path.join(root, name))
                            except:
                                pass
                    try:
                        os.rmdir(project_path)
                    except:
                        pass
                
                # Final check
                if os.path.exists(project_path):
                    messagebox.showerror("Error", 
                                       f"Failed to delete project completely.\n\n"
                                       f"Some files may be in use by another application.\n"
                                       f"Please close all applications using files from:\n"
                                       f"{project_path}\n\n"
                                       f"Then try deleting again or delete manually.")
                    return
                
                # Update active project if it was the deleted one
                if self.config['active_project'] == project_to_delete:
                    self.config['active_project'] = None
                    self.active_project_label.config(text="None")
                
                # Remove from recent projects
                if project_to_delete in self.config['recent_projects']:
                    self.config['recent_projects'].remove(project_to_delete)
                
                self.save_config()
                
                # Refresh UI
                self.refresh_projects_list()
                self.projects_var.set('')
                self.update_button_states()
                self.update_project_info()
                
                messagebox.showinfo("Success", 
                                  f"Project '{project_to_delete}' has been deleted successfully.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete project:\n{e}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to list projects:\n{e}")
    
    def _launch_app(self, app_filename, app_display_name, app_class=None, required_files=None, pre_launch_checks=None):
        """Generic method to launch any TSP application
        
        Args:
            app_filename: Name of the Python file to launch (e.g., 'postcode_distance_app.py')
            app_display_name: Display name for notifications (e.g., 'Postcode Distance Calculator')
            app_class: The app class to instantiate (if imported)
            required_files: Optional list of (filename, display_name) tuples to check before launching
            pre_launch_checks: Optional function for custom validation logic
        """
        if not self.config['active_project']:
            messagebox.showwarning("No Project", "Please select or create a project first.")
            return
        
        project_path = os.path.join(self.config['projects_directory'], 
                                   self.config['active_project'])
        
        # Run custom pre-launch checks if provided
        if pre_launch_checks and not pre_launch_checks(project_path):
            return
        
        # Check required files if specified
        if required_files:
            for filename, display_name in required_files:
                file_path = os.path.join(project_path, filename)
                if not os.path.exists(file_path):
                    messagebox.showwarning("Missing File", 
                                         f"{display_name} not found in project.\n\n"
                                         f"Please ensure required files are available.")
                    return
        
        # Determine if we're running as an EXE
        is_frozen = getattr(sys, 'frozen', False)
        
        print(f"DEBUG: is_frozen={is_frozen}, APPS_IMPORTED={APPS_IMPORTED}, app_class={app_class}")
        
        # If running as EXE, we MUST use the imported classes
        if is_frozen:
            if APPS_IMPORTED and app_class:
                try:
                    # Launch directly in main thread for proper tkinter behavior
                    new_root = tk.Toplevel(self.root)
                    app_class(new_root, project_dir=project_path)
                    return
                except Exception as e:
                    import traceback
                    error_details = traceback.format_exc()
                    print(f"Failed to launch via import: {error_details}")
                    messagebox.showerror("Error", 
                                       f"Failed to launch {app_display_name}:\n{e}\n\n"
                                       f"Please check the error log.")
                    return
            else:
                # Running as EXE but imports failed - cannot proceed
                messagebox.showerror("Error", 
                                   f"Failed to launch {app_display_name}.\n\n"
                                   f"The application components could not be loaded.\n"
                                   f"APPS_IMPORTED={APPS_IMPORTED}")
                return
        
        # Not running as EXE (development mode) - try import first, then subprocess
        if APPS_IMPORTED and app_class:
            try:
                # Try to launch with imported class
                new_root = tk.Toplevel(self.root)
                app_class(new_root, project_dir=project_path)
                return
            except Exception as e:
                print(f"Failed to launch via import, trying subprocess: {e}")
                # Fall through to subprocess method
        
        # Subprocess method (development mode only)
        app_path = os.path.join(self.app_directory, app_filename)
        
        if not os.path.exists(app_path):
            messagebox.showerror("Error", 
                               f"{app_display_name} not found:\n{app_path}")
            return
        
        try:
            subprocess.Popen([sys.executable, app_path, project_path])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch application:\n{e}")
    
    def launch_distance_app(self):
        """Launch the postcode distance calculator"""
        self._launch_app("postcode_distance_app.py", "Postcode Distance Calculator",
                        app_class=PostcodeDistanceApp if APPS_IMPORTED else None)
    
    def launch_clustering_app(self):
        """Launch the TSP clustering optimizer"""
        def check_clustering_requirements(project_path):
            locations_file = os.path.join(project_path, "locations.csv")
            distances_file = os.path.join(project_path, "distances.csv")
            
            if not os.path.exists(locations_file):
                messagebox.showwarning("Missing File", 
                                     "locations.csv not found in project.\n\n"
                                     "Please add the locations file first.")
                return False
            
            if not os.path.exists(distances_file):
                response = messagebox.askyesno("Missing Distance Data", 
                                              "distances.csv not found in project.\n\n"
                                              "You need to run the Postcode Distance Calculator first.\n\n"
                                              "Launch it now?")
                if response:
                    self.launch_distance_app()
                return False
            
            return True
        
        self._launch_app("tsp_clustering_app.py", "TSP Clustering Optimizer",
                        app_class=TSPClusteringApp if APPS_IMPORTED else None,
                        pre_launch_checks=check_clustering_requirements)
    
    def launch_scheduler_app(self):
        """Launch the Calendar Organizer"""
        self._launch_app("calendar_organizer_app.py", "Calendar Organizer",
                        app_class=CalendarOrganizerApp if APPS_IMPORTED else None)
    
    def launch_smart_scheduler_app(self):
        """Launch the Smart Scheduler"""
        self._launch_app("smart_scheduler_app.py", "Smart Scheduler",
                        app_class=SmartSchedulerApp if APPS_IMPORTED else None)


def main():
    root = tk.Tk()
    app = ProjectLauncher(root)
    root.mainloop()


if __name__ == "__main__":
    main()
