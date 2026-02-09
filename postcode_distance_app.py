import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import requests
import time
import threading
from itertools import combinations
import os
import sys


class PostcodeDistanceApp:
    def __init__(self, root, project_dir=None):
        self.root = root
        self.root.title("Postcode Distance Calculator")
        self.root.geometry("900x700")
        
        # Project directory from command line
        self.project_dir = project_dir
        
        # Variables
        self.input_file = None
        self.output_dir = None
        self.postcodes = []
        
        self.setup_ui()
        
        # Auto-load project files if project directory provided
        if self.project_dir:
            self.auto_load_project_files()
        
    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Postcode Distance Calculator", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # File Selection and Output Directory sections removed - using project-based workflow
        
        # Generate Button
        generate_frame = ttk.Frame(main_frame)
        generate_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        self.generate_btn = ttk.Button(generate_frame, text="Generate CSV Files", 
                                       command=self.start_generation, state=tk.DISABLED)
        self.generate_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(generate_frame, text="(This may take several minutes)", 
                 foreground="gray").pack(side=tk.LEFT)
        
        # Progress Section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=400)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        progress_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(progress_frame, text="Ready", foreground="green")
        self.status_label.grid(row=1, column=0, sticky=tk.W)
        
        # Log Section
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.rowconfigure(6, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        main_frame.columnconfigure(0, weight=1)
    
    def auto_load_project_files(self):
        """Auto-load files from project directory"""
        if not self.project_dir or not os.path.exists(self.project_dir):
            return
        
        project_name = os.path.basename(self.project_dir)
        self.root.title(f"Postcode Distance Calculator - Project: {project_name}")
        
        # Set output directory to project directory
        self.output_dir = self.project_dir
        
        # Load locations.csv
        locations_path = os.path.join(self.project_dir, "locations.csv")
        if os.path.exists(locations_path):
            self.input_file = locations_path
            self.log(f"✓ Auto-loaded: {locations_path}")
            
            # Load and parse postcodes
            try:
                # First try to read with headers
                df = pd.read_csv(locations_path)
                
                # Check if 'postcode' column exists
                if 'postcode' not in df.columns:
                    # Check if first value looks like a postcode (not a header)
                    first_value = str(df.columns[0])
                    # Simple heuristic: if it looks like a postcode, assume no header
                    if any(char.isdigit() for char in first_value) or len(first_value) <= 10:
                        self.log("⚠ No 'postcode' header found - reading as headerless CSV")
                        # Reload without header
                        df = pd.read_csv(locations_path, header=None, names=['postcode'])
                        # Save with proper header
                        df.to_csv(locations_path, index=False)
                        self.log(f"✓ Added 'postcode' header to {os.path.basename(locations_path)}")
                    else:
                        # First row looks like a header but not 'postcode' - rename it
                        df = df.rename(columns={df.columns[0]: 'postcode'})
                        df.to_csv(locations_path, index=False)
                        self.log(f"✓ Renamed first column to 'postcode'")
                
                self.postcodes = df['postcode'].dropna().str.strip().unique().tolist()
                self.log(f"✓ Loaded {len(self.postcodes)} unique postcodes")
                self.log(f"\n✓ Project '{project_name}' loaded successfully")
                
                # Enable generate button since we have input file and output directory
                self.generate_btn.config(state=tk.NORMAL)
                
            except Exception as e:
                self.log(f"✗ Error loading postcodes: {e}")
        else:
            self.log(f"⚠ locations.csv not found in project directory")
    
    def log(self, message):
        """Add message to log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def update_status(self, message, color="black"):
        """Update status label"""
        self.status_label.config(text=message, foreground=color)
        self.root.update_idletasks()
            
    def browse_output_dir(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(title="Select Output Directory")
        
        if directory:
            self.output_dir = directory
            self.log(f"Selected output directory: {directory}")
            
            # Enable generate button if we have both input and output
            if self.input_file and self.output_dir:
                self.generate_btn.config(state=tk.NORMAL)
            
    def load_postcodes(self):
        """Load postcodes from CSV and extract prefixes"""
        try:
            df = pd.read_csv(self.input_file, header=None, names=['postcode'])
            self.postcodes = df['postcode'].str.strip().unique().tolist()
            
            self.log(f"Loaded {len(self.postcodes)} postcodes")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load postcodes: {e}")
            self.log(f"ERROR: {e}")
        
    def start_generation(self):
        """Start the generation process in a separate thread"""
        self.generate_btn.config(state=tk.DISABLED)
        self.update_status("Processing...", "orange")
        
        # Run in separate thread to keep UI responsive
        thread = threading.Thread(target=self.generate_files)
        thread.daemon = True
        thread.start()
        
    def generate_files(self):
        """Generate the two CSV files"""
        try:
            self.log("\n" + "="*70)
            self.log("Starting postcode processing...")
            self.log("="*70)
            
            # Step 1: Geocode postcodes
            self.log("\nSTEP 1: Geocoding postcodes...")
            postcode_coords = {}
            total_postcodes = len(self.postcodes)
            
            for i, postcode in enumerate(self.postcodes, 1):
                coords = self.get_coordinates_from_postcode(postcode)
                if coords:
                    postcode_coords[postcode] = coords
                    if i % 5 == 0 or i == total_postcodes:
                        progress = (i / total_postcodes) * 30  # First 30% for geocoding
                        self.progress_bar['value'] = progress
                        self.log(f"  Geocoded {i}/{total_postcodes} ({i*100//total_postcodes}%)")
                else:
                    self.log(f"  ✗ Failed: {postcode}")
                time.sleep(0.1)
            
            self.log(f"✓ Successfully geocoded {len(postcode_coords)} postcodes")
            
            # Save coordinates file as distance_matrix.csv
            coords_output = os.path.join(self.output_dir, "distance_matrix.csv")
            postcode_list_df = pd.DataFrame([
                {'postcode': pc, 'latitude': coords['latitude'], 'longitude': coords['longitude']}
                for pc, coords in sorted(postcode_coords.items())
            ])
            postcode_list_df.to_csv(coords_output, index=False)
            self.log(f"✓ Saved coordinates to: {coords_output}")
            
            # Step 2: Calculate distances
            self.log("\nSTEP 2: Calculating driving distances...")
            
            # Generate all pairs
            all_pairs = list(combinations(sorted(postcode_coords.keys()), 2))
            
            self.log(f"Total pairs to calculate: {len(all_pairs)}")
            self.log(f"Estimated time: ~{len(all_pairs) * 1.2 / 60:.1f} minutes\n")
            
            results = []
            processed = 0
            failed = 0
            
            for origin, dest in all_pairs:
                origin_coords = postcode_coords[origin]
                dest_coords = postcode_coords[dest]
                
                driving_data = self.get_driving_time_osrm(origin_coords, dest_coords)
                
                if driving_data:
                    results.append({
                        'origin': origin,
                        'destination': dest,
                        'driving_time_minutes': driving_data['duration_minutes'],
                        'distance_km': driving_data['distance_km']
                    })
                    processed += 1
                    
                    if processed % 20 == 0 or processed == len(all_pairs):
                        progress = 30 + (processed / len(all_pairs)) * 70  # Last 70%
                        self.progress_bar['value'] = progress
                        self.log(f"  Progress: {processed}/{len(all_pairs)} routes ({processed*100//len(all_pairs)}%) | Failed: {failed}")
                else:
                    failed += 1
                
                time.sleep(1)  # Be respectful to the API
            
            self.log(f"\n✓ Successfully calculated {processed} routes")
            if failed > 0:
                self.log(f"⚠ Failed to calculate {failed} routes")
            
            # Save distances file
            distances_output = os.path.join(self.output_dir, "distances.csv")
            results_df = pd.DataFrame(results)
            results_df.to_csv(distances_output, index=False)
            self.log(f"✓ Saved distances to: {distances_output}")
            
            # Summary
            self.log("\n" + "="*70)
            self.log("SUMMARY")
            self.log("="*70)
            self.log(f"Postcodes processed: {len(postcode_coords)}")
            self.log(f"Routes calculated: {len(results)}")
            self.log(f"\nFiles created:")
            self.log(f"  1. {coords_output}")
            self.log(f"  2. {distances_output}")
            self.log("\n✓ COMPLETE!")
            self.log("="*70)
            
            self.progress_bar['value'] = 100
            self.update_status("Complete!", "green")
            
            messagebox.showinfo("Success", 
                              f"CSV files generated successfully!\n\n"
                              f"Postcodes: {len(postcode_coords)}\n"
                              f"Routes calculated: {len(results)}\n\n"
                              f"Files saved to:\n{self.output_dir}")
            
        except Exception as e:
            self.log(f"\n✗ ERROR: {e}")
            self.update_status("Error!", "red")
            messagebox.showerror("Error", f"An error occurred:\n{e}")
        
        finally:
            self.generate_btn.config(state=tk.NORMAL)
            
    def get_coordinates_from_postcode(self, postcode):
        """Get coordinates from postcode using postcodes.io API"""
        postcode_clean = postcode.replace(" ", "")
        url = f"https://api.postcodes.io/postcodes/{postcode_clean}"
        
        try:
            response = requests.get(url)
            if response.status_code == 200:
                data = response.json()
                return {
                    'latitude': data['result']['latitude'],
                    'longitude': data['result']['longitude']
                }
        except Exception as e:
            self.log(f"Error geocoding {postcode}: {e}")
        return None
        
    def get_driving_time_osrm(self, origin_coords, dest_coords):
        """Get driving time using OSRM API"""
        url = f"http://router.project-osrm.org/route/v1/driving/{origin_coords['longitude']},{origin_coords['latitude']};{dest_coords['longitude']},{dest_coords['latitude']}"
        params = {'overview': 'false'}
        
        try:
            response = requests.get(url, params=params)
            if response.status_code == 200:
                data = response.json()
                if data['code'] == 'Ok':
                    route = data['routes'][0]
                    duration_minutes = route['duration'] / 60
                    distance_km = route['distance'] / 1000
                    return {
                        'duration_minutes': round(duration_minutes, 2),
                        'distance_km': round(distance_km, 2)
                    }
        except Exception as e:
            self.log(f"Error getting route: {e}")
        return None


def main():
    # Check for project directory argument
    project_dir = None
    if len(sys.argv) > 1:
        project_dir = sys.argv[1]
        if not os.path.exists(project_dir):
            print(f"Warning: Project directory not found: {project_dir}")
            project_dir = None
    
    root = tk.Tk()
    app = PostcodeDistanceApp(root, project_dir=project_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
