import pandas as pd
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import time
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import os

class GeocodingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Address Geocoding Tool")
        self.root.geometry("600x400")
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Input file section
        ttk.Label(self.main_frame, text="Input CSV File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_path = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(self.main_frame, text="Browse", command=self.browse_input).grid(row=0, column=2)
        
        # Output file section
        ttk.Label(self.main_frame, text="Output Excel File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_path = tk.StringVar()
        ttk.Entry(self.main_frame, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(self.main_frame, text="Browse", command=self.browse_output).grid(row=1, column=2)
        
        # Progress section
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Status message
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status_var)
        self.status_label.grid(row=3, column=0, columnspan=3, sticky=tk.W)
        
        # Process button
        self.process_button = ttk.Button(self.main_frame, text="Start Processing", command=self.process_file)
        self.process_button.grid(row=4, column=0, columnspan=3, pady=10)
        
    def browse_input(self):
        filename = filedialog.askopenfilename(
            title="Select Input CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.input_path.set(filename)
            
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Save Output Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_path.set(filename)
    
    def geocode_address(self, address):
        try:
            geolocator = Nominatim(user_agent="my_geocoder")
            time.sleep(1)
            location = geolocator.geocode(address)
            
            if location:
                return location.latitude, location.longitude
            return None, None
            
        except (GeocoderTimedOut, Exception) as e:
            self.status_var.set(f"Error geocoding address {address}: {str(e)}")
            return None, None
    
    def process_file(self):
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showerror("Error", "Please select both input and output files.")
            return
        
        try:
            # Read the CSV file
            df = pd.read_csv(self.input_path.get(), 
                           sep=';',
                           encoding='utf-8-sig')
            
            # Create new columns for coordinates
            df['latitude'] = None
            df['longitude'] = None
            
            total_rows = len(df)
            
            # Process each row
            for index, row in df.iterrows():
                # Update progress
                progress = (index + 1) / total_rows * 100
                self.progress_var.set(progress)
                self.status_var.set(f"Processing: {row['address']}")
                self.root.update()
                
                # Geocode
                lat, lon = self.geocode_address(row['address'])
                df.at[index, 'latitude'] = lat
                df.at[index, 'longitude'] = lon
            
            # Save results
            df.to_excel(self.output_path.get(), index=False, engine='openpyxl')
            
            self.status_var.set("Processing completed successfully!")
            messagebox.showinfo("Success", "Geocoding completed successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_var.set(f"Error: {str(e)}")

def main():
    root = tk.Tk()
    app = GeocodingApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()