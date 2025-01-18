import pandas as pd
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import time
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import os

class GeocodingApp:
    """
    A GUI application for geocoding addresses from Excel files.
    
    This application allows users to:
    1. Load addresses from an Excel file
    2. Standardize German addresses (convert umlauts, street names)
    3. Geocode addresses to obtain latitude and longitude
    4. Save results to a new Excel file

    Attributes:
        root (tk.Tk): The main window of the application
        input_path (tk.StringVar): Stores the path to input Excel file
        output_path (tk.StringVar): Stores the path to output Excel file
        progress_var (tk.DoubleVar): Stores the progress bar value
        status_var (tk.StringVar): Stores the current status message
    """
    def __init__(self, root):
        
        """
        Initialize the GeocodingApp with its GUI components.

        Args:
            root (tk.Tk): The main window of the application
        """
        self.root = root
        self.root.title("Address Geocoding Tool")
        self.root.geometry("600x400")
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Input file section
        ttk.Label(self.main_frame, text="Input Excel File:").grid(row=0, column=0, sticky=tk.W, pady=5)
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
        
        """
        Open a file dialog for selecting the input Excel file.
        
        Updates the input_path StringVar with the selected file path.
        Only shows Excel files (.xlsx, .xls) in the file dialog.
        """
        
        
        filename = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_path.set(filename)
            
    def browse_output(self):
        
        """
        Open a file dialog for selecting the output Excel file location.
        
        Updates the output_path StringVar with the selected file path.
        Automatically adds .xlsx extension if not specified.
        """
        
        filename = filedialog.asksaveasfilename(
            title="Save Output Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_path.set(filename)
    
    def merge_addresses(self, df):
        """
        Standardize and merge address components into a single address string.

        This method:
        1. Converts numeric fields (house numbers, postal codes) to proper format
        2. Concatenates address components with appropriate separators
        3. Converts German umlauts to their standard representations
        4. Standardizes street name variations

        Args:
            df (pandas.DataFrame): DataFrame containing address components in separate columns

        Returns:
            pandas.DataFrame: DataFrame with added 'address' column containing standardized addresses

        Required columns in input DataFrame:
            - street: Street name
            - hs_nr: House number
            - hs_nr_x: House number extension (if any)
            - plz: Postal code
            - ort: City
            - State: State/Region
            - country: Country
        """
        # Convert numbers to integers before string conversion
        df['hs_nr'] = df['hs_nr'].fillna('').apply(lambda x: str(int(float(x))) if x != '' else '')
        df['plz'] = df['plz'].fillna('').apply(lambda x: str(int(float(x))) if x != '' else '')
        
        # Create complete address using concat, skip NaN values
        df['address'] = df['street'].str.cat(
            df['hs_nr'].astype(str),
            sep=' ',
            na_rep=''
        ).str.cat(
            df['hs_nr_x'],
            sep='',
            na_rep=''
        ).str.cat(
            df['plz'].astype(str),
            sep=', ',
            na_rep=''
        ).str.cat(
            df['ort'],
            sep=' ',
            na_rep=''
        ).str.cat(
            df['State'],
            sep=', ',
            na_rep=''
        ).str.cat(
            df['country '],
            sep=', ',
            na_rep=''
        )

        # Define umlaut mapping
        umlaut_mapping = {
            'ä': 'ae',
            'ö': 'oe',
            'ü': 'ue',  
            'ß': 'ss'
        }
        
        # Apply umlaut conversions
        for umlaut, replacement in umlaut_mapping.items():
            df['address'] = df['address'].str.replace(umlaut, replacement)
        
        # Apply street variations conversions
        street_variations = {
            'strasse': 'str.',
            'Strasse': 'str.',
            'straße': 'str.',
            'Straße': 'str.'
        }
        
        for variant, replacement in street_variations.items():
            df['address'] = df['address'].str.replace(variant, replacement)
        

    
    
    # Additional cleaning steps
    
    
        df['address'] = df['address'].apply(lambda x: x.strip() if isinstance(x, str) else x)  # Remove extra spaces
        df['address'] = df['address'].str.replace('  ', ' ')  # Remove double spaces
        df['address'] = df['address'].str.replace(' ,', ',')  # Fix spacing around commas
        
        # Handle empty components better
        df['address'] = df['address'].str.replace(',,', ',')  # Remove empty components
        df['address'] = df['address'].str.replace(', ,', ',')  # Remove empty components
        
        return df

    def geocode_address(self, address):
        
        """
        Geocode a single address using Nominatim service.

        Uses a 1-second delay between requests to comply with Nominatim's usage policy.

        Args:
            address (str): The complete address string to geocode

        Returns:
            tuple: (latitude, longitude) if successful, (None, None) if geocoding fails

        Note:
            Handles both GeocoderTimedOut and other exceptions, updating status message
            on error.
        """
        
        max_retries = 3  # Number of retry attempts
        for attempt in range(max_retries):
            try:
                geolocator = Nominatim(user_agent="my_geocoder")
                time.sleep(1)  # Respect Nominatim's usage policy
                
                # Try variations of the address if previous attempts failed
                if attempt == 1:
                    # Try without house number
                    address = ' '.join(address.split(',')[1:])
                elif attempt == 2:
                    # Try just city and country
                    parts = address.split(',')
                    if len(parts) >= 2:
                        address = f"{parts[-2]}, {parts[-1]}"
                
                location = geolocator.geocode(address, exactly_one=True, addressdetails=True)
                
                if location:
                    return location.latitude, location.longitude
                
            except GeocoderTimedOut:
                self.status_var.set(f"Timeout - Attempt {attempt + 1} for: {address}")
                time.sleep(2)  # Wait longer before retry
            except Exception as e:
                self.status_var.set(f"Error - Attempt {attempt + 1} for {address}: {str(e)}")
                time.sleep(2)
        
        return None, None
    
    def process_file(self):
        
        """
        Main processing method that handles the complete geocoding workflow.

        This method:
        1. Validates input/output file selections
        2. Reads the input Excel file
        3. Standardizes addresses using merge_addresses
        4. Geocodes each address
        5. Saves results to the output Excel file

        Updates the progress bar and status message throughout processing.
        Shows error messages if any step fails.

        Note:
            Progress updates are shown in real-time through the GUI.
            All exceptions are caught and displayed to the user.
        """
        
        
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showerror("Error", "Please select both input and output files.")
            return
        
        try:
            # Read the Excel file
            self.status_var.set("Reading Excel file...")
            df = pd.read_excel(self.input_path.get())
            
            # Standardize and merge addresses
            self.status_var.set("Standardizing addresses...")
            df = self.merge_addresses(df)
            
            # Create new columns for coordinates
            df['latitude'] = None
            df['longitude'] = None
            
            total_rows = len(df)
            
            # Process each row
            missed_addresses = []
            for index, row in df.iterrows():
                # Update progress
                progress = (index + 1) / total_rows * 100
                self.progress_var.set(progress)
                self.status_var.set(f"Processing: {row['address']}")
                self.root.update()
                
                # Geocode
                lat, lon = self.geocode_address(row['address'])
                if lat is None or lon is None:
                    missed_addresses.append(row['address'])
                df.at[index, 'latitude'] = lat
                df.at[index, 'longitude'] = lon
            
            # Save missed addresses to a separate file
            if missed_addresses:
                missed_df = pd.DataFrame({'missed_addresses': missed_addresses})
                missed_file = os.path.splitext(self.output_path.get())[0] + '_missed.xlsx'
                missed_df.to_excel(missed_file, index=False)
                self.status_var.set(f"Completed. {len(missed_addresses)} addresses were not found.")
            
            # Save results
            df.to_excel(self.output_path.get(), index=False, engine='openpyxl')
            
            self.status_var.set("Processing completed successfully!")
            messagebox.showinfo("Success", "Geocoding completed successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_var.set(f"Error: {str(e)}")

def main():
    
    """
    Main entry point of the application.
    
    Creates the main window and starts the application's event loop.
    """
    
    
    root = tk.Tk()
    app = GeocodingApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()