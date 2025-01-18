# Address Geocoding Tool
A Python GUI application that converts German addresses from Excel files into geographical coordinates (latitude and longitude) using the Nominatim geocoding service.

## Prerequisites
### Required Python packages
pip install pandas geopy openpyxl pyinstaller

## Features
- User-friendly GUI interface
- Reads from Excel files (.xlsx, .xls)
- Standardizes German addresses (converts umlauts and street names)
- Geocodes addresses to obtain latitude and longitude
- Saves results to Excel
- Progress tracking with status updates
- Handles failed geocoding attempts with retry mechanism

## Installation & Setup
### Clone the repository:
git clone "repository-url"

### Create executable:
pyinstaller --onefile --windowed --name GeocodingTool test.py

The executable will be created in the dist folder.

## Using the Application

- Launch the application
- Select input Excel file
- Choose output location
- Click "Start Processing"

## Input Excel Format
### Required columns:

- street: Street name
- hs_nr: House number
- hs_nr_x: House number extension (optional)
- plz: Postal code
- ort: City
- State: State/Region
- country: Country

## Output

### Original columns plus:

- address: Standardized complete address
- latitude & longitude coordinates
- Failed addresses saved in *_missed.xlsx
