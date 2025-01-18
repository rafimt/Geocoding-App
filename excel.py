import pandas as pd
import numpy as np

# Create sample data
data = {
    'street': [
        'Kurfürstendamm', 'Friedrichstraße', 'Unter den Linden', 'Alexanderplatz',
        'Potsdamer Straße', 'Karl-Marx-Allee', 'Torstraße', 'Prenzlauer Allee',
        'Kantstraße', 'Warschauer Straße', 'Schönhauser Allee', 'Frankfurter Allee',
        'Leipziger Straße', 'Oranienstraße', 'Brunnenstraße', 'Karl-Liebknecht-Straße',
        'Mehringdamm', 'Petersburger Straße', 'Rosa-Luxemburg-Straße', 'Bernauer Straße'
    ],
    'hs_nr': np.random.randint(1, 200, 20),  # Random house numbers between 1 and 200
    'hs_nr_x': [''] * 20,  # Empty suffix for house numbers
    'plz': [
        '10719', '10117', '10117', '10178', '10785', '10243', '10119', '10405',
        '10627', '10243', '10435', '10247', '10117', '10999', '10119', '10178',
        '10965', '10249', '10178', '10119'
    ],
    'ort': ['Berlin'] * 20,
    'country': ['Germany'] * 20,
    'state': ['Berlin'] * 20
}

# Create DataFrame
df = pd.DataFrame(data)

# Add some random house number suffixes (a, b, c) to about 25% of the addresses
random_indices = np.random.choice(20, 5, replace=False)  # Select 5 random indices
suffixes = ['a', 'b', 'c']
for idx in random_indices:
    df.loc[idx, 'hs_nr_x'] = np.random.choice(suffixes)

# Display the first few rows
print(df.head())

# To save to Excel (commented out, but you can use this)
# df.to_excel('berlin_addresses.xlsx', index=False)

# Save the DataFrame to Excel
df.to_excel('berlin_addresses.xlsx', index=False, sheet_name='Addresses')