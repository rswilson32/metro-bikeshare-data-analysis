import pandas as pd
import unicodedata
import re
import os
import glob

def clean_text(text, name=False):
    if not isinstance(text, str) or not text.strip():
        return ""  # Return empty string for None or non-string values

    # ignore non UTF-8 characters
    cleaned = unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode("utf-8")
    if name:
        cleaned = re.sub(r"[^\w\s]", "", cleaned)  # Remove special characters
        cleaned = re.sub(r"\s+", "_", cleaned)  # Replace multiple spaces with a single underscore
    return cleaned.strip()

# Convert DMS (Degrees, Minutes, Seconds) to Longitude and Latitude
def dms_to_decimal_degrees(dms_string):
    try:
        lat_str, lon_str = dms_string.strip().split()

        def convert(coord):
            match = re.match(r"(\d+(?:\.\d+)?)°([NSEW])", coord)  # Decimal Degrees
            if match:
                deg, direction = match.groups()
                decimal = float(deg)
                return decimal * (-1 if direction in ["W", "S"] else 1)

            match = re.match(r"(\d+)°(\d+)'([\d.]+)\"([NSEW])", coord)  # DMS Format
            if match:
                deg, minutes, seconds, direction = match.groups()
                decimal = float(deg) + float(minutes) / 60 + float(seconds) / 3600
                return decimal * (-1 if direction in ["W", "S"] else 1)

            raise ValueError(f"Invalid coordinate format: {coord}")

        return convert(lat_str), convert(lon_str)

    except Exception as e:
        print(f"Error converting coordinates: {dms_string} - {e}")
        return None, None  # Return NaN-compatible values

# Process Excel file and extract relevant columns
def process_facilities_file(file_path):
    df = pd.read_excel(file_path)
    
    if 'coordinates' not in df.columns:
        print(f"Skipping {file_path}: 'coordinates' column missing")
        return None

    # Apply coordinate conversion and filter out invalid rows
    valid_coords = df['coordinates'].dropna().map(dms_to_decimal_degrees)
    df[['lat', 'lon']] = pd.DataFrame(valid_coords.tolist(), index=df.index).dropna()

    # Clean text columns
    df['facility_name'] = df['facility_name'].apply(lambda x: clean_text(x, name=True))
    df['short_description'] = df['short_description'].apply(clean_text)
    df['long_description'] = df['long_description'].apply(clean_text)

    return df[['facility_name', 'lat', 'lon', 'short_description', 'long_description']]

# Save cleaned data to CSV
def save_csv(df, file_path):
    df.to_csv(file_path, index=False)

# Main function to process multiple Excel files
def main():
    files = glob.glob('*.xlsx')  # Find all Excel files
    for file in files:
        cleaned_df = process_facilities_file(file)
        if cleaned_df is not None:  # Ensure valid data before saving
            save_csv(cleaned_df, f"{os.path.splitext(file)[0]}_cleaned.csv")

if __name__ == "__main__":
    main()