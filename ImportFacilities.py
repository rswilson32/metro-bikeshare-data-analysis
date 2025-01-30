import pandas as pd
from comtypes.client import GetActiveObject
import glob
import os

def ImportFacilities(STKVersion, filepath):
    # ImportFacilities attaches to an open instance of STK and imports position data
    # from an Excel spreadsheet. Inputs include STK whole number version as an
    # integer and Excel file path. Units are assumed to be degrees and meters with a
    # header row in the Excel file for ID, LAT, LON, ALT. This function requires the
    # pandas Python library.

    try:
        # Grab a running instance of STK
        uiApplication = GetActiveObject(f"STK{STKVersion}.Application")
        root = uiApplication.Personality2
    except Exception as e:
        print(f"Error: Could not connect to STK version {STKVersion}. Is STK running? {e}")
        return

    from comtypes.gen import STKObjects

    # Grab current scenario
    scenario = root.CurrentScenario
    uiApplication.Visible = True
    uiApplication.UserControl = True

    # Change the latitude and longitude to degrees
    root.UnitPreferences.Item("Latitude").SetCurrentUnit("deg")
    root.UnitPreferences.Item("Longitude").SetCurrentUnit("deg")

    # Change the distance to meters
    root.UnitPreferences.SetCurrentUnit("Distance", "m")

    # Use pandas to read in excel sheet as a dataframe
    if not os.path.exists(filepath):
        print(f"Warning: File {filepath} does not exist.")
        return

    try:
        df = pd.read_csv(filepath)  # Assuming CSV format, adjust if it's an Excel file
    except Exception as e:
        print(f"Error reading file {filepath}: {e}")
        return

    # Iterate through each row
    for _, row in df.iterrows():
        facName = row["facility_name"]
        lat = row["lat"]
        lon = row["lon"]
        shortDesc = row["short_description"]
        longDesc = row["long_description"]

        # There cannot be two objects with the same name in STK, so
        # if there is already a facility with the same name, delete it.
        if scenario.Children.Contains(STKObjects.eFacility, facName):
            obj = scenario.Children.Item(facName)
            obj.Unload()

        # Create the facility
        fac = scenario.Children.New(STKObjects.eFacility, facName)
        fac2 = fac.QueryInterface(STKObjects.IAgFacility)

        # Choose to use terrain
        fac2.UseTerrain = True

        # Set the latitude, longitude, and altitude to 0
        fac2.Position.AssignGeodetic(lat, lon, 0)  # Altitude set to 0
        fac.ShortDescription = str(shortDesc)
        fac.LongDescription = str(longDesc)

    print(f"Successfully imported facilities from {filepath}.")

def main():
    files = glob.glob('*.csv')  # Find all csv files
    if not files:
        print("Warning: No .csv files found in the current directory.")
        return

    for file in files:
        print(f"Processing file: {file}")
        ImportFacilities(12, file)

if __name__ == "__main__":
    main()
