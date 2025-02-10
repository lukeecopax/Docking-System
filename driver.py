import requests
import pandas as pd

def download_excel_file(url, output_filename):
    """
    Downloads an Excel file from the given URL and saves it locally.
    """
    print("Downloading Excel file...")
    response = requests.get(url)
    if response.status_code == 200:
        with open(output_filename, "wb") as f:
            f.write(response.content)
        print("Download completed.")
        return output_filename
    else:
        print(f"Error downloading file: {response.status_code}")
        return None

def get_available_drivers(filename):
    """
    Reads the 'Driver' sheet from the Excel file and returns a sorted list of unique driver names
    where the 'Available' column is 'YES'.
    """
    try:
        driver_df = pd.read_excel(filename, sheet_name="Driver", engine="openpyxl")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []
    
    # Ensure the 'Available' column is treated as string and filter rows where it equals "YES"
    available_df = driver_df[driver_df["Available"].astype(str).str.strip().str.upper() == "YES"]
    
    # Convert the 'Driver Name' column to string, drop missing values, and strip spaces
    drivers = available_df["Driver Name"].dropna().apply(lambda x: str(x).strip()).unique().tolist()
    
    # Filter out any entries that are empty or equal to "nan" (case-insensitive)
    drivers = [d for d in drivers if d and d.lower() != "nan"]
    
    return sorted(drivers)

def main():
    url = "https://netorg245148-my.sharepoint.com/:x:/g/personal/lhuang_ecopaxinc_com/EdUeZd99AJJFn9GIOYn6s74BxnH194ZxUJq5DfehBlI9EQ?e=TlZyOt&download=1"
    output_file = "Docking_System_-_Pando_(V5).xlsx"
    
    # Download the Excel file
    file_path = download_excel_file(url, output_file)
    if file_path is None:
        print("Exiting due to download error.")
        return
    
    # Get the list of available drivers
    drivers = get_available_drivers(file_path)
    if drivers:
        print("\nAvailable Drivers:")
        for driver in drivers:
            print(f"- {driver}")
    else:
        print("No available drivers found.")

if __name__ == "__main__":
    main()
