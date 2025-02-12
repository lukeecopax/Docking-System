"""
Driver Collaboration Analysis Tool

This script analyzes collaboration patterns between drivers and loaders based on Excel data.
It downloads an Excel file from a specified URL and processes the data to show how often
a selected driver works with other team members.
"""

import logging
from typing import List, Optional, Tuple
import pandas as pd
import requests
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Configuration
CONFIG = {
    'excel_url': ("https://netorg245148-my.sharepoint.com/:x:/g/"
                  "personal/lhuang_ecopaxinc_com/EdUeZd99AJJFn9GIOYn6s74BxnH194ZxUJq5DfehBlI9EQ?"
                  "e=TlZyOt&download=1"),
    'output_file': "Docking_System_-_Pando_(V5).xlsx",
    'sheets': {
        'records': "Cleaned Archive",
        'drivers': "Driver"
    }
}

def download_excel_file(url: str, output_path: str) -> bool:
    """
    Download an Excel file from the specified URL.

    Args:
        url (str): The URL to download the file from
        output_path (str): The path where the file should be saved

    Returns:
        bool: True if download was successful, False otherwise
    """
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        
        with open(output_path, "wb") as f:
            f.write(response.content)
        logger.info("Excel file downloaded successfully")
        return True
    
    except requests.RequestException as e:
        logger.error(f"Error downloading file: {str(e)}")
        return False
    except IOError as e:
        logger.error(f"Error saving file: {str(e)}")
        return False

def load_excel_data() -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
    """
    Load and validate the Excel data from the specified sheets.

    Returns:
        Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]: Cleaned records and driver DataFrames
    """
    try:
        df_clean = pd.read_excel(
            CONFIG['output_file'],
            sheet_name=CONFIG['sheets']['records'],
            engine="openpyxl"
        )
        driver_df = pd.read_excel(
            CONFIG['output_file'],
            sheet_name=CONFIG['sheets']['drivers'],
            engine="openpyxl"
        )
        return df_clean, driver_df
    except Exception as e:
        logger.error(f"Error loading Excel data: {str(e)}")
        return None, None

def get_valid_drivers(driver_df: pd.DataFrame) -> List[str]:
    """
    Process the driver sheet to get a list of valid driver names.

    Args:
        driver_df (pd.DataFrame): DataFrame containing driver information

    Returns:
        List[str]: List of valid driver names
    """
    try:
        # Clean and standardize the data
        driver_df = driver_df.copy()
        driver_df["Available"] = driver_df["Available"].astype(str).str.strip().str.upper()
        driver_df["Driver Name"] = driver_df["Driver Name"].astype(str).str.strip()
        
        # Filter for available drivers
        valid_drivers = driver_df[
            driver_df["Available"] == "YES"
        ]["Driver Name"].dropna().unique().tolist()
        
        # Remove invalid entries
        valid_drivers = [d for d in valid_drivers if d.lower() != "nan" and d != ""]
        
        return valid_drivers
    except Exception as e:
        logger.error(f"Error processing driver data: {str(e)}")
        return []

def analyze_collaborations(
    df_clean: pd.DataFrame,
    selected_driver: str,
    valid_drivers: List[str]
) -> Optional[pd.DataFrame]:
    """
    Analyze collaboration patterns for a selected driver.

    Args:
        df_clean (pd.DataFrame): DataFrame containing the cleaned records
        selected_driver (str): Name of the driver to analyze
        valid_drivers (List[str]): List of valid driver names

    Returns:
        Optional[pd.DataFrame]: DataFrame containing collaboration statistics
    """
    try:
        # Validate input
        if selected_driver not in valid_drivers:
            logger.error(f"Selected driver '{selected_driver}' not found in valid drivers list")
            return None

        # Filter records for selected driver
        collab_base = df_clean[
            (df_clean["Driver"] == selected_driver) |
            (df_clean["Loader 1"] == selected_driver) |
            (df_clean["Loader 2"] == selected_driver)
        ]

        if collab_base.empty:
            logger.warning(f"No records found for driver: {selected_driver}")
            return None

        # Reshape data to list all collaborators
        melted = collab_base.melt(
            id_vars=[col for col in collab_base.columns if col not in ["Driver", "Loader 1", "Loader 2"]],
            value_vars=["Driver", "Loader 1", "Loader 2"],
            value_name="Collaborator"
        )
        melted["Collaborator"] = melted["Collaborator"].astype(str).str.strip()

        # Filter and count collaborations
        collaborations = melted[
            (melted["Collaborator"] != selected_driver) &
            (melted["Collaborator"].isin(valid_drivers))
        ]
        
        collab_counts = collaborations.groupby("Collaborator").size().reset_index(name="Collaborations")
        collab_counts = collab_counts.sort_values("Collaborations", ascending=False)
        
        return collab_counts

    except Exception as e:
        logger.error(f"Error analyzing collaborations: {str(e)}")
        return None

def main(selected_driver: str = "Carlos Rodriguez Escobar"):
    """
    Main function to orchestrate the collaboration analysis process.

    Args:
        selected_driver (str): Name of the driver to analyze
    """
    # Ensure output file directory exists
    Path(CONFIG['output_file']).parent.mkdir(parents=True, exist_ok=True)

    # Download Excel file if it doesn't exist
    if not Path(CONFIG['output_file']).exists():
        if not download_excel_file(CONFIG['excel_url'], CONFIG['output_file']):
            return

    # Load Excel data
    df_clean, driver_df = load_excel_data()
    if df_clean is None or driver_df is None:
        return

    # Get valid drivers
    valid_drivers = get_valid_drivers(driver_df)
    if not valid_drivers:
        logger.error("No valid drivers found")
        return

    # Analyze collaborations
    collab_counts = analyze_collaborations(df_clean, selected_driver, valid_drivers)
    
    # Output results
    if collab_counts is None or collab_counts.empty:
        logger.info(f"No sufficient collaboration data found for {selected_driver}")
    else:
        print(f"\nCollaborations for {selected_driver}:")
        print(collab_counts.to_string(index=False))
        print(f"\nTotal unique collaborators: {len(collab_counts)}")
        print(f"Total collaborations: {collab_counts['Collaborations'].sum()}")

if __name__ == "__main__":
    main()
