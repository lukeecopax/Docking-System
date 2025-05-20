# Added new Performance Rankings report to view all individuals ranked by load type
# Enhanced Individual Performance with detailed "Show All Loads" feature including duration comparisons
# Expanded Team Performance report to support all load types instead of just "Import 40"


import streamlit as st
import warnings
import requests
import pandas as pd
import numpy as np
from datetime import datetime
import streamlit.components.v1 as components
import matplotlib.pyplot as plt
from io import BytesIO
import logging
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import calendar
from datetime import timedelta

# Set page config and suppress warnings
st.set_page_config(layout="wide")
warnings.filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

##################################
# Part 1: Download Files
##################################
def download_files():
    # --- Download 2025 Active File via SharePoint ---
    site_url = st.secrets["sharepoint"]["site_url"]
    username = st.secrets["sharepoint"]["username"]
    password = st.secrets["sharepoint"]["password"]
    
    ctx_auth = AuthenticationContext(site_url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(site_url, ctx_auth)
        file_relative_url = "/sites/PandoShipping/Shared Documents/Bethlehem Shipping/Docking System - Pando 2025.xlsm"
        try:
            with open("Docking System - Pando 2025.xlsm", "wb") as local_file:
                file_obj = ctx.web.get_file_by_server_relative_url(file_relative_url)
                file_obj.download(local_file).execute_query()
            logging.info("2025 file downloaded successfully!")
        except Exception as e:
            logging.error("Error downloading 2025 file: %s", e)
    else:
        logging.error("Authentication failed for 2025 file.")
    
    # --- Download 2024 Archived File via Direct URL ---
    direct_download_url = "https://netorg245148-my.sharepoint.com/:x:/g/personal/lhuang_ecopaxinc_com/EdUeZd99AJJFn9GIOYn6s74BxnH194ZxUJq5DfehBlI9EQ?e=TlZyOt&download=1"
    output_file = "Docking System - Pando (V5).xlsx"
    try:
        response = requests.get(direct_download_url)
        with open(output_file, "wb") as f:
            f.write(response.content)
        logging.info("2024 archived file downloaded successfully!")
    except Exception as e:
        logging.error("Error downloading 2024 file: %s", e)

##################################
# Part 2: Helper Functions for Cleaning
##################################
def clean_load_type_column(df, col_name="Load Type"):
    """Cleans the given load type column: trims whitespace, converts to title case, 
       and replaces variants of 'Ltl' with 'LTL'."""
    df[col_name] = (
        df[col_name]
        .astype(str)
        .str.strip()
        .replace('nan', np.nan)
        .str.title()
        .str.replace(r'^Ltl', 'LTL', regex=True)
    )
    return df

def compute_actual_duration(df):
    """Compute datetime columns and the actual duration (in minutes) using the Day Archive, Start Time, and Actual End Time."""
    df['Day Archive'] = pd.to_datetime(df['Day Archive'], errors='coerce')
    df['Start Datetime'] = pd.to_datetime(
        df['Day Archive'].dt.strftime('%Y-%m-%d') + ' ' + df['Start Time'].astype(str),
        errors='coerce'
    )
    df['Actual End Datetime'] = pd.to_datetime(
        df['Day Archive'].dt.strftime('%Y-%m-%d') + ' ' + df['Actual End Time'].astype(str),
        errors='coerce'
    )
    df['Actual Duration'] = ((df['Actual End Datetime'] - df['Start Datetime'])
                             .dt.total_seconds() / 60)
    return df

def drop_column_if_exists(df, col):
    """Drop a column if it exists in the DataFrame."""
    if col in df.columns:
        df.drop(columns=[col], inplace=True)
    return df

##################################
# Part 3: Define Cleaning Functions for Each Dataset
##################################
def clean_2024_archive(file_path):
    # Read sheets from the 2024 file
    archive_df = pd.read_excel(file_path, sheet_name="Cleaned Archive", engine="openpyxl")
    driver_df  = pd.read_excel(file_path, sheet_name="Driver", engine="openpyxl")
    load_type_df = pd.read_excel(file_path, sheet_name="Load Type", engine="openpyxl")
    
    archive_df = clean_load_type_column(archive_df, "Load Type")
    load_type_df = clean_load_type_column(load_type_df, "Load Type")
    
    # Merge archive with load type info
    archive_merged_df = archive_df.merge(load_type_df, how="left", on="Load Type")
    
    # Create individual driver DataFrame (from the Driver sheet)
    individual_df = pd.DataFrame(driver_df["Driver Name"].dropna().unique(), columns=["Driver Name"])
    
    # Standardize column names (title case)
    archive_merged_df.columns = archive_merged_df.columns.str.strip().str.title()
    
    # Drop unwanted columns
    archive_merged_df = drop_column_if_exists(archive_merged_df, "Minute Late (Neg = Early)")
    
    # Compute Actual Duration
    archive_merged_df = compute_actual_duration(archive_merged_df)
    
    # Additional cleaning
    archive_merged_df = archive_merged_df[archive_merged_df['Actual Duration'] >= 0]
    archive_merged_df.drop_duplicates(inplace=True)
    archive_merged_df = drop_column_if_exists(archive_merged_df, "Start Time")
    archive_merged_df = drop_column_if_exists(archive_merged_df, "Actual End Time")
    
    return archive_merged_df, individual_df

def clean_2025_archive(file_path):
    # Read sheets from the 2025 file
    archive_df = pd.read_excel(file_path, sheet_name="Archive", engine="openpyxl")
    driver_df  = pd.read_excel(file_path, sheet_name="Driver", engine="openpyxl")
    load_type_df = pd.read_excel(file_path, sheet_name="Load Type", engine="openpyxl")
    
    archive_df = clean_load_type_column(archive_df, "Load Type")
    # Note: In load_type_df, the key column is "Load type" (with lowercase t)
    load_type_df = clean_load_type_column(load_type_df, "Load type")
    
    # Merge archive with load type info (using different key names)
    archive_merged_df = archive_df.merge(load_type_df, how="left", left_on="Load Type", right_on="Load type")
    archive_merged_df = drop_column_if_exists(archive_merged_df, "Load type")
    
    individual_df = pd.DataFrame(driver_df["Driver Name"].dropna().unique(), columns=["Driver Name"])
    
    archive_merged_df.columns = archive_merged_df.columns.str.strip().str.title()
    
    # Drop unwanted columns for 2025
    archive_merged_df = drop_column_if_exists(archive_merged_df, "Column1")
    archive_merged_df = drop_column_if_exists(archive_merged_df, "Minute Late (Neg = Early)")
    
    archive_merged_df = compute_actual_duration(archive_merged_df)
    
    archive_merged_df = archive_merged_df[archive_merged_df['Actual Duration'] >= 0]
    archive_merged_df.drop_duplicates(inplace=True)
    archive_merged_df = drop_column_if_exists(archive_merged_df, "Start Time")
    archive_merged_df = drop_column_if_exists(archive_merged_df, "Actual End Time")
    
    return archive_merged_df, individual_df

##################################
# Part 4: Clean Both Datasets and Merge
##################################
def clean_and_merge_datasets():
    # Clean the 2024 and 2025 files
    archive_2024, individual_2024 = clean_2024_archive("Docking System - Pando (V5).xlsx")
    archive_2025, individual_2025 = clean_2025_archive("Docking System - Pando 2025.xlsm")
    
    # Merge the two archive DataFrames (vertical concatenation)
    merged_archive = pd.concat([archive_2024, archive_2025], ignore_index=True)
    
    # Combine the individual driver lists
    individual_df_combined = pd.concat([individual_2024, individual_2025]).drop_duplicates().reset_index(drop=True)
    
    # Clean and standardize driver names
    individual_df_combined["Driver Name"] = individual_df_combined["Driver Name"].apply(lambda n: n.strip().title())
    all_individuals = set(individual_df_combined["Driver Name"].unique())
    
    return merged_archive, individual_df_combined, all_individuals

##################################
# Part 5: Compute Statistics and Available Individuals
##################################
def compute_statistics(merged_archive, all_individuals):
    # Compute year distribution
    merged_archive['Year'] = merged_archive['Day Archive'].dt.year
    year_distribution = merged_archive['Year'].value_counts().sort_index()
    
    # Get unique load types
    unique_load_types = merged_archive['Load Type'].dropna().unique()
    
    # Determine available individuals (active in the past month)
    today = pd.to_datetime("2025-02-01")
    past_month = today - pd.DateOffset(months=1)
    
    recent_records = merged_archive[(merged_archive["Day Archive"] >= past_month) &
                                    (merged_archive["Day Archive"] <= today)]
    
    individual_cols = ["Driver", "Loader 1", "Loader 2"]
    active_names = set()
    for col in individual_cols:
        if col in recent_records.columns:
            active_names.update(recent_records[col].dropna().unique())
    
    active_names_cleaned = {n.strip().title() for n in active_names}
    available_individuals = all_individuals.intersection(active_names_cleaned)
    
    return year_distribution, unique_load_types, available_individuals

##################################
# Download and Clean Data
##################################
download_files()
merged_archive, individual_df_combined, all_individuals = clean_and_merge_datasets()
year_distribution, unique_load_types, available_individuals = compute_statistics(merged_archive, all_individuals)

# Create the cleaned DataFrame for the app (filter out "Assignment" rows)
df_clean = merged_archive[merged_archive["Load Type"].str.lower() != "assignment"].copy()

# Ensure date and numeric columns are properly formatted
df_clean["Day Archive"] = pd.to_datetime(df_clean["Day Archive"], errors="coerce")
df_clean["Actual Duration"] = pd.to_numeric(df_clean["Actual Duration"], errors="coerce")

# Standardize text fields for loaders
df_clean["Loader 1"] = df_clean["Loader 1"].fillna("").astype(str).str.strip()
df_clean["Loader 2"] = df_clean["Loader 2"].fillna("").astype(str).str.strip()

# Create role flag columns: is_solo and is_team
df_clean["is_solo"] = ((df_clean["Loader 1"].isin(["", "nan", "None"])) & (df_clean["Loader 2"].isin(["", "nan", "None"])))
df_clean["is_team"] = ~df_clean["is_solo"]

##################################
# Sidebar and UI Setup
##################################

st.title("Pando Docking Performance Reports v1.2")
st.markdown(
    """
    **Latest Update:** *May 20, 2025*  
    - Enhanced Team Averages vs Goals report with improved summary display
    - Fixed Year to Date functionality
    """
)

# Report type selection remains as before
report_choice = st.sidebar.selectbox(
    "Select Report Type:",
    [
        "Individual Performance",
        "Team Performance",
        "Performance Rankings",
        "Team Averages vs Goals"
    ],
    index=2
)

# Global Date Range Selection (inserted in the sidebar)
date_range_option = st.sidebar.selectbox(
    "Select Date Range Option",
    ["This Month", "Year to Date", "Year 2024", "Year 2025", "Last Month", "Custom"],
    index=0  # default is "This Month"
)

today = datetime.today()
current_year = today.year

if date_range_option == "Custom":
    start_date = st.sidebar.date_input("Start Date:", datetime(2025, 1, 1))
    end_date = st.sidebar.date_input("End Date:", datetime(2025, 6, 1))
elif date_range_option == "Year 2024":
    start_date = datetime(2024, 1, 1)
    end_date = datetime(2024, 12, 31)
elif date_range_option == "Year 2025":
    start_date = datetime(2025, 1, 1)
    end_date = datetime(2025, 12, 31)
elif date_range_option == "Year to Date":
    start_date = datetime(current_year, 1, 1)
    end_date = today
elif date_range_option == "Last Month":
    first_of_current = today.replace(day=1)
    last_month_end = first_of_current - timedelta(days=1)
    last_month_start = last_month_end.replace(day=1)
    start_date = last_month_start
    end_date = last_month_end
elif date_range_option == "This Month":
    start_date = today.replace(day=1)
    last_day = calendar.monthrange(today.year, today.month)[1]
    end_date = today.replace(day=last_day)

start_date_str = start_date.strftime("%Y-%m-%d")
end_date_str = end_date.strftime("%Y-%m-%d")

# Display the computed global date range beneath the select box on the sidebar with extra spacing
st.sidebar.markdown(
    f"<strong>Selected Date Range:</strong><br>{start_date_str} to {end_date_str}<br><br><br>",
    unsafe_allow_html=True
)

# Existing Individuals Scope selection and explanation remain unchanged
individual_scope = st.sidebar.radio("**Select Individuals Scope:**", ["Only Available Individuals", "All Individuals"])
st.sidebar.markdown(
    "<strong>Explanation:</strong><br><br>"
    "- <strong>Only Available Individuals:</strong> Displays only those individuals who have been active in the past month (i.e. they have recent records in the dataset).<br><br>"
    "- <strong>All Individuals:</strong> Includes every individual from the dataset, regardless of recent activity.<br><br>",
    unsafe_allow_html=True
)
if individual_scope == "Only Available Individuals":
    valid_drivers = sorted(list(available_individuals))
else:
    valid_drivers = sorted(list(all_individuals))

css = """
<style>
body {
    font-family: Arial, sans-serif;
}
table {
    max-width: 1000px;
    margin: 20px;
    padding: 20px;
    border: 1px solid #ddd;
    border-radius: 8px;
    border-collapse: separate;
    border-spacing: 20px;
    width: 100%;
}
th, td {
    padding: 12px 25px;
    border-bottom: 1px solid #ddd;
    white-space: nowrap;
}
th {
    font-weight: bold;
}
td {
    min-width: 30px;
}
.explanation {
    max-width: 1000px;
    margin: 20px;
    padding: 15px;
    line-height: 1.6;
    background-color: #f7f7f7;
    border: 1px solid #ddd;
    border-radius: 8px;
    font-size: 15px;
}
.modules-container {
    border: 2px solid #ddd;
    max-width: 1100px;
    border-radius: 8px;
    margin: 20px;
    padding: 20px;
}
.status-dot {
    display: inline-block;
    width: 12px;
    height: 12px;
    border-radius: 50%;
    margin-left: 5px;
}
.good { background-color: #4CAF50; }
.avg { background-color: #FFC107; }
.bad { background-color: #F44336; }
.teams-tables table {
    border-spacing: 10px !important;
}
.teams-tables th, .teams-tables td {
    padding: 8px 15px !important;
}
</style>
"""

##################################
# Helper Function: Vectorized Status Calculation
##################################
def compute_status_vectorized(avg_series, target_series, threshold=0.15):
    conditions = [avg_series < target_series,
                  (avg_series >= target_series) & ((avg_series - target_series) < threshold * target_series)]
    choices = ["good", "avg"]
    return np.select(conditions, choices, default="bad")

##################################
# Function: Process Individual/Team Performance (Vectorized)
##################################
def process_performance(individual_df, overall_df, module_title, driver, start_date_str, end_date_str, valid_drivers):
    # Group individual performance by Load Type
    indiv_grouped = individual_df.groupby("Load Type").agg(
        Average_Duration_Numeric=("Actual Duration", "mean"),
        Loads_Count=("Actual Duration", "count")
    ).reset_index()
   
    # Compute overall average performance for each Load Type
    overall_avg = overall_df.groupby("Load Type")["Actual Duration"].mean().reset_index().rename(columns={"Actual Duration": "Overall_Average"})
    indiv_grouped = indiv_grouped.merge(overall_avg, on="Load Type", how="left")
    indiv_grouped["vs_Average_Numeric"] = indiv_grouped["Average_Duration_Numeric"] - indiv_grouped["Overall_Average"]
   
    # Merge with target minutes from df_clean (assumes a column "Target Minutes" exists)
    merged_df = pd.merge(indiv_grouped, df_clean[["Load Type", "Target Minutes"]].drop_duplicates(), on="Load Type", how="left")
    merged_df["vs_Target_Numeric"] = merged_df["Average_Duration_Numeric"] - merged_df["Target Minutes"]
   
    # Compute statuses
    merged_df["target_status"] = compute_status_vectorized(merged_df["Average_Duration_Numeric"], merged_df["Target Minutes"])
    merged_df["average_status"] = compute_status_vectorized(merged_df["Average_Duration_Numeric"], merged_df["Overall_Average"])
   
    # Tooltips for statuses
    tooltip_target = np.select(
        [merged_df["target_status"]=="good", merged_df["target_status"]=="avg", merged_df["target_status"]=="bad"],
        ["Good: Performance is below target", "Average: Performance is slightly above target", "Bad: Performance is significantly above target"],
        default=""
    )
    tooltip_average = np.select(
        [merged_df["average_status"]=="good", merged_df["average_status"]=="avg", merged_df["average_status"]=="bad"],
        ["Good: Performance is below overall average", "Average: Performance is slightly above overall average", "Bad: Performance is significantly above overall average"],
        default=""
    )
   
    # Ensure each component is cast to string
    merged_df["vs Target"] = merged_df["vs_Target_Numeric"].map("{:.1f} min".format).astype(str) + \
        " <span class='status-dot " + merged_df["target_status"].astype(str) + \
        "' title='" + pd.Series(tooltip_target).astype(str) + "'></span>"
       
    merged_df["vs Average"] = merged_df["vs_Average_Numeric"].map("{:.1f} min".format).astype(str) + \
        " <span class='status-dot " + merged_df["average_status"].astype(str) + \
        "' title='" + pd.Series(tooltip_average).astype(str) + "'></span>"
   
    merged_df["Average Duration"] = merged_df["Average_Duration_Numeric"].map("{:.1f} min".format)
   
    # --- New Metric: vs Top ---
    if module_title == "Team":
        melt_df = overall_df.melt(id_vars=[col for col in overall_df.columns if col not in ["Driver", "Loader 1", "Loader 2"]],
                                   value_vars=["Driver", "Loader 1", "Loader 2"],
                                   value_name="Participant")
        melt_df["Participant"] = melt_df["Participant"].astype(str).str.strip()
        melt_df = melt_df[melt_df["Participant"].isin(valid_drivers)]
        eligible_df = melt_df.groupby(["Load Type", "Participant"]).agg(
            avg_duration=("Actual Duration", "mean"),
            loads_count=("Actual Duration", "count")
        ).reset_index().rename(columns={"Participant": "Driver"})
    else:
        eligible_df = overall_df[overall_df["Driver"].isin(valid_drivers)].groupby(["Load Type", "Driver"]).agg(
            avg_duration=("Actual Duration", "mean"),
            loads_count=("Actual Duration", "count")
        ).reset_index()
    eligible_df = eligible_df[eligible_df["loads_count"] > 5]
    top_df = eligible_df.groupby("Load Type")["avg_duration"].min().reset_index().rename(columns={"avg_duration": "top_avg"})
    driver_avg_df = eligible_df[eligible_df["Driver"] == driver][["Load Type", "avg_duration"]].rename(columns={"avg_duration": "driver_avg"})
    vs_top_df = pd.merge(driver_avg_df, top_df, on="Load Type", how="left")
    vs_top_df["vs_top_ratio"] = np.where(vs_top_df["driver_avg"] > 0, (vs_top_df["top_avg"] / vs_top_df["driver_avg"]) * 100, np.nan)
    merged_df = merged_df.merge(vs_top_df[["Load Type", "vs_top_ratio"]], on="Load Type", how="left")
    merged_df["vs Top"] = np.where(merged_df["Loads_Count"] > 5,
                                   merged_df["vs_top_ratio"].map(lambda x: f"{x:.0f}% of #1" if not pd.isna(x) else "N/A"),
                                   "N/A")
   
    # --- New Metric: Ranking ---
    if module_title == "Team":
        melt_df = overall_df.melt(id_vars=[col for col in overall_df.columns if col not in ["Driver", "Loader 1", "Loader 2"]],
                                   value_vars=["Driver", "Loader 1", "Loader 2"],
                                   value_name="Participant")
        melt_df["Participant"] = melt_df["Participant"].astype(str).str.strip()
        melt_df = melt_df[melt_df["Participant"].isin(valid_drivers)]
        ranking_df = melt_df.groupby(["Load Type", "Participant"]).agg(
            avg_duration=("Actual Duration", "mean"),
            loads_count=("Actual Duration", "count")
        ).reset_index().rename(columns={"Participant": "Driver"})
    else:
        ranking_df = overall_df[overall_df["Driver"].isin(valid_drivers)].groupby(["Load Type", "Driver"]).agg(
            avg_duration=("Actual Duration", "mean"),
            loads_count=("Actual Duration", "count")
        ).reset_index()
    ranking_df = ranking_df[ranking_df["loads_count"] > 5]
    ranking_df["Rank"] = ranking_df.groupby("Load Type")["avg_duration"].rank(method="min", ascending=True)
    driver_ranks = ranking_df[ranking_df["Driver"] == driver][["Load Type", "Rank"]]
    merged_df = merged_df.merge(driver_ranks, on="Load Type", how="left")
    merged_df["Ranking"] = merged_df["Rank"].apply(lambda x: f"Rank: {int(x)}" if not pd.isna(x) else "N/A")
    merged_df.drop(columns=["Rank"], inplace=True)
   
    # --- Overall Metrics ---
    total_loads = merged_df["Loads_Count"].sum()
    if total_loads > 0:
        weighted_avg_duration = (merged_df["Average_Duration_Numeric"] * merged_df["Loads_Count"]).sum() / total_loads
        weighted_target = (merged_df["Target Minutes"] * merged_df["Loads_Count"]).sum() / total_loads
        overall_diff = weighted_avg_duration - weighted_target
    else:
        weighted_avg_duration = weighted_target = overall_diff = 0
   
    total_status = compute_status_vectorized(pd.Series([weighted_avg_duration]), pd.Series([weighted_target]))[0]
    tooltip_total = {"good": "Good: Performance is below target",
                     "avg": "Average: Performance is slightly above target",
                     "bad": "Bad: Performance is significantly above target"}.get(total_status, "")
    total_vs_target = f"{overall_diff:+.1f} min <span class='status-dot {total_status}' title='{tooltip_total}'></span>"
   
    overall_module_avg = overall_df["Actual Duration"].mean() if not overall_df.empty else 0
    overall_diff_vs_avg = weighted_avg_duration - overall_module_avg
    total_status_vs_avg = compute_status_vectorized(pd.Series([weighted_avg_duration]), pd.Series([overall_module_avg]))[0]
    tooltip_total_vs_avg = {"good": "Good: Performance is below overall average",
                            "avg": "Average: Performance is slightly above overall average",
                            "bad": "Bad: Performance is significantly above overall average"}.get(total_status_vs_avg, "")
    total_vs_average = f"{overall_diff_vs_avg:+.1f} min <span class='status-dot {total_status_vs_avg}' title='{tooltip_total_vs_avg}'></span>"
   
    summary = {
        "Load Type": "Total",
        "Average Duration": f"{weighted_avg_duration:.1f} min",
        "# Loads": total_loads,
        "Target Minutes": f"{weighted_target:.1f}",
        "vs Target": total_vs_target,
        "vs Average": total_vs_average,
        "vs Top": "",
        "Ranking": ""
    }
    summary_df = pd.DataFrame([summary])
    merged_df = pd.concat([merged_df, summary_df], ignore_index=True)
   
    merged_df = merged_df[["Load Type", "Average Duration", "Loads_Count", "vs Target", "vs Average", "vs Top", "Ranking"]]
    merged_df.rename(columns={"Loads_Count": "# Loads"}, inplace=True)
   
    module_html = f"<h2>{module_title} Performance for {driver}</h2>"
    module_html += f"<h3>Date Range: <strong>{start_date_str}</strong> to <strong>{end_date_str}</strong></h3>"
   
    if module_title == "Solo":
        explanation_text = (
            "This table displays your solo performance—operating as the sole driver without any additional loaders. "
            "It presents the average duration for each load type and compares your performance against both the overall solo averages and the target goals, along with additional insights "
            "such as 'vs Top' and 'Ranking'.<br><br>"
            "<span class='status-dot good'></span> <strong>Green:</strong> Performance is better than (i.e., lower than) the target or overall average.<br>"
            "<span class='status-dot avg'></span> <strong>Yellow:</strong> Performance is within 15% of the target or overall average.<br>"
            "<span class='status-dot bad'></span> <strong>Red:</strong> Performance deviates by 15% or more from the target or overall average.<br><br>"
            "<em>Note: For each load type, you must have at least 5 load records to be eligible for ranking.</em>"
        )
    elif module_title == "Team":
        explanation_text = (
            "This table displays your performance when you work as part of a team (as either a driver or loader). "
            "It includes metrics comparing your team's averages to the overall team averages, along with additional insights "
            "such as 'vs Top' and 'Ranking'.<br><br>"
            "<span class='status-dot good'></span> <strong>Green:</strong> Performance is better than (i.e., lower than) the target or overall average.<br>"
            "<span class='status-dot avg'></span> <strong>Yellow:</strong> Performance is within 15% of the target or overall average.<br>"
            "<span class='status-dot bad'></span> <strong>Red:</strong> Performance deviates by 15% or more from the target or overall average.<br><br>"
            "<em>Note: For each load type, you must have at least 5 load records to be eligible for ranking.</em>"
        )
   
    module_html += f"<div class='explanation'>{explanation_text}</div>"
    module_html += merged_df.to_html(index=False, escape=False)
    return module_html

##################################
# Function: Generate Team Performance Report (2024)
##################################
def generate_team_report(load_type_filter="Import 40", min_loads=10, valid_drivers=None):
    # In the Team Performance Report Section:
    df_team = df_clean.copy()
    df_team = df_team[df_team["Driver"].notna() & (df_team["Driver"].str.strip() != "")]
    df_team = df_team[df_team["Loader 1"].notna() | df_team["Loader 2"].notna()]
    # Use the global date range for filtering
    df_team = df_team[(df_team["Day Archive"] >= pd.to_datetime(start_date)) & (df_team["Day Archive"] <= pd.to_datetime(end_date))]
    
    df_team["Diff"] = df_team["Actual Duration"] - df_team["Target Minutes"]
   
    df_team = df_team[df_team["Load Type"] == load_type_filter]
    team_group = df_team.groupby(["Driver", "Loader 1", "Loader 2"]).agg(
        avg_duration=("Actual Duration", "mean"),
        total_diff=("Diff", "sum"),
        avg_diff=("Diff", "mean"),
        load_count=("Diff", "count")
    ).reset_index()
   
    teams_valid = team_group[team_group["load_count"] >= min_loads]
    top_teams = teams_valid.nsmallest(10, "avg_diff") if not teams_valid.empty else None
   
    df_team["Diff"] = df_team["Actual Duration"] - df_team["Target Minutes"]
    melted = df_team.melt(id_vars=[col for col in df_team.columns if col not in ["Driver", "Loader 1", "Loader 2"]],
                           value_vars=["Driver", "Loader 1", "Loader 2"],
                           value_name="Team Member")
    melted["Team Member"] = melted["Team Member"].astype(str).str.strip()
    melted = melted[melted["Team Member"].isin(valid_drivers)]
    individual_stats = melted.groupby("Team Member").agg(
        avg_diff=("Diff", "mean"),
        load_count=("Diff", "count")
    ).reset_index()
    individual_stats = individual_stats[individual_stats["load_count"] >= 5]
    overall_avg_diff = df_team["Diff"].mean()
    individual_stats["plus_minus"] = individual_stats["avg_diff"] - overall_avg_diff
    individual_stats = individual_stats.sort_values("plus_minus", ascending=False)
   
    html_report = css
    html_report += """
    <h2>Team Performance Report for Load Type: {}</h2>
    <div class="explanation" style="background-color: #fff3cd; border-left: 5px solid #ffa000; padding: 10px; margin-bottom: 20px;">
      <p><strong>Note:</strong> Each team must have at least <strong>{}</strong> load records to be eligible for review.</p>
    </div>
    <p>This report identifies the top 10 best-performing teams and analyzes individual contributions using a plus/minus metric.</p>
    """.format(load_type_filter, min_loads)
   
    if top_teams is not None and not teams_valid.empty:
        html_report += "<div class='teams-tables'><h2>🏆 Top 10 Best Teams</h2>"
        html_report += "<table><tr><th>Rank</th><th>Driver</th><th>Loader 1</th><th>Loader 2</th><th>Loads Count</th><th>Avg Duration</th><th>Total Diff</th><th>Avg Diff</th></tr>"
        for i, (_, team) in enumerate(top_teams.iterrows(), 1):
            html_report += f"""
            <tr>
                <td>#{i}</td>
                <td>{team["Driver"]}</td>
                <td>{team["Loader 1"]}</td>
                <td>{team["Loader 2"]}</td>
                <td>{team["load_count"]}</td>
                <td>{team["avg_duration"]:.1f} min</td>
                <td>{team["total_diff"]:.1f} min</td>
                <td>{team["avg_diff"]:.1f} min</td>
            </tr>
            """
        html_report += "</table></div>"
    else:
        html_report += "<p>No teams met the minimum load requirement.</p>"
   
    worst_teams = teams_valid.nlargest(10, "avg_diff")
    # Avoid overlap: remove teams that are in the top_teams from worst_teams
    if top_teams is not None and not worst_teams.empty:
        overlap_indices = top_teams.index.intersection(worst_teams.index)
        if not overlap_indices.empty:
            worst_teams = worst_teams[~worst_teams.index.isin(overlap_indices)]
   
    if not worst_teams.empty:
        html_report += "<div class='teams-tables'><h2>💀 Bottom 10 Teams (Worst Performance)</h2>"
        html_report += "<table><tr><th>Rank</th><th>Driver</th><th>Loader 1</th><th>Loader 2</th><th>Loads Count</th><th>Avg Duration</th><th>Total Diff</th><th>Avg Diff</th></tr>"
        for i, (_, team) in enumerate(worst_teams.iterrows(), 1):
            html_report += f"""
            <tr>
                <td>#{i}</td>
                <td>{team["Driver"]}</td>
                <td>{team["Loader 1"]}</td>
                <td>{team["Loader 2"]}</td>
                <td>{team["load_count"]}</td>
                <td>{team["avg_duration"]:.1f} min</td>
                <td>{team["total_diff"]:.1f} min</td>
                <td>{team["avg_diff"]:.1f} min</td>
            </tr>
            """
        html_report += "</table></div>"
    else:
        html_report += "<p>No teams met the minimum load requirement for worst performance analysis.</p>"
   
    html_report += "<h2>⏳ Individual Performance</h2>"
    html_report += """
    <div class="explanation">
      <h3>Plus/Minus Concept Explained</h3>
      <p><strong>TL;DR:</strong> Individuals with the highest plus/minus values tend to slow down the process. Negative values indicate efficiency.</p>
      <p>This concept is borrowed from NBA analytics, where a player's plus/minus score reflects their impact on the team's overall performance. Here, it is calculated as the difference between an individual's average load completion difference and the overall average.</p>
      <p>Interpretation:</p>
      <ul>
        <li><span class="status-dot" style="background-color: green;"></span> Negative plus/minus indicates faster load completions than average.</li>
        <li><span class="status-dot" style="background-color: red;"></span> Positive plus/minus indicates slower load completions than average.</li>
      </ul>
    </div>
    """
    html_report += "<table><tr><th>Team Member</th><th>Loads</th><th>Avg Diff (min)</th><th>Plus/Minus</th></tr>"
    for _, row in individual_stats.iterrows():
        color = "green" if row["plus_minus"] < 0 else "red"
        html_report += f"""
        <tr>
            <td>{row["Team Member"]}</td>
            <td>{row["load_count"]}</td>
            <td>{row["avg_diff"]:.1f} min</td>
            <td>{row["plus_minus"]:.1f} <span class="status-dot" style="background-color: {color};"></span></td>
        </tr>
        """
    html_report += "</table>"
    html_report += f"<p><strong>Overall Average Diff:</strong> {overall_avg_diff:.1f} min</p>"
    return html_report

##################################
# Individual Performance Report Section
##################################
if report_choice == "Individual Performance":
    st.header("Individual Performance Report")
   
    if not valid_drivers:
        st.error("No individuals found in the Driver sheet.")
        st.stop()
   
    # In your Individual Performance Report Section, after the driver selectbox:
    default_driver = "Carlos Rodriguez Escobar"
    if default_driver in valid_drivers:
        default_index = valid_drivers.index(default_driver)
    else:
        default_index = 0
    driver = st.selectbox("Name:", valid_drivers, index=default_index)
    
    team_role_filter = st.selectbox("Team Role Filter (applied to Team Performance):", ["Either", "Driver Only", "Loader Only"])
   
    df = df_clean.copy()
    date_filtered_df = df[(df["Day Archive"] >= pd.to_datetime(start_date)) & (df["Day Archive"] <= pd.to_datetime(end_date))]
   
    # Workload Summary Card for the Selected Individual
    workload_df = df[(df["Day Archive"] >= pd.to_datetime(start_date)) & (df["Day Archive"] <= pd.to_datetime(end_date))].copy()
    melted_workload = workload_df.melt(id_vars=[col for col in workload_df.columns if col not in ["Driver", "Loader 1", "Loader 2"]],
                                       value_vars=["Driver", "Loader 1", "Loader 2"],
                                       value_name="Involved")
    melted_workload["Involved"] = melted_workload["Involved"].astype(str).str.strip()
    melted_workload = melted_workload[melted_workload["Involved"].isin(valid_drivers)]
    individual_metrics = melted_workload.groupby("Involved").agg(
        load_count=('Involved', 'count'),
        total_duration=('Actual Duration', 'sum')
    ).reset_index()
   
    selected_stats = individual_metrics[individual_metrics["Involved"] == driver]
    if not selected_stats.empty:
        load_count = selected_stats["load_count"].iloc[0]
        total_duration = selected_stats["total_duration"].iloc[0]
    else:
        load_count = 0
        total_duration = 0
   
    median_load_count = individual_metrics["load_count"].median() if not individual_metrics.empty else 0
    avg_load_count = individual_metrics["load_count"].mean() if not individual_metrics.empty else 0
    median_total_duration = individual_metrics["total_duration"].median() if not individual_metrics.empty else 0
    avg_total_duration = individual_metrics["total_duration"].mean() if not individual_metrics.empty else 0
   
    percent_vs_median_load = (load_count / median_load_count * 100) if median_load_count else 0
    percent_vs_avg_load = (load_count / avg_load_count * 100) if avg_load_count else 0
    percent_vs_median_duration = (total_duration / median_total_duration * 100) if median_total_duration else 0
    percent_vs_avg_duration = (total_duration / avg_total_duration * 100) if avg_total_duration else 0
   
    card_html = f"""
    <div class='explanation' style="background-color: #dcecf8; font-size: 17px;">
        <h3>Workload Summary for {driver}</h3>
        <p><strong>Total Loads:</strong> {load_count}</p>
        <p><strong>Total Duration:</strong> {total_duration:.1f} min</p>
        <p><strong>Load Count Comparison:</strong> {percent_vs_median_load:.1f}% of Median ({median_load_count:.1f} loads), {percent_vs_avg_load:.1f}% of Average ({avg_load_count:.1f} loads)</p>
        <p><strong>Total Duration Comparison:</strong> {percent_vs_median_duration:.1f}% of Median ({median_total_duration:.1f} min), {percent_vs_avg_duration:.1f}% of Average ({avg_total_duration:.1f} min)</p>
        <p>
            <strong>Explanation:</strong> <br>
            <strong>Total Loads</strong> represents the number of load events you were involved in (as a driver or loader), while <strong>Total Duration</strong> is the sum of the durations for those loads. The <strong>Load Count Comparison</strong> and <strong>Total Duration Comparison</strong> metrics indicate how your performance measures against the overall median and average values among your peers.
        </p>
    </div>
    """
    st.markdown(card_html, unsafe_allow_html=True)
   
    individual_loads_df = date_filtered_df[date_filtered_df.apply(
        lambda row: driver in {str(row["Driver"]).strip(), str(row["Loader 1"]).strip(), str(row["Loader 2"]).strip()},
        axis=1
    )]
    individual_counts = individual_loads_df.groupby("Load Type").size()
   
    melted_counts = date_filtered_df.melt(id_vars=[col for col in date_filtered_df.columns if col not in ["Driver", "Loader 1", "Loader 2"]],
                                          value_vars=["Driver", "Loader 1", "Loader 2"],
                                          value_name="Involved")
    melted_counts["Involved"] = melted_counts["Involved"].astype(str).str.strip()
    melted_counts = melted_counts[melted_counts["Involved"].isin(valid_drivers)]
    individual_loads_all = melted_counts.groupby(["Load Type", "Involved"]).size().reset_index(name="load_count")
    team_avg = individual_loads_all.groupby("Load Type")["load_count"].mean()
   
    counts_df = pd.concat([individual_counts, team_avg], axis=1, keys=['Individual', 'Team_Avg']).fillna(0).reset_index()
    counts_df = counts_df.sort_values(by='Individual', ascending=False)
   
    fig, ax = plt.subplots(figsize=(10, 6), dpi=80)
    bar_width = 0.35
    y_pos = np.arange(len(counts_df))
    ax.barh(y_pos - bar_width/2, counts_df['Individual'], height=bar_width, color='#086ccc', label=driver)
    ax.barh(y_pos + bar_width/2, counts_df['Team_Avg'], height=bar_width, color='gray', label='Team Average')
    ax.set_xlabel('Number of Loads')
    ax.set_title(f'Total Loads Handled per Category vs Team Average for {driver}')
    ax.set_yticks(y_pos)
    ax.set_yticklabels(counts_df['Load Type'])
    ax.invert_yaxis()
    ax.legend(loc='best')
    plt.tight_layout()
   
    buf = BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    st.image(buf, width=1000)
   
    overall_solo_df = date_filtered_df[df_clean["is_solo"]]
    solo_df = overall_solo_df[overall_solo_df["Driver"] == driver]
    overall_team_df = date_filtered_df[~df_clean["is_solo"]]
   
    if team_role_filter == "Driver Only":
        team_df = overall_team_df[overall_team_df["Driver"] == driver]
    elif team_role_filter == "Loader Only":
        team_df = overall_team_df[(overall_team_df["Loader 1"] == driver) | (overall_team_df["Loader 2"] == driver)]
    else:
        team_df = overall_team_df[(overall_team_df["Driver"] == driver) |
                                  (overall_team_df["Loader 1"] == driver) |
                                  (overall_team_df["Loader 2"] == driver)]
   
    solo_html = process_performance(solo_df, overall_solo_df, "Solo", driver, start_date_str, end_date_str, valid_drivers)
    team_html = process_performance(team_df, overall_team_df, "Team", driver, start_date_str, end_date_str, valid_drivers)
   
    final_html = css + "<div class='modules-container'>" + solo_html + "<br><br>" + team_html + "</div>"
    st.markdown(final_html, unsafe_allow_html=True)
   
    def collaborator_analysis(df, selected_driver, valid_drivers, min_collabs=5):
        team_records = df[~(
            (df["Loader 1"].str.strip().isin(["", "nan"])) &
            (df["Loader 2"].str.strip().isin(["", "nan"]))
        )].copy()
 
        team_records = team_records[
            (team_records["Driver"] == selected_driver) |
            (team_records["Loader 1"] == selected_driver) |
            (team_records["Loader 2"] == selected_driver)
        ]
 
        if "Target Minutes" not in team_records.columns:
            team_records = pd.merge(team_records, df_clean[["Load Type", "Target Minutes"]].drop_duplicates(),
                                  on="Load Type", how="left")
       
        if "Diff" not in team_records.columns:
            team_records["Diff"] = team_records["Actual Duration"] - team_records["Target Minutes"]
 
        collab_data = {}
       
        for _, row in team_records.iterrows():
            team_members = {
                str(row["Driver"]).strip(),
                str(row["Loader 1"]).strip(),
                str(row["Loader 2"]).strip()
            }
            team_members = {m for m in team_members if m and m.lower() != "nan" and m != "None"}
           
            if selected_driver in team_members:
                for member in team_members:
                    if member != selected_driver and member in valid_drivers:
                        if member not in collab_data:
                            collab_data[member] = []
                        collab_data[member].append(row["Diff"])
 
        collab_records = []
        for collaborator, diffs in collab_data.items():
            if len(diffs) >= 5:
                avg_diff = np.mean(diffs)
                collab_count = len(diffs)
                collab_records.append({
                    "Collaborator": collaborator,
                    "Collaborations": collab_count,
                    "Average_Diff": round(avg_diff, 1)
                })
 
        if not collab_records:
            return None, None
 
        collab_df = pd.DataFrame(collab_records)
        total_collaborators = len(collab_df)
 
        if total_collaborators >= 10:
            best_collab = collab_df.nsmallest(5, "Average_Diff")
            worst_collab = collab_df.nlargest(5, "Average_Diff")
        else:
            mid_point = total_collaborators // 2
            best_collab = collab_df.nsmallest(mid_point, "Average_Diff")
            worst_collab = collab_df.nlargest(total_collaborators - mid_point, "Average_Diff")
            overlap_indices = best_collab.index.intersection(worst_collab.index)
            if not overlap_indices.empty:
                worst_collab = worst_collab[~worst_collab.index.isin(overlap_indices)]
       
        return best_collab, worst_collab
   
    best_collab, worst_collab = collaborator_analysis(date_filtered_df, driver, valid_drivers, min_collabs=5)
   
    collab_html = "<div class='modules-container'>"
    collab_html += "<h2>Collaborator Analysis for {}</h2>".format(driver)
    collab_html += """
    <div class='explanation'>
        <p><strong>Average Diff (min):</strong> This value shows, on average, how many minutes the actual load durations differed from the target time when working with each collaborator.
        A lower (or negative) number means loads were generally completed faster than the target, while a higher number indicates loads took longer.</p>
        <br><em>Note:</em> Only collaborators who have been involved in at least 5 team records are included in this analysis.</p>
    </div>
    """
    if best_collab is not None and not best_collab.empty:
        collab_html += "<h3>Best Collaborators</h3>"
        collab_html += best_collab.to_html(index=False, escape=False)
    else:
        collab_html += "<p>No sufficient collaboration data available to determine best collaborators.</p>"
   
    if worst_collab is not None and not worst_collab.empty:
        collab_html += "<h3>Poor Collaborators</h3>"
        collab_html += worst_collab.to_html(index=False, escape=False)
    else:
        collab_html += "<p>No sufficient collaboration data available to determine poor collaborators.</p>"
    collab_html += "</div>"
   
    st.markdown(collab_html, unsafe_allow_html=True)

    # Add toggle to show all loads
    show_all_loads = st.checkbox("Show All Loads Completed by This Individual")

    if show_all_loads:
        st.subheader(f"All Loads Completed by {driver}")
        
        # Get all loads where this person was involved
        all_individual_loads = date_filtered_df[
            (date_filtered_df["Driver"] == driver) | 
            (date_filtered_df["Loader 1"] == driver) | 
            (date_filtered_df["Loader 2"] == driver)
        ].copy()
        
        if not all_individual_loads.empty:
            # Calculate comparison metrics
            all_individual_loads["vs_target"] = all_individual_loads["Actual Duration"] - all_individual_loads["Target Minutes"]
            
            # Calculate avg for each load type to compare
            load_type_avgs = date_filtered_df.groupby("Load Type")["Actual Duration"].mean().to_dict()
            all_individual_loads["load_type_avg"] = all_individual_loads["Load Type"].map(load_type_avgs)
            all_individual_loads["vs_avg"] = all_individual_loads["Actual Duration"] - all_individual_loads["load_type_avg"]
            
            # Format dates and determine role
            all_individual_loads["Day"] = all_individual_loads["Day Archive"].dt.strftime("%Y-%m-%d")
            all_individual_loads["Role"] = "Unknown"
            all_individual_loads.loc[all_individual_loads["Driver"] == driver, "Role"] = "Driver"
            all_individual_loads.loc[all_individual_loads["Loader 1"] == driver, "Role"] = "Loader 1"
            all_individual_loads.loc[all_individual_loads["Loader 2"] == driver, "Role"] = "Loader 2"
            
            # Select and rename columns for display
            display_df = all_individual_loads[["Day", "Role", "Load Type", "Actual Duration", "Target Minutes", "vs_target", "vs_avg"]].copy()
            display_df.columns = ["Date", "Role", "Load Type", "Actual Duration (min)", "Target (min)", "vs Target (min)", "vs Avg (min)"]
            
            # Sort by date, most recent first
            display_df = display_df.sort_values("Date", ascending=False)
            
            # Display the data
            st.dataframe(display_df,width=900, height=500)
        else:
            st.write("No loads found for this individual in the selected date range.")

##################################
# Team Performance Report Section
##################################
# Team Performance Report Section - Fixed
# Fixed Rankings Dashboard Section - Aggregation Function
elif report_choice == "Team Performance":
    st.header("Team Performance Report")
    
    # Create date_filtered_df in Team Performance section
    df = df_clean.copy()
    date_filtered_df = df[(df["Day Archive"] >= pd.to_datetime(start_date)) & (df["Day Archive"] <= pd.to_datetime(end_date))]
    
    # Use all load types from the data instead of just "Import 40"
    load_type_filter = st.selectbox("Select Load Type:", sorted(df_clean["Load Type"].dropna().unique()))
    
    min_loads = st.number_input("Minimum Loads Required:", min_value=0, max_value=50, value=10)
   
    team_report_html = generate_team_report(load_type_filter=load_type_filter, min_loads=min_loads, valid_drivers=valid_drivers)
   
    full_html = f"""
    <html>
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        {css}
      </head>
      <body>
        {team_report_html}
      </body>
    </html>
    """
    components.html(full_html, height=3200, scrolling=False)

##################################
# Performance Rankings Section
##################################
elif report_choice == "Performance Rankings":
    st.header("Performance Rankings by Load Type")
    st.markdown("Showing all individuals ranked from fastest to slowest for each load type")

    # Create date_filtered_df for this section
    df = df_clean.copy()
    date_filtered_df = df[(df["Day Archive"] >= pd.to_datetime(start_date)) & (df["Day Archive"] <= pd.to_datetime(end_date))]
    
    # Add min loads parameter
    min_loads_ranking = st.number_input("Minimum Loads Required for Ranking:", min_value=1, max_value=100, value=3)
    
    # Get unique load types
    unique_load_types = sorted(date_filtered_df["Load Type"].dropna().unique())

    for load_type in unique_load_types:
        with st.expander(f"Rankings for {load_type}", expanded=True):
            # Filter for this load type
            load_df = date_filtered_df[date_filtered_df["Load Type"] == load_type]
            
            # Process all individuals whether they were Driver, Loader1, or Loader2
            melted = load_df.melt(id_vars=[col for col in load_df.columns if col not in ["Driver", "Loader 1", "Loader 2"]],
                                  value_vars=["Driver", "Loader 1", "Loader 2"],
                                  value_name="Person")
            melted["Person"] = melted["Person"].astype(str).str.strip()
            melted = melted[melted["Person"].isin(valid_drivers)]
            
            # Calculate target difference before grouping
            melted["diff_vs_target"] = melted["Actual Duration"] - melted["Target Minutes"]
            
            # Group by person with corrected aggregation
            person_stats = melted.groupby("Person").agg(
                avg_duration=("Actual Duration", "mean"),
                vs_target=("diff_vs_target", "mean"),
                load_count=("Actual Duration", "count")
            ).reset_index()
            
            # Only include persons with minimum loads - use the parameter
            person_stats = person_stats[person_stats["load_count"] >= min_loads_ranking]
            
            # Sort by average duration (fastest first)
            person_stats = person_stats.sort_values("avg_duration")
            
            if not person_stats.empty:
                # Display as a dataframe
                person_stats["avg_duration"] = person_stats["avg_duration"].round(1).astype(str) + " min"
                person_stats["vs_target"] = person_stats["vs_target"].round(1).astype(str) + " min"
                person_stats.columns = ["Person", "Average Duration", "vs Target", "Load Count"]
                st.dataframe(person_stats)
            else:
                st.write(f"No individuals have at least {min_loads_ranking} loads for this load type.")

##################################
# Team Averages vs Goals Section
##################################
elif report_choice == "Team Averages vs Goals":
    st.header("Team Averages vs Goals Report")
    st.markdown(
        """
        This report provides an overview of team performance for each load type, comparing average completion times 
        to target goals. This information can be posted for the team to see collective performance without singling out individuals.
        """
    )

    # Filter data by date range
    df = df_clean.copy()
    date_filtered_df = df[(df["Day Archive"] >= pd.to_datetime(start_date)) & (df["Day Archive"] <= pd.to_datetime(end_date))]

    # Group by load type and calculate stats
    team_avg_by_load = date_filtered_df.groupby("Load Type").agg(
        avg_duration=("Actual Duration", "mean"),
        min_duration=("Actual Duration", "min"),
        max_duration=("Actual Duration", "max"),
        median_duration=("Actual Duration", "median"),
        load_count=("Actual Duration", "count")
    ).reset_index()

    # Merge with target goals
    team_avg_by_load = team_avg_by_load.merge(
        date_filtered_df[["Load Type", "Target Minutes"]].dropna().drop_duplicates(),
        on="Load Type", how="left"
    )

    # Calculate differences and percentages
    team_avg_by_load["diff_vs_target"] = team_avg_by_load["avg_duration"] - team_avg_by_load["Target Minutes"]
    team_avg_by_load["percent_of_target"] = (team_avg_by_load["avg_duration"] / team_avg_by_load["Target Minutes"]) * 100
    
    # Sort by load count (most common first)
    team_avg_by_load = team_avg_by_load.sort_values("load_count", ascending=False)

    # Create a summary card for posting
    st.subheader("Pando Shipping Team Performance Summary")
    st.markdown(f"**Date Range: {start_date_str} to {end_date_str}**")
    
    # Use Streamlit's native components instead of complex HTML
    st.markdown("### Team Performance Summary", unsafe_allow_html=True)
    
    # Create a styled DataFrame to display instead of custom HTML
    performance_df = team_avg_by_load.copy()
    performance_df['Status'] = performance_df.apply(
        lambda row: "✅ Under Goal" if row["diff_vs_target"] < 0 else 
                  "⚠️ Near Goal" if row["diff_vs_target"] < 0.15 * row["Target Minutes"] else 
                  "❌ Over Goal", 
        axis=1
    )
    
    # Format the display columns for better readability
    display_df = performance_df.copy()
    display_df["Completed"] = display_df["load_count"].astype(int)
    display_df["Avg Time"] = display_df["avg_duration"].map("{:.1f} min".format)
    display_df["Goal"] = display_df["Target Minutes"].map("{:.1f} min".format)
    display_df["vs Goal"] = display_df["diff_vs_target"].map("{:+.1f} min".format)
    
    # Select and reorder columns for display
    display_df = display_df[["Load Type", "Completed", "Avg Time", "Goal", "vs Goal", "Status"]]
    
    # Display the DataFrame with styling
    st.dataframe(
        display_df,
        column_config={
            "Load Type": st.column_config.TextColumn("Load Type", width="medium"),
            "Completed": st.column_config.NumberColumn("Completed", width="small"),
            "Avg Time": st.column_config.TextColumn("Avg Time", width="small"),
            "Goal": st.column_config.TextColumn("Goal", width="small"),
            "vs Goal": st.column_config.TextColumn("vs Goal", width="small"),
            "Status": st.column_config.TextColumn("Status", width="medium")
        },
        use_container_width=True
    )
    
    # Add a simple legend
    st.markdown("""
    **Legend:**
    - ✅ **Under Goal**: Performance is better than target (negative values)
    - ⚠️ **Near Goal**: Performance is slightly above target (within 15%)
    - ❌ **Over Goal**: Performance is significantly above target (more than 15%)
    """)
    
    # Add a download button for the summary
    csv_data = team_avg_by_load.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Summary as CSV",
        data=csv_data,
        file_name=f"pando_team_performance_{start_date_str}_to_{end_date_str}.csv",
        mime='text/csv',
    )

    # Color coding for performance in detailed table
    def perf_color(diff, target):
        if pd.isna(diff) or pd.isna(target):
            return ""
        if diff < 0:
            return "background-color: #d4edda;"  # green
        elif diff < 0.15 * target:
            return "background-color: #fff3cd;"  # yellow
        else:
            return "background-color: #f8d7da;"  # red

    styled_table = team_avg_by_load.style.apply(
        lambda row: [perf_color(row["diff_vs_target"], row["Target Minutes"])] * len(row), axis=1
    ).format({
        "avg_duration": "{:.1f} min",
        "min_duration": "{:.1f} min",
        "max_duration": "{:.1f} min",
        "median_duration": "{:.1f} min",
        "Target Minutes": "{:.1f} min",
        "diff_vs_target": "{:+.1f} min",
        "percent_of_target": "{:.0f}%"
    })

    st.markdown("<h3>Detailed Performance Table</h3>", unsafe_allow_html=True)
    st.dataframe(styled_table, use_container_width=True)

    # Bar chart: Average Duration vs Target
    st.markdown("<h3>Average Duration vs Target by Load Type</h3>", unsafe_allow_html=True)
    fig, ax = plt.subplots(figsize=(10, 6), dpi=80)
    x = team_avg_by_load["Load Type"]
    avg = team_avg_by_load["avg_duration"]
    target = team_avg_by_load["Target Minutes"]
    diff = team_avg_by_load["diff_vs_target"]
    
    # Create an improved bar chart showing actual time, target, and difference
    bar_width = 0.35
    x_pos = np.arange(len(x))
    
    bars1 = ax.bar(x_pos - bar_width/2, avg, width=bar_width, color="#086ccc", label="Team Avg")
    bars2 = ax.bar(x_pos + bar_width/2, target, width=bar_width, color="#4CAF50", label="Target")
    
    # Add data labels showing the difference
    for i, (bar1, d) in enumerate(zip(bars1, diff)):
        label_color = 'green' if d < 0 else 'red'
        ax.text(i - bar_width/2, bar1.get_height() + 1, 
                f"{d:+.1f}", 
                ha='center', va='bottom', 
                color=label_color, fontweight='bold')
    
    ax.set_xticks(x_pos)
    ax.set_xticklabels(x, rotation=45, ha="right")
    ax.set_ylabel("Minutes")
    ax.set_title("Team Average Duration vs Target by Load Type")
    ax.legend()
    plt.tight_layout()
    st.pyplot(fig)
