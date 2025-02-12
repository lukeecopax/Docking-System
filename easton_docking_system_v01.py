import streamlit as st
import warnings
import requests
import pandas as pd
import numpy as np
from datetime import datetime
import streamlit.components.v1 as components
import matplotlib.pyplot as plt
from io import BytesIO
 
st.set_page_config(layout="wide")  # Enables wide mode by default
 
# Suppress specific warnings
warnings.filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")
 
# -----------------------------------------------------------
# Parameters for performance thresholds and minimum counts
# -----------------------------------------------------------
THRESHOLD_PERCENT = 0.15
MIN_LOADS_RANK = 5
MIN_COLLABS = 5
 
# -----------------------------------------------------------
# Utility Function: Download and Load Excel Data (Cached)
# -----------------------------------------------------------
def load_excel_data():
    direct_download_url = "https://netorg245148-my.sharepoint.com/:x:/g/personal/lhuang_ecopaxinc_com/EWcwHdeMwsxFmh4s6blHiiABxhwd0T31H8GD5uY9M26cmg?e=CguuMW&download=1"
    output_file = "Docking_System_-_Easton_(V2).xlsx"
    response = requests.get(direct_download_url)
    if response.status_code == 200:
        with open(output_file, "wb") as f:
            f.write(response.content)
    else:
        st.error(f"Error downloading file: {response.status_code}")
        return None, None, None
    # Load sheets
    df = pd.read_excel(output_file, sheet_name="Archive", engine="openpyxl")
    load_type_df = pd.read_excel(output_file, sheet_name="Load Type", engine="openpyxl")
    driver_df = pd.read_excel(output_file, sheet_name="Driver", engine="openpyxl")
   
    # Strip extra whitespace from column names
    df.columns = df.columns.str.strip()
    if load_type_df is not None:
        load_type_df.columns = load_type_df.columns.str.strip()
        load_type_df.rename(columns={"Load type": "Load Type", "Require mins": "Target minutes"}, inplace=True)
    if driver_df is not None:
        driver_df.columns = driver_df.columns.str.strip()
   
    return df, load_type_df, driver_df
 
# Load data once
df_raw, load_type_df, driver_df = load_excel_data()
if df_raw is None:
    st.stop()
 
# -----------------------------------------------------------
# Preprocessing: Create a Clean DataFrame and Role Flags
# -----------------------------------------------------------
# Filter out "Assignment" rows once
df_clean = df_raw[df_raw["Load Type"].str.lower() != "assignment"].copy()
 
# Convert date columns
df_clean["Day Archive"] = pd.to_datetime(df_clean["Day Archive"], errors="coerce")
# Combine the "Day Archive" with the time values to create full datetime values
df_clean["Start Time"] = pd.to_datetime(
    df_clean["Day Archive"].astype(str) + " " + df_clean["Start Time"].astype(str),
    errors="coerce"
)
 
# Now "Actual End Time" should match since whitespace was stripped.
if "Actual End Time" not in df_clean.columns:
    if "End Time" in df_clean.columns:
        df_clean.rename(columns={"End Time": "Actual End Time"}, inplace=True)
    else:
        st.error("The 'Actual End Time' (or 'End Time') column was not found in the Archive sheet. Please check the file.")
        st.stop()
 
df_clean["Actual End Time"] = pd.to_datetime(
    df_clean["Day Archive"].astype(str) + " " + df_clean["Actual End Time"].astype(str),
    errors="coerce"
)
# Compute Actual Duration in minutes as (Actual End Time - Start Time)
df_clean["Actual Duration"] = (df_clean["Actual End Time"] - df_clean["Start Time"]).dt.total_seconds() / 60
 
# Standardize text fields for loaders
df_clean["Loader 1"] = df_clean["Loader 1"].astype(str).str.strip()
df_clean["Loader 2"] = df_clean["Loader 2"].astype(str).str.strip()
 
# Create role flag columns: is_solo and is_team
df_clean["is_solo"] = ((df_clean["Loader 1"].isin(["", "nan", "None"])) & (df_clean["Loader 2"].isin(["", "nan", "None"])))
df_clean["is_team"] = ~df_clean["is_solo"]
 
# -----------------------------------------------------------
# Process Driver Sheet: Compute both Available and All lists
# -----------------------------------------------------------
if driver_df is not None:
    driver_df["Available"] = driver_df["Available"].astype(str).str.strip().str.upper()
    driver_df["Driver Name"] = driver_df["Driver Name"].astype(str).str.strip()
   
    # Build the available drivers list while excluding names starting with "Waiting"
    available_driver_options = driver_df[driver_df["Available"] == "YES"]["Driver Name"].dropna().unique().tolist()
    available_driver_options = [d for d in available_driver_options if d.lower() != "nan" and d != ""]
    available_driver_options = [d for d in available_driver_options if not d.lower().startswith("waiting")]
    available_driver_options = sorted(available_driver_options)
   
    # Build the all drivers list while excluding names starting with "Waiting"
    all_driver_options = driver_df["Driver Name"].dropna().unique().tolist()
    all_driver_options = [d for d in all_driver_options if d.lower() != "nan" and d != ""]
    all_driver_options = [d for d in all_driver_options if not d.lower().startswith("waiting")]
    all_driver_options = sorted(all_driver_options)
else:
    available_driver_options = []
    all_driver_options = []
 
# -----------------------------------------------------------
# Sidebar: Report Type and Individuals Scope Selection
# -----------------------------------------------------------
st.sidebar.title("Easton Docking System Performance Reports v0.1")
report_choice = st.sidebar.selectbox("Select Report Type:", ["Individual Performance", "Team Performance"])
 
# Updated Report Title
st.title("Easton Docking System Performance Dashboard v0.1")
 
# New selection for individuals scope
individual_scope = st.sidebar.radio("Select Individuals Scope:", ["All Individuals","Only Available Individuals"])
if individual_scope == "All Individuals":
    valid_drivers = all_driver_options
else:
    valid_drivers = available_driver_options
 
# -----------------------------------------------------------
# Common CSS for HTML Reports
# -----------------------------------------------------------
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
 
/* New CSS for Team Performance tables only */
.teams-tables table {
    border-spacing: 10px !important;
}
.teams-tables th, .teams-tables td {
    padding: 8px 15px !important;
}
</style>
"""
 
# -----------------------------------------------------------
# Helper Function: Vectorized Status Calculation
# -----------------------------------------------------------
def compute_status_vectorized(avg_series, target_series, threshold=THRESHOLD_PERCENT):
    conditions = [avg_series < target_series,
                  (avg_series >= target_series) & ((avg_series - target_series) < threshold * target_series)]
    choices = ["good", "avg"]
    return np.select(conditions, choices, default="bad")
 
# -----------------------------------------------------------
# Function: Process Individual/Team Performance (Vectorized)
# -----------------------------------------------------------
def process_performance(individual_df, overall_df, module_title, driver, start_date_str, end_date_str, valid_drivers):
    # Group individual performance by Load Type.
    indiv_grouped = individual_df.groupby("Load Type").agg(
        Average_Duration_Numeric=("Actual Duration", "mean"),
        Loads_Count=("Actual Duration", "count")
    ).reset_index()
   
    # Compute overall average performance for each Load Type.
    overall_avg = overall_df.groupby("Load Type")["Actual Duration"].mean().reset_index().rename(columns={"Actual Duration": "Overall_Average"})
    indiv_grouped = indiv_grouped.merge(overall_avg, on="Load Type", how="left")
    indiv_grouped["vs_Average_Numeric"] = indiv_grouped["Average_Duration_Numeric"] - indiv_grouped["Overall_Average"]
   
    # Merge with Load Type info.
    merged_df = pd.merge(indiv_grouped, load_type_df, on="Load Type", how="left")
   
    # Calculate vs Target (numeric).
    merged_df["vs_Target_Numeric"] = merged_df["Average_Duration_Numeric"] - merged_df["Target minutes"]
   
    # Compute status for vs Target and vs Average using vectorized operations.
    merged_df["target_status"] = compute_status_vectorized(merged_df["Average_Duration_Numeric"], merged_df["Target minutes"])
    merged_df["average_status"] = compute_status_vectorized(merged_df["Average_Duration_Numeric"], merged_df["Overall_Average"])
   
    # Tooltips for statuses.
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
   
    # Create display columns using vectorized string concatenation.
    merged_df["vs Target"] = merged_df["vs_Target_Numeric"].map("{:.1f} min".format).fillna("N/A").astype(str) + \
        " <span class='status-dot " + merged_df["target_status"].astype(str) + \
        "' title='" + pd.Series(tooltip_target).astype(str) + "'></span>"
       
    merged_df["vs Average"] = merged_df["vs_Average_Numeric"].map("{:.1f} min".format).fillna("N/A").astype(str) + \
        " <span class='status-dot " + merged_df["average_status"].astype(str) + \
        "' title='" + pd.Series(tooltip_average).astype(str) + "'></span>"
   
    # Format Average Duration.
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
    eligible_df = eligible_df[eligible_df["loads_count"] > MIN_LOADS_RANK]
    top_df = eligible_df.groupby("Load Type")["avg_duration"].min().reset_index().rename(columns={"avg_duration": "top_avg"})
    driver_avg_df = eligible_df[eligible_df["Driver"] == driver][["Load Type", "avg_duration"]].rename(columns={"avg_duration": "driver_avg"})
    vs_top_df = pd.merge(driver_avg_df, top_df, on="Load Type", how="left")
    vs_top_df["vs_top_ratio"] = np.where(vs_top_df["driver_avg"] > 0, (vs_top_df["top_avg"] / vs_top_df["driver_avg"]) * 100, np.nan)
    merged_df = merged_df.merge(vs_top_df[["Load Type", "vs_top_ratio"]], on="Load Type", how="left")
    merged_df["vs Top"] = np.where(merged_df["Loads_Count"] > MIN_LOADS_RANK,
                                   merged_df["vs_top_ratio"].map(lambda x: f"{x:.0f}% of #1" if not pd.isna(x) else "N/A"),
                                   "N/A")
   
    # --- New Metric: Ranking ---
    if module_title == "Team":
        if 'Participant' not in overall_df.columns:
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
    ranking_df = ranking_df[ranking_df["loads_count"] > MIN_LOADS_RANK]
    ranking_df["Rank"] = ranking_df.groupby("Load Type")["avg_duration"].rank(method="min", ascending=True)
    driver_ranks = ranking_df[ranking_df["Driver"] == driver][["Load Type", "Rank"]]
    merged_df = merged_df.merge(driver_ranks, on="Load Type", how="left")
    merged_df["Ranking"] = merged_df["Rank"].apply(lambda x: f"Rank: {int(x)}" if not pd.isna(x) else "N/A")
    merged_df.drop(columns=["Rank"], inplace=True)
   
    # --- Overall Metrics ---
    total_loads = merged_df["Loads_Count"].sum()
    if total_loads > 0:
        weighted_avg_duration = (merged_df["Average_Duration_Numeric"] * merged_df["Loads_Count"]).sum() / total_loads
        weighted_target = (merged_df["Target minutes"] * merged_df["Loads_Count"]).sum() / total_loads
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
        "Target minutes": f"{weighted_target:.1f}",
        "vs Target": total_vs_target,
        "vs Average": total_vs_average,
        "vs Top": "",
        "Ranking": ""
    }
    summary_df = pd.DataFrame([summary])
    merged_df = pd.concat([merged_df, summary_df], ignore_index=True)
   
    # Reorder columns to match desired output
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
 
# -----------------------------------------------------------
# Function: Generate Team Performance Report (2024)
# -----------------------------------------------------------
def generate_team_report(load_type_filter="Import 40", min_loads=10, valid_drivers=None):
    df_team = df_clean.copy()
    df_team = df_team[df_team["Driver"].notna() & (df_team["Driver"].str.strip() != "")]
    # Use the precomputed role flag to keep team records
    df_team = df_team[~df_team["is_solo"]]
    # Debug: Check rows for the year 2024
    df_team_year = df_team[df_team["Day Archive"].dt.year == 2024]
    
    # Filter for 2024 records
    df_team = df_team_year.copy()
    
    # Filter to only consider the specified load type.
    df_team = df_team[df_team["Load Type"].str.strip().str.lower() == load_type_filter.lower()]
    
    # Merge load type info to get "Target minutes"
    df_team = pd.merge(df_team, load_type_df[["Load Type", "Target minutes"]], on="Load Type", how="left")
    
    # Compute "Diff" before grouping
    df_team["Diff"] = df_team["Actual Duration"] - df_team["Target minutes"]
    
    # ----------------- New Change: Sort loader assignments so order does not matter -----------------
    def sort_loaders(row):
        l1 = str(row["Loader 1"]).strip()
        l2 = str(row["Loader 2"]).strip()
        loaders = [x for x in [l1, l2] if x and x.lower() != "nan" and x != "None"]
        sorted_loaders = sorted(loaders)
        if len(sorted_loaders) == 0:
            return pd.Series(["", ""])
        elif len(sorted_loaders) == 1:
            return pd.Series([sorted_loaders[0], ""])
        else:
            return pd.Series(sorted_loaders)
    df_team[['Loader 1', 'Loader 2']] = df_team.apply(sort_loaders, axis=1)
    # -----------------------------------------------------------------------------------------------
    
    team_group = df_team.groupby(["Driver", "Loader 1", "Loader 2"]).agg(
        avg_duration=("Actual Duration", "mean"),
        total_diff=("Diff", "sum"),
        avg_diff=("Diff", "mean"),
        load_count=("Diff", "count")
    ).reset_index()
   
    teams_valid = team_group[team_group["load_count"] >= min_loads]
    top_teams = teams_valid.nsmallest(10, "avg_diff") if not teams_valid.empty else None
   
    # Collaborate individual stats using melt for vectorized approach.
    df_team["Diff"] = df_team["Actual Duration"] - df_team["Target minutes"]
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
    <h2>Team Performance Report (2024) for Load Type: {}</h2>
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
 
# -----------------------------------------------------------
# Individual Performance Report Section
# -----------------------------------------------------------
if report_choice == "Individual Performance":
    st.header("Individual Performance Report")
   
    if not valid_drivers:
        st.error("No individuals found in the Driver sheet.")
        st.stop()
   
    # Define the default driver you want
    default_driver = "Carlos Rodriguez Escobar"
   
    # Determine the index of the default driver in the valid_drivers list
    if default_driver in valid_drivers:
        default_index = valid_drivers.index(default_driver)
    else:
        default_index = 0  # fallback to the first driver if not found
   
    # Create the selectbox with the default index
    driver = st.selectbox("Name:", valid_drivers, index=default_index)
    start_date = st.date_input("Start Date:", datetime(2024, 1, 1))
    end_date = st.date_input("End Date:", datetime(2024, 12, 31))
    start_date_str = start_date.strftime("%Y-%m-%d")
    end_date_str = end_date.strftime("%Y-%m-%d")
   
    # NEW: Team Role Filter (applied only to Team Performance section)
    team_role_filter = st.selectbox("Team Role Filter (applied to Team Performance):", ["Either", "Driver Only", "Loader Only"])
   
    # Use df_clean for filtering instead of df_raw
    df = df_clean.copy()
    # Filter by date range
    date_filtered_df = df[(df["Day Archive"] >= pd.to_datetime(start_date)) & (df["Day Archive"] <= pd.to_datetime(end_date))]
   
    # Workload Summary uses all load types from df_clean (assignments already filtered out)
    workload_df = df[(df["Day Archive"] >= pd.to_datetime(start_date)) & (df["Day Archive"] <= pd.to_datetime(end_date))].copy()
   
    # -----------------------------------------------------------
    # New Section: Workload Summary Card for the Selected Individual
    # -----------------------------------------------------------
    # Compute involved individuals using melt for vectorization
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
            <strong>Total Loads</strong> represents the number of load events you were involved in (as a driver or loader), while <strong>Total Duration</strong> is the sum of the durations for those loads. The <strong>Load Count Comparison</strong> and <strong>Total Duration Comparison</strong> metrics indicate how your performance measures against the overall median and average values among your peers. Values above 100% mean your metric is higher than the benchmark, and values below 100% indicate a lower workload or duration.
        </p>
    </div>
    """
    st.markdown(card_html, unsafe_allow_html=True)
   
    # -----------------------------------------------------------
    # New Visualization: Total Loads Handled per Category vs Entire Team Average
    # -----------------------------------------------------------
    # Compute individual load counts per Load Type for the selected driver
    individual_loads_df = date_filtered_df[date_filtered_df.apply(
        lambda row: driver in {str(row["Driver"]).strip(), str(row["Loader 1"]).strip(), str(row["Loader 2"]).strip()},
        axis=1
    )]
    individual_counts = individual_loads_df.groupby("Load Type").size()
   
    # Compute team average load counts per Load Type using melt
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
   
    # -----------------------------------------------------------
    # Continue with existing Individual Performance filtering
    # -----------------------------------------------------------
    overall_solo_df = date_filtered_df[date_filtered_df["is_solo"]]
    solo_df = overall_solo_df[overall_solo_df["Driver"] == driver]
   
    overall_team_df = date_filtered_df[~date_filtered_df["is_solo"]]
   
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
   
    # -----------------------------------------------------------
    # New Section: Collaborator Analysis for the Selected Individual
    # -----------------------------------------------------------
    def collaborator_analysis(df, selected_driver, valid_drivers, min_collabs=MIN_COLLABS):
        # Filter to team records where at least one loader is present
        team_records = df[~(
            (df["Loader 1"].str.strip().isin(["", "nan"])) &
            (df["Loader 2"].str.strip().isin(["", "nan"]))
        )].copy()
 
        # Get records where the selected driver is involved
        team_records = team_records[
            (team_records["Driver"] == selected_driver) |
            (team_records["Loader 1"] == selected_driver) |
            (team_records["Loader 2"] == selected_driver)
        ]
 
        # Merge to get target minutes if not present
        if "Target minutes" not in team_records.columns:
            team_records = pd.merge(team_records, load_type_df[["Load Type", "Target minutes"]],
                                  on="Load Type", how="left")
       
        # Calculate Diff if not present
        if "Diff" not in team_records.columns:
            team_records["Diff"] = team_records["Actual Duration"] - team_records["Target minutes"]
 
        # Initialize a list to store collaborator data
        collab_records = []
       
        # Process each record to identify collaborators
        for _, row in team_records.iterrows():
            # Get all team members for this record
            team_members = {
                str(row["Driver"]).strip(),
                str(row["Loader 1"]).strip(),
                str(row["Loader 2"]).strip()
            }
            # Remove invalid values
            team_members = {m for m in team_members if m and m.lower() != "nan" and m != "None"}
           
            # If selected driver is in this team, process each collaborator
            if selected_driver in team_members:
                for member in team_members:
                    if member != selected_driver and member in valid_drivers:
                        # Only add if enough collaborations are met later
                        collab_records.append((member, row["Diff"]))
 
        # Create a DataFrame from the records and group by collaborator
        if not collab_records:
            return None, None
        collab_df = pd.DataFrame(collab_records, columns=["Collaborator", "Diff"])
        grouped = collab_df.groupby("Collaborator").agg(
            Collaborations=("Diff", "count"),
            Average_Diff=("Diff", "mean")
        ).reset_index()
 
        # Only include collaborators with at least min_collabs collaborations
        grouped = grouped[grouped["Collaborations"] >= min_collabs]
        total_collaborators = len(grouped)
 
        if total_collaborators >= 10:
            best_collab = grouped.nsmallest(5, "Average_Diff")
            worst_collab = grouped.nlargest(5, "Average_Diff")
        else:
            # Sort the grouped dataframe and split it nearly 50/50
            grouped_sorted = grouped.sort_values("Average_Diff")
            mid = int(np.ceil(total_collaborators / 2))
            best_collab = grouped_sorted.iloc[:mid]
            worst_collab = grouped_sorted.iloc[mid:]
 
        return best_collab, worst_collab
   
    best_collab, worst_collab = collaborator_analysis(date_filtered_df, driver, valid_drivers, min_collabs=MIN_COLLABS)
   
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
 
# -----------------------------------------------------------
# Team Performance Report Section
# -----------------------------------------------------------
elif report_choice == "Team Performance":
    st.header("Team Performance Report (2024)")
    # ----------------- New Change: Limit load type options -----------------
    load_type_filter = st.selectbox("Select Load Type:", ["Floor Load", "Pallet Load"])
    # ----------------- New Change: Default Minimum Loads Required set to 3 -----------------
    min_loads = st.number_input("Minimum Loads Required:", min_value=0, max_value=50, value=3)
   
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