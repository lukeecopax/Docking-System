## optimized for printing

# app.py
import streamlit as st
import warnings
import requests
import pandas as pd
import numpy as np
from datetime import datetime
import streamlit.components.v1 as components

st.set_page_config(layout="wide")  # Enables wide mode by default

# Suppress specific warnings
warnings.filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")

# -----------------------------------------------------------
# Utility Function: Download and Load Excel Data (Cached)
# -----------------------------------------------------------
def load_excel_data():
    direct_download_url = "https://netorg245148-my.sharepoint.com/:x:/g/personal/lhuang_ecopaxinc_com/EdUeZd99AJJFn9GIOYn6s74BxnH194ZxUJq5DfehBlI9EQ?e=TlZyOt&download=1"
    output_file = "Docking_System_-_Pando_(V5).xlsx"
    response = requests.get(direct_download_url)
    if response.status_code == 200:
        with open(output_file, "wb") as f:
            f.write(response.content)
    else:
        st.error(f"Error downloading file: {response.status_code}")
        return None, None, None
    # Load sheets
    df = pd.read_excel(output_file, sheet_name="Cleaned Archive", engine="openpyxl")
    load_type_df = pd.read_excel(output_file, sheet_name="Load Type", engine="openpyxl")
    driver_df = pd.read_excel(output_file, sheet_name="Driver", engine="openpyxl")
    return df, load_type_df, driver_df

# Load data once
df_raw, load_type_df, driver_df = load_excel_data()
if df_raw is None:
    st.stop()

# -----------------------------------------------------------
# Process Driver Sheet: Compute both Available and All lists
# -----------------------------------------------------------
if driver_df is not None:
    driver_df["Available"] = driver_df["Available"].astype(str).str.strip().str.upper()
    driver_df["Driver Name"] = driver_df["Driver Name"].astype(str).str.strip()
    available_driver_options = driver_df[driver_df["Available"] == "YES"]["Driver Name"].dropna().unique().tolist()
    available_driver_options = [d for d in available_driver_options if d.lower() != "nan" and d != ""]
    available_driver_options = sorted(available_driver_options)
    
    all_driver_options = driver_df["Driver Name"].dropna().unique().tolist()
    all_driver_options = [d for d in all_driver_options if d.lower() != "nan" and d != ""]
    all_driver_options = sorted(all_driver_options)
else:
    available_driver_options = []
    all_driver_options = []

# -----------------------------------------------------------
# Sidebar: Report Type and Individuals Scope Selection
# -----------------------------------------------------------
st.sidebar.title("Docking System Performance Reports")
report_choice = st.sidebar.selectbox("Select Report Type:", ["Individual Performance", "Team Performance"])

st.title("Docking System App v0.5")

# New selection for individuals scope
individual_scope = st.sidebar.radio("Select Individuals Scope:", ["Only Available Individuals", "All Individuals"])
if individual_scope == "Only Available Individuals":
    valid_drivers = available_driver_options
else:
    valid_drivers = all_driver_options

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
# Function: Process Individual/Team Performance
# (Now accepts valid_drivers as an extra parameter)
# -----------------------------------------------------------
def process_performance(individual_df, overall_df, module_title, driver, start_date_str, end_date_str, valid_drivers):
    # Exclude rows with load type "Assignment" from both datasets
    individual_df = individual_df[individual_df["Load Type"].str.lower() != "assignment"]
    overall_df = overall_df[overall_df["Load Type"].str.lower() != "assignment"]

    # In Solo mode, limit overall_df to rows where the Driver is valid.
    if module_title == "Solo":
        overall_df = overall_df[overall_df["Driver"].isin(valid_drivers)]
    
    if individual_df.empty:
        return f"<p>No records found for <strong>{module_title}</strong> performance for {driver} between <strong>{start_date_str}</strong> and <strong>{end_date_str}</strong>.</p>"

    # 1. Group individual performance by Load Type.
    indiv_grouped = individual_df.groupby("Load Type").agg({
        "Actual Duration": ["mean", "count"]
    }).reset_index()
    indiv_grouped.columns = ["Load Type", "Average Duration Numeric", "# Loads"]

    # 2. Compute overall average performance for each Load Type.
    overall_avg = overall_df.groupby("Load Type")["Actual Duration"].mean().to_dict()
    indiv_grouped["vs Average (numeric)"] = indiv_grouped.apply(
        lambda row: row["Average Duration Numeric"] - overall_avg.get(row["Load Type"], float('nan')),
        axis=1
    )
    indiv_grouped["Overall Average"] = indiv_grouped["Load Type"].map(overall_avg)

    # 3. Merge with Load Type info.
    merged_df = pd.merge(indiv_grouped, load_type_df, on="Load Type", how="left")

    # 4. Calculate vs Target (numeric).
    merged_df["vs Target (numeric)"] = merged_df["Average Duration Numeric"] - merged_df["Target minutes"]

    # Define a common status function for color coding.
    def get_status(avg, target):
        if avg < target:
            return "good"
        else:
            diff = avg - target
            if diff < 0.15 * target:
                return "avg"
            else:
                return "bad"

    # Display function for vs Target.
    def vs_target_display(row):
        diff = row["vs Target (numeric)"]
        target_val = row["Target minutes"]
        avg_val = row["Average Duration Numeric"]
        status_class = get_status(avg_val, target_val)
        tooltip_text = {
            "good": "Good: Performance is below target",
            "avg": "Average: Performance is slightly above target",
            "bad": "Bad: Performance is significantly above target"
        }.get(status_class, "")
        return f"{diff:.1f} min <span class='status-dot {status_class}' title='{tooltip_text}'></span>"

    # Display function for vs Average.
    def vs_average_display(row):
        diff = row["vs Average (numeric)"]
        overall_avg_val = row["Overall Average"]
        indiv_avg_val = row["Average Duration Numeric"]
        status_class = get_status(indiv_avg_val, overall_avg_val)
        tooltip_text = {
            "good": "Good: Performance is below overall average",
            "avg": "Average: Performance is slightly above overall average",
            "bad": "Bad: Performance is significantly above overall average"
        }.get(status_class, "")
        return f"{diff:.1f} min <span class='status-dot {status_class}' title='{tooltip_text}'></span>"

    merged_df["vs Target"] = merged_df.apply(vs_target_display, axis=1)
    merged_df["vs Average"] = merged_df.apply(vs_average_display, axis=1)

    # Format Average Duration.
    merged_df["Average Duration"] = merged_df["Average Duration Numeric"].apply(lambda x: f"{x:.1f} min")

    # --- New Metric: vs Top ---
    if module_title == "Team":
        participants_df = overall_df.copy()
        participants_df['Participants'] = participants_df.apply(
            lambda row: [str(row['Driver']).strip(), str(row['Loader 1']).strip(), str(row['Loader 2']).strip()],
            axis=1
        )
        participants_df = participants_df.explode('Participants')
        participants_df = participants_df[
            (participants_df['Participants'] != '') &
            (participants_df['Participants'].str.lower() != 'nan') &
            (participants_df['Participants'] != 'None')
        ]
        participants_df = participants_df[participants_df['Participants'].isin(valid_drivers)]
        eligible_df = participants_df.groupby(["Participants", "Load Type"]).agg(
            avg_duration=pd.NamedAgg(column="Actual Duration", aggfunc="mean"),
            loads_count=pd.NamedAgg(column="Actual Duration", aggfunc="count")
        ).reset_index().rename(columns={'Participants': 'Driver'})
    else:
        eligible_df = overall_df[overall_df["Driver"].isin(valid_drivers)].groupby(["Driver", "Load Type"]).agg(
            avg_duration=pd.NamedAgg(column="Actual Duration", aggfunc="mean"),
            loads_count=pd.NamedAgg(column="Actual Duration", aggfunc="count")
        ).reset_index()

    eligible_df = eligible_df[eligible_df["loads_count"] > 5]
    eligible_top = eligible_df.groupby("Load Type")["avg_duration"].min().to_dict()
    eligible_driver_avg = {(row["Load Type"], row["Driver"]): row["avg_duration"] for _, row in eligible_df.iterrows()}

    def vs_top_display(row):
        driver_key = (row["Load Type"], driver)
        if row["# Loads"] <= 5 or driver_key not in eligible_driver_avg:
            return "N/A"
        driver_avg = eligible_driver_avg.get(driver_key)
        top_value = eligible_top.get(row["Load Type"], None)
        if top_value is None or pd.isna(driver_avg) or driver_avg == 0:
            return "N/A"
        ratio = (top_value / driver_avg) * 100
        return f"{ratio:.0f}% of #1"

    merged_df["vs Top"] = merged_df.apply(vs_top_display, axis=1)

    # --- New Metric: Ranking ---
    if module_title == "Team":
        ranking_df = participants_df.groupby(["Participants", "Load Type"]).agg(
            avg_duration=pd.NamedAgg(column="Actual Duration", aggfunc="mean"),
            loads_count=pd.NamedAgg(column="Actual Duration", aggfunc="count")
        ).reset_index().rename(columns={'Participants': 'Driver'})
    else:
        ranking_df = overall_df[overall_df["Driver"].isin(valid_drivers)].groupby(["Driver", "Load Type"]).agg(
            avg_duration=pd.NamedAgg(column="Actual Duration", aggfunc="mean"),
            loads_count=pd.NamedAgg(column="Actual Duration", aggfunc="count")
        ).reset_index()
    ranking_df = ranking_df[ranking_df["loads_count"] > 5]
    ranking_df['rank'] = ranking_df.groupby("Load Type")["avg_duration"].rank(method="min", ascending=True)
    ranking_dict = {(row["Load Type"], row["Driver"]): int(row["rank"]) for _, row in ranking_df.iterrows()}

    def ranking_display(row):
        rank = ranking_dict.get((row["Load Type"], driver), None)
        return f"Rank: {rank}" if rank is not None else "N/A"

    merged_df["Ranking"] = merged_df.apply(ranking_display, axis=1)

    # --- Overall Metrics ---
    total_loads = merged_df["# Loads"].sum()
    if total_loads > 0:
        weighted_avg_duration = (merged_df["Average Duration Numeric"] * merged_df["# Loads"]).sum() / total_loads
        weighted_target = (merged_df["Target minutes"] * merged_df["# Loads"]).sum() / total_loads
        overall_diff = weighted_avg_duration - weighted_target
    else:
        weighted_avg_duration = weighted_target = overall_diff = 0

    total_status = get_status(weighted_avg_duration, weighted_target)
    tooltip_total = {
        "good": "Good: Performance is below target",
        "avg": "Average: Performance is slightly above target",
        "bad": "Bad: Performance is significantly above target"
    }.get(total_status, "")
    total_vs_target = f"{overall_diff:+.1f} min <span class='status-dot {total_status}' title='{tooltip_total}'></span>"

    if not overall_df.empty:
       overall_module_avg = overall_df["Actual Duration"].mean()
    else:
       overall_module_avg = 0

    overall_diff_vs_avg = weighted_avg_duration - overall_module_avg
    total_status_vs_avg = get_status(weighted_avg_duration, overall_module_avg)
    tooltip_total_vs_avg = {
       "good": "Good: Performance is below overall average",
       "avg": "Average: Performance is slightly above overall average",
       "bad": "Bad: Performance is significantly above overall average"
    }.get(total_status_vs_avg, "")
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

    # Ensure "Assignment" rows are not displayed (they were already filtered earlier)
    merged_df = merged_df[merged_df["Load Type"].str.lower() != "assignment"]

    merged_df = merged_df[["Load Type", "Average Duration", "# Loads", "vs Target", "vs Average", "vs Top", "Ranking"]]

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
# (Accepts valid_drivers as extra parameter)
# -----------------------------------------------------------
def generate_team_report(load_type_filter="Import 40", min_loads=10, valid_drivers=None):
    df_team = df_raw.copy()
    df_team = df_team[df_team["Driver"].notna() & (df_team["Driver"].str.strip() != "")]
    df_team = df_team[df_team["Loader 1"].notna() | df_team["Loader 2"].notna()]
    df_team["Day Archive"] = pd.to_datetime(df_team["Day Archive"], errors="coerce")
    df_team = df_team[df_team["Day Archive"].dt.year == 2024]
    df_team["Actual Duration"] = pd.to_numeric(df_team["Actual Duration"], errors="coerce")
    df_team = pd.merge(df_team, load_type_df[["Load Type", "Target minutes"]], on="Load Type", how="left")
    df_team["Diff"] = df_team["Actual Duration"] - df_team["Target minutes"]
    
    # Exclude rows with Load Type "Assignment"
    df_team = df_team[df_team["Load Type"].str.lower() != "assignment"]
    
    # Only consider rows where the primary Driver is in valid_drivers.
    df_team = df_team[df_team["Driver"].isin(valid_drivers)]
    
    # Filter to only include the specified load type.
    df_team = df_team[df_team["Load Type"] == load_type_filter]

    team_group = df_team.groupby(["Driver", "Loader 1", "Loader 2"]).agg(
        avg_duration=("Actual Duration", "mean"),
        total_diff=("Diff", "sum"),
        avg_diff=("Diff", "mean"),
        load_count=("Diff", "count")
    ).reset_index()

    teams_valid = team_group[team_group["load_count"] >= min_loads]

    top_teams = teams_valid.nsmallest(10, "avg_diff") if not teams_valid.empty else None

    df_team["Team Members"] = df_team.apply(lambda row: [row["Driver"], row["Loader 1"], row["Loader 2"]], axis=1)
    df_exploded = df_team.explode("Team Members")
    df_exploded = df_exploded[df_exploded["Team Members"].isin(valid_drivers)]
    individual_stats = df_exploded.groupby("Team Members").agg(
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

    # Wrap the Top 10 Best Teams section in a div with class "teams-tables"
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

    # Wrap the Bottom 10 Teams section in a div with class "teams-tables"
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
            <td>{row["Team Members"]}</td>
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



    
    df = df_raw.copy()
    df["Actual Duration"] = pd.to_numeric(df["Actual Duration"], errors="coerce")
    df["Day Archive"] = pd.to_datetime(df["Day Archive"], errors="coerce")
    df["Loader 1"] = df["Loader 1"].astype(str)
    df["Loader 2"] = df["Loader 2"].astype(str)
    date_filtered_df = df[(df["Day Archive"] >= pd.to_datetime(start_date)) & (df["Day Archive"] <= pd.to_datetime(end_date))]
    # Exclude rows with Load Type "Assignment"
    date_filtered_df = date_filtered_df[date_filtered_df["Load Type"].str.lower() != "assignment"]
    # This is for workload summary, thus include all load types, including assignment
    workload_df = df[(df["Day Archive"] >= pd.to_datetime(start_date)) & 
                 (df["Day Archive"] <= pd.to_datetime(end_date))].copy()
    
    # -----------------------------------------------------------
    # New Section: Workload Summary Card for the Selected Individual
    # -----------------------------------------------------------
    def get_involved(row):
        names = {str(row["Driver"]).strip(), str(row["Loader 1"]).strip(), str(row["Loader 2"]).strip()}
        return [name for name in names if name and name.lower() != "nan" and name != "None"]
    
    workload_df["Involved"] = workload_df.apply(get_involved, axis=1)
    involved_exploded = workload_df.explode("Involved")
    involved_exploded = involved_exploded[involved_exploded["Involved"].isin(valid_drivers)]
    individual_metrics = involved_exploded.groupby("Involved").agg(
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
    import matplotlib.pyplot as plt
    from io import BytesIO
    
    # Compute individual load counts per Load Type for the selected driver
    individual_loads_df = date_filtered_df[date_filtered_df.apply(
        lambda row: driver in {str(row["Driver"]).strip(), str(row["Loader 1"]).strip(), str(row["Loader 2"]).strip()},
        axis=1
    )]
    individual_counts = individual_loads_df.groupby("Load Type").size()
    
    # Compute team average load counts per Load Type
    # First, define a helper function to get all involved individuals in a load record
    def get_involved(row):
        names = {str(row["Driver"]).strip(), str(row["Loader 1"]).strip(), str(row["Loader 2"]).strip()}
        return [name for name in names if name and name.lower() != "nan" and name != "None"]
    
    # Explode the records so that each individual is counted separately
    exploded_df = date_filtered_df.copy()
    exploded_df["Involved"] = exploded_df.apply(get_involved, axis=1)
    exploded_df = exploded_df.explode("Involved")
    exploded_df = exploded_df[exploded_df["Involved"].isin(valid_drivers)]
    
    # Compute individual load counts per Load Type for each involved team member
    individual_loads = exploded_df.groupby(["Load Type", "Involved"]).size().reset_index(name="load_count")
    # Then compute the team average load count per load type (averaging over individuals)
    team_avg = individual_loads.groupby("Load Type")["load_count"].mean()
    
    # Combine the individual's counts and the team average into a single DataFrame
    counts_df = pd.concat([individual_counts, team_avg], axis=1, keys=['Individual', 'Team_Avg']).fillna(0).reset_index()
    
    # Sort by total loads handled (Individual count) descending so the load type with the most loads appears upper
    counts_df = counts_df.sort_values(by='Individual', ascending=False)
    
    # Create a grouped horizontal bar chart comparing the individual's counts vs the team average for each Load Type
    fig, ax = plt.subplots(figsize=(10, 6), dpi=80)
    bar_width = 0.35
    y_pos = np.arange(len(counts_df))
    # Use the individual's name for the first bar's label instead of "Individual"
    ax.barh(y_pos - bar_width/2, counts_df['Individual'], height=bar_width, color='#086ccc', label=driver)
    ax.barh(y_pos + bar_width/2, counts_df['Team_Avg'], height=bar_width, color='gray', label='Team Average')
    ax.set_xlabel('Number of Loads')
    # Include the individual's name in the chart title
    ax.set_title(f'Total Loads Handled per Category vs Team Average for {driver}')
    ax.set_yticks(y_pos)
    ax.set_yticklabels(counts_df['Load Type'])
    ax.invert_yaxis()  # Ensures that the load type with the highest count is at the top
    ax.legend(loc='best')  # Adds a legend with the updated labels
    plt.tight_layout()
    
    # Display the chart
    buf = BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    st.image(buf, width=1000)


    
    # -----------------------------------------------------------
    # Continue with existing Individual Performance filtering
    # -----------------------------------------------------------
    overall_solo_df = date_filtered_df[
        (date_filtered_df["Loader 1"].str.strip().isin(["", "nan"])) &
        (date_filtered_df["Loader 2"].str.strip().isin(["", "nan"]))
    ]
    solo_df = overall_solo_df[overall_solo_df["Driver"] == driver]
    
    overall_team_df = date_filtered_df[
        ~((date_filtered_df["Loader 1"].str.strip().isin(["", "nan"])) &
          (date_filtered_df["Loader 2"].str.strip().isin(["", "nan"])))
    ]
    
    # Apply the new Team Role Filter only to the Team Performance section
    if team_role_filter == "Driver Only":
        team_df = overall_team_df[overall_team_df["Driver"] == driver]
    elif team_role_filter == "Loader Only":
        team_df = overall_team_df[
            (overall_team_df["Loader 1"].str.strip() == driver) |
            (overall_team_df["Loader 2"].str.strip() == driver)
        ]
    else:  # "Either" – default view where driver appears in any role
        team_df = overall_team_df[
            (overall_team_df["Driver"] == driver) |
            (overall_team_df["Loader 1"].str.strip() == driver) |
            (overall_team_df["Loader 2"].str.strip() == driver)
        ]
    
    solo_html = process_performance(solo_df, overall_solo_df, "Solo", driver, start_date_str, end_date_str, valid_drivers)
    team_html = process_performance(team_df, overall_team_df, "Team", driver, start_date_str, end_date_str, valid_drivers)
    
    final_html = css + "<div class='modules-container'>" + solo_html + "<br><br>" + team_html + "</div>"
    st.markdown(final_html, unsafe_allow_html=True)
    
    # -----------------------------------------------------------
    # New Section: Collaborator Analysis for the Selected Individual
    # -----------------------------------------------------------
    def collaborator_analysis(df, selected_driver, valid_drivers, min_collabs=5):
        # Filter to team records: where at least one loader is present
        team_records = df[~(
            (df["Loader 1"].str.strip().isin(["", "nan"])) &
            (df["Loader 2"].str.strip().isin(["", "nan"]))
        )].copy()
    
        # Merge to obtain the target minutes for each load type so we can compute Diff if not already present
        team_records = pd.merge(team_records, load_type_df[["Load Type", "Target minutes"]],
                                on="Load Type", how="left")
        team_records["Diff"] = team_records["Actual Duration"] - team_records["Target minutes"]
    
        # Filter to records that include the selected driver in any role
        team_records = team_records[team_records.apply(
            lambda row: selected_driver in {str(row["Driver"]).strip(), str(row["Loader 1"]).strip(), str(row["Loader 2"]).strip()},
            axis=1
        )]
    
        # Initialize a dictionary to hold Diff values for each collaborator
        collab_data = {}
        for _, row in team_records.iterrows():
            # Get unique team members from the row
            members = {str(row["Driver"]).strip(), str(row["Loader 1"]).strip(), str(row["Loader 2"]).strip()}
            # Remove invalid values
            members = {m for m in members if m and m.lower() != "nan" and m != "None"}
            if selected_driver in members:
                for member in members:
                    if member != selected_driver and member in valid_drivers:
                        if member not in collab_data:
                            collab_data[member] = []
                        collab_data[member].append(row["Diff"])
    
        # Build a DataFrame for collaborators with at least min_collabs collaborations
        records = []
        for member, diffs in collab_data.items():
            if len(diffs) >= min_collabs:
                avg_diff = round(np.mean(diffs), 1)  # format to 1 decimal place
                count = len(diffs)
                records.append((member, count, avg_diff))
        collab_df = pd.DataFrame(records, columns=["Collaborator", "Collaborations", "Average Diff (min)"])
        if collab_df.empty:
            return None, None
        
        sorted_df = collab_df.sort_values("Average Diff (min)")
        n = len(sorted_df)
        if n >= 10:
            best_collab = sorted_df.head(5)
            worst_collab = sorted_df.sort_values("Average Diff (min)", ascending=False).head(5)
        else:
            best_count = int(np.ceil(n / 2))
            worst_count = n - best_count
            best_collab = sorted_df.head(best_count)
            worst_collab = sorted_df.tail(worst_count)
        return best_collab, worst_collab
    
    best_collab, worst_collab = collaborator_analysis(date_filtered_df, driver, valid_drivers, min_collabs=5)
    
    collab_html = "<div class='modules-container'>"
    collab_html += "<h2>Collaborator Analysis for {}</h2>".format(driver)
    collab_html += """
    <div class='explanation'>
        <p><strong>Average Diff (min):</strong> This value shows, on average, how many minutes the actual load durations differed from the target time when working with each collaborator.
        A lower (or negative) number means loads were generally completed faster than the target, while a higher number indicates loads took longer.</p>
        <br><em>Note:</em> Only collaborators who have been involved in at least 5 team records (min_collabs=5) are included in this analysis.</p>
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
# Team Performance Report Section (Solution 3 Final)
# -----------------------------------------------------------
elif report_choice == "Team Performance":
    st.header("Team Performance Report (2024)")
    load_type_filter = st.selectbox("Select Load Type:", ["Import 40"])
    min_loads = st.number_input("Minimum Loads Required:", min_value=0, max_value=50, value=10)
    
    team_report_html = generate_team_report(load_type_filter=load_type_filter, min_loads=min_loads, valid_drivers=valid_drivers)
    
    # Wrap the team report HTML in a full HTML document including the CSS exactly as in your original output
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
    # Render the full HTML document using components.html with a high fixed height (adjust if necessary)
    components.html(full_html, height=3200, scrolling=False)
