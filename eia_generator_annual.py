import requests
import calendar
from io import BytesIO
import zipfile
import openpyxl
import streamlit as st
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt

STATE_ABBR_MAP = {
    'AK': 'Alaska', 'AL': 'Alabama', 'AR': 'Arkansas', 'AZ': 'Arizona', 'CA': 'California',
    'CO': 'Colorado', 'CT': 'Connecticut', 'DC': 'District of Columbia', 'DE': 'Delaware',
    'FL': 'Florida', 'GA': 'Georgia', 'HI': 'Hawaii', 'IA': 'Iowa', 'ID': 'Idaho', 'IL': 'Illinois',
    'IN': 'Indiana', 'KS': 'Kansas', 'KY': 'Kentucky', 'LA': 'Louisiana', 'MA': 'Massachusetts',
    'MD': 'Maryland', 'ME': 'Maine', 'MI': 'Michigan', 'MN': 'Minnesota', 'MO': 'Missouri',
    'MS': 'Mississippi', 'MT': 'Montana', 'NC': 'North Carolina', 'ND': 'North Dakota',
    'NE': 'Nebraska', 'NH': 'New Hampshire', 'NJ': 'New Jersey', 'NM': 'New Mexico', 'NV': 'Nevada',
    'NY': 'New York', 'OH': 'Ohio', 'OK': 'Oklahoma', 'OR': 'Oregon', 'PA': 'Pennsylvania',
    'RI': 'Rhode Island', 'SC': 'South Carolina', 'SD': 'South Dakota', 'TN': 'Tennessee',
    'TX': 'Texas', 'UT': 'Utah', 'VA': 'Virginia', 'VT': 'Vermont', 'WA': 'Washington',
    'WI': 'Wisconsin', 'WV': 'West Virginia', 'WY': 'Wyoming'
}

def get_file_url(base_url, year):
    """Constructs the URL for a given year."""
    file_url = f"{base_url}{year}.zip"
    return file_url

def file_exists(url):
    """Checks if the file at the given URL exists (ZIP or Excel)."""
    try:
        response = requests.head(url, allow_redirects=True)
        if response.status_code == 200:
            content_type = response.headers.get('Content-Type', '').lower()
            return 'zip' in content_type or 'excel' in content_type
        return False
    except requests.RequestException:
        return False

def get_latest_file_url(base_url):
    """Finds the latest file URL."""
    now = datetime.now()
    current_year = now.year-1

    
    while True:
        file_url = get_file_url(base_url, current_year)
        ER_status = False
        if file_exists(file_url):
            st.write(f"Latest Year: {current_year}")
            return file_url, current_year, ER_status
        else:
            st.write(f"Latest Year Early Release: {current_year}")
            file_url = f"{base_url}{current_year}ER.zip"
            ER_status = True
            return file_url, current_year, ER_status

def get_previous_file_url(base_url, year):
    """Constructs the URL for the previous year's file."""
    
    return f"{base_url}{year-1}.zip"

@st.cache_data
def download_excel_file(url, year, ER_status):
    """
    Downloads and extracts the specific Excel file from the ZIP archive.
    Returns tuple: (BytesIO of Excel file, Excel filename string)
    """
    with st.spinner("Downloading file..."):
        try:
            response = requests.get(url)
            response.raise_for_status()

            with zipfile.ZipFile(BytesIO(response.content)) as z:
                # Build expected Excel filename inside ZIP
                filename_expected = f"3_1_Generator_Y{year}"
                if ER_status:
                    filename_expected += "_Early_Release"
                filename_expected += ".xlsx"

                if filename_expected not in z.namelist():
                    st.error(f"Excel file '{filename_expected}' not found inside ZIP archive.")
                    return None, None

                excel_data = z.read(filename_expected)
                return BytesIO(excel_data), filename_expected

        except requests.exceptions.RequestException as err:
            st.error(f"Request Exception occurred: {err}")
            return None, None
        except zipfile.BadZipFile:
            st.error("Downloaded file is not a valid ZIP archive.")
            return None, None

def rename_columns(df, sheet_name=None):
    """Renames columns based on the sheet name."""
    
    # Common renames
    if df.columns[15] == 'Unnamed: 15':
        df.rename(columns={df.columns[15]: 'Nameplate Capacity (MW)'}, inplace=True)
    if df.columns[7] == 'Unnamed: 7':
        df.rename(columns={df.columns[7]: 'Technology'}, inplace=True)
    if df.columns[2] == 'Unnamed: 2':
        df.rename(columns={df.columns[2]: 'Plant ID'}, inplace=True)
    if df.columns[4] == 'Unnamed: 4':
        df.rename(columns={df.columns[4]: 'State'}, inplace=True)
        df['State'] = df['State'].map(STATE_ABBR_MAP).fillna(df['State'])

    # Conditional rename for Energy Source column
    if sheet_name == 'Proposed':
        energy_source_col_index = 28
    else:
        energy_source_col_index = 33

    if df.columns[energy_source_col_index] == f'Unnamed: {energy_source_col_index}':
        df.rename(columns={df.columns[energy_source_col_index]: 'Energy Source'}, inplace=True)


    energy_code_conversion = {
        "AB": "Other",
        "BFG": "Other",
        "BIT": "Coal",
        "BLQ": "Other",
        "DFO": "Oil",
        "GEO": "Geothermal",
        "JF": "Other",
        "KER": "Oil",
        "LFG": "Other",
        "LIG": "Coal",
        "MSW": "Other",
        "MWH": "Storage",
        "NG": "Natural Gas",
        "NUC": "Nuclear",
        "OBG": "Other RE",
        "OGB": "Other RE",
        "OBL": "Other RE",
        "OBS": "Other RE",
        "OG": "Other",
        "OTH": "Other",
        "PC": "Other",
        "PG": "Oil",
        "PUR": "Other",
        "RC": "Coal",
        "RFO": "Oil",
        "SGC":"Coal",
        "SUB": "Coal",
        "SUN": "Solar",
        "WAT": "Hydro",
        "WC": "Coal",
        "WDL": "Other RE",
        "WDS": "Other RE",
        "WH": "Other",
        "WND": "Wind",
        "WO": "Oil"
    }

    df['Energy Source'] = df['Energy Source'].map(energy_code_conversion).fillna('Unknown')

    return df 

def sum_nameplate_capacity(df, sheet_name=None):
    """Sums up the 'Nameplate Capacity (MW)' column from the given DataFrame, starting from row 2, converting non-numeric values to NaN."""
    df = rename_columns(df, sheet_name)  # Ensure columns are correctly named
    if 'Nameplate Capacity (MW)' in df.columns:
        df['Nameplate Capacity (MW)'] = pd.to_numeric(df['Nameplate Capacity (MW)'], errors='coerce')  # Convert to numeric, setting errors to NaN
        
        # Slice the DataFrame to start from row 2
        df_to_sum = df.iloc[1:]
        
        return df_to_sum['Nameplate Capacity (MW)'].sum()  # Sum up the column starting from row 2
    else:
        st.warning("'Nameplate Capacity (MW)' column not found.")
        return 0

def count_plants(df):
    """Counts the number of plants in the given DataFrame."""
    return df.shape[0]

def plot_comparison(latest_data, previous_data, capacity_type):
    """Plots a bar graph comparing capacities of the latest and previous months for a specific type,
       with energy source distribution stacked within the bars."""

    def aggregate_fuel_type(df):
        """Aggregates energy sources and their total capacities."""
        rename_dict = {
                'Coal': 'cCoal',         # Light brownish
                'Oil': 'dOil',          # Dark gray
                'Natural Gas': 'eNatural Gas',          # Almost black
                'Nuclear': 'sNuclear',      # Pink
                'Hydro': 'wHydro',        # Navy blue
                'Solar': "ySolar",        # Yellowish (towards yellow)
                'Wind': 'xWind',         # Light emerald green
                'Storage': 'zStorage',      # Light gray
                'Other': 'aOther',        # Medium gray
                'Other RE': 'bOther RE'      # Light tan
            }
        
        if "Energy Source" in df.columns and "Nameplate Capacity (MW)" in df.columns:
            df = df[df["Energy Source"] != "Unknown"]
            df['Energy Source'] = df['Energy Source'].map(rename_dict)
            df = df.groupby("Energy Source")["Nameplate Capacity (MW)"].sum()
            df = df / 1000  # Convert to GW

            return df
        else:
            st.warning("'Energy Source' or 'Nameplate Capacity (MW)' column not found.")
            return pd.Series()

    # Ensure that all values are DataFrames
    if not all(isinstance(df, pd.DataFrame) for df in latest_data.values()):
        st.error("Some values in latest_data are not DataFrames.")
        return
    if not all(isinstance(df, pd.DataFrame) for df in previous_data.values()):
        st.error("Some values in previous_data are not DataFrames.")
        return

    # Process data for each sheet separately
    for sheet_name in latest_data.keys():
        latest_df = latest_data[sheet_name]
        previous_df = previous_data[sheet_name]
        
        # Aggregate by energy source for each sheet
        latest_energy_sources = aggregate_fuel_type(latest_df)
        previous_energy_sources = aggregate_fuel_type(previous_df)

    # Create DataFrame for comparison with energy sources as rows and months as columns
    df_comparison = pd.DataFrame({
        'Previous Year': previous_energy_sources,
        'Latest Year': latest_energy_sources
    }).fillna(0).T

    # Define the color dictionary
    color_dict = {
        'cCoal': '#a0522d',         # Light brownish
        'dOil': '#2f4f4f',          # Dark gray
        'eNatural Gas': '#1f77b4',          # Almost black
        'sNuclear': '#ff69b4',      # Pink
        'wHydro': '#003366',        # Navy blue
        'ySolar': "#F5B800",        # Yellowish (towards yellow)
        'xWind': '#39FF14',         # Light emerald green
        'zStorage': '#d3d3d3',      # Light gray
        'aOther': '#808080',        # Medium gray
        'bOther RE': '#d2b48c'      # Light tan
    }
        
    # Reorder DataFrame based on custom order
    
    # Plot stacked bar chart
    fig, ax = plt.subplots(figsize=(12, 8))
    df_comparison.plot(kind='bar', stacked=True, ax=ax, color=[color_dict.get(x, '#d3d3d3') for x in df_comparison.columns])
    
    plt.title(f'Comparison of {capacity_type} Capacity by Energy Source (GW)')
    plt.ylabel('Total Nameplate Capacity (GW)')
    
    # Adjust axis formatting to avoid scientific notation
    ax.ticklabel_format(style='plain', axis='y')

    # Set y-axis limits to fit data
    max_value = df_comparison.max().max()
    ax.set_ylim(0, 4 * max_value)

    handles, labels = ax.get_legend_handles_labels()
        
        # Update legend labels as needed
    updated_labels = [label.replace('cCoal', 'Coal').replace('dOil', 'Oil').replace('eNatural Gas', 'Natural Gas').replace('sNuclear', 'Nuclear').replace('wHydro', 'Hydro').replace('ySolar', 'Solar').replace('xWind', 'Wind').replace('zStorage', 'Storage').replace('aOther', 'Other').replace('bOther RE', 'Other RE') for label in labels]  # Example replacements

        # Set the updated legend
    ax.legend(handles, updated_labels, title='Energy Source', title_fontsize='13', fontsize='11', loc='upper left', bbox_to_anchor=(1, 1))


    # Display the plot
    st.pyplot(fig)
    
    # Calculate and display changes
    changes = {sheet: latest_data[sheet]["Nameplate Capacity (MW)"].sum() - previous_data[sheet]["Nameplate Capacity (MW)"].sum() for sheet in latest_data}
    st.write(f"Change in {capacity_type} Capacity (GW):")
    for sheet, change in changes.items():
        st.write(f"{change / 1000} GW. {latest_data[sheet]['Nameplate Capacity (MW)'].sum() / 1000} GW in latest year, {previous_data[sheet]['Nameplate Capacity (MW)'].sum() / 1000} GW in previous year.")

def plot_plant_comparison(latest_counts, previous_counts, capacity_type):
    """Plots a bar graph comparing the number of plants for the latest and previous months for a specific type."""
    df_comparison = pd.DataFrame({
        'Sheet': latest_counts.keys(),
        'Previous Year': previous_counts.values(),
        'Latest Year': latest_counts.values(),
    })

    fig, ax = plt.subplots()
    df_comparison.plot(x='Sheet', kind='bar', ax=ax)
    plt.title(f'Comparison of Number of {capacity_type} Plants')
    plt.ylabel('Number of Plants')
    
    # Adjust axis formatting to avoid scientific notation
    ax.ticklabel_format(style='plain', axis='y')
    
    # Calculate the min and max values for setting y-axis limits
    all_values = list(latest_counts.values()) + list(previous_counts.values())
    min_value = min(all_values)
    max_value = max(all_values)
    
    # Set y-axis limits with a 10% margin
    margin = 0.1 * (max_value - min_value)
    ax.set_ylim(min_value - margin, max_value + margin)

    st.pyplot(fig)

def plot_technology_pie_charts(latest_df, previous_df, sheet):
    """Plots side-by-side pie charts of technology distribution for the latest and previous months for a specific type."""
    latest_technology_counts = latest_df['Energy Source'].value_counts()
    previous_technology_counts = previous_df['Energy Source'].value_counts()

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 7))

    ax1.pie(latest_technology_counts, labels=latest_technology_counts.index, autopct='%1.1f%%')
    ax1.set_title(f'Latest Year - {sheet} Technology Distribution')

    ax2.pie(previous_technology_counts, labels=previous_technology_counts.index, autopct='%1.1f%%')
    ax2.set_title(f'Previous Year - {sheet} Technology Distribution')

    st.pyplot(fig)

def new_plant_comparison(latest_df, previous_df, sheet):
    """Identifies new plants in the latest month that were not present in the previous month,
    and returns a DataFrame containing the Plant ID and Energy Source."""

    # Ensure the columns are renamed properly
    
    
    # Get the set of Plant IDs for both months
    latest_plant_ids = set(latest_df['Plant ID'].dropna().unique())
    previous_plant_ids = set(previous_df['Plant ID'].dropna().unique())
    
    # Identify new Plant IDs in the latest month
    new_plant_ids = latest_plant_ids - previous_plant_ids
    
    # Create a DataFrame with the new plants and their Energy Sources
    new_plants_df = latest_df[latest_df['Plant ID'].isin(new_plant_ids)][['Plant ID', 'Energy Source', 'Nameplate Capacity (MW)']]

    st.write(f"Number of new plants in latest year for {sheet}: ", len(new_plants_df))
    
    return new_plants_df

def plot_new_plant_pie_chart(new_plants_df, sheet):
    """Plots a pie chart of the energy sources for the new plants with counts and capacities in GW, and aligns colors with the bar graph."""
    
    # Define the color dictionary (aligned with bar graphs)
    color_dict = {
        'cCoal': '#a0522d',         # Light brownish
        'dOil': '#2f4f4f',          # Dark gray
        'eNatural Gas': '#4682b4',  # Slightly more blue for Natural Gas
        'sNuclear': '#ff69b4',      # Pink
        'wHydro': '#003366',        # Navy blue
        'ySolar': "#F5B800",        # Yellowish (towards yellow)
        'xWind': '#39FF14',         # Light emerald green
        'zStorage': '#d3d3d3',      # Light gray
        'aOther': '#808080',        # Medium gray
        'bOther RE': '#d2b48c'      # Light tan
    }

    # Map the 'Energy Source' to the color labels used in the bar chart
    rename_dict = {
        'Coal': 'cCoal',
        'Oil': 'dOil',
        'Natural Gas': 'eNatural Gas',
        'Nuclear': 'sNuclear',
        'Hydro': 'wHydro',
        'Solar': "ySolar",
        'Wind': 'xWind',
        'Storage': 'zStorage',
        'Other': 'aOther',
        'Other RE': 'bOther RE'
    }

    # Apply renaming to ensure consistency in labels
    new_plants_df['Energy Source'] = new_plants_df['Energy Source'].map(rename_dict)

    # Pie chart for plant counts by energy source
    energy_source_counts = new_plants_df['Energy Source'].value_counts()
    
    fig, ax1 = plt.subplots(figsize=(7, 7))
    ax1.pie(energy_source_counts, labels=energy_source_counts.index.map(lambda x: x[1:]), 
            autopct=lambda p: f'{int(p * sum(energy_source_counts) / 100)} plants', 
            colors=[color_dict.get(x, '#d3d3d3') for x in energy_source_counts.index])
    ax1.set_title(f'New {sheet} Plants by Energy Source')
    
    st.pyplot(fig)

    # Pie chart for capacity distribution by energy source (in GW)
    energy_source_capacity = new_plants_df.groupby('Energy Source')['Nameplate Capacity (MW)'].sum()
    total_capacity_mw = new_plants_df['Nameplate Capacity (MW)'].sum()
    st.write("Total capacity for new units: ", total_capacity_mw)
    
    fig, ax2 = plt.subplots(figsize=(7, 7))
    ax2.pie(energy_source_capacity, labels=energy_source_capacity.index.map(lambda x: x[1:]), 
            autopct=lambda p: f'{p * sum(energy_source_capacity) / 100:.1f} MW', 
            colors=[color_dict.get(x, '#d3d3d3') for x in energy_source_capacity.index])
    ax2.set_title(f'New {sheet} Plants Capacity Distribution by Energy Source (MW)')
    
    st.pyplot(fig)



def main():
    st.title("EIA-860 Capacity Data Comparison")

    base_url = "https://www.eia.gov/electricity/data/eia860/xls/eia860"

    latest_url, latest_year, ER_status = get_latest_file_url(base_url)

    if latest_url:
        st.write(f"Latest year ({latest_year}) file URL: {latest_url}")

        previous_url = get_previous_file_url(base_url, latest_year)
        st.write(f"Previous year ({latest_year-1}) file URL: {previous_url}")

        latest_file_io, latest_file = download_excel_file(latest_url, latest_year, ER_status)
        previous_year = latest_year-1
        previous_file_io, previous_file = download_excel_file(previous_url, previous_year, ER_status=False)

        if latest_file and previous_file:
            sheets = ['Operable', 'Proposed', 'Retired and Canceled']
            
            latest_dfs = pd.read_excel(latest_file_io, sheet_name=sheets)
            previous_dfs = pd.read_excel(previous_file_io, sheet_name=sheets)

            latest_sums = {sheet: sum_nameplate_capacity(latest_dfs[sheet], sheet) for sheet in sheets}
            previous_sums = {sheet: sum_nameplate_capacity(previous_dfs[sheet], sheet) for sheet in sheets}

            latest_counts = {sheet: count_plants(latest_dfs[sheet]) for sheet in sheets}
            previous_counts = {sheet: count_plants(previous_dfs[sheet]) for sheet in sheets}

            tab1, tab2, tab3 = st.tabs(sheets)

            with tab1:
                st.header(f"{sheets[0]} Plants")
                states = sorted(latest_dfs[sheets[0]]['State'].dropna().unique())
                selected_state = st.selectbox("Select a state", ["All"] + states, key="op")
                if selected_state != "All":
                    latest_df = latest_dfs[sheets[0]][latest_dfs[sheets[0]]['State'] == selected_state]
                    previous_df = previous_dfs[sheets[0]][previous_dfs[sheets[0]]['State'] == selected_state]
                else:
                    latest_df = latest_dfs[sheets[0]]
                    previous_df = previous_dfs[sheets[0]]

                plot_comparison({sheets[0]: latest_df}, {sheets[0]: previous_df}, 'Operating')
                new_df = new_plant_comparison(latest_df, previous_df, 'Operating')
                plot_new_plant_pie_chart(new_df, 'Operating')

            with tab2:
                st.header(f"{sheets[1]} Plants")
                states = sorted(latest_dfs[sheets[1]]['State'].dropna().unique())
                selected_state = st.selectbox("Select a state", ["All"] + states, key="pl")
                if selected_state != "All":
                    latest_df = latest_dfs[sheets[1]][latest_dfs[sheets[1]]['State'] == selected_state]
                    previous_df = previous_dfs[sheets[1]][previous_dfs[sheets[1]]['State'] == selected_state]
                else:
                    latest_df = latest_dfs[sheets[1]]
                    previous_df = previous_dfs[sheets[1]]

                plot_comparison({sheets[1]: latest_df}, {sheets[1]: previous_df}, 'Proposed')
                new_df = new_plant_comparison(latest_df, previous_df, 'Proposed')
                plot_new_plant_pie_chart(new_df, 'Proposed')

            with tab3:
                st.header(f"{sheets[2]} Plants")
                states = sorted(latest_dfs[sheets[2]]['State'].dropna().unique())
                selected_state = st.selectbox("Select a state", ["All"] + states, key="rt")
                if selected_state != "All":
                    latest_df = latest_dfs[sheets[2]][latest_dfs[sheets[2]]['State'] == selected_state]
                    previous_df = previous_dfs[sheets[2]][previous_dfs[sheets[2]]['State'] == selected_state]
                else:
                    latest_df = latest_dfs[sheets[2]]
                    previous_df = previous_dfs[sheets[2]]

                plot_comparison({sheets[2]: latest_df}, {sheets[2]: previous_df}, 'Retired and Canceled')
                new_df = new_plant_comparison(latest_df, previous_df, 'Retired and Canceled')
                plot_new_plant_pie_chart(new_df, 'Retired and Canceled')

if __name__ == "__main__":
    main()
