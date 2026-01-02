"""
Mastersheet.py

Summary:
This script loads a meteorology master file (Meteorology.txt) and several
auxiliary sensor files (PM.xlsx, UFP.xlsx, Memocomp.xlsx, BC.csv), filters
data to a fixed time window, and merges records by an exact timestamp match
based on seconds from midnight. It outputs a merged Excel file (Mastersheet.xlsx)
and optionally generates pollutant heatmaps (HTML) if Folium is installed.

Inputs:
- Meteorology.txt (required): master file with a timestamp column and
    optional GPS and meteorology columns.
- PM.xlsx, UFP.xlsx, Memocomp.xlsx, BC.csv (optional): auxiliary sensor files
    with time/date columns and pollutant data.

Outputs:
- Mastersheet.xlsx: Excel file with merged columns for the selected time window.
- HTML heatmaps: One HTML map per pollutant (requires Folium/Banca).

Usage:
Run the script directly with Python; ensure input files are in the same
directory. Example: `python Mastersheet.py`.
"""

import pandas as pd
import numpy as np
import os
import datetime
import webbrowser
import tempfile

try:
    import folium
    from folium.plugins import HeatMap
except Exception:
    folium = None

try:
    from branca.colormap import LinearColormap
except Exception:
    LinearColormap = None

POLLUTANTS = ['PM10', 'PM2_5', 'PM1', 'UFP', 'NO2', 'NO', 'O3', 'BC']

# Heatmap palette choice: use the custom AQI-like palette by default.
HEATMAP_PALETTE = 'AQI_CUSTOM'

# Palette definitions (hex color stops from light -> intense)
PALETTES = {
    'YlOrRd': ['#ffffb2', '#fecc5c', '#fd8d3c', '#f03b20', '#bd0026'],
    'Inferno': ['#000004', '#3b0f70', '#8c2981', '#de4968', '#fe9f6d'],
    'Viridis': ['#440154', '#31688e', '#35b778', '#fde725', '#ffffbf'],
    # Custom AQI-style palette (Good -> Very Poor)
    'AQI_CUSTOM': ['#50F0E6', '#50CCAA', '#F0E641', '#FF5050', '#960032']
}

# --- Configuration ---
START_TIME = datetime.time(21, 49)
END_TIME = datetime.time(22, 58)
OUTPUT_FILENAME = 'Mastersheet.xlsx'


def get_seconds_from_midnight(dt_series):
    """
    Converts a datetime series to an integer representing seconds from midnight.
    Example: 21:49:00 -> 78540
    This allows exact matching even if the dates (years/months) are different.
    """
    return dt_series.dt.hour * 3600 + dt_series.dt.minute * 60 + dt_series.dt.second


def load_master_meteorology():
    print("--- Loading Meteorology.txt (Master) ---")
    if not os.path.exists('Meteorology.txt'):
        print("CRITICAL ERROR: Meteorology.txt not found.")
        return None

    # Read file
    try:
        df = pd.read_csv('Meteorology.txt', sep=',')
    except:
        # Fallback if file encoding is weird
        df = pd.read_csv('Meteorology.txt', sep=',', encoding='latin1')

    df.columns = df.columns.str.strip()

    # Identify Timestamp (Assuming first column)
    time_col = df.columns[0]

    # Convert to Datetime
    df['Temp_Timestamp'] = pd.to_datetime(
        df[time_col], dayfirst=True, errors='coerce')

    # Filter Time Window (21:49 to 22:58)
    df = df[df['Temp_Timestamp'].dt.time.between(START_TIME, END_TIME)].copy()

    if df.empty:
        print("Error: No data found in Meteorology for specified time.")
        return None

    # Create Universal Merge Key (Seconds)
    df['Merge_Seconds'] = get_seconds_from_midnight(df['Temp_Timestamp'])

    # Clean Date and Time columns for final output
    df['Date'] = df['Temp_Timestamp'].dt.date
    df['Time'] = df['Temp_Timestamp'].dt.time
    df['S_Date_and_Time'] = df['Temp_Timestamp']

    # --- GPS SPLIT LOGIC (COLON SEPARATOR) ---
    if 'GPS_Data' in df.columns:
        print("Splitting GPS Data by ':'...")
        # Split by colon, expand=True creates new columns
        gps = df['GPS_Data'].astype(str).str.split(':', expand=True)

        # Map split columns safely
        if gps.shape[1] >= 1:
            df['GPS_Latitude'] = gps[0].str.strip()
        else:
            df['GPS_Latitude'] = np.nan

        if gps.shape[1] >= 2:
            df['GPS_Longitude'] = gps[1].str.strip()
        else:
            df['GPS_Longitude'] = np.nan

        if gps.shape[1] >= 3:
            df['GPS_Height_Location'] = gps[2].str.strip()
        else:
            df['GPS_Height_Location'] = np.nan
    else:
        df['GPS_Latitude'] = np.nan
        df['GPS_Longitude'] = np.nan
        df['GPS_Height_Location'] = np.nan

    print(f"Master loaded: {len(df)} rows.")

    # Preview some loaded values to help confirm correct parsing
    try:
        preview_cols = ['Temp_Timestamp', 'GPS_Latitude',
                        'GPS_Longitude', 'S_Date_and_Time']
        preview_cols = [c for c in preview_cols if c in df.columns]
        if preview_cols:
            print("Master preview (first 5 rows):")
            print(df[preview_cols].head(5).to_string(index=False))

        # Numeric columns that are commonly expected in Meteorology
        num_preview = [c for c in ['Temperature', 'Relative_Humidity',
                                   'Pressure', 'Solar_Radiation'] if c in df.columns]
        if num_preview:
            print("Numeric preview (first 5 rows):")
            print(df[num_preview].head(5).to_string(index=False))
    except Exception as e:
        print(f"Preview error: {e}")

    return df


def load_aux_file(filename, required_cols, data_name, sep=None):
    """
    Generic function to load PM, UFP, Memocomp, BC.
    It returns a DataFrame with 'Merge_Seconds' and the Data Columns.
    """
    print(f"--- Loading {data_name} ({filename}) ---")
    if not os.path.exists(filename):
        print(f"Warning: {filename} not found. Columns will be empty.")
        return None

    try:
        # Determine file type
        if filename.endswith('.xlsx'):
            df = pd.read_excel(filename)
        elif filename.endswith('.csv') or filename.endswith('.txt'):
            df = pd.read_csv(filename, sep=sep if sep else ',')

        df.columns = df.columns.str.strip()

        # Find Date and Time columns to construct a timestamp
        # Priority: Look for 'Time' and 'Date' columns
        time_col = next((c for c in df.columns if 'time' in c.lower()), None)
        date_col = next((c for c in df.columns if 'date' in c.lower()), None)

        combined_dt = None

        if time_col and date_col:
            # Combine Date + Time string to handle AM/PM correctly
            combined_dt = pd.to_datetime(
                df[date_col].astype(str) + ' ' + df[time_col].astype(str),
                errors='coerce',
                dayfirst=True
            )
        elif time_col:
            # Just Time available
            combined_dt = pd.to_datetime(
                df[time_col], errors='coerce', dayfirst=True)

        if combined_dt is None:
            print(f"Error: Could not parse time in {filename}")
            return None

        # Drop bad time rows
        df = df.loc[combined_dt.notna()].copy()
        combined_dt = combined_dt[combined_dt.notna()]

        # Filter Window
        mask = combined_dt.dt.time.between(START_TIME, END_TIME)
        df = df.loc[mask].copy()
        combined_dt = combined_dt.loc[mask]

        # Create Key
        df['Merge_Seconds'] = get_seconds_from_midnight(combined_dt)

        # Keep only Seconds Key + Required Data Columns
        # Filter required_cols to only those that actually exist in file
        actual_cols = [c for c in required_cols if c in df.columns]

        if not actual_cols:
            print(
                f"Warning: None of the columns {required_cols} found in {filename}")
            return None

        # Handle duplicates: If two rows have exact same second, keep first.
        df = df.drop_duplicates(subset=['Merge_Seconds'])

        # Print a small preview so the user can confirm correct columns/values
        try:
            print(f"Found columns in {filename}: {actual_cols}")
            preview_cols = ['Merge_Seconds'] + actual_cols
            print(f"{data_name} preview (first 5 rows):")
            print(df[preview_cols].head(5).to_string(index=False))
            print("Column dtypes:")
            print(df[actual_cols].dtypes.to_string())
        except Exception as e:
            print(f"Preview error for {filename}: {e}")

        return df[['Merge_Seconds'] + actual_cols]

    except Exception as e:
        print(f"Error processing {filename}: {e}")
        return None


def main():
    # 1. Load Master
    master_df = load_master_meteorology()
    if master_df is None:
        return

    # 2. Load Aux Files
    # PM Data
    pm_df = load_aux_file('PM.xlsx', ['PM10', 'PM2_5', 'PM1'], 'PM Data')

    # UFP Data
    ufp_df = load_aux_file('UFP.xlsx', ['UFP'], 'UFP Data')

    # Memocomp Data (NO2, NO, O3)
    memo_df = load_aux_file(
        'Memocomp.xlsx', ['NO2', 'NO', 'O3'], 'Memocomp Data')

    # BC Data
    bc_df = load_aux_file('BC.csv', ['BC'], 'BC Data', sep=';')

    # 3. MERGE - EXACT MATCH (Left Join)
    # This ensures we DO NOT REPEAT values. If time matches, copy. If not, leave blank.
    print("\n--- Merging Data (Exact Match Only) ---")

    if pm_df is not None:
        print(f"Merging PM... ({len(pm_df)} rows found)")
        master_df = pd.merge(master_df, pm_df, on='Merge_Seconds', how='left')

    if ufp_df is not None:
        print(f"Merging UFP... ({len(ufp_df)} rows found)")
        master_df = pd.merge(master_df, ufp_df, on='Merge_Seconds', how='left')

    if memo_df is not None:
        print(f"Merging Memocomp... ({len(memo_df)} rows found)")
        master_df = pd.merge(master_df, memo_df,
                             on='Merge_Seconds', how='left')

    if bc_df is not None:
        print(f"Merging BC... ({len(bc_df)} rows found)")
        master_df = pd.merge(master_df, bc_df, on='Merge_Seconds', how='left')

    # 4. Formatting and Saving
    final_cols = [
        'Date', 'Time',
        'PM10', 'PM2_5', 'PM1',
        'UFP',
        'NO2', 'NO', 'O3',
        'BC',
        'Corrected_Wind_Direction', 'GPS_Corrected_Speed',
        'Pressure', 'Relative_Humidity', 'Temperature', 'Solar_Radiation',
        'GPS_Data', 'S_Date_and_Time',
        'GPS_Latitude', 'GPS_Longitude', 'GPS_Height_Location'
    ]

    # Fill missing columns with NaN so Excel structure is maintained
    for col in final_cols:
        if col not in master_df.columns:
            master_df[col] = np.nan

    master_df = master_df[final_cols]

    try:
        print(f"Preparing to save to {OUTPUT_FILENAME}...")
        try:
            print("Merged master preview (first 5 rows):")
            print(master_df.head(5).to_string(index=False))
        except Exception:
            pass
        print(f"Saving to {OUTPUT_FILENAME}...")
        master_df.to_excel(OUTPUT_FILENAME, index=False)
        print(f"\nSUCCESS: {OUTPUT_FILENAME} created.")
    except Exception as e:
        # Handle permission denied (file open in Excel) by saving to an alternative filename
        if isinstance(e, PermissionError) or 'Permission denied' in str(e):
            alt_name = OUTPUT_FILENAME
            if alt_name.lower().endswith('.xlsx'):
                alt_name = alt_name[:-5] + '_' + \
                    datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx'
            else:
                alt_name = alt_name + '_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            try:
                master_df.to_excel(alt_name, index=False)
                print(
                    f"Warning: Could not overwrite {OUTPUT_FILENAME}. Saved as {alt_name}")
            except Exception as e2:
                print(f"Error saving to alternative file {alt_name}: {e2}")
        else:
            print(f"Error saving file: {e}")

    # After creating the Mastersheet, generate heatmaps (HTML + optional PNG)
    try:
        if folium is None:
            print('Folium not installed; skipping heatmap generation.')
        else:
            generate_heatmaps_from_master(master_df)
    except Exception as e:
        print(f"Error generating heatmaps: {e}")


def format_pollutant_html(pollutant):
    """
    Convert pollutant names to HTML with subscripts.
    E.g., 'PM10' -> 'PM<sub>10</sub>', 'PM2_5' -> 'PM<sub>2.5</sub>', 'O3' -> 'O<sub>3</sub>', 'NO2' -> 'NO<sub>2</sub>'
    """
    mapping = {
        'PM10': 'PM<sub>10</sub> (µg/m<sup>3</sup>)',
        'PM2_5': 'PM<sub>2.5</sub> (µg/m<sup>3</sup>)',
        'PM1': 'PM<sub>1</sub> (µg/m<sup>3</sup>)',
        'O3': 'O<sub>3</sub> (µg/m<sup>3</sup>)',
        'NO2': 'NO<sub>2</sub> (µg/m<sup>3</sup>)',
        'UFP': 'UFP (numbers/cm<sup>3</sup>)',
        'NO': 'NO (µg/m<sup>3</sup>)',
        'BC': 'BC (ng/m<sup>3</sup>)'
    }
    return mapping.get(pollutant, pollutant)


def clean_coord(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    s = s.replace(',', '.')
    filtered = ''.join(ch for ch in s if (ch.isdigit() or ch in '+-.'))
    try:
        return float(filtered)
    except Exception:
        return None


def generate_heatmaps_from_master(master_df):
    """
    Create folium heatmaps for pollutants using `GPS_Latitude` and `GPS_Longitude` columns
    and save as HTML files. Each HTML map is opened in the default web browser.
    """
    if 'GPS_Latitude' not in master_df.columns or 'GPS_Longitude' not in master_df.columns:
        print('No GPS columns found in master dataframe; skipping heatmaps.')
        return

    # Clean coordinates
    master_df = master_df.copy()
    master_df['_lat'] = master_df['GPS_Latitude'].apply(clean_coord)
    master_df['_lon'] = master_df['GPS_Longitude'].apply(clean_coord)
    master_df = master_df[master_df['_lat'].notna() &
                          master_df['_lon'].notna()]
    master_df = master_df[master_df['_lat'].between(
        -90, 90) & master_df['_lon'].between(-180, 180)]

    if master_df.empty:
        print('No valid GPS points after cleaning. Cannot create heatmaps.')
        return

    center = [float(master_df['_lat'].mean()), float(master_df['_lon'].mean())]

    # Extract start and end times from the data for the title
    start_dt = None
    end_dt = None
    if 'S_Date_and_Time' in master_df.columns:
        times = master_df['S_Date_and_Time'].dropna()
        if not times.empty:
            start_dt = times.min()
            end_dt = times.max()

    # (PNG rendering via Selenium was removed — the script only writes HTML and opens it.)

    for idx, pollutant in enumerate(POLLUTANTS):
        print(
            f"Generating heatmap for {pollutant} ({idx+1}/{len(POLLUTANTS)})...")
        if pollutant not in master_df.columns:
            print(f'Skipping {pollutant}: column not found')
            continue

        sub = master_df[master_df[pollutant].notna()].copy()
        if sub.empty:
            print(f'Skipping {pollutant}: no non-null values')
            continue

        heat_data = []
        for _, r in sub.iterrows():
            try:
                w = float(r[pollutant])
            except Exception:
                continue
            if np.isnan(w):
                continue
            heat_data.append([r['_lat'], r['_lon'], w])

        if not heat_data:
            print(f'No usable points for {pollutant}')
            continue

        weights = [w for _, _, w in heat_data]
        wmin, wmax = min(weights), max(weights)
        print(f"{pollutant}: {len(heat_data)} points, value range {wmin} — {wmax}")
        if wmax > wmin:
            # Normalize to 0..1 then apply a gamma to increase visual contrast.
            # gamma < 1 spreads low/medium values (makes yellows/reds more visible).
            gamma = 0.6
            norm = [[lat, lon, ((val - wmin) / (wmax - wmin)) ** gamma]
                    for lat, lon, val in heat_data]
        else:
            norm = [[lat, lon, 1.0] for lat, lon, val in heat_data]

        m = folium.Map(location=center, zoom_start=13)
        # Select palette stops (light -> intense). If requested palette missing,
        # fall back to 'Inferno'. Create gradient stops evenly spaced.
        palette = PALETTES.get(HEATMAP_PALETTE, PALETTES['Inferno'])
        stops = [0.0, 0.25, 0.5, 0.75, 1.0]
        gradient = {s: c for s, c in zip(stops, palette)}
        HeatMap(norm, gradient=gradient, min_opacity=0.35,
                radius=15, blur=20, max_zoom=18).add_to(m)

        # Add a custom title overlay with investigation description, timestamps, and formatted pollutant name
        title_lines = [
            "Investigation of Air Quality in Stuttgart using Bicycle Measurements"
        ]

        # Add start and end times if available
        if start_dt and end_dt:
            title_lines.append(
                f"Start: {start_dt.strftime('%Y-%m-%d %H:%M:%S')}  |  End: {end_dt.strftime('%Y-%m-%d %H:%M:%S')}")

        # Add formatted pollutant name
        pollutant_html = format_pollutant_html(pollutant)
        title_lines.append(f"Pollutant: {pollutant_html}")

        # Build the HTML title box with multiple lines
        title_content = '<br>'.join(title_lines)
        title_html = f"""
        <div style="position: fixed; top: 10px; left:50%; transform: translateX(-50%); z-index:9999; background: rgba(255,255,255,0.95); padding:10px 15px; border-radius:6px; font-weight:600; text-align: center; max-width:700px;">
            {title_content}
        </div>
        """
        m.get_root().html.add_child(folium.Element(title_html))

        # Add a colorbar / legend. Prefer branca LinearColormap if available.
        try:
            if LinearColormap is not None:
                # Use the same palette for the colorbar/legend
                palette = PALETTES.get(HEATMAP_PALETTE, PALETTES['Inferno'])
                cmap = LinearColormap(palette, vmin=wmin, vmax=wmax)
                # Get unit for pollutant
                unit = "numbers/cm³" if pollutant == "UFP" else "ng/m³" if pollutant == "BC" else "µg/m³"
                cmap.caption = f"{pollutant} ({unit}, {wmin:.2f} — {wmax:.2f})"
                cmap.add_to(m)
            else:
                # Fallback simple legend box showing min/max with units
                legend_html = f"""
                <div style="position: fixed; bottom: 50px; left: 10px; z-index:9999; background: white; padding:8px; border-radius:5px; box-shadow: 0 0 6px rgba(0,0,0,0.3);">
                    <b>{pollutant} ({unit})</b><br>
                    min: {wmin:.2f} {unit}<br>
                    max: {wmax:.2f} {unit}
                </div>
                """
                m.get_root().html.add_child(folium.Element(legend_html))
            # Also add a categorical legend (Good -> Very Poor) with the
            # user-provided colors so the same labels appear on every map.
            try:
                cat_palette = PALETTES.get('AQI_CUSTOM')
                if cat_palette:
                    # Labels in order corresponding to the palette (low -> high)
                    cat_labels = ['Good', 'Fair',
                                  'Moderate', 'Poor', 'Very Poor']
                    items_html = ''.join([
                        f"<div style='display:flex;align-items:center;margin:2px 0;'>"
                        f"<div style='width:18px;height:12px;background:{col};margin-right:8px;border:1px solid #444;'></div>"
                        f"<div style='font-size:12px;color:#222;'>{label}</div>"
                        f"</div>"
                        for col, label in zip(cat_palette, cat_labels)
                    ])

                    categories_html = f"""
                    <div style="position: fixed; bottom: 10px; left: 10px; z-index:9999; background: rgba(255,255,255,0.95); padding:8px; border-radius:5px; box-shadow: 0 0 6px rgba(0,0,0,0.25);">
                        <b style="display:block;margin-bottom:6px;">Legend</b>
                        {items_html}
                    </div>
                    """
                    m.get_root().html.add_child(folium.Element(categories_html))
            except Exception:
                pass
        except Exception:
            # If legend insertion fails, continue silently
            pass

        # Render HTML to a temporary file (do not keep files in working directory)
        html_str = m.get_root().render()
        tmp = tempfile.NamedTemporaryFile(
            delete=False, suffix=f'_heatmap_{pollutant}.html')
        tmp_path = tmp.name
        try:
            tmp.write(html_str.encode('utf-8'))
            tmp.flush()
            tmp.close()
        except Exception:
            try:
                tmp.close()
            except Exception:
                pass

        print(f'Saved temporary heatmap HTML: {tmp_path} ({len(norm)} points)')
        try:
            webbrowser.open('file:///' + os.path.abspath(tmp_path))
        except Exception:
            pass


if __name__ == "__main__":
    main()
