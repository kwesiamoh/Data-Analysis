"""
Air Quality Data Analysis and Correction Script
===============================================
This script loads air quality data from Exercise_1b.xlsx, performs linear regression
correction on raw device measurements, and generates three pages of visualizations:
1. Raw data time series (3 pollutants)
2. Corrected data time series (3 pollutants)
3. Side-by-side raw vs. corrected comparison (3 pollutants)
"""

from datetime import date
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from scipy.stats import linregress

# Helper function to format pollutant names with subscripts


def format_pollutant(pollutant):
    """Convert pollutant names to subscript format for plotting."""
    if pollutant == 'PM10':
        return 'PM$_{10}$'
    elif pollutant == 'PM25':
        return 'PM$_{2.5}$'
    elif pollutant == 'PM1':
        return 'PM$_{1}$'
    return pollutant

# Helper function to format time axis


def format_time_axis(ax):
    """Format the x-axis to show only time (HH:MM:SS)."""
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))

# ==============================================================================
# STEP 1: LOAD DATA
# ==============================================================================


# Load the Excel file
df = pd.read_excel('Exercise_1b.xlsx')

# Parse the Time column as datetime
# Handle the case where Time column contains datetime.time objects
if df['Time'].dtype == 'object':
    # Convert time objects to datetime using a reference date (1900-01-01)
    df['Time'] = df['Time'].apply(lambda x: pd.Timestamp(
        datetime.combine(date(1900, 1, 1), x)))

print("Data loaded successfully!")
print(f"Shape: {df.shape}")
print(f"Columns: {df.columns.tolist()}")
print(f"\nFirst few rows:\n{df.head()}")
print('\nPreview specific columns to verify correct loading:')
sample_raw_cols = [c for c in df.columns if 'RawData' in c][:6]
preview_cols = ['Time'] + sample_raw_cols
print(df[preview_cols].head(5).to_string(index=False))

# ==============================================================================
# STEP 2: DATA CORRECTION - LINEAR REGRESSION
# ==============================================================================

# Define pollutants and devices
pollutants = ['PM10', 'PM25', 'PM1']
devices = [1, 2, 3, 4]

# Dictionary to store regression parameters for each device/pollutant
regression_params = {}

# Step 2.1: Calculate average for each pollutant across all devices
for pollutant in pollutants:
    # Get raw data columns for this pollutant
    raw_cols = [f'Dev{d}_RawData_{pollutant}' for d in devices]

    # Calculate average (reference value)
    df[f'Avg_{pollutant}'] = df[raw_cols].mean(axis=1)

    print(f"\nProcessing {pollutant}:")
    print(f"  Average column created: Avg_{pollutant}")

# After creating averages, show a short preview per pollutant
print('\nPreview of Avg and raw columns (first 5 rows)')
for pollutant in pollutants:
    raw_cols = [f'Dev{d}_RawData_{pollutant}' for d in devices]
    avg_col = f'Avg_{pollutant}'
    available = [c for c in (['Time'] + raw_cols +
                             [avg_col]) if c in df.columns]
    print(f"\n{pollutant} columns: {', '.join(available)}")
    print(df[available].head(5).to_string(index=False))
    for c in raw_cols:
        if c in df.columns:
            print(
                f"  {c}: {df[c].notna().sum()} non-null, {df[c].isna().sum()} NaN")

# Step 2.2: Perform linear regression for each device/pollutant combination
for pollutant in pollutants:
    for device in devices:
        raw_col = f'Dev{device}_RawData_{pollutant}'
        avg_col = f'Avg_{pollutant}'

        # Get data (remove NaN values)
        valid_mask = df[[raw_col, avg_col]].notna().all(axis=1)
        x = df.loc[valid_mask, avg_col].values
        y = df.loc[valid_mask, raw_col].values

        # Perform linear regression: y = mx + c
        slope, intercept, r_value, p_value, std_err = linregress(x, y)

        # Store regression parameters
        regression_params[(device, pollutant)] = {
            'slope': slope,
            'intercept': intercept,
            'r_squared': r_value ** 2
        }

        print(
            f"  Dev{device}: y = {slope:.4f}x + {intercept:.4f} (R² = {r_value**2:.4f})")

# Step 2.3: Calculate corrected values using inverse regression
# Corrected Value = (Raw Device Value - c) / m
for pollutant in pollutants:
    for device in devices:
        raw_col = f'Dev{device}_RawData_{pollutant}'
        corrected_col = f'Dev{device}_Corrected_{pollutant}'

        m = regression_params[(device, pollutant)]['slope']
        c = regression_params[(device, pollutant)]['intercept']

        # Apply inverse regression formula
        df[corrected_col] = (df[raw_col] - c) / m

        print(f"  Created corrected column: {corrected_col}")

print("\n" + "="*80)
print("Data correction completed!")
print("="*80)

print('\nPreview of raw vs corrected columns (first 5 rows):')
for pollutant in pollutants:
    raw_cols = [f'Dev{d}_RawData_{pollutant}' for d in devices]
    corr_cols = [f'Dev{d}_Corrected_{pollutant}' for d in devices]
    cols_to_show = ['Time'] + [c for c in raw_cols +
                               corr_cols + [f'Avg_{pollutant}'] if c in df.columns]
    print(f"\n{pollutant} preview columns: {', '.join(cols_to_show)}")
    print(df[cols_to_show].head(5).to_string(index=False))
    for c in raw_cols + corr_cols:
        if c in df.columns:
            print(
                f"  {c}: {df[c].notna().sum()} non-null, {df[c].isna().sum()} NaN")

# ==============================================================================
# STEP 3: VISUALIZATION - PAGE 1: RAW DATA TIME SERIES
# ==============================================================================

print('Plotting page 1: Raw data time series...')
fig1, axes1 = plt.subplots(3, 1, figsize=(14, 10))
fig1.suptitle('Raw PM Data Time Series',
              fontsize=16, fontweight='bold', y=0.995)

for idx, pollutant in enumerate(pollutants):
    ax = axes1[idx]

    for device in devices:
        raw_col = f'Dev{device}_RawData_{pollutant}'
        ax.plot(df['Time'], df[raw_col],
                label=f'Device {device}', linewidth=1.5)

    pollutant_formatted = format_pollutant(pollutant)
    ax.set_ylabel(f'{pollutant_formatted} Concentration (µg/m³)',
                  fontsize=11, fontweight='bold')
    ax.set_title(f'Raw {pollutant_formatted} Data - All Devices',
                 fontsize=12, fontweight='bold')
    ax.legend(loc='best', ncol=4, framealpha=0.9)
    ax.grid(True, alpha=0.3)

    # Format x-axis
    if idx == 2:
        format_time_axis(ax)
        ax.set_xlabel('Time', fontsize=11, fontweight='bold')
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
    else:
        ax.set_xticklabels([])

plt.tight_layout()
plt.show()

# ==============================================================================
# STEP 4: VISUALIZATION - PAGE 2: CORRECTED DATA TIME SERIES
# ==============================================================================

print('Plotting page 2: Corrected data time series...')
fig2, axes2 = plt.subplots(3, 1, figsize=(14, 10))
fig2.suptitle('Corrected PM Data Time Series',
              fontsize=16, fontweight='bold', y=0.995)

for idx, pollutant in enumerate(pollutants):
    ax = axes2[idx]

    for device in devices:
        corrected_col = f'Dev{device}_Corrected_{pollutant}'
        ax.plot(df['Time'], df[corrected_col],
                label=f'Device {device}', linewidth=1.5)

    pollutant_formatted = format_pollutant(pollutant)
    ax.set_ylabel(f'{pollutant_formatted} Concentration (µg/m³)',
                  fontsize=11, fontweight='bold')
    ax.set_title(f'Corrected {pollutant_formatted} Data - All Devices',
                 fontsize=12, fontweight='bold')
    ax.legend(loc='best', ncol=4, framealpha=0.9)
    ax.grid(True, alpha=0.3)

    # Format x-axis
    if idx == 2:
        format_time_axis(ax)
        ax.set_xlabel('Time', fontsize=11, fontweight='bold')
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
    else:
        ax.set_xticklabels([])

plt.tight_layout()
plt.show()

# ==============================================================================
# STEP 5: VISUALIZATION - PAGE 3: SIDE-BY-SIDE COMPARISON (3x2 GRID)
# ==============================================================================

print('Plotting page 3: Side-by-side raw vs corrected comparison...')
fig3, axes3 = plt.subplots(3, 2, figsize=(16, 11))
fig3.suptitle('Raw vs. Corrected PM Data Comparison',
              fontsize=16, fontweight='bold', y=0.995)

for row, pollutant in enumerate(pollutants):
    # Left column: Raw data
    ax_left = axes3[row, 0]
    for device in devices:
        raw_col = f'Dev{device}_RawData_{pollutant}'
        ax_left.plot(df['Time'], df[raw_col],
                     label=f'Device {device}', linewidth=1.5)

    pollutant_formatted = format_pollutant(pollutant)
    ax_left.set_ylabel(f'{pollutant_formatted} (µg/m³)',
                       fontsize=11, fontweight='bold')
    ax_left.set_title(f'Raw {pollutant_formatted} Data',
                      fontsize=12, fontweight='bold')
    ax_left.legend(loc='best', ncol=2, framealpha=0.9)
    ax_left.grid(True, alpha=0.3)

    # Right column: Corrected data
    ax_right = axes3[row, 1]
    for device in devices:
        corrected_col = f'Dev{device}_Corrected_{pollutant}'
        ax_right.plot(df['Time'], df[corrected_col],
                      label=f'Device {device}', linewidth=1.5)

    ax_right.set_ylabel(f'{pollutant_formatted} (µg/m³)',
                        fontsize=11, fontweight='bold')
    ax_right.set_title(f'Corrected {pollutant_formatted} Data',
                       fontsize=12, fontweight='bold')
    ax_right.legend(loc='best', ncol=2, framealpha=0.9)
    ax_right.grid(True, alpha=0.3)

    # Format x-axis for bottom row only
    if row == 2:
        format_time_axis(ax_left)
        format_time_axis(ax_right)
        ax_left.set_xlabel('Time', fontsize=11, fontweight='bold')
        ax_right.set_xlabel('Time', fontsize=11, fontweight='bold')
        plt.setp(ax_left.xaxis.get_majorticklabels(), rotation=45, ha='right')
        plt.setp(ax_right.xaxis.get_majorticklabels(), rotation=45, ha='right')
    else:
        ax_left.set_xticklabels([])
        ax_right.set_xticklabels([])

plt.tight_layout()
plt.show()

# ==============================================================================
# SUMMARY
# ==============================================================================

print("\n" + "="*80)
print("VISUALIZATION COMPLETE!")
print("="*80)
print("\nThree visualization pages have been generated:")
print("Raw PM Data Time Series (3 subplots)")
print("Corrected PM Data Time Series (3 subplots)")
print("Side-by-Side Raw vs. Corrected Comparison (3x2 grid)")
print("\nRegression Parameters Summary:")
print("-" * 80)

for pollutant in pollutants:
    print(f"\n{pollutant}:")
    for device in devices:
        m = regression_params[(device, pollutant)]['slope']
        c = regression_params[(device, pollutant)]['intercept']
        r2 = regression_params[(device, pollutant)]['r_squared']
        print(f"  Device {device}: y = {m:.6f}x + {c:.6f} (R² = {r2:.6f})")
