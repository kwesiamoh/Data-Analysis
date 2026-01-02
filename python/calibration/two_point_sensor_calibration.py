"""
Simple two-point calibration script for Exercise_1a.xlsx

This script:
 - Reads sheets named 'Zero' and 'Span' (case-insensitive) from Exercise_1a.xlsx
 - Plots Time vs Concentration for each sheet (separate plots)
 - Computes the average concentration for Zero (all values)
 - Computes the average concentration for Span using only measured concentrations > 400
 - Uses actual concentrations [0, 438] and measured averages [zero_avg, span_avg_above_400]
   to fit a linear calibration y = m*x + b (y = measured, x = actual)
 - Plots the calibration points and the fitted calibration line and annotates the equation

Requirements: pandas, numpy, matplotlib, scikit-learn (sklearn is optional; numpy.polyfit is used here)

Usage (PowerShell):
 C:\path\to\python.exe .\Exercise_1a.py

The script assumes the Excel file is in the same folder as this script.
"""

from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


def read_sheet_case_insensitive(xl_path, target):
    xl = pd.ExcelFile(xl_path)
    for name in xl.sheet_names:
        if name.strip().lower() == target.strip().lower():
            return pd.read_excel(xl_path, sheet_name=name)
    raise ValueError(f"Sheet named '{target}' not found in {xl_path}")


def pick_time_conc_columns(df):
    cols = [c for c in df.columns]
    # prefer common names
    lower = {c.strip().lower(): c for c in cols}
    time_col = lower.get('time') or lower.get('timestamp') or cols[0]
    conc_col = lower.get('concentration') or lower.get(
        'conc') or (cols[1] if len(cols) > 1 else cols[0])
    return time_col, conc_col


def main():
    script_dir = Path(__file__).parent
    excel_path = script_dir / 'Exercise_1a.xlsx'
    if not excel_path.exists():
        print(f"Excel file not found at {excel_path}")
        return

    # Read sheets
    zero_df = read_sheet_case_insensitive(excel_path, 'Zero')
    span_df = read_sheet_case_insensitive(excel_path, 'Span')
    print(
        f"Loaded 'Zero' sheet: {zero_df.shape[0]} rows, {zero_df.shape[1]} columns")
    print(
        f"Loaded 'Span' sheet: {span_df.shape[0]} rows, {span_df.shape[1]} columns")

    # Pick columns
    z_time_col, z_conc_col = pick_time_conc_columns(zero_df)
    s_time_col, s_conc_col = pick_time_conc_columns(span_df)
    print(
        f"Detected Zero time column: '{z_time_col}'  concentration column: '{z_conc_col}'")
    print(
        f"Detected Span time column: '{s_time_col}'  concentration column: '{s_conc_col}'")

    # Convert time column to pandas datetime where possible for nicer plotting
    def normalize_time_col(series):
        # First try pandas parsing
        parsed = pd.to_datetime(series, errors='coerce')
        # If parsed contains NaT for everything or remains object (e.g., python time objects),
        # try to convert python datetime.time to Timestamp on a dummy date.
        if parsed.isna().all() or parsed.dtype == 'O':
            def to_timestamp(x):
                if pd.isna(x):
                    return pd.NaT
                # already a Timestamp
                if isinstance(x, pd.Timestamp):
                    return x
                # string or numeric - try parsing
                if isinstance(x, str) or not hasattr(x, 'hour'):
                    try:
                        return pd.to_datetime(x)
                    except Exception:
                        return pd.NaT
                # likely a datetime.time-like object
                try:
                    hour = getattr(x, 'hour', 0)
                    minute = getattr(x, 'minute', 0)
                    second = getattr(x, 'second', 0)
                    micro = getattr(x, 'microsecond', 0)
                    return pd.Timestamp(1900, 1, 1, hour, minute, second, micro)
                except Exception:
                    return pd.NaT

            return series.apply(to_timestamp)
        return parsed

    zero_df['__Time_parsed'] = normalize_time_col(zero_df[z_time_col])
    span_df['__Time_parsed'] = normalize_time_col(span_df[s_time_col])

    # Convert concentration to numeric
    zero_df['__Concentration'] = pd.to_numeric(
        zero_df[z_conc_col], errors='coerce')
    span_df['__Concentration'] = pd.to_numeric(
        span_df[s_conc_col], errors='coerce')

    # Preview loaded/converted columns to verify correctness
    print('\nPreview Zero raw columns:')
    try:
        print(zero_df[[z_time_col, z_conc_col]].head(5).to_string(index=False))
    except Exception:
        print(zero_df.head(5).to_string(index=False))
    print('\nPreview Zero parsed/converted columns:')
    print(zero_df[['__Time_parsed', '__Concentration']].head(
        5).to_string(index=False))
    print(
        f"Zero: {zero_df['__Concentration'].notna().sum()} non-null concentrations, {zero_df['__Concentration'].isna().sum()} NaN")

    print('\nPreview Span raw columns:')
    try:
        print(span_df[[s_time_col, s_conc_col]].head(5).to_string(index=False))
    except Exception:
        print(span_df.head(5).to_string(index=False))
    print('\nPreview Span parsed/converted columns:')
    print(span_df[['__Time_parsed', '__Concentration']].head(
        5).to_string(index=False))
    print(
        f"Span: {span_df['__Concentration'].notna().sum()} non-null concentrations, {span_df['__Concentration'].isna().sum()} NaN\n")

    print('Plotting original Time vs Concentration (Zero & Span)...')
    # Plot original Time vs Concentration data (both Zero and Span)
    plt.figure(figsize=(10, 6))
    plt.plot(zero_df['__Time_parsed'], zero_df['__Concentration'],
             'bo-', label='Zero', markersize=4)
    plt.plot(span_df['__Time_parsed'], span_df['__Concentration'],
             'ro-', label='Span', markersize=4)
    plt.title('Original Measurements: Zero and Span')
    plt.xlabel('Time')
    plt.ylabel('NO Concentration, ppb')
    plt.grid(True)
    plt.legend()
    # Format x-axis to show only time
    plt.gcf().axes[0].xaxis.set_major_formatter(
        plt.matplotlib.dates.DateFormatter('%H:%M:%S'))
    plt.tight_layout()
    plt.show()

    # Compute averages
    zero_avg = zero_df['__Concentration'].mean()
    # Span: only include values above 400
    span_filtered = span_df[span_df['__Concentration']
                            > 400]['__Concentration']
    if len(span_filtered) == 0:
        print('Warning: no Span concentration values above 400 found. Using all Span values for average.')
        span_avg = span_df['__Concentration'].mean()
    else:
        span_avg = span_filtered.mean()

    print(f"Zero average (all values): {zero_avg:.6g}")
    print(f"Span average (values > 400): {span_avg:.6g}")
    # Show how many span values were above 400 and a short preview
    count_above_400 = int((span_df['__Concentration'] > 400).sum())
    print(f"Span: {count_above_400} values > 400 used for average")
    if count_above_400 > 0:
        print('Preview of Span values > 400:')
        print(span_filtered.head(10).to_string(index=False))

    # Actual values for calibration
    x_actual = np.array([0.0, 438.0])
    y_measured = np.array([zero_avg, span_avg])

    # Fit linear regression (measured = m * actual + b)
    m, b = np.polyfit(x_actual, y_measured, 1)
    print(f"\nCalibration fit: measured = {m:.6g} * actual + {b:.6g}")

    # Calculate corrected concentration values using inverse calibration
    def correct_concentration(measured, m, b):
        """Apply inverse calibration: actual = (measured - b) / m"""
        return (measured - b) / m

    # Add corrected concentrations to both dataframes
    zero_df['Corrected_Concentration'] = correct_concentration(
        zero_df['__Concentration'], m, b)
    span_df['Corrected_Concentration'] = correct_concentration(
        span_df['__Concentration'], m, b)

    # Print summary statistics of corrected values
    print("\nCorrected concentration summary:")
    print("Zero data:")
    print(f"  Mean: {zero_df['Corrected_Concentration'].mean():.6g}")
    print(f"  Min:  {zero_df['Corrected_Concentration'].min():.6g}")
    print(f"  Max:  {zero_df['Corrected_Concentration'].max():.6g}")
    print("Span data:")
    print(f"  Mean: {span_df['Corrected_Concentration'].mean():.6g}")
    print(f"  Min:  {span_df['Corrected_Concentration'].min():.6g}")
    print(f"  Max:  {span_df['Corrected_Concentration'].max():.6g}")

    # Actual values
    x_actual = np.array([0.0, 438.0])
    y_measured = np.array([zero_avg, span_avg])

    # Fit linear regression (measured = m * actual + b)
    # Use numpy.polyfit for degree 1
    m, b = np.polyfit(x_actual, y_measured, 1)
    print(f"Calibration fit: measured = {m:.6g} * actual + {b:.6g}")

    # Plot calibration (actual on x-axis, measured on y-axis)
    plt.figure(figsize=(7, 6))
    # Scatter the two calibration points
    plt.scatter(x_actual, y_measured, color='red', label='Calibration points')
    # Plot regression line across a little beyond 0..438 for clarity
    x_line = np.linspace(0, 450, 200)
    y_line = m * x_line + b
    plt.plot(x_line, y_line, color='blue',
             label=f'Fit: y = {m:.4g}x + {b:.4g}')
    plt.xlabel('Actual concentration (ppb)')
    plt.ylabel('Measured concentration (ppb)')
    plt.title('Calibration Curve')
    plt.legend()
    # Annotate equation on the plot
    eq_text = f"y = {m:.4g}x + {b:.4g}"
    plt.text(0.05, 0.95, eq_text, transform=plt.gca().transAxes,
             fontsize=10, verticalalignment='top', bbox=dict(facecolor='white', alpha=0.8))
    plt.grid(True)
    plt.tight_layout()
    plt.show()

    print('Plotting corrected concentrations...')
    # Plot only corrected concentrations
    plt.figure(figsize=(10, 6))
    plt.plot(zero_df['__Time_parsed'], zero_df['Corrected_Concentration'],
             'yo-', label='Zero (Corrected)', markersize=4)
    plt.plot(span_df['__Time_parsed'], span_df['Corrected_Concentration'],
             'go-', label='Span (Corrected)', markersize=4)
    plt.title('Corrected Concentration Data (0 - 438 ppb Calibration)')
    plt.xlabel('Time')
    plt.ylabel('NO Concentration, ppb')
    plt.legend()
    plt.grid(True)
    # Format x-axis to show only time
    plt.gcf().axes[0].xaxis.set_major_formatter(
        plt.matplotlib.dates.DateFormatter('%H:%M:%S'))
    plt.tight_layout()
    plt.show()

    # Plot comparing original vs corrected concentrations
    plt.figure(figsize=(10, 6))

    # Zero data
    plt.plot(zero_df['__Time_parsed'], zero_df['__Concentration'],
             'b-', label='Zero (Measured)', marker='o', markersize=2)
    plt.plot(zero_df['__Time_parsed'], zero_df['Corrected_Concentration'],
             'y--', label='Zero (Corrected)', marker='s', markersize=2)

    # Span data
    plt.plot(span_df['__Time_parsed'], span_df['__Concentration'],
             'r-', label='Span (Measured)', marker='o', markersize=2)
    plt.plot(span_df['__Time_parsed'], span_df['Corrected_Concentration'],
             'g--', label='Span (Corrected)', marker='s', markersize=2)

    plt.title('Measured vs Corrected Concentrations (Overlapped)')
    plt.xlabel('Time')
    plt.ylabel('NO Concentration, ppb')
    plt.legend()
    plt.grid(True)
    # Format x-axis to show only time
    plt.gcf().axes[0].xaxis.set_major_formatter(
        plt.matplotlib.dates.DateFormatter('%H:%M:%S'))
    plt.tight_layout()
    plt.show()

    print('\nProcessing complete.')


if __name__ == '__main__':
    main()
