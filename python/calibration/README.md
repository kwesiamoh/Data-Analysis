# Sensor Calibration Scripts

This folder contains Python scripts related to sensor calibration and signal correction.

## Scripts

"""
Script: two_point_sensor_calibration.py
Performs a two-point calibration of NO concentration sensor data using zero and span reference measurements stored in an Excel file.  
The script computes calibration coefficients, applies inverse calibration, and visualizes measured versus corrected concentration values over time.



"""
Script: air_quality_multi_sensor_correction.py

Purpose:
Correct air quality sensor measurements using linear regression against
device-averaged reference values.

Description:
- Loads particulate matter (PM1, PM2.5, PM10) data from an Excel file
- Computes average pollutant concentrations across multiple devices
- Performs linear regression for each device–pollutant combination
- Applies inverse regression to correct raw sensor measurements
- Generates time-series visualizations of raw, corrected, and comparative data

Tools:
pandas, matplotlib, scipy

Input:
- Exercise_1b.xlsx (must be located in the same directory as the script)
"""
