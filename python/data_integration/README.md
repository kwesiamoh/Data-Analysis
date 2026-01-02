"""
Script: air_quality_data_integration_and_mapping.py

Purpose:
Integrate meteorological data and multiple air quality sensor datasets into a
time-aligned master table and generate spatial heatmaps of pollutant concentrations.

Description:
- Loads a meteorological master file and multiple auxiliary sensor datasets
- Filters data to a fixed time window
- Aligns records using seconds-from-midnight for exact timestamp matching
- Merges pollutant and meteorological variables into a single dataset
- Exports a consolidated Excel mastersheet
- Optionally generates interactive HTML heatmaps using GPS coordinates

Tools:
pandas, numpy, folium (optional)

Inputs:
- Meteorology.txt
- PM.xlsx, UFP.xlsx, Memocomp.xlsx, BC.csv

Outputs:
- Mastersheet.xlsx
- Interactive HTML heatmaps for selected pollutants
"""
