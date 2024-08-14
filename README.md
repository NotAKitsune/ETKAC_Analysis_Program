# ETKAC Analysis Program

## Description
This program is designed to analyze ETKAC (Erythrocyte Transketolase Activity Coefficient) experiment data. It processes export files from ETKAC experiments and layout files to generate comprehensive analysis reports.

## Features
- User-friendly GUI for easy file selection and operation
- Processes ETKAC export files and layout files
- Generates detailed analysis reports in Excel format
- Calculates and displays statistics including averages, standard deviations, and coefficients of variation
- Color-coded results for easy interpretation
- Supports multiple patient data analysis in a single run

## Requirements
- Python 3.x
- Required Python libraries: pandas, openpyxl, numpy, tkinter

## Installation
1. Ensure Python 3.x is installed on your system.
2. Install required libraries using pip:
   i. pip install pandas openpyxl numpy
3. Download the ETKAC Analysis program script.

## Usage
1. Run the script to launch the GUI.
2. Enter the technician's name.
3. Select the ETKAC export file.
4. Select the layout file.
5. Click "Create Analysis Report" to generate and save the analysis.

## Output
The program generates an Excel file containing:
- A summary sheet with overall results
- Detailed ETKAC analysis for each patient
- Basal activity calculations

## Notes
- Ensure that the export and layout files are in the correct format.
- The program will flag any invalid blank readings.
- Results are categorized as Sufficient, Insufficient, or Deficient based on ETKAC values.
