import pandas as pd
import xlwings as xw
import os

# Prompt user for the file paths
excel_file = input("Enter the path to the Excel file (e.g., AMPERSAND_OCTOBER_MIS_Summary_2024.xlsx): ")

# Ensure the excel_file path is absolute
excel_file = os.path.abspath(excel_file)

# Check if the file exists, if not, try to find a similar file
if not os.path.exists(excel_file):
    # Get the directory of the script or current working directory
    script_dir = os.path.dirname(excel_file) or os.getcwd()
    
    # List all Excel files in the directory
    excel_files = [f for f in os.listdir(script_dir) if f.endswith('.xlsx')]
    
    if excel_files:
        # Find the most recently modified Excel file
        excel_files_with_path = [os.path.join(script_dir, f) for f in excel_files]
        excel_file = max(excel_files_with_path, key=os.path.getmtime)
        print(f"Using file: {excel_file}")
    else:
        print("No Excel file found!")
        exit()

rates_file = input("Enter the path to the rates CSV file (e.g., rates.csv): ")

# Generate output file path by appending '_updated' to the original file name
directory, original_filename = os.path.split(excel_file)
filename, extension = os.path.splitext(original_filename)
output_file = os.path.join(directory, f"{filename}_updated{extension}")

# Load rates CSV
rates_df = pd.read_csv(rates_file)

# Open the Excel workbook
app = xw.App(visible=False)  # Run Excel in the background
try:
    wb = app.books.open(excel_file)

    # Iterate over sheets and process each one
    for sheet in wb.sheets:
        sheet_name = sheet.name
        
        # Get M4 and R17 values
        m4_value = sheet.range("M4").value
        r17_value = sheet.range("R17").value
        
        # Filter rates for the specific cross_dock_name (sheet name matches)
        filtered_rates = rates_df[rates_df['cross_dock_name'] == sheet_name]
        
        # Find the appropriate rate for M4 and write it to N4
        m4_rate = filtered_rates[
            (filtered_rates['Lower_Bound'] <= m4_value) &
            (filtered_rates['Upper_Bound'] >= m4_value)
        ]['Rate'].values
        if m4_rate:
            sheet.range("N4").value = m4_rate[0]  # Get the first value as the rate
        else:
            sheet.range("N4").value = "No rate found"  # Handle missing rates gracefully

        # Find the appropriate rate for R17 and write it to S17
        r17_rate = filtered_rates[
            (filtered_rates['Lower_Bound'] <= r17_value) &
            (filtered_rates['Upper_Bound'] >= r17_value)
        ]['Rate'].values
        if r17_rate:
            sheet.range("S17").value = r17_rate[0]  # Get the first value as the rate
        else:
            sheet.range("S17").value = "No rate found"  # Handle missing rates gracefully

    # Save the workbook with the new name
    wb.save(output_file)
    wb.close()

except Exception as e:
    print(f"An error occurred: {e}")
finally:
    app.quit()

print(f"Rates successfully updated and saved in {output_file}.")