import glob
import pandas as pd

# Get the list of .xlsx files
xlsx_files = glob.glob("janssons_kranar_*.xlsx")

# Initialize an empty list to store DataFrames
all_dataframes = []

# Loop through the list of files and read each one into a DataFrame
for file in xlsx_files:
    data = pd.read_excel(file)

    # Add the data from the current file to the list of DataFrames
    all_dataframes.append(data)

# Concatenate all the DataFrames in the list
merged_data = pd.concat(all_dataframes, ignore_index=True)

# Save the merged data to a new .xlsx file
merged_data.to_excel("merged_janssons_kranar.xlsx", index=False)
