import os
import pandas as pd
from google.colab import files

def merge_excel_files(folder_path):
    all_data = pd.DataFrame()

    for root, dirs, files in os.walk(folder_path):
        for i, file_name in enumerate(files):
            if file_name.endswith('.xls') or file_name.endswith('.xlsx'):
                file_path = os.path.join(root, file_name)
                df = pd.read_excel(file_path)

                # Remove the first row for files after the first one
                if i > 0:
                    df = df.iloc[1:]

                # Remove the last row
                df = df.iloc[:-1]

                # Add a new column for file name
                df['File_Name'] = file_name

                # Print columns with numeric values
                numeric_columns = df.select_dtypes(include=['number']).columns
                print(f"\nNumeric columns for {file_name}:\n{numeric_columns}")

                # Concatenate DataFrames
                all_data = pd.concat([all_data, df], ignore_index=True)

    return all_data

# Upload a zip file containing the folder with Excel files
from google.colab import files
uploaded = files.upload()

# Extract the contents of the uploaded zip file
import zipfile
import io

for zip_filename, content in uploaded.items():
    with zipfile.ZipFile(io.BytesIO(content), 'r') as zip_ref:
        zip_ref.extractall('/content/uploaded_folder/')

# Specify the folder path in Google Colab
folder_path = '/content/uploaded_folder/'

# Merge Excel files
all_data = merge_excel_files(folder_path)

# Display the first few rows of the merged DataFrame
print(all_data.head())

# Save merged DataFrame to a new Excel file
merged_file_path = '/content/merged_file.xlsx'
all_data.to_excel(merged_file_path, index=False)

# Download the Excel file
files.download(merged_file_path)
