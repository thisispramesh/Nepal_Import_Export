import pandas as pd

# Define the input and output file paths
input_file = r'D:\Pramesh\Power BI\Nepal Import Export\FTS_2080_81.xlsx'
output_file = r'D:\Pramesh\Power BI\Nepal Import Export\Import_Export.xlsx'

# Read the Excel file
xls = pd.ExcelFile(input_file)


# Function to remove the first two letters from sheet names
def remove_first_two_letters(sheet_name):
    return sheet_name[2:]


# Create a writer object to save the modified sheets
with pd.ExcelWriter(output_file) as writer:
    for sheet_name in xls.sheet_names:
        # Read each sheet into a DataFrame
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        # Extract the data starting from the 3rd row and treat it as the header
        new_header = df.iloc[2]  # Get the 3rd row as header
        df = df[3:]  # Take the data from the 4th row onward
        df.columns = new_header  # Set the new header

        # Remove the first two letters from the sheet name
        new_sheet_name = remove_first_two_letters(sheet_name)

        # Save the extracted table to a new Excel file
        df.to_excel(writer, sheet_name=new_sheet_name, index=False)

print(f'Modified Excel file saved as {output_file}')
