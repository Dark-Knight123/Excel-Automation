import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import numbers
from openpyxl.drawing.image import Image
import os

# print the current working directory
print("Current Working Directory (Before):", os.getcwd())

# dynamically set the working directory to the folder containing the script
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)
print("Updated Working Directory (After):", os.getcwd())

# print the full paths of the files to verify they exist
file_name = "Book.xlsx"
image_path = r"R.jpg"  # assuming the image is in the same directory as the script
print("Excel File Path:", os.path.abspath(file_name))
print("Image File Path:", os.path.abspath(image_path))

# check if the files exist
if not os.path.exists(file_name):
    print(f"Error: Excel file not found at {os.path.abspath(file_name)}")
if not os.path.exists(image_path):
    print(f"Error: Image file not found at {os.path.abspath(image_path)}")

# read the Excel file into a DataFrame
# if you want to read a specific sheet, you can specify the sheet name using the 'sheet_name' parameter (e.g., sheet_name="Sheet1")
df = pd.read_excel(file_name)

# extract unique client names from the "Source" column
# filter out entries where the "Source" value has only one word (e.g., "Unknown")
# this assumes that valid client names have more than one word
names = list(filter(lambda x: len(str(x).split()) > 1, df.groupby("Source").count().reset_index()["Source"]))

# initialize an empty list to store the starting indices of each client's data
name_indices = []

# iterate through the list of client names
for name in names:
    # find the index of the first occurrence of the client name in the "Source" column
    # np.nonzero() returns the indices where the condition is True
    # np.array(df["Source"]) converts the "Source" column to a NumPy array for comparison
    name_indices.append(np.nonzero(np.array(df["Source"]) == name)[0][0])

# sort the indices in ascending order and add the length of the DataFrame as the last index
# adding len(df) ensures that the last client's data is included in the slicing
name_indices = np.array(sorted(name_indices) + [len(df)])

# displays first 5 rows of the DataFrame
df.head()

# read only the "Reference" column from the Excel file "Book.xlsx"
# the 'usecols' parameter ensures that only the specified column is read into the DataFrame
df_references = pd.read_excel(file_name, usecols=["Reference"])

# extract the first entry (value) from the "Reference" column
# 'iloc[0, 0]' accesses the first row and first column of the DataFrame
opening_balance = df_references.iloc[0, 0]

# print the extracted opening balance to the console
print("Opening Balance:", opening_balance)

# read the Excel file specified by the variable 'file_name' into a DataFrame
# the DataFrame 'df' will contain the data from the Excel file, including all columns and rows
df = pd.read_excel(file_name)

# print the column names of the DataFrame to the console
# this helps verify the structure of the DataFrame and ensures the expected columns are present
print("Column names:", df.columns)

# iterate through each client using the indices in name_indices
for i in range(len(name_indices) - 1):
    # extract the current client's data from the DataFrame using slicing
    # the range is determined by consecutive indices in name_indices
    current_df = df.iloc[name_indices[i]: name_indices[i + 1]]
    
    # extract the client name from the first row of the "Source" column
    name = current_df["Source"].iloc[0]
    
    # drop the "Account Number/Year/ Prd." column if it exists
    # the 'errors="ignore"' parameter ensures no error is raised if the column is missing
    current_df = current_df.drop(labels=["Account Number/Year/ Prd."], axis=1)
    
    # rename the "Credits" column to "Deposits" and the "Debits" column to "Withdrawals"
    current_df = current_df.rename(columns={"Credits": "Deposits", "Debits": "Withdrawals"})
    
    # add a new "Balance" column with empty values for all rows
    values = [""] * len(current_df)
    current_df["Balance"] = values
    
    # remove rows with null values from the DataFrame
    current_df = current_df.dropna()
    
    # format the "Doc. Date" column to a standard date format (MM/DD/YYYY)
    # the 'errors="coerce"' parameter ensures invalid dates are converted to NaT
    current_df['Doc. Date'] = pd.to_datetime(current_df['Doc. Date']).dt.strftime('%m/%d/%Y')
    
    # replace zero entries in the DataFrame with empty strings
    current_df.replace(0, "", inplace=True)
    
    # ensure the "Withdrawals" and "Deposits" columns are numeric
    # invalid values are coerced to NaN
    current_df["Withdrawals"] = pd.to_numeric(current_df["Withdrawals"], errors='coerce')
    current_df["Deposits"] = pd.to_numeric(current_df["Deposits"], errors='coerce')
    
    # calculate the sum of the "Withdrawals" and "Deposits" columns
    withdrawals_sum = current_df["Withdrawals"].sum()
    deposits_sum = current_df["Deposits"].sum()
    
    # perform the desired mathematical operation: subtract withdrawals from deposits
    total_value = deposits_sum - withdrawals_sum
    
    # read only the "Reference" column from the Excel file "Book.xlsx"
    # the 'usecols' parameter ensures only the specified column is read
    df_references = pd.read_excel(file_name, usecols=["Reference"])
    
    # extract the opening balance for the current client using the index from name_indices
    opening_balance = df_references.iloc[name_indices[i], 0]
    print(opening_balance)  # Print the opening balance for debugging purposes
    
    # calculate the ending balance by adding the total value to the opening balance
    ending_balance = opening_balance + total_value
    
    # create a DataFrame for the opening balance row
    new_first = pd.DataFrame(data=[["", "", "", "Opening Balance:", "", "", opening_balance]], columns=current_df.columns)
    
    # create a DataFrame for the ending balance row
    new_end = pd.DataFrame(data=[["", "", "", "Ending Balance:", "", "", ending_balance]], columns=current_df.columns)
    
    # create a DataFrame for additional rows with the client name in the first row and an empty second row
    additional_rows = pd.DataFrame(data=[[name] + [""] * (len(current_df.columns) - 1), [""] * len(current_df.columns)], columns=current_df.columns)
    
    # concatenate the additional rows, opening balance row, current_df, and ending balance row
    final_df = pd.concat([additional_rows, new_first, current_df, new_end]).reset_index(drop=True)
    
    # save the modified DataFrame to a new Excel file named after the client
    client_file_name = f"{name}.xlsx"  # Correct file name for the current client
    final_df.to_excel(client_file_name, index=False)
    
    # apply Excel-specific number formatting using openpyxl
    workbook = load_workbook(client_file_name)  # Load the correct file for the current client
    sheet = workbook.active

    # apply dollar formatting to the "Withdrawals," "Deposits," and "Balance" columns
    for col in ["Withdrawals", "Deposits", "Balance"]:
        if col in final_df.columns:
            col_index = final_df.columns.get_loc(col) + 1  # Get the 1-based column index for openpyxl
            for row in range(2, sheet.max_row + 1):  # Start from row 2 to skip the header
                cell = sheet.cell(row=row, column=col_index)
                if isinstance(cell.value, (int, float)):  # Ensure the value is numeric
                    cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE  # Apply dollar formatting

    # save the workbook with the applied formatting
    workbook.save(client_file_name)

    # insert an empty row and add the image
    # reload the workbook to ensure changes are applied
    workbook = load_workbook(client_file_name)
    sheet = workbook.active

    # insert 8 empty rows at the top of the worksheet (shift all rows down)
    for _ in range(8):  # loop 8 times
        sheet.insert_rows(1)

    # add the image to the worksheet 
    # r gets rid of type casting, use 'ls' in windows terminal to find the path of image
    # image_path = r"\Users\frank_\OneDrive - University of Windsor\Desktop\github\autoExcel-1\R.jpg"  
    # img = Image(image_path)
    # img.anchor = "A1"  # place the image in cell A1 (top-left corner)
    
    image_path = r"R.jpg"  
    img = Image(image_path)
    img.anchor = "A1"  # place the image in cell A1 (top-left corner)

    # resize the image (optional)
    img.width = 200  # adjust width
    img.height = 150  # adjust height

    # add the image to the worksheet
    sheet.add_image(img)

    # adjust page layout settings for PDF conversion
    sheet.page_setup.fitToPage = True # enable "Fit to Page" mode
    sheet.page_setup.fitToWidth = 1  # fit all columns to one page width
    sheet.page_setup.fitToHeight = 0  # allow rows to span multiple pages
    sheet.page_setup.orientation = "landscape"  # landscape orientation for wide tables
    sheet.page_setup.paperSize = 9  # use 9 for A4 paper size (constant value in openpyxl)
    sheet.print_area = f"A1:{sheet.cell(sheet.max_row, sheet.max_column).coordinate}"  # set the print area to include all data

    # save the updated workbook with the image
    workbook.save(client_file_name)

sheet.print_area = f"A1:{sheet.cell(sheet.max_row, sheet.max_column).coordinate}"

print(sheet.print_area)


