import os
import openpyxl
import re

def change_working_directory():
    """Change the current working directory."""
    os.chdir(r'C:\Users\user\OneDrive\文件\Python')

def get_input_file():
    """Prompt user to input the name of the Excel file."""
    while True:
        input_file = input("Enter the name of the input Excel file (with .xlsx extension): ")
        try:
            wb = openpyxl.load_workbook(input_file)  # Attempt to load the workbook
            return input_file
        except FileNotFoundError:
            print(f"File '{input_file}' not found. Please try again.")

def get_columns_to_clean():
    """Prompt user to input the column indices to clean."""
    while True:
        try:
            columns_input = input("Enter the column indices to clean (comma-separated, e.g., 1,2 for columns A and B): ")
            columns = [int(col.strip()) for col in columns_input.split(',')]  # Return as is (1-based index)
            if any(col < 1 for col in columns):
                print("Column index must be at least 1.")
                continue
            return columns  # Return list of column indices
        except ValueError:
            print("Invalid input. Please enter valid integers.")

def get_output_file_details(default_directory):
    """Prompt user for output file name and directory."""
    output_directory = input(f"Enter the output directory (or '0' to use '{default_directory}'): ")
    
    if output_directory == '0':
        output_directory = default_directory
    
    output_file_name = input("Enter the output file name (without extension): ")
    
    full_output_path = os.path.join(output_directory, f"{output_file_name}.xlsx")
    
    return full_output_path

def format_date(date_string):
    """Format date strings to 'XXXX年XX月XX日'."""
    cleaned_string = ''.join(date_string.split())
    
    # Match date format YYYY-MM-DD
    match = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})', cleaned_string)
    if match:
        year = match.group(1)
        month = int(match.group(2))
        day = int(match.group(3))
        return f"{year}年{month}月{day}日"
    
    # If date format does not match, return None
    return None

def clean_dates(input_file, column_indices, output_file):
    """Clean and format dates in specified columns of an Excel file."""
    wb = openpyxl.load_workbook(input_file)
    ws = wb.active

    for col in column_indices:
        for row in range(1, ws.max_row + 1):  # Start from row 1 (including header)
            cell = ws.cell(row=row, column=col)
            if isinstance(cell.value, str):  # Check if the cell contains text
                formatted_date = format_date(cell.value.strip())
                if formatted_date:
                    cell.value = formatted_date  # Update cell value with formatted date

    # Save the changes to a new Excel file
    wb.save(output_file)
    print(f"All dates have been cleaned and saved to {output_file}")

# Main execution flow
change_working_directory()  # Change working directory at start

# Get user input for file name and columns to clean
input_excel_file = get_input_file()
columns_to_clean = get_columns_to_clean()

# Get default directory where the script is running
default_directory = os.getcwd()

# Get output file details from user
output_excel_file = get_output_file_details(default_directory)

# Clean dates in specified columns and save to output file
clean_dates(input_excel_file, columns_to_clean, output_excel_file)
