import os
import openpyxl
import re
from opencc import OpenCC

def get_input_file(default_directory):
    """Prompt user to input the name of the Excel file."""
    while True:
        input_file = input("Enter the name of the input Excel file (with .xlsx extension): ")
        full_path = os.path.join(default_directory, input_file)
        try:
            wb = openpyxl.load_workbook(full_path)
            return full_path
        except FileNotFoundError:
            print(f"File '{full_path}' not found. Please try again.")

def get_output_file_details(default_directory):
    """Prompt user for output file name and directory."""
    output_directory = input(f"Enter the output directory (or '0' to use '{default_directory}'): ")
    
    if output_directory == '0':
        output_directory = default_directory
    
    output_file_name = input("Enter the output file name (without extension): ")
    
    full_output_path = os.path.join(output_directory, f"{output_file_name}.xlsx")
    
    return full_output_path

def convert_simplified_to_traditional(input_file, output_file):
    """Convert Simplified Chinese text in an Excel file to Traditional Chinese."""
    cc = OpenCC('s2t')
    wb = openpyxl.load_workbook(input_file)
    
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cell.value = cc.convert(cell.value)
    
    wb.save(output_file)
    print("Chinese conversion completed.")
    return output_file

def get_column_headers(file_path):
    """Get the headers (first row) of the Excel file."""
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    headers = []
    for cell in ws[1]:
        headers.append(cell.value if cell.value is not None else "")
    return headers

def process_text_column(ws, col):
    """Process text column - Remove spaces"""
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if isinstance(cell.value, str):
            cell.value = ''.join(cell.value.split())

def process_date_column(ws, col):
    """Process date column - Remove spaces and format date"""
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if isinstance(cell.value, str):
            # First remove spaces
            cleaned_value = ''.join(cell.value.split())
            
            # Then format date
            date_patterns = [
                r'(\d{4})-(\d{1,2})-(\d{1,2})',
                r'(\d{4})年(\d{1,2})月(\d{1,2})日',
                r'(\d{4})(\d{2})(\d{2})'
            ]
            
            formatted_date = None
            for pattern in date_patterns:
                match = re.match(pattern, cleaned_value)
                if match:
                    year = match.group(1)
                    month = str(int(match.group(2)))
                    day = str(int(match.group(3)))
                    formatted_date = f"{year}年{month}月{day}日"
                    break
            
            if formatted_date:
                cell.value = formatted_date

def process_time_column(ws, col):
    """Process time column - Format time or replace with N/A"""
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        
        # Initialize cleaned_value
        cleaned_value = ""
        
        # Check if cell value is string
        if isinstance(cell.value, str):
            # Try to match time patterns first
            time_str = cell.value.strip()
            # Pattern for H:MM or HH:MM
            match = re.match(r'^(\d{1,2}):(\d{2})$', time_str)
            if match:
                hours = int(match.group(1))
                minutes = int(match.group(2))
                if 0 <= hours <= 23 and 0 <= minutes <= 59:
                    cell.value = f"{hours:02d}:{minutes:02d}"
                    continue
            
            # If no match, try to extract numbers
            cleaned_value = ''.join(filter(str.isdigit, time_str))
        
        # Format time if we have valid numeric data
        if cleaned_value:
            if len(cleaned_value) >= 4:
                hours = int(cleaned_value[:2])
                minutes = int(cleaned_value[2:4])
                if 0 <= hours <= 23 and 0 <= minutes <= 59:
                    cell.value = f"{hours:02d}:{minutes:02d}"
                else:
                    cell.value = "N/A"
            elif len(cleaned_value) == 3:  # Handle cases like "8:18" -> "818"
                hours = int(cleaned_value[0])
                minutes = int(cleaned_value[1:])
                if 0 <= hours <= 23 and 0 <= minutes <= 59:
                    cell.value = f"{hours:02d}:{minutes:02d}"
                else:
                    cell.value = "N/A"
            else:
                cell.value = "N/A"
        else:
            cell.value = "N/A"

def process_number_column(ws, col):
    """Process number column - Remove text and spaces"""
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if isinstance(cell.value, str):
            # Keep only numeric characters
            cleaned_value = ''.join(filter(str.isdigit, cell.value))
            cell.value = int(cleaned_value) if cleaned_value else None

def main():
    default_directory = os.getcwd()
    
    # Step 1: Get input file and convert Chinese characters
    input_file = get_input_file(default_directory)
    temp_output = os.path.join(default_directory, "temp_converted.xlsx")
    converted_file = convert_simplified_to_traditional(input_file, temp_output)
    
    # Step 2: Get column headers and process based on type
    headers = get_column_headers(converted_file)
    print("\nColumn headers:")
    for i, header in enumerate(headers, 1):
        print(f"{i}. {header}")
    
    # Create new workbook for final output
    wb = openpyxl.load_workbook(converted_file)
    ws = wb.active
    
    # Process each column
    for col_num in range(1, len(headers) + 1):
        while True:
            print(f"\nFor column {col_num} ({headers[col_num-1]}), choose data type:")
            print("1. Text")
            print("2. Date")
            print("3. Time")
            print("4. Number")
            print("5. None")
            choice = input("Enter choice (1-5): ")
            
            if choice in ['1', '2', '3', '4', '5']:
                if choice == '1':  # Text
                    process_text_column(ws, col_num)
                elif choice == '2':  # Date
                    #process_text_column(ws, col_num)
                    process_date_column(ws, col_num)
                elif choice == '3':  # Time
                    process_time_column(ws, col_num)
                elif choice == '4':  # Number
                    process_number_column(ws, col_num)
                    process_text_column(ws, col_num)
                # Choice 5 (None) does nothing
                break
            else:
                print("Invalid choice. Please enter a number between 1 and 5.")
    
    # Save final output
    final_output = get_output_file_details(default_directory)
    wb.save(final_output)
    
    # Clean up temporary file
    try:
        os.remove(temp_output)
    except:
        pass
    
    print(f"\nProcessing complete. Final output saved to: {final_output}")

if __name__ == "__main__":
    main()