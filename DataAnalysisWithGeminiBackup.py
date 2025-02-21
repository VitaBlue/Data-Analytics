import pandas as pd
import google.generativeai as gemini
from datetime import datetime

# Step 1: Configure the Gemini API with your API key
gemini.configure(api_key="AIzaSyBsiclFvN2oV8rJphikFMIqmtWYh6zoizY")  # Replace with your actual API key

# Step 2: Load the XLSX file into a DataFrame
def load_xlsx(file_path):
    try:
        df = pd.read_excel(file_path)
        print("XLSX file loaded successfully!")
        return df
    except Exception as e:
        print(f"Error loading XLSX file: {e}")
        return None

def get_date_input():
    """
    Prompts the user to enter a date in YYYY-MM-DD format and validates the input.
    """
    while True:
        date_str = input("Enter the date (YYYY-MM-DD): ")
        try:
            date_object = datetime.strptime(date_str, '%Y-%m-%d')
            return date_object
        except ValueError:
            print("Invalid date format. Please use YYYY-MM-DD.")

def calculate_profit_by_date(file_path, start_date, end_date):
    """
    Calculates the sum of profit within a specified date range from an Excel file.

    Args:
        file_path (str): The path to the Excel file.
        start_date (datetime): The start date for filtering.
        end_date (datetime): The end date for filtering.
    """
    try:
        # Load the Excel file into a Pandas DataFrame
        df = pd.read_excel(file_path)

        # Ensure the DataFrame is loaded correctly
        if df is None or df.empty:
            print("Error: DataFrame is empty or could not be loaded.")
            return None

        # Print available columns for user reference
        print("Available columns:", df.columns.tolist())

        # Prompt the user for the index of the date column
        date_index = int(input("Enter the index of the date column (starting from 1): ")) - 1

        # Prompt the user for the index of the profit column
        profit_index = int(input("Enter the index of the profit column (starting from 1): ")) - 1

        # Verify that indices are within range
        if date_index < 0 or date_index >= len(df.columns) or profit_index < 0 or profit_index >= len(df.columns):
            print("Error: One or both specified indices are out of range.")
            return None
        
        # Select columns by index using iloc
        date_column = df.iloc[:, date_index]
        profit_column = df.iloc[:, profit_index]

        # Convert the date column to datetime objects
        df.iloc[:, date_index] = pd.to_datetime(date_column)

        # Filter the DataFrame based on the date range
        filtered_df = df[(df.iloc[:, date_index] >= start_date) & (df.iloc[:, date_index] <= end_date)]

        # Calculate the sum of profit for the filtered date range
        total_profit = filtered_df.iloc[:, profit_index].sum()

        # Print the total profit
        print(f"Total profit between {start_date.strftime('%Y-%m-%d')} and {end_date.strftime('%Y-%m-%d')}: {total_profit}")

        dates = "{} 至 {}".format(start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))

        return dates

    except FileNotFoundError:
        print("Error: File not found.")
        return None
    except ValueError:
        print("Error: Invalid input. Please enter numeric values for indices and valid dates.")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None
    
# Step 3: Define a function to ask questions based on the XLSX data
def answer_question_with_gemini(question, df):
    try:
        # Convert the DataFrame to a string for context
        context = df.to_string(index=False)
        
        # Generate a response using the Gemini model
        response = gemini.GenerativeModel('gemini-2.0-flash').generate_content(
            f"Answer the following question based on the data:\n\n{context}\n\nQuestion: {question}"
        )
        
        return response.text
    except Exception as e:
        print(f"Error generating response: {e}")
        return None

# Step 4: Main function to load XLSX and query Gemini
def main():
    # Prompt user for directory and file name
    directory = input("Enter the directory path where your XLSX file is located: ")
    file_name = input("Enter the name of your XLSX file (with .xlsx extension): ")
    
    # Combine directory and file name to create full path
    xlsx_file_path = f"{directory}/{file_name}"  # Adjust path separator if needed
    
    # Load the XLSX file
    df = load_xlsx(xlsx_file_path)
    
    if df is not None:
        # Get start and end dates from user input
        print("Enter the start date:")
        start_date = get_date_input()
        
        print("Enter the end date:")
        end_date = get_date_input()

        dates = calculate_profit_by_date(xlsx_file_path, start_date, end_date)

        if dates:
            question = "請為{}生成一份簡單的商業報告, 其中包括商品類別銷售分析、熱銷商品分析、銷售趨勢分析、問題與商業建議和商業決策。每部份用少於100字。格式:純文字 不要有開頭例如好的。".format(dates)
        
            # Get an answer from Gemini
            answer = answer_question_with_gemini(question, df)
        
            if answer:
                print(f"{answer}")
            else:
                print("Failed to get an answer from Gemini.")
        else:
            print("Failed to calculate profit by date.")
    else:
        print("Failed to load the XLSX file.")

# Run the script
if __name__ == "__main__":
    main()