import pandas as pd
from tkinter import Tk, filedialog
import os

# Function to load an Excel file
def load_excel():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    if file_path:
        return file_path
    else:
        return None

# Function to process the Excel file
def process_excel(file_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Check for required columns and notify if any are missing
    required_columns = ['Checked on', 'Closed on', 'Reported on', 'Status']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print(f"Missing required columns: {', '.join(missing_columns)}")
        return

    # Create new columns based on conditions
    df['Closed on Check Day?'] = df.apply(lambda row: 'Yes' if row['Checked on'] == row['Closed on'] else 'No', axis=1)
    df['Reported on Day?'] = df.apply(lambda row: 'Yes' if row['Reported on'] == row['Checked on'] else 'No', axis=1)
    df['Open?'] = df.apply(lambda row: 'Yes' if row['Status'] in ['In Bearbeitung', 'Offen'] else 'No', axis=1)

   # Count 'Yes' in each new column and write it only in the first row
    df['Count Closed on Check Day'] = None  # Initialize the column with None
    df['Count Reported on Day'] = None      # Initialize the column with None
    df['Count Open'] = None                 # Initialize the column with None

    # Assign the counts only to the first row
    df.at[0, 'Count Closed on Check Day'] = (df['Closed on Check Day?'] == 'Yes').sum()
    df.at[0, 'Count Reported on Day'] = (df['Reported on Day?'] == 'Yes').sum()
    df.at[0, 'Count Open'] = (df['Open?'] == 'Yes').sum()


    # Save the new Excel file
    new_file_path = os.path.splitext(file_path)[0] + '_Reporting.xlsx'
    df.to_excel(new_file_path, index=False)
    print(f'Processed file saved as: {new_file_path}')

# Main function to execute the program
def main():
    print("Please select an Excel file to process...")
    file_path = load_excel()
    if file_path:
        process_excel(file_path)
    else:
        print("No file selected. Exiting.")

if __name__ == "__main__":
    main()
