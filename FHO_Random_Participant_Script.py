import pandas as pd
import random
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from pathlib import Path


def select_random_rows(file_path, number_of_attendees,num_alternates):
    # Read the CSV file
    data = pd.read_csv(file_path, sep='\t', encoding='utf-16')
    
    # Check if num_rows is greater than the number of rows in the file
    if number_of_attendees+num_alternates > len(data):
        print("Requested number of rows exceeds total rows in the file.")
        return None
    
    # Randomly select rows
    selected_rows = data.sample(n=number_of_attendees, random_state=random.randint(0, 10000))

    # Exclude the selected rows to ensure alternates are distinct
    remaining_data = data.drop(selected_rows.index)
    
    # Randomly select the alternate rows
    alternate_rows = remaining_data.sample(n=num_alternates, random_state=random.randint(0, 10000))
    
    return selected_rows, alternate_rows

def save_to_excel(file_name_hunt,selected_rows, alternate_rows):
    # Get current timestamp for unique file naming
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = file_name_hunt + f"_selected_{timestamp}.xlsx"
    
    # Create an Excel writer

    with pd.ExcelWriter(file_name) as writer:
        # Write primary selected rows to the first sheet
        selected_rows.to_excel(writer, sheet_name="Primary Rows", index=False)
        
        # Write alternate rows to the second sheet
        alternate_rows.to_excel(writer, sheet_name="Alternate Rows", index=False)
    
    print(f"Data saved to {file_name}")

def main():
    file_path = askopenfilename(
        filetypes=[("CSV Files", "*.csv"), ("TSV Files", "*.tsv")],
        title="Select the CSV/TSV File"
    )

    print(file_path)

    number_of_attendees = input("Input number of attendees: ")
    number_of_alternates = input("Input number of alternates needed: ")


    selected_data,alternate_rows = select_random_rows(file_path, int(number_of_attendees),int(number_of_alternates))

    if selected_data is not None and alternate_rows is not None:
        file_name_hunt = Path(file_path).name
        file_name_hunt = file_name_hunt.split(".")[0]
        save_to_excel(file_name_hunt,selected_data,alternate_rows)


main()