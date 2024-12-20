import random
from datetime import datetime
from tkinter.filedialog import askopenfilename
from pathlib import Path
import pandas as pd

def clean_and_upper(df):
    #This function takes a pandas dataframe, runs through each entry,
    # removes whitespace, and makes the text uppercase.

    return df.applymap(lambda x: x.strip().upper() if isinstance(x, str) else x)

def load_banned_list(banned_file_path):
    # Load the banned list as a DataFrame
    banned_data = pd.read_csv(banned_file_path, sep=',')
    return banned_data

def filter_list(original_data, banned_list,previous_attendee_list):

    #Makes sure case-insensitive and removes whitespace
    original_data = clean_and_upper(original_data)
    banned_list = clean_and_upper(banned_list)
    previous_attendee_list = clean_and_upper(previous_attendee_list)

    merged = original_data.merge(
        banned_list, on=['First name','Last name'],how='left',indicator=True
        )
    filtered_data = merged[merged['_merge'] == 'left_only'].drop(columns='_merge')

    merged = filtered_data.merge(
        previous_attendee_list, on=['First name','Last name'],how='left',indicator=True
        )
    filtered_data = merged[merged['_merge'] == 'left_only'].drop(columns='_merge')

    previous_attendee_list_rsvps = original_data.merge(
        previous_attendee_list,on=['First name','Last name'],how='inner'
        )

    return filtered_data, previous_attendee_list_rsvps


def load_previous_attendee_members(previous_attendee_path):
    # Load the already attended list as a DataFrame
    already_attended_data = pd.read_csv(previous_attendee_path, sep=',')
    already_attended_data = already_attended_data.dropna()
    return already_attended_data


def select_random_rows(filtered_data, number_of_attendees):

    selected_rows = pd.DataFrame()
    alternate_rows = pd.DataFrame()

    actual_number_of_attendees = number_of_attendees
    actual_num_alternates = len(filtered_data) - number_of_attendees

    # Check if num_rows is greater than the number of rows in the file
    if number_of_attendees > len(filtered_data):
        actual_number_of_attendees = len(filtered_data)
        actual_num_alternates = 0

    # Randomly select rows
    if actual_number_of_attendees > 0 :
        selected_rows = filtered_data.sample(
            n=actual_number_of_attendees, random_state=random.randint(0, 10000))

        # Exclude the selected rows to ensure alternates are distinct
        remaining_data = filtered_data.drop(selected_rows.index)

        if actual_num_alternates > 0:
            # Randomly select the alternate rows
            alternate_rows = remaining_data.sample(
                n=actual_num_alternates, random_state=random.randint(0, 10000))

    return selected_rows, alternate_rows

def save_to_excel(file_name_hunt,selected_rows, alternate_rows,randomized_prev_attendee_list):
    # Get current timestamp for unique file naming
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"selected_{timestamp}_" + file_name_hunt + ".xlsx"
    saved_file_path = "/Users/triciadang/Downloads/" + file_name
    # Create an Excel writer
    with pd.ExcelWriter(saved_file_path) as writer:
        # Write primary selected rows to the first sheet
        selected_rows.to_excel(writer, sheet_name="Primary Rows", index=False)

        # Write alternate rows to the second sheet
        alternate_rows.to_excel(writer, sheet_name="Alternate Rows", index=False)

        # Write previous attendees randomized to the third sheet
        randomized_prev_attendee_list.to_excel(writer, \
            sheet_name = "Randomized Previous Attendees", \
            index = False)
    print(f"Data saved to {file_name}")

# Randomize all of the attendees who have attended past events
def randomize_previous_attendee_list(previous_attendee_list):
    randomized_prev_attendee_list = previous_attendee_list.sample(frac=1).reset_index(drop=True)
    return randomized_prev_attendee_list


def read_csv_with_fallback_encoding(file_path):
    encodings=['utf-8', 'utf-16']
    for encoding in encodings:
        try:
            original_data = pd.read_csv(file_path, sep='\t',encoding=encoding)
            return original_data
        except UnicodeDecodeError:
            continue

    raise UnicodeDecodeError("Unable to decode file with any of the specified encodings.")


def main():

    option = 2

    if option == 1:
        # Get input file - Windows
        file_path = askopenfilename(
            filetypes=[("CSV Files", "*.csv"), ("TSV Files", "*.tsv")],
            title="Select the CSV/TSV File"
            )

    else:
        # Hardcoded file
        file_path = "/Users/triciadang/Downloads/\
        Guest list fho-axis-whitetail-veterans-hunt 2024-11-13"

    # Read the CSV file of event RSVPs and try different encodings
    original_data = read_csv_with_fallback_encoding(file_path)

    # Load in list of all banned members
    banned_list = load_banned_list("/Users/triciadang/Documents/Banned_Members.csv")

    # Load in list of all previous attendees
    previous_attendee_list = load_banned_list("/Users/triciadang/Documents/Already_Attended.csv")

    # Take out all banned members AND previous attendees from consideration
    filtered_data,previous_attendee_list_rsvps = \
        filter_list(original_data,banned_list,previous_attendee_list)

    # Get numbers you need
    number_of_attendees = input("Input number of attendees: ")

    # Randomly selects winners
    selected_data,alternate_rows = select_random_rows(filtered_data, int(number_of_attendees))

    randomized_prev_attendee_list = pd.DataFrame()

    #Randomize list of those who already attended
    randomized_prev_attendee_list = randomize_previous_attendee_list(previous_attendee_list_rsvps)

    # Write to an excel file
    if selected_data is not None and alternate_rows is not None:
        file_name_hunt = Path(file_path).name
        file_name_hunt = file_name_hunt.split(".")[0].split("Guest list ")[1]

        save_to_excel(file_name_hunt,selected_data,alternate_rows,randomized_prev_attendee_list)


if __name__ == '__main__':
    main()
