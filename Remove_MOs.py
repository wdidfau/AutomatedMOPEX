import pandas as pd
from datetime import datetime
import argparse
import os

def remove_resigned_officers(officer_rankings, removed_officers_path):
    removed_officers = pd.read_excel(removed_officers_path, sheet_name="Exclude from MOPEX")
    removed_Employee_ID = removed_officers['Employee ID'].tolist()
    officer_rankings = officer_rankings[~officer_rankings['Employee ID'].isin(removed_Employee_ID)]
    return officer_rankings

def posting_blacklist(mo_ranking_data, posting_blacklist_data):
    # Create a dictionary mapping each Employee ID to a set of their blacklisted PMS Codes.
    blacklist_dict = posting_blacklist_data.groupby('Employee ID')['PMS Code'].apply(set).to_dict()

    choice_columns = [col for col in mo_ranking_data.columns if "choice" in col]

    # Iterate over each officer in the main DataFrame.
    for index, officer_row in mo_ranking_data.iterrows():
        employee_id = officer_row['Employee ID']

        # Check if the current officer has any blacklist entries.
        if employee_id in blacklist_dict:
            pms_codes_to_blacklist = blacklist_dict[employee_id]

            # Iterate through each of the officer's choices.
            for col in choice_columns:
                # If a choice is in the set of blacklisted codes, modify it.
                if officer_row[col] in pms_codes_to_blacklist:
                    # Use .at for direct, fast, label-based modification.
                    mo_ranking_data.at[index, col] = "Blacklisted-" + str(officer_row[col])
    
    return mo_ranking_data

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--MORankFile', type=str, required=True, help="The file name of MO Ranking Consolidated")
    parser.add_argument('--RemovedMOFile', type=str, required=True, help="The file name of Removed MOs") 
    args = parser.parse_args()
    
    officer_rankings_path = args.MORankFile
    removed_officers_path = args.RemovedMOFile 

    if os.path.exists(officer_rankings_path):
        officer_rankings = pd.read_excel(officer_rankings_path)
    else:
        print(f"No file found at: {officer_rankings_path}")
        exit()

    # Removing resigned officers
    officer_rankings = remove_resigned_officers(officer_rankings, removed_officers_path)
    
    # Applying posting blacklist
    posting_blacklist_data = pd.read_excel(removed_officers_path, sheet_name="Posting Blacklist")
    updated_mo_ranking_data = posting_blacklist(officer_rankings, posting_blacklist_data)

    # Extract directory from MORankFile path
    output_dir = os.path.dirname(officer_rankings_path) 

    # Save the updated rankings to a new file in the same directory
    date_string = datetime.today().strftime('%d-%m-%Y')
    output_filename = os.path.join(output_dir, f'Updated MO Ranking Consolidated {date_string}.xlsx')
    updated_mo_ranking_data.to_excel(output_filename, index=False) 
