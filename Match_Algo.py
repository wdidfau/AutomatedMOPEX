import pandas as pd
import argparse
import os
import numpy as np

def gale_shapley_1(hod_rankings, officer_choices, no_vacancies_departments):
    # Prepare the officers list and departments list
    officers_list = officer_choices['Employee ID'].tolist()
    departments_list = hod_rankings['PMS Code'].tolist()

    # Prepare departments vacancies
    department_vacancies = hod_rankings.set_index('PMS Code')['Vacancies'].to_dict()

    # Prepare officers and departments preferences
    # Replace empty strings with NaN in the preference columns
    preference_columns = [col for col in officer_choices.columns if 'choice' in col.lower()]  # Assuming 'choice' in officers data
    officer_choices[preference_columns] = officer_choices[preference_columns].replace('', np.nan)

    # Create dictionary of Officer Preferences
    officers_pref = officer_choices.set_index('Employee ID')[preference_columns].to_dict('split')
    officers_pref = {officer: choices for officer, choices in zip(officers_pref['index'], officers_pref['data'])}

    # Create dictionary of Department Preferences
    department_pref_columns = [col for col in hod_rankings.columns if 'Match' in col]
    departments_pref = hod_rankings.set_index('PMS Code')[department_pref_columns].replace(np.nan, '').T.to_dict('list')

    # Create dictionaries to map MCR Number to Employee Name and PMS Code to Posting, add GDFM Bump comments
    officer_names = officer_choices.set_index('Employee ID')['Employee Name'].to_dict()
    department_postings = hod_rankings.set_index('PMS Code')['Postings'].to_dict()
    officer_comments = pd.Series(officer_choices.Comment.values, index=officer_choices['Employee ID']).to_dict()

    # Initialize
    tentative_matches = {department: [] for department in departments_list}
    first_unmatched_officers = []
    
    exceptions = []  # To track the exceptions
    
    while officers_list:
        officer = officers_list.pop(0)
        officer_preferences = officers_pref[officer]
        
        for department in officer_preferences:
            if pd.isna(department):  # This exception handles officers who have listed <10 choices
                first_unmatched_officers.append(officer)
                break
            try:
                department_preferences = departments_pref[department]
                if officer not in department_preferences:  # If officer not in department preferences, reject and continue to next preference
                    continue
                if department_vacancies[department] > len(tentative_matches[department]):
                    # There's an available vacancy
                    tentative_matches[department].append(officer)
                    break
                else:
                    # No vacancy available, check if officer is a preferred candidate
                    current_worst_officer = max(tentative_matches[department], key=department_preferences.index)
                    
                    if department_preferences.index(officer) < department_preferences.index(current_worst_officer):
                        # Officer is a preferred candidate
                        tentative_matches[department].remove(current_worst_officer)
                        tentative_matches[department].append(officer)
                        officers_list.append(current_worst_officer)
                        break

            #Debugs for unexpected PMS Code i.e. when MO applies for department not available in HOD Ranking List
            except KeyError:
                if department not in no_vacancies_departments['PMS Code'].values: #Checks if PMS Code was already flagged up previously as having no vacancy and thus removed
                    exceptions.append((officer, department))
                    
        else:
            # Officer could not be matched, add to unmatched list
            first_unmatched_officers.append(officer)
    
    # Returns dataframes of unmatched officers, departments with vacancies, and matches to main function for writing
    first_unmatched_officers_df = officer_choices[officer_choices['Employee ID'].isin(first_unmatched_officers)]
    departments_with_vacancies = {dept: vacancies - len(tentative_matches[dept]) for dept, vacancies in department_vacancies.items() if vacancies - len(tentative_matches[dept]) > 0}
    first_departments_with_vacancies_df = pd.DataFrame.from_dict(departments_with_vacancies, orient='index', columns=['Remaining Vacancies'])
    first_departments_with_vacancies_df.reset_index(inplace=True)
    first_departments_with_vacancies_df.columns = ['Department', 'Remaining Vacancies']
    mutualmatch_df = pd.DataFrame([(officer, officer_names[officer], department, department_postings[department], officer_comments[officer]) for department, officers in tentative_matches.items() for officer in officers], columns=['Employee ID', 'Employee Name', 'PMS Code', 'Posting', 'Comment'])
    return mutualmatch_df, first_unmatched_officers_df, first_departments_with_vacancies_df, exceptions

def gale_shapley_2(first_unmatched_officers_df, first_departments_with_vacancies_df, hod_rankings):
    officers_list = first_unmatched_officers_df['Employee ID'].tolist()
    departments_list = first_departments_with_vacancies_df['Department'].tolist()

    # Prepare departments vacancies
    department_vacancies = first_departments_with_vacancies_df.set_index('Department')['Remaining Vacancies'].to_dict()

    # Prepare officers and departments preferences
    # Replace empty strings with NaN in the preference columns
    preference_columns = [col for col in first_unmatched_officers_df.columns if 'choice' in col.lower()]  # Assuming 'choice' in officers data
    first_unmatched_officers_df.loc[:, preference_columns] = first_unmatched_officers_df.loc[:, preference_columns].replace('', np.nan)

    # Create dictionary of Officer Preferences
    officers_pref = first_unmatched_officers_df.set_index('Employee ID')[preference_columns].to_dict('split')
    officers_pref = {officer: choices for officer, choices in zip(officers_pref['index'], officers_pref['data'])}

    # Create dictionaries to map MCR Number to Employee Name and PMS Code to Posting, add GDFM Bump comments
    officer_names = first_unmatched_officers_df.set_index('Employee ID')['Employee Name'].to_dict()
    department_postings = hod_rankings.set_index('PMS Code')['Postings'].to_dict()
    officer_comments = pd.Series(first_unmatched_officers_df.Comment.values, index=first_unmatched_officers_df['Employee ID']).to_dict()

    # Initialize
    tentative_matches = {department: [] for department in departments_list}
    second_unmatched_officers = []

    while officers_list:
        officer = officers_list.pop(0)
        officer_preferences = officers_pref[officer]

        for rank, department in enumerate(officer_preferences):
            if pd.isna(department):  # This exception handles officers who have listed <10 choices
                second_unmatched_officers.append(officer)
                break

            # If department is full, reject and move on
            if department not in department_vacancies:
                continue

            #If department has vacancy, add MO in
            if department_vacancies[department] > len(tentative_matches[department]):
                tentative_matches[department].append((officer, rank))
                break
            else:
                # No vacancy available, rank the existing MOs and the new MO
                ranked_officers = sorted(tentative_matches[department] + [(officer, rank)], key=lambda x: (x[1], x[0])) #sorts based on MO preference, then by employee ID
                if ranked_officers.index((officer, rank)) < department_vacancies[department]:
                    # MO is accepted, remove the last MO
                    rejected_officer = ranked_officers.pop(-1)[0]
                    tentative_matches[department] = ranked_officers
                    officers_list.append(rejected_officer)
                    break

        else:
            # Officer could not be matched, add to unmatched list
            second_unmatched_officers.append(officer)

    tentative_matches = {department: [officer for officer, _ in officers] for department, officers in tentative_matches.items()}

    # Returns dataframes of unmatched officers, departments with vacancies, and matches to main function for writing
    second_unmatched_officers_df = first_unmatched_officers_df[first_unmatched_officers_df['Employee ID'].isin(second_unmatched_officers)]
    departments_with_vacancies = {dept: vacancies - len(tentative_matches[dept]) for dept, vacancies in department_vacancies.items() if vacancies - len(tentative_matches[dept]) > 0}
    second_departments_with_vacancies_df = pd.DataFrame.from_dict(departments_with_vacancies, orient='index', columns=['Remaining Vacancies'])
    second_match_df = pd.DataFrame([(officer, officer_names[officer], department, department_postings[department], officer_comments[officer]) for department, officers in tentative_matches.items() for officer in officers], columns=['Employee ID', 'Employee Name', 'PMS Code', 'Posting', 'Comment'])

    return second_match_df, second_unmatched_officers_df, second_departments_with_vacancies_df

def license_satisfies_requirement(officer_license, requirement):
    license_hierarchy = ["Conditional-L1", "Conditional-L2", "Conditional-L3", "Full"]
    return license_hierarchy.index(officer_license) >= license_hierarchy.index(requirement)

def check_license(officer_choices, posting_license_requirement):
    # Iterate over each officer's choices
    for idx, officer in officer_choices.iterrows():
        for col in officer_choices.columns:
            if 'choice' in col.lower():  # Assuming 'choice' in officers data
                department = officer[col]
                if pd.isna(department):
                    continue
                # Check if department has license requirement
                if department in posting_license_requirement['PMS Code'].values:
                    requirement = posting_license_requirement[posting_license_requirement['PMS Code'] == department]['Requirement'].values[0]
                    # Check if officer meets the requirement
                    if pd.notna(officer['Registration Type']):
                        if not license_satisfies_requirement(officer['Registration Type'], requirement):
                            # Officer does not meet the requirement, modify the department code
                            officer_choices.at[idx, col] = department + '_' + officer['Registration Type']
    return officer_choices

def GDFM_bump(GDFM1, GDFM2, officer_choices, hod_rankings):

    officer_choices['Comment'] = ''
    polyclinic_postings = ['NHGPlyNHGPly', 'SHSPlySHSPly', 'NUPNUP']

    # Initializing priority lists and counters for each polyclinic cluster
    priority_lists = {posting: [] for posting in polyclinic_postings}
    priority_counters = {posting: 0 for posting in polyclinic_postings}
    original_positions = {}  # Dictionary to store original positions

    def prioritize_candidates(GDFM, priority_lists, priority_counters, original_positions, posting):

        for _, row in GDFM.iterrows():
            if row['Eligible for Prioritisation'] == 'Y' and pd.notna(row['Employee ID']):
                employee_id = row['Employee ID']
                try:
                    first_choice = officer_choices.loc[officer_choices['Employee ID'] == employee_id, '1st choice'].values[0]
                except IndexError:
                    print(f'Error: Officer {employee_id} was not found in officer_choices. Please check the data.')
                    continue

                if first_choice == posting:
                    hod_row = hod_rankings.loc[hod_rankings['PMS Code'] == first_choice]
                    match_cols = [col for col in hod_rankings.columns if 'Match' in col]

                    if employee_id in hod_row[match_cols].values:
                        # Store original position
                        original_position = hod_row[match_cols].apply(lambda row: next((i for i, col in enumerate(row) if str(col) == str(employee_id)), None), axis=1).values[0]
                        original_positions[employee_id] = original_position

                        priority_lists[first_choice].append(employee_id)
                        priority_counters[first_choice] += 1

                        # Remove the prioritized officer from their current position
                        hod_rankings.loc[hod_row.index, match_cols] = hod_row[match_cols].apply(lambda x: x.replace(str(employee_id), ''))
        return priority_counters

    # Check for postings with 0 vacancies
    postings_with_vacancies = []
    for posting in polyclinic_postings:
        try:
            if hod_rankings.loc[hod_rankings['PMS Code'] == posting, 'Vacancies'].values[0] > 0:
                postings_with_vacancies.append(posting)
            else:
                print(f"Skipping prioritization for {posting} as there are no vacancies.")
        except IndexError:
            print(f"Could not find '{posting}' in HOD Rankings file!")
            continue

    # Prioritize candidates from GDFM cohort 1 only for postings with vacancies
    for posting in postings_with_vacancies:
        priority_counters = prioritize_candidates(GDFM1, priority_lists, priority_counters, original_positions, posting)

    #Prioritize candidates from 2nd GDFM cohort if there are still vacancies
    for posting in postings_with_vacancies:
        
        if priority_counters[posting] < hod_rankings.loc[hod_rankings['PMS Code'] == posting, 'Vacancies'].values[0]:
           priority_counters = prioritize_candidates(GDFM2, priority_lists, priority_counters, original_positions, posting)

    # Rearrange the remaining officers and insert the prioritized officers
    for posting, priority_list in priority_lists.items():
        if priority_list:
            hod_row = hod_rankings.loc[hod_rankings['PMS Code'] == posting]
            match_cols = [col for col in hod_rankings.columns if 'Match' in col]
            remaining_officers = pd.Series(hod_row[match_cols].values.flatten()).replace('', pd.NA).dropna().tolist()

            # Separate priority list into sheet1 and sheet2 lists
            GDFM1_priority = [officer for officer in priority_list if officer in GDFM1['Employee ID'].values]
            GDFM2_priority = [officer for officer in priority_list if officer in GDFM2['Employee ID'].values]

            # Sort each priority list based on original HOD ranking
            def sort_by_hod_ranking(priority_list):
                return [x for _, x in sorted(zip(map(original_positions.get, priority_list), priority_list))]

            GDFM1_priority = sort_by_hod_ranking(GDFM1_priority)
            GDFM2_priority = sort_by_hod_ranking(GDFM2_priority)

            # Combine the sorted priority lists
            final_priority_list = GDFM1_priority + GDFM2_priority

            # Ensure the remaining officers are rearranged to have no gaps
            specified_rows = hod_rankings[hod_rankings['PMS Code'] == posting].index
            cols_to_modify = match_cols[len(final_priority_list):len(final_priority_list)+len(remaining_officers)]
            hod_rankings.loc[specified_rows, cols_to_modify] = remaining_officers

            # Insert the prioritized officers at the top
            hod_rankings.loc[specified_rows, match_cols[:len(final_priority_list)]] = final_priority_list

            # Add comment to officer_choices
            officer_choices.loc[officer_choices['Employee ID'].isin(GDFM1_priority), 'Comment'] = 'GDFM Bump Intake 1'
            officer_choices.loc[officer_choices['Employee ID'].isin(GDFM2_priority), 'Comment'] = 'GDFM Bump Intake 2'

    return hod_rankings, officer_choices

def main(hod_rankings_path, officer_choices_path, GDFM_list_path):
    # Preparing the output folder
    output_folder = os.path.dirname(officer_choices_path) 
    output_folder_path = os.path.join(output_folder, 'output')
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    # Input file paths
    hod_rankings_path = args.HODRankFile
    officer_choices_path = args.MORankFile
    GDFM_list_path = args.GDFM
    posting_license_requirement_path = os.path.join(output_folder, 'Posting License Requirements.xlsx')

    #Output file paths
    no_choice_officers_path     = os.path.join(output_folder_path, 'no_choice_officers.csv')
    no_vacancies_departments_path = os.path.join(output_folder_path, 'no_vacancies_departments.csv')
    mutualmatches_path = os.path.join(output_folder_path, 'mutualmatches.xlsx')
    exceptions_path = os.path.join(output_folder_path, 'exceptions.txt')
    second_departments_with_vacancies_path = os.path.join(output_folder_path, 'Departments with vacancies.xlsx')
    second_unmatched_officers_path = os.path.join(output_folder_path, 'Unmatched officers.xlsx')
    combined_match_path = os.path.join(output_folder_path, 'Final Matches.xlsx')
    
    # Load data from Excel files.
    hod_rankings = pd.read_excel(hod_rankings_path)
    officer_choices = pd.read_excel(officer_choices_path, dtype={'Employee ID': str})
    posting_license_requirement = pd.read_excel(posting_license_requirement_path)
    GDFM1 = pd.read_excel(GDFM_list_path, sheet_name=0, header=1, dtype={'Employee ID': str})
    GDFM2 = pd.read_excel(GDFM_list_path, sheet_name=1, header=1, dtype={'Employee ID': str})

    #Ensure employee ID data type in HOD Rankings is formatted as string
    match_cols = [col for col in hod_rankings.columns if 'Match' in col]
    hod_rankings[match_cols] = hod_rankings[match_cols].astype(str).replace('nan', '')
    
    # Identify choice columns
    choice_columns = [col for col in officer_choices.columns if 'choice' in col.lower()]

    # Handle officers who didn't rank any departments
    no_choice_columns = ['Employee Name', 'Employee ID'] + choice_columns
    no_choice_officers = officer_choices[officer_choices[choice_columns].isnull().all(axis=1)][no_choice_columns]
    no_choice_officers.to_csv(no_choice_officers_path, sep=',', index=False)

    # Remove officers with no choices from the original DataFrame
    officer_choices = officer_choices[~officer_choices[choice_columns].isnull().all(axis=1)]

    # Handle departments with no initial vacancies.
    no_vacancies_columns = ['Postings', 'PMS Code', 'Vacancies']
    no_vacancies_departments = hod_rankings[hod_rankings['Vacancies'] == 0][no_vacancies_columns]
    no_vacancies_departments.to_csv(no_vacancies_departments_path, sep=',', index=False)

    # Remove departments with no vacancies from the original DataFrame.
    hod_rankings = hod_rankings[hod_rankings['Vacancies'] > 0]

    #Filter out based on license requirements
    officer_choices['Registration Type'] = officer_choices['Registration Type'].replace(['Provisional'], 'Full') #Assume all Provisional licenses will promote to Full licenses
    officer_choices = check_license(officer_choices, posting_license_requirement)

    #Bump selected GDFM candidates
    hod_rankings, officer_choices = GDFM_bump(GDFM1, GDFM2, officer_choices, hod_rankings)

    # Call both rounds of matching and combine the final matches
    mutualmatch_df, first_unmatched_officers_df, first_departments_with_vacancies_df, exceptions = gale_shapley_1(hod_rankings, officer_choices, no_vacancies_departments)
    second_match_df, second_unmatched_officers_df, second_departments_with_vacancies_df = gale_shapley_2(first_unmatched_officers_df,first_departments_with_vacancies_df, hod_rankings)
    combined_match_df = pd.concat([mutualmatch_df, second_match_df], axis=0)
    #combined_match_df = pd.concat([preallocated_matches_df, combined_match_df], axis=0)
    combined_match_df=combined_match_df.sort_values(by=['PMS Code','Employee ID'])

    # Write the exceptions to a new txt file.
    with open(exceptions_path, 'w') as f:
        for officer, department in exceptions:
            f.write(f"Employee ID: {officer}, Department PMS Code: {department}\n")

    # Write outputs to a new Excel files.
    mutualmatch_df.to_excel(mutualmatches_path, index=False)
    second_unmatched_officers_df.to_excel(second_unmatched_officers_path, index=False)
    second_departments_with_vacancies_df.to_excel(second_departments_with_vacancies_path)
    combined_match_df.to_excel(combined_match_path, index=False)

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--HODRankFile', type=str, required=True, help="The file name of HOD Rankings Consolidated")
    parser.add_argument('--MORankFile', type=str, required=True, help="The file name of MO Ranking Consolidated")
    parser.add_argument('--GDFM', type=str, required=True, help="The file name of GDFM List")
    args = parser.parse_args()
    main(args.HODRankFile, args.MORankFile, args.GDFM)
