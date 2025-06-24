"""
README:
This code will compile each department's HOD ranking into a consolidated Excel file for matching.
It assumes that all raw data files are located within the same folder and are labelled correctly based on PMS Code.
"""

# import necessary libraries
import pandas as pd
import os
import argparse

# Load the MOPEX 25 Staff List and create MCR to Employee ID mapping
def load_mcr_to_employee_mapping(staff_list_path):
    staff_list = pd.read_excel(staff_list_path)
    mcr_to_employee_id = dict(zip(staff_list['MCR No.'], staff_list['Employee ID']))
    return mcr_to_employee_id

# Load and map MCR/DCR to Employee ID in raw data files
def load_and_map_data(file_path, mcr_to_employee_id):
    data = pd.read_excel(file_path)
    
    # Map MCR/DCR Number to Employee ID if the column exists
    if 'MCR/DCR Number' in data.columns:
        data['Employee ID'] = data['MCR/DCR Number'].map(mcr_to_employee_id)
    return data

# Process multiple raw data files based on PMS codes and maps MCR to employee ID
def process_raw_data_files(pms_codes, raw_data_directory, mcr_to_employee_id):
    processed_data = {}
    
    for pms_code in pms_codes:
        file_name = f"{pms_code}.xlsx"
        file_path = os.path.join(raw_data_directory, file_name)
        
        if os.path.exists(file_path):
            processed_data[pms_code] = load_and_map_data(file_path, mcr_to_employee_id)
        else:
            print(f"File not found for PMS Code: {pms_code}")
    
    return processed_data

# Load the consolidated HOD Rankings
def load_consolidated_hod_rankings(hod_rankings_path):
    return pd.read_excel(hod_rankings_path)

# Process department rankings, sort by HOD and MO rankings, and update the consolidated file
def process_and_update_rankings(dept_df, consolidated_df, index, filename):
    # Convert 'HOD Ranking' and 'MO Ranking' to numeric, non-numeric values will be NaN
    dept_df['HOD Ranking'] = pd.to_numeric(dept_df['HOD Ranking'], errors='coerce')
    dept_df['MO Ranking'] = pd.to_numeric(dept_df['MO Ranking'], errors='coerce')

    # Sort the data by 'HOD Ranking', 'MO Ranking', and 'Employee ID'
    dept_df = dept_df.sort_values(by=['HOD Ranking', 'MO Ranking', 'Employee ID'], na_position='last')

    # Save the sorted data back to the department's file
    dept_df.to_excel(filename, index=False)

    # Filter the DataFrame to include only rows with numerical 'HOD Ranking'
    filtered_dept_df = dept_df.dropna(subset=['HOD Ranking'])

    # Copy the 'Employee ID' from the filtered DataFrame and transpose it
    employee_ids = filtered_dept_df['Employee ID'].tolist()

    # Create a dictionary with column keys and corresponding Employee IDs
    match_columns = {f'Match {i+1}': employee_id for i, employee_id in enumerate(employee_ids)}

    # Add the new columns to the consolidated DataFrame
    consolidated_df.loc[index, match_columns.keys()] = match_columns.values()

# Main function to compile rankings and process raw data
def compile_rankings(consolidated_df, raw_data_directory, mcr_to_employee_id):
    pms_codes = consolidated_df['PMS Code'].unique()
    processed_data_files = process_raw_data_files(pms_codes, raw_data_directory, mcr_to_employee_id)

    # Iterate over each row in the consolidated HOD rankings
    for index, row in consolidated_df.iterrows():
        pms_code = row['PMS Code']
        
        # Check if processed data exists for the current PMS code
        if pms_code in processed_data_files:
            dept_df = processed_data_files[pms_code]
            filename = os.path.join(raw_data_directory, f"{pms_code}.xlsx")
            
            # Process and update the department rankings
            process_and_update_rankings(dept_df, consolidated_df, index, filename)

    return consolidated_df

#Check for blank MCR numbers
def check_blank_column_errors(df):
    """
    Checks for rows where a blank cell is followed by a non-blank cell.
    Returns a list of indices of rows with such errors.
    """
    error_rows = []
    for index, row in df.iterrows():
        if any(pd.isna(row[i]) and not pd.isna(row[i + 1]) for i in range(len(row) - 1)):
            error_rows.append(index)
    return error_rows

# Main entry point
if __name__ == "__main__":
    # Create the parser and add arguments
    parser = argparse.ArgumentParser()
    parser.add_argument('--dir', type=str, required=True, help="The target directory path")
    parser.add_argument('--HODRankFile', type=str, required=True, help="The file name of consolidated HOD Rankings")
    parser.add_argument('--MOPEXStaffList', type=str, required=True, help="The file name of MOPEX 25 Staff List")
    args = parser.parse_args()
    
    # Set the directory path and file names
    raw_dir_path = args.dir
    hod_rankings_path = args.HODRankFile
    output_dir_path = os.path.dirname(hod_rankings_path)
    staff_list_path = args.MOPEXStaffList
    
    # Load the MCR to Employee ID mapping
    mcr_to_employee_id = load_mcr_to_employee_mapping(staff_list_path)
    
    # Load the consolidated HOD rankings
    consolidated_df = load_consolidated_hod_rankings(hod_rankings_path)
    
    # Compile rankings and process raw data files
    processed_data = compile_rankings(consolidated_df, raw_dir_path, mcr_to_employee_id)

    # Validate for blank column errors
    error_rows = check_blank_column_errors(consolidated_df)

    if error_rows:
        print(f"Error: The following rows contain invalid blank columns: {error_rows}")
        print("Please ensure the staff listing and raw data files match correctly.")
        exit(1)  # Exit the script if errors are found

    # save the consolidated file
    consolidated_df.to_excel(os.path.join(output_dir_path, "HOD Rankings Consolidated.xlsx"), index=False)