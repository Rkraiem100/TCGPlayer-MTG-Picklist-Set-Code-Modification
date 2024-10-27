import os
import configparser
import pandas as pd
import sys
import time

def create_default_ini(ini_path):
    # Creating default INI file with the provided mappings
    config = configparser.ConfigParser()
    config['SetNames'] = {
        'Magic Origins': 'ORI',
        'Wilds of Eldraine': 'WOE',
        'Ravnica Remastered': 'RVR',
        'Modern Horizons 2': 'MH2',
        'Theros Beyond Death': 'THB',
        'Commander Legends Battle for Baldur\'s Gate': 'CLB',
        'Strixhaven: School of Mages': 'STX',
        'Ikoria: Lair of Behemoths': 'IKO',
        'Core Set 2021': 'M21',
        'Outlaws of Thunder Junction': 'OTJ',
        'Zendikar Rising': 'ZNR',
        'Kaldheim': 'KHM',
        'Zendikar': 'No set acronym',
        'The Brothers\' War: Retro Frame Artifacts': 'No set acronym',
        'Commander Legends': 'CMR'
    }
    with open(ini_path, 'w') as configfile:
        config.write(configfile)
    print(f"Default INI file created at {ini_path}. Please modify it with additional set names as needed.")

def load_set_mappings(ini_path):
    # Load the set name mappings from the INI file
    config = configparser.ConfigParser()
    config.read(ini_path)
    set_mappings = {}
    if 'SetNames' in config:
        set_mappings = {k.strip().lower(): v.strip() for k, v in config['SetNames'].items()}
    return set_mappings

def modify_data(df, set_mappings):
    # Ensure the output column "Set Code" exists
    if 'Set Code' not in df.columns:
        df['Set Code'] = ''

    # Iterate over each row starting from the first data row (ignoring the header)
    for index, row in df.iterrows():
        if str(row['Product Line']).strip().lower() == 'magic':
            product_name = str(row['Product Name']).strip().lower()
            if '(retro frame)' in product_name:
                # If the product name contains "(Retro Frame)", set the code to "Retro Frame"
                set_code = 'Retro Frame'
            else:
                # Otherwise, use the default set code logic
                set_name = str(row['Set']).strip().lower()
                set_code = set_mappings.get(set_name, 'null')
            # Update the "Set Code" column
            df.at[index, 'Set Code'] = set_code
    return df

def modify_file(file_path, set_mappings):
    # Determine the input file type (CSV or Excel)
    if file_path.endswith('.csv'):
        # Load the CSV file into a DataFrame
        df = pd.read_csv(file_path, delimiter=',')
    elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        # Load the Excel file into a DataFrame
        df = pd.read_excel(file_path, engine='openpyxl')
    else:
        print("Unsupported file format. Please provide a CSV or Excel file.")
        sys.exit(1)

    # Modify the data
    df = modify_data(df, set_mappings)

    # Save the modified data to a new Excel file named "Modified Pullsheet.xlsx"
    output_file_path = os.path.join(os.path.dirname(file_path), 'Modified Pullsheet.xlsx')
    df.to_excel(output_file_path, index=False, engine='openpyxl')
    print(f"Modifications complete and saved to Excel file: {output_file_path}")

def main():
    # Check if the script is executed with a file argument
    if len(sys.argv) < 2:
        print("Usage: Drag and drop the CSV or Excel file onto the executable to run the modifications.")
        sys.exit(1)

    # Get the path to the input file from the command-line arguments
    file_path = sys.argv[1]

    # Define the INI file name and path
    ini_file_name = 'TCGPlayer Picklist Mod.ini'
    ini_file_path = os.path.join(os.path.dirname(sys.argv[0]), ini_file_name)

    # Check if the INI file exists, otherwise create a default one
    if not os.path.exists(ini_file_path):
        create_default_ini(ini_file_path)
    
    # Load the set mappings from the INI file
    set_mappings = load_set_mappings(ini_file_path)

    # Modify the input file and save it as an Excel file
    modify_file(file_path, set_mappings)

    # Display confirmation and wait for user input to exit
    print("All modifications are complete.")
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
