import os
import sys
import time
import pandas as pd

start = time.ctime()
print('#######################################################')
print('#                                                     #')
print('#             Samsung ENB Part 1 Precheck             #')
print('#                                                     #')
print('#######################################################')

directory = 'Raw Data'
tech = 'ENB'
level = 'Part 1'

def Select_subMarket():
    print('Available subMarkets:')
    while True:
        items = os.listdir('Raw Data')
        for item in items:
            if os.path.isdir('Raw Data/'+item) is True:
                print(item)
        choice = input("\nPlease select subMarket from the list: ")
        choice = choice.upper()
        if choice in items:
            return choice
        else:
            print("\nsubMarket not found, please try again.")
            print('\nAvailable subMarkets:')

def Samsung_Parameters():
    """Loads the required parameters and groups them by their Script Group Name."""
#    file = 'Samsung Required Parameters Master.xlsx'
    file = 'Demo_GPL_Samsung.xlsx'
    
    df = pd.read_excel(file, sheet_name='master', index_col=False, dtype=str)
    enb_group1 = df.loc[(df['Technology'] == 'ENB') & (df['Group'] == 'Part 1')]
    group_script_names = enb_group1['Group Script Name'].unique().tolist()
    
    script_name_parameter_map = []
    for name in group_script_names:
        if type(name) is str:
            reqd_params = []
            group_parameters = enb_group1.loc[(enb_group1['Group Script Name'] == name)]
            reqd_parameters_raw = group_parameters['Parameter'].unique().tolist()
            
            for item in reqd_parameters_raw:
                reqd_parameter = item.split('[')[0]
                if reqd_parameter not in reqd_params:
                    reqd_params.append(reqd_parameter)
                    
            script_name_parameter_map.append([name, reqd_params])
    
    return script_name_parameter_map

def loading_data_files_old(subMarket):
    """Loads individual CSV headers, checks for duplicate datasets, and verifies master structures."""
    print('\n > Loading raw data files for subMarket:', subMarket)
    fpath = directory + '/' + subMarket + '/' + tech + '/' + level
    
    if not os.path.exists(fpath):
        print(f"❌ Error: Folder path tree does not exist -> {fpath}")
        sys.exit()
        
    files = os.listdir(fpath)
    data_files = [file for file in files if file.endswith('.csv')]
    
    conf_parameters_file_map = []
    conf_parameter_sets = []
    seen_file_signatures = []  # Tracks unique parameter footprints to catch duplicate files
    
    for file in data_files:
        # Performance: read only the header columns
        df = pd.read_csv(fpath + '/' + file, dtype=str, index_col=False, nrows=0)
        SubMarket = pd.read_csv(fpath + '/' + file, dtype=str, index_col=False, nrows=1)['subMarket'].iloc[0].split('-')[-1]
        
        print(f"   Checking File SubMarket: {SubMarket} | File Name: {file}")
        
        if subMarket == SubMarket:
            headers = df.columns.tolist()
            
            # Isolate parameter metric column names
            paramIDs = [h for h in headers if h.startswith('paramId') or h.startswith('enb-')]
            for p_id in paramIDs:
                headers.remove(p_id)
        
            parameters = headers[5:]
            
            # --- CRITICAL RESTORED SAFEGARD: Duplicate File Detection ---
            # Using frozenset allows order-independent identification of duplicate files
            file_signature = frozenset(parameters)
            if file_signature in seen_file_signatures:
                print('\n   PRECHECK ABORTED: Duplicate parameters / dataset found!')
                print(f"   The file '{file}' contains the exact same metric data structure as another file in this folder.")
                print('   Please remove the duplicate report file and run the script again.')
                sys.exit()
            else:
                seen_file_signatures.append(file_signature)
                conf_parameter_sets.append(set(parameters))
                conf_parameters_file_map.append([file, parameters])
        else:
            print('\n❌ Invalid submarket data detected!')
            print('Download the needed file[s] and run script again.')
            sys.exit()
        
    # Checking for missing required parameters using set logic
    print('\n > Checking for missing required parameters against Master Groups...')
    
    overall_passed = True
    for item in script_name_parameter_map:
        name, reqd_parameters = item
        reqd_set = set(reqd_parameters)
        
        matched = False
        for file_set in conf_parameter_sets:
            if reqd_set.issubset(file_set):
                matched = True
                break
                
        if matched:
            print(f"   ✅ PASS: Group '{name}' columns verified in raw files.")
        else:
            print(f"   ❌ FAIL: Missing required parameters for Group Script Name: {name}")
            print(f"      ↳ Missing Columns from: {reqd_parameters}")
            overall_passed = False
            
    if not overall_passed:
        print('\nDownload the needed file[s] or check your Airwave script.')
        print('Exiting script...')
        sys.exit()
        
    print(f'\n   Process completed successfully!')
    print('\n   Safe to run the next script: Samsung ENB Prep & Merge.')
    return (conf_parameter_sets, conf_parameters_file_map)

# ==============================================================================
# PIPELINE EXECUTION
# ==============================================================================

# Selecting subMarket
subMarket = Select_subMarket()
print(f'\n > Processing data for {subMarket}...')

# Load Samsung required parameters
script_name_parameter_map = Samsung_Parameters()

# Read and structural check raw data files
conf_parameters, conf_parameters_file_map = loading_data_files_old(subMarket)

end = time.ctime()
print('')
print('Start:', start)
print('End  :', end)
