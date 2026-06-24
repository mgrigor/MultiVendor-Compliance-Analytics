import os
import sys
import time
import pandas as pd

start = time.ctime()
print('''
################################################################################
#                                                                              #
#                       Samsung CELL Part 1 Precheck v1.0                      #
#                                                                              #
################################################################################
''')

directory = 'Raw Data'
tech = 'CELL'
level = 'Part 1'

def Select_subMarket():
    """Prompts user for regional submarket validation."""
    print('Available subMarkets:')
    while True:
        items = os.listdir('Raw Data')
        valid_folders = [item for item in items if os.path.isdir('Raw Data/' + item)]
        for folder in valid_folders:
            print(f" ↳ {folder}")
            
        choice = input("\nPlease select subMarket from the list: ").upper()
        if choice in valid_folders:
            return choice
        else:
            print("\n❌ subMarket not found, please try again.\n")

def Samsung_Parameters():
    """Extracts required CELL parameters dynamically grouped by Script Name."""
    print('\n > Generating dataframe for required parameters...')
#    file = 'Samsung Required Parameters Master.xlsx'
    file = 'Demo_GPL_Samsung.xlsx'
    
    df = pd.read_excel(file, sheet_name='master', index_col=False, dtype=str)

    # Filter master criteria rules specifically for CELL Part 1 parameters
    cell_group = df.loc[(df['Technology'] == tech) & (df['Group'] == level)]
    group_script_names = cell_group['Group Script Name'].unique().tolist()
    
    reqd_parameters = []
    group_script_name_dict = {}
    groupby_script_name = cell_group.groupby('Group Script Name')
    
    for group_script, data in groupby_script_name:
        parameters = data['Parameter'].unique().tolist()
        parameters.sort()
        group_script_name_dict.update({tuple(parameters): group_script})
        reqd_parameters.append(parameters)

    # Baseline multi-part tracking schema for cell sectors
    labels = ['date', 'subMarket', 'market', 'enbname', 'enbid', 'cellId', 'sector', 'carrier']
    headers_and_values = {label: '' for label in labels}
    
    for _, row in cell_group.iterrows():
        headers_and_values.update({row['Parameter']: row['Airwave Setting']})
        
    df_headers = pd.DataFrame(headers_and_values, index=[0])
    return (reqd_parameters, group_script_name_dict, df_headers)

def checking_data_consistency1(reqd_parameters, group_script_name_dict):
    print('\n > Checking data structure integrity and footprint consistency...')
    path = directory + '/' + subMarket + '/' + tech + '/' + level
    
    if not os.path.exists(path):
        # Fallback path traversal support if level nested folders aren't present
        path = directory + '/' + subMarket + '/' + tech

    if not os.path.exists(path):
        print(f"❌ Error: Target data directory does not exist -> {path}")
        sys.exit()

    labels = ['date', 'subMarket', 'market', 'enbid', 'enbname', 'cellId', 'sector', 'carrier']
    files = os.listdir(path)
    csvfiles = [f for f in files if f.endswith('.csv')]
    
    conf_parameter_sets = []
    seen_file_signatures = []  # Robust order-independent duplicate file catcher
    
    for csvfile in csvfiles:
        # Performance: Read row 0 to extract header schema instantly
        df_headers = pd.read_csv(path + '/' + csvfile, nrows=0)
        
        # Read row 1 to extract the file metadata market name
        df_market_check = pd.read_csv(path + '/' + csvfile, nrows=1)
        if df_market_check.empty:
            print(f"⚠️ Warning: File {csvfile} is empty. Skipping.")
            continue
            
        SubMarket = df_market_check['subMarket'].iloc[0].split('-')[-1]
        print(f"   Checking File SubMarket: {SubMarket} | File Name: {csvfile}")
        
        # 1. Enforce strict geographical region alignment
        if subMarket != SubMarket:
            print(f'\n❌ CRITICAL: Invalid submarket data detected inside {csvfile}!')
            print('Please download the correct data for this market region.')
            sys.exit()
            
        # 2. Extract core parameter metrics
        parameters = [col for col in df_headers.columns if not col.startswith('paramId') and col not in labels]
        
        # RESTORED CRITICAL RULE: Append QCI suffix to map correctly with Master guidelines
        if 'qci-value' in parameters:
            parameters = [p + '(QCI)' for p in parameters]
            
        parameters.sort()
        file_signature = frozenset(parameters)

        # 3. Restored Safeguard: Duplicate Dataset Verification
        if file_signature in seen_file_signatures:
            print(f'\n  PRECHECK ABORTED: Duplicate tracking dataset found inside folder!')
            print(f'   The file "{csvfile}" contains the exact same metric layout as another file.')
            print('   Please remove the duplicate configuration report and rerun the script.')
            sys.exit()
        else:
            seen_file_signatures.append(file_signature)
            conf_parameter_sets.append(set(parameters))
            
        print('   ↳ Parameters staged successfully.')

    # 4. Set-Driven Structural Validation (Replaces brittle list matching)
    print('\n > Auditing staged files against Master Blueprint Groups...')
    overall_passed = True
    
    for req_list in reqd_parameters:
        req_set = set(req_list)
        group_script_name = group_script_name_dict[tuple(req_list)]
        
        matched = False
        for file_set in conf_parameter_sets:
            if req_set.issubset(file_set):
                matched = True
                break
                
        if matched:
            print(f"   ✅ PASS: Group '{group_script_name}' metrics verified.")
        else:
            print(f"   ❌ FAIL: Missing parameters for Group Script Name: {group_script_name}")
            print(f"      ↳ Missing Columns: {req_list}")
            overall_passed = False
            
    if not overall_passed:
        print('\n   PRECHECK FAILED: Structural anomalies discovered inside folder data.')
        print('\nPlease update your master configurations or extract missing data views.')
        sys.exit()
           
    print('\n   Process completed successfully')
    print('\n   Safe to run the next script: Samsung CELL Part 1 Prep & Merge.')

# ==============================================================================
# PIPELINE EXECUTION
# ==============================================================================    
subMarket = Select_subMarket()
reqd_params, group_dict, df_hdrs = Samsung_Parameters()
checking_data_consistency1(reqd_params, group_dict)

end = time.ctime()

print('\nStart:', start)
print('End  :', end)
