import os
import sys
import time
import pandas as pd

start = time.ctime()
print('''
################################################################################
#                                                                              #
#                        Samsung CELL Part 2 Precheck v1.0                     #
#                                                                              #
################################################################################
''')

directory = 'Raw Data'
tech = 'CELL'
level = 'Part 2'

def Select_subMarket():
    """Prompts the user to select a valid regional submarket folder."""
    print('Available subMarkets:')
    while True:
        items = os.listdir('Raw Data')
        valid_folders = [item for item in items if os.path.isdir(os.path.join('Raw Data', item))]
        for folder in valid_folders:
            print(f" ↳ {folder}")
        choice = input("\nPlease select subMarket from the list: ").upper()
        if choice in valid_folders:
            return choice
        else:
            print("\n❌ subMarket not found, please try again.\nAvailable subMarkets:")

def Samsung_Parameters():
    """Dynamically loads required parameters from the master blueprint sheet."""
#    file = 'Samsung Required Parameters Master.xlsx'    # For live network
    file = 'Demo_GPL_Samsung.xlsx'   # For demo purposes
    
    df = pd.read_excel(file, sheet_name='master', index_col=False, dtype=str)

    cell_group = df.loc[(df['Technology'] == tech) & (df['Group'] == level)]
    reqd_parameters = []
    group_script_name_dict = {}
    groupby_script_name = cell_group.groupby('Group Script Name')
    
    for group_script, data in groupby_script_name:
        # Isolate base parameter names from status brackets cleanly
        parameters = [item.split('[')[0] for item in data['Parameter'].unique().tolist()]
        parameters = sorted(list(set(parameters)))
        reqd_parameters.append(parameters)
        group_script_name_dict.update({tuple(parameters): group_script})

    labels = ['date', 'subMarket', 'market', 'enbname', 'enbid', 'cellId', 'sector', 'carrier']
    headers_and_values = {label: '' for label in labels}
    for i in range(cell_group.shape[0]):
        headers_and_values.update({cell_group.iloc[i]['Parameter']: cell_group.iloc[i]['Airwave Setting']})
        
    df_headers = pd.DataFrame(headers_and_values, index=[0])
    return (reqd_parameters, group_script_name_dict, df_headers)

def transform_dataframe(df):
    """Vectorized multi-attribute pivot engine with corrected identity label protections."""
    # PROTECTED: added carrierId as a clean identity key to exclude it from parameter modifications
    labels = ['date', 'subMarket', 'market', 'enbid', 'enbname', 'cellId', 'sector', 'carrier', 'carrierId']

    # Isolate status type values from brackets dynamically
    df['paramId1'] = df['paramId1'].apply(lambda x: x.split('=')[1].strip(']'))
    df = df[df['paramId1'] != 'bar-manual']
    
    # Pristine Selection: Picks up only pure parameter variables to pivot
    parameters = [col for col in df.columns if col not in labels and col != 'paramId1' and not col.startswith('paramId')]

    # Pivot un-flattened status arrays horizontally at high speed
    pivot_df = df.pivot(index=labels, columns='paramId1', values=parameters)
    
    # Collapse MultiIndex hierarchies cleanly into standard "parameter_name[status_type]"
    pivot_df.columns = [f'{param}[{status}]' for param, status in pivot_df.columns]
    
    final_df = pivot_df.reset_index()
    
    # Restore standard structural column positioning orientation (carrierId at position 8)
    columns_list = final_df.columns.tolist()
    columns_list.remove('carrierId')
    columns_list.insert(8, 'carrierId')
    return final_df[columns_list]

def checking_data_consistency1(reqd_parameters, group_script_name_dict):
    print(' > Checking data structure integrity and footprint consistency...')
    labels = ['date', 'subMarket', 'market', 'enbid', 'enbname', 'cellId', 'sector', 'carrier']
    path = os.path.join(directory, subMarket, tech, level)
    
    if not os.path.exists(path):
        path = os.path.join(directory, subMarket, tech)
        
    files = os.listdir(path)
    csvfiles = [file for file in files if file.endswith('.csv') and not file.endswith('-x.csv') and not file.startswith('~$')]
    
    conf_parameter_sets = []
    seen_file_signatures = []  # Robust order-independent duplicate file catcher list
    
    for csvfile in csvfiles:
        df = pd.read_csv(os.path.join(path, csvfile), index_col=False, dtype=str, sep=',')
        
        # SEQUENCE RULE 1: EXCLUDE NON-OPERATIONAL "GROW" CELLS INSTANTLY
        df = df[~df['enbname'].str.lower().str.startswith('grow', na=False)].reset_index(drop=True)
        if df.empty:
            continue
            
        # Verify submarket tracking region matches target selection
        submarket_check = df['subMarket'].dropna().unique().tolist()[0].split('-')[-1]
        print(f"   Checking File SubMarket: {submarket_check} | File Name: {csvfile}")
        
        if subMarket != submarket_check:
            print(f'   ❌ CRITICAL ERROR: Invalid submarket data detected inside {csvfile}! Exiting.')
            sys.exit()

        # Isolate baseline parameters
        parameters = [col for col in df.columns if not (col.startswith('param') or col.startswith('path'))]
        for label in labels:
            if label in parameters:
                parameters.remove(label)
        parameters.sort()
        
        # Order-independent duplicate file catcher using frozensets
        file_signature = frozenset(parameters)
        if file_signature in seen_file_signatures:
            print(f'   🚨 PRECHECK ABORTED: Duplicate tracking dataset configuration detected for {csvfile}!')
            sys.exit()
        seen_file_signatures.append(file_signature)

        # Validate structure against Master blueprints before proceeding to save intermediate file
        matched_any_group = False
        for r_p in reqd_parameters:
            if set(r_p).issubset(set(parameters)):
                matched_any_group = True
                conf_parameter_sets.append(set(parameters))
                
                # SEQUENCE RULE 2: APPLY LIVE 4G5G PRIORITIZATION FILTER
                df['enbid'] = df['enbid'].astype(str).str.zfill(6)
                for label in ['cellId', 'sector', 'carrier']:
                    df[label] = df[label].astype(float).astype(int).astype(str)
                df['carrierId'] = df['enbid'] + '-' + df['cellId'] + '-' + df['sector'] + '-' + df['carrier']
                
                df['is_modernized'] = df['enbname'].str.contains('4G5G', na=False, case=False)
                has_modernized = df.groupby('carrierId')['is_modernized'].transform('any')
                df = df[df['is_modernized'] | ~has_modernized].reset_index(drop=True)
                df.drop(columns=['is_modernized'], inplace=True)

                # Execute horizontal flattening transformation and save intermediate snapshot
                transformed_df = transform_dataframe(df)
                new_csvfile = csvfile.replace('.csv', '-x.csv')
                transformed_df.to_csv(os.path.join(path, new_csvfile), index=False)
                
                print('   ↳ Parameters staged and transformed successfully.')
                break
                
        if not matched_any_group:
            print(f'   ❌ Invalid parameter structure found inside {csvfile}! Check source columns.')
            sys.exit()

    # ─── THE POLISHED PASS/FAIL AUDIT BLOCK ───────────────────────────────────
    print('\n > Auditing staged files against Master Blueprint Groups...')
    overall_passed = True
    
    for parameters_list in reqd_parameters:
        req_set = set(parameters_list)
        group_script_name = group_script_name_dict[tuple(parameters_list)]
        
        matched = False
        for file_set in conf_parameter_sets:
            if req_set.issubset(file_set):
                matched = True
                break
                
        if matched:
            print(f"   ✅ PASS: Group '{group_script_name}' metrics verified.")
        else:
            print(f"   ❌ FAIL: Missing parameters for Group Script Name: {group_script_name}")
            print(f"      ↳ Expected base metrics: {parameters_list}")
            overall_passed = False
            
    if not overall_passed:
        print('\n🛑 Precheck Gate Notice: Data parameters are missing or group counts are incomplete.')
        sys.exit()
      
    print('\n   Process completed successfully!.')
    print('\n   Ready to run the next script: Samsung CELL Part 2 Prep Merge & Audit.')

# ==============================================================================
# PIPELINE EXECUTION
# ==============================================================================
subMarket = Select_subMarket()
print(f'\n > Processing data for market region: {subMarket}...\n')

reqd_parameters, group_script_name_dict, df_headers = Samsung_Parameters()
checking_data_consistency1(reqd_parameters, group_script_name_dict)

end = time.ctime()
print(f'\nStart: {start}\nEnd  : {end}')
