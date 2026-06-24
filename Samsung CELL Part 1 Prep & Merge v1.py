import os
import sys
import time
import xlsxwriter
import pandas as pd

start = time.ctime()
print('''
################################################################################
#                                                                              #
#                     Samsung CELL Part 1 Prep & Merge v1.0                    #
#                                                                              #
################################################################################
''')

directory = 'Raw Data'
tech = 'CELL'
level = 'Part 1'


def Select_subMarket():
    """Prompts the user to select a valid regional submarket folder."""
    print('Select subMarkets to process:\n')
    while True:
        items = os.listdir('Raw Data')
        valid_markets = [item for item in items if os.path.isdir('Raw Data/' + item)]
        for market in valid_markets:
            print(f" ↳ {market}")
        choice = input("\nPlease select subMarket from the list: ").upper()
        if choice in valid_markets:
            return choice
        print("\n❌ subMarket not found, please try again.")


def Samsung_Parameters():
    """Dynamically extracts parameters from the master blueprint sheet to build base frame columns."""
    print('\n > Generating dataframe headers dynamically from blueprint...')
#    file = 'Samsung Required Parameters Master.xlsx' #for live network
    file = 'Demo_GPL_Samsung.xlsx' # for demo purposes
    
    df = pd.read_excel(file, sheet_name='master', index_col=False, dtype=str)

    # Filter to isolate target blueprint parameters matching CELL Part 1 definitions
    cell_group = df.loc[(df['Technology'] == tech) & (df['Group'] == level)]
    reqd_parameters = []
    group_script_name_dict = {}
    groupby_script_name = cell_group.groupby('Group Script Name')
    
    for group_script, data in groupby_script_name:
        parameters = data[data['Group Script Name'] == group_script]['Parameter'].unique().tolist()
        group_script_name_dict.update({tuple(parameters): group_script})
        reqd_parameters.append(parameters)

    # Core multi-part sector metadata tracking row layout
    labels = ['date', 'subMarket', 'market', 'enbname', 'enbid', 'cellId', 'sector', 'carrier']
    headers_and_values = {label: '' for label in labels}
    
    # Map master targets to baseline columns
    param_mappings = dict(zip(cell_group['Parameter'], cell_group['Airwave Setting']))
    headers_and_values.update(param_mappings)
    
    df_headers = pd.DataFrame(headers_and_values, index=[0])
    
    # Inject carrierId at standard index position 8 to act as our master structural key anchor
    df_headers['carrierId'] = ''
    columns = df_headers.columns.tolist()
    columns.remove('carrierId')
    columns.insert(8, 'carrierId')
    df_headers = df_headers[columns]
    
    df_headers = df_headers.map(lambda x: x.lower().strip() if isinstance(x, str) else x)
    return (reqd_parameters, group_script_name_dict, df_headers)


def loading_data_files(subMarket, reqd_parameters, group_script_name_dict):
    """Loads CSV files, executes immediate front-door GROW cell removal, and handles deduplication."""
    path = directory + '/' + subMarket + '/' + tech + '/' + level
    if not os.path.exists(path):
        path = directory + '/' + subMarket + '/' + tech
        
    print(f'\n > Loading raw data CSV logs from: {path}')
    files = os.listdir(path)
    data_files = [file for file in files if file.endswith('.csv') and not file.startswith('~$')]
    
    conf_parameters = {}
    frames = []
    labels = ['date', 'subMarket', 'market', 'enbname', 'enbid', 'cellId', 'sector', 'carrier']
    
    for file in data_files:
        df = pd.read_csv(path + '/' + file, index_col=False, sep=',', dtype=str)
        df['enbid'] = df['enbid'].astype(str).str.zfill(6)
        
        submarket_check = df['subMarket'].dropna().unique().tolist()[0].split('-')[-1]
        print(f"     Staging File: {file} (Market Check: {submarket_check})")
        
        if submarket_check != subMarket:
            print('       ❌ CRITICAL ERROR: Invalid submarket tracking data detected within dataset!')
            sys.exit()
        
        # ─── RULE 1 IMPLEMENTATION: FRONT-DOOR "GROW" DATA PURGE ───
        # Wipe out pre-operational staging rows before running any deduplication filters
        init_rows = df.shape[0]
        df = df[~df['enbname'].str.lower().str.startswith('grow', na=False)].reset_index(drop=True)
        if df.empty:
            print(f"      ⚠️ Skip: '{file}' contained only non-operational GROW entries after purge.")
            continue
        grow_purged = init_rows - df.shape[0]
        if grow_purged > 0:
            print(f"      [+] Wiped out {grow_purged} pre-operational GROW rows from memory.")
        
        # 1. Cleanse metadata counters and index strings (paramId columns)
        df = df.loc[:, ~df.columns.str.startswith('param')]
        
        # 2. Dynamic QCI Parameter Header Substitution Rule
        if 'qci-value' in df.columns:
            print("      [+] Appending '(QCI)' syntax mapping onto metrics...")  
            for col in df.columns:
                if col not in labels:
                    df.rename(columns={col: col + '(QCI)'}, inplace=True)
        
        # 3. Clean up sector labels to eliminate decimal notation padding (.0)
        for label in ['cellId', 'sector', 'carrier']:
            df[label] = df[label].astype(float).astype(int).astype(str)
            
        # ─── SAFE WHOLE-NUMBER CLEANER FOR ALL METRICS ───────────────────────
        # This scans all columns, leaves text alone, and strips .0 off numbers
        for col in df.columns:
            if col not in labels and col != 'carrierId':
                # Only apply to rows that end in an actual decimal zero
                df[col] = df[col].str.replace(r'\.0$', '', regex=True)
        # ──────────────────────────────────────────────────────────────────────
            
        # 4. Formulate the sector antenna composite layout mapping key
        df['carrierId'] = df['enbid'] + '-' + df['cellId'] + '-' + df['sector'] + '-' + df['carrier']
        
        # ─── RULE 2 IMPLEMENTATION: VECTORIZED 4G5G MODERNIZATION PRIORITY ───
        # Guarantees 4G5G Live rows sort to the top and win over legacy entries
        df['is_modernized'] = df['enbname'].str.contains('4G5G', na=False, case=False)
        df.sort_values(by=['carrierId', 'is_modernized'], ascending=[True, False], inplace=True)
        
        pre_dedup_rows = df.shape[0]
        df.drop_duplicates(subset=['carrierId'], keep='first', inplace=True)
        post_dedup_rows = df.shape[0]
        df.drop(columns=['is_modernized'], inplace=True)
        
        if pre_dedup_rows != post_dedup_rows:
            print(f"     ↳ Compressed {pre_dedup_rows - post_dedup_rows} multi-PLMN layout elements/redundant rows.")

        # Re-enforce clean column alignment before tracking footprints
        columns = df.columns.tolist()
        columns.remove('carrierId')
        columns.insert(8, 'carrierId')
        df = df[columns]
        
        # Capture current file parameter footprint to verify against rules ledger
        parameters_tuple = tuple(sorted([c for c in df.columns if c not in labels and c != 'carrierId']))
        conf_parameters[parameters_tuple] = file
        frames.append(df)

    # 6. Structural Validation Cross-Check (UPDATED TO FIX SUBSET BUG)
    # Flatten all parameters found across all loaded files into a single pool
    all_staged_parameters = set()
    for file_params in conf_parameters.keys():
        all_staged_parameters.update(file_params)

    for req_list in reqd_parameters:
        # Check if any required parameter in this blueprint group is missing from our pool
        missing_params = [p for p in req_list if p not in all_staged_parameters]
        
        if missing_params:
            group_script_name = group_script_name_dict[tuple(req_list)]
            print(f'   ❌ Critical Mandatory Metric Group Missing: {group_script_name}')
            print(f'      ↳ Missing columns: {missing_params}')
            sys.exit()
                      
    return frames


def merging_dataframes_to_single_dataframe(frames, df_headers, subMarket):
    """Executes multi-source horizontal merge routines, isolates CBRS cells, and exports report."""
    print('\n > Executing side-by-side horizontal outer merges on common keys...')
    
    merged_df = frames[0]
    labels_keys = ['date', 'subMarket', 'market', 'enbid', 'enbname', 'cellId', 'sector', 'carrier', 'carrierId']
    
    for next_frame in frames[1:]:
        common_join_keys = [c for c in labels_keys if c in merged_df.columns and c in next_frame.columns]
        merged_df = pd.merge(merged_df, next_frame, on=common_join_keys, how='outer')

    # Re-enforce exact blueprint layout sorting structure
    headers = df_headers.columns.tolist()
    for col in headers:
        if col not in merged_df.columns:
            merged_df[col] = ""
    merged_df = merged_df[headers]

    # Isolate and export specialized high-capacity CBRS cells
    print('   > Separating CBRS sectors (Max Call Boundary Rule = 400)...')
    CBRS = merged_df.loc[merged_df['max-call-count'] == '400']
    
    fpath_out = directory + '/' + subMarket
    os.makedirs(fpath_out, exist_ok=True)
    cbrs_path = f"{fpath_out}/Samsung_{subMarket}_CBRS.xlsx"
    CBRS.to_excel(cbrs_path, index=False)
    
    # Live rows are safely isolated (already cleared of GROW data up front)
    live_data_rows = merged_df.loc[merged_df['max-call-count'] != '400']
    
    # Prepend row 0 blueprint targets rule onto our live layout rows
    subMarket_report_new = pd.concat([df_headers, live_data_rows], ignore_index=True)

    # ─── XLSXWRITER STRUCTURAL EXPORT ENGINE ──────────────────────────────────
    filename = f"{fpath_out}/Samsung_{subMarket}_CELL_Part_1.xlsx"
    print(f' > Saving finalized consolidated report to Excel: {filename}')
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet("CELL Part 1 Baseline")

    header1_format = workbook.add_format({'font_color': 'black', 'bold': True})
    header2_format = workbook.add_format({'font_color': 'red', 'bold': True})
    standard_format = workbook.add_format({'font_color': 'black'})

    column_labels = subMarket_report_new.columns.tolist()
    first_row_values = subMarket_report_new.iloc[0].tolist()

    worksheet.write_row(0, 0, column_labels, header1_format)
    worksheet.write_row(1, 0, first_row_values, header2_format)

    live_data_matrix = subMarket_report_new.iloc[1:].values.tolist()
    for row_idx, row_values in enumerate(live_data_matrix, start=2):
        worksheet.write_row(row_idx, 0, row_values, standard_format)
    
    workbook.close()
    return subMarket_report_new, CBRS


# =========================================================================
# RUN CELL DATA INGESTION PIPELINE
# =========================================================================
if __name__ == "__main__":
    subMarket = Select_subMarket()
    print(f'\n > Processing Data Preparation Engine for: {subMarket}...')

    # 1. Fetch layout schema targets from master criteria rules sheet
    reqd_parameters, group_script_name_dict, df_headers = Samsung_Parameters()

    # 2. Extract, clean, and deduplicate vertical source data files
    frames = loading_data_files(subMarket, reqd_parameters, group_script_name_dict)

    # 3. Consolidate tables horizontally and finalize styled spreadsheet
    if frames:
        subMarket_df, cbrs_df = merging_dataframes_to_single_dataframe(frames, df_headers, subMarket)

    end = time.ctime()
    print(f'\n   Process completed successfully!')
    print(f'\n   Safe to run the next script: CELL Part 2 Precheck.')
    print(f'\nStart: {start}\nEnd  : {end}')
