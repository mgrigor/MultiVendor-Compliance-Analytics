import os
import sys
import time
import xlsxwriter
import pandas as pd

start = time.ctime()
print('''
################################################################################
#                                                                              #
#                           Samsung ENB Prep & Merge v.1                       #
#                                                                              #
################################################################################
''')

directory = 'Raw Data'
tech = 'ENB'
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
    """Dynamically extracts parameters from the master blueprint sheet to build the base frame columns."""
    print('\n > Generating dataframe headers dynamically from blueprint...')
#    file = 'Samsung Required Parameters Master.xlsx'
    file = 'Demo_GPL_Samsung.xlsx'
    
    df = pd.read_excel(file, sheet_name='master', index_col=False, dtype=str)

    # Filter to isolate target blueprint parameters matching ENB Part 1 definitions
    enb_group = df.loc[(df['Technology'] == tech) & (df['Group'] == level)]
    reqd_parameters = []
    group_script_name_dict = {}
    groupby_script_name = enb_group.groupby('Group Script Name')
    
    for group_script, data in groupby_script_name:
        parameters = data[data['Group Script Name'] == group_script]['Parameter'].unique().tolist()
        group_script_name_dict.update({tuple(parameters): group_script})
        reqd_parameters.append(parameters)

    # Core site metadata tracking row layout
    labels = ['date', 'subMarket', 'market', 'enbname', 'enbid']
    headers_and_values = {label: '' for label in labels}
    
    # Map master targets to baseline columns
    param_mappings = dict(zip(enb_group['Parameter'], enb_group['Airwave Setting']))
    headers_and_values.update(param_mappings)
    
    df_headers = pd.DataFrame(headers_and_values, index=[0])
    df_headers = df_headers.map(lambda x: x.lower().strip() if isinstance(x, str) else x)
    
    return (reqd_parameters, group_script_name_dict, df_headers)

def loading_data_files_old(subMarket):
    """Loads CSV files, purges legacy sync data cleanly, and runs QCI compression loops."""
    print('\n > Loading raw data files for subMarket:', subMarket)
    fpath = directory + '/' + subMarket + '/' + tech + '/' + level
    
    # Path resilience fallback check
    if not os.path.exists(fpath):
        fpath = directory + '/' + subMarket + '/' + tech
        
    files = os.listdir(fpath)
    data_files = [file for file in files if file.endswith('.csv')]
    
    frames = []
    string_qci = 'enb-function_qci-backhaul-bandwidth-info_qci'
    
    for file in data_files:
        df = pd.read_csv(fpath + '/' + file, dtype=str, index_col=False)
        df['enbid'] = df['enbid'].astype(str).str.zfill(6)
        
        SubMarket_check = df['subMarket'].unique().tolist()[0].split('-')[-1]
        print(f"     Staging File: {file} (Market Check: {SubMarket_check})")

        # Strip system index counters and individual vertical QCI labels
        cols_to_keep = [c for c in df.columns if c != string_qci and not c.startswith('paramId')]
        df = df[cols_to_keep]
        
        columns = df.columns.tolist()
        parameters = columns[5:]
        
        # Create a flag tracking modernized status
        df['is_modernized'] = df['enbname'].str.contains('4G5G', na=False, case=False)
                   
        # ─── THE TRANSFORMATION ENGINE ────────────────────────────────────────
        if parameters and parameters[0].startswith('backhaul'):
            print('     ↳Backhaul dataset detected. Purging legacy profiles & compressing QCI (0-9)...')
            
            # CRITICAL FILTER RULE: Vectorized legacy purge. If a 4G5G site exists, drop its legacy rows.
            # (This vectorized approach also safely bypasses the deprecated .apply() warning)
            has_modernized = df.groupby('enbid')['is_modernized'].transform('any')
            df = df[df['is_modernized'] | ~has_modernized].reset_index(drop=True)
            
            Data = []
            grouped = df.groupby('enbid')
            for enbid, data in grouped:
                rowData = {}
                for label in columns[:5]:
                    rowData.update({label: data[label].unique().tolist()[0]})
                for parameter in parameters:
                    # Extracts unique config thresholds across all vertical rows
                    conf_values = data[parameter].dropna().unique().tolist()
                    rowData.update({parameter: conf_values[0] if len(conf_values) == 1 else "/".join(conf_values)})
                Data.append(pd.DataFrame(rowData, index=[0]))
            frame = pd.concat(Data, ignore_index=True)
            frames.append(frame)
        else:
            # CRITICAL FILTER RULE: Standard files. Put 4G5G first, then drop duplicates to discard legacy data.
            df.sort_values(by=['enbid', 'is_modernized'], ascending=[True, False], inplace=True)
            
            init_rows = df.shape[0]
            df.drop_duplicates(subset=['enbid'], keep='first', inplace=True)
            final_rows = df.shape[0]
            df.drop(columns=['is_modernized'], inplace=True)
            
            if init_rows != final_rows:
                print(f"     ↳Dropped {init_rows - final_rows} legacy sync data rows from memory.")
            frames.append(df)
            
    return frames

def merging_dataframes_to_single_dataframe(frames, df_headers, subMarket):
    """Marries pristine, pre-deduplicated data tables horizontally and exports via xlsxwriter."""
    print('\n > Merging groups of single-row dataframes into a single DataFrame...')
    
    # Sequential horizontal outer join on the clean, unique node profiles
    merged_df = frames[0]
    join_keys = ['date', 'subMarket', 'market', 'enbid', 'enbname']
    
    for frame in frames[1:]:
        parameters = frame.columns.tolist()[5:]
        parameters.insert(0, 'enbid')
        merged_df = pd.merge(merged_df, frame[parameters], on='enbid', how='outer')
        
    # Enforce blueprint column layout ordering 
    headers = df_headers.columns.tolist()
    for col in headers:
        if col not in merged_df.columns:
            merged_df[col] = ""
            
    merged_df = merged_df[headers]

    print('   ↳Scrubbing non-production "GROW" testing cells...')
    final_live_df = merged_df[~merged_df['enbname'].str.lower().str.startswith('grow', na=False)]   

    # Combine the template configuration rules blueprint onto row 0
    subMarket_report_new = pd.concat([df_headers, final_live_df], ignore_index=True)

    # ─── XLSXWRITER STRUCTURAL EXPORT ENGINE ──────────────────────────────────
    fpath_out = directory + '/' + subMarket
    os.makedirs(fpath_out, exist_ok=True)
    filename = f"{fpath_out}/Samsung_{subMarket}_ENB.xlsx"
    print(f'\n > Saving finalized consolidated report to Excel via xlsxwriter: {filename}')
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet("ENB Master Baseline")

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
    return subMarket_report_new, filename
        
# =========================================================================
# RUN ENB DATA INGESTION PIPELINE
# =========================================================================
if __name__ == "__main__":
    subMarket = Select_subMarket()
    print(f'\n > Processing Data Preparation Engine for: {subMarket}...')

    # 1. Fetch parameters from spreadsheet master layout
    reqd_parameters, group_script_name_dict, df_headers = Samsung_Parameters()

    # 2. Extract and transform vertical data dumps (Legacy rows dropped symmetrically)
    frames = loading_data_files_old(subMarket)

    # 3. Marry and flatten structures side-by-side horizontally
    if frames:
        subMarket_df, file = merging_dataframes_to_single_dataframe(frames, df_headers, subMarket)

    end = time.ctime()
    print(f'\n  Process completed successfully!')
    print(f'\n  Safe to run the next script: Samsung CELL Part 1 Precheck.')
    print('')      
    print('Start:', start)
    print('End  :', end)
