import xlwings as xw
import pandas as pd
import os
import json

BASE_DIR = r"D:\DB Merging Automation\DB Merging Automation"
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
MAPPER_PATH = os.path.join(BASE_DIR, "Column_Mapper.xlsx")

def test_discovery(file_path, mapping_key, mode="Export"):
    print(f"Testing discovery for {file_path} with key {mapping_key}...")
    
    # Mock enough of the engine logic
    df_map = pd.read_excel(MAPPER_PATH, sheet_name=mode)
    row = df_map[df_map['Key'] == mapping_key].iloc[0]
    
    mapping = {
        "source_sheet": row["Source_Sheet"],
        "date_config": {
            "type": row.get("Date_Type", "Split"),
            "date_col": row.get("Date_Col"),
            "year_col": row.get("Year_Col"),
            "month_col": row.get("Month_Col")
        }
    }
    
    date_cfg = mapping['date_config']
    method = date_cfg.get('type', 'Split')
    sheet_name = mapping['source_sheet']
    
    app = xw.App(visible=False)
    try:
        # Optimization check: Read only + No update links
        wb = app.books.open(file_path, read_only=True, update_links=False)
        source_sheet = wb.sheets[sheet_name]
        
        # Fresh unique temp name
        temp_name = "ST_Discovery_Debug"
        if temp_name in [s.name for s in wb.sheets]:
            wb.sheets[temp_name].delete()
        temp_sheet = wb.sheets.add(name=temp_name, visible=False)
        
        headers = source_sheet.range('1:1').value
        print(f"Headers found: {headers[:5]}...")

        if method == 'Direct':
            date_col = date_cfg['date_col']
            print(f"Direct Method: Looking for {date_col}")
            col_idx = headers.index(date_col) + 1
            col_letter = xw.utils.get_address(source_sheet.range(1, col_idx).address).split('$')[1]
            formula = f"=UNIQUE('{sheet_name}'!{col_letter}2:{col_letter}1048576)"
            temp_sheet.range('A1').formula = formula
            
            # Explicit calculation wait
            app.api.Calculate()
            
            raw_values = temp_sheet.range('A1:A100').value
            print(f"Raw values from formula: {raw_values[:10]}")
            
        else: # Split
            yr_col = date_cfg['year_col']
            mo_col = date_cfg['month_col']
            print(f"Split Method: Looking for {yr_col}, {mo_col}")
            yr_idx = headers.index(yr_col) + 1
            mo_idx = headers.index(mo_col) + 1
            yr_letter = xw.utils.get_address(source_sheet.range(1, yr_idx).address).split('$')[1]
            mo_letter = xw.utils.get_address(source_sheet.range(1, mo_idx).address).split('$')[1]
            
            temp_sheet.range('A1').formula = f"=UNIQUE('{sheet_name}'!{yr_letter}2:{yr_letter}1048576)"
            temp_sheet.range('B1').formula = f"=UNIQUE('{sheet_name}'!{mo_letter}2:{mo_letter}1048576)"
            
            app.api.Calculate()
            
            yrs = temp_sheet.range('A1:A100').value
            mos = temp_sheet.range('B1:B100').value
            print(f"Yrs found: {yrs[:10]}")
            print(f"Mos found: {mos[:10]}")

        wb.close()
        app.quit()
    except Exception as e:
        print(f"ERROR: {e}")
        if 'app' in locals(): app.quit()

if __name__ == "__main__":
    # Test with the massive file
    test_discovery(r"D:\DB Merging Automation\DB Merging Automation\data\base_files\Stress_Large_Excel_Limit.xlsx", "STRESS_MAX")
