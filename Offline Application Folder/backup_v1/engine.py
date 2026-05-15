import xlwings as xw
import pandas as pd
import json
import os

class ExcelEngine:
    def __init__(self, mapper_path, config_path, mode="Export"):
        self.mode = mode
        with open(config_path, 'r') as f:
            self.config = json.load(f)
            
        # Select correct dashboard based on mode
        if self.mode == "Import":
            self.dashboard_path = self.config['import_dashboard_path']
        else:
            self.dashboard_path = self.config['export_dashboard_path']

        # Load mappings from specific sheet in Excel Matrix
        try:
            df_map = pd.read_excel(mapper_path, sheet_name=self.mode)
            self.mappings = {}
            for _, row in df_map.iterrows():
                key = row["Key"]
                self.mappings[key] = {
                    "source_sheet": row["Source_Sheet"],
                    "target_sheet": row["Target_Sheet"],
                    "date_config": {
                        "type": row.get("Date_Type", "Split"),
                        "date_col": row.get("Date_Col"),
                        "year_col": row.get("Year_Col"),
                        "month_col": row.get("Month_Col")
                    },
                    "column_mapping": {
                        # Key is Source_Header (cell value), Value is Dashboard_Header (column name)
                        row[col]: col for col in df_map.columns[7:] if pd.notna(row[col]) and pd.notna(row[col])
                    }
                }
        except Exception as e:
            self.mappings = {}
            print(f"Error loading {self.mode} mappings: {e}")

    def get_sheet_names(self, file_path):
        """Quickly get sheet names without full app visibility."""
        with xw.App(visible=False) as app:
            wb = app.books.open(file_path)
            names = [s.name for s in wb.sheets]
            wb.close()
            return names

    def merge_file(self, base_file_path, country_key, target_month, target_year):
        """The core surgical merge logic."""
        mapping = self.mappings.get(country_key)
        if not mapping:
            return False, f"No mapping found for {country_key}"

        dashboard_path = self.dashboard_path
        
        try:
            # Start Excel App
            app = xw.App(visible=False)
            wb_dash = app.books.open(dashboard_path)
            wb_base = app.books.open(base_file_path)
            
            # 1. Read Data from Base
            source_sheet = wb_base.sheets[mapping['source_sheet']]
            # Convert to DataFrame for easy filtering
            df_base = source_sheet.used_range.options(pd.DataFrame, index=False).value
            
            # 2. Filter for Month/Year using centralized helper
            df_filtered, err = self._apply_date_filter(df_base, mapping, target_month, target_year)
            
            if not err == "" or df_filtered.empty:
                wb_base.close()
                wb_dash.close()
                app.quit()
                return False, f"{err if err else 'No data found'} for {target_month} {target_year} in {os.path.basename(base_file_path)}"

            # 3. Map Columns Dynamically based on Dashboard Headers
            target_sheet_name = mapping['target_sheet']
            target_sheet = wb_dash.sheets[target_sheet_name]
            
            # Fetch actual headers from Dashboard to ensure perfect alignment
            dash_row = target_sheet.range('1:1').value
            if not isinstance(dash_row, list): dash_row = [dash_row]
            dash_headers = [h for h in dash_row if h is not None]
            
            # Create a blank template matching the Dashboard structure
            df_to_append = pd.DataFrame(columns=dash_headers)
            
            # Fill using strict mapping: Target columns must exactly match dashboard headers
            for src_col, target_col in mapping['column_mapping'].items():
                if src_col in df_filtered.columns and target_col in dash_headers:
                    df_to_append[target_col] = df_filtered[src_col]
            
            # Optional: If 'Date' is a column in dashboard but wasn't explicitly mapped 
            # and we are in 'Split' mode, we can auto-fill it for convenience
            if "Date" in dash_headers and df_to_append["Date"].isnull().all():
                if mapping['date_config']['type'] == 'Split':
                    yr_col = mapping['date_config']['year_col']
                    mo_col = mapping['date_config']['month_col']
                    if yr_col in df_filtered.columns and mo_col in df_filtered.columns:
                        # Construct "01-MMM-YYYY" string
                        df_to_append["Date"] = "01-" + df_filtered[mo_col].astype(str) + "-" + df_filtered[yr_col].astype(str)
            
            # 4. Append to Dashboard (with Empty Row + Repeating Headers)
            last_row = target_sheet.range('A' + str(target_sheet.cells.last_cell.row)).end('up').row
            # If dashboard is currently empty (only row 1 has contents), we start at row 3
            start_row = last_row + 2
            
            # Write Bold Headers
            header_cell = target_sheet.range(f"A{start_row}")
            header_cell.value = [df_to_append.columns.tolist()]
            header_cell.expand('right').api.Font.Bold = True
            
            # Write Data starting directly after the new header row
            target_sheet.range(f"A{start_row + 1}").value = df_to_append.values
            
            # 5. Slicer & Pivot Protection (Disabled)
            # last_col_letter = target_sheet.range(1, len(dash_headers)).address.split('$')[1]
            
            wb_dash.save()
            wb_base.close()
            wb_dash.close()
            app.quit()
            return True, f"Successfully merged {len(df_to_append)} rows (with headers)."
            
        except Exception as e:
            if 'app' in locals(): app.quit()
            return False, str(e)

    def merge_files_batch(self, tasks):
        """Processes multiple merge tasks in a single Excel session efficiently."""
        results = []
        app = None
        try:
            app = xw.App(visible=False)
            wb_dash = app.books.open(self.dashboard_path)
            
            for task in tasks:
                base_file_path = task['path']
                country_key = task['key']
                target_month = task['month']
                target_year = task['year']
                
                try:
                    mapping = self.mappings.get(country_key)
                    if not mapping:
                        results.append({"file": os.path.basename(base_file_path), "status": "Failed", "msg": f"No mapping for {country_key}"})
                        continue
                        
                    wb_base = app.books.open(base_file_path)
                    # 1. Read Data
                    source_sheet = wb_base.sheets[mapping['source_sheet']]
                    df_base = source_sheet.used_range.options(pd.DataFrame, index=False).value
                    
                    # 2. Filter
                    df_filtered, err = self._apply_date_filter(df_base, mapping, target_month, target_year)
                    
                    if not err == "" or df_filtered.empty:
                        results.append({"file": os.path.basename(base_file_path), "status": "Failed", "msg": f"{err if err else 'No data for date'}"})
                        wb_base.close()
                        continue
                    
                    # 3. Map Dynamically based on Dashboard Headers
                    target_sheet = wb_dash.sheets[mapping['target_sheet']]
                    dash_row = target_sheet.range('1:1').value
                    if not isinstance(dash_row, list): dash_row = [dash_row]
                    dash_headers = [h for h in dash_row if h is not None]
                    
                    # Create aligned DataFrame
                    df_to_append = pd.DataFrame(columns=dash_headers)
                    
                    # Fill using strict mapping
                    for src_col, target_col in mapping['column_mapping'].items():
                        if src_col in df_filtered.columns and target_col in dash_headers:
                            df_to_append[target_col] = df_filtered[src_col]
                    
                    # Auto-fill Date if missing
                    if "Date" in dash_headers and df_to_append["Date"].isnull().all():
                        if mapping['date_config']['type'] == 'Split':
                            yr_col = mapping['date_config']['year_col']
                            mo_col = mapping['date_config']['month_col']
                            if yr_col in df_filtered.columns and mo_col in df_filtered.columns:
                                df_to_append["Date"] = "01-" + df_filtered[mo_col].astype(str) + "-" + df_filtered[yr_col].astype(str)
                            
                    # 4. Append with Spacing + Headers
                    last_row = target_sheet.range('A' + str(target_sheet.cells.last_cell.row)).end('up').row
                    start_row = last_row + 2
                    
                    # Write Bold Headers
                    header_cell = target_sheet.range(f"A{start_row}")
                    header_cell.value = [df_to_append.columns.tolist()]
                    header_cell.expand('right').api.Font.Bold = True
                    
                    # Write Data
                    target_sheet.range(f"A{start_row + 1}").value = df_to_append.values
                    
                    wb_base.close()
                    results.append({"file": os.path.basename(base_file_path), "status": "Success", "msg": f"Merged {len(df_to_append)} rows"})
                    
                except Exception as e:
                    results.append({"file": os.path.basename(base_file_path), "status": "Failed", "msg": str(e)})
            
            wb_dash.save()
            wb_dash.close()
            app.quit()
            
        except Exception as e:
            if app: app.quit()
            return [{"file": "Batch Process", "status": "Error", "msg": f"Critical Error: {str(e)}"}]
            
        return results

    def get_available_dates(self, file_path, mapping_key):
        """Discovers unique years and months using native Office 365 =UNIQUE() formulas for instant results."""
        mapping = self.mappings.get(mapping_key)
        if not mapping:
            return {"years": [], "months": []}
        
        date_cfg = mapping['date_config']
        method = date_cfg.get('type', 'Split')
        sheet_name = mapping['source_sheet']
        app = None
        
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            
            # Optimization: Open as Read Only for speed
            wb = app.books.open(file_path, read_only=True, update_links=False)
            source_sheet = wb.sheets[sheet_name]
            
            # 1. Create a temporary sheet for calculations
            temp_name = "ST_Discovery_Temp"
            try:
                if temp_name in [s.name for s in wb.sheets]:
                    wb.sheets[temp_name].delete()
            except: pass
            
            # Corrected: Sheets.add() does not support visible=False
            temp_sheet = wb.sheets.add(name=temp_name)
            # You can hide it via API if necessary: temp_sheet.api.Visible = False
            
            # Fetch all headers to find positions
            headers = source_sheet.range('1:1').value
            
            found_data = {"years": [], "months": []}

            if method == 'Direct':
                date_col = date_cfg['date_col']
                if date_col in headers:
                    col_idx = headers.index(date_col) + 1
                    col_letter = source_sheet.range(1, col_idx).address.split('$')[1]
                    
                    # Native O365 Formula - Instant processing of up to 1M rows
                    formula = f"=UNIQUE('{sheet_name}'!{col_letter}2:{col_letter}1048576)"
                    temp_sheet.range('A1').formula = formula
                    
                    # Ensure Excel calculation is done
                    app.calculate()
                    
                    # Read the spilled results
                    # (Reading a small fixed block like A1:A100 is safer and faster than expanding)
                    raw_values = temp_sheet.range('A1:A100').value
                    # Filter out Nones
                    clean_values = [v for v in raw_values if v is not None and v != 0]
                    
                    s = pd.Series(clean_values).dropna()
                    dates = pd.to_datetime(s, errors='coerce').dropna()
                    found_data["years"] = sorted(dates.dt.year.unique().tolist())
                    month_num_map = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun", 
                                     7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
                    found_data["months"] = [month_num_map[m] for m in sorted(dates.dt.month.unique().tolist())]
            
            else: # Split method
                yr_col = date_cfg['year_col']
                mo_col = date_cfg['month_col']
                if yr_col in headers and mo_col in headers:
                    yr_idx = headers.index(yr_col) + 1
                    mo_idx = headers.index(mo_col) + 1
                    yr_letter = source_sheet.range(1, yr_idx).address.split('$')[1]
                    mo_letter = source_sheet.range(1, mo_idx).address.split('$')[1]
                    
                    # Unique of two columns at once
                    temp_sheet.range('A1').formula = f"=UNIQUE('{sheet_name}'!{yr_letter}2:{yr_letter}1048576)"
                    temp_sheet.range('B1').formula = f"=UNIQUE('{sheet_name}'!{mo_letter}2:{mo_letter}1048576)"
                    
                    app.calculate()
                    
                    yrs = [v for v in temp_sheet.range('A1:A100').value if v is not None and v != 0]
                    mos = [v for v in temp_sheet.range('B1:B100').value if v is not None and v != 0]
                    
                    found_data["years"] = sorted([int(y) for y in set(yrs)])
                    months_found = list(set(mos))
                    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                    found_data["months"] = sorted(months_found, key=lambda x: month_order.index(x) if x in month_order else 99)

            # Cleanup
            temp_sheet.delete()
            wb.close()
            app.quit()
            return found_data

        except Exception as e:
            if app: 
                try: app.quit()
                except: pass
            print(f"Native discovery error: {e}")
            return {"years": [], "months": []}

    def _apply_date_filter(self, df, mapping, target_months, target_years):
        """Filters the dataframe based on lists of years and months."""
        date_cfg = mapping['date_config']
        method = date_cfg.get('type', 'Split')

        # If target_months is actually a list of (Year, Month) tuples (passed from new UI)
        # then we use that for precise filtering across multiple years.
        if isinstance(target_months, list) and len(target_months) > 0 and isinstance(target_months[0], tuple):
            valid_pairs = target_months
        else:
            # Fallback for old list/scalar calls: construct all possible pairs
            if not isinstance(target_months, list): target_months = [target_months]
            if not isinstance(target_years, list): target_years = [target_years]
            valid_pairs = [(y, m) for y in target_years for m in target_months]

        try:
            if method == 'Direct':
                date_col = date_cfg.get('date_col')
                if not date_col or date_col not in df.columns:
                    return pd.DataFrame(), f"Date col '{date_col}' missing"
                
                temp_date = pd.to_datetime(df[date_col], errors='coerce')
                month_map = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6, "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
                
                # Construct vectorized conditions for all valid pairs
                conditions = None
                for yr, mo_str in valid_pairs:
                    mo_num = month_map.get(mo_str)
                    cond = (temp_date.dt.year == yr) & (temp_date.dt.month == mo_num)
                    conditions = cond if conditions is None else (conditions | cond)
                
                return df[conditions] if conditions is not None else pd.DataFrame(), ""
            
            else: # Split method
                yr_col = date_cfg.get('year_col')
                mo_col = date_cfg.get('month_col')
                if not yr_col or yr_col not in df.columns or not mo_col or mo_col not in df.columns:
                    return pd.DataFrame(), f"Split cols '{yr_col}'/'{mo_col}' missing"
                
                # Check for (Year, Month) pair membership
                # We normalize Year to an integer string (removing any .0 from floats)
                # and strip any hidden spaces from the Month string.
                def normalize_yr(val):
                    try: return str(int(float(val))) if pd.notna(val) else ""
                    except: return str(val)

                temp_yr = df[yr_col].apply(normalize_yr)
                temp_mo = df[mo_col].astype(str).str.strip()
                temp_s = temp_yr + "_" + temp_mo
                
                valid_set = {f"{y}_{m}" for y, m in valid_pairs}
                mask = temp_s.isin(valid_set)
                return df[mask], ""
        except Exception as e:
            return pd.DataFrame(), f"Filter error: {str(e)}"
    def _update_pivots_and_slicers(self, wb, new_range_address):
        """Disconnect slicers, update pivot source, reconnect."""
        # This uses the underlying win32com API through xlwings (.api)
        try:
            slicer_connections = []
            
            # 1. Store and Disconnect Slicers
            for sc in wb.api.SlicerCaches:
                for pt in sc.PivotTables:
                    slicer_connections.append((sc, pt))
                    # Note: We don't always need to disconnect if we use ChangePivotCache
                    # But the user specifically requested it.
                # In VBA: sc.PivotTables.RemovePivotTable(pt)
            
            # 2. Update Pivot Ranges
            for sheet in wb.sheets:
                for pt in sheet.api.PivotTables():
                    # Update SourceData
                    # SourceData must be in R1C1 format usually for ChangePivotCache
                    pt.ChangePivotCache(wb.api.PivotCaches().Create(SourceType=1, SourceData=new_range_address))
            
            # 3. Refresh All
            wb.api.RefreshAll()
            
        except Exception as e:
            print(f"Pivot update error: {e}")

if __name__ == "__main__":
    # Test logic
    engine = ExcelEngine('D:/DB Merging Automation/DB Merging Automation/mapping_registry.json', 'D:/DB Merging Automation/DB Merging Automation/config.json')
    # success, msg = engine.merge_file('...', 'Argentina_MAB_Export', 'Mar', 2024)
    # print(msg)
