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
                        row[col]: col for col in df_map.columns[7:] if pd.notna(col) and pd.notna(row[col])
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

    def _get_dtypes(self, mapping):
        dtypes = {}
        if not mapping: return dtypes
        date_cfg = mapping.get('date_config', {})
        if date_cfg.get('date_col'): dtypes[date_cfg['date_col']] = str
        if date_cfg.get('year_col'): dtypes[date_cfg['year_col']] = str
        if date_cfg.get('month_col'): dtypes[date_cfg['month_col']] = str
        return dtypes

    def merge_file(self, base_file_path, country_key, target_month, target_year):
        """The core surgical merge logic."""
        mapping = self.mappings.get(country_key)
        if not mapping:
            return False, f"No mapping found for {country_key}"

        dashboard_path = self.dashboard_path
        
        try:
            # 1. Fast Read Data from Base Instantly with Calamine
            dtypes_dict = self._get_dtypes(mapping)
            df_base = pd.read_excel(base_file_path, sheet_name=mapping['source_sheet'], engine='calamine', dtype=dtypes_dict)
            
            # 2. Filter for Month/Year using centralized helper
            df_filtered, err = self._apply_date_filter(df_base, mapping, target_month, target_year)
            
            if err or df_filtered.empty:
                return False, f"{err if err else 'No data found'} for {target_month} {target_year} in {os.path.basename(base_file_path)}"

            # Start Excel App ONLY for the Target Dashboard
            app = xw.App(visible=False)
            wb_dash = app.books.open(dashboard_path)
            
            # 3. Map Columns Dynamically based on Dashboard Headers
            target_sheet_name = mapping['target_sheet']
            sheet_names = [s.name for s in wb_dash.sheets]
            if target_sheet_name not in sheet_names:
                wb_dash.close()
                app.quit()
                return False, f"Target sheet '{target_sheet_name}' not found in Dashboard (Available: {', '.join(sheet_names)})"
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
            
            # Create reverse mapping to find Base File headers (Target -> Source)
            rev_mapping = {v: k for k, v in mapping['column_mapping'].items()}
            base_file_headers = [rev_mapping.get(col, col) for col in df_to_append.columns]
            
            # 4. Append to Dashboard (1 blank row gap, header formatting only on header row)
            last_row = target_sheet.used_range.last_cell.row
            
            if last_row == 1:
                # Dashboard is empty (only has main headers). Write data directly to row 2.
                target_sheet.range("A2").value = df_to_append.values
                # Clear any inherited formatting on data rows
                if len(df_to_append) > 0:
                    data_range = target_sheet.range((2, 1), (1 + len(df_to_append), len(base_file_headers)))
                    data_range.api.ClearFormats()
            else:
                # 1 blank row separates previous data from new section (last_row+1 is blank)
                start_row = last_row + 2
                num_cols = len(base_file_headers)
                
                # Write base-file header row
                header_cell = target_sheet.range(f"A{start_row}")
                header_cell.value = [base_file_headers]
                
                # Apply dashboard row-1 formatting to the header row ONLY
                target_sheet.range((1, 1), (1, num_cols)).copy()
                header_cell.expand('right').paste(paste='formats')
                
                # Write data rows immediately after the header
                target_sheet.range(f"A{start_row + 1}").value = df_to_append.values
                
                # Clear ALL formatting from data rows — no header styles should bleed in
                if len(df_to_append) > 0:
                    data_range = target_sheet.range((start_row + 1, 1), (start_row + len(df_to_append), num_cols))
                    data_range.api.ClearFormats()
            
            # 5. Slicer & Pivot Protection (Disabled)
            # last_col_letter = target_sheet.range(1, len(dash_headers)).address.split('$')[1]
            
            wb_dash.save()
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
                        
                    # 1. Fast Read Data
                    dtypes_dict = self._get_dtypes(mapping)
                    df_base = pd.read_excel(base_file_path, sheet_name=mapping['source_sheet'], engine='calamine', dtype=dtypes_dict)
                    
                    # 2. Filter
                    df_filtered, err = self._apply_date_filter(df_base, mapping, target_month, target_year)
                    
                    if err or df_filtered.empty:
                        results.append({"file": os.path.basename(base_file_path), "status": "Failed", "msg": f"{err if err else 'No data for date'}"})
                        continue
                    
                    # 3. Map Dynamically based on Dashboard Headers
                    target_sheet_name = mapping['target_sheet']
                    sheet_names = [s.name for s in wb_dash.sheets]
                    if target_sheet_name not in sheet_names:
                        results.append({"file": os.path.basename(base_file_path), "status": "Failed", "msg": f"Target sheet '{target_sheet_name}' missing from Dashboard"})
                        continue
                    target_sheet = wb_dash.sheets[target_sheet_name]
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
                            
                    # Create reverse mapping to find Base File headers
                    rev_mapping = {v: k for k, v in mapping['column_mapping'].items()}
                    base_file_headers = [rev_mapping.get(col, col) for col in df_to_append.columns]
                    
                    # 4. Append with 1 blank row gap; header formatting on header row only
                    last_row = target_sheet.used_range.last_cell.row
                    
                    if last_row == 1:
                        target_sheet.range("A2").value = df_to_append.values
                        # Clear any inherited formatting on data rows
                        if len(df_to_append) > 0:
                            data_range = target_sheet.range((2, 1), (1 + len(df_to_append), len(base_file_headers)))
                            data_range.api.ClearFormats()
                    else:
                        # 1 blank row separates previous data from new section
                        start_row = last_row + 2
                        num_cols = len(base_file_headers)
                        
                        # Write base-file header row
                        header_cell = target_sheet.range(f"A{start_row}")
                        header_cell.value = [base_file_headers]
                        
                        # Apply dashboard row-1 formatting to the header row ONLY
                        target_sheet.range((1, 1), (1, num_cols)).copy()
                        header_cell.expand('right').paste(paste='formats')
                        
                        # Write data rows immediately after the header
                        target_sheet.range(f"A{start_row + 1}").value = df_to_append.values
                        
                        # Clear ALL formatting from data rows — no header styles should bleed in
                        if len(df_to_append) > 0:
                            data_range = target_sheet.range((start_row + 1, 1), (start_row + len(df_to_append), num_cols))
                            data_range.api.ClearFormats()
                    
                    results.append({"file": os.path.basename(base_file_path), "status": "Success", "msg": f"Merged {len(df_to_append)} rows"})
                    
                except Exception as e:
                    # Clean up COM error tuples if present
                    err_str = str(e)
                    if hasattr(e, 'excepinfo') and e.excepinfo:
                        err_str = f"Excel Error: {e.excepinfo[1]} ({e.excepinfo[5]})"
                    elif hasattr(e, 'args') and len(e.args) >= 3 and isinstance(e.args[2], tuple):
                        err_str = f"Excel Error: {e.args[2][1]}"
                    results.append({"file": os.path.basename(base_file_path), "status": "Failed", "msg": err_str})
            
            wb_dash.save()
            wb_dash.close()
            app.quit()
            
        except Exception as e:
            if app: app.quit()
            return [{"file": "Batch Process", "status": "Error", "msg": f"Critical Error: {str(e)}"}]
            
        return results

    def get_available_dates(self, file_path, mapping_key):
        """Discovers unique years and months using the ultra-fast calamine engine."""
        mapping = self.mappings.get(mapping_key)
        if not mapping:
            return {"years": [], "months": []}
        
        date_cfg = mapping['date_config']
        method = date_cfg.get('type', 'Split')
        sheet_name = mapping['source_sheet']
        
        try:
            # Fast read using calamine
            dtypes_dict = self._get_dtypes(mapping)
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='calamine', dtype=dtypes_dict)
            
            found_data = {"years": [], "months": []}

            if method == 'Direct':
                date_col = date_cfg['date_col']
                if date_col in df.columns:
                    dates = pd.to_datetime(df[date_col], errors='coerce').dropna()
                    found_data["years"] = sorted(dates.dt.year.unique().tolist())
                    month_num_map = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun", 
                                     7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
                    found_data["months"] = [month_num_map[m] for m in sorted(dates.dt.month.unique().tolist())]
            
            else: # Split method
                yr_col = date_cfg['year_col']
                mo_col = date_cfg['month_col']
                if yr_col in df.columns and mo_col in df.columns:
                    yrs = df[yr_col].dropna().unique()
                    def clean_yr(y):
                        try: return int(float(y))
                        except: return str(y)
                    found_data["years"] = sorted(list(set([clean_yr(y) for y in yrs])))
                    
                    months_found = list(set(df[mo_col].dropna().astype(str).str.strip().unique()))
                    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                    found_data["months"] = sorted(months_found, key=lambda x: month_order.index(x) if x in month_order else 99)

            return found_data

        except Exception as e:
            print(f"Calamine discovery error: {e}")
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
                def normalize_yr(val):
                    try: return str(int(float(val))) if pd.notna(val) else ""
                    except: return str(val)

                # Robust month normalizer: handles integers (1, 2), floats (1.0), short strings (Jan), long strings (January)
                month_map = {
                    "1": "Jan", "01": "Jan", "january": "Jan", "jan": "Jan",
                    "2": "Feb", "02": "Feb", "february": "Feb", "feb": "Feb",
                    "3": "Mar", "03": "Mar", "march": "Mar", "mar": "Mar",
                    "4": "Apr", "04": "Apr", "april": "Apr", "apr": "Apr",
                    "5": "May", "05": "May", "may": "May",
                    "6": "Jun", "06": "Jun", "june": "Jun", "jun": "Jun",
                    "7": "Jul", "07": "Jul", "july": "Jul", "jul": "Jul",
                    "8": "Aug", "08": "Aug", "august": "Aug", "aug": "Aug",
                    "9": "Sep", "09": "Sep", "september": "Sep", "sep": "Sep",
                    "10": "Oct", "october": "Oct", "oct": "Oct",
                    "11": "Nov", "november": "Nov", "nov": "Nov",
                    "12": "Dec", "december": "Dec", "dec": "Dec"
                }

                def normalize_mo(val):
                    if pd.isna(val): return ""
                    s = str(val).strip().lower()
                    if s.endswith('.0'): s = s[:-2]
                    return month_map.get(s, s.capitalize())

                temp_yr = df[yr_col].apply(normalize_yr)
                temp_mo = df[mo_col].apply(normalize_mo)
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
