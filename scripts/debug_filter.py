# -*- coding: utf-8 -*-
"""
Debug script to test _apply_date_filter with various date formats.
Run from the LAB DB Merging directory:
    .\\venv\\Scripts\\python.exe scripts\\debug_filter.py
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
from scripts.engine import ExcelEngine

def ok(cond):  return "PASS" if cond else "FAIL"

# ─── Synthetic test data covering all known formats ──────────────────────────
print("=" * 70)
print("TEST 1: SPLIT mode -- various year/month formats")
print("=" * 70)

df_split = pd.DataFrame({
    "Year":  ["2024",   "2024",   "2024",    2024,    2024.0,  "2025",  "2025"],
    "Month": ["Jan",    "Feb",    "january", 3,       4.0,     "May",   "Jun"],
    "Value": [100,      200,      300,       400,     500,     600,     700],
})
print("Input data:"); print(df_split.to_string(index=False))

mapping_split = {
    "date_config": {"type": "Split", "year_col": "Year", "month_col": "Month", "date_col": None}
}

engine = ExcelEngine.__new__(ExcelEngine)  # skip __init__

# Pairs from UI — year is numpy int64 (from Period.year), month is str
import numpy as np
test_pairs = [(np.int64(2024), "Jan"), (np.int64(2024), "Feb"),
              (np.int64(2024), "Mar"), (np.int64(2024), "Apr")]

result, err = engine._apply_date_filter(df_split, mapping_split, test_pairs, [])
print(f"\nFilter for year=2024, months=Jan/Feb/Mar/Apr")
print(f"Error  : '{err}'")
print(f"Rows   : {len(result)}")
print(result.to_string(index=False) if not result.empty else "(empty)")

expected_values = {100, 200, 300, 400, 500}
got_values = set(result["Value"].tolist()) if not result.empty else set()
print(f"\n[{ok(got_values == expected_values)}] expected {expected_values}, got {got_values}")
if got_values != expected_values:
    missing = expected_values - got_values
    extra   = got_values - expected_values
    if missing: print(f"  MISSING values: {missing}")
    if extra:   print(f"  EXTRA   values: {extra}")

# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("TEST 2: DIRECT mode -- date column as string dates")
print("=" * 70)

df_direct = pd.DataFrame({
    "Date":  ["2024-01-15", "2024-02-20", "2024-03-10", "2025-01-05", "2024-04-01"],
    "Value": [10, 20, 30, 40, 50],
})
print("Input data:"); print(df_direct.to_string(index=False))

mapping_direct = {
    "date_config": {"type": "Direct", "date_col": "Date", "year_col": None, "month_col": None}
}

test_pairs_d = [(np.int64(2024), "Jan"), (np.int64(2024), "Feb"), (np.int64(2024), "Mar")]
result_d, err_d = engine._apply_date_filter(df_direct, mapping_direct, test_pairs_d, [])
print(f"\nFilter for 2024 Jan/Feb/Mar")
print(f"Error  : '{err_d}'")
print(f"Rows   : {len(result_d)}")
print(result_d.to_string(index=False) if not result_d.empty else "(empty)")

expected_d = {10, 20, 30}
got_d = set(result_d["Value"].tolist()) if not result_d.empty else set()
print(f"\n[{ok(got_d == expected_d)}] expected {expected_d}, got {got_d}")

# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("TEST 3: SPLIT mode -- numeric month column (integers 1-12) coerced to str")
print("=" * 70)

df_num_mo = pd.DataFrame({
    "Year":  ["2024", "2024", "2024", "2024", "2024"],
    "Month": ["1",    "2",    "3",    "4",    "5"],   # dtype=str from calamine
    "Value": [10,     20,     30,     40,     50],
})
print("Input data:"); print(df_num_mo.to_string(index=False))

test_pairs_n = [(np.int64(2024), "Jan"), (np.int64(2024), "Feb"), (np.int64(2024), "Mar")]
result_n, err_n = engine._apply_date_filter(df_num_mo, mapping_split, test_pairs_n, [])
print(f"\nFilter for 2024 Jan/Feb/Mar (months stored as '1','2','3')")
print(f"Error  : '{err_n}'")
print(f"Rows   : {len(result_n)}")
print(result_n.to_string(index=False) if not result_n.empty else "(empty)")

expected_n = {10, 20, 30}
got_n = set(result_n["Value"].tolist()) if not result_n.empty else set()
print(f"\n[{ok(got_n == expected_n)}] expected {expected_n}, got {got_n}")

# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("TEST 4: SPLIT mode -- float year/month strings (e.g. '2024.0', '1.0')")
print("=" * 70)

df_float = pd.DataFrame({
    "Year":  ["2024.0", "2024.0", "2025.0"],
    "Month": ["1.0",    "2.0",    "1.0"],
    "Value": [10,       20,       30],
})
print("Input data:"); print(df_float.to_string(index=False))

result_f, err_f = engine._apply_date_filter(
    df_float, mapping_split, [(np.int64(2024), "Jan"), (np.int64(2024), "Feb")], [])
print(f"\nFilter for 2024 Jan/Feb (data has '2024.0','1.0' format)")
print(f"Error  : '{err_f}'")
print(f"Rows   : {len(result_f)}")
print(result_f.to_string(index=False) if not result_f.empty else "(empty)")

expected_f = {10, 20}
got_f = set(result_f["Value"].tolist()) if not result_f.empty else set()
print(f"\n[{ok(got_f == expected_f)}] expected {expected_f}, got {got_f}")

# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("TEST 5: Edge cases -- empty pairs, NaN rows")
print("=" * 70)

result_e, err_e = engine._apply_date_filter(df_split, mapping_split, [], [])
print(f"Empty pairs => rows: {len(result_e)}, err: '{err_e}'")
print(f"[{ok(len(result_e) == 0)}] Empty pairs should return 0 rows")

df_nan = df_split.copy()
df_nan.loc[len(df_nan)] = [None, None, 999]
result_nan, _ = engine._apply_date_filter(df_nan, mapping_split, [(np.int64(2024), "Jan")], [])
nan_leaked = 999 in (result_nan["Value"].tolist() if not result_nan.empty else [])
print(f"\nNaN row 999 leaked into results: {nan_leaked}")
print(f"[{ok(not nan_leaked)}] NaN rows should be excluded")

# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("TEST 6: SPLIT mode -- month stored as full name ('January','February')")
print("=" * 70)

df_fullname = pd.DataFrame({
    "Year":  ["2024",      "2024",      "2024"],
    "Month": ["January",   "February",  "March"],
    "Value": [10,          20,          30],
})
print("Input data:"); print(df_fullname.to_string(index=False))

result_fn, err_fn = engine._apply_date_filter(
    df_fullname, mapping_split, [(np.int64(2024), "Jan"), (np.int64(2024), "Feb")], [])
print(f"\nFilter for 2024 Jan/Feb (data has full names)")
print(f"Error  : '{err_fn}'")
print(f"Rows   : {len(result_fn)}")
print(result_fn.to_string(index=False) if not result_fn.empty else "(empty)")
expected_fn = {10, 20}
got_fn = set(result_fn["Value"].tolist()) if not result_fn.empty else set()
print(f"\n[{ok(got_fn == expected_fn)}] expected {expected_fn}, got {got_fn}")

print("\n" + "=" * 70)
print("ALL TESTS COMPLETE")
print("=" * 70)
