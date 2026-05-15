# LAB DB Merging Application — Complete Documentation

This document explains the architecture, the technology stack, and the exact step-by-step logic of how the LAB DB Merging application works. 

---

## 1. What does this application do?
The **LAB DB Merging Application** is a high-performance tool designed to take monthly data from multiple, inconsistently formatted Excel "Base Files" and surgically merge them into central "Master Dashboards" (Import or Export). 

Its main value proposition is **formatting preservation**. Instead of destroying the target dashboard or converting everything to plain CSV data, it carefully appends new rows into the existing dashboard while keeping all Pivot Tables, charts, and Excel formatting completely intact.

---

## 2. Core Libraries & Technologies Used

The tool utilizes a carefully selected stack of Python libraries to achieve both speed and surgical precision:

| Library | Role in Application | Why we use it |
|---|---|---|
| **Streamlit** | Frontend UI | Builds the web-based interactive interface (`app.py`). It handles file selections, batch queues, buttons, and user notifications. |
| **Pandas** | Data Processing | The core data engine. It creates DataFrames to filter rows, map headers dynamically, and hold the extracted monthly data in memory before writing. |
| **python-calamine** | Ultra-fast Reading Engine | Used specifically to *read* massive Excel files. Calamine is written in Rust and is exceptionally fast. It allows the tool to scan large base files and extract dates almost instantly without opening Excel. |
| **xlwings** | Surgical Writing Engine | Used exclusively to *write* data to the Dashboard. Unlike Pandas (which destroys existing Excel files when saving), `xlwings` connects directly to the Microsoft Excel application in the background (via COM). It behaves exactly like a human copying and pasting data, preserving all Pivot Tables, colors, and layout. |
| **tkinter** | Native Dialogs | A built-in Python UI library used to open native Windows "Browse File" and "Browse Folder" popups. |
| **pywin32** | Windows COM Bridge | A background dependency that allows `xlwings` to talk directly to the Windows Microsoft Excel `.exe` application. |

---

## 3. How the Application Works (Step-by-Step)

The application logic is split into two main components: the UI (`app.py`) and the Engine (`scripts/engine.py`).

### Step A: Configuration & Mapping
1. **`config.json`**: Stores the file paths (where your base files are, where your dashboards are).
2. **`Column_Mapper.xlsx`**: The rulebook. The application reads this file to understand how to translate column headers from the Base File (e.g., "Importing Country") to the Dashboard headers (e.g., "Country").

### Step B: Discovery (UI Phase)
When you select a Base File and a mapping from the dropdown:
1. The UI uses the `calamine` engine to quickly scan the Base file.
2. It identifies the dates available in the file (handling both direct Date formats and Split Year/Month formats).
3. The UI presents a Calendar so the user can select exactly which months they want to merge.

### Step C: The Merge Execution (`engine.py`)
When you click **"Run Single Merge"** or **"Process Batch Queue"**, the following sequence happens:

1. **Ultra-fast Read (`calamine`)**: 
   The engine reads the selected Base File into a Pandas DataFrame. It reads all data as strings or raw numbers to prevent Excel from corrupting dates.
2. **Date Filtering**:
   The `_apply_date_filter` function strips down the massive dataset so only the exact months and years you selected remain in memory.
3. **Column Mapping Alignment**:
   The engine opens the target Dashboard in the background using `xlwings`. It scans Row 1 of the Dashboard to get the exact header names. It then renames and aligns the Base File data to perfectly match the Dashboard layout. Missing columns are ignored; unmapped target columns remain blank.
4. **Locating the Drop Zone**:
   The engine uses `target_sheet.used_range.last_cell.row` to find the absolute bottom of the existing data in the Dashboard. This ensures it never overwrites existing rows.
5. **Appending & Formatting**:
   * It skips one blank row below the existing data.
   * It writes the **Base File Headers** in the new row and copies the header formatting (bold, colors) from Row 1 of the Dashboard.
   * It pastes the **Data Rows** immediately beneath the headers.
   * Finally, it runs `ClearFormats()` on the newly pasted data rows to ensure no header styling "bleeds" into the data cells.
6. **Save & Close**:
   The engine saves the Dashboard and forcefully quits the background Excel process, returning a success message to the Streamlit UI.

---

## 4. Key Architectural Decisions
* **Hybrid Excel Architecture**: Reading using `calamine` + Writing using `xlwings`. This hybrid approach solves the classic Python problem where fast libraries (like openpyxl) destroy formatting, and formatting-safe libraries (xlwings) are too slow for reading huge files.
* **Batch Processing**: The batch queue logic initializes the `xlwings` background Excel application only **once**. It merges 10 files in a single session and saves at the very end. This drastically reduces processing time compared to opening and closing Excel 10 times.
