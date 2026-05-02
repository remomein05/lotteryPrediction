# Backtest Automation Plan

## Objective
Automate the manual extraction and cross-referencing steps currently required to populate `backtest_vellam.xlsx`. By providing the user-selected anchor points (Input, Yellow, and Red cells), a new script will automatically calculate the Family Tree matches, generate the predicted pool, and cross-reference the numbers against the actual podium draws for Day N, N+1, N+2, and N+3. **Crucially, Sheet 1 will remain entirely untouched.**

## Implementation Steps

### 1. Create a Dedicated Backtesting Script
*   **File:** Create a new script named `backtest_data.py` to handle the batch processing.
*   **Data Sources:** Load `backtest_vellam.xlsx` along with `podium.xlsx` and `family_tree.xlsx` (in read-only, data-only mode to optimize performance).

### 2. Safeguard Sheet 1 and Setup Sheet 2
*   **Strict Rule:** Under no circumstances will the script edit Sheet 1 of `backtest_vellam.xlsx`.
*   **Copy Data:** The script will copy the contents of Sheet 1 into **Sheet 2**. All subsequent reads for pending calculations and all data updates will be performed strictly on Sheet 2.

### 3. Parse Anchor Data from Sheet 2
*   Iterate through the rows of Sheet 2 in `backtest_vellam.xlsx` (starting from the first data row, e.g., Row 4).
*   For pending rows, read the manually provided anchor points:
    *   **Input Cell** (Col C) and Value (Col D)
    *   **Yellow Cell** (Col E) and Value (Col F)
    *   **Red Cell** (Col G) and Value (Col H)

### 4. Automate Family Tree Extraction (Step 13)
*   Implement the core logic from `step13_highlight_family_tree` without the Excel styling overhead:
    *   Locate the columns in `family_tree.xlsx` that contain the Yellow Cell Value.
    *   Search those specific columns for the calculated neighbours of the Red Cell Value.
    *   Store the matching coordinates and unique values.
*   **Write Output:** Populate **Col J** (Family Tree Highlight) and **Col K** (Highlighted Numbers) in **Sheet 2** with comma-separated strings of the matches.

### 5. Automate Prediction & Verification (Step 14)
*   **Generate Pool:** Replicate the pool generation from `step14`. For each unique blue value (from Col K), generate its **block** (8 numbers). 
    *   **Pool M (Blocks):** The collection of all generated 8 block numbers.
    *   **Pool L (Neighbours of Blocks):** The neighbours of all the numbers in Pool M.
*   **Cross-Reference:** 
    *   Parse the numeric row index from the Input Cell coordinate (e.g., `N1120` -> Row `1120`). This is **Day N**.
    *   Identify the target company offset for the Input Cell.
    *   Check `podium.xlsx` against the generated pools with the correct company filter.
*   **Write Output (to Sheet 2):** 
    *   **Col L:** Check for **Pool L** (neighbours of the 8 block numbers) that hit on **Day N** (the same day/row as the input number in Col C). Write these hits to Col L.
    *   **Col M:** Check `podium.xlsx` for any **Pool M** hits (the original 8 block numbers) found on **Day N+1**, **Day N+2**, and **Day N+3**. Write these results formatted into Col M (e.g., "N+1 - 1234, N+2 - 5678").
    *   **Col N:** Discard (do not write to Col N).

### 6. Finalization
*   Save the updated `backtest_vellam.xlsx` with the populated **Sheet 2**.
*   Output a brief console summary of rows processed.
