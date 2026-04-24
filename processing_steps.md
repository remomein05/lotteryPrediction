# Data Processing and Prediction Steps

This document tracks the sequential steps required to process the lottery dataset (`podium.xlsx`) and generate future predictions. These steps must be rigorously followed in order.

## Step 1: User Cell Selection
- When `main.py` starts, it automatically opens `podium.xlsx` in Excel so the user can visually browse the dataset while choosing a cell.
- `main.py` then prompts the user to input a specific cell coordinate (e.g., "H7") in the terminal.
- The script reads the `podium.xlsx` file (already loaded into memory), locates the provided cell, and extracts the 4-digit number inside it.
- This extracted number becomes the focal point/baseline for the subsequent processing steps.

## Step 2: Permutation Limitation Check
- The script calculates the mathematical permutations of the extracted 4-digit number.
- **Rule**: Processing only continues if the number has **12 or fewer unique permutations**. 
  - *Context*: A 4-digit number with 4 entirely distinct digits (e.g., `8610`) has 24 permutations. A number must have at least one repeating digit (e.g., `9772`) to have 12 or fewer permutations.
- **New Rule**: The number must have exactly **one** repeating digit. Numbers with two different sets of repeating digits (e.g., `6644` with two 6s and two 4s) are rejected, and the user must select another number.
- If the number fails these checks, the script pauses and asks the user for a new cell coordinate (the reference `podium.xlsx` window remains open for re-selection).
- Once all checks pass, the reference `podium.xlsx` Excel window is **automatically closed** before processing continues.

## Step 3: Generating Neighbours
- For a valid number with exactly one repeating digit, the script identifies its non-repeating digits.
- **Rule**: For each non-repeating digit, the script creates two "Neighbours" by incrementing and decrementing that specific digit by 1.
  - *Context*: The operation wraps around using modulo 10 (e.g., `0 - 1` becomes `9`, and `9 + 1` becomes `0`).
- After incrementing or decrementing, all four digits of the resulting Neighbour are sorted in **descending order**.
*Examples:*
- **Standard Case (`4330`)**: Non-repeating digits are `4` and `0`. Results: `9433`, `4331`, `3330`, `5330`.
- **First Two Repeat Case (`3332`)**: First two are `3,3`. Modifies indices 2 and 3.
    - Modifying index 2 (`3`): `4332`, `3322`
    - Modifying index 3 (`2`): `3333`, `3331`
- The resulting list of Neighbours is printed.

## Step 4: Searching and Highlighting
- **Company Filtering**: The script identifies which lottery company (Magnum, Kuda, Toto, etc.) the selected cell belongs to based on its relative row position within its layout block.
- The script searches the entire dataset (`podium.xlsx`) looking ONLY for duplicates of the Neighbours that occur within that **exact same company**.
- Any matching cells belonging to the same company are visually highlighted in **light yellow**.
- The specific cell originally selected by the user is highlighted in **light green** for easy visual reference.
- To prevent accidental data loss, the actively highlighted sheet is saved as a new file named `podium_highlighted.xlsx`.

## Step 5: Highlighted Cell Selection
- After generating and automatically opening the `podium_highlighted.xlsx` dataset in Excel, the script pauses in the terminal.
- The user visually reviews the newly colored Neighbours and explicitly inputs the cell coordinate of one of those matches back into the terminal.
- **Validation Rule**: The script mathematically verifies that the coordinate entered actually belongs to one of the valid computed duplicate Neighbours for that company. If the user selects an invalid cell, the script prompts them to rectify the mistake until a valid cell is chosen (the `podium_highlighted.xlsx` window remains open for re-reference).
- Once a valid cell is chosen, the `podium_highlighted.xlsx` Excel window is **automatically closed** before processing continues.
- The script validates the value as **B** for Step 6.

## Step 6: Family Generation
- From the validated input **B**, the script generates a "Family" of 4 additional numbers.
- **Rule**: Each subsequent number is created by adding 1 to each digit of the previous number (digit-wise addition modulo 10).
- *Example*: For `3100`:
  - `3100 + 1` = `4211`
  - `4211 + 1` = `5322`
  - `5322 + 1` = `6433`
  - `6433 + 1` = `7544`
- The script prints these Family numbers clearly in the terminal.

## Step 7: Blocks for Family
- For each of the 4 Family numbers generated in Step 6, the script calculates its full 8-member **Block** (B, SB, POT, POT SB, FA, FA SB, FA POT, FA POT SB).
- Each block is printed sequentially in the terminal.
- All 32 generated numbers (4 families * 8 members) are added to the global processing pool.

## Step 8: Generate Block
From the validated input **B**, the script calculates its 7 remaining "Block" members (8 numbers total in a Block):
- **B**: The starting validated highlighted number (e.g., `3200`).
- **SB**: Computed by taking `10 - digit` for each digit in **B**. If the digit is `0`, it elegantly remains `0` (calculated programmatically via modulo `(10 - digit) % 10`). The resulting 4 digits are then mathematically sorted in strict descending order (e.g., from `3200` to `7800` then finally sorted to `8700`).
- **POT**: Computed by multiplying each digit in **B** by 3, and keeping only the last digit of the result if it exceeds 9 (calculated programmatically via modulo `(digit * 3) % 10`). The final 4 digits are completely sorted in strict descending order (e.g., `4300` mathematically maps to `2, 9, 0, 0` and is sorted descending to `9200`).
- **POT SB**: Computed by taking the mathematically derived **POT** number, and applying the exact same **SB** subtraction logic (`(10 - digit) % 10`) to those digits, then sorting descending again.
- **FA**: Computed by taking each digit of **B**, seamlessly adding 5, and keeping only the final digit of the result (`(digit + 5) % 10`). The final 4 digits are completely sorted in strict descending order.
- **FA SB**: Computed by applying the exact **SB** subtraction logic cascade onto the newly generated **FA** number, strictly sorting descending again.
- **FA POT**: Computed by systematically applying the **POT** multiplier logic cascade directly to the **FA** number, extracting the 2nd digits, and sorting descending.
- **FA POT SB**: Computed by rigorously applying the **SB** subtraction logic cascade to the **FA POT** number, and sorting descending.

## Step 9: Neighbors for the 40-Member Pool
- The script uniquely aggregates all 40 mathematical numbers generated in Step 7 and Step 8.
- For *every* number in this pool, the script calculates the standard **Neighbours** for that number.
- **Neighbour Rule**: The script increments and decrements each **non-repeating** digit by 1 (modulo 10). 
- *Clarification on Repeats*: 
  - If the first two digits are identical (e.g., `3332`), the script acts on the **last two digits** (indices 2 and 3). This ensures that even repeating digits are processed if they fall in those positions.
  - For all other numbers, the script acts only on the **non-repeating** digits.
- Each newly generated Neighbour is mathematically sorted in descending order.

## Step 10: 24-Permutation Removal Filter
- The script iterates through the derived list of ~320 Neighbors from Step 9.
- For each Neighbor, it computes its base mathematical permutations.
- **Critical Rule**: If a number naturally possesses *exactly* **24 permutations** (all 4 digits unique), it is **immediately removed** from the pool.
- **Result**: Only numbers with repeating digits (4, 6, or 12 permutations) are retained for highlighting.

## Step 11: Secondary Cross-Dataset Highlighting
- The script sequentially iterates across the **entirety** of the dataset (all columns D-R) targeting the newly filtered list of Neighbors from Step 10.
- **Directional Constraint (Strictly Past)**: The script only highlights matches that are positioned **before** the user's selected draw block in reading order.
  - *Above*: All matches in rows located above the current block's starting row (regardless of column).
  - *Left*: Matches in the same row as the selected block, but in columns to the left.
  - *Excluded*: Any block to the right of the selection (in the same row) or in any row below the selection is skipped.
- **Company Partnership Rule**: Highlighting is generally limited to the same company. However, **Magnum (Offset 0)** and **Singapore (Offset 5)** are treated as a partnership. If the user selects a cell from either company, the script highlights matches from both.
- **Empty Row Handling**: In "Special Case Layouts" (Columns D-F) where Singapore is excluded/empty, the script gracefully skips the empty row while maintaining correct offset alignment for the other companies in the block.
- Any cell coordinate values fulfilling these rules are highlighted in **light red** (`#FFFF9999`).
- Once complete, the resulting structural dataset is saved into `podium_highlighted.xlsx`.

## Step 12: Validate Red-Highlighted Cell and Generate Neighbours (Interactive Loop)
- The script prompts the user to input a cell coordinate from the light-red highlighted cells in `podium_highlighted.xlsx`.
- **Iterative Process**: After inputting a valid coordinate, the script closes `podium_highlighted.xlsx`, processes everything through to Step 14, then re-opens `next_podiums.xlsx` and the updated `podium_highlighted.xlsx` for review.
- Once Step 14 concludes, the script returns to the Step 12 prompt. This allows the user to select another red-highlighted cell and see new results without restarting the entire script.
- The user can type "done" or "exit" at the Step 12 prompt to finish the script once they are satisfied with the results.
- The generated Neighbours for the selected red cell are printed in the terminal as before.

## Step 13: Search and Highlight Neighbours in family_tree.xlsx
- The script loads `family_tree.xlsx` (data range: C3 to Y46).
- It validates that the Step 12 selected value (red cell) is one of the known finalists; if not, it returns to Step 12.
- **First**: The script locates the Step 5 selected value (light yellow highlighted cell from podium_highlighted.xlsx) in family_tree.xlsx and highlights it in **grey** (`#D3D3D3`).
- **Second**: It identifies all columns in family_tree.xlsx that contain the Step 5 selected value.
- **Third**: It generates neighbours for the Step 12 selected value (red cell).
- **Fourth**: It searches for these neighbours ONLY within the identified columns from Step 13 Second.
- Only exact 4-digit match values are considered; any cell values that cannot be normalized to 4 digits are ignored.
- Matching neighbour cells are highlighted in **light blue** (`#CCFFFF`).
- If no matches are found, the script prints a warning and allows the user to choose another red highlighted cell for Step 12 (without terminating).
- The results are saved as `family_tree_highlighted.xlsx` and opened automatically.
- A summary of all grey and blue-highlighted matches (coordinate + value) is printed in the terminal.

## Step 14: Generate Blocks and Neighbours from Blue-Highlighted Cells, then Highlight Duplicates
- The script first collects all light-blue highlighted cell numbers from `family_tree.xlsx` and removes any duplicates.
- The next processing is done on the resulting unique blue-highlighted numbers only.
- For each unique blue-highlighted number, the script generates its full 8-member Block (B, SB, POT, POT SB, FA, FA SB, FA POT, FA POT SB).
- All unique numbers from that blue number's block are collected.
- For each of those block numbers, the script generates their Neighbours using the predefined Neighbour generation rule.
- The script then searches `podium_highlighted.xlsx` for duplicates of these neighbours, but only within the same company as the original green-highlighted cell (e.g., Sabah only).
- **Special Case**: For Magnum (Offset 0) or Singapore (Offset 5), matches from both companies are highlighted (following Step 11 rules).
- Matched cells for each unique blue number are highlighted with a dedicated color:
  - 1st unique blue number: `#f2dceb`
  - 2nd unique blue number: `#79d0f2`
  - 3rd unique blue number: `#6f98b8`
  - 4th unique blue number: `#b0c8e9`
- The updated file is saved as a new file named `next_podiums.xlsx` and opened automatically.
- A summary of all color-coded highlighted matches is printed in the terminal.
