from __future__ import annotations

# type: ignore
import os
import sys
import time
from collections import Counter
from typing import Any, cast, List, Set, Tuple

from logger import setup_logging
from constants import COMPANIES
from utils import (
    calculate_permutations,
    get_sb,
    get_pot,
    get_fa,
    get_family,
    get_block,
    get_neighbours,
    get_block_origin,
    get_company_offset,
    open_file,
    close_file,
    is_file_locked_by_libreoffice,
    wait_for_file_unlock,
    get_terminal_color,
)

try:
    import openpyxl
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Error: The 'openpyxl' library is required to read Excel files.")
    print("Please install it by running: pip install openpyxl")
    sys.exit(1)

# Start dynamically teeing stdout into result.md and hook input()
setup_logging("result.md")


def main():
    # Use absolute paths so the script works regardless of the working directory
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_filename = os.path.join(_script_dir, "podium.xlsx")
    output_filename = os.path.join(_script_dir, "podium_highlighted.xlsx")
    
    if not os.path.exists(excel_filename):
        print(f"Error: Could not find '{excel_filename}' in the current directory.")
        sys.exit(1)
        
    print(f"Loading '{excel_filename}'... (this might take a few seconds)")
    try:
        # Load the workbook in standard mode so styles and highlight fills are preserved on save
        wb = openpyxl.load_workbook(excel_filename, data_only=False)
        ws = cast(Any, wb.active)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    # Initialize variables to satisfy static analysis and avoid unbound variable warnings
    sel_val: str = ""
    blue_matches_summary: List[Tuple[str, str]] = []
    unique_blue_values: List[str] = []
    color_map: List[str] = []
    ph_ws = None
    ph_wb = None

    # Open podium.xlsx in Excel so the user can reference it while choosing a cell
    print("Opening 'podium.xlsx' for reference...")
    open_file(excel_filename)
        
    while True:
        # --- STEP 1: User Cell Selection ---
        print("\n--- Step 1: Initialization ---")
        cell_coord = input("Enter the cell coordinate to select (e.g., Q1086): ").strip().upper()
        if not cell_coord:
            print("Invalid input.")
            sys.exit(1)
            
        try:
            cell_value = ws[cell_coord].value
        except ValueError:
            print(f"Error: Invalid cell coordinate '{cell_coord}'. Please try again.")
            continue
        if cell_value is None:
            print(f"Error: Cell {cell_coord} is empty. Please try again.")
            continue
            
        # Ensure the value is treated as a string to preserve the 4-digit format easily
        number = str(cell_value).strip()
        
        target_offset = get_company_offset(ws, ws[cell_coord].row, ws[cell_coord].column)
        company_name = COMPANIES.get(target_offset, f"Unknown ({target_offset})")
        
        print("\n--- Summary ---")
        print(f"Selected Cell: {cell_coord}")
        print(f"Target Number: {number}")
        print(f"Identified Company: {company_name}")
        
        # --- STEP 2: Permutation Limitation Check ---
        print("\n--- Step 2: Permutation Limitation Check ---")
        perms = calculate_permutations(number)
        print(f"Number '{number}' generates {perms} unique permutations.")
        
        if perms > 12:
            print(f"Rule Triggered: Number '{number}' has {perms} permutations (must be 12 or less).")
            print("Please select another number with at least one repeating digit.")
            continue
        
        # New Rule: Check for exactly one repeating digit.
        counts = Counter(number)
        repeating_digits = [digit for digit, count in counts.items() if count > 1]
        
        if len(repeating_digits) != 1:
            print(f"Rule Triggered: Number '{number}' does not have exactly one repeating digit.")
            print("Please select another number (e.g., one with a double or triple digit).")
            continue
            
        print("Check passed: Permutations are 12 or less and number has exactly one repeating digit.")

        # Close the reference podium.xlsx Excel window now that input is accepted
        close_file("podium")
        print("Closed reference 'podium.xlsx' window.")

        # --- STEP 3: Generate Neighbours ---
        print("\n--- Step 3: Generating Neighbours ---")
        neighbours = get_neighbours(number)
        print(f"Neighbours for '{number}': {', '.join(neighbours)}")
        
        # --- STEP 4: Search and Highlight ---
        print("\n--- Step 4: Searching and Highlighting ---")
        light_yellow_fill = PatternFill(start_color="FFFFFF99", end_color="FFFFFF99", fill_type="solid")
        light_green_fill = PatternFill(start_color="FFCCFFCC", end_color="FFCCFFCC", fill_type="solid")
        
        found_matches: int = 0
        for row_idx in range(7, ws.max_row + 1):
            for col_idx in range(4, 19): # Columns D to R
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value is not None and str(cell.value).strip() != "":
                    cell_val_str = str(cell.value).strip().zfill(4)
                    if cell_val_str in neighbours and get_company_offset(ws, row_idx, col_idx) == target_offset:
                        cell.fill = light_yellow_fill
                        found_matches += 1
                        
        print(f"Found and highlighted {found_matches} occurrence(s) of Neighbours within '{company_name}'.")
        
        # Highlight the originally selected target cell in light green
        cast(Any, ws)[cell_coord].fill = light_green_fill # type: ignore
        print(f"Highlighted target cell {cell_coord} in light green.")
        
        print(f"Saving results to '{output_filename}'...")
        
        while True:
            try:
                wb.save(output_filename)
                break
            except PermissionError:
                print(f"\n[!] Error: '{output_filename}' is currently open in another program (like Excel).")
                input("    Please close the file and press ENTER to try saving again...")
                
        print("Done! Automatically opening the generated file...")
        open_file(output_filename)
        time.sleep(2)  # Give time for the spreadsheet application to fully open and display the file
            
        # --- STEP 5: Select Highlighted Cell ---
        print("\n--- Step 5: Select Highlighted Cell ---")
        print("Please review the highlighted dataset in Excel.")
        
        user_exited = False
        selected_highlight_coord = cell_coord
        while True:
            selected_highlight = input("Enter the cell coordinate of one of the highlighted cells (or 'exit' to quit): ").strip().upper()

            if selected_highlight == 'EXIT' or not selected_highlight:
                print("Exiting...")
                user_exited = True
                break

            try:
                sel_cell = cast(Any, ws)[selected_highlight] # type: ignore
                sel_val = str(sel_cell.value).strip() if sel_cell.value is not None else ""
            except ValueError:
                print(f"Error: Invalid cell coordinate '{selected_highlight}'.")
                continue

            # Check if it is a valid highlighted neighbour
            if sel_val in neighbours and get_company_offset(ws, int(sel_cell.row), int(sel_cell.column)) == target_offset: # type: ignore
                selected_highlight_coord = selected_highlight
                # Close podium_highlighted.xlsx now that the cell has been selected
                close_file(output_filename)
                if is_file_locked_by_libreoffice(output_filename):
                    print("Waiting for LibreOffice to release the lock on 'podium_highlighted.xlsx'...")
                    if not wait_for_file_unlock(output_filename, timeout_seconds=10):
                        input("Please close 'podium_highlighted.xlsx' and press ENTER to continue...")
                print("Closed 'podium_highlighted.xlsx' window.")
                break
            else:
                print(f"Error: Cell {selected_highlight} is not one of the highlighted Neighbours. Please rectify and try again.")
                
        if user_exited:
            break
                
        # --- STEP 6: Family Generation ---
        print("\n--- Step 6: Family Generation ---")
        family_numbers = get_family(sel_val)
        for fn in family_numbers:
            print(fn)
            
        # --- STEP 7: Blocks for Family ---
        print("\n--- Step 7: Blocks for Family ---")
        all_generated_numbers = set()
        for fn in family_numbers:
            all_generated_numbers.add(fn)
            print(f"--- Block for Family Number {fn} ---")
            fn_block = get_block(fn)
            for key, val in fn_block.items():
                all_generated_numbers.add(val)
                print(f"{key:<10}: {val}")
            print()
            
        # --- STEP 8: Generate Block for B ---
        print("\n--- Step 8: Generating Block ---")
        b_val = sel_val
        b_block = get_block(b_val)
        for key, val in b_block.items():
            print(f"{key:<10}: {val}")
            
        # Collect all generated numbers sequentially for Step 10/11
        all_generated_numbers.update(b_block.values())
        
        # --- STEP 9: Neighbors for the 40-Member Pool ---
        print("\n--- Step 9: Generating Neighbors for the 40-Member Pool ---")
        neighbour_pool = set()
        # all_generated_numbers already contains the 40 members from Step 7 and 8
        for num in sorted(list(all_generated_numbers)):
            nbs = get_neighbours(num)
            for nb in nbs:
                neighbour_pool.add(nb)
        print(f"Total unique Neighbors generated: {len(neighbour_pool)}")
                
        # --- STEP 10: 24-Permutation Removal Filter ---
        print("\n--- Step 10: 24-Permutation Removal Filter ---")
        final_highlights: Set[str] = set()
        removed_cnt: int = 0
        for nb in sorted(list(neighbour_pool)):
            if calculate_permutations(nb) == 24: # type: ignore
                removed_cnt += 1 # type: ignore
            else:
                final_highlights.add(nb)
        
        print(f"Removed {removed_cnt} numbers with 24 permutations.")
        print(f"Retained {len(final_highlights)} numbers with repeating digits (4, 6, or 12 perms).")
        
        # --- STEP 11: Searching and Highlighting Final Neighbors ---
        print(f"\n--- Step 11: Searching and Highlighting Final Neighbors ---")
        close_file("podium_highlighted")
        if is_file_locked_by_libreoffice(output_filename):
            print("Waiting for LibreOffice to release the lock on 'podium_highlighted.xlsx' before saving final red highlights...")
            if not wait_for_file_unlock(output_filename, timeout_seconds=10):
                input("Please close 'podium_highlighted.xlsx' and press ENTER to continue...")
        light_red_fill = PatternFill(start_color="FFFF9999", end_color="FFFF9999", fill_type="solid")
        found_red_cnt: int = 0
        red_matches_summary: List[Tuple[str, str]] = []
        
        # Identify the selected block's origin for the "Left or Above" filter
        ts_cell = cast(Any, ws)[selected_highlight_coord]
        target_os_row, target_os_col = get_block_origin(ws, ts_cell.row, ts_cell.column) # type: ignore
        print(f"Target Block Origin: Row {target_os_row}, Col {target_os_col}")
        
        # Expand search to all relevant columns D (4) to R (18)
        for col_idx in range(4, 19):
            for row_idx in range(7, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx) # type: ignore
                if cell.value is not None and str(cell.value).strip() != "":
                    val_str = str(cell.value).strip().zfill(4)
                    if val_str not in final_highlights:
                        continue

                    # Verify Company
                    m_start_row, m_start_col = get_block_origin(ws, row_idx, col_idx)
                    m_offset = int(row_idx) - int(m_start_row)
                    
                    # Step 11 Special Rule: Treat Magnum (0) and Singapore (5) as a combined entity.
                    # This allows matches from both companies to be highlighted if either was selected.
                    # Note: If Singapore is empty/missing in a block, the value check above skips it automatically.
                    if (m_offset == target_offset) or ({m_offset, target_offset} == {0, 5}):
                        # Verify Directional Constraint: Strictly Before (Above OR Left-Same-Row)
                        is_above = m_start_row < target_os_row
                        is_left_same_row = (m_start_row == target_os_row) and (m_start_col < target_os_col)
                        
                        if is_above or is_left_same_row:
                            cell.fill = light_red_fill
                            found_red_cnt += 1
                            red_matches_summary.append((cell.coordinate, val_str))
                            
        print(f"Found and highlighted {found_red_cnt} occurrence(s) in light red (Same Company, Left/Above only).")
        
        print("\nSaving final results to 'podium_highlighted.xlsx'...")
        while True:
            try:
                wb.save(output_filename)
                break
            except PermissionError:
                input(f"\n[!] Error: '{output_filename}' is currently open in another program (like Excel).\n    Please close the file and press ENTER to try saving again...")
                
        print("Done! Automatically opening the generated file...")
        open_file(output_filename)
                
        if red_matches_summary:
            print("\n--- Final Results (Light Red Highlights) ---")
            for coord, val in red_matches_summary:
                print(f"Cell {coord}: {val}")

        # --- STEP 12: Validate red-highlighted cell and generate its Neighbours (Iterative Loop) ---
        print("\n--- Step 12: Select Red Highlighted Cell and Generate Neighbours ---")
        valid_red_cells = {coord: val for coord, val in red_matches_summary}

        if not valid_red_cells:
            print("No light-red highlighted cells were found; skipping Step 12-14.")
        else:
            # One-time prompt: read the user's choice once and process it (no iterative loop)
            ft_ws = None
            blue_matches_summary = []
            unique_blue_values = []

            step12_coord = input("\nEnter the cell coordinate of a light-red highlighted cell (or type 'done' to finish): ").strip().upper()
            if step12_coord == 'DONE' or step12_coord == 'EXIT' or not step12_coord:
                print("Exiting iterative processing.")
            elif step12_coord not in valid_red_cells:
                print(f"Error: {step12_coord} is not in the red highlighted list. Skipping iterative processing.")
            else:
                step12_val = valid_red_cells[step12_coord]
                print(f"Selected red highlighted cell {step12_coord} with value {step12_val}.")

                step12_neighbours = get_neighbours(step12_val)
                print(f"Neighbours for {step12_val}: {', '.join(step12_neighbours)}")

                # Close podium_highlighted.xlsx after user input
                close_file("podium_highlighted.xlsx")
                time.sleep(1)

                # --- STEP 13: Search and Highlight Neighbours in family_tree.xlsx ---
            # --- STEP 12: Select Red Highlighted Cell and Generate Neighbours ---
            print("\n--- Step 12: Select Red Highlighted Cell and Generate Neighbours ---")
            valid_red_cells = {coord: val for coord, val in red_matches_summary}

            if not valid_red_cells:
                print("No light-red highlighted cells were found; skipping Step 12-14.")
            else:
                # Try automatic selection: pick the first red cell that has exactly 12 permutations
                # and is not from the same day (block origin row) as the originally selected yellow cell.
                ft_ws = None
                blue_matches_summary = []
                unique_blue_values = []

                # Prompt the user first. If they provide a selection, honor it (after same-day validation).
                step12_coord = input("\nEnter the cell coordinate of a light-red highlighted cell (or type 'done' to finish): ").strip().upper()
                user_provided = not (step12_coord == 'DONE' or step12_coord == 'EXIT' or not step12_coord)

                step12_val = None
                process_step12 = False

                if user_provided:
                    if step12_coord not in valid_red_cells:
                        print(f"Error: {step12_coord} is not in the red highlighted list. Will attempt auto-selection.")
                    else:
                        # Validate same-day constraint for user selection
                        try:
                            rc_cell = cast(Any, ws)[step12_coord]
                            r_start_row, r_start_col = get_block_origin(ws, int(rc_cell.row), int(rc_cell.column))
                            if r_start_row == target_os_row and r_start_col == target_os_col:
                                print("Selected red cell is from the same day as the yellow cell. Will attempt auto-selection.")
                            else:
                                step12_val = valid_red_cells[step12_coord]
                                process_step12 = True
                                print(f"User-selected red highlighted cell {step12_coord} with value {step12_val}.")
                        except Exception as e:
                            print(f"Error validating selected red cell {step12_coord}: {e}. Will attempt auto-selection.")

                if not process_step12:
                    # Auto-selection disabled: require explicit user selection.
                    print("No valid red-cell selection provided; auto-selection is disabled. Skipping Step 12 iterative processing.")

                if process_step12 and step12_val is not None:
                    step12_neighbours = get_neighbours(step12_val)
                    print(f"Neighbours for {step12_val}: {', '.join(step12_neighbours)}")

                    # Close podium_highlighted.xlsx after selection
                    close_file("podium_highlighted.xlsx")
                    time.sleep(1)

                    # --- STEP 13: Search and Highlight Neighbours in family_tree.xlsx ---
                print("\n--- Step 13: Searching and Highlighting Neighbours in family_tree.xlsx ---")
                family_tree_filename = os.path.join(_script_dir, "family_tree.xlsx")
                family_tree_output_filename = os.path.join(_script_dir, "family_tree_highlighted.xlsx")
                
                if not os.path.exists(family_tree_filename):
                    print(f"Error: Could not find '{family_tree_filename}' in the current directory.")
                    print("Skipping Step 13.")
                else:
                    print(f"Loading '{family_tree_filename}'... (this might take a few seconds)")
                    try:
                        ft_wb = openpyxl.load_workbook(family_tree_filename, data_only=False)
                        ft_ws = cast(Any, ft_wb.active)
                    except Exception as e:
                        print(f"Error reading family_tree.xlsx: {e}")
                        print("Skipping Step 13.")
                    else:
                        grey_fill = PatternFill(start_color="FFD3D3D3", end_color="FFD3D3D3", fill_type="solid")
                        light_blue_fill = PatternFill(start_color="FFCCFFFF", end_color="FFCCFFFF", fill_type="solid")
                        found_grey_cnt: int = 0
                        found_blue_cnt: int = 0
                        grey_matches_summary: List[Tuple[str, str]] = []
                        
                        # First: Locate the Step 5 selected value and highlight in grey, collect columns
                        columns_with_step5: Set[int] = set()
                        step5_val_str = str(sel_val).strip().zfill(4)
                        for row_idx in range(3, 47):
                            for col_idx in range(3, 26):
                                cell = ft_ws.cell(row=row_idx, column=col_idx)
                                if cell.value is not None and str(cell.value).strip() != "":
                                    cell_val_str = str(cell.value).strip().zfill(4)
                                    if cell_val_str == step5_val_str:
                                        cell.fill = grey_fill
                                        found_grey_cnt += 1
                                        grey_matches_summary.append((cell.coordinate, cell_val_str))
                                        columns_with_step5.add(col_idx)
                        
                        print(f"Found and highlighted {found_grey_cnt} occurrence(s) of Step 5 value {step5_val_str} in grey.")
                        print(f"Identified columns with Step 5 value: {sorted(columns_with_step5)}")
                        
                        # Third: Generate neighbours for Step 12 selected value
                        step12_neighbours = get_neighbours(step12_val)
                        print(f"Neighbours for {step12_val}: {', '.join(step12_neighbours)}")
                        
                        # Fourth: Search for neighbours only in identified columns
                        for row_idx in range(3, 47):
                            for col_idx in range(3, 26):
                                if col_idx in columns_with_step5:
                                    cell = ft_ws.cell(row=row_idx, column=col_idx)
                                    if cell.value is not None and str(cell.value).strip() != "":
                                        cell_val_str = str(cell.value).strip().zfill(4)
                                        if cell_val_str in step12_neighbours:
                                            cell.fill = light_blue_fill
                                            found_blue_cnt += 1
                                            blue_matches_summary.append((cell.coordinate, cell_val_str))
                        
                        print(f"Found and highlighted {found_blue_cnt} occurrence(s) of Neighbours in light blue.")
                        
                        print(f"Saving results to '{family_tree_output_filename}'...")
                        while True:
                            try:
                                ft_wb.save(family_tree_output_filename)
                                break
                            except PermissionError:
                                input(f"\n[!] Error: '{family_tree_output_filename}' is currently open in another program (like Excel).\n    Please close the file and press ENTER to try saving again...")
                                
                        print("Done! Automatically opening the generated file...")
                        open_file(family_tree_output_filename)
                        
                        if grey_matches_summary or blue_matches_summary:
                            print("\n--- Grey and Blue Highlight Results ---")
                            if grey_matches_summary:
                                print("Grey highlights (Step 5 value):")
                                for coord, val in grey_matches_summary:
                                    print(f"Cell {coord}: {val}")
                            if blue_matches_summary:
                                print("Blue highlights (Neighbours):")
                                for coord, val in blue_matches_summary:
                                    print(f"Cell {coord}: {val}")

                # --- STEP 14: Generate Blocks and Neighbours, then Highlight Duplicates ---
                print("\n--- Step 14: Generating Blocks and Neighbours, then Highlighting Duplicates ---")
                if ft_ws and blue_matches_summary:
                    # Deduplicate blue-highlighted cell numbers before further processing
                    unique_blue_values = sorted({val for _, val in blue_matches_summary})
                    print(f"Found {len(blue_matches_summary)} blue-highlighted entries, {len(unique_blue_values)} unique values.")

                    # Use requested 4 distinct colors exactly for each unique blue number
                    color_map = ["#f2dceb", "#79d0f2", "#357826", "#b0c8e9"]

                    # Re-load podium_highlighted.xlsx to modify it
                    next_podiums_filename = os.path.join(_script_dir, "next_podiums.xlsx")
                    try:
                        ph_wb = openpyxl.load_workbook(output_filename, data_only=False)
                        ph_ws = cast(Any, ph_wb.active)
                    except Exception as e:
                        print(f"Error re-loading podium_highlighted.xlsx: {e}")
                        # Fallback: if we cannot load the file, skip this step
                        print("Skipping Step 14 due to file load error.")
                        ph_ws = None

                    if ph_ws is not None:
                        # Remove all light red highlights before applying new color-coded highlights
                        light_red_fill = PatternFill(start_color="FFFF9999", end_color="FFFF9999", fill_type="solid")
                        for row in ph_ws.iter_rows(min_row=7, max_row=ph_ws.max_row, min_col=4, max_col=18):
                            for cell in row:
                                if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb == "FFFF9999":
                                    cell.fill = PatternFill(fill_type=None)
                        
                        print("Cleared light red highlights from podium_highlighted.xlsx.")
                        
                        # Get the company of the green-highlighted cell
                        green_offset = get_company_offset(ph_ws, ph_ws[cell_coord].row, ph_ws[cell_coord].column)
                        print(f"Green-highlighted cell: {cell_coord}, Company offset: {green_offset} ({COMPANIES.get(green_offset, 'Unknown')})")

                        total_purple_matches = 0
                        purple_matches_summary: List[Tuple[str, str, str]] = []

                        for idx, blue_val in enumerate(unique_blue_values):
                            active_color = color_map[idx % len(color_map)]
                            fill = PatternFill(start_color=f"FF{active_color.lstrip('#')}", end_color=f"FF{active_color.lstrip('#')}", fill_type="solid")

                            # Generate block and neighbours for this unique blue value
                            block_values = set(get_block(blue_val).values())
                            neighbours_from_this_block = set()
                            for num in block_values:
                                neighbours_from_this_block.update(get_neighbours(num))

                            print(f"Processing blue value #{idx+1} '{blue_val}': block members {sorted(block_values)}, neighbours {sorted(neighbours_from_this_block)}")

                            # Search podium_highlighted for this set of neighbours and highlight with active color
                            matched_count = 0
                            for row_idx in range(7, ph_ws.max_row + 1):
                                for col_idx in range(4, 19):  # Columns D to R
                                    cell = ph_ws.cell(row=row_idx, column=col_idx)
                                    if cell.value is not None and str(cell.value).strip() != "":
                                        cell_val_str = str(cell.value).strip().zfill(4)
                                        if cell_val_str in neighbours_from_this_block:
                                            # Check company
                                            m_start_row, m_start_col = get_block_origin(ph_ws, row_idx, col_idx)
                                            m_offset = int(row_idx) - int(m_start_row)
                                            if (m_offset == green_offset) or ({m_offset, green_offset} == {0, 5}):
                                                cell.fill = fill
                                                matched_count += 1
                                                total_purple_matches += 1
                                                purple_matches_summary.append((cell.coordinate, cell_val_str, active_color))

                            print(f"Highlighted {matched_count} cells for blue value '{blue_val}' using color {active_color}.")

                        print(f"Found and highlighted {total_purple_matches} occurrence(s) of neighbours from blocks in color-coded highlights (company filter applied).")

                        if ph_wb is not None:
                            print(f"Saving results to '{next_podiums_filename}'...")
                            while True:
                                try:
                                    ph_wb.save(next_podiums_filename)
                                    break
                                except PermissionError:
                                    input(f"\n[!] Error: '{next_podiums_filename}' is currently open in another program (like Excel).\n    Please close the file and press ENTER to try saving again...")
                        else:
                            print("Error: workbook object is missing, cannot save next_podiums.xlsx")

                        print("Done! Attempting to open the generated files...")
                        abs_next = os.path.abspath(next_podiums_filename)
                        opened_next = open_file(abs_next)
                        if not opened_next:
                            print(f"Note: Could not open '{abs_next}' automatically. You can open it manually.")
                        abs_out = os.path.abspath(output_filename)
                        opened_out = open_file(abs_out)
                        if not opened_out:
                            print(f"Note: Could not open '{abs_out}' automatically. You can open it manually.")
                        
                        if purple_matches_summary:
                            print("\n--- Four Colour Highlight Results ---")
                            for coord, val, color in purple_matches_summary:
                                print(f"Cell {coord}: {val} (color {color})")
                else:
                    print("No blue highlights found or family_tree.xlsx not loaded; skipping Step 14 for this cell.")

                print(f"\n[Finished processing results for cell {step12_coord}]")

                # Final user-facing summary: print the 4 unique blue values + color codes used in Step 14
                if unique_blue_values:
                    print("\n--- Final Unique Blue Value -> Color Mapping ---")
                    for idx, blue_val in enumerate(unique_blue_values[:4]):
                        active_color = color_map[idx % len(color_map)]
                        colored_val = get_terminal_color(blue_val, active_color)
                        print(f"{idx+1}. {colored_val}")

                print("\n'next_podiums.xlsx' and updated 'podium_highlighted.xlsx' are now open for your review.")

        break

if __name__ == "__main__":
    try:
        main()
    finally:
        # Safely shut down the custom logger.
        stdout_obj = cast(Any, sys.stdout)
        if hasattr(stdout_obj, "close_log"):
            stdout_obj.close_log()
