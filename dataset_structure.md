# Lottery Dataset Structure (`podium.xlsx`)

## Overview
- **File**: `podium.xlsx`
- **Data Range**: Columns D to R, Rows 7 onwards (the dataset continuously grows downwards as new draws are appended manually at the bottom).
- **Content**: Past winning numbers (1st, 2nd, and 3rd place) for up to 7 different lottery companies.
- **Chronology**: Data is ordered sequentially from top to bottom (i.e., data at the bottom of the dataset is newer than the data above it).

## Section Layout
The data is grouped into "sections". Each section corresponds to a draw, with the rows representing the lottery companies and the columns representing the prize rankings.
- **Columns**: Exactly 3 columns per section (1st prize, 2nd prize, 3rd prize).
- **Spacing**: Results from a new draw are separated from an older draw by **more than 1 empty row** vertically and **at least 1 empty column** horizontally.
- **Rows**: The companies within a section are vertically ordered as follows:
  1. Magnum
  2. Kuda
  3. Toto
  4. Sandakan
  5. Sarawak
  6. Singapore *(Note: Excluded in the Special Case layout)*
  7. Sabah

## Range Nuances
- **Standard Layout (Columns G to R)**:
  - Spans from Row 7 to the bottom of the dataset.
  - Section Size: A typical contiguous block of **7 Rows x 3 Columns**.
  - Includes all 7 lottery companies.
- **Special Case Layout (Columns D, E, F)**:
  - Spans from Row 7 to the bottom of the dataset.
  - Section Size: A block of **5 Rows x 3 Columns**, followed by **1 empty row**, followed by a **1 Row x 3 Columns** block.
  - **Exception**: The "Singapore" company (normally the 6th row) is excluded from these sections, which accounts for the empty row.

## Data Point Formatting
- **Digit Sorting**: The 4-digit winning numbers in every cell have been pre-processed so that their digits are arranged in **strict descending order** (highest to lowest). 
  - *Example*: A literal winning number of `3204` is stored in the cell as `4320`.
- **Length**: A valid cell will always contain exactly **4 digits**.
