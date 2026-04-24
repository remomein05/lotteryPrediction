# Terminal Execution Log
```text
Loading '/home/muru/Repo/selvam_viruthi/podium.xlsx'... (this might take a few seconds)
Opening 'podium.xlsx' for reference...

--- Step 1: Initialization ---
Enter the cell coordinate to select (e.g., Q1086): M1126

--- Summary ---
Selected Cell: M1126
Target Number: 9744
Identified Company: Sabah

--- Step 2: Permutation Limitation Check ---
Number '9744' generates 12 unique permutations.
Check passed: Permutations are 12 or less and number has exactly one repeating digit.
Closed reference 'podium.xlsx' window.

--- Step 3: Generating Neighbours ---
Neighbours for '9744': 7440, 8744, 9844, 9644

--- Step 4: Searching and Highlighting ---
Found and highlighted 3 occurrence(s) of Neighbours within 'Sabah'.
Highlighted target cell M1126 in light green.
Saving results to '/home/muru/Repo/selvam_viruthi/podium_highlighted.xlsx'...
Done! Automatically opening the generated file...

--- Step 5: Select Highlighted Cell ---
Please review the highlighted dataset in Excel.
Enter the cell coordinate of one of the highlighted cells (or 'exit' to quit): R816
Waiting for LibreOffice to release the lock on 'podium_highlighted.xlsx'...
Closed 'podium_highlighted.xlsx' window.

--- Step 6: Family Generation ---
8551
9662
0773
1884

--- Step 7: Blocks for Family ---
--- Block for Family Number 8551 ---
B         : 8551
SB        : 9552
POT       : 5543
POT SB    : 7655
FA        : 6300
FA SB     : 7400
FA POT    : 9800
FA POT SB : 2100

--- Block for Family Number 9662 ---
B         : 9662
SB        : 8441
POT       : 8876
POT SB    : 4322
FA        : 7411
FA SB     : 9963
FA POT    : 3321
FA POT SB : 9877

--- Block for Family Number 0773 ---
B         : 0773
SB        : 7330
POT       : 9110
POT SB    : 9910
FA        : 8522
FA SB     : 8852
FA POT    : 6654
FA POT SB : 6544

--- Block for Family Number 1884 ---
B         : 1884
SB        : 9622
POT       : 4432
POT SB    : 8766
FA        : 9633
FA SB     : 7741
FA POT    : 9987
FA POT SB : 3211


--- Step 8: Generating Block ---
B         : 7440
SB        : 6630
POT       : 2210
POT SB    : 9880
FA        : 9952
FA SB     : 8511
FA POT    : 7765
FA POT SB : 5433

--- Step 9: Generating Neighbors for the 40-Member Pool ---
Total unique Neighbors generated: 120

--- Step 10: 24-Permutation Removal Filter ---
Removed 0 numbers with 24 permutations.
Retained 120 numbers with repeating digits (4, 6, or 12 perms).

--- Step 11: Searching and Highlighting Final Neighbors ---
Target Block Origin: Row 810, Col 16
Found and highlighted 97 occurrence(s) in light red (Same Company, Left/Above only).

Saving final results to 'podium_highlighted.xlsx'...
Done! Automatically opening the generated file...

--- Final Results (Light Red Highlights) ---
Cell E672: 9953
Cell F73: 7764
Cell F708: 6653
Cell H151: 8666
Cell H330: 6653
Cell H390: 8800
Cell H430: 7441
Cell H439: 8840
Cell H469: 5520
Cell H499: 6330
Cell H518: 7500
Cell H573: 9942
Cell H609: 8422
Cell H645: 6640
Cell H726: 7775
Cell I151: 8411
Cell I181: 8633
Cell I201: 8411
Cell I221: 7731
Cell I231: 8853
Cell I290: 8840
Cell I518: 9522
Cell I789: 5520
Cell J131: 8662
Cell J151: 8831
Cell J221: 8851
Cell J231: 9663
Cell J271: 9533
Cell J340: 8622
Cell J370: 8880
Cell J518: 8840
Cell J527: 8552
Cell J744: 5300
Cell J753: 7311
Cell J771: 5300
Cell J789: 7751
Cell J807: 8442
Cell L43: 9920
Cell L53: 9942
Cell L83: 6640
Cell L171: 6640
Cell L400: 8440
Cell L430: 9744
Cell L527: 8550
Cell L600: 9991
Cell L618: 8655
Cell L645: 7311
Cell L708: 3100
Cell L717: 6655
Cell L780: 8411
Cell M33: 6440
Cell M122: 2220
Cell M181: 9951
Cell M191: 7522
Cell M320: 9511
Cell M340: 8400
Cell M360: 8840
Cell M480: 8851
Cell M554: 3100
Cell M699: 6400
Cell N83: 8831
Cell N141: 9942
Cell N221: 6653
Cell N281: 8875
Cell N469: 9942
Cell N480: 9663
Cell N489: 8877
Cell N509: 6433
Cell N591: 7331
Cell N654: 8442
Cell N798: 1100
Cell P131: 8840
Cell P211: 8862
Cell P261: 6664
Cell P271: 8110
Cell P281: 9553
Cell P330: 6400
Cell P400: 4211
Cell P439: 9441
Cell P449: 7755
Cell P469: 8442
Cell P527: 6200
Cell P573: 7764
Cell P663: 8411
Cell P771: 6433
Cell Q73: 7742
Cell Q350: 4431
Cell Q400: 9700
Cell Q636: 7522
Cell Q645: 6631
Cell Q672: 9663
Cell Q780: 7441
Cell R23: 9744
Cell R33: 6440
Cell R400: 8875
Cell R420: 9221
Cell R762: 9663

--- Step 12: Select Red Highlighted Cell and Generate Neighbours ---

Enter the cell coordinate of a light-red highlighted cell (or type 'done' to finish): N798
Selected red highlighted cell N798 with value 1100.
Neighbours for 1100: 1110, 9110

--- Step 13: Searching and Highlighting Neighbours in family_tree.xlsx ---
Loading '/home/muru/Repo/selvam_viruthi/family_tree.xlsx'... (this might take a few seconds)
Found and highlighted 2 occurrence(s) of Step 5 value 7440 in grey.
Identified columns with Step 5 value: [3, 5]
Neighbours for 1100: 1110, 9110
Found and highlighted 2 occurrence(s) of Neighbours in light blue.
Saving results to '/home/muru/Repo/selvam_viruthi/family_tree_highlighted.xlsx'...
Done! Automatically opening the generated file...

--- Grey and Blue Highlight Results ---
Grey highlights (Step 5 value):
Cell C27: 7440
Cell E45: 7440
Blue highlights (Neighbours):
Cell E32: 9110
Cell C33: 9110

--- Step 14: Generating Blocks and Neighbours, then Highlighting Duplicates ---
Found 2 blue-highlighted entries, 1 unique values.
Cleared light red highlights from podium_highlighted.xlsx.
Green-highlighted cell: M1126, Company offset: 6 (Sabah)
Processing blue value #1 '9110': block members ['6544', '6654', '7330', '7730', '8522', '8852', '9110', '9910'], neighbours ['1100', '5544', '6330', '6444', '6644', '6653', '6655', '6664', '7331', '7522', '7544', '7720', '7731', '7740', '8110', '8330', '8422', '8622', '8842', '8851', '8853', '8862', '9111', '9522', '9733', '9773', '9900', '9911', '9920', '9991']
Highlighted 30 cells for blue value '9110' using color #f2dceb.
Found and highlighted 30 occurrence(s) of neighbours from blocks in color-coded highlights (company filter applied).
Saving results to '/home/muru/Repo/selvam_viruthi/next_podiums.xlsx'...
Done! Automatically opening the generated file...

--- Four Colour Highlight Results ---
Cell L43: 9920 (color #f2dceb)
Cell M191: 7522 (color #f2dceb)
Cell P211: 8862 (color #f2dceb)
Cell I221: 7731 (color #f2dceb)
Cell J221: 8851 (color #f2dceb)
Cell N221: 6653 (color #f2dceb)
Cell I231: 8853 (color #f2dceb)
Cell P261: 6664 (color #f2dceb)
Cell P271: 8110 (color #f2dceb)
Cell H330: 6653 (color #f2dceb)
Cell J340: 8622 (color #f2dceb)
Cell M480: 8851 (color #f2dceb)
Cell H499: 6330 (color #f2dceb)
Cell I518: 9522 (color #f2dceb)
Cell N591: 7331 (color #f2dceb)
Cell L600: 9991 (color #f2dceb)
Cell H609: 8422 (color #f2dceb)
Cell Q636: 7522 (color #f2dceb)
Cell F708: 6653 (color #f2dceb)
Cell L717: 6655 (color #f2dceb)
Cell N798: 1100 (color #f2dceb)
Cell J870: 8851 (color #f2dceb)
Cell Q888: 8622 (color #f2dceb)
Cell H933: 6655 (color #f2dceb)
Cell P960: 6655 (color #f2dceb)
Cell R969: 7740 (color #f2dceb)
Cell I1006: 9522 (color #f2dceb)
Cell J1006: 7544 (color #f2dceb)
Cell I1015: 7731 (color #f2dceb)
Cell L1099: 8330 (color #f2dceb)

[Finished processing results for cell N798]

--- Final Unique Blue Value -> Color Mapping ---
1. [30;48;2;242;220;235m 9110 [0m

'next_podiums.xlsx' and updated 'podium_highlighted.xlsx' are now open for your review.
You may now select another light-red cell or type 'done' to exit.

Enter the cell coordinate of a light-red highlighted cell (or type 'done' to finish): DONE
Exiting iterative processing.

```
