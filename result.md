# Terminal Execution Log
```text
Loading '/home/muru/Repo/selvam_viruthi/podium.xlsx'... (this might take a few seconds)
Opening 'podium.xlsx' for reference...

--- Step 1: Initialization ---
Enter the cell coordinate to select (e.g., Q1086): h1123

--- Summary ---
Selected Cell: H1123
Target Number: 8550
Identified Company: Sandakan

--- Step 2: Permutation Limitation Check ---
Number '8550' generates 12 unique permutations.
Check passed: Permutations are 12 or less and number has exactly one repeating digit.
Closed reference 'podium.xlsx' window.

--- Step 3: Generating Neighbours ---
Neighbours for '8550': 9550, 7550, 8551, 9855

--- Step 4: Searching and Highlighting ---
Found and highlighted 5 occurrence(s) of Neighbours within 'Sandakan'.
Highlighted target cell H1123 in light green.
Saving results to '/home/muru/Repo/selvam_viruthi/podium_highlighted.xlsx'...
Done! Automatically opening the generated file...

--- Step 5: Select Highlighted Cell ---
Please review the highlighted dataset in Excel.
Enter the cell coordinate of one of the highlighted cells (or 'exit' to quit): r759
Closed 'podium_highlighted.xlsx' window.

--- Step 6: Family Generation ---
8661
9772
0883
1994

--- Step 7: Blocks for Family ---
--- Block for Family Number 8661 ---
B         : 8661
SB        : 9442
POT       : 8843
POT SB    : 7622
FA        : 6311
FA SB     : 9974
FA POT    : 9833
FA POT SB : 7721

--- Block for Family Number 9772 ---
B         : 9772
SB        : 8331
POT       : 7611
POT SB    : 9943
FA        : 7422
FA SB     : 8863
FA POT    : 6621
FA POT SB : 9844

--- Block for Family Number 0883 ---
B         : 0883
SB        : 7220
POT       : 9440
POT SB    : 6610
FA        : 8533
FA SB     : 7752
FA POT    : 9954
FA POT SB : 6511

--- Block for Family Number 1994 ---
B         : 1994
SB        : 9611
POT       : 7732
POT SB    : 8733
FA        : 9644
FA SB     : 6641
FA POT    : 8722
FA POT SB : 8832


--- Step 8: Generating Block ---
B         : 7550
SB        : 5530
POT       : 5510
POT SB    : 9550
FA        : 5200
FA SB     : 8500
FA POT    : 6500
FA POT SB : 5400

--- Step 9: Generating Neighbors for the 40-Member Pool ---
Total unique Neighbors generated: 105

--- Step 10: 24-Permutation Removal Filter ---
Removed 0 numbers with 24 permutations.
Retained 105 numbers with repeating digits (4, 6, or 12 perms).

--- Step 11: Searching and Highlighting Final Neighbors ---
Target Block Origin: Row 756, Col 16
Found and highlighted 72 occurrence(s) in light red (Same Company, Left/Above only).

Saving final results to 'podium_highlighted.xlsx'...
Done! Automatically opening the generated file...

--- Final Results (Light Red Highlights) ---
Cell D208: 9940
Cell D248: 9722
Cell D705: 9733
Cell F89: 9953
Cell H80: 6550
Cell H327: 7533
Cell H417: 8864
Cell H466: 6220
Cell H542: 8442
Cell I158: 7522
Cell I218: 5100
Cell I278: 8422
Cell I317: 6400
Cell I327: 6600
Cell I387: 8550
Cell I486: 8772
Cell I515: 8440
Cell I615: 7753
Cell J50: 6400
Cell J80: 8633
Cell J148: 8644
Cell J218: 8833
Cell J317: 9722
Cell J397: 9544
Cell J542: 9331
Cell J560: 8864
Cell J624: 7762
Cell L109: 5520
Cell L327: 8842
Cell L367: 7511
Cell L407: 9544
Cell L446: 9771
Cell L606: 9443
Cell M80: 9975
Cell M148: 9551
Cell M268: 5520
Cell M327: 5100
Cell M579: 8330
Cell M696: 7751
Cell N99: 6400
Cell N427: 7311
Cell N515: 7731
Cell N542: 4200
Cell N624: 6220
Cell N651: 7322
Cell N696: 9940
Cell N714: 8831
Cell N741: 8644
Cell N750: 9942
Cell P347: 9973
Cell P407: 9553
Cell P466: 9744
Cell P560: 8220
Cell P606: 9544
Cell P615: 8332
Cell P678: 6400
Cell P741: 9443
Cell Q80: 9443
Cell Q168: 4420
Cell Q228: 9975
Cell Q705: 5300
Cell Q750: 7311
Cell R80: 8440
Cell R228: 8862
Cell R357: 6640
Cell R407: 5500
Cell R496: 7742
Cell R515: 6422
Cell R542: 6220
Cell R705: 6640
Cell R732: 7511
Cell R750: 7331

--- Step 12: Select Red Highlighted Cell and Generate Neighbours ---

Enter the cell coordinate of a light-red highlighted cell (or type 'done' to finish): r750
Selected red highlighted cell R750 with value 7331.
Neighbours for 7331: 8331, 6331, 7332, 7330

--- Step 12: Select Red Highlighted Cell and Generate Neighbours ---
Auto-selected red cell D208 with value 9940 (12 permutations, different day).
Selected red highlighted cell D208 with value 9940.
Neighbours for 9940: 9950, 9930, 9941, 9994

--- Step 13: Searching and Highlighting Neighbours in family_tree.xlsx ---
Loading '/home/muru/Repo/selvam_viruthi/family_tree.xlsx'... (this might take a few seconds)
Found and highlighted 1 occurrence(s) of Step 5 value 7550 in grey.
Identified columns with Step 5 value: [13]
Neighbours for 9940: 9950, 9930, 9941, 9994
Found and highlighted 1 occurrence(s) of Neighbours in light blue.
Saving results to '/home/muru/Repo/selvam_viruthi/family_tree_highlighted.xlsx'...
Done! Automatically opening the generated file...

--- Grey and Blue Highlight Results ---
Grey highlights (Step 5 value):
Cell M46: 7550
Blue highlights (Neighbours):
Cell M17: 9941

--- Step 14: Generating Blocks and Neighbours, then Highlighting Duplicates ---
Found 1 blue-highlighted entries, 1 unique values.
Cleared light red highlights from podium_highlighted.xlsx.
Green-highlighted cell: H1123, Company offset: 3 (Sandakan)
Processing blue value #1 '9941': block members ['6641', '7732', '8722', '8733', '8832', '9611', '9644', '9941'], neighbours ['6110', '6440', '6631', '6640', '6642', '6651', '7722', '7731', '7733', '7742', '8611', '8622', '8633', '8644', '8822', '8831', '8833', '8842', '9511', '9544', '9711', '9722', '9733', '9744', '9931', '9940', '9942', '9951']
Highlighted 32 cells for blue value '9941' using color #f2dceb.
Found and highlighted 32 occurrence(s) of neighbours from blocks in color-coded highlights (company filter applied).
Saving results to '/home/muru/Repo/selvam_viruthi/next_podiums.xlsx'...
Done! Attempting to open the generated files...

--- Four Colour Highlight Results ---
Cell J80: 8633 (color #f2dceb)
Cell J148: 8644 (color #f2dceb)
Cell D208: 9940 (color #f2dceb)
Cell J218: 8833 (color #f2dceb)
Cell D248: 9722 (color #f2dceb)
Cell J317: 9722 (color #f2dceb)
Cell L327: 8842 (color #f2dceb)
Cell R357: 6640 (color #f2dceb)
Cell J397: 9544 (color #f2dceb)
Cell L407: 9544 (color #f2dceb)
Cell P466: 9744 (color #f2dceb)
Cell R496: 7742 (color #f2dceb)
Cell N515: 7731 (color #f2dceb)
Cell P606: 9544 (color #f2dceb)
Cell N696: 9940 (color #f2dceb)
Cell D705: 9733 (color #f2dceb)
Cell R705: 6640 (color #f2dceb)
Cell N714: 8831 (color #f2dceb)
Cell N741: 8644 (color #f2dceb)
Cell N750: 9942 (color #f2dceb)
Cell M813: 9744 (color #f2dceb)
Cell Q840: 8833 (color #f2dceb)
Cell H867: 8842 (color #f2dceb)
Cell R867: 7742 (color #f2dceb)
Cell Q885: 8644 (color #f2dceb)
Cell N930: 9544 (color #f2dceb)
Cell F985: 9942 (color #f2dceb)
Cell Q1003: 8644 (color #f2dceb)
Cell M1022: 8644 (color #f2dceb)
Cell M1041: 9940 (color #f2dceb)
Cell H1050: 9940 (color #f2dceb)
Cell I1087: 9942 (color #f2dceb)

[Finished processing results for cell D208]

--- Final Unique Blue Value -> Color Mapping ---
1. [30;48;2;242;220;235m 9941 [0m

'next_podiums.xlsx' and updated 'podium_highlighted.xlsx' are now open for your review.

```
