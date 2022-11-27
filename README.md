# Overview

Updated Nov 2022 with a new Python file, tentatively named recon2022highlight.py. 

Execute with `python recon2022highlight.py <filename> <sheet1> <sheet2>`.

This script is very specific to someone's use case that I wrote it for. If you're looking for something to perform cash reconciliation or something similar, you'll need to tweak it a lot. 

It collects all unique deposit ids and the sums of the transaction amounts for each id in two sheets of the Excel file. It then cross-references the ids with those in the second one (asymmetrically, so it's only checking if the first sheet's ids are in the second), and if the values match exactly then it colors the cells that composed the total green. If they don't match exactly but do have values, it colors the cells yellow, and if there is no result for the id from the first sheet in the second it colors it orange. 

This iteration of the script is much less likely to result in false matches that the old version was. Mostly because I used the wrong data structure for it when I first wrote it. 

I've left the old description below since the older Python files are there, but they're not particularly accurate.

# Old Description

## What's this?
Performs transaction reconciliation on two sheets of an Excel spreadsheet. Helps automate some drudge work. This was written based off of alifarhadd's auto_bank_reconciliation script, modified to suit a very specific purpose.

- verysimplerecon.py does not perform the sum matching. I found that recon and simplerecon make mistakes on various edge cases. 
- recon_revamp.py is a work in progress to use pandas to perform the work, since doing it all in openpyxl seems error-prone (for me).

## Function
This script is organized as such:

1. Looks for matching absolute values of transaction amounts in Column I from the first sheet and Column J from the second sheet. It highlights matches green and nonmatches orange.
2. Checks Column P in the first sheet for the last 6 characters. If they begin with V, indicating an account number type, then it takes the remaining 5 digits of the number and compares it with the last five digits of column H in the second sheet. It highlights these matches green.
3. Checks the transaction amounts in Column J of the second sheet nearby (within 8 rows) for those that match the account number from the previous function. It sums them up and if they then match a total amount from Column I in the first sheet, it highlights the amounts green. 

In order to make this script work you need 2 things:
 * Python3
 * an .XLSX file with two sheets
 
## Instructions
1. Excel file contains two sheets: (sheet_A, sheet_B)
2. Relevant columns in "Sheet_A" are I and P

| Column I | Column P         |
| ------------- |:----------:|
|   79356       | Housing Construction V83248			 |
|   20243       | Solar Panel Company      |
| 	94319       | Mech Equipment V81349      |

3. Relevant columns in "Sheet_B" are H and J

| Column H        | Column J        |
| ------------- |:-------------:|
|   Construction #V83248  | 79356		 |
|  Design Architecture      | 19288      |
| 	Equipment #V81349    | 94319      |

4. Keep your Excel file in the same folder as the script recon.py
5. ```python recon.py```
6. Follow the prompt
7. After the script is done, your results file ready.xlsx will have these two cells highlighted in both sheets since they were the same in both sheets:

| Column I | Column P         |
| ------------- |:----------:|
|   79356       | Housing Construction V83248			 |
| 	94319       | Mech Equipment V81349      |

| Column H        | Column J        |
| ------------- |:-------------:|
|   Construction #V83248  | 79356		 |
| 	Equipment #V81349    | 94319      |
