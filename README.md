# Randomised-ethereum-wallet-address
VBA (Visual Basic for Applications) code that produces a randomised wallet address in th0x e Ethereum address format intended to produce dummy data in Excel.

Produces strings starting with '0x' followed by 40 random lower-case letters, upper-case letters, and numbers.

# How To Use The Code
With an Excel spreadsheet open;
- Press Alt + F11
- In the insert tab, click the 'Module' option to open a new module.
- Copy and paste the code from the "Index.txt" file, into the new module.
- Save the spreadsheet as an .xlsm file (Macro Enabled Spreadsheet).
- In the relevant spreadsheet cell, enter in "=RandomizeAddress(0, 0)"

It will produce a string in the same format as this: '0x5skN54D85U3hOeniXWOjGG5Kp6zL4ISM4s7N42l2'

XLSM files will not push to a git repo due to the VBA code within them. Export the file to a normal workbook or CSV file to push to a repository of your own without the VBA code.
