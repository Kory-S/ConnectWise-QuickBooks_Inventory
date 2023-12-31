# ConnectWise-QuickBooks_Inventory
Compares inventory reports between ConnectWise and QuickBooks


# Compare inventory reports between ConnectWise and QuickBooks Online

This requires:
  - Python3
  - TKinter
  - CustomTKinter
  - OpenPyXL

This will compare each item in the ConnectWise inventory report with the items in the QuickBooks Online report.
This will export an excel spreadsheet containing 3 books

1st book is called "Accurate Count"
- This contains discrepencies where the inventory count does not match in CW and QBO.
- It will display the Product ID, the inventory count for CW and the inventory count for QBO
- It will also list to the right, everything that was found/is correct in the same format

2nd book is called "ConnectWise Only"
- This will call out items that only exist in CW but does not exist in QBO
- It will display the "Valuation Count" next to the "Product ID".
- It will also display to the right of the "CW Only" count, every item in the CW report

3rd book is called "QuickBooks Only"
- This will call out items that only exist in QBO but does not exist in CW
- It will display the "Valuation Count" next to the "Product ID".
- It will also display to the right of the "QBO Only" count, every item in the QBO report

## How to use

Make sure inventory.py and inventory_gui.py are in the same folder

In terminal (Preferred)
cd to location of inventory
C:...$ python inventory_gui.py
This will opena new window allowing you to select the reports from CW and QBO.
When using this through the terminal, it will open the folder to search for the reports in the current directory the application is ran from.

or...
Right click inventory_gui.py > run with... python
This will open a new window allowing you to select the reports from CW and QBO.