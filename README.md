# PowerBi Auto Refresh
It is not currently possible to periodically refresh PowerBi Desktop data when not imported from a database - in which case you could use the DirectQuery capabilities of PowerBi.

When working with a simple Excel file, I did not find any way to automate the refresh, and ended up using PyWinAuto instead. This is not a pretty solution, but it works reliably and does not require much client input - simply double-click the launcher instead of the .pbix file.

This program detects file modifications on the input excel file - when the file gets saved, and auto-refreshes data on powerbi.

## Installation
Needs a working install of Python3.X - tested with Python 3.8.
Dependencies: can be found in requirements.txt and installed with `pip install -r requirements.txt`.

## Configuration
The launch.dat needs to point to you python executable, as well as your .pbix file.
Depending on your locale, you must change the buttons labels.
Example: win.Actualiser in french becomes win.Refresh in english.
