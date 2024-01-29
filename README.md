# spreadsheet-error-analysis

Scripts for computing weighted linear least square fit in both Excel and GoogleSheet. In addition, for GoogleSheet, provide utilities to plot non-uniform error bars.

## Adding the `.vba` script to Excel file

To add the `.vba` script to an Excel file, first make sure the file is saved as a "macro-enabled workbook," with extension `.xlsm`. Then, open the Visual Basic Editor (If you enabled the "Developer" tab in the ribbon, the icon for the editor will show up there. Otherwise, use the shortcut `Alt + F11` for Windows and `Opt + F11` or `Fn + Opt + F11` on Mac).

Once in the Visual Basic editor, use the `insert` tab on top to insert a new module. Copy the content of the `.vba` file to the empty module, and close the editor. Whoa, the `WLINEST()` and `WLINEST_HELP()` function should be executable from the main workbook.

Alternatively, you can just copy a file with the module already defined. To help you get started, I have hosted one such file on my GoogleDrive, which you can access and download from [this link](https://docs.google.com/spreadsheets/d/1EqxN7xtYww0SaZHJSzJkBYhbNKfUlJEP/).

## Adding the `.gs` scripts to GoogleSheet

To add the `.gs` scripts to a GoogleSheet, first open the sheet online. Then, from the `Extension` tab click on `App Script`. This should bring up a script editor. Just add two script files and copy the content of each `.gs` file into a separate script file. Close the script editor when you're done, and reopen your sheet. Whoa the sheet should now have the added functionality.

_Note_ #1: when you use the script-defined functionality for the first time, Google will ask you to authorize the script, which you should do so.

_Note_ #2: it is possible to add just one of the two scripts if you only need part of the functionality.

Alternatively, you can just copy a sheet with the scripts already loaded. To help you get started, I have hosted one such file on my GoogleDrive, which you can access and copy from [this link](https://docs.google.com/spreadsheets/d/1kUjKvUM2l_IY2ujVOliPjyVWOOYWmDuccRAWg9UFzRg/).

## `errorBarChart.gs`: plot data with non-uniform error bars

In Excel, it is possible to plot non-uniform error bars ("custom" in Excel's menu option), i.e., for each data point, a different value of error can be specified. Unfortunately, this functionality is absent in GoogleSheet.

The known workaround for this problem is to create one data series for each data point, then enter the values of error manually for each point (see examples [here](https://www.youtube.com/watch?v=B-zKcSoYMq0) and [here](https://www.youtube.com/watch?v=Dj5kRkdtFNE)). As one can imagine, doing this for each data point can be quite tedious.

The `errorBarChart.gs` script is designed to automate this process. The script creates a new menu named `ErrorAnalysis`, under which there are options to make chart and add new data series to an existing chart (the chart being modified is always the last chart inserted. By copy-and-paste you can make an older chart the "last"). In either case, an auxillary sheet is created that allow each data point to be treated as a separate data series. The script then create the chart (or plot the new data series) for you, with error (a.k.a. uncertainty) automatically entered by the script.

## `WLINEST`: weighed linear least square fit

When the uncertainty of a data series is non-constant, a proper construction of trend line needs to weight different data point differently. Intuitively, the deviation of a data point from its trend line value should be compared relative to the size of the uncertainty, and the trend line is the line for which the sum of deviation thus weighed is minimized. More technically, a _**weighed** least square fit_, as opposed to linear least square fit, is required in such scenario.

In both Excel and GoogleSheet, the least square fit can be found using the `LINEST()` function. The function gives you the slope and intercept of the trend line, and optionally their uncertainties as well as additional statistical information.

Unfortunately, neither Excel nor GoogleSheet provides a function for weighed least square fit. This is where the `WLINEST` script (`.vba` for Excel and `.gs` for GoogleSheet) comes in. It defines the user function `WLINEST()` that has a similar interface as the built-in `LINEST()` function.

Both scripts also provide some form of help message for the `WLINEST()` function. In Excel this is accessed through the `WLINEST_help()` function. On Google the tooltip will automatically provide such information as you type the function. In addition, if `errorBarChart.gs` is loaded a help message can also be accessed from the `ErrorAnalysis` menu.

Furthermore, the two files with the scripts pre-loaded ([Excel](https://docs.google.com/spreadsheets/d/1EqxN7xtYww0SaZHJSzJkBYhbNKfUlJEP/) and [GoogleSheet](https://docs.google.com/spreadsheets/d/1kUjKvUM2l_IY2ujVOliPjyVWOOYWmDuccRAWg9UFzRg/)) are themselves filled with mock data that illustrate the use of the `WLINEST()` function.
