Forked from my internship Github account, to my personal Github account

# excelsdmx

Excel add-in for accessing PDH.STAT SDMX API.

This add-in creates a "live" connection between your Excel Workbook and the Pacific Data Hub's .Stat API.

It provides a list of available dataflows (datasets), and choosing one of these dataflows returns all of the corresponding data in a new sheet.

Thanks to Mark from "Excel Off the Grid" for his [tutorial](https://exceloffthegrid.com/inserting-a-dynamic-drop-down-in-ribbon/) on "Dynamic Drop-down Menus", and thanks to the Bank of Italy for their "getTimeSeries" [function](https://github.com/amattioc/SDMX/tree/master/EXCEL).

Excel add-in instructions adapted from [here](https://support.office.com/en-us/article/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460)

## Installation

Download `sdmxpdh.xlam` from this repository, put it somewhere you'll remember

Open a new Excel Workbook

Go `File` > `Options` > `Add-Ins`

In the `Manage` box, click `Excel Add-ins` > `Go`

In the `Add-ins` box which appears, click `Browse`

Navigate to where you saved `sdmxpdh.xlam`, select it and hit `Open`

Check that the add-in's box is "ticked" in the `Add-ins` dialog box

Click `OK`

You should now see `PDH .Stat` in Excel's ribbon menu (among `File`, `Home`, `Data`, `View` etc.)

## Quick start

Click `PDH .Stat` in the ribbon menu

You will see a drop-down menu called `Dataflow` (this has all available dataflows fetched from PDH .Stat)

Choose one of the dataflows to get its data (**a blank window will pop-up temporarily, don't close it (this is the program fetching data)**)

The data will be returned in a new Sheet with the name of the selected dataflow

## Further development

The source files for the add-in are found in the directory `main`. The file `add_in.bas` fetches dataflows and updates the drop-down menu whenever the Workbook is opened. The file `SDMX.bas` is a basic version of the Bank of Italy's `getTimeSeries` function. It is activated when you click on a drop-down menu item, sending a request to the PDH .Stat API for that dataflow's data.

Rebuilding the add-in with new changes is as easy as `File` > `Save as` > then choose `Excel Add-in` for `Save as type`.
