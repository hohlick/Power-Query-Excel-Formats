# Rows Outline

*Easiest part* :)

## Idea

As Power Query import data from Excel spreadsheet into a table structure, then row outline levels could be stored as additional column for such tables.

***

## Realisation

    fnGetRowsOutline(
        FullPath as text, 
        optional SheetNames as any, 
        optional AddOutlinesToData as nullable logical
        ) as table

Sample is [there](#sample)

### Code:
[ExcelWorksheetsRowOutlines.pq](../RowsOutline/ExcelWorksheetsRowOutlines.pq)

### Description:

Returns spreadsheets (not tables) data from Excel workbook (xlsx or xlsm tested), adding information about rows outline levels.
As rows outline levels is the property of rows (not cells), it is possible to return outline level for each used row.

Based on `Excel.Workbook` built-in function, but adds (one or two, depending on the third argument) additional columns to its result:

* `RowsOutline` column with a table of two columns: 
    * `RowIndex` as number, (zero-based) - an index to further relations to [Data] column contents
    * `outlineLevel` as Int64.Type
* `DataWithOutline` column, where `outlineLevel` column is added as the first column to raw sheet data (`Excel.Workbook` `[Data]` column).

### Function arguments:

#### `FullPath`

*Type:* text,

*Description:* full path to workbook. **Mandatory**

*Example:* "C:\PQ\Outline\test2.xlsx"
  
#### `SheetNames`

*Type*: any
   
*Description*: text or list of worksheet names to extract. **Optional**
   
   If argument: ExcelWorksheetsRowOutlines_TesterDataSet.xlsx 
   
   - not provided,
   - or null,
   - or empty list {}, 
   - or argument type is different from text/list, 

then all worksheets from workbook will be analyzed.
    
*Example*: 
* {"Sheet1", "Sheet3"}
* "Sheet1"


#### `AddOutlinesToData`

*Type*: nullable logical

*Description*: defines whether add outlineLevel column to the sheet [Data] table. **Optional**

If null or not provided then `true`
        
*Example*: 
* true, 
* false, 
* null

***

## Notes:
1. Included copy of [Mark White's UnZip function](../UnZip.pq).
2. Both functions (`Excel.Workbook` and `ExcelWorksheetsRowOutlines`) return cells range from worksheet, based on `UsedRange` VBA property (or `dimension` sheet atteribute in SpreadsheetML schema).

***

## [Sample](#sample)
There two files:
* [`ExcelWorksheetsRowOutlines_TesterDataSet.xlsx`](ExcelWorksheetsRowOutlines_TesterDataSet.xlsx), which contains simple workbook with some rows grouping
* [`ExcelWorksheetsRowOutlines_Tester.xlsx`](ExcelWorksheetsRowOutlines_Tester.xlsx), which contains code and several queries to show how function works

Download both files, place them somewhere, and write down a full path to the `ExcelWorksheetsRowOutlines_TesterDataSet.xlsx` in the `ExcelWorksheetsRowOutlines_Tester.xlsx ` (see parameter table inside), then refresh loaded query. You can also look on the other possible function results in Power Query Editor.
