## Excel.GetNumberFormats 

### Function

**`ExcelGetNumberFormats`**
(
`FullPath` *as text*, 
`SheetName` *as text*, 
*optional* `ColumnN` *as nullable number*, 
*optional* `AddToTable` *as nullable logical*
) as table

### Description

Get cell number formats from specified Excel workbook, worksheet and column. The output is similar to Excel.Workbook, but additional column will be added
		
### Arguments:
		
- **`FullPath`** *(text)*: full *.XLSX/M file path (for example, "C:\Temp\Test.xlsx")
- **`SheetName`** *(text)*: single worksheet name to get formats from (for example, "Sheet1", "Report Table" etc.)
- **`ColumnN`** *(whole number)*: optional column number in R1C1 notation (column C = 3, column D = 5 etc.). Default is 1
- **`AddToTable`** *(logical)*: optional selector, defines whether to add column formats directly to datasheet table (`true`) or as separate table (`false`). Default is `true` 
		
### Examples
**Description:**  Get fromats from 1st column (column "A") on "Sheet1" from file located at "C:\Temp\Test.xlsx" and add them to data table

**Code:** `ExcelGetNumberFormats("C:\Temp\Test.xlsx", "Sheet1", 1)`

**Result:** Formats from column A from Sheet1 worksheet added to Sheet1 data table as "Column1.NumberFormat" column, whole result placed in "DataWithFormats" column
