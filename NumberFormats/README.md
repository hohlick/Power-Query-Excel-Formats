## Excel.GetNumberFormats 

### Code

[Excel.GetNumberFormats.pq](/NumberFormats/Excel.GetNumberFormats.pq)

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
**Description:**  Get fromats from 2nd column (column "B") on "Sample Sheet" from file located at "C:\Temp\Test.xlsx" and add them to data table

Download a [sample file](/NumberFormats/Docs%20and%20samples/ExcelGetNumberFormats%20Sample.xlsx)

![Source data sample](/NumberFormats/Docs%20and%20samples/Sample%20Source%20data.JPG "Source data sample")

This source data contains some currency formats (you can see different EUR, USD and RUB financial and currency formats applied) and custom text formats (defining visible indent of text in the cells). If you try to get this data by built-in Excel.Workbook function, custom number formats will be almost completely lost (except date format):

![Sample load by Excel.Workbook function](/NumberFormats/Docs%20and%20samples/Sample%20ExcelWorkbook%20load.JPG "Sample load by Excel.Workbook function")

**Code:** `ExcelGetNumberFormats("C:\Temp\Test.xlsx", "Sample Sheet", 2)`

**Result:** Formats from column B from `Sample Sheet` worksheet added to `Sample Sheet` data table as ["Column1.NumberFormat"](#note) column, whole result placed in "DataWithFormats" column. This column could be extracted to the table and then parsed as needed:

![Excel.GetNumberFormats Function Output](/NumberFormats/Docs%20and%20samples/Sample%20ExcelGetNumberFormats%20Function%20Output.JPG "Excel.GetNumberFormats Function Output")

#### Note:
As worksheet used range in the sample file starts from column `B`, after import it became `Column1` (empty column `A` ignored). So Number formats column in this case also became `Column1.NumberFormat`.
