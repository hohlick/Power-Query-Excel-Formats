# Rows Outline

Easiest part :)

## Idea

As Power Query import data from Excel spreadsheet into a table structure, then row outline levels could be stored as additional column for such tables.

### Mandatory arguments

* `FullPath` as text - to UnZip XLSX structure and gain access to SpreadsheetML elements.
* `SheetName` as text - to get a specific sheet data directly or via iteration.

### Optional arguments

* `separate` as nullable logical - to select output type.

If `true`, function output is a table of two columns:
    * `RowIndex` as number, (zero-based)
    * `outlineLevel` as Int64.Type
  
If `false`, function will add such columns to an imported spreadsheet
Default value is `true` *(??)*
