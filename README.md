[README на русском](README_RUS.md)

# Power-Query-Excel-Formats
A collection of M code to get various formats from Excel sheets in Power Query

## Main purpose

Information, stored in the Excel workbooks, often has additional metadata, important for analyzis. This metadata could be stored in various forms, mostly as cell formats, number formats, colours, etc. Often a row, column or cell format is a critical element of the workbook data set.

At the moment (Aug 2017) the Microsoft Power Query and corresponding "Query Editor" in Microsoft Power BI do not allow users to get additional information (stored in Excel workbooks and spreadsheets as various applied formats) natively, except (sometimes) the data types of calculated values.

A wide range of formats and the complexity of extracting their parameters by other tools, such as Power Query, lead to the loss of a noticeable piece of information. Additional problem is storing extracted formats data in Power Query for further use.
Задачи и методы

## Tasks

Develop a set of functions to extract/import specific info about sheet and/or cell formats into Power Query.

In the future - develop universal functions:

* spreadsheet information (info about rows, columns, sheet in whole)
* cells info (colors, fonts, alignment, number formats, indents etc.)

The versatility of the methods due to the same tools (unzip and XML parsing) and the similarity of data sources. Specific kind of function result can be selected via function argument.

---

### Methods

#### Unzip

Main method is unpacking of XLSX/XLSM as zip and working with XML documents inside. Unpack performed via custom function [UnZip.pq](UnZip.pq) by Mike White. But any other analogue to unpack zip archives in Power Query can be used.

#### XML Parsing

After UnZip the XML files (`binary` type) from workbook structure become available for the (current) main function. Possible parse methods - with built-in functions `Xml.Tables` or `Xml.Document`, or with other suitable XML parsing methods.

* Main problem: cell formats stored separate from cells, cells itself stored inside row element, cell address stored in A1 notation (need additional convert to R1C1-style or similar).
* Additional problem: linking/mapping extracted format info with cell position in Power Query table.

---
## Work plan

1. Sheet structure: 
    - [rows outline levels](../../tree/master/RowsOutline),
    - columns outline levels,
    - extended rows state (visibility, spans, outlines, collapsed, etc.),
    - extended columns state.
2. Cell indents and alignment
3. Cell number formats
4. Cell color
5. Top-left rows and columns addition to UsedRange/dimension (see this [post about UsedRange pitfall](http://excel-inside.pro/blog/2017/05/23/excel-sheet-as-a-source-to-power-query-and-power-bi-a-pitfall-of-usedrange/))
6. Additional formats, conditional formats and further development
