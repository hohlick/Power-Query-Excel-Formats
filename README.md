# Power-Query-Excel-Formats
A collection of M code to get various formats from Excel sheets in Power Query

## Main purpose
A lot of information, stored in the Excel workbooks, has additional metadata . This metadata could be stored in various forms, mostly as cell formats, number formats, colours, etc.

A wide range of formats and the complexity of extracting their parameters by other tools, such as Power Query, lead to the loss of a noticeable piece of information. Often the format of a row, column or cell is a critical element of the data set. 

At the moment (Aug 2017) the Microsoft Power Query and corresponding "Query Editor" in Microsoft Power BI do not allow users to get additional information, stored in Excel workbooks as various applied formats, except (sometimes) the data types of calculated values.

Additional problem is storing extracted formats data in Power Query for further use
## Work plan

1. Sheet structure: 
    - rows outline levels
    - columns outline levels
    - outline state (collapsed or not)
    - visibility state of rows/columns.
2. Cell indents and alignment
3. Cell number formats
4. Cell color
5. Top-left rows and columns addition to UsedRange/dimension
6. Additional formats and further development
