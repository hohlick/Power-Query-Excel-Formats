/*
	Author: Maxim Zelensky, info@excel-inside.pro, http://excel-inside.pro
	Date Updated: 2017-08-14
	Purpose: Get outline levels of rows from Excel worksheets in Power Query.
	Note: Included copy of Mark White's UnZip function.

// Parameters:

// FullPath
	Type: text, 
	Description: full path to workbook. Mandatory
	Example: "C:\PQ\Outline\test2.xlsx"
  
// SheetNames
	Type: any
	Description: text or list of worksheet names. Optional
		If argument: 
			not provided, 
			or null,
			or empty list {}, 
			or argument type is different from text/list, 
		then all worksheets from workbook will be analyzed.
    
	Example: {"Sheet1", "Sheet3"}
    Example: "Sheet1"
    
// AddOutlinesToData
	Type: nullable logical
	Description: defines whether add outlineLevel column to the sheet [Data] table. Optional
    	If null or not provided then = true
	Example: true, false, null

*/

// fnGetRowsOutline

(FullPath as text, optional SheetNames as any, optional AddOutlinesToData as nullable logical) as table =>

let


/*
/////////////////////////////////////////////////////////////////////////////////////////////
// Functions
/////////////////////////////////////////////////////////////////////////////////////////////
*/

// UnZip function copy

fnUnZip = (ZIPFile) => 
let
    Header = BinaryFormat.Record([
        MiscHeader = BinaryFormat.Binary(14),
        BinarySize = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
        FileSize   = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
        FileNameLen= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian),
        ExtrasLen  = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian)    
    ]),

    HeaderChoice = BinaryFormat.Choice(
        BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
        each if _ <> 67324752             // not the IsValid number? then return a dummy formatter
            then BinaryFormat.Record([IsValid = false, Filename=null, Content=null])
            else BinaryFormat.Choice(
                    BinaryFormat.Binary(26),      // Header payload - 14+4+4+2+2
                    each BinaryFormat.Record([
                        IsValid  = true,
                        Filename = BinaryFormat.Text(Header(_)[FileNameLen]), 
                        Extras   = BinaryFormat.Text(Header(_)[ExtrasLen]), 
                        Content  = BinaryFormat.Transform(
                            BinaryFormat.Binary(Header(_)[BinarySize]),
                            (x) => try Binary.Buffer(Binary.Decompress(x, Compression.Deflate)) otherwise null
                        )
                        ]),
                        type binary                   // enable streaming
                )
    ),

    ZipFormat = BinaryFormat.List(HeaderChoice, each _[IsValid] = true),

    Entries = List.Transform(
        List.RemoveLastN( ZipFormat(ZIPFile), 1),
        (e) => [FileName = e[Filename], Content = e[Content] ]
    )
in
    Table.FromRecords(Entries),

// RowsOutline
    fnRowsOutline = (WSrel as text) => 
        let
            Source = UnZipped,
            Sheet1 = Source{[FileName="xl/" & WSrel]}[Content],
            XMLContent = Xml.Tables(Sheet1,null,65001),
            sheetDataTable = XMLContent{[Name="sheetData"]}[Table],
            Filtered = Table.SelectRows(sheetDataTable, each [Name] = "row"),
            rowTable = Filtered{0}[Table],
            RemovedOtherColumns = Table.SelectColumns(rowTable,{"Attribute:r", "Attribute:outlineLevel"}, MissingField.UseNull),
            RenamedColumns = Table.RenameColumns(RemovedOtherColumns,{{"Attribute:r", "RowN"}, {"Attribute:outlineLevel", "outlineLevel"}}),
            Typed = Table.TransformColumnTypes(RenamedColumns,{{"RowN", Int64.Type}, {"outlineLevel", Int64.Type}}),
            RowIndices = List.Zip( { {List.Min(Typed[RowN])..(List.Max(Typed[RowN]))} }),
            RowIndTable = Table.AddIndexColumn(#table(type table [RowNumber=Int64.Type], RowIndices), "Index"),
            ExpandMissedRows = Table.Join(RowIndTable, {"RowNumber"}, Typed, {"RowN"},JoinKind.LeftOuter),
            RemovedOtherColumns1 = Table.SelectColumns(ExpandMissedRows,{"Index", "outlineLevel"}),
            Result = if Table.IsEmpty(sheetDataTable) then #table(type table [Index = number, outlineLevel = Int64.Type],{}) else RemovedOtherColumns1
        in
            Result,
/////////////////////////////////////////////////////////////////////////////////////////////


    Source = Excel.Workbook(File.Contents(FullPath), false, true),

// leave sheets only
    FilteredSheets = Table.SelectRows(Source, each ([Kind] = "Sheet")),

// sheets in PQ initially in appearance order, i.e. sheets index (despite visibility)
    AddSheetsIndex = Table.AddIndexColumn(FilteredSheets, "Index", 1, 1),

// check SheetNames parameter
    SheetNames = if SheetNames is text then {SheetNames} else if SheetNames is list then SheetNames else null,

// filter sheets by name if provided
    FilteredByNames = if SheetNames = null or List.IsEmpty(SheetNames) then AddSheetsIndex else Table.SelectRows(AddSheetsIndex, each List.Contains(SheetNames, [Name])),

// UnZip file
UnZipped = Table.Buffer(fnUnZip(File.Contents(FullPath))),
/*
    let
        Source = Folder.Files(Folder),
        file = Source{[Name = FileName, Folder Path = Folder & "\"]}[Content],
        UnZippedFile = Table.Buffer(fnUnZip(file))
    in
        Table.Buffer(UnZippedFile),
*/

// relations id table for sheets
    workbook =
         let
            Source = UnZipped,
            Content = Source{[FileName ="xl/workbook.xml"]}[Content],
            ImportedXML = Xml.Tables(Content,null,TextEncoding.Utf8),
            sheetsTable = ImportedXML{[Name = "sheets"]}[Table],
            sheetTable = sheetsTable{[Name = "sheet"]}[Table],
            ExpandedRel = Table.ExpandTableColumn(sheetTable, "http://schemas.openxmlformats.org/officeDocument/2006/relationships", {"Attribute:id"}, {"Attribute:id"}),
            typed = Table.TransformColumnTypes(ExpandedRel,{{"Attribute:name", type text}, {"Attribute:sheetId", Int64.Type}, {"Attribute:id", type text}})
        in
            typed,

// sheets relations id to XML target files
    workbook_rels = 
        let
            Source = UnZipped,
            Filtered = Table.SelectRows(Source, each [FileName]="xl/_rels/workbook.xml.rels"),
            GetXML = Table.TransformColumns(Filtered, {"Content", each Xml.Tables(_,null,65001)}),
            XMLContent = GetXML{0}[Content]{[Name="Relationship"]}[Table],
            FilteredSheetsRel = Table.SelectRows(XMLContent, each [#"Attribute:Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
            Removed = Table.RemoveColumns(FilteredSheetsRel,{"Attribute:Type"})
        in
            Removed,

// merge relations id (via sheets index)
    MergedRelationsID = Table.Join(FilteredByNames, {"Index"}, workbook, {"Attribute:sheetId"}), 

// join workbook relations
    MergedRelationsTarget = Table.Join(MergedRelationsID,{"Attribute:id"},workbook_rels,{"Attribute:Id"}),

// invoke RowsOutline function to get separate index+outlineLevel table
    AddOutlineTable = Table.AddColumn(MergedRelationsTarget, "RowsOutline", each fnRowsOutline([#"Attribute:Target"]), type table),
    // Clean table
    RemovedColumns = Table.RemoveColumns(AddOutlineTable,{"Index", "Attribute:Id", "Attribute:Target", "Attribute:name", "Attribute:sheetId", "Attribute:id"}),

// join [AddOutlineTable] and [Data] by rows index
    AddedRowsOutline = Table.AddColumn(RemovedColumns, "DataWithOutline", each Table.RemoveColumns(Table.Join([RowsOutline], {"Index"}, Table.AddIndexColumn([Data],"Index"), {"Index"}),{"Index"}), type table),

// add outlines to Data or not
    Result = if AddOutlinesToData or AddOutlinesToData = null then AddedRowsOutline else RemovedColumns
in
    Result
