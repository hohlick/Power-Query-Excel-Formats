// Folder
"here:\goes\folder\path" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]

// FileName
"test2.xlsx" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]

// Sheets
let

    Source = Excel.Workbook(File.Contents(Folder & "\" & FileName), null, true),

    FilteredSheets = Table.SelectRows(Source, each ([Kind] = "Sheet")),

    AddSheetsIndex = Table.AddIndexColumn(FilteredSheets, "Index", 1, 1),

    AddIndPrefix = Table.TransformColumns(AddSheetsIndex, {{"Index", each "rId" & Text.From(_), type text}}),

    MergedRelations = Table.Join(AddIndPrefix,{"Index"},workbook_rels,{"Attribute:Id"}),

    RemovedOther = Table.RemoveColumns(MergedRelations,{"Item", "Kind", "Hidden", "Attribute:Id"}),

    AddRelationTable = Table.AddColumn(RemovedOther, "RowsOutline", each RowsOutline([#"Attribute:Target"])),

    AddedRowsOutline = Table.AddColumn(AddRelationTable, "DataWithOutlines", each Table.Join([RowsOutline], {"Index"}, Table.AddIndexColumn([Data],"Index"), {"Index"}))

in
    AddedRowsOutline

// UnZip
(ZIPFile) => 
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
    Table.FromRecords(Entries)

// UnZipped
let
    Source = Folder.Files(Folder),
    file = Source{[Name = FileName]}[Content],
    Custom1 = UnZip(file)
in
    Custom1

// workbook_rels
let
    Source = UnZipped,
    #"Filtered Rows2" = Table.SelectRows(Source, each [FileName]="xl/_rels/workbook.xml.rels"),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows2", "XML Content", each Xml.Tables([Content],null,1251)),
    #"XML Content1" = #"Added Custom"{0}[XML Content]{[Name="Relationship"]}[Table],
    #"Removed Columns" = Table.RemoveColumns(#"XML Content1",{"Attribute:Type"})
in
    #"Removed Columns"

// RowsOutline
let
    Source = (WSrel as text) => let
    Source = UnZipped,
    Sheet1 = Source{[FileName="xl/" & WSrel]}[Content],
    #"XML Content" = Xml.Tables(Sheet1,null,65001),
    Table = #"XML Content"{[Name="sheetData"]}[Table],
    Filtered = Table.SelectRows(Table, each [Name] = "row"),
    Table1 = Filtered{0}[Table],
    #"Removed Other Columns" = Table.SelectColumns(Table1,{"Attribute:r", "Attribute:outlineLevel"}, MissingField.UseNull),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"Attribute:r", "RowN"}, {"Attribute:outlineLevel", "outlineLevel"}}),
    Typed = Table.TransformColumnTypes(#"Renamed Columns",{{"RowN", Int64.Type}, {"outlineLevel", Int64.Type}}),
    RowIndices = List.Zip( { {List.Min(Typed[RowN])..(List.Max(Typed[RowN]))} }),
    RowIndTable = Table.AddIndexColumn(#table(type table [RowNumber=Int64.Type], RowIndices), "Index"),
    ExpandMissedRows = Table.Join(RowIndTable, {"RowNumber"}, Typed, {"RowN"},JoinKind.LeftOuter),
        #"Removed Other Columns1" = Table.SelectColumns(ExpandMissedRows,{"Index", "outlineLevel"}),
    Result = if Table.IsEmpty(Filtered) then #table(type table [Index = number, outlineLevel = Int64.Type],{}) else #"Removed Other Columns1"
    in
        Result
in
    Source
