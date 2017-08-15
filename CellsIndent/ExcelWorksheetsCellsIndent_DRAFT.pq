// Draft version 15 Aug 2017

// getIndents
let
// WSrel
WSrel = "worksheets/sheet1.xml", // meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]
// FullPath
FullPath = "\test2.xlsx", // meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]
// SheetNames
//SheetNames = "Sheet1", // meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=false]

////////////
// COMMON //
////////////


Source = Excel.Workbook(File.Contents(FullPath), false, true), 
  
 // leave sheets only 
     FilteredSheets = Table.SelectRows(Source, each ([Kind] = "Sheet")), 
  
 // sheets in PQ initially in appearance order, i.e. sheets index (despite visibility) 
     AddSheetsIndex = Table.AddIndexColumn(FilteredSheets, "Index", 1, 1), 
  
 // check SheetNames parameter 
     SheetNames = if SheetNames is text then {SheetNames} else if SheetNames is list then SheetNames else null, 
  
 // filter sheets by name if provided 
     FilteredByNames = if SheetNames = null or List.IsEmpty(SheetNames) then AddSheetsIndex else Table.SelectRows(AddSheetsIndex, each List.Contains(SheetNames, [Name])), 

UnZip =
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
    Table.FromRecords(Entries),


UnZipped =
let
    UnZippedFile = Table.Buffer(UnZip(File.Contents(FullPath)))
in
    UnZippedFile,  

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
  
 // merge relations id (via sheets name) 
     MergedRelationsID = Table.Join(FilteredByNames, {"Name"}, workbook, {"Attribute:name"}),  
  
 // join workbook relations 
     MergedRelationsTarget = Table.Join(MergedRelationsID,{"Attribute:id"},workbook_rels,{"Attribute:Id"}), 

//////////////////////////////////////////////////////////////////////////////////////////////////////////////


fnA1toR1C1 =
    (dim as text) => let
    Source = Text.Upper(dim),
    Split = Text.Split(Source, ":"),
    TopLeftCell = Split{0},
    BottomRightCell = try Split{1} otherwise null,
    TLC_List = Text.ToList(TopLeftCell),
    BRC_List = Text.ToList(BottomRightCell),

    TopRow = Number.From(Text.Combine(List.Transform(TLC_List, each if Text.Lower(_)=_ then _ else null))),
    BottomRow = if BottomRightCell <> null then Number.From(Text.Combine(List.Transform(BRC_List, each if Text.Lower(_)=_ then _ else null))) else null,

    TopLeftColumnLetters = Text.Combine(List.Transform(TLC_List, each if Text.Lower(_)<>_ then _ else null)),
    BottomRightColumnLetters = Text.Combine(List.Transform(BRC_List, each if Text.Lower(_)<>_ then _ else null)),

    One = {"A".."Z"},
    Two = List.Combine(List.Transform(One, each List.Transform(One, (x) => _ & x))),
    Three = List.Combine(List.Transform(Two, each List.Transform(One, (x) => _ & x))),
    Full = One & Two & Three,

    LeftColumn = List.PositionOf(Full, TopLeftColumnLetters)+1,
    RightColumn = if BottomRightCell <> null then List.PositionOf(Full, BottomRightColumnLetters)+1 else null,

    Result = [TR = TopRow, LC = LeftColumn, BR = BottomRow, RC = RightColumn]
in
    Result,


//fnGetIndents = (WSRel)

xfIndents =
let
    styles = UnZipped{[FileName = "xl/styles.xml"]}[Content],
    ToXML = Xml.Tables(styles, null, 65001),
    cellXfs = ToXML{[Name="cellXfs"]}[Table],
    xf = cellXfs{0}[xf],
    #"Added Index" = Table.AddIndexColumn(xf, "styleID", 0, 1),
    #"Removed Other Columns" = Table.SelectColumns(#"Added Index",{"styleID", "alignment"}, MissingField.UseNull),
    #"Expanded {0}" = Table.ExpandTableColumn(#"Removed Other Columns", "alignment", {"Attribute:horizontal", "Attribute:indent"}, {"Attribute:horizontal", "Attribute:indent"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded {0}",{{"Attribute:indent", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Indent", each if [#"Attribute:horizontal"]="right" then [#"Attribute:indent"]*-1 else [#"Attribute:indent"], Int64.Type),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Attribute:horizontal", "Attribute:indent"})
in
    #"Removed Columns",

OneSheet = 
let
    Sheet = UnZipped{[FileName="xl/" & WSrel]}[Content],
    XMLContent = Xml.Tables(Sheet, null, 65001)
in
    XMLContent,

Cells =
let
//    Sheet = UnZipped{[FileName="xl/" & WSrel]}[Content],
//    XMLContent = Xml.Tables(Sheet, null, 65001),
    Table = OneSheet{[Name="sheetData"]}[Table],
    Rows = Table{[Name="row"]}[Table],
    RemovedColumns = Table.SelectColumns(Rows,{"c", "Attribute:r"}),
    RenamedColumns = Table.RenameColumns(RemovedColumns,{{"Attribute:r", "R1"}}),
    ExpandCells = Table.ExpandTableColumn(RenamedColumns, "c", {"Attribute:r", "Attribute:s"}, {"A1", "styleRef"}),
    toR1C1 = Table.AddColumn(ExpandCells, "R1C1", each fnA1toR1C1([A1])),
    ExpandC1 = Table.ExpandRecordColumn(toR1C1, "R1C1", {"LC"}, {"C1"}),
    #"Changed Type" = Table.TransformColumnTypes(ExpandC1,{{"C1", Int64.Type}, {"styleRef", type number}, {"R1", Int64.Type}, {"A1", type text}}),
    ReplacedNulls = Table.ReplaceValue(#"Changed Type",null,0,Replacer.ReplaceValue,{"styleRef"}),
    AddRowIndex = Table.AddColumn(ReplacedNulls, "RowIndex", each [R1]-sheet_dimension[TR], Int64.Type),
    AddColumnIndex = Table.AddColumn(AddRowIndex, "ColumnIndex", each [C1]-sheet_dimension[LC], Int64.Type),
    #"Merged Queries" = Table.Join(AddColumnIndex,{"styleRef"},xfIndents,{"styleID"},JoinKind.LeftOuter)
in
    #"Merged Queries",



sheet_dimension = 
let
    dimension_ref = OneSheet{[Name="dimension"]}[Table],
    dim = dimension_ref{0}[#"Attribute:ref"],
    dimR1C1 = fnA1toR1C1(dim)
in
    dimR1C1,

IndentsMap =
let
    Source = Cells,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Indent] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "ColumnName", each "Column" & Text.From([ColumnIndex]+1), type text),
    #"Grouped Rows" = Table.Group(#"Added Custom", {"RowIndex", "ColumnName"}, {{"tab", each _{0}[Indent], Int64.Type}}),
    #"Pivoted Column" = Table.Pivot(#"Grouped Rows", List.Sort(List.Distinct(#"Grouped Rows"[ColumnName])), "ColumnName", "tab")
in
    #"Pivoted Column"
in
    IndentsMap

