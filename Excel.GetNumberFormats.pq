/*
    Name =  Excel.GetNumberFormats, 
    Description = 
		Get cell number formats from specified Excel workbook, worksheet and column. The output is similar to Excel.Workbook, but additional column will be added
		
        Arguments:
		
        FullPath (text): full *.XLSX/M file path (for example, "C:\Temp\Test.xlsx")
		
        SheetName (text): single worksheet name to get formats from (for example, "Sheet1", "Report Table" etc.)
		
        ColumnN (whole number): optional column number in R1C1 notation (column C = 3, column D = 5 etc.). Default is 1
		
        AddToTable (logical): optional selector, defines whether to add column formats directly to datasheet table (true) or as separate table (false). Default is true 
		
	Category = Accessing data functions, 
    Source = local, 
    Author = Maxim Zelensky: http://Excel-Inside.pro 
	Project = https://github.com/hohlick/Power-Query-Excel-Formats
    Examples = 
        Description:  Get fromats from 1st column (""A"") on Sheet1 from file located at C:\Temp\Test.xlsx and add them to data table , 
        Code: ExcelGetNumberFormats(""C:\Temp\Test.xlsx"", ""Sheet1"", 1), 
        Result: "Formats from column A from Sheet1 worksheet added to Sheet1 data table as ""Column1.NumberFormat"" column, whole result placed in ""DataWithFormats"" column"
*/
let 
functionbody = 
(	FullPath as text, 
	SheetName as text, 
	optional ColumnN as nullable number, 
	optional AddToTable as nullable logical
) as table => 

let

// Parameters check
SheetName = {SheetName}, 
/* list conversion for further sheets list compatibility
 if SheetName is list then SheetName else 
	if not (try Text.From(SheetName))[HasError] then {SheetName} else  if SheetName is null then {null}
*/
ColumnN = if ColumnN = null then 1 else ColumnN, // 1 by default
AddToTable = if AddToTable = null then true else AddToTable,


// UnZip
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

// fnColumnA1Number
fnColumnA1Number = 
	(A1ref as text) as record => 
	
	let
	//    A1ref = "XFD125",
	    Source = Text.Upper(A1ref),
	    ToList = Text.ToList(Source),
	    RowN = Number.From(
	        Text.Combine(
	            List.Transform(
	                ToList, 
	                each if Text.Lower(_)=_ then _ else null)
	            )
	        ) as number,
	
	    ColumnLettersRev = 
	        List.Buffer(
	            List.Reverse(
	                List.RemoveNulls(
	                    List.Transform(ToList, each if Text.Lower(_)<>_ then _ else null) ) ) ),
	
	    ListAZ = List.Buffer({"A".."Z"}),
	    ColumnN = 
	        List.Accumulate(
	            List.Positions(ColumnLettersRev), 
	            0, 
	            (s, c) => s + (List.PositionOf(ListAZ, ColumnLettersRev{c} ) + 1) * ( {1,26,676}{c} ) 
	        ) as number
	in
	    [Row = RowN, Column = ColumnN],

// UnZipped

    UnZipped = Table.Buffer(UnZip(File.Contents(FullPath))),

// fnLinkFormatsToCells
fnLinkFormatsToCells = (SheetRef) as table =>
	let
		// styles
		styles =
			let
			    Source = UnZipped{[FileName="xl/styles.xml"]}[Content],
			    ImportedXML = Table.Buffer(Xml.Tables(Source, null, 65001))
			in
			    ImportedXML,
		
		// cellXfs
		cellXfs = 
			let
			    Source = styles,
			    Table = Source{[Name = "cellXfs"]}[Table],
			    xf = Table{0}[xf],
			    AddedStyleID = Table.AddIndexColumn(xf, "StyleID", 0, 1),
			    ChangedType = Table.TransformColumnTypes(AddedStyleID,{{"Attribute:numFmtId", type number}}),
			    Merged_numFmts = Table.Join(ChangedType, {"Attribute:numFmtId"}, numFmts, {"numFmtId"}),
			    SelectedColumns = Table.SelectColumns(Merged_numFmts,{"StyleID", "formatCode"})
			in
			    SelectedColumns,
		
		// numFmts
		numFmts = 
			let
			
			    numFmtsTable = styles{[Name = "numFmts"]}[Table][numFmt]{0},
			    Renamed = Table.RenameColumns(numFmtsTable,{{"Attribute:numFmtId", "numFmtId"}, {"Attribute:formatCode", "formatCode"}}),
			    ChangedType = Table.TransformColumnTypes(Renamed,{{"numFmtId", type number}, {"formatCode", type text}}),
			    numFmts_builtin =
			        let
			            Source = Table.FromRows(
				            Json.Document(
				                Binary.Decompress(
									Binary.FromText(
	"VZE7D8IgEID/SkNjorEP2lL7cKgOxsnEuBIGE5p0EAc78e8F7hCcet93jxxXzgklGbnO7/nzfBGRcVIZpi6qbVRQgMZAmqUpEPPk04Ot3cAAin2IFeJlj7WVHZwmUzkBNoCGUdjpSuVS5lqDaY2RuTLSm4M3gJ1rCene8DIqlZxv5f0GbkA3rmuka4oaqApFIOyyqpSl1smvqun8+5Pj1n134Pvg+WOW4i85REfDNhNBktE4GfWGitY90e/F7AH4IqJVGRxhXQv8SW4bO8KcHoxd4USE+AI=", 
										BinaryEncoding.Base64), 
				                    Compression.Deflate)
				                ), 
				            	let _t = ((type text) meta [Serialized.Text = true]) in type table [numFmtId = _t, formatCode = _t]),
			        	Typed = Table.TransformColumnTypes(Source, {{"numFmtId", type number}})
			    	in
			        	Typed,
	
			    AppendedBiultIn = Table.Combine({numFmts_builtin, ChangedType})
			in
			    AppendedBiultIn,
	
		// LinkFormatsToCells itself 
	    Source = Xml.Tables(UnZipped{[FileName = SheetRef]}[Content]),
	
	    dim = Source{[Name="dimension"]}[Table][#"Attribute:ref"]{0},
	    TLC = fnColumnA1Number(Text.Split(dim, ":"){0}),
	    GetRows = Source{[Name="sheetData"]}[Table],
	    GR2 =Table.SelectColumns(GetRows[Table]{0}, {"c", "Attribute:r"}),
	    Renamed = Table.RenameColumns(GR2, {{"Attribute:r", "row"}}),
	    Expanded  = Table.ExpandTableColumn(Renamed, "c", {"Attribute:r", "Attribute:s"}, {"cell", "StyleID"}),
	    GetColumnNumber = Table.AddColumn(Expanded, "column", each fnColumnA1Number([cell])[Column], type number),
	    FilteredByColumnNumber = Table.SelectRows(GetColumnNumber, each ([column] = ColumnN)),
	    ChangedType = Table.TransformColumnTypes(FilteredByColumnNumber,{{"StyleID", type number}, {"row", type number}, {"cell", type text}}),
	    AddRowIndex = Table.AddColumn(ChangedType, "rowIndex", each [row]-TLC[Row], type number),
	    MergeByStyleRef = Table.Join(AddRowIndex, {"StyleID"}, cellXfs, {"StyleID"}),
	    Removed = Table.SelectColumns(MergeByStyleRef,{"rowIndex", "formatCode"}),
	    Sorted = Table.Sort(Removed,{{"rowIndex", Order.Ascending}}),
	    Result = if Table.IsEmpty(GetRows) 
	        then 
	            #table(
	                type table [
	                    rowIndex = number, 
	                    formatCode=text
	                    ],
	                {}) 
	        else Sorted
	in
    	Table.RenameColumns(Result,{{"formatCode", "Column" & Text.From(ColumnN) & ".NumberFormat"}}),

// SourceSheets
SourceSheets = 
	let
	    Source = Excel.Workbook(File.Contents(FullPath), null, true),
	    SheetsOnly = Table.SelectRows(Source, each ([Kind] = "Sheet")),
	    SheetsIndex = Table.AddIndexColumn(SheetsOnly, "Index", 1, 1)
	in
	    Table.SelectRows(SheetsIndex, each List.Contains(SheetName, [Name])),
// 
SheetsWithFormats = 
	let
	    AddSheetRef = Table.TransformColumns(SourceSheets, {{"Index", each "xl/worksheets/sheet" & Text.From(_) & ".xml"}}),
	    DataWithFormats = Table.AddColumn(
			AddSheetRef, 
			"DataWithFormats", 
			each Table.RemoveColumns(
				Table.Join(
					fnLinkFormatsToCells([Index]), 
					{"rowIndex"}, 
					Table.AddIndexColumn([Data], "Index", 0, 1), 
					{"Index"}, 
					JoinKind.RightOuter
					), 
				{"rowIndex", "Index"}
			)
		),
	    ColumnFormatsOnly = Table.AddColumn(
			AddSheetRef, 
			"Column" & Text.From(ColumnN) &".Formats", 
			each fnLinkFormatsToCells([Index])
		)
	in
		if AddToTable then DataWithFormats else ColumnFormatsOnly,

Result = if Table.IsEmpty(SourceSheets) then SourceSheets else SheetsWithFormats
in 
    Result,
Description = [
    Documentation.Name =  "Excel.GetNumberFormats", 
    Documentation.Description = "Get cell number formats from specified Excel workbook, worksheet and column" , 
    Documentation.LongDescription =  "<p>Get cell number formats from specified Excel workbook, worksheet and column. The output is similar to <code>Excel.Workbook</code>, but additional column will be added</p>
        <p><b>Arguments:</b></p>
        <p><b><code>FullPath</code></b> <code>(text)</code>: full *.XLSX/M file path (for example, <code>""C:\Temp\Test.xlsx""</code>)<br>
        <b><code>SheetName</code></b> <code>(text)</code>: single worksheet name to get formats from (for example, <code>""Sheet1""</code>, <code>""Report Table""</code> etc.)<br>
        <b><code>ColumnN</code></b> <code>(whole number)</code>: <i>optional</i> column number in R1C1 notation (column C = 3, column D = 5 etc.). Default is 1<br>
        <b><code>AddToTable</code></b> <code>(logical)</code>: <i>optional</i> selector, defines whether to add column formats directly to datasheet table (<code>true</code>) or as separate table (<code>false</code>). Default is <code>true</code></p>
		<p>Part of the GitHub project https://github.com/hohlick/Power-Query-Excel-Formats</p>", 
    Documentation.Category = "Accessing data functions", 
    Documentation.Source = "local", 
    Documentation.Author = "Maxim Zelensky: http://Excel-Inside.pro", 
    Documentation.Examples = {
        [Description =  "Get fromats from 1st column (""A"") on <code>Sheet1</code> from file located at <code>C:\Temp\Test.xlsx</code> and add them to data table" , 
        Code = "ExcelGetNumberFormats(""C:\Temp\Test.xlsx"", ""Sheet1"", 1)", 
        Result = "Formats from column A from Sheet1 worksheet added to Sheet1 data table as ""Column1.NumberFormat"" column, whole result placed in ""DataWithFormats"" column"]
        }
    ] 
in 
    Value.ReplaceType(functionbody, Value.ReplaceMetadata(Value.Type(functionbody), Description))