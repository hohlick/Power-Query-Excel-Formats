/*
  // Author: Mark White
  // Reference: http://sql10.blogspot.ru/2016/06/reading-zip-files-in-powerquery-m.html
  // Code Source: https://pastebin.com/KJ8MyAPb
  
  // ZIPFile should be a binary!
*/

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
