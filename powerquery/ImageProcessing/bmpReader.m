(fpath,optional color)=>
let    
    BYTE2LONG = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
    BYTE2INT  = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian),

    BITMAPFILEHEADER = BinaryFormat.Record([
        etc1 = BinaryFormat.Binary(18),
        biWidth= BinaryFormat.Binary(4),
        biHeight = BinaryFormat.Binary(4),
        biPlanes = BinaryFormat.Binary(2),
        biBitCount = BinaryFormat.Binary(2),
        etc2 = BinaryFormat.Binary(24),
        biData = BinaryFormat.Binary()
    ]),

    imageData = BinaryFormat.ByteOrder(BinaryFormat.Byte, ByteOrder.LittleEndian),

    GetPixels = BinaryFormat.List(imageData),

    FileHeader = BITMAPFILEHEADER(File.Contents(fpath)),
    BitCount   = BYTE2INT(FileHeader[biBitCount]),
    ByteCount  = (BitCount / 8),
    
    width      = BYTE2LONG(FileHeader[biWidth]),
    Amari      = if Number.Mod(ByteCount * width,4) = 0 then 0 else 4 - Number.Mod(ByteCount * width,4),

    height     = BYTE2LONG(FileHeader[biHeight]),

    buffer     = List.Buffer(GetPixels(FileHeader[biData])),
    count      = List.Count(buffer),

    Entries    = List.Buffer(List.Generate(()=>
                        [
                            Index  = 0,
                            Row    = 0,
                            Column = 0,
                            Val    = GrayScale(buffer,Index,color)
                        ],
                 each 
                        [Index] < count,
                 each 
                        if [Column] + 1 = width then
                        [
                            Index  = [Index] + ByteCount + Amari,
                            Row    = [Row]   + 1,
                            Column = 0,
                            Val    = GrayScale(buffer,Index,color)
                        ]
                        else
                        [
                            Index  = [Index] + ByteCount,
                            Row    = [Row],
                            Column = [Column] + 1,
                            Val    = GrayScale(buffer,Index,color)                
                        ]

    )),

    GrayScale = (buf,index,optional color) => 
            if color = 0 then 
                (buf{index}+buf{index+1}+buf{index+2})*0.333
            else if color = 1 then
                buf{index}
            else if color = 2 then
                buf{index+1}
            else if color = 3 then
                buf{index+2}
            else
                buf{index+3},
    
    List2Table  = Table.FromList(Entries, Splitter.SplitByNothing(), null, null, ExtraValues.Error),

    Expanded    = Table.ExpandRecordColumn(List2Table, "Column1", {"Index", "Row", "Column", "Val"}, {"Index", "Row", "Column", "Val"}),
    Removed     = Table.RemoveColumns(Expanded,{"Index"}),

    Transformed = Table.TransformColumnTypes(Removed,{{"Val", Number.Type},{"Column", Int64.Type},{"Row", Int64.Type}}),
    Result      = Table.Buffer(Table.Distinct(Transformed, {"Row", "Column"}))
in
    Result
