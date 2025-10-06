(STLFile) => 
let
    Vertex = BinaryFormat.Record([
        NX1 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        NY1 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        NZ1 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        X1 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        Y1 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        Z1 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        X2 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        Y2 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        Z2 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        X3 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        Y3 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        Z3 = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        Blank = BinaryFormat.ByteOrder(BinaryFormat.SignedInteger16, ByteOrder.LittleEndian)
    ]),
    GetData = BinaryFormat.List(Vertex),
    Entries = List.Transform(GetData(Binary.Range(STLFile,84)),
                            (e) => 
                                {
                                    [I=1,X=e[X1],Y=e[Y1],Z=e[Z1]],
                                    [I=2,X=e[X2],Y=e[Y2],Z=e[Z2]],
                                    [I=3,X=e[X3],Y=e[Y3],Z=e[Z3]]
                                }
                            ),
    Expanded = Table.ExpandListColumn(Table.FromList(Entries, Splitter.SplitByNothing(), null, null, ExtraValues.Error),"Column1"),
    Expanded2 = Table.ExpandRecordColumn(Expanded, "Column1", {"I", "X", "Y", "Z"}, {"I", "X", "Y", "Z"}),
    Transformed = Table.TransformColumnTypes(Expanded2,{{"I", Int64.Type}, {"X", type number}, {"Y", type number}, {"Z", type number}})
in
    Transformed
