(STLBinary as binary) =>
let
    DataStart = Binary.Range(STLBinary, 84),
    TriangleFormat = BinaryFormat.Record([
        NormalX = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        NormalY = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        NormalZ = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        V1X = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        V1Y = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        V1Z = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        V2X = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        V2Y = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        V2Z = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        V3X = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        V3Y = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        V3Z = BinaryFormat.ByteOrder(BinaryFormat.Single, ByteOrder.LittleEndian),
        AttributeByteCount = BinaryFormat.ByteOrder(BinaryFormat.SignedInteger16, ByteOrder.LittleEndian)
    ]),
    TriangleList = BinaryFormat.List(TriangleFormat)(DataStart),

    AllPoints = List.Combine(
        List.Transform(TriangleList, each {
            {[V1X],[V1Y],[V1Z]},
            {[V2X],[V2Y],[V2Z]},
            {[V3X],[V3Y],[V3Z]}
        })
    ),
    Xs = List.Transform(AllPoints, each _{0}),
    Ys = List.Transform(AllPoints, each _{1}),
    Zs = List.Transform(AllPoints, each _{2}),
    MinX = List.Min(Xs), MaxX = List.Max(Xs),
    MinY = List.Min(Ys), MaxY = List.Max(Ys),
    MinZ = List.Min(Zs), MaxZ = List.Max(Zs),
    Width  = MaxX - MinX,
    Length = MaxY - MinY,
    Height = MaxZ - MinZ,

    // 全頂点を抽出して重心を計算
    G = {
        List.Average(Xs),
        List.Average(Ys),
        List.Average(Zs)
    },
    // 三角形ごとの面積と体積（重心基準）を計算
    AddCalc = List.Transform(TriangleList, each
        let
            A = {[V1X],[V1Y],[V1Z]},
            B = {[V2X],[V2Y],[V2Z]},
            C = {[V3X],[V3Y],[V3Z]},
            // 相対座標
            Arel = {A{0}-G{0}, A{1}-G{1}, A{2}-G{2}},
            Brel = {B{0}-G{0}, B{1}-G{1}, B{2}-G{2}},
            Crel = {C{0}-G{0}, C{1}-G{1}, C{2}-G{2}},
            // 面積
            AB = {B{0}-A{0}, B{1}-A{1}, B{2}-A{2}},
            AC = {C{0}-A{0}, C{1}-A{1}, C{2}-A{2}},
            Cross = {
                AB{1}*AC{2} - AB{2}*AC{1},
                AB{2}*AC{0} - AB{0}*AC{2},
                AB{0}*AC{1} - AB{1}*AC{0}
            },
            Area = 0.5 * Number.Sqrt(Cross{0}*Cross{0} + Cross{1}*Cross{1} + Cross{2}*Cross{2}),
            // 体積（重心基準）
            Volume = (Arel{0}*(Brel{1}*Crel{2} - Brel{2}*Crel{1})
                    - Arel{1}*(Brel{0}*Crel{2} - Brel{2}*Crel{0})
                    + Arel{2}*(Brel{0}*Crel{1} - Brel{1}*Crel{0})) / 6
        in
            [Area=Area, Volume=Volume]
    ),
    TableCalc = Table.FromList(AddCalc, Splitter.SplitByNothing(), {"Record"}),
    Expanded = Table.ExpandRecordColumn(TableCalc, "Record", {"Area","Volume"}),
    TotalArea = List.Sum(Expanded[Area]),
    TotalVolume = Number.Abs(List.Sum(Expanded[Volume])),
    Weight = (TotalVolume / 1000) * 1.24,
    Cost = Weight * 3,
    Result = #table(
        {"Width[mm]","Length[mm]","Height[mm]","SurfaceArea[mm^2]","Volume[mm^3]","Weight[g]","Cost[yen]"},
        {{Width as number, Length as number, Height as number, TotalArea as number, TotalVolume as number,Weight as number,Cost as number}}
    )
in
    Result
