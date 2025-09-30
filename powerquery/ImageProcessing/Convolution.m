(bmpTable as table,filter as list)=>
let
    f        = List.Buffer(filter),
    vals     = List.Buffer(bmpTable[Val]),
    clmax    = Table.Max(bmpTable, "Column")[Column],
    rwmax    = Table.Max(bmpTable, "Row")[Row],
    AddIndex = Table.AddIndexColumn(bmpTable, "Index", 0, 1),
    buf      = Table.Buffer(Table.RemoveColumns(AddIndex,{"Val"})),
    cnvl     = Table.AddColumn(buf, "Val", each if [Column] = 0 or [Column] = clmax or [Row] = 0 or [Row] = rwmax then
                                         0
                                     else
                                       + f{0}*vals{[Index] - clmax - 1}
                                       + f{1}*vals{[Index] - clmax} 
                                       + f{2}*vals{[Index] - clmax + 1}
                                       + f{3}*vals{[Index] - 1}     
                                       + f{4}*vals{[Index]}   
                                       + f{5}*vals{[Index] + 1} 
                                       + f{6}*vals{[Index] + clmax - 1} 
                                       + f{7}*vals{[Index] + clmax} 
                                       + f{8}*vals{[Index] + clmax + 1}),


    Result  = Table.TransformColumnTypes(Table.RemoveColumns(cnvl,{"Index"}),{{"Val", type number}})
in
    Result
