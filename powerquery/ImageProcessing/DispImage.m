(bmpTable as table)=>
let
    Table   = Table.Buffer(Table.TransformColumnTypes(bmpTable,{{"Column", Text.Type}})),
    Columns = List.Buffer(List.Distinct(Table[Column])),
    Result  = Table.Sort(Table.Pivot(Table,Columns, "Column", "Val"),{{"Row", Order.Descending}})
in
    Result