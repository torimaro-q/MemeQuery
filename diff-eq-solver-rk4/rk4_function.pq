(f as function,h as number,steps as number, initial_values as list)=>
let
    rk4 = (x as list) as list =>
    let
        k1 = f(x),
        xk2 = List.Transform({0..List.Count(x)-1}, each x{_} + (h/2.0) * k1{_}),
        k2 = f(xk2),
        xk3 = List.Transform({0..List.Count(x)-1}, each x{_} + (h/2.0) * k2{_}),
        k3 = f(xk3),
        xk4 = List.Transform({0..List.Count(x)-1}, each x{_} + h * k3{_}),
        k4 = f(xk4),
        next = List.Transform({0..List.Count(x)-1}, each x{_} + (h/6.0) * (k1{_} + 2.0*k2{_} + 2.0*k3{_} + k4{_}))
    in
        List.Buffer(next),

    seq = List.Generate(
        ()=> [step = 0, t = 0.0, val = initial_values],
        each [step] <= steps,
        each [step=[step]+1, t=[t]+h, val=rk4([val])]
    ),

    fieldNames = List.Transform({1..List.Count(initial_values)}, each "x" & Text.From(_)),
    tbl = Table.FromList(seq, Splitter.SplitByNothing(), {"Record"}),

    Expanded1 = Table.ExpandRecordColumn(tbl, "Record", {"step","t","val"}, {"step","t","val"}),

    AddRec = Table.TransformColumns(Expanded1, {"val", each Record.FromList(_, fieldNames)}),
    Expanded2 = Table.ExpandRecordColumn(AddRec, "val", fieldNames),

    result = Table.TransformColumnTypes(
        Expanded2,
        {{"step", type number},{"t", type number}} & List.Transform(fieldNames, each {_, type number})
    )
in
    result
