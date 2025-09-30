let
    source           = Folder.Files("K:\stl"),
    row_selected     = Table.SelectRows(source, each [Extension] = ".stl"),
    column_added     = Table.AddColumn(row_selected, "stl", each STLReader([Content])),
    column_selected  = Table.SelectColumns(column_added,{"Name", "stl"}),
    column_names     = Table.ColumnNames(column_selected[stl]{0}),
    column_expanded  = Table.ExpandTableColumn(column_selected, "stl", column_names, column_names)
in
    column_expanded