let
    Source = Folder.Files("D:\Volkswagen AG\Audi AS - To be Archived\AGP AGA Raw Data"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File from Audi AS - AGP AGA Raw Data", each #"Transform File from Audi AS - AGP AGA Raw Data"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File from Audi AS - AGP AGA Raw Data"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File from Audi AS - AGP AGA Raw Data", Table.ColumnNames(#"Transform File from Audi AS - AGP AGA Raw Data"(#"Sample File"))),
    #"Filtered Rows" = Table.SelectRows(#"Expanded Table Column1", each [#"Sold-to party"] = "Audi 南屯" or [#"Sold-to party"] = "Audi 台中" or [#"Sold-to party"] = "Audi 台北旗艦中心 (內湖)" or [#"Sold-to party"] = "Audi 岡山 (台南)" or [#"Sold-to party"] = "Audi 岡山 (高雄)" or [#"Sold-to party"] = "Audi 花蓮" or [#"Sold-to party"] = "Audi 高雄" or [#"Sold-to party"] = "Audi 鳳山" or [#"Sold-to party"] = "PDI (AUDI)" or [#"Sold-to party"] = "台奧北區股份有限公司" or [#"Sold-to party"] = "和順興汽車企業股份有限公司員林營業所" or [#"Sold-to party"] = "奧迪北區股份有限公司南港分公司" or [#"Sold-to party"] = "奧迪南部汽車股份有限公司台南分公司" or [#"Sold-to party"] = "豐仕汽車股份有限公司北投營業所" or [#"Sold-to party"] = "豐仕汽車股份有限公司濱江營業所" or [#"Sold-to party"] = "He Shun Sing"),
    #"Changed Type2" = Table.TransformColumnTypes(#"Filtered Rows",{{"RRP Total Amount", Int64.Type}, {"DNP Total Amount", Int64.Type}}),
    #"Inserted Year" = Table.AddColumn(#"Changed Type2", "Year", each Date.Year([Created on]), Int64.Type),
    #"Inserted Month" = Table.AddColumn(#"Inserted Year", "Month", each Date.Month([Created on]), Int64.Type),
    #"Inserted Last Characters" = Table.AddColumn(#"Inserted Month", "Last Characters", each Text.End([Order], 1), type text),
    #"Added Conditional Column" = Table.AddColumn(#"Inserted Last Characters", "Order ends 689", each if Text.EndsWith([Order], "6") then true else if Text.EndsWith([Order], "8") then true else if Text.EndsWith([Order], "9") then true else false),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Conditional Column",{{"Order ends 689", type logical}, {"Dv", type text}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Type", each if [PGr]= "TW1" and [Order ends 689]= true then "AGP"
else if [PGr]= "ZFI" and [Order ends 689]= true then "AGP"

else if [PGr]="TW2" and [Dv]="94" then "AGA"

else if [PGr]="TW3" and [MG 1] <> "i11" and [MG 1] <> "i12" and [BillT]="ZFVA" then "Local"

else if [PGr]="TW3" and [MG 1] <> "i11" and [MG 1] <> "i12" and [BillT]="ZG2" then "Local"

else ""),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "Tire", each if [PGr]="TW3" and [Order] = "AC285_0102" and [BillT]="ZFVA" or [BillT]="ZG2" then "True"

else "False"),
    #"Grouped Rows" = Table.Group(#"Added Custom1", {"Sold-to pt", "Year", "Month", "Type", "Tire"}, {{"Total DNP", each List.Sum([DNP Total Amount]), type number}, {"Tire Billed Qty", each List.Sum([Billed]), type number}}),
    #"Merged Queries" = Table.NestedJoin(#"Grouped Rows",{"Sold-to pt"},#"Workshop Name",{"Sold-to pt"},"Workshop Name",JoinKind.Inner),
    #"Expanded Workshop Name" = Table.ExpandTableColumn(#"Merged Queries", "Workshop Name", {"Workshop"}, {"Workshop"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded Workshop Name",{"Workshop", "Sold-to pt", "Year", "Month", "Type", "Tire", "Total DNP", "Tire Billed Qty"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Reordered Columns",{{"Tire", type logical}}),
    #"Appended Query" = Table.Combine({#"Changed Type1", Adjust})
in
    #"Appended Query"
