let
    //Use Relative path for M Language to ensure it works on every person's computer
    MyFolder = Excel.CurrentWorkbook(){[Name = "Registration"]}[Content][Column1]{0},
    Source = Folder.Files(MyFolder),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File from Registration", each #"Transform File from Registration"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File from Registration"}),
    #"Expanded Transform File from Registration" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File from Registration", {"Name", "Data", "Item", "Kind", "Hidden"}, {"Name", "Data", "Item", "Kind", "Hidden"}),
    #"Filtered Rows1" = Table.SelectRows(#"Expanded Transform File from Registration", each ([Hidden] = false) and ([Name] = "領牌" or [Name] = "領牌明細")),
    #"Expanded Data" = Table.ExpandTableColumn(#"Filtered Rows1", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "SN", "領牌日", "型式", "訂單編號", "車身號碼", "車牌", "車主名稱", "車主身份證/統一編號字號", "出生#(lf)年月日", "姓別", "郵遞區號", "地址", "聯絡住址", "住家#(lf)電話", "公司電話", "行動#(lf)電話", "出廠年月", "使用人名稱", "銷售人員", "經銷商", "showroom", "備註:#(lf)(客戶來源: A/B/C)", "備註(二)"}, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "SN", "領牌日", "型式", "訂單編號", "車身號碼", "車牌", "車主名稱", "車主身份證/統一編號字號", "出生#(lf)年月日", "姓別", "郵遞區號", "地址", "聯絡住址", "住家#(lf)電話", "公司電話", "行動#(lf)電話", "出廠年月", "使用人名稱", "銷售人員", "經銷商", "showroom", "備註:#(lf)(客戶來源: A/B/C)", "備註(二)"}),
    #"Filtered Rows2" = Table.SelectRows(#"Expanded Data", each ([Column4] <> null)),
    #"Promoted Headers1" = Table.PromoteHeaders(#"Filtered Rows2", [PromoteAllScalars=true]),
    #"Removed Columns" = Table.RemoveColumns(#"Promoted Headers1",{"202209_September registration.xlsx", "領牌", "Column3", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46", "Column47", "Column48"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns", each ([領牌日] <> "領牌日")),
    #"Added Conditional Column" = Table.AddColumn(#"Filtered Rows", "Category", each if Text.StartsWith([訂單編號], "TPNFU") then "Function Car Sales" else if Text.StartsWith([訂單編號], "TNSFU") then "Function Car Sales" else if Text.StartsWith([訂單編號], "TFSFUN") then "Function Car Sales" else if Text.StartsWith([訂單編號], "KSSFN") then "Function Car Sales" else if [車主名稱] = "和順利汽車股份有限公司" then "Function Car Sales" else if [使用人名稱] = "和順利汽車股份有限公司" then "Function Car Sales" else if [經銷商] = "VGT" then "Function Car Sales" else "Retail Car Sales"),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Conditional Column",{{"領牌日", type date}}),
    #"Inserted Year" = Table.AddColumn(#"Changed Type", "Year", each Date.Year([領牌日]), Int64.Type),
    #"Inserted Month" = Table.AddColumn(#"Inserted Year", "Month", each Date.Month([領牌日]), Int64.Type),
    #"Grouped Rows" = Table.Group(#"Inserted Month", {"經銷商", "showroom", "Year", "Month", "Category"}, {{"Sales Actual", each Table.RowCount(_), type number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Grouped Rows",{{"經銷商", "Investor"}, {"showroom", "Outlet"}}),
in
    #"Renamed Columns"
