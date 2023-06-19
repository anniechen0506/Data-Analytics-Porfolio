# -*- coding: utf-8 -*-

import pandas as pd

#read excel file
excelPath = r"D:\Volkswagen AG\Audi AS - 1.3.1 CASIS report\2023 CASIS\202305\CASIS_Monthly KPI_202305 (v.20230601).xlsx"
xls = pd.ExcelFile(excelPath)
sheetNameList = xls.sheet_names
df = pd.DataFrame()

# for loop every sheet name in the Excel file
# transform every sheets into Dataframe
sheetDataDict = dict()
for i in sheetNameList[1:]:
    fh_tmp = pd.concat([pd.read_excel(excelPath, sheet_name = i).assign(sheet_names = i) for i in sheetNameList[1:]])
    sheetDataDict[i] = fh_tmp
    col = fh_tmp.iloc[2]
    col[0] = "Attribute"
    col[1] = "Detail"
    col[-1] = "Outlet"
    fh_tmp.columns = col
    fh_tmp = fh_tmp.drop([0,1,2]) # drop the first three column
    first_column = fh_tmp.pop("Outlet")
    fh_tmp.insert(0, "Outlet", first_column) # move the additional Outlet to the first column
    fh_tmp = fh_tmp.iloc[:, :16]
    fh_tmp.iloc[:, 1] = fh_tmp.iloc[:, 1].fillna(method = 'ffill') # Down fill the first column
    fh_tmp.iloc[:, 1:] = fh_tmp.iloc[:, 1:].fillna(0)
    #fh_tmp = fh_tmp.reset_index(drop = True)
    df = pd.concat([df, fh_tmp], axis = 0, ignore_index = True) # Concat all dataframe into a Dictionary
    df = df.replace("Summary", "National")


df_unpivoted = df.melt(id_vars = ['Outlet', 'Attribute', "Detail"], var_name = 'Month', value_name = 'value')
df_unpivoted = df_unpivoted.assign(year = '2023')

df_unpivoted.to_excel(r'D:\Volkswagen AG\Audi BI - Reference\Raw Data\AS\AS_CASIS_report_202305.xlsx', sheet_name='CASIS', index=False)

