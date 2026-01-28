
import pandas as pd
from sqlalchemy import create_engine, types


excel_files = [
    r"E:\Projects\WriteBalance\Build_DataBase_Test\table1.xlsx",
    r"E:\Projects\WriteBalance\Build_DataBase_Test\table2.xlsx",
    r"E:\Projects\WriteBalance\Build_DataBase_Test\table4.xlsx",
]

sheet_name = "Dev"   


engine = create_engine(
    "mssql+pyodbc://localhost/Database1"
    "?driver=ODBC+Driver+17+for+SQL+Server&charset=utf8"
)


for excel_path in excel_files:
    print(f"Importing: {excel_path}")

    df = pd.read_excel(
        excel_path,
        sheet_name=sheet_name,
        engine="openpyxl",
        dtype={
            "Kol_Code": str,
            "Moeen_Code": str,
            "FinApplication_Title": str,
            "AccountNature_Title": str,
            "Flow_Credit":float,
            "Flow_Debit":float,
            "Remain_Last_Credit":float,
            "Remain_last_Debit":float,
        }
    )

    df = df.fillna(0)

    df = df.map(lambda x: str(x).strip() if isinstance(x, str) else x)

    df.to_sql(
        "DWProxyDB",       
        con=engine,
        if_exists="append",    
        index=False,
        dtype={
            "Kol_Title": types.NVARCHAR(length=255),
            "Moeen_Title": types.NVARCHAR(length=255),
            "Kol_Code": types.NVARCHAR(length=255),
            "Moeen_Code": types.NVARCHAR(length=255),
            "Tafzil_Code": types.INTEGER,
            "Tafzil_Tilte": types.INTEGER,
            "AccountNature_ID":types.INTEGER,
            "FinApplication_Title": types.NVARCHAR(length=255),
            "AccountNature_Title": types.NVARCHAR(length=255),
            "Motamam": types.INTEGER,
            "Remain_First_Credit":types.DECIMAL(38,0),
            "Remain_First_Debit":types.DECIMAL(38,0),
            "Flow_Credit":types.DECIMAL(38,0),
            "Flow_Debit":types.DECIMAL(38,0),
            "Remain_Last_Credit":types.DECIMAL(38,0),
            "Remain_last_Debit":types.DECIMAL(38,0),
        },
    )

print("Done!")




# import pandas as pd
# from sqlalchemy import create_engine, types

# excel_path = r"E:\Projects\WriteBalance\table4.xlsx"
# df = pd.read_excel(
#     excel_path,
#     sheet_name="Dev",
#     engine="openpyxl",
#     dtype={
#         "Kol_Code": str,
#         "Moeen_Code": str,
#         "Mande_Bed":"Int64",
#         "Mande_Bes":"Int64" ,
#     }
# )


# df = df.fillna(0)

# df = df.map(lambda x: str(x).strip() if isinstance(x, str) else x)


# engine = create_engine(
#     "mssql+pyodbc://localhost/Database1?driver=ODBC+Driver+17+for+SQL+Server&charset=utf8"
# )


# df.to_sql(
#     "MyExcelImport4",
#     con=engine,
#     if_exists="append",
#     index=False,
#     dtype={
#         "Kol_Title": types.NVARCHAR(length=255),
#         "Moeen_Title": types.NVARCHAR(length=255),
#         "Kol_Code": types.NVARCHAR(length=255),
#         "Moeen_Code": types.NVARCHAR(length=255),
#     },
# )
