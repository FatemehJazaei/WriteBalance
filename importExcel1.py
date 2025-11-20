
import pandas as pd
from sqlalchemy import create_engine, types


excel_files = [
    r"E:\Projects\WriteBalance\table1.xlsx",
    r"E:\Projects\WriteBalance\table2.xlsx",
    r"E:\Projects\WriteBalance\table4.xlsx",
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
            "Tafzil_Code": types.NVARCHAR(length=255),
            "Tafzil_Tilte": types.NVARCHAR(length=255),
            "FinApplication_Title": types.NVARCHAR(length=255),
            "Gardersh_Bed":types.DECIMAL,
            "Gardersh_Bes":types.DECIMAL,
            "Mande_Bed":types.DECIMAL,
            "Mande_Bes":types.DECIMAL,
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
