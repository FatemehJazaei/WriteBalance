import pandas as pd
from sqlalchemy import create_engine, types

excel_path = r"E:\Projects\WriteBalance\table4.xlsx"
df = pd.read_excel(
    excel_path,
    sheet_name="Dev",
    engine="openpyxl",
    dtype={
        "Kol_Code": str,
        "Moeen_Code": str,
        "Tafzil_Code": str,  
    }
)


df = df.map(lambda x: str(x).strip() if isinstance(x, str) else x)


engine = create_engine(
    "mssql+pyodbc://localhost/Database1?driver=ODBC+Driver+17+for+SQL+Server&charset=utf8"
)


df.to_sql(
    "MyExcelImport4",
    con=engine,
    if_exists="append",
    index=False,
    dtype={
        "Kol_Title": types.NVARCHAR(length=255),
        "Moeen_Title": types.NVARCHAR(length=255),
        "Kol_Code": types.NVARCHAR(length=255),
        "Moeen_Code": types.NVARCHAR(length=255),
        "Tafzil_Tilte": types.NVARCHAR(length=255),
        "FinApplication_Title": types.NVARCHAR(length=255),
    },
)
