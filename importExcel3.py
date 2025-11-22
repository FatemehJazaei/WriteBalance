import pandas as pd
from sqlalchemy import create_engine, types

excel_path = r"E:\Projects\WriteBalance\table3.xlsx"
df = pd.read_excel(
    excel_path,
    sheet_name="Dev",
    engine="openpyxl",
    dtype={
        "حساب کل": str,
        "حساب معین": str,
        "حساب تفصیلی": str,
        "حساب جز1": str,
        "حساب جز2": str,
        "کد مرکز هزینه": str,
        "کد واحد عملیاتی": str,
        "نام واحد عملیاتی": str,
        "کد پرونده": str,
        "نام پرونده": str,
    }
)

df = df.map(lambda x: str(x).strip() if isinstance(x, str) else x)

engine = create_engine(
    "mssql+pyodbc://localhost/Database1?driver=ODBC+Driver+17+for+SQL+Server&charset=utf8"
)

df.to_sql(
    "AccountingDB",
    con=engine,
    if_exists="append",
    index=False,
    dtype={
        "کد گروه": types.NVARCHAR(length=255),
        "نام گروه": types.NVARCHAR(length=255),
        "نام حساب کل": types.NVARCHAR(length=255),
        "نام حساب معین": types.NVARCHAR(length=255),
        "نام حساب تفصیلی": types.NVARCHAR(length=255),
        "نام حساب جز1": types.NVARCHAR(length=255),
        "نام حساب جز2": types.NVARCHAR(length=255),
        "حساب کل": types.NVARCHAR(length=255),
        "حساب معین": types.NVARCHAR(length=255),
        "حساب تفصیلی": types.NVARCHAR(length=255),
        "حساب جز1": types.NVARCHAR(length=255),
        "حساب جز2": types.NVARCHAR(length=255),
        "کد مرکز هزینه":  types.NVARCHAR(length=255),
        "کد واحد عملیاتی":  types.NVARCHAR(length=255),
        "نام واحد عملیاتی":  types.NVARCHAR(length=255),
        "کد پرونده":  types.NVARCHAR(length=255),
        "نام پرونده":  types.NVARCHAR(length=255),
        "مانده اول دوره": types.Float,
        "بدهکار": types.Float,
        "بستانکار": types.Float,
        "مانده بدهکار": types.Float,
        "مانده بستانکار": types.Float,
    },
)


print("Done!")