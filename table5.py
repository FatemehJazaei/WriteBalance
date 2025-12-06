
import pandas as pd
from sqlalchemy import create_engine, types


excel_files = [
    r"E:\Projects\WriteBalance\table5.xlsx",
]

sheet_name = "Dev"   


engine = create_engine(
    "mssql+pyodbc://localhost/Refah"
    "?driver=ODBC+Driver+17+for+SQL+Server&charset=utf8"
)


for excel_path in excel_files:
    print(f"Importing: {excel_path}")

    df = pd.read_excel(
        excel_path,
        sheet_name=sheet_name,
        engine="openpyxl",
        dtype={
            "cntrlbmi": str,
            "cntrlfdsc":str,
            "memcod":str,
            "detail":str,
            "abbr":str,
            "abbrfdsc":str,
            "drbal":float,
            "crbal":float,
            "drbaleq":float,
            "crbaleq":float,
            "drbsequv":float,
            "crbsequv":float,
            "dramnt":float,
            "cramnt":float,
        }
    )


    df = df.fillna(0)

    df = df.map(lambda x: str(x).strip() if isinstance(x, str) else x)

    df.to_sql(
        "Arzi",       
        con=engine,
        if_exists="append",    
        index=False,
        dtype={
            "cntrlbmi":  types.NVARCHAR(length=255),
            "cntrlfdsc": types.NVARCHAR(length=255),
            "memcod": types.NVARCHAR(length=255),
            "detail": types.NVARCHAR(length=255),
            "abbr": types.NVARCHAR(length=255),
            "abbrfdsc": types.NVARCHAR(length=255),
            "trz_dt": types.INTEGER,
            "brn_cod": types.INTEGER,
            "cntrl": types.INTEGER,
            "curr": types.INTEGER,
            "Cust_no": types.INTEGER,
            "Cust_kd": types.INTEGER,
            "drbal":types.DECIMAL(38,0),
            "crbal":types.DECIMAL(38,0),
            "drbaleq":types.DECIMAL(38,0),
            "crbaleq":types.DECIMAL(38,0),
            "drbsequv":types.DECIMAL(38,0),
            "crbsequv":types.DECIMAL(38,0),
            "dramnt":types.DECIMAL(38,0),
            "cramnt":types.DECIMAL(38,0),
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
