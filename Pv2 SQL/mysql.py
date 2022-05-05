import pandas as pd
from sqlalchemy import create_engine

columns = ["Company name",
           "Email",
           "First Name",
           "Last Name",
           "Job Title",
           "Phone Number",
           "Linkedin Bio",
           "Sucursales",
           "Country",
           "Industry",
           "Contact Owner",
           "Procedencia de Lead",
           "Lead Status",
           "Prospectador",
           "Hunter Value",
           "Motivo de Prospeccion",
           "Fecha de solicitud de prospeccion"]

df = pd.read_excel(
    r"C:\Users\nico\Google Drive (nmenzaghi@debmedia.com)\prospecciones\prospectar\Argentina\Ejemplo\p4 1.xlsx",
    sheet_name="Hoja2", index_col=None, header=None)  # header=None,index=False  usecols=columns

df.columns = columns

print(df)

df.to_sql(
    name="pu_contacts",
    con=create_engine('mysql+pymysql://mresta:mresta@integraciones.debmedia.com:3306/mresta_prueba'),
    index=False,
    if_exists="append",
    method='multi')

print("done")

