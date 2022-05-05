from prospectador import Prospectador

# ---- general ----
prospectador = "Prospectador 2"
g_drive_path = r"C:\Users\nico\Google Drive (nmenzaghi@debmedia.com)"
activity = "prospectar"
country = "Argentina"
industry = "Ejemplo"
pagina = 1
countries_list = [country]
motivo_de_prospeccion = "asdasdasd"
fecha_de_solicitud_de_prospeccion = "2010-11-02" # 	YYYY-MM-DD

# ---- general ----
job_title_list = ["agencias", "ceo", "Chief Executive Director", "cio", "coo", "cto", "Director", "experience",
                  "experiencia", "innovacion", "innovation", "IT", "TI", "marketing", "mercadeo", "mejora continua",
                  "operaciones", "operations", "presidente", "procesos", "proyecto", "sistemas", "sucursales",
                  "tecnología", "proyectos", "customer success"]

# ---- partners ----
# job_title_list = ["ceo", "Chief Executive Director", "cio", "coo", "cto", "Director",
#                   "ventas", "sales", "marketing", "mercadeo",
#                   "operaciones", "operations", "presidente", "proyecto", "proyectos", "comercial"]


p = Prospectador(prospectador, g_drive_path, activity, country, industry, pagina, motivo_de_prospeccion, fecha_de_solicitud_de_prospeccion,  job_titles=job_title_list,
                 region_list=countries_list)

p.session_cookie = "AQEDATIP1boFDWZJAAABes_StFcAAAF73x0D_U0AGqrOXauw1UwLTgw6U6ZPHD0wTxuruGZ1db_54fUG9qIYKX6Ez_21osnX" \
                   "yWmd3g7UfnfOb0GHtfssTrcpeAQSR_TzTalbeYhMlUyzSu_joHC3l6EE"

# p.particionador_p1()
p.pros_1()
# p.pros_2_pb()
# p.pros_3_pb()
# p.unidor_p4()

# p.carga_sql()
# correr solo despues de unidor_p4 con motivo de prospeccion y fecha (expresada dd/mm/yyyy funciona, pero creo que otros formatos tambien)
# No tiene que tener headers tampoco

# ----- poco frecuentes -----
# p.pros_1_no_google()
# p.pros_2()
# p.pros_3()
# p.pros_2_pb_no_country()


