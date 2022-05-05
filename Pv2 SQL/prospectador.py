from empresa import Empresa
from contacto import Contacto
import pandas as pd
import sqlalchemy
import xlsxwriter
import os
import traceback
import math


class Prospectador(Empresa, Contacto):
    hunterkey = '124c421abfa9dfd4a6f5a6fa6b6803786886de40'
    chrome_driver_path = r"C:\Program Files (x86)\chromedriver.exe"
    id_agent = "4772522250592905"
    container = "3317408268735804"
    url_sheets = "https://docs.google.com/spreadsheets/d/1GCtuwfO8mRK371PpR4FQghxvKShppO6ve8IgukIHEhA/edit?usp=sharing"
    session_cookie = "AQEDATIP1boCWrzBAAABeWfNhNYAAAF6CrXXJE0AOKgQ69ah3a3kP0y7063vFYSiurMpju-k5jAZS4gL9pBL_XA8rhp4Ts" \
                     "HnQIRJsY9m_McBLfVbQwDqSMJQT0MTWqGm2DKCU0CZEZ1ujz5e4uCUWw9h"
    results_per_search = 100
    pb_key = "SDX6wzVNo5jgBPwOjSSmLNRy09i1GJOCR7NN7qrUFv0"
    database = "mysql+pymysql://mresta:mresta@integraciones.debmedia.com:3306/mresta_prueba"
    engine = sqlalchemy.create_engine(database)
    hubspot_mails = pd.read_sql_table("hubspot_contacts", engine)["email"].tolist()

    def __init__(self, prospectador, g_drive_path, activity, country, industry, pag, motivo_de_prospeccion,
                 fecha_de_solicitud_de_prospeccion,
                 region_list=None,
                 job_titles=None):

        self.prospectador = prospectador
        self.g_drive_path = g_drive_path
        self.activity = activity  # a futuro estaria reprospectar
        self.country = country
        self.industry = industry
        self.pag = str(pag)
        self.motivo_de_prospeccion = motivo_de_prospeccion
        self.fecha_de_solicitud_de_prospeccion = fecha_de_solicitud_de_prospeccion


        self.carpeta_contenedora = self.g_drive_path + r"\prospecciones\{0}\{1}".format(self.activity, self.country)
        self.carpeta = self.g_drive_path + r"\prospecciones\{0}\{1}\{2}".format(self.activity, self.country,
                                                                                self.industry)
        self.p1c = self.carpeta + r"\p1.xlsx"
        self.p1 = self.carpeta + r"\p1 {}.xlsx".format(self.pag)
        self.p2 = self.carpeta + r"\p2 {}.xlsx".format(self.pag)
        self.p3 = self.carpeta + r"\p3 {}.xlsx".format(self.pag)
        self.p4 = self.carpeta + r"\p4 {}.xlsx".format(self.pag)

        self.region_list = region_list
        self.job_titles = job_titles
        self.file_pb = prospectador + industry  # + str(pag)

        self.region_list_definer()
        # self.path_contactos = self.g_drive_path + r"\hubspot integracion\contactos.csv"
        # self.hubspot_mails = pd.read_csv(self.path_contactos)['Email'].tolist()

    def region_list_definer(self):
        if self.region_list is None:
            self.region_list = [self.country]
        elif self.region_list == "n":
            self.region_list = None
        # si le di una lista de busqueda seria otro elif y dsp pass

    # si no exite la carpeta la creo
    def create_folder(self):
        for carp in [self.carpeta_contenedora, self.carpeta]:
            if not os.path.exists(carp):
                os.makedirs(carp)

    def pros_1(self):
        df = pd.read_excel(self.p1)
        workbook = xlsxwriter.Workbook(self.p2)
        Empresa.chrome_starter(self.chrome_driver_path)
        for z in range(len(df.index)):
            empresa = Empresa(df["enumerado"][z], df["nombres"][z], df["links"][z], df["sucursales"][z])
            empresa.from_link_to_domain()
            empresa.from_domain_to_hunter_info(self.hunterkey)

            if z == 0:  # inicio busqueda pimer dominio
                empresa.g_first_search()

            elif z != len(df.index) - 1 and z != 0:  # inicio busqueda dominios segundo hasta el ultimo
                try:
                    empresa.g_next_search()
                except:
                    hold_on = input("manda algo una vez pasado el captcha")
                    empresa.g_next_search()

            pages_str = empresa.g_pages_to_pages_str()
            empresa.gmails_finder(pages_str)
            workbook = empresa.p1_display(workbook)

        Empresa.driver.close()
        workbook.close()

    def pros_1_no_google(self):
        df = pd.read_excel(self.p1)
        workbook = xlsxwriter.Workbook(self.p2)
        for z in range(len(df.index)):
            empresa = Empresa(df["enumerado"][z], df["nombres"][z], df["links"][z], df["sucursales"][z])
            empresa.from_link_to_domain()
            empresa.from_domain_to_hunter_info(self.hunterkey)

            # if z == 0:  # inicio busqueda pimer dominio
            #     empresa.g_first_search()
            #
            # elif z != len(df.index) - 1 and z != 0:  # inicio busqueda dominios segundo hasta el ultimo
            #     try:
            #         empresa.g_next_search()
            #     except:
            #         hold_on = input("manda algo una vez pasado el captcha")
            #         empresa.g_next_search()
            #
            # pages_str = empresa.g_pages_to_pages_str()
            # empresa.gmails_finder(pages_str)
            workbook = empresa.p1_display(workbook)
        workbook.close()

    def pros_2(self):
        xl = pd.ExcelFile(self.p2)
        workbook = xlsxwriter.Workbook(self.p3)

        cantidad = len(xl.sheet_names)
        plantilla = workbook.add_worksheet("plantilla")

        plantilla.write("B1", "Links")
        plantilla.write("C1", "Company Name")
        plantilla.write("D1", "Ps1")
        plantilla.write("E1", "Ps2")
        plantilla.write("F1", "Ps3")
        plantilla.write("G1", "Ps4")
        plantilla.write("H1", "Dominio sug")
        plantilla.write("I1", "Email")
        plantilla.write("J1", "First Name")
        plantilla.write("K1", "Last Name")
        plantilla.write("L1", "Job Title")
        plantilla.write("M1", "Phone Number")
        plantilla.write("N1", "Linkedin Bio")
        plantilla.write("O1", "Country")
        plantilla.write("P1", "sucursales")
        plantilla.write("Q1", "Industry")
        plantilla.write("R1", "Contact Owner")
        plantilla.write("S1", "Procedencia de Lead")
        plantilla.write("T1", "Lead Status")
        plantilla.write("U1", "Prospectador")
        plantilla.write("V1", "Cs1")
        plantilla.write("W1", "Cs2")
        plantilla.write("X1", "Cs3")
        plantilla.write("Y1", "Cs4")

        plantilla.set_column('A:A', 0)
        plantilla.set_column('B:B', 0)
        plantilla.set_column('C:C', 25)
        plantilla.set_column('D:D', 3)
        plantilla.set_column('E:H', 0)
        plantilla.set_column('I:I', 5)
        plantilla.set_column('L:L', 25)
        plantilla.set_column('M:M', 3)
        plantilla.set_column('P:P', 5)
        plantilla.set_column('Q:U', 5)
        plantilla.set_column('V:Y', 3)

        cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
        plantilla.set_row(0, 15, cell_format)

        inter_primer_lista = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]

        for i in range(cantidad):
            intervalo_lista = []
            for a in inter_primer_lista:
                intervalo_lista.append(a + (10 * i))

            df = pd.read_excel(self.p2, sheet_name=xl.sheet_names[i])

            for nfila in intervalo_lista:
                plantilla.write("B" + str(nfila), str(df["dominio link"][0]))
                plantilla.write("C" + str(nfila), str(df["nombre"][0]))

                plantilla.write("D" + str(nfila), str(df["formato sugerido"][0]))
                plantilla.write("E" + str(nfila), str(df["formato sugerido"][1]))
                plantilla.write("F" + str(nfila), str(df["formato sugerido"][2]))
                plantilla.write("G" + str(nfila), str(df["formato sugerido"][3]))
                plantilla.write("H" + str(nfila), str("@" + df["dominio sugerido"][0]))

                plantilla.write("N" + str(nfila), "=geturl(J" + str(nfila) + ")")
                plantilla.write("O" + str(nfila), self.country)
                plantilla.write("P" + str(nfila), int(df["sucursales"][0]))
                plantilla.write("Q" + str(nfila), self.industry)
                plantilla.write("R" + str(nfila), "Matías Restahinoch")
                plantilla.write("S" + str(nfila), "Outbound")
                plantilla.write("T" + str(nfila), "Prospectado")
                plantilla.write("U" + str(nfila), self.prospectador)

                plantilla.write("V" + str(nfila), 0)
                plantilla.write("W" + str(nfila), 0)
                plantilla.write("X" + str(nfila), 0)
                plantilla.write("Y" + str(nfila), 0)

        workbook.close()

    def pros_2_pb(self):
        self.from_fdm_to_data(self.p2, self.job_titles, self.region_list)
        self.run_and_download_phantom(self.id_agent, self.container, self.session_cookie, self.url_sheets,
                                      self.results_per_search,
                                      self.file_pb, self.pb_key)
        self.from_dict_and_fdm_to_fdp(self.file_pb, self.p2, self.p3, self.country, self.industry, self.prospectador)

    def pros_2_pb_no_country(self):
        self.from_fdm_to_data_no_country(self.p2, self.job_titles)
        self.run_and_download_phantom(self.id_agent, self.container, self.session_cookie, self.url_sheets,
                                      self.results_per_search,
                                      self.file_pb, self.pb_key)
        self.from_dict_and_fdm_to_fdp(self.file_pb, self.p2, self.p3, self.country, self.industry, self.prospectador)

    def pros_3(self):
        df = pd.read_excel(self.p3)

        df['Email'] = df['Email'].apply(lambda x: x if pd.isnull(x) else str(x))
        df['Email'] = df['Email'].astype(str)

        print(range(len(df.index)))
        for i in list(df.index):
            try:
                best_mail_list = self.best_mail(self.mails_posibles(df, i), self.hunterkey, self.hubspot_mails)
                if best_mail_list:
                    print(best_mail_list)
                    df.at[i, "Email"] = best_mail_list[0]
                    df.at[i, "Cs1"] = best_mail_list[1]
            except:
                traceback.print_exc()

        writer = pd.ExcelWriter(self.p4, engine='xlsxwriter')  # ARCHIVO QUE ESCRIBO
        df.to_excel(writer, sheet_name='hoja1')
        writer.save()

    def pros_3_pb(self):
        xl = pd.ExcelFile(self.p3)
        frames = [xl.parse(i) for i in xl.sheet_names]

        df = pd.concat(frames)
        df = df[df['Criterio'] != 0]
        df.reset_index(inplace=True)

        # Estas dos lineas esta pq sino no funciona para que lo lea como str, ambas son necesarias
        df['Email'] = df['Email'].apply(lambda x: x if pd.isnull(x) else str(x))
        df['Email'] = df['Email'].astype(str)

        print(range(len(df.index)))
        for i in list(df.index):
            try:
                best_mail_list = self.best_mail(self.mails_posibles(df, i), self.hunterkey, self.hubspot_mails)
                if best_mail_list:
                    print(best_mail_list)
                    df.at[i, "Email"] = best_mail_list[0]
                    df.at[i, "Cs1"] = best_mail_list[1]
            except:
                traceback.print_exc()

        writer = pd.ExcelWriter(self.p4, engine='xlsxwriter')  # ARCHIVO QUE ESCRIBO
        df.to_excel(writer, sheet_name='hoja1')
        writer.save()

    def particionador_p1(self):

        df = pd.read_excel(self.p1c)
        len_p1s = 15
        c_p = len(df.index) / len_p1s
        c_paginas = math.ceil(c_p)

        for i in range(c_paginas):
            y = i * len_p1s
            z = i * len_p1s + len_p1s

            if i == (c_paginas - 1):
                data = df[y:]
            else:
                data = df[y:z]

            wrpath = self.p1c.replace(".xlsx", " " + str(i + 1) + ".xlsx")
            writer = pd.ExcelWriter(wrpath, engine='xlsxwriter')
            data.to_excel(writer, sheet_name='hoja1', index=False)
            writer.save()

    def unidor_p4(self):
        paths = os.listdir(self.carpeta)
        e = []  # paths de interes
        for i in paths:
            if "p4 " in i:
                e.append(os.path.join(self.carpeta, i))

        lista_dfs = []
        for p4 in e:
            print(p4)
            # noinspection PyTypeChecker
            lista_dfs.append(pd.read_excel(p4, sheet_name="Hoja2", index_col=None, header=None))
            # Hoja2 pq queda ahi dsp de correr el macro de xlsx

        df_c = pd.concat(lista_dfs)
        df_c["Motivo de Prospeccion"] = self.motivo_de_prospeccion
        df_c["Fecha de solicitud de prospeccion"] = self.fecha_de_solicitud_de_prospeccion
        print(df_c)

        writer = pd.ExcelWriter(self.carpeta + r"\p4c.xlsx", engine='xlsxwriter')
        df_c.to_excel(writer, sheet_name='hoja1', index=False, header=False)
        writer.save()

    def carga_sql(self):
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
        df = pd.read_excel(self.carpeta + r"\p4c.xlsx", sheet_name="hoja1", index_col=None, header=None)
        df.columns = columns

        df.to_sql(
            name="pu_contacts",
            con=self.engine,
            index=False,
            if_exists="append",
            method='multi')
        print("done")  #falta testear
