import pandas as pd
import json
import gspread
from oauth2client import service_account
import requests
import xlsxwriter
from pyhunter import PyHunter


class Contacto:

    def __init__(self, company_name, first, last, emails_potenciales=None, linkedin_bio=None,
                 hunter_values=None):
        self.company_name = company_name
        self.first = first
        self.last = last
        self.fullname = first + " " + last
        self.linkedin_bio = linkedin_bio
        self.emails_potenciales = emails_potenciales
        self.hunter_values = hunter_values
        self.pot_emails_values = dict(zip(emails_potenciales, hunter_values))

    @staticmethod
    def from_data_to_sn_url(nombres_empresa, cargos, paises=None):
        pre_url_str = "https://www.linkedin.com/sales/search/people?"
        parameters_list = []

        # companyIncluded
        if nombres_empresa is not None:
            empresas_parse = '%2C'.join([elem for elem in nombres_empresa]).replace(" ", "%2520")
            empresas_str = "companyIncluded=" + empresas_parse + "&companyTimeScope=CURRENT"
            parameters_list.append(empresas_str)

        # titleIncluded
        if cargos is not None:
            cargos_parse = '%2C'.join([elem for elem in cargos]).replace(" ", "%2520")
            cargos_str = "titleIncluded=" + cargos_parse + "&titleTimeScope=CURRENT"
            parameters_list.append(cargos_str)

        # geoIncluded
        paises_code = []
        paises_dict = {"Argentina": "100446943", "México": "103323778", "Chile": "104621616", "Uruguay": "100867946",
                       "Paraguay": "104065273", "República Dominicana": "105057336", "Puerto Rico": "105245958",
                       "Perú": "102927786", "Panamá": "100808673", "Honduras": "101937718", "Guatemala": "100877388",
                       "España": "105646813", "El Salvador": "106522560", "Costa Rica": "101739942",
                       "Colombia": "100876405", "Bolivia": "104379274"}
        if paises is not None:
            for i in paises:
                paises_code.append(paises_dict[i])
            paises_parse = '%2C'.join([elem for elem in paises_code])
            paises_str = "geoIncluded=" + paises_parse
            parameters_list.append(paises_str)

        post_url_str = "&".join(parameters_list)
        return pre_url_str + post_url_str

    @staticmethod
    def update_gsheet(lista_urls):
        SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        CREDENTIALS = service_account.ServiceAccountCredentials.from_json_keyfile_name(
            "credentials.json", SCOPE)
        c = gspread.authorize(credentials=CREDENTIALS)
        spread_sheet = c.open("SN URLs")
        active_sheet = spread_sheet.worksheet("URLs")
        spread_sheet.values_clear("URLs!A:A")
        for i in range(len(lista_urls)):
            active_sheet.update('A' + str(i + 1), lista_urls[i])

    def from_fdm_to_data(self, fdm_file, cargos, paises):
        xl = pd.ExcelFile(fdm_file)
        url_list = []
        for i in xl.sheet_names:
            dic = json.loads(xl.parse(i).to_json())
            dic_company_names = []
            for nombre in dic["nombre"]:
                if dic["nombre"][nombre] is not None:
                    dic_company_names.append(dic["nombre"][nombre])
            url_list.append(self.from_data_to_sn_url(dic_company_names, cargos, paises))
        self.update_gsheet(url_list)

    def from_fdm_to_data_no_country(self, fdm_file, cargos):
        xl = pd.ExcelFile(fdm_file)
        url_list = []
        for i in xl.sheet_names:
            dic = json.loads(xl.parse(i).to_json())
            dic_company_names = []
            for nombre in dic["nombre"]:
                if dic["nombre"][nombre] is not None:
                    dic_company_names.append(dic["nombre"][nombre])
            url_list.append(self.from_data_to_sn_url(nombres_empresa=dic_company_names, cargos=cargos))
        self.update_gsheet(url_list)

    def run_and_download_phantom(self, id_agent, container, session_cookie, url_sheets, results_per_search, file_name, pb_key):
        headers = {
            'Content-Type': 'application/json',
            'x-phantombuster-key': pb_key,
        }
        data = '{"id":"' + id_agent + \
               '","argument":{"extractDefaultUrl":false,"removeDuplicateProfiles":false,"sessionCookie":"' \
               + session_cookie + '","searches":"' + url_sheets + '","numberOfResultsPerSearch":' \
               + str(results_per_search) + ',"csvName":"' + file_name + '"}}'
        print(requests.post('https://api.phantombuster.com/api/v2/agents/launch', headers=headers, data=data))
        print("https://phantombuster.com/" + container + "/phantoms/" + id_agent + "/console")
        hold_on = input("hold on...")
        with requests.get(self.download_json_phantom(id_agent, file_name, pb_key)) as rq:
            with open(file_name + ".json", 'wb') as file:
                file.write(rq.content)

    @staticmethod
    def download_json_phantom(id_agent, file_name, pb_key):
        url = "https://api.phantombuster.com/api/v2/agents/fetch"
        headers = {"Accept": "application/json",
                   'x-phantombuster-key': pb_key}
        params = {"id": id_agent}
        response = requests.request("GET", url, headers=headers, params=params)
        resp = response.json()
        s3Folder = resp["s3Folder"]
        orgS3Folder = resp["orgS3Folder"]
        return "https://phantombuster.s3.amazonaws.com/{1}/{0}/{2}.json".format(s3Folder, orgS3Folder, file_name)

    @staticmethod
    def from_json_to_fdp_dict(data, fdm):
        queries = []
        dic_queries = {}
        for diccionario in data:
            query = diccionario['query']
            if query not in queries:
                # dic_queries[query] = [diccionario['profileUrl']]
                dic_queries[query] = [diccionario]
                queries.append(query)
            else:
                dic_queries[query].append(diccionario)
        x = pd.ExcelFile(fdm)
        dict_new = {}
        for a in range(len(x.sheet_names)):
            dict_new[x.sheet_names[a]] = list(dic_queries.values())[a]
        return dict_new

    def from_dict_and_fdm_to_fdp(self, file_name, fdm, fdp, pais, categoria, prospectador):
        with open(file_name + ".json", encoding="utf8") as json_file:
            print(file_name)
            data = json.load(json_file)

        dict_for_fdp = self.from_json_to_fdp_dict(data, fdm=fdm)
        workbook = xlsxwriter.Workbook(fdp)
        cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})

        for sheet in dict_for_fdp:
            df = pd.read_excel(fdm, sheet_name=sheet)

            if str(df["formato sugerido"][0]) == "n" or "error" in dict_for_fdp[sheet][0]:
                continue

            pl = workbook.add_worksheet(sheet)
            pl.write("B1", "Links")
            pl.write("C1", "Company Name")
            pl.write("D1", "Ps1")
            pl.write("E1", "Ps2")
            pl.write("F1", "Ps3")
            pl.write("G1", "Ps4")
            pl.write("H1", "Dominio sug")
            pl.write("I1", "Email")
            pl.write("J1", "First Name")  # Dict
            pl.write("K1", "Last Name")  # Dict
            pl.write("L1", "Job Title")  # Dict
            pl.write("M1", "Phone Number")
            pl.write("N1", "Linkedin Bio")  # Dict
            pl.write("O1", "Country")
            pl.write("P1", "sucursales")
            pl.write("Q1", "Industry")
            pl.write("R1", "Contact Owner")
            pl.write("S1", "Procedencia de Lead")
            pl.write("T1", "Lead Status")
            pl.write("U1", "Prospectador")
            pl.write("V1", "Cs1")
            pl.write("W1", "Cs2")
            pl.write("X1", "Cs3")
            pl.write("Y1", "Cs4")
            pl.write("Z1", "Company")
            pl.write("AA1", "Linkedin Comp")
            pl.write("AB1", "Criterio")

            pl.set_column('A:A', 0)
            pl.set_column('B:B', 0)
            pl.set_column('C:C', 0)
            pl.set_column('D:D', 0)  # ps1
            pl.set_column('E:I', 0)
            # pl.set_column('I:I', 5)
            pl.set_column('L:L', 75)
            pl.set_column('M:N', 0)
            pl.set_column('O:Y', 0)
            pl.set_column('Z:AA', 20)

            pl.set_row(0, 15, cell_format)

            for contacto in dict_for_fdp[sheet]:
                nfila = 2 + dict_for_fdp[sheet].index(contacto)
                pl.set_row(row=(nfila - 1), height=18)

                pl.write("B" + str(nfila), str(df["dominio link"][0]))
                pl.write("C" + str(nfila), str(df["nombre"][0]))

                pl.write("D" + str(nfila), str(df["formato sugerido"][0]))
                pl.write("E" + str(nfila), str(df["formato sugerido"][1]))
                pl.write("F" + str(nfila), str(df["formato sugerido"][2]))
                pl.write("G" + str(nfila), str(df["formato sugerido"][3]))
                pl.write("H" + str(nfila), str("@" + df["dominio sugerido"][0]))

                pl.write("J" + str(nfila), contacto["firstName"])  # first name
                try:  # last name try/ex: los nombres chinos no tienen apellido y tmb hay caracteres que complican
                    pl.write("K" + str(nfila), contacto["lastName"])
                    pl.write("L" + str(nfila), contacto["title"])
                except:
                    pass
                pl.write("N" + str(nfila), contacto["profileUrl"])
                # pl.write("N" + str(nfila), "=geturl(J" + str(nfila) + ")")
                pl.write("O" + str(nfila), pais)
                # pl.write("P" + str(nfila), int(df["sucursales"][0]))
                pl.write("Q" + str(nfila), categoria)
                pl.write("R" + str(nfila), "Matías Restahinoch")
                pl.write("S" + str(nfila), "Outbound")
                pl.write("T" + str(nfila), "Prospectado")
                pl.write("U" + str(nfila), prospectador)

                pl.write("V" + str(nfila), 0)
                pl.write("W" + str(nfila), 0)
                pl.write("X" + str(nfila), 0)
                pl.write("Y" + str(nfila), 0)

                pl.write("Z" + str(nfila), str(df["nombre"][0]))
                pl.write("AA" + str(nfila), contacto["companyName"])
                pl.write("AB" + str(nfila), 0)

        workbook.close()

    @staticmethod
    def generador_mails(psug, dsug, fname, lname):
        try:
            for i in range(6):
                if fname[0] == " ":
                    fname = fname[1:]
                if fname[-1] == " ":
                    fname = fname[:-1]
                if lname[0] == " ":
                    lname = lname[1:]
                if lname[-1] == " ":
                    lname = lname[:-1]

            formatoMail = lambda x: x.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó",
                                                                                                            "o").replace(
                "ú", "u").replace("ñ", "n").replace("-", "").replace(",", "").replace(".", "")
            f = formatoMail(fname[0])
            l = formatoMail(lname[0])
            first = formatoMail(fname.split(" ")[0])
            last = formatoMail(lname.split(" ")[0])

            if " " in fname:
                f2 = formatoMail(fname.split(" ")[1][0])
                first2 = formatoMail(fname.split(" ")[1])
            else:
                f2 = ""
                first2 = ""

            if " " in lname:
                l2 = formatoMail(lname.split(" ")[1][0])
                last2 = formatoMail(lname.split(" ")[1])
            else:
                l2 = ""
                last2 = ""

            listaDeFormatos = [f, l, first, last, f2, l2, first2, last2]
            listaDeStrsFormatos = ["{f}", "{l}", "{first}", "{last}", "{f2}", "{l2}", "{first2}", "{last2}"]

            pParteMail = psug
            for i in listaDeStrsFormatos:
                if i in psug:
                    pParteMail = pParteMail.replace(i, listaDeFormatos[listaDeStrsFormatos.index(i)])

            mail = pParteMail + dsug
            print(mail)
            return mail
        except:
            return None

    def mails_posibles(self, frame, row):
        mailsPosibles = []
        for e in range(1, 5):
            posMail = self.generador_mails(frame["Ps" + str(e)][row], frame["Dominio sug"][row],
                                           frame["First Name"][row],
                                           frame["Last Name"][row])
            if posMail:
                mailsPosibles.append(posMail)
        return mailsPosibles

    #  hasta aca se usa el self solo para metodos, en realidad esta tod o  estatico

    @staticmethod
    def best_mail(lista_mails, hunterkey, hubspot_mails):
        hunter = PyHunter(hunterkey)
        if lista_mails:
            mails_score = {}
            for i in lista_mails:
                if i in hubspot_mails:
                    mails_score[i] = 1
                else:
                    mails_score[i] = hunter.email_verifier(i)["score"]
                print(mails_score[i])
            max_mail = max(mails_score)
            max_score = max(list(mails_score.values()))
            return [max_mail, max_score]
        else:
            return None
