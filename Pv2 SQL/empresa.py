import random
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from pyhunter import PyHunter


class Empresa:

    sleep_time = random.uniform(0.8, 1.1)
    formatos_p = ["{f}{last}", "{first}.{last}", "{first}{last}", "{f}.{last}" , "{f}{f2}{last}", "{first}{f2}{last}",
                  "{first}{last}{last2}", "{first}.{last}{last2}", "{last}{f}", "{last}.{f}", "{last}{first}",
                  "{last}.{first}", "{first}{l}", "{first}.{l}", "{l}{first}", "{l}.{first}", "{first}", "{last}"]
    paginas_a_iterar = 7
    q_de_empresas = 0

    def __init__(self, enumerado, nombre, link, sucursales, domain=None, gmails=None, hunter_mails=None,
                 hunter_pattern=None):
        self.enumerado = enumerado
        self.nombre = nombre
        self.link = link
        self.sucursales = sucursales
        self.domain = domain
        self.gmails = gmails
        self.hunter_mails = hunter_mails
        self.hunter_pattern = hunter_pattern

        Empresa.q_de_empresas += 1

    def from_link_to_domain(self):
        a = str(self.link).lower()
        a = a.replace("www.", "").replace("www2.", "").replace("https://", "").replace("http://", "")
        try:
            b = a.replace("/", " ").split(" ")
            self.domain = b[0]
        except:
            self.domain = a


    def gmails_finder(self, page_str):
        if self.gmails is None:
            self.gmails = []
        lista = page_str.split()
        for i in lista:
            if "@" in i:
                self.gmails.append(i)

    def from_domain_to_hunter_info(self, hunterkey):
        info_hunter = PyHunter(hunterkey).domain_search(self.domain)
        self.hunter_mails = []
        for n in info_hunter['emails']:
            self.hunter_mails.append(n["value"])
        self.hunter_pattern = info_hunter['pattern']

    def p1_display(self, wb):

        pagina_ind = wb.add_worksheet(str(self.enumerado))

        pagina_ind.write("A1", "nombre")
        pagina_ind.write("B1", "dominio link")
        pagina_ind.write("C1", "hunter mails")
        pagina_ind.write("D1", "patron sug hunter")
        pagina_ind.write("E1", "google mails")
        pagina_ind.write("F1", "formatos p")
        pagina_ind.write("G1", "formato sugerido")
        pagina_ind.write("H1", "dominio sugerido")
        pagina_ind.write("I1", "sucursales")

        pagina_ind.write("A2", self.nombre)
        pagina_ind.write("B2", self.domain)
        pagina_ind.write("D2", self.hunter_pattern)
        pagina_ind.write("I2", self.sucursales)

        if self.hunter_mails is not None:
            for email in self.hunter_mails:
                pagina_ind.write("C" + str(self.hunter_mails.index(email) + 2), email)

        print(self.gmails)
        if self.gmails is not None:
            for gmail in self.gmails:
                pagina_ind.write("E" + str(self.gmails.index(gmail) + 2), gmail)

        for formato in self.formatos_p:
            pagina_ind.write("F" + str(self.formatos_p.index(formato) + 2), formato)

        pagina_ind.set_column('A:A', 20)
        pagina_ind.set_column('B:B', 30)
        pagina_ind.set_column('C:C', 40)
        pagina_ind.set_column('D:D', 17)
        pagina_ind.set_column('E:E', 40)
        pagina_ind.set_column('F:F', 17)
        pagina_ind.set_column('G:H', 25)

        cell_format = wb.add_format({'bold': True, 'font_color': 'red'})
        pagina_ind.set_row(0, 15, cell_format)

        return wb

    @classmethod
    def chrome_starter(cls, chrome_driver_path):
        cls.driver = webdriver.Chrome(executable_path=chrome_driver_path)

    def g_first_search(self):
        time.sleep(self.sleep_time)
        self.driver.get("https://www.google.com")
        inputElem = self.driver.find_element_by_css_selector("input[name=q]")
        inputElem.send_keys("@" + self.domain)
        inputElem.send_keys(Keys.ENTER)

    def g_next_search(self):
        time.sleep(self.sleep_time)
        inputElem = self.driver.find_element_by_css_selector("input[name=q]")
        inputElem.clear()
        inputElem.send_keys("@" + self.domain)  # para esto generar una lista de dominios
        inputElem.send_keys(Keys.ENTER)

    def g_pages_to_pages_str(self):
        descriptions = []
        for e in range(self.paginas_a_iterar):  # pasaje de paginas (observar si se pierde info)
            time.sleep(self.sleep_time)
            elements = self.driver.find_elements_by_class_name("IsZvec")
            for element in elements:
                descriptions.append(element.text)
            try:
                self.driver.find_element_by_id("pnnext").click()
            except Exception:
                break
        return " ".join(descriptions)
