import platform
import time
import wget
import zipfile
import requests
import os
import datetime
import pdfkit

from tkinter import messagebox
from tkinter import *
from selenium import webdriver
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

APP_PATH = os.path.dirname(os.path.realpath(__file__))

class AutoEmpresas:
    def __init__(self):
        self.driver = None
        self.download_dir = None

    def run(self):
        self.create_main_view()

    def create_main_view(self):
        root = Tk()
        root.title("AutoEmpresas")

        lb_title = Label(root, width=30, text="Lista de empresas: ")
        lb_title.grid(row=1, column=0, padx=5, pady=5)

        mt_ids = Text(root, height=15)
        mt_ids.grid(row=1, column=1, padx=0, pady=5)

        scrollbar = Scrollbar(root)
        scrollbar.grid(row=1, column=2, sticky="ns")
        mt_ids.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=mt_ids.yview)

        btn_start_workflow = Button(root, text='Obtener capturas',
                                    command=lambda: self.main_workflow(mt_ids.get('1.0', 'end-1c').splitlines()))
        btn_start_workflow.grid(row=2, columnspan=3, padx=5, pady=5)

        btn_open_folder = Button(root, text='Abrir carpeta de capturas',
                                 command=lambda: self.open_capture_folder())
        btn_open_folder.grid(row=3, columnspan=3, padx=5, pady=5)

        root.mainloop()

    def open_capture_folder(self):
        capture_folder = os.path.join(APP_PATH, "capturas")
        os.startfile(capture_folder)

    def main_workflow(self, ids_list):
        self.get_latest_driver()
        self.download_workflow(ids_list)

    def download_workflow(self, nombres):
        self.driver = self.create_driver()
        os_separator = os.sep
        self.download_dir = os.path.join(APP_PATH, "capturas")

        for nombre in nombres:
            while True:
                try:
                    self.process_sancionados_page(nombre)
                    self.process_dof_page(nombre)

                except (WebDriverException, NoSuchElementException) as e:
                    self.handle_error(e)
                    continue
                break

        self.driver.quit()
        messagebox.showinfo("Proceso Terminado", "Capturas descargadas")

    def process_sancionados_page(self, nombre):
        self.driver.get('https://directoriosancionados.apps.funcionpublica.gob.mx/SanFicTec/jsp/Ficha_Tecnica/SancionadosN.htm')
        combo_box_filter = self.driver.find_element(By.XPATH,
                                                    '/html/body/app-root/div/app-body/form/div/div[2]/div[5]/select')
        select = Select(combo_box_filter)
        options = select.options
        last_option = options[-1]
        select.select_by_value(last_option.get_attribute('value'))

        table_locator = (By.XPATH,
                         '/html/body/app-root/div/app-body/form/div/div[3]/div[2]/body-seleccion-prov/form/div[1]/table/tbody/tr[4]/th')

        wait = WebDriverWait(self.driver, 5)
        table_head = wait.until(EC.visibility_of_element_located(table_locator))

        search_box_name = self.driver.find_element(By.XPATH, '/html/body/app-root/div/app-body/form/div/div[3]/div[2]/body-seleccion-prov/form/div[1]/mat-form-field/div/div[1]/div/input')
        search_box_name.send_keys(nombre)

        btn_search = self.driver.find_element(By.XPATH, "/html/body/app-root/div/app-body/form/div/div[3]/div[2]/body-seleccion-prov/form/div[1]/button[1]")
        self.driver.execute_script("arguments[0].click();", btn_search)


        os.makedirs(os.path.join(self.download_dir, nombre), exist_ok=True)
        download_dir_temp = os.path.join(self.download_dir, nombre, 'directoriosancionados.png')
        search_box_name.send_keys(Keys.TAB)
        self.save_screenshot(download_dir_temp)
        print('Captura guardada en:', download_dir_temp)

    def process_dof_page(self, nombre):
        today = datetime.date.today()
        one_year_ago = today - datetime.timedelta(days=365)
        today_date = today.strftime("%d-%m-%Y")
        one_year_ago_date = one_year_ago.strftime("%d-%m-%Y")
        url = "https://sidof.segob.gob.mx/busquedaAvanzada/busqueda?tipo=C&tipotexto=F&texto=CHANGEME&fechainicio=DATE&fechahasta=DATE&organismos=&sinonimos=false"
        new_url = url.replace("CHANGEME", nombre).replace("DATE", one_year_ago_date, 1).replace("DATE", today_date)
        self.driver.get(new_url)

        # Get the current window handle
        current_window = self.driver.current_window_handle

        # Perform Ctrl+P action to open print dialog
        ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('p').key_up(Keys.CONTROL).perform()

        # Switch to the print dialog window
        for window_handle in self.driver.window_handles:
            if window_handle != current_window:
                self.driver.switch_to.window(window_handle)
                break

        # Select "Save as PDF" option
        save_as_pdf_button = self.driver.find_element_by_xpath("//button[@aria-label='Save as PDF']")
        save_as_pdf_button.click()

        # Choose the destination directory and file name
        save_dialog = self.driver.find_element_by_xpath("//input[@type='file']")
        save_dialog.send_keys(os.path.join(self.download_dir, nombre, 'DOF.pdf'))

        # Close the print dialog window
        self.driver.close()

        # Switch back to the original window
        self.driver.switch_to.window(current_window)

        print('PDF saved in:', os.path.join(self.download_dir, nombre, 'DOF.pdf'))

    def save_screenshot(self, path):
        self.driver.maximize_window()
        time.sleep(2)
        self.driver.save_screenshot(path)

    def get_latest_driver(self):
        url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE'
        response = requests.get(url)
        latest_version = response.text

        if os.name == 'nt':
            download_url = f"https://chromedriver.storage.googleapis.com/{latest_version}/chromedriver_win32.zip"
        else:
            download_url = f"https://chromedriver.storage.googleapis.com/{latest_version}/chromedriver_linux64.zip"

        latest_driver_zip = wget.download(download_url, 'chromedriver.zip')

        with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
            zip_ref.extractall()

        os.remove(latest_driver_zip)

    def create_driver(self):
        chromedriver = ""
        if platform.system() == 'Windows':
            chromedriver = "chromedriver.exe"
        elif platform.system() == 'Linux':
            chromedriver = "chromedriver"
        else:
            messagebox.showerror("Unsupported OS", "Operating system not supported.")
            exit(1)

        download_dir = os.path.join(APP_PATH, "capturas")

        chrome_options = Options()
        chrome_options.add_experimental_option('prefs', {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing_for_trusted_sources_enabled": False,
            "safebrowsing.enabled": False,
            'profile.default_content_setting_values.automatic_downloads': 1
        })
        chrome_options.add_argument('--no-sandbox')
        # chrome_options.add_argument('--headless')  # Run Chrome in headless mode (no GUI)
        chrome_options.add_argument('--disable-gpu')  # Disable GPU acceleration
        chrome_options.add_argument('--print-to-pdf')  # Enable printing to PDF

        if platform.system() != 'Windows':
            os.chmod(chromedriver, 0o755)

        service = Service(chromedriver)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver

    def handle_error(self, e):
        time.sleep(4)
        self.driver.quit()
        self.driver = self.create_driver()
        print(e)


if __name__ == '__main__':
    auto_empresas = AutoEmpresas()
    auto_empresas.run()
