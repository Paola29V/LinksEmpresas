from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
from datetime import datetime
import os

class BuscadorContratos:
    def __init__(self):
        self.options = Options()
        self.options.add_argument("--headless")
        self.options.add_argument("--disable-gpu")
        self.options.add_argument("--no-sandbox")
        self.options.add_argument("--log-level=3")
        
        # Lista para almacenar resultados
        self.resultados = []
        
        self.empresas = {
            "ELECTROHUILA": {
                "url": "https://electrohuila.com.co/contratacion/",
                "xpath_busqueda": '//*[@id="tblContratos_filter"]/label/input',
                "xpath_sin_resultados": '//*[@id="tblContratos"]/tbody/tr/td',
                "xpath_resultados": '//*[@id="tblContratos"]/tbody/tr[not(contains(@class, "dataTables_empty"))]',
                "texto_sin_resultados": "No se encontraron registros"
            },
            "EPM": {
                "url": "https://www13.epm.com.co/ComprasMayores/ConsultarProcesosContratacion.aspx?t=a",
                "xpath_busqueda": '//*[@id="ctl00_ContentPlaceHolderContenido_txtObjeto"]',
                "xpath_sin_resultados": '//*[@id="ctl00_ContentPlaceHolderContenido_Panel1"]/div/div/div[5]/div[4]',
                "xpath_resultados": '//*[@id="ctl00_ContentPlaceHolderContenido_Panel1"]//div[string-length(normalize-space(text())) > 0]',
                "texto_sin_resultados": "No se encontraron resultados"
            },
            "EMSA": {
                "url": "https://www.electrificadoradelmeta.com.co/auctions",
                "xpath_busqueda": '//*[@id="DataTables_Table_0_wrapper"]/div[1]/div[4]/input',
                "xpath_sin_resultados": '//*[@id="DataTables_Table_0_wrapper"]/div[2]/table',
                "xpath_resultados": '//*[@id="DataTables_Table_0"]/tbody/tr[not(contains(@class, "DataTables_empty"))]',
                "texto_sin_resultados": "No se encontraron resultados"
            }
        }

    def buscar_contratos(self, empresa, palabras_clave):
        config = self.empresas[empresa]
        driver = webdriver.Chrome(options=self.options)
        wait = WebDriverWait(driver, 10)

        try:
            print(f"\n🔍 Buscando en {empresa}")
            print("=" * 50)
            driver.get(config["url"])

            campo_busqueda = wait.until(EC.presence_of_element_located(
                (By.XPATH, config["xpath_busqueda"])
            ))

            for palabra in palabras_clave:
                print(f"\n📝 Buscando: '{palabra}'")

                campo_busqueda.clear()
                campo_busqueda.send_keys(palabra)
                campo_busqueda.send_keys(Keys.RETURN)
                time.sleep(2)

                encontrado = False
                num_resultados = 0

                try:
                    sin_resultados = driver.find_element(By.XPATH, config["xpath_sin_resultados"])
                    if config["texto_sin_resultados"] in sin_resultados.text:
                        print(f"❌ No se encontraron contratos para '{palabra}'")
                    else:
                        filas = driver.find_elements(By.XPATH, config["xpath_resultados"])
                        if filas:
                            num_resultados = len(filas)
                            encontrado = True
                            print(f"✅ Se encontraron {num_resultados} contratos para '{palabra}'")
                except:
                    filas = driver.find_elements(By.XPATH, config["xpath_resultados"])
                    if filas:
                        num_resultados = len(filas)
                        encontrado = True
                        print(f"✅ Se encontraron {num_resultados} contratos para '{palabra}'")

                # Guardar resultados
                self.resultados.append({
                    'Fecha_Búsqueda': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'Empresa': empresa,
                    'URL': config["url"],
                    'Palabra_Clave': palabra,
                    'Encontrado': "Sí" if encontrado else "No",
                    'Cantidad_Resultados': num_resultados
                })

        except Exception as e:
            print(f"⚠️ Error en {empresa}: {str(e)}")
        finally:
            driver.quit()

    def guardar_resultados(self, archivo_salida=None, nombre_hoja='Búsqueda_Contratos'):
        """
        Guarda los resultados en Excel.
        
        Args:
            archivo_salida (str): Ruta del archivo donde guardar los resultados
            nombre_hoja (str): Nombre de la hoja en el archivo Excel
        """
        try:
            # Crear DataFrame con los resultados
            df = pd.DataFrame(self.resultados)
        
        # Si no se especifica archivo de salida, crear uno nuevo con fecha y hora
            if archivo_salida is None:
                fecha_hora_actual = datetime.now().strftime('%Y%m%d_%H%M%S')
                archivo_salida = f'resultados_combinados_{fecha_hora_actual}.xlsx'
        
        # Guardar resultados
            if os.path.exists(archivo_salida):
                 with pd.ExcelWriter(archivo_salida, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                     df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            else:
                with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
                print(f"✅ Resultados guardados en: {archivo_salida}")

        except Exception as e:
         print(f"❌ Error al guardar resultados: {str(e)}")