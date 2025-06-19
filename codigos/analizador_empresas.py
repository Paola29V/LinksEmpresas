import pandas as pd
import schedule
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from datetime import datetime
import time
import os

class AnalizadorEmpresas:
    def __init__(self):
        # Configurar el driver con opciones adicionales para cookies y pop-ups
        self.chrome_options = Options()
        self.chrome_options.add_argument('--no-sandbox')
        self.chrome_options.add_argument('--disable-dev-shm-usage')
        self.chrome_options.add_argument('--disable-notifications')
        self.chrome_options.add_argument('--disable-popup-blocking')
        self.chrome_options.add_argument('--incognito')
        self.chrome_options.add_argument("--headless")
        
        
        self.driver = webdriver.Chrome(
            options=self.chrome_options
        )
        
        self.carpeta_resultados = 'resultados_analisisCol'
        if not os.path.exists(self.carpeta_resultados):
            os.makedirs(self.carpeta_resultados)

    def manejar_cookies_y_popups(self):
        try:
            selectores_cookies = [
                "button[id*='accept']",
                "button[id*='cookie']",
                "button[class*='cookie']",
                "a[id*='accept']",
                "a[class*='accept']",
                "#acceptCookies",
                ".accept-cookies",
                "button:contains('Aceptar')",
                "button:contains('Accept')",
                ".modal-close",
                ".popup-close",
                "#close-modal",
                ".close-button",
                "[aria-label='Close']",
                "#onetrust-accept-btn-handler",
                ".cc-dismiss",
                ".cc-allow"
            ]

            for selector in selectores_cookies:
                try:
                    elemento = WebDriverWait(self.driver, 2).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    try:
                        elemento.click()
                        print(f"Cerrado elemento con selector: {selector}")
                        time.sleep(0.5)
                    except ElementClickInterceptedException:
                        self.driver.execute_script("arguments[0].click();", elemento)
                except TimeoutException:
                    continue

            # Manejar iframes
            try:
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                for iframe in iframes:
                    try:
                        if "cookie" in iframe.get_attribute("src").lower():
                            self.driver.switch_to.frame(iframe)
                            for selector in selectores_cookies:
                                try:
                                    elemento = WebDriverWait(self.driver, 1).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                                    )
                                    elemento.click()
                                    break
                                except:
                                    continue
                            self.driver.switch_to.default_content()
                    except:
                        continue
            except:
                pass

        except Exception as e:
            print(f"Error manejando cookies/pop-ups: {str(e)}")

    def eliminar_cookies(self):
        try:
            self.driver.delete_all_cookies()
            print("Cookies eliminadas")
        except Exception as e:
            print(f"Error eliminando cookies: {str(e)}")

    def manejar_captcha(self):
        try:
            iframe = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']"))
            )
            self.driver.switch_to.frame(iframe)
            checkbox = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.ID, "recaptcha-anchor"))
            )
            checkbox.click()
            time.sleep(3)
            self.driver.switch_to.default_content()
            return True
        except:
            return False

    def analizar_pagina(self, url, nombre_empresa, palabras_clave):
        try:
            print(f"Analizando {nombre_empresa}: {url}")
            self.driver.get(url)
            
            # Manejar cookies y pop-ups
            self.manejar_cookies_y_popups()
            
            # Manejar CAPTCHA si existe
            self.manejar_captcha()
            
            time.sleep(3)
            
            contenido = self.driver.page_source.lower()
            
            resultados = {
                'Fecha_Análisis': datetime.now().strftime('%Y-%m-%d'),
                'Hora_Análisis': datetime.now().strftime('%H:%M:%S'),
                'Empresa': nombre_empresa,
                'URL': url,
                'Accesible': True
            }
            
            for palabra in palabras_clave:
                encontrada = palabra.lower() in contenido
                conteo = contenido.count(palabra.lower())
                resultados[f'Contiene_{palabra}'] = 'Sí' if encontrada else 'No'
                resultados[f'Cantidad_{palabra}'] = conteo
            
            # Limpiar cookies
            self.eliminar_cookies()
            
            return resultados
            
        except Exception as e:
            return {
                'Fecha_Análisis': datetime.now().strftime('%Y-%m-%d'),
                'Hora_Análisis': datetime.now().strftime('%H:%M:%S'),
                'Empresa': nombre_empresa,
                'URL': url,
                'Accesible': False,
                'Error': str(e)
            }

    def procesar_empresas(self, archivo_excel, palabras_clave, archivo_salida=None, nombre_hoja='Análisis_Páginas'):
        """
        Procesa las empresas y guarda los resultados en un archivo Excel.
        """
        try:
            # Leer archivo de empresas
            df = pd.read_excel(archivo_excel)
            resultados = []
            
            # Analizar cada empresa
            for index, row in df.iterrows():
                resultado = self.analizar_pagina(
                    row['URL'],
                    row['Empresa'],
                    palabras_clave
                )
                resultados.append(resultado)
                time.sleep(2)
            
            # Crear DataFrame con resultados
            df_resultados = pd.DataFrame(resultados)
            
            # Determinar nombre del archivo de salida
            if archivo_salida is None:
                fecha_hora_actual = datetime.now().strftime('%Y%m%d_%H%M%S')
                archivo_salida = os.path.join(
                    self.carpeta_resultados, 
                    f'resultados_combinados_{fecha_hora_actual}.xlsx'
                )
            
            # Guardar resultados
            self._guardar_resultados(df_resultados, archivo_salida, nombre_hoja)
            print(f"✅ Resultados guardados en: {archivo_salida}")
            
            return archivo_salida
                
        except Exception as e:
            print(f"❌ Error en procesar_empresas: {str(e)}")
            return None
        finally:
            self.driver.quit()

    def _guardar_resultados(self, df_resultados, archivo_salida, nombre_hoja):
        """
    Guarda los resultados en Excel.
    """
        try:
            if os.path.exists(archivo_salida):
            # Si el archivo existe, sobrescribir la hoja existente
               with pd.ExcelWriter(archivo_salida, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_resultados.to_excel(writer, sheet_name=nombre_hoja, index=False)
            else:
            # Si el archivo no existe, crearlo
              with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
                df_resultados.to_excel(writer, sheet_name=nombre_hoja, index=False)
            print(f"✅ Resultados guardados en: {archivo_salida}")
        except Exception as e:
           print(f"❌ Error al guardar resultados: {str(e)}")
        raise