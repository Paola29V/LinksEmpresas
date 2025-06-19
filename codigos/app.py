import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
from analizador_empresas import AnalizadorEmpresas
from buscador_contratos import BuscadorContratos

def configurar_pagina():
    st.set_page_config(
        page_title="Analizador de Empresas",
        page_icon="📊",
        layout="wide"
    )

def crear_carpeta_resultados():
    carpeta_resultados = 'resultados_analisisCol'
    if not os.path.exists(carpeta_resultados):
        os.makedirs(carpeta_resultados)
    return carpeta_resultados

def ejecutar_analisis_completo(archivo_excel, palabras_clave, progress_bar, status_text):
    try:
        # Configurar nombre del archivo de salida
        fecha_actual = datetime.now().strftime('%Y%m%d')
        archivo_salida = f'resultados_combinados_{fecha_actual}.xlsx'
        
        # Ejecutar primer análisis (AnalizadorEmpresas)
        status_text.write("🔍 Iniciando análisis de páginas web...")
        progress_bar.progress(25)
        
        analizador = AnalizadorEmpresas()
        analizador.procesar_empresas(
            archivo_excel=archivo_excel,
            palabras_clave=palabras_clave,
            archivo_salida=archivo_salida,
            nombre_hoja='Análisis_Páginas'
        )

        # Ejecutar segundo análisis (BuscadorContratos)
        status_text.write("🔍 Iniciando búsqueda de contratos...")
        progress_bar.progress(50)
        
        buscador = BuscadorContratos()
        for empresa in buscador.empresas.keys():
            buscador.buscar_contratos(empresa, palabras_clave)
            progress_bar.progress(75)
        
        buscador.guardar_resultados(
            archivo_salida=archivo_salida,
            nombre_hoja='Búsqueda_Contratos'
        )

        progress_bar.progress(100)
        status_text.write(f"✅ Análisis completo guardado en: {archivo_salida}")
        
        return archivo_salida

    except Exception as e:
        status_text.error(f"❌ Error durante la ejecución: {str(e)}")
        raise e

def main():
    configurar_pagina()
    carpeta_resultados = crear_carpeta_resultados()

    st.title("Analizador de Empresas y Contratos")

    # Crear dos columnas
    col1, col2 = st.columns([2, 1])

    with col1:
        # Opción para cargar archivo o usar archivo predeterminado
        opcion_archivo = st.radio(
            "Seleccione la fuente de datos:",
            ["Cargar archivo Excel", "Usar archivo predeterminado"]
        )

        archivo_excel = None
        if opcion_archivo == "Cargar archivo Excel":
            uploaded_file = st.file_uploader("Cargar archivo Excel", type=['xlsx'])
            if uploaded_file is not None:
                # Guardar archivo temporal
                temp_file = "temp_empresas.xlsx"
                with open(temp_file, "wb") as f:
                    f.write(uploaded_file.getvalue())
                archivo_excel = temp_file
        else:
            archivo_excel = "C:/LinksEmpresas/codigos/EmpresasCol.xlsx"

        # Palabras clave predeterminadas
        default_palabras = """estudio de flujo de carga
cortocircuito
coordinación de protecciones
análisis de estabilidad
arranque de motores
análisis de armónicos
análisis de fenómenos transitorios
análisis de confiabilidad
sistemas de puesta a tierra
coordinación de aislamiento
2025"""

        # Área de texto para palabras clave
        palabras_clave = st.text_area(
            "Palabras clave (una por línea)",
            value=default_palabras
        ).split('\n')

        # Botón para ejecutar análisis
        if st.button("🚀 Ejecutar Análisis", type="primary"):
            if archivo_excel:
                try:
                    # Contenedores para progreso y estado
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    # Ejecutar análisis
                    archivo_salida = ejecutar_analisis_completo(
                        archivo_excel,
                        palabras_clave,
                        progress_bar,
                        status_text
                    )

                    # Botón de descarga
                    if os.path.exists(archivo_salida):
                        with open(archivo_salida, "rb") as file:
                            st.download_button(
                                label="📥 Descargar Resultados",
                                data=file,
                                file_name=archivo_salida,
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )

                except Exception as e:
                    st.error(f"Error durante el análisis: {str(e)}")
                finally:
                    # Limpiar archivo temporal si existe
                    if opcion_archivo == "Cargar archivo Excel" and os.path.exists("temp_empresas.xlsx"):
                        os.remove("temp_empresas.xlsx")
            else:
                st.warning("Por favor, seleccione o cargue un archivo Excel.")

    with col2:
        # Instrucciones
        with st.expander("📖 Instrucciones de uso", expanded=True):
            st.write("""
            1. Seleccione la fuente de datos:
               - Cargar nuevo archivo Excel
               - Usar archivo predeterminado
            2. Revise y modifique las palabras clave si es necesario
            3. Haga clic en 'Ejecutar Análisis'
            4. Espere a que el proceso termine
            5. Descargue los resultados
            """)

        # Historial de análisis
        st.markdown("### 📜 Historial de Análisis")
        archivos = [f for f in os.listdir(carpeta_resultados) if f.endswith('.xlsx')]
        if archivos:
            for archivo in archivos:
                col1, col2 = st.columns([3,1])
                with col1:
                    st.text(archivo)
                with col2:
                    with open(os.path.join(carpeta_resultados, archivo), "rb") as file:
                        st.download_button(
                            label="📥",
                            data=file,
                            file_name=archivo,
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            key=archivo
                        )
        else:
            st.info("No hay análisis previos")

if __name__ == "__main__":
    main()