import os
from datetime import datetime
import schedule
import time
from analizador_empresas import AnalizadorEmpresas
from buscador_contratos import BuscadorContratos

def ejecutar_analisis_completo():
    # Configurar nombre del archivo de salida
    fecha_actual = datetime.now().strftime('%Y%m%d')
    archivo_salida = f'resultados_combinados_{fecha_actual}.xlsx'
    
    # Palabras clave para ambos análisis
    palabras_clave = [
        'estudio de flujo de carga',
        'cortocircuito',
        'coordinación de protecciones',
        'análisis de estabilidad',
        'arranque de motores',
        'análisis de armónicos',
        'análisis de fenómenos transitorios',
        'análisis de confiabilidad',
        'sistemas de puesta a tierra',
        'coordinación de aislamiento'
    ]

    try:
        # Ejecutar primer análisis (AnalizadorEmpresas)
        print("\n" + "="*50)
        print("🔍 Iniciando análisis de páginas web...")
        print("="*50)
        
        analizador = AnalizadorEmpresas()
        analizador.procesar_empresas(
            archivo_excel="C:/LinksEmpresas/codigos/EmpresasCol.xlsx",
            palabras_clave=palabras_clave,
            archivo_salida=archivo_salida,
            nombre_hoja='Análisis_Páginas'
        )

        # Ejecutar segundo análisis (BuscadorContratos)
        print("\n" + "="*50)
        print("🔍 Iniciando búsqueda de contratos...")
        print("="*50)
        
        buscador = BuscadorContratos()
        for empresa in buscador.empresas.keys():
            buscador.buscar_contratos(empresa, palabras_clave)
        buscador.guardar_resultados(
            archivo_salida=archivo_salida,
            nombre_hoja='Búsqueda_Contratos'
        )

        print("\n" + "="*50)
        print(f"✅ Análisis completo guardado en: {archivo_salida}")
        print("="*50)

    except Exception as e:
        print("\n" + "="*50)
        print(f"❌ Error durante la ejecución: {str(e)}")
        print("="*50)

if __name__ == "__main__":
    # Crear carpeta para resultados si no existe
    carpeta_resultados = 'resultados_analisisCol'
    if not os.path.exists(carpeta_resultados):
        os.makedirs(carpeta_resultados)

    print("\n" + "="*50)
    print("🚀 Iniciando programa de análisis")
    print("="*50)

    # Primera ejecución
    ejecutar_analisis_completo()

    # Programar ejecuciones futuras
    schedule.every().monday.at("09:00").do(ejecutar_analisis_completo)

    print("\n" + "="*50)
    print("⏰ Programa en espera de próxima ejecución programada")
    print("Presiona Ctrl+C para detener")
    print("="*50)

    # Bucle de ejecución programada
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)
    except KeyboardInterrupt:
        print("\n" + "="*50)
        print("👋 Programa detenido por el usuario")
        print("="*50)