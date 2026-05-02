import app

"""
Este módulo se encarga de imprimir la documentación de las funciones definidas en el módulo main.py.

"""
def imprimir_documentacion():   

    funciones = [app.borrarPantalla, app.obtener_ruta_ejecucion, app.obtener_ruta_carpeta, app.imprimir_barra_progreso, app.ejecutar_automatizacion, app.convertir_doc_a_docx ]
    for func in funciones:
        print(f"Función: {func.__name__}")
        print(f"Docstring: {func.__doc__}")
        print(f"Anotaciones: {func.__annotations__}")
        print("-" * 50)

if __name__ == "__main__":
    imprimir_documentacion()
    