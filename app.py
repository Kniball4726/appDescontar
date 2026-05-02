"""
Automatización para descuento de productos en Excel a partir de documentos Word.
Pasos:
1. Verificar y crear carpetas 'Descontar' y 'Descontados' en el directorio de ejecución.
2. Procesar archivos Word (.docx y .doc) en 'Descontar':
   - Extraer líneas con formato: (cantidad) (producto), total (cantidad) (
   producto), o con "etiquetas" después del número.
   - Mover archivos procesados a 'Descontados'.
   3. Crear carpeta con fecha actual en 'Descontar' y generar Excel consolidado con columnas "Cantidad" y "Producto".
    - El Excel se guardará en la carpeta del día con el nombre "consolidado_YYYY-MM-DD.xlsx".
Requisitos:
- Python 3.8+
- Librerías: python-docx, pandas, openpyxl (para Excel), pywin32 (solo para convertir .doc a .docx en Windows)
Nota:
- La conversión de .doc a .docx solo funciona en Windows con Microsoft Word instalado.
- El programa muestra una barra de progreso durante el procesamiento de archivos.
- Si no se encuentran archivos para procesar, se muestra un mensaje y se espera que el usuario presione Enter para salir.
- Si es el primer uso, se crean las carpetas necesarias y se solicita al usuario que coloque los archivos .docx en 'Descontar' antes de ejecutar nuevamente.
"""

import os
import re
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from docx import Document
import pandas as pd



def borrarPantalla():
    """Función para limpiar la pantalla de la consola.
    Detecta el sistema operativo y ejecuta el comando adecuado para limpiar la pantalla.
    En Windows, se utiliza 'cls', mientras que en sistemas Unix/Linux/Mac se utiliza 'clear'.

    Esta función es útil para mantener la consola limpia y mejorar la legibilidad de la salida del programa.
    """
    if os.name == "nt":
        os.system("cls")
    else:
        os.system("clear")


borrarPantalla()

def convertir_doc_a_docx(archivo_doc: Path, archivo_docx: Path) -> bool:
    """Convierte un archivo .doc a .docx utilizando Microsoft Word a través de COM.
    Solo funciona en Windows con Microsoft Word instalado.
            
        Args:
            archivo_doc (Path): Ruta del archivo .doc a convertir.
            archivo_docx (Path): Ruta donde se guardará el archivo .docx convertido.
        
        Returns:
            bool: True si la conversión fue exitosa, False en caso de error o si no se puede realizar la conversión.
        
        Nota:
            La función utiliza la biblioteca win32com para interactuar con Microsoft Word a través de COM.
            La función verifica si el sistema operativo es Windows antes de intentar la conversión.
            Si la biblioteca win32com no está disponible, se muestra un mensaje de error y se devuelve False.
            Si ocurre cualquier error durante la conversión, se captura la excepción, se muestra un mensaje de error y se devuelve False.
            Si la conversión es exitosa, se devuelve True.
    """
    if os.name != "nt":
        print("La conversión de archivos .doc solo está disponible en Windows.")
        return False

    try:
        from win32com.client import Dispatch
    except ImportError as exc:
        print(f"No se encuentra win32com: {exc}")
        return False

    word = None
    try:
        word = Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(archivo_doc), ReadOnly=1)
        doc.SaveAs(str(archivo_docx), FileFormat=16)
        doc.Close(False)
        return True
    except Exception as exc:
        print(f"Error convirtiendo {archivo_doc.name}: {exc}")
        return False
    finally:
        if word is not None:
            word.Quit()


def obtener_ruta_ejecucion() -> Path:
    """Obtiene la ruta del directorio de ejecución del programa.
    Esta función devuelve la ruta absoluta del directorio desde el cual se está ejecutando el programa.
    
    Returns:
        Path: La ruta absoluta del directorio de ejecución.
    
    Nota:
    - Se utiliza Path.cwd() para obtener el directorio de trabajo actual y resolve() para obtener la ruta absoluta. Esto asegura que se obtiene la ruta correcta independientemente de cómo se ejecute el programa (desde un IDE, terminal, etc.).
    """
    ruta_actual = Path.cwd().resolve()
    return ruta_actual


def obtener_ruta_carpeta(base: Path, nombres: list[str], crear: bool = False) -> Path:
    """Busca una carpeta dentro de la ruta base con alguno de los nombres especificados.
    
    Args:
        base (Path): La ruta base donde buscar la carpeta.
        nombres (list[str]): Lista de nombres de carpetas a buscar.
        crear (bool, optional): Si True, crea la carpeta si no existe. Por defecto es False.
    
    Returns:
        Path: La ruta de la carpeta encontrada o creada.
    
    Raises:
        FileNotFoundError: Si no se encuentra ninguna de las carpetas especificadas y no se crea.
    
    Nota:
    - La función itera sobre la lista de nombres de carpetas y verifica si alguna de ellas existe en la ruta base.
    - Si se encuentra una carpeta, se devuelve su ruta absoluta.    
    - Si no se encuentra ninguna carpeta y el parámetro 'crear' es True, se crea la primera carpeta de la lista y se devuelve su ruta.
    - Si no se encuentra ninguna carpeta y 'crear' es False, se lanza una excepción FileNotFoundError con un mensaje que indica que no se encontraron las carpetas esperadas.
    """
    for nombre in nombres:
        ruta = base / nombre
        if ruta.exists():
            return ruta.resolve()

    if crear:
        ruta = base / nombres[0]
        ruta.mkdir(parents=True, exist_ok=True)
        return ruta.resolve()

    raise FileNotFoundError(
        f"No se encontró ninguna de las carpetas esperadas en {base}: {', '.join(nombres)}"

    )


def imprimir_barra_progreso(actual: int, total: int, prefix: str = "Progreso:", suffix: str = "Completado", length: int = 40) -> None:
    """Imprime una barra de progreso en la consola.
    Args:
        actual (int): La cantidad de elementos procesados.
        total (int): La cantidad total de elementos.
        prefix (str, optional): El prefijo para la barra de progreso. Por defecto es "Progreso:".
        suffix (str, optional): El sufijo para la barra de progreso. Por defecto es "Completado".
        length (int, optional): La longitud de la barra de progreso. Por defecto es 40.
    Returns:
        None: Esta función no devuelve ningún valor, solo imprime la barra de progreso en la consola
        Nota:
        - Se utiliza para mostrar el progreso de una operación que se realiza en múltiples pasos.
        - La función calcula el porcentaje de progreso y la cantidad de caracteres '#' y '-' para representar visualmente el progreso.
        - Si el total es 0, la función no hace nada para evitar divisiones por cero.
        - Al finalizar el progreso (cuando actual es igual a total), se imprime una nueva línea para finalizar la barra de progreso.
    """
    if total == 0:
        return
    porcentaje = actual / total
    bloque = int(round(length * porcentaje))
    barra = "#" * bloque + "-" * (length - bloque)
    print(f"\r{prefix} |{barra}| {int(porcentaje * 100):3d}% {suffix}", end="")
    if actual == total:
        print()


def ejecutar_automatizacion():
    """Función principal que ejecuta la automatización para el descuento de productos.
    Esta función realiza los siguientes pasos:
    
     1. Configura las rutas necesarias para las carpetas 'Descontar' y 'Descontados' utilizando el directorio de ejecución actual.
     2. Verifica si ya se realizó el descuento hoy verificando la existencia de una carpeta con la fecha actual dentro de 'Descontar'.
     3. Procesa los archivos Word (.docx y .doc) encontrados en la carpeta 'Descontar', extrayendo la información de cantidad y producto utilizando expresiones regulares.
     4. Mueve los archivos procesados a la carpeta 'Descontados'.
     5. Crea un archivo Excel consolidado con la información extraída, agrupando por producto y sumando las cantidades, y lo guarda en una carpeta con la fecha actual dentro de 'Descontar'.
    Returns:
        None: Esta función no devuelve ningún valor, solo ejecuta la automatización.
     Nota:
     - La función maneja la creación de carpetas necesarias y verifica si es el primer uso para guiar al usuario a colocar los archivos .docx en la carpeta 'Descontar'.
     - Si ya se realizó el descuento hoy, muestra un mensaje de aviso y espera que el usuario presione Enter para salir.
     - Durante el procesamiento de archivos, muestra una barra de progreso para indicar el avance."""
    
    # 1. Configuración de rutas usando el directorio de ejecución actual    
    carpeta_base = obtener_ruta_ejecucion()
    print(f"Directorio de ejecución: {carpeta_base}")
    ruta_descontar = carpeta_base / "Descontar"
    ruta_descontados = carpeta_base / "Descontados"
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    ruta_hoy = ruta_descontar / fecha_hoy

    primer_uso = False
    if not ruta_descontar.exists():
        ruta_descontar.mkdir(parents=True, exist_ok=True)
        primer_uso = True
        print(f"Carpeta 'Descontar' creada en: {ruta_descontar}")

    if not ruta_descontados.exists():
        ruta_descontados.mkdir(parents=True, exist_ok=True)
        primer_uso = True
        print(f"Carpeta 'Descontados' creada en: {ruta_descontados}")

    if primer_uso:
        print(
            "Bienvenid@s para el primer uso del descuento de productos debe guardar el archivo.docx en la carpeta Descontar y correr nuevamente el programa"
        )
        input("Presione Enter para salir...")   
        return

    # Verificar si ya se realizó el descuento hoy
    if ruta_hoy.exists():
        print(f"Aviso: La carpeta '{fecha_hoy}' ya existe. Descuento ya realizado.")
        input("Presione Enter para salir...")
        return
    
    # 2. Verificar documentos Word (.docx) y .doc DENTRO de la carpeta 'descontar'
    archivos_word = [f for f in ruta_descontar.iterdir() if f.is_file() and f.suffix.lower() in {".docx", ".doc"}]

    if not archivos_word:
        print("No se encontraron documentos de Word (.docx) o .doc pendientes en la carpeta 'descontar'.")
        input("Presione Enter para salir...")
        return

    datos_extraidos = []
    total_archivos = len(archivos_word)
    procesados = 0
    print(f"Procesando {total_archivos} archivo(s) de descuento...")
    
    # Procesamiento de documentos
    for archivo in archivos_word:
        procesados += 1
        doc_temporal = None
        archivo_para_procesar = archivo

        if archivo.suffix.lower() == ".doc":
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            doc_temporal = Path(tempfile.gettempdir()) / f"{archivo.stem}_{timestamp}.docx"
            print(f"Convirtiendo {archivo.name} a {doc_temporal.name}...")
            if not convertir_doc_a_docx(archivo, doc_temporal):
                print(f"No se pudo convertir {archivo.name}. Se omite este archivo.")
                input("Presione Enter para continuar con el siguiente archivo...")
                continue
            archivo_para_procesar = doc_temporal

        try:
            doc = Document(archivo_para_procesar)
            for parrafo in doc.paragraphs:
                linea = parrafo.text.strip()
                if not linea:
                    continue
                
                # Regex para: (cantidad) (producto completo), total (cantidad) (producto completo),
                # o con la palabra etiquetas después del número.
                match = re.search(
                    r"^(?:total\s+)?(\d+)(?:\s+etiquetas)?\s+(.+)$",
                    linea,
                    re.IGNORECASE,
                )

                if match:
                    cantidad = int(match.group(1))
                    producto = match.group(2).strip()
                    datos_extraidos.append({"Producto": producto, "Cantidad": cantidad})

            # 3. Cortar y pegar (mover) a la carpeta 'descontados'
            destino = ruta_descontados / archivo.name
            
            # Si el archivo ya existe en el destino, le agregamos un timestamp para no sobrescribir
            if destino.exists():
                destino = ruta_descontados / f"{archivo.stem}_{datetime.now().strftime('%H%M%S')}{archivo.suffix}"
            
            shutil.move(str(archivo), str(destino))
            imprimir_barra_progreso(procesados, total_archivos, prefix="Procesando archivos:", suffix="completado")
            print(f"Archivo procesado y movido a 'descontados': {archivo.name}")

        except Exception as e:
            print(f"Error procesando {archivo.name}: {e}")

        finally:
            if doc_temporal and doc_temporal.exists():
                doc_temporal.unlink()
            imprimir_barra_progreso(procesados, total_archivos, prefix="Procesando archivos:", suffix="completado")

    # 4. Crear el archivo Excel consolidado
    if datos_extraidos:
        if not ruta_hoy.exists():
            ruta_hoy.mkdir(parents=True, exist_ok=True)
            print(f"Carpeta del día creada: {ruta_hoy}")

        df = pd.DataFrame(datos_extraidos)
        df_consolidado = df.groupby("Producto", as_index=False)["Cantidad"].sum()
        df_consolidado = df_consolidado[["Cantidad", "Producto"]]

        ruta_excel = ruta_hoy / f"consolidado_{fecha_hoy}.xlsx"
        df_consolidado.to_excel(ruta_excel, index=False)
        print(f"\nProceso exitoso. Excel creado en la carpeta del día: {ruta_excel}")
        input("Presione Enter para salir...")
    else:
        print("\nNo se pudo extraer información válida de los archivos.")
        input("Presione Enter para salir...")

if __name__ == '__main__':
    """Punto de entrada del programa.
    Se encarga de ejecutar la función principal de automatización y manejar cualquier excepción de tipo FileNotFoundError que pueda ocurrir durante la ejecución, mostrando el mensaje de error y finalizando el programa con un código de salida 1.
    
    Nota:
    - La función ejecutar_automatizacion() contiene toda la lógica principal del programa, y cualquier error relacionado con archivos no encontrados será capturado y manejado en este bloque principal.
    - Si ocurre un FileNotFoundError, se imprime el mensaje de error y se llama a sys.exit(1) para finalizar el programa indicando que hubo un error.
    
    """
    try:
        ejecutar_automatizacion()
    except FileNotFoundError as exc:
        print(exc)
        sys.exit(1)