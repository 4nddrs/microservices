from datetime import datetime
import sys
import os
import fitz
import re
import pytesseract
import numpy as np
from PIL import Image, ImageDraw
from pdf2image import convert_from_path
import pyodbc

# Regiones para lectura con caracteres (fitz)
REGIONES_CARACTERES = {
    "ID_Mesa": [(90, 210, 160, 220)],
    "ID_Recinto": [(80, 220, 160, 240)],
    "numMesa": [(80, 260, 150, 320)],
    "total": [(90, 310, 160, 370)],
    "totalAnforas": [(80, 380, 190, 430)],
    "totalNoUtilizadas": [(80, 440, 190, 495)],
    "MAS": [(460, 160, 570, 215)],
    "SUMATE": [(460, 215, 570, 255)],
    "VOTEXCHI": [(460, 254, 570, 300)],
    "CHUPACOTO": [(460, 290, 570, 350)],
    "votosValidos": [(460, 370, 570, 420)],
    "votosBlancos": [(460, 420, 570, 465)],
    "votosNulos": [(460, 465, 570, 510)],
}

# Regiones para lectura desde imagen (OCR)
REGIONES_IMAGEN = {
    "ID_Mesa": [(90, 210, 170, 235)],
    "ID_Recinto": [(80, 230, 160, 250)],
    "numMesa": [(80, 270, 140, 310)],
    "total": [(90, 330, 160, 370)],
    "totalAnforas": [(80, 400, 190, 430)],
    "totalNoUtilizadas": [(80, 440, 190, 495)],
    "MAS": [(460, 180, 570, 215)],
    "SUMATE": [(460, 220, 570, 257)],
    "VOTEXCHI": [(460, 265, 568, 300)],
    "CHUPACOTO": [(460, 310, 570, 350)],
    "votosValidos": [(460, 380, 570, 420)],
    "votosBlancos": [(460, 420, 570, 465)],
    "votosNulos": [(460, 467, 570, 510)],
}

def int_or_zero(valor):
    try:
        return int(valor) if valor is not None else 0
    except:
        return 0 
    
def extraer_datos_fitz(pdf_path):
    datos = {}
    with fitz.open(pdf_path) as doc:
        pagina = doc[0]
        palabras = pagina.get_text("words")
        for etiqueta, zonas in REGIONES_CARACTERES.items():
            digitos = []
            for zona in zonas:
                x0, y0, x1, y1 = zona
                zona_digitos = []
                for palabra in palabras:
                    px0, py0, px1, py1, texto, *_ = palabra
                    texto = texto.strip()
                    if not re.fullmatch(r"\d+-?\d*", texto):
                        continue
                    if x0 <= px0 <= x1 and y0 <= py0 <= y1:
                        zona_digitos.append((px0, texto))
                zona_digitos.sort()
                digitos.extend([d[1] for d in zona_digitos])
            valor = "".join(digitos) if digitos else "[NO ENCONTRADO]"
            if etiqueta == "ID_Mesa" and valor != "[NO ENCONTRADO]":
                valor = valor[:5]
            if etiqueta == "numMesa" and valor != "[NO ENCONTRADO]":
                valor = valor[:-2]
            if etiqueta == "total" and valor != "[NO ENCONTRADO]":
                valor = valor[:-3]
            datos[etiqueta] = valor
    return datos

def resaltar_azul(imagen):
    np_image = np.array(imagen)
    azul_bajo = np.array([0, 0, 100])
    azul_alto = np.array([100, 100, 255])
    mask = np.all(np.logical_and(np_image >= azul_bajo, np_image <= azul_alto), axis=-1)
    np_image[mask] = [255, 255, 255]
    np_image[~mask] = [0, 0, 0]
    return Image.fromarray(np_image)

def detectar_rotacion(imagen):
    """Devuelve el ángulo necesario para rotar la imagen a su orientación correcta."""
    try:
        osd = pytesseract.image_to_osd(imagen)
        rotacion = int(re.search(r"Rotate: (\d+)", osd).group(1))
        return rotacion
    except Exception as e:
        print("⚠️ No se pudo detectar la rotación:", e)
        return 0  # Asumir que no necesita rotación si falla


def extraer_datos_ocr_faltantes(pdf_path, datos_actuales):
    imagenes = convert_from_path(pdf_path, dpi=122)
    imagen = imagenes[0]

    # Detectar rotación y corregirla
    angulo = detectar_rotacion(imagen)
    if angulo != 0:
        print(f"↩️ Imagen rotada {angulo}°, corrigiendo...")
        imagen = imagen.rotate(-angulo, expand=True)

    for etiqueta, zonas in REGIONES_IMAGEN.items():
        if datos_actuales[etiqueta] != "[NO ENCONTRADO]":
            continue
        for zona in zonas:
            x0, y0, x1, y1 = zona
            cropped = imagen.crop((x0, y0, x1, y1))
            texto = pytesseract.image_to_string(cropped, config="--psm 6 digits").strip()
            if not texto:
                cropped = resaltar_azul(cropped)
                texto = pytesseract.image_to_string(cropped, config="--psm 6 digits").strip()
            if texto:
                texto = texto.strip()
                texto = texto.replace(".", "")  # elimina puntos como en 3.17
                texto = re.sub(r"\D", "", texto)  # elimina cualquier carácter no numérico

                if etiqueta == "ID_Mesa":
                    texto = texto[:5]  # solo los primeros 5 caracteres

                datos_actuales[etiqueta] = texto if texto else "[NO ENCONTRADO]"
    return datos_actuales

def verificar_datos(datos):
    errores = []

    # Verificar si todos los campos tienen datos válidos
    for k in datos:
        if datos[k] in [None, "[NO ENCONTRADO]", ""]:
            errores.append("Los datos no se leyeron correctamente |")
            datos[k] = 0

    try:
        datos_int = {k: int(datos[k]) for k in datos if k not in ["ID_Mesa", "ID_Recinto", "numMesa"]}
    except ValueError as e:
        errores.append(f"Error al convertir datos a enteros: {e} |")
        return errores

    suma_1 = datos_int["MAS"] + datos_int["SUMATE"] + datos_int["VOTEXCHI"] + datos_int["CHUPACOTO"] + datos_int["votosBlancos"]
    if suma_1 != datos_int["votosValidos"]:
        errores.append(
            f"Suma Votos Validos: MAS({datos_int['MAS']}) + SUMATE({datos_int['SUMATE']}) + VOTEXCHI({datos_int['VOTEXCHI']}) + CHUPACOTO({datos_int['CHUPACOTO']}) + Blancos({datos_int['votosBlancos']}) = {suma_1} != votosValidos({datos_int['votosValidos']}) |"
        )

    suma_2 = datos_int["totalAnforas"] + datos_int["totalNoUtilizadas"]
    if suma_2 != datos_int["total"]:
        errores.append(
            f"Cantidad Total no Coincide: totalAnforas({datos_int['totalAnforas']}) + totalNoUtilizadas({datos_int['totalNoUtilizadas']}) = {suma_2} != total({datos_int['total']}) |"
        )

    suma_3 = datos_int["totalAnforas"] - datos_int["votosNulos"]
    if suma_3 != datos_int["votosValidos"]:
        errores.append(
            f"Total Validos no coincide: totalAnforas({datos_int['totalAnforas']}) - votosNulos({datos_int['votosNulos']}) = {suma_3} != votosValidos({datos_int['votosValidos']})"
        )

    if errores:
        print("📌 Errores encontrados:")
        for e in errores:
            print("   -", e)
    else:
        print("✅ Verificación de consistencia correcta.")
    return errores


def insertar_en_sql_server(nombre_archivo, datos, errores, conexion_str):
    try:
        conn = pyodbc.connect(conexion_str)
        cursor = conn.cursor()

        # Buscar idMesa desde MesaElectoral usando codigoMesa (ID_Mesa)
        cursor.execute("""
            SELECT TOP 1 idMesa FROM MesaElectoral WHERE codigoMesa = ?
        """, (datos["ID_Mesa"],))
        resultado = cursor.fetchone()

        if not resultado:
            print(f"❌ No se encontró idMesa para el código {datos['ID_Mesa']}. No se insertará el acta.")
            return

        idMesa_fk = resultado[0]

        # Preparar datos adicionales
        hora_registro = datetime.now()
        peso_kb = os.path.getsize(nombre_archivo) // 1024
        observaciones = "\n".join(errores) if errores else "No hay observaciones"
        estado_acta = "A" if errores else "V"

        cursor.execute("""
                INSERT INTO ActaElectoral (
                    papeletasNoUsadas, cantidadAnfora, votosValidos,
                    votosBlancos, votosNulos, observacion, estadoActa,
                    horaRegistro, cantidadKB, fk_idMesa,
                    MAS_IPSP, SUMATE, VOTEXCHI, CHUPACOTO
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                int_or_zero(datos["totalNoUtilizadas"]),
                int_or_zero(datos["totalAnforas"]),
                int_or_zero(datos["votosValidos"]),
                int_or_zero(datos["votosBlancos"]),
                int_or_zero(datos["votosNulos"]),
                observaciones,
                estado_acta,
                hora_registro,
                peso_kb,
                idMesa_fk,
                int_or_zero(datos["MAS"]),
                int_or_zero(datos["SUMATE"]),
                int_or_zero(datos["VOTEXCHI"]),
                int_or_zero(datos["CHUPACOTO"])
        ))
        cursor.execute("""
            UPDATE re
            SET re.estadoRecinto = 'Activo'
            FROM RecintoElectoral re
            JOIN MesaElectoral me ON re.idRecintoElectoral = me.fk_idRecintoElectoral
            WHERE me.idMesa = ?
        """, (idMesa_fk,))


        conn.commit()
        print("📥 Acta insertada correctamente en la base de datos.")

    except Exception as e:
        print("❌ Error al insertar en SQL Server:", e)
    finally:
        conn.close()

def int_or_none(valor):
    try:
        return int(valor)
    except:
        return None


def procesar_carpeta(carpeta_path):
    for archivo in os.listdir(carpeta_path):
        if archivo.lower().endswith(".pdf"):
            print(f"\n📄 Procesando: {archivo}")
            ruta_pdf = os.path.join(carpeta_path, archivo)
            datos = extraer_datos_fitz(ruta_pdf)

            if "[NO ENCONTRADO]" in datos.values():
                print("🔍 Datos faltantes detectados, aplicando OCR sobre imagen...")
                datos = extraer_datos_ocr_faltantes(ruta_pdf, datos)

            for k, v in datos.items():
                print(f"{k}: {v}")
            errores = verificar_datos(datos)
            conexion_str = (
                                "DRIVER={ODBC Driver 17 for SQL Server};"
                                "SERVER=LAPTOP_ANDRES\SQLEXPRESS;" 
                                "DATABASE=Trep;"
                                "Trusted_Connection=yes;"
                            )

            insertar_en_sql_server(ruta_pdf, datos, errores, conexion_str)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python verificar_datos_pdfs.py <carpeta_pdfs>")
        sys.exit(1)
    procesar_carpeta(sys.argv[1])
