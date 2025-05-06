import sys
import pandas as pd
import pyodbc
from datetime import datetime
import os

conexion_str = (
                    "DRIVER={ODBC Driver 17 for SQL Server};"
                    "SERVER=LAPTOP_ANDRES\SQLEXPRESS;" 
                    "DATABASE=TrepOficial;"
                    "Trusted_Connection=yes;"
                )

def int_or_none(valor):
    try:
        return int(valor)
    except:
        return None


def obtener_cantidad_habilitada(codigo_mesa, conexion_str):
    try:
        conn = pyodbc.connect(conexion_str)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT cantidadHabilitada 
            FROM MesaElectoral 
            WHERE codigoMesa = ?
        """, (codigo_mesa,))
        resultado = cursor.fetchone()
        return int(resultado[0]) if resultado else None
    except Exception as e:
        print("❌ Error al obtener cantidadHabilitada:", e)
        return None
    finally:
        conn.close()


def insertar_en_sql_server(nombre_archivo, datos, errores, conexion_str):
    try:
        conn = pyodbc.connect(conexion_str)
        cursor = conn.cursor()

        # Buscar idMesa desde MesaElectoral usando codigoMesa (ID_Mesa)
        cursor.execute("""
            SELECT TOP 1 idMesa FROM MesaElectoral WHERE codigoMesa = ?
        """, (datos["CodigoMesa"],))
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

        print("📤 Datos a insertar en SQL Server:")
        print({
            "papeletasNoUsadas": int_or_none(datos["PapeletasNoUsadas"]),
            "cantidadAnfora": int_or_none(datos["CantidadAnfora"]),
            "votosValidos": int_or_none(datos["Validos"]),
            "votosBlancos": int_or_none(datos["Blancos"]),
            "votosNulos": int_or_none(datos["Nulos"]),
            "observacion": observaciones,
            "estadoActa": estado_acta,
            "horaRegistro": hora_registro,
            "cantidadKB": peso_kb,
            "fk_idMesa": idMesa_fk,
            "MAS_IPSP": int_or_none(datos["Partido1"]),
            "SUMATE": int_or_none(datos["Partido2"]),
            "VOTEXCHI": int_or_none(datos["Partido3"]),
            "CHUPACOTO": int_or_none(datos["Partido4"])
        })


        cursor.execute("""
                INSERT INTO ActaElectoral (
                    papeletasNoUsadas, cantidadAnfora, votosValidos,
                    votosBlancos, votosNulos, observacion, estadoActa,
                    horaRegistro, cantidadKB, fk_idMesa,
                    MAS_IPSP, SUMATE, VOTEXCHI, CHUPACOTO
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                int_or_none(datos["PapeletasNoUsadas"]),
                int_or_none(datos["CantidadAnfora"]),
                int_or_none(datos["Validos"]),
                int_or_none(datos["Blancos"]),
                int_or_none(datos["Nulos"]),
                observaciones,
                estado_acta,
                hora_registro,
                peso_kb,
                idMesa_fk,
                int_or_none(datos["Partido1"]),
                int_or_none(datos["Partido2"]),
                int_or_none(datos["Partido3"]),
                int_or_none(datos["Partido4"])
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


def verificar_datos_fila(fila, i):
    try:
        datos = {k: int(fila[k]) for k in [
            "Partido1", "Partido2", "Partido3", "Partido4",
            "Blancos", "Validos", "CantidadAnfora",
            "PapeletasNoUsadas", "Nulos", "CodigoMesa"
        ]}
    except Exception as e:
        print(f"❌ Error fila {i}: datos inválidos - {e}")
        return
    cantidad_habilitada = obtener_cantidad_habilitada(datos["CodigoMesa"], conexion_str)
    errores = []

    # Verificar si hay valores vacíos o no encontrados
    if any(fila[k] in [None, "[NO ENCONTRADO]", ""] for k in datos):
        errores.append("Los datos no se leyeron correctamente |")

    # Reglas de consistencia
    try:
        suma_1 = datos["Partido1"] + datos["Partido2"] + datos["Partido3"] + datos["Partido4"] + datos["Blancos"]
        if suma_1 != datos["Validos"]:
            errores.append(
                f"Suma Votos Validos: Partido1({datos['Partido1']}) + Partido2({datos['Partido2']}) + Partido3({datos['Partido3']}) + Partido4({datos['Partido4']}) + Blancos({datos['Blancos']}) = {suma_1} != Validos({datos['Validos']}) |"
            )

        suma_2 = datos["CantidadAnfora"] + datos["PapeletasNoUsadas"]
        if cantidad_habilitada is None:
            errores.append(f"No se pudo obtener cantidadHabilitada para código {datos['CodigoMesa']} |")
        else:
            suma_2 = datos["CantidadAnfora"] + datos["PapeletasNoUsadas"]
            if suma_2 != cantidad_habilitada:
                errores.append(
                    f"Cantidad Total no Coincide: CantidadAnfora({datos['CantidadAnfora']}) + PapeletasNoUsadas({datos['PapeletasNoUsadas']}) = {suma_2} != cantidadHabilitada({cantidad_habilitada}) |"
                )

        suma_3 = datos["CantidadAnfora"] - datos["Nulos"]
        if suma_3 != datos["Validos"]:
            errores.append(
                f"Total Validos no coincide: CantidadAnfora({datos['CantidadAnfora']}) - Nulos({datos['Nulos']}) = {suma_3} != Validos + Blancos = {datos['Validos']}"
            )

    except Exception as e:
        print(f"❌ Error en reglas de validación: {e}")

    if errores:
        print("📌 Errores encontrados:")
        for e in errores:
            print("   -", e)
    else:
        print("✅ Verificación de consistencia correcta.")

    # Siempre insertar en la base de datos
    insertar_en_sql_server(sys.argv[1], datos, errores, conexion_str)
    return errores


def verificar_excel(ruta_excel):
    df = pd.read_excel(ruta_excel)
    for i, fila in df.iterrows():
        verificar_datos_fila(fila, i + 1)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python verificar_datos_excel.py <archivo_excel.xlsx>")
        sys.exit(1)
    verificar_excel(sys.argv[1])

