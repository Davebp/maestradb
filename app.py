import os

import pandas as pd
import sqlite3
import streamlit as st
import io

# Crear una función para guardar las tablas en la base de datos
def guardar_tablas_en_bd(df1, df2, columnas_df1, columnas_df2):
    # Crear una conexión a la base de datos SQLite
    conn = sqlite3.connect('maestra.db')

    # Guardar las columnas del primer archivo en una tabla en la base de datos
    for columna in columnas_df1:
        if columna == 'nmalmcn':
            df1[columna] = df1[columna].astype(str)  # Convertir la columna en tipo de dato "texto"
        if columna == 'calmcn':
            df1[columna] = df1[columna].astype(str).str.zfill(5)
        if columna == 'fsrgstro':
            df1[columna] = pd.to_datetime(df1[columna]).dt.strftime('%Y-%m-%d')
        if columna == 'cartclo':
            df1[columna] = df1[columna].astype(str).str.zfill(14).str.replace(" ", "")
    df1[columnas_df1].to_sql('movimientos', conn, if_exists='replace', index=False)


    for columna2 in columnas_df2:
        if columna2 == 'nmalmcn':
            df2[columna2] = df2[columna2].astype(str)  # Convertir la columna en tipo de dato "texto"
        if columna2 == 'calmcn':
            df2[columna2] = df2[columna2].astype(str).str.zfill(5)
        if columna2 == 'fvlote':
            df2[columna2] = pd.to_datetime(df2[columna2]).dt.strftime('%Y-%m-%d')
        if columna2 == 'cartclo':
            df2[columna2] = df2[columna2].str.slice(10).str.zfill(14).str.replace(" ", "")
        if columna2 == 'comprobante':
            df2[columna2] = df2[columna2].str.slice(5)

    # Guardar las columnas del segundo archivo en otra tabla en la base de datos
    df2[columnas_df2].to_sql('kardex', conn, if_exists='replace', index=False)



    # Cerrar la conexión a la base de datos
    conn.close()

def borrar_datos_bd():
    # Conectar a la base de datos
    conn = sqlite3.connect('maestra.db')
    c = conn.cursor()

    # Obtener el nombre de todas las tablas en la base de datos
    c.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tablas = c.fetchall()
    tablas = [tabla[0] for tabla in tablas]

    # Borrar todas las tablas
    for tabla in tablas:
        c.execute(f"DROP TABLE {tabla}")

    # Cerrar la conexión a la base de datos
    conn.close()



def realizar_consulta():
    # Conectar a la base de datos SQLite
    conn = sqlite3.connect('maestra.db')
    c = conn.cursor()

    # Ejecutar la consulta
    c.execute('''
     SELECT 'MAPFRE PERU' AS codigo, movimientos.calmcn, movimientos.cartclo, movimientos.dartclo, strftime('%d/%m/%Y', movimientos.fsrgstro), movimientos.gmvmnto, 
        movimientos.ndrcpcndo, movimientos.des_cli, movimientos.cfmvmnto, movimientos.emfrccn, movimientos.tmalmcn, 
        movimientos.nmalmcn,movimientos.generico, movimientos.comercial, strftime('%d/%m/%Y',kardex.fvlote), kardex.lote, kardex.des_mov,'15-25° C' AS temperatura, julianday(kardex.fvlote) - julianday(movimientos.fsrgstro) AS  carta_canje, 
        CASE 
            WHEN julianday(kardex.fvlote) - julianday(movimientos.fsrgstro) < 366 THEN 'SI'
            ELSE 'NO'
        END AS carta
        FROM movimientos
        JOIN 
            kardex ON kardex.nmalmcn = movimientos.nmalmcn
            AND kardex.cartclo = movimientos.cartclo
            AND kardex.comprobante = movimientos.ndrcpcndo
        WHERE movimientos.ndrcpcndo IS NOT NULL AND kardex.comprobante IS NOT NULL
        UNION
        SELECT 'MAPFRE PERU' AS codigo, movimientos2.calmcn, movimientos2.cartclo, movimientos2.dartclo, strftime('%d/%m/%Y',movimientos2.fsrgstro), movimientos2.gmvmnto, 
        movimientos2.ndrcpcndo, movimientos2.des_cli, movimientos2.cfmvmnto, movimientos2.emfrccn, movimientos2.tmalmcn, 
        movimientos2.nmalmcn, movimientos2.generico, movimientos2.comercial, strftime('%d/%m/%Y',kardex.fvlote), kardex.lote, kardex.des_mov,'15-25° C' AS temperatura, julianday(kardex.fvlote) - julianday(movimientos2.fsrgstro) AS  carta_canje,
        CASE 
            WHEN julianday(kardex.fvlote) - julianday(movimientos2.fsrgstro) < 366 THEN 'SI'
            ELSE 'NO'
        END AS carta
        FROM 
            movimientos AS movimientos2
        LEFT JOIN 
            kardex ON kardex.nmalmcn = movimientos2.nmalmcn
            AND kardex.cartclo = movimientos2.cartclo
            AND kardex.tipo = movimientos2.tmalmcn
        
    ''')

    # Obtener los resultados de la consulta
    resultados = c.fetchall()

    # Cerrar la conexión a la base de datos
    conn.close()

    return resultados








# Configurar la página Streamlit
st.title("Cargar archivos de Excel y guardar tablas en una base de datos")

# Cargar los archivos de Excel mediante la interfaz de Streamlit
archivo1 = st.file_uploader("Cargar archivo movimientos", type="xls")
archivo2 = st.file_uploader("Cargar archivo kardex", type="xls")

# Verificar si los archivos han sido cargados
if archivo1 and archivo2:
    # Leer los archivos de Excel en DataFrames
    df1 = pd.read_excel(archivo1)
    df2 = pd.read_excel(archivo2)

    # Seleccionar las columnas de interés mediante la interfaz de Streamlit
    columnas_df1 = st.multiselect("Seleccionar columnas del archivo 1", df1.columns)
    columnas_df2 = st.multiselect("Seleccionar columnas del archivo 2", df2.columns)

    # Verificar si se han seleccionado columnas en ambos archivos
    if columnas_df1 and columnas_df2:
        # Mostrar un botón para guardar las tablas en la base de datos
        if st.button("Guardar tablas en la base de datos"):
            guardar_tablas_en_bd(df1, df2, columnas_df1, columnas_df2)
            st.success("Las tablas se han guardado exitosamente en la base de datos.")
    else:
        st.warning("Debes seleccionar al menos una columna en ambos archivos.")



st.title("Consulta TABLA")

# Ruta de la base de datos
db_path = os.path.join('maestra.db')

# Conexión a la base de datos
conn = sqlite3.connect(db_path)
c = conn.cursor()

# Obtener el nombre de todas las tablas en la base de datos
c.execute("SELECT name FROM sqlite_master WHERE type='table';")
tablas = c.fetchall()
tablas = [tabla[0] for tabla in tablas]

# Widget para seleccionar la tabla
tabla_seleccionada = st.selectbox("Seleccionar tabla", tablas)

if tabla_seleccionada:
    # Consultar todos los datos de la tabla seleccionada
    consulta = f"SELECT * FROM {tabla_seleccionada}"
    c.execute(consulta)
    resultados = c.fetchall()

    if resultados:
        # Crear un DataFrame con los resultados de la consulta
        columnas_df = [descripcion[0] for descripcion in c.description]
        datos_df = pd.DataFrame(resultados, columns=columnas_df)

        # Mostrar los datos en Streamlit
        st.write(datos_df)
    else:
        st.write("No se encontraron datos en la tabla seleccionada.")

conn.close()


# Mostrar un botón para borrar todos los datos de la base de datos
if st.button("Borrar todos los datos de la base de datos"):
    borrar_datos_bd()
    st.success("Se han borrado todos los datos, columnas y la tabla de la base de datos.")




# Configurar la página Streamlit
st.title("Consulta de coincidencias")

# Botón para realizar la consulta
if st.button("Consultar las coincidencias"):
    # Realizar la consulta
    resultados = realizar_consulta()

    # Mostrar los resultados en una tabla
    if len(resultados) > 0:
        st.write("Resultados:")
        df = pd.DataFrame(resultados, columns=["codigo","calmcn", "cartclo", "dartclo", "fsrgstro", "gmvmnto", "ndrcpcndo", "des_cli", "cfmvmnto", "emfrccn", "tmalmcn", "nmalmcn", "generico", "comercial", "fvlote", "lote", "des_mov","temperatura","carta_canje","carta"])
        st.dataframe(df)




    else:
        st.write("No se encontraron resultados.")



