# %%
# librerías necesarias

import os
import requests
import sqlite3
from openpyxl import load_workbook
import json
import math
from datetime import datetime, date
import unicodedata
import zipfile
import io

# %%
# necesaria para la función PintarMenu()

def EsEntero(num):
    try:
        k = int(num)
        return k
    except:
        return -1

# %%
def PintarMenu(listaOpc):
    """
    Muestra un menú y comprueba que la opción elegida es correcta
    Recibe listaOpc, que es una lista con las opciones que muestra
    Devuelve k, la opción elegida por el usuario
    """
    seguir = True
    while(seguir):
        for i,opcion in enumerate(listaOpc):
            print(f"{i}.-{opcion}")
        cad = input("Inserte una opción: ")
        k = EsEntero(cad)
        if 0<=k and k< len(listaOpc):
            return k 
        elif k == -1:
            print("\nDebe insertar un número entero.\n")
        else:
            print(f"\nEl número debe estar entre 0 y {len(listaOpc)-1}.\n")

# %%
def ComprobarCarpeta():
    if not os.path.exists("./datos"):
        os.makedirs("./datos")

# %%
# 1. CREAR BASE DE DATOS

def CrearBase():

    ComprobarCarpeta()

    try:
        fich = open("./datos/meteo.db","r")
        fich.close()
        print("La base de datos ya existe\n")
    except:
        with open("./datos/meteo.db", "w") as fich:
            pass
        print("Base de datos vacía creada\n")

# %%
# CrearBase()

# %%
def QuitarAcentos(texto):
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

# %%
# 2. DESCARGAR DATOS DE PROVINCIAS

def DescargarProvincias():

    print("Descargando datos...")
    
    try:
        r = requests.get("https://ucadrive.uca.es/index.php/s/iZQNdrybL9nnB2z/download")
        r.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return

    rutadic = "datos/diccionario.xlsx"

    with open(rutadic,"wb") as fich:
        fich.write(r.content)
    
    excel_municipios=load_workbook(rutadic)
    hojaMunicipio=excel_municipios['provincias']
    lCodigoProvincias=[]
    for i,linea in enumerate(hojaMunicipio):
        lCodigoProvincias.append((linea[0].value,QuitarAcentos((linea[1].value).upper())))

    # print(lCodigoProvincias)

    con = sqlite3.connect("./datos/meteo.db")
    cur=con.cursor()
    cur.executescript('''
        DROP TABLE IF EXISTS Provincias;
        CREATE TABLE Provincias(
            cprov INTEGER PRIMARY KEY,
            nombre varchar(20) NOT NULL UNIQUE
        );
    ''')

    con.commit()
    cur.executemany("INSERT INTO Provincias VALUES(?, ?)", lCodigoProvincias)
    con.commit()
    
    # cur.execute("SELECT * FROM Provincias")
    # resultado=cur.fetchall()
    # print(resultado)

    con.close()
    print("Se han insertado los datos en la tabla Provincias\n")


# %%
# DescargarProvincias()

# %%
EQUIVALENCIAS = {
    "BALEARES": "ISLAS BALEARES",
    "ILLES BALEARS": "ISLAS BALEARES",
    "STA. CRUZ DE TENERIFE": "SANTA CRUZ DE TENERIFE",
    "ARABA/ALAVA": "ALAVA",
    "CASTELLON": "CASTELLON DE LA PLANA",
    "ALACANT": "ALICANTE",
    "ARABA": "ALAVA",
    "GIPUZKOA": "GUIPUZCOA",
    "BIZKAIA": "VIZCAYA",
    "GIRONA": "GERONA",
    "LLEIDA": "LERIDA",
    "A CORUÑA": "LA CORUNA",
    "OURENSE": "ORENSE",
}

# %%
def ExisteTabla(cursor, nombre_tabla):
    cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{nombre_tabla}'")
    return cursor.fetchone() is not None

# %%
# 3. DESCARGAR DATOS DE ESTACIONES

def DescargarEstaciones(apikey):
    url = "https://opendata.aemet.es/opendata/api/valores/climatologicos/inventarioestaciones/todasestaciones/"
    querystring = {"api_key":apikey}
    headers = {'cache-control': "no-cache"}
    
    try:
        response = requests.request("GET", url, headers=headers, params=querystring)
        response.raise_for_status() 
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return

    dic1=json.loads(response.text)

    if "datos" not in dic1:
        raise RuntimeError("La API no devolvió la URL de datos")
        
    dic1['datos'] 

    if "datos" not in dic1:
        raise RuntimeError("La API no devolvió la URL de datos")

    try:
        r2 = requests.request("GET", dic1['datos'], headers=headers, params=querystring)
        r2.raise_for_status() 
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return

    dic2=json.loads(r2.text)

    con = sqlite3.connect("./datos/meteo.db")
    cur = con.cursor()
    infoEst = []

    if not ExisteTabla(cur, "Provincias"):
        print("La tabla Provincias no existe. Abortando...")
        return

    cur.executescript('''
        DROP TABLE IF EXISTS Estaciones;
        CREATE TABLE Estaciones(
            idema VARCHAR(5) PRIMARY KEY,
            nombre VARCHAR(20) NOT NULL,
            latitud VARCHAR(15),
            longitud VARCHAR(15),
            cprov INTEGER NOT NULL,
            FOREIGN KEY (cprov) REFERENCES Provincias(cprov)
                ON DELETE CASCADE
                ON UPDATE CASCADE
        );
    ''')
    con.commit()

    for x in dic2:
        lat = f"{int(x['latitud'][0:2])}º{int(x['latitud'][2:4])}'{int(x['latitud'][4:6])}\""

        if x['longitud'][-1] == "E":
            lon = f"{int(x['longitud'][0:2])}º{int(x['longitud'][2:4])}'{int(x['longitud'][4:6])}\""
        else:
            lon = f"-{int(x['longitud'][0:2])}º{int(x['longitud'][2:4])}'{int(x['longitud'][4:6])}\""

        prov = x['provincia'].upper()

        nombre_ = EQUIVALENCIAS.get(prov, prov)
        
        cur.execute(
            "SELECT cprov FROM Provincias WHERE nombre = ?",
            (nombre_,)
        )
        fila = cur.fetchone()
        if fila:
            cur.execute("""
                INSERT INTO Estaciones (idema, nombre, latitud, longitud, cprov)
                VALUES (?, ?, ?, ?, ?)
            """, (
                x['indicativo'],
                x['nombre'],
                lat,
                lon,
                fila[0]
            ))
        else:
            print(f"Provincia no encontrada: {nombre_}")
    con.commit()
    
    # cur.execute("SELECT * FROM Provincias")
    # resultado=cur.fetchall()
    # print(resultado)

    con.close()
    print("Se han insertado los datos en la tabla Estaciones\n")

# %%
# DescargarEstaciones(apikey)

# %%
# 4. DESCARGAR DATOS DE MUNICIPIOS

def DescargarMunicipios(apikey):
    url = "https://opendata.aemet.es/opendata/api/maestro/municipios/"
    querystring = {"api_key":apikey}
    headers = {'cache-control': "no-cache"}

    try:
        response = requests.request("GET", url, headers=headers, params=querystring)
        response.raise_for_status() 
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return

    dic1=json.loads(response.text)

    if "datos" not in dic1:
        raise RuntimeError("La API no devolvió la URL de datos")

    dic1['datos'] 

    if "datos" not in dic1:
        raise RuntimeError("La API no devolvió la URL de datos")

    try:
        r2 = requests.request("GET", dic1['datos'], headers=headers, params=querystring)
        r2.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return

    dic2=json.loads(r2.text)

    infoMun = []
    [infoMun.append((int(x['id'][2:4]),int(x['id'][2:]),x['nombre'],x['latitud'],x['longitud'])) for x in dic2]
    
    con = sqlite3.connect("./datos/meteo.db")
    cur = con.cursor()

    if not ExisteTabla(cur, "Provincias"):
        print("La tabla Provincias no existe. Abortando...")
        return
    
    cur.executescript('''
        PRAGMA foreign_keys = ON;
        DROP TABLE IF EXISTS Municipios;
        CREATE TABLE Municipios(
            cprov INTEGER NOT NULL,
            cmun INTEGER NOT NULL,
            nombre VARCHAR(50) NOT NULL,
            latitud VARCHAR(15),
            longitud VARCHAR(15),
            PRIMARY KEY (cprov, cmun),
            FOREIGN KEY (cprov) REFERENCES Provincias(cprov)
                ON DELETE CASCADE
                ON UPDATE CASCADE
        );
    ''')

    con.commit()
    cur.executemany("INSERT INTO Municipios VALUES(?, ?, ?, ?, ?)", infoMun)
    con.commit()
    
    # cur.execute("SELECT * FROM Provincias")
    # resultado=cur.fetchall()
    # print(resultado)

    con.close()
    print("Se han insertado los datos en la tabla Municipios\n")

# %%
# DescargarMunicipios(apikey)

# %%
# gms = "-34º34\'34.434222\""
# s = gms.strip()
# s = s.replace('º', ' ').replace('°', ' ')
# s = s.replace("''", ' ').replace('"', ' ').replace("'", ' ')
# partes = s.split()
# partes

# %%
def gms_a_decimal(gms):
    """
    Convierte un string formato "aºb'c''" a decimal.
    Si falla, devuelve 0.0.
    """
    if gms is None:
        return 0.0
    
    s = gms.strip()
    if s == "":
        return 0.0
    
    try:
        s = s.replace('º', ' ').replace('°', ' ')
        s = s.replace("''", ' ').replace('"', ' ').replace("'", ' ')
            
        partes = s.split()
        
        grados = 0.0
        minutos = 0.0
        segundos = 0.0
        
        if len(partes) > 0:
            grados = float(partes[0])
        if len(partes) > 1:
            minutos = float(partes[1])
        if len(partes) > 2:
            segundos = float(partes[2])
            
        resultado = grados + (minutos / 60) + (segundos / 3600)
        
        return math.radians(resultado)

    except Exception as e:
        # Esto imprime el error en la consola para que lo veas, 
        # pero devuelve 0.0 para que SQL siga trabajando.
        print(f"Error convirtiendo el valor '{gms}': {e}")
        return 0.0

# %%
# 5. CREACIÓN DE LA TABLA DISTANCIA

def CrearDistancia():
    con = sqlite3.connect("./datos/meteo.db")
    cur = con.cursor()

    if not ExisteTabla(cur, "Municipios"):
        print("La tabla Municipios no existe. Abortando...")
        return
    
    if not ExisteTabla(cur, "Estaciones"):
        print("La tabla Estaciones no existe. Abortando...")
        return
    
    cur.executescript('''
        PRAGMA foreign_keys = ON;
        DROP TABLE IF EXISTS Distancia;
        CREATE TABLE Distancia (
            cprov INTEGER NOT NULL,
            cmun INTEGER NOT NULL,
            idema VARCHAR(5) NOT NULL,
            kms DECIMAL(10, 4) NOT NULL,     
            PRIMARY KEY (cprov, cmun, idema),
            FOREIGN KEY (cprov, cmun) REFERENCES Municipios(cprov, cmun),
            FOREIGN KEY (idema) REFERENCES Estaciones(idema)
                ON DELETE CASCADE
                ON UPDATE CASCADE
        );
    ''')

    con.commit()

    print("Calculando distancias...")

    cur.execute("SELECT cprov, cmun, latitud, longitud FROM Municipios")
    resultadoMunicipios=cur.fetchall()

    cur.execute("SELECT idema, latitud, longitud FROM Estaciones")
    resultadoEstaciones=cur.fetchall()
    InfoDist = []
    R = 6371.0  # km
    for prov, mun, m_lat, m_lon in resultadoMunicipios:
        for id, e_lat, e_lon in resultadoEstaciones:
            
            distancia = 2 * R * math.asin(math.sqrt(math.sin((gms_a_decimal(m_lat) - gms_a_decimal(e_lat)) / 2)**2 + math.cos(gms_a_decimal(m_lat)) * math.cos(gms_a_decimal(e_lat)) * math.sin((gms_a_decimal(m_lon) - gms_a_decimal(e_lon)) / 2)**2)) 
            if distancia is not None:
                InfoDist.append((int(prov), int(mun), id, distancia))

    cur.executemany("INSERT INTO Distancia VALUES (?, ?, ?, ?)", InfoDist)
    con.commit()

    # cur.execute("SELECT * FROM Provincias")
    # resultado=cur.fetchall()
    # print(resultado)

    con.close()
    print("Se ha creado la tabla Distancia\n")

# %%
# CrearDistancia()

# %%
# 6. NUEVOS DATOS HORARIOS DE LAS ÚLTIMAS 12 HORAS

def DescargarNuevosDatos(apikey):
    url = "https://opendata.aemet.es/opendata/api/observacion/convencional/todas/"
    querystring = {"api_key":apikey}
    headers = {'cache-control': "no-cache"}

    try:
        response = requests.request("GET", url, headers=headers, params=querystring)
        response.raise_for_status() 
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return

    dic1=json.loads(response.text)

    if "datos" not in dic1:
        raise RuntimeError("La API no devolvió la URL de datos")

    dic1['datos'] 

    if "datos" not in dic1:
        raise RuntimeError("La API no devolvió la URL de datos")

    try:
        r2 = requests.request("GET", dic1['datos'], headers=headers, params=querystring)
        r2.raise_for_status() 
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return
        
    dic2=json.loads(r2.text)

    infoND = []
    info = []

    con = sqlite3.connect("./datos/meteo.db")
    cur = con.cursor()

    if not ExisteTabla(cur, "Estaciones"):
        print("La tabla Estaciones no existe. Abortando...")
        return

    cur.execute("SELECT idema FROM Estaciones")
    estaciones = {fila[0] for fila in cur.fetchall()}

    for x in dic2:
        if x['idema'] in estaciones:   
            dt = datetime.strptime(x['fint'], '%Y-%m-%dT%H:%M:%S%z')
            dt_ = dt.replace(tzinfo=None)  # quitar zona horaria
            datos = ['prec','vmax','vv','dv','hr','tamin','ta','tamax']
            for y in datos:
                try:
                    info.append(float(x[y]))
                except:
                    info.append(None) 
            infoND.append((x['idema'], dt_, info[0], info[1], info[2], info[3], info[4], info[5], info[6], info[7]))
        else:
            print(f"La estación con idema ", x["idema"]," no aparece en los datos horarios de las últimas 12 horas.")
        info.clear()

    cur = con.cursor()

    cur.executescript("""
        PRAGMA foreign_keys = ON;
        DROP TABLE IF EXISTS DatosHoras;
        CREATE TABLE DatosHoras(
            idema VARCHAR(5) NOT NULL,
            fint DATETIME NOT NULL,
            prec FLOAT,
            vmax FLOAT,
            vv FLOAT,
            dv FLOAT,
            hr FLOAT,
            tamin FLOAT,
            ta FLOAT,
            tamax FLOAT,
            PRIMARY KEY (idema, fint),
            FOREIGN KEY (idema) REFERENCES Estaciones(idema)
                ON DELETE CASCADE
                ON UPDATE CASCADE
        );
    """)

    con.commit()
    cur.executemany(
        "INSERT INTO DatosHoras VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", infoND
    )

    con.commit()
    con.close()
    # cur.execute("SELECT * FROM Provincias")
    # resultado=cur.fetchall()
    # print(resultado)
    
    print("Se han insertado los datos en la tabla DatosHoras\n")

# %%
# DescargarNuevosDatos(apikey)    

# %%
# 7. NUEVOS DATOS HORARIOS DE LAS ÚLTIMAS 12 HORAS DESDE FICHERO

def DescargarDatosHora():
    url = f"https://ucadrive.uca.es/index.php/s/H8DLzpnKjYitSMY/download?path=%2FficherosP2"
    carpeta = "datos"

    ComprobarCarpeta()

    print("Accediendo a la carpeta compartida...")

    try:
        response = requests.get(url)
        response.raise_for_status() 
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return

    cumplencondiciones = []
    with zipfile.ZipFile(io.BytesIO(response.content)) as archivo_zip:
        for i,fichero in enumerate(archivo_zip.namelist()):
            nombre = fichero.split("/")[-1]
            if nombre.endswith(".txt") and nombre.startswith("datosHora"):
                cumplencondiciones.append(nombre)

        if not cumplencondiciones:
            print("No se encontraron ficheros válidos.")
            return
        
        print("\nSelecciona uno de los ficheros:")
        seleccionado = PintarMenu(cumplencondiciones)

        print(f"\nFichero elegido: {cumplencondiciones[seleccionado]}")

        ruta_destino = os.path.join(carpeta, cumplencondiciones[seleccionado])

        carpetadrive = "ficherosP2/"
        
        with open(ruta_destino, 'wb') as f:
            f.write(archivo_zip.read(f"{carpetadrive}{cumplencondiciones[seleccionado]}"))
    
    with open(ruta_destino, 'r') as f:
        dic1 = json.load(f)

    con = sqlite3.connect("./datos/meteo.db")
    cur = con.cursor()

    if not ExisteTabla(cur, "Estaciones"):
        print("La tabla Estaciones no existe. Abortando...")
        return

    cur.execute("SELECT idema FROM Estaciones")
    estaciones = {fila[0] for fila in cur.fetchall()}

    infoND = []
    info = []

    for x in dic1:
        if x['idema'] in estaciones:   
            dt = datetime.strptime(x['fint'], '%Y-%m-%dT%H:%M:%S%z')
            dt_ = dt.replace(tzinfo=None)  # quitar zona horaria
            datos = ['prec','vmax','vv','dv','hr','tamin','ta','tamax']
            for y in datos:
                try:
                    info.append(float(x[y]))
                except:
                    info.append(None) 
            infoND.append((x['idema'], dt_, info[0], info[1], info[2], info[3], info[4], info[5], info[6], info[7]))
        else:
            print(f"La estación con idema ", x["idema"]," no aparece en los datos horarios del fichero.")
        info.clear()

    cur = con.cursor()

    cur.executescript("""
        PRAGMA foreign_keys = ON;
        DROP TABLE IF EXISTS DatosHoras;
        CREATE TABLE DatosHoras(
            idema VARCHAR(5) NOT NULL,
            fint DATETIME NOT NULL,
            prec FLOAT,
            vmax FLOAT,
            vv FLOAT,
            dv FLOAT,
            hr FLOAT,
            tamin FLOAT,
            ta FLOAT,
            tamax FLOAT,
            PRIMARY KEY (idema, fint),
            FOREIGN KEY (idema) REFERENCES Estaciones(idema)
                ON DELETE CASCADE
                ON UPDATE CASCADE
        );
    """)

    con.commit()
    cur.executemany(
        "INSERT INTO DatosHoras VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", infoND
    )

    con.commit()
    con.close()
    # cur.execute("SELECT * FROM Provincias")
    # resultado=cur.fetchall()
    # print(resultado)

    print(f"Se ha cargado en la base de datos el contenido de {cumplencondiciones[seleccionado]}.")


# %%
# DescargarDatosHora()

# %%
# 8. NUEVOS DATOS CLIMATOLÓGICOS DIARIOS

def PedirFechas():
    seguir = True
    while seguir:
        print("Introduzca dos fechas que no estén separadas entre más de 15 días en el siguiente formato: AAAA-MM-DD\n")
        f1 = input("\t- Fecha 1: ")
        f2 = input("\t- Fecha 2: ")
        try:
            fecha1 = datetime.strptime(f1, "%Y-%m-%d").date()
            fecha2 = datetime.strptime(f2, "%Y-%m-%d").date()
            hoy = date.today()

            if fecha1 > hoy or fecha2 > hoy:
                raise ValueError("Las fechas deben ser anteriores al día de hoy")
            if fecha1 > fecha2:
                raise ValueError("La Fecha 1 debe ser anterior o igual a la Fecha 2")
            if (fecha2 - fecha1).days > 15:
                raise ValueError("Las fechas no pueden separarse más de 15 días")

            return fecha1.strftime("%Y-%m-%d"), fecha2.strftime("%Y-%m-%d")

        except ValueError as e:
            print("\nError:", e)
            print("Por favor, introduzca las fechas correctamente.\n")


def DatosEntreFechas(apikey):
    print("\nEsta opción te permite insertar en la base de datos los valores entre dos fechas.")

    f1,f2 = PedirFechas()
    url = f"https://opendata.aemet.es/opendata/api/valores/climatologicos/diarios/datos/fechaini/{f1}T00%3A00%3A00UTC/fechafin/{f2}T23%3A59%3A59UTC/todasestaciones"
    
    querystring = {"api_key":apikey}
    headers = {'cache-control': "no-cache"}

    try:
        response = requests.get(url, headers=headers, params=querystring)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return

    dic1 = response.json() 

    if "datos" not in dic1:
        raise RuntimeError("La API no devolvió la URL de datos")

    try:
        r2 = requests.request("GET", dic1['datos'], headers=headers, params=querystring)
        r2.raise_for_status() 
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return

    dic2=json.loads(r2.text)

    infoND = []
    info = []
    infohoras = []

    con = sqlite3.connect("./datos/meteo.db")
    cur = con.cursor()

    if not ExisteTabla(cur, "Estaciones"):
        print("La tabla Estaciones no existe. Abortando...")
        return

    cur.execute("SELECT idema FROM Estaciones")
    estaciones = {fila[0] for fila in cur.fetchall()}

    for x in dic2:
        if x['indicativo'] in estaciones:   
            fecha = datetime.strptime(x['fecha'], "%Y-%m-%d").date()
            datoshora = ['horaracha','horahrMax','horahrMin','horatmin','horatmax']
            datos = ['prec','racha','velmedia','dir','hrMedia','hrMax','hrMin','tmin','tmax']
            
            for y in datoshora:
                try:
                    infohoras.append(datetime.strptime(x[y], "%H:%M").strftime("%H:%M"))
                except:
                    infohoras.append(None)

            for y in datos:
                try:
                    info.append(float(x[y]))
                except:
                    info.append(None) 

            infoND.append((x['indicativo'], fecha, info[0], info[1], info[2], infohoras[0], info[3], info[4], info[5], infohoras[1], info[6], 
                        infohoras[2], info[7], infohoras[3], info[8], infohoras[4]))
        else:
            print(f"La estación con idema ", x["indicativo"]," no aparece en los datos diarios de entre {f1} y {f2}.")
        info.clear()
        infohoras.clear()

    cur = con.cursor()

    cur.executescript("""
        PRAGMA foreign_keys = ON;
        DROP TABLE IF EXISTS DatosDias;
        CREATE TABLE DatosDias(
            idema VARCHAR(5) NOT NULL,
            fecha DATE NOT NULL,
            prec FLOAT,
            racha FLOAT,
            velmedia FLOAT,
            horaracha TIME,
            dv FLOAT,
            hrMedia FLOAT,
            hrMax FLOAT,
            horahrMax TIME,
            hrMin FLOAT,
            horahrMin TIME,
            tmin FLOAT,
            horatmin TIME,
            tmax FLOAT,
            horatmax TIME,
            PRIMARY KEY (idema, fecha),
            FOREIGN KEY (idema) REFERENCES Estaciones(idema)
                ON DELETE CASCADE
                ON UPDATE CASCADE
        );
    """)

    con.commit()
    cur.executemany(
        "INSERT INTO DatosDias VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", infoND
    )

    con.commit()
    con.close()
    # cur.execute("SELECT * FROM Provincias")
    # resultado=cur.fetchall()
    # print(resultado)
    
    print("Se han insertado los datos en la tabla DatosDias\n")
    

# %%
# DatosEntreFechas(apikey)

# %%
# 9. NUEVOS DATOS CLIMATOLÓGICOS DIARIOS DESDE FICHERO

def DescargarDatosDias():
    url = f"https://ucadrive.uca.es/index.php/s/H8DLzpnKjYitSMY/download?path=%2FficherosP2"
    carpeta = "datos"

    ComprobarCarpeta()

    print("Accediendo a la carpeta compartida...")

    try:
        response = requests.get(url)
        response.raise_for_status() 
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión: {e}")
        return

    cumplencondiciones = []
    with zipfile.ZipFile(io.BytesIO(response.content)) as archivo_zip:
        for i,fichero in enumerate(archivo_zip.namelist()):
            nombre = fichero.split("/")[-1]
            if nombre.endswith(".txt") and nombre.startswith("datosDias"):
                cumplencondiciones.append(nombre)

        if not cumplencondiciones:
            print("No se encontraron ficheros válidos.")
            return
        
        print("\nSelecciona uno de los ficheros:")
        seleccionado = PintarMenu(cumplencondiciones)

        print(f"\nFichero elegido: {cumplencondiciones[seleccionado]}")

        ruta_destino = os.path.join(carpeta, cumplencondiciones[seleccionado])

        carpetadrive = "ficherosP2/"
        
        with open(ruta_destino, 'wb') as f:
            f.write(archivo_zip.read(f"{carpetadrive}{cumplencondiciones[seleccionado]}"))
    
    with open(ruta_destino, 'r') as f:
        dic1 = json.load(f)

    con = sqlite3.connect("./datos/meteo.db")
    cur = con.cursor()

    if not ExisteTabla(cur, "Estaciones"):
        print("La tabla Estaciones no existe. Abortando...")
        return

    cur.execute("SELECT idema FROM Estaciones")
    estaciones = {fila[0] for fila in cur.fetchall()}

    infoND = []
    info = []
    infohoras = []

    for x in dic1:
        if x['indicativo'] in estaciones:   
            fecha = datetime.strptime(x['fecha'], "%Y-%m-%d").date()
            datoshora = ['horaracha','horahrMax','horahrMin','horatmin','horatmax']
            datos = ['prec','racha','velmedia','dir','hrMedia','hrMax','hrMin','tmin','tmax']
            
            for y in datoshora:
                try:
                    infohoras.append(datetime.strptime(x[y], "%H:%M").strftime("%H:%M"))
                except:
                    infohoras.append(None)

            for y in datos:
                try:
                    info.append(float(x[y]))
                except:
                    info.append(None) 

            infoND.append((x['indicativo'], fecha, info[0], info[1], info[2], infohoras[0], info[3], info[4], info[5], infohoras[1], info[6], 
                        infohoras[2], info[7], infohoras[3], info[8], infohoras[4]))
        else:
            print(f"La estación con idema ", x["indicativo"]," no aparece en los datos diarios de entre {f1} y {f2}.")
        info.clear()
        infohoras.clear()

    cur = con.cursor()

    cur.executescript("""
        PRAGMA foreign_keys = ON;
        DROP TABLE IF EXISTS DatosDias;
        CREATE TABLE DatosDias(
            idema VARCHAR(5) NOT NULL,
            fecha DATE NOT NULL,
            prec FLOAT,
            racha FLOAT,
            velmedia FLOAT,
            horaracha TIME,
            dv FLOAT,
            hrMedia FLOAT,
            hrMax FLOAT,
            horahrMax TIME,
            hrMin FLOAT,
            horahrMin TIME,
            tmin FLOAT,
            horatmin TIME,
            tmax FLOAT,
            horatmax TIME,
            PRIMARY KEY (idema, fecha),
            FOREIGN KEY (idema) REFERENCES Estaciones(idema)
                ON DELETE CASCADE
                ON UPDATE CASCADE
        );
    """)

    con.commit()
    cur.executemany(
        "INSERT INTO DatosDias VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", infoND
    )

    con.commit()
    con.close()
    # cur.execute("SELECT * FROM Provincias")
    # resultado=cur.fetchall()
    # print(resultado)

    print(f"Se ha cargado en la base de datos el contenido de {cumplencondiciones[seleccionado]}.")

# %%
# DescargarDatosDias()

# %%
# 10. 

def Formatear(valor):
    if valor is None:
        return "N/A"
    return round(valor, 2)

def PedirFloat(mensaje):
    while True:
        try:
            return float(input(mensaje))
        except ValueError:
            print("Entrada inválida. Por favor, introduce un número.")

def MenuSecundarioApartado10(con, cprov, cmun, nombre_mun):
    cur = con.cursor()

    seguir = True
    while seguir:
        listaOpciones = ["Volver al menú principal",
                        "Estaciones más cercanas",
                        "Balance día hora",
                        "Balance día", 
                        "Predicción día municipio"]
        
        k = PintarMenu(listaOpciones)

        match k:
            case 0:
                print("\nVolviendo al menú principal...\n")
                seguir = False
            case 1:
                try:
                    distancia_max = PedirFloat("Introduce la distancia máxima en Kms: ")

                    if not ExisteTabla(cur, "Distancia"):
                        print("La tabla Distancia no existe. Abortando...")
                        continue

                    if not ExisteTabla(cur, "Estaciones"):
                        print("La tabla Estaciones no existe. Abortando...")
                        continue

                    consulta = """
                    SELECT E.nombre, D.kms, E.idema FROM Distancia D JOIN Estaciones E ON D.idema = E.idema
                    WHERE D.cprov = ? AND D.cmun = ? AND D.kms <= ?
                    ORDER BY D.kms ASC
                    """
                    cur.execute(consulta, (cprov, cmun, distancia_max))
                    estaciones = cur.fetchall()
                    
                    print(f"\nEstaciones a menos de {distancia_max} km:")
                    if not estaciones:
                        print("No hay estaciones en ese rango.")
                    for est in estaciones:
                        print(f"- {est[0]} ({est[2]}): a {est[1]:.2f} km")
                except ValueError:
                    print("Por favor, introduce un número válido.")

            case 2:
                try:
                    if not ExisteTabla(cur, "DatosHoras"):
                        print("La tabla DatosHoras no existe. Abortando...")
                        continue

                    cur.execute("SELECT MIN(fint), MAX(fint) FROM DatosHoras")
                    min_f, max_f = cur.fetchone()

                    if not min_f:
                        print("No hay datos horarios disponibles en la base de datos.")
                        continue

                    print(f"\nRango disponible: {min_f} hasta {max_f}")                   

                    if input("¿Desea continuar? (Y/N): ").upper() != "Y":
                        continue

                    fecha_str = input("Introduce fecha (AAAA-MM-DD): ")
                    hora_str = input("Introduce hora (HH:MM): ")

                    fint = datetime.strptime(f"{fecha_str} {hora_str}", "%Y-%m-%d %H:%M")

                    fint_ = fint.strftime("%Y-%m-%d %H:%M:%S")

                    if not ExisteTabla(cur, "Distancia"):
                        print("La tabla Distancia no existe. Abortando...")
                        continue

                    if min_f <= fint_ <= max_f:
                        dist_max = PedirFloat("Introduce la distancia máxima en Kms: ")
                        consulta = """
                        SELECT H.idema, H.prec, H.vmax, H.vv, H.dv, H.hr, H.tamin, H.ta, H.tamax, D.kms,
                        AVG(H.prec), AVG(H.vmax), AVG(H.vv), AVG(H.dv), AVG(H.hr), AVG(H.tamin), AVG(H.ta), AVG(H.tamax), AVG(D.kms)
                        FROM Distancia D
                        JOIN DatosHoras H ON D.idema = H.idema
                        WHERE H.fint = ? AND D.cprov = ? AND D.cmun = ? AND D.kms <= ?
                        ORDER BY D.kms ASC
                        """
                        cur.execute(consulta, (fint_, cprov, cmun, dist_max))
                        res = cur.fetchall()

                        if res:
                            print(f"\nDatos para {fint_}:")
                            for r in res:
                                print(f"ID: {r[0]} | Prec: {Formatear(r[1])} | Racha: {Formatear(r[2])}km/h | VMedia: {Formatear(r[3])}km/h | DV: {Formatear(r[4])} | Hum: {Formatear(r[5])}% | TempMin: {Formatear(r[6])}°C | Temp: {Formatear(r[7])}°C | TempMax: {Formatear(r[8])}°C | Dist: {Formatear(r[9])}km")

                            print(f"Valor medio | Prec: {Formatear(r[10])} | Racha: {Formatear(r[11])}km/h | VMedia: {Formatear(r[12])}km/h | DV: {Formatear(r[13])} | Hum: {Formatear(r[14])}% | TempMin: {Formatear(r[15])}°C | Temp: {Formatear(r[16])}°C | TempMax: {Formatear(r[17])}°C | Dist: {Formatear(r[18])}km")
                        else:
                            print("No se encontraron datos para los criterios seleccionados.")
                    else:
                        print("Error: La fecha está fuera del rango disponible.")
    
                except ValueError:
                    print("Error: Formato de datos incorrecto.")

            case 3:
                try:
                    if not ExisteTabla(cur, "DatosDias"):
                        print("La tabla DatosDias no existe. Abortando...")
                        continue

                    cur.execute("SELECT MIN(fecha), MAX(fecha) FROM DatosDias")
                    min_str, max_str = cur.fetchone()

                    min_f = date.fromisoformat(min_str) 
                    max_f = date.fromisoformat(max_str)

                    if not min_f:
                        print("No hay datos diarios disponibles en la base de datos.")
                        continue

                    print(f"\nRango disponible: {min_f} hasta {max_f}")                   

                    if input("¿Desea continuar? (Y/N): ").upper() != "Y":
                        continue

                    fecha_str = input("Introduce fecha (AAAA-MM-DD): ")
                    fint = datetime.strptime(fecha_str, "%Y-%m-%d").date()

                    if not ExisteTabla(cur, "Distancia"):
                        print("La tabla Distancia no existe. Abortando...")
                        continue

                    if min_f <= fint <= max_f:
                        dist_max = PedirFloat("Introduce la distancia máxima en Kms: ")
                        consulta = """
                        SELECT H.idema, H.prec, H.racha, H.velmedia, H.dv, H.hrMedia, H.hrMax, H.hrMin, H.tmin, H.tmax, D.kms, 
                        AVG(H.prec), AVG(H.racha), AVG(H.velmedia), AVG(H.dv), AVG(H.hrMedia), AVG(H.hrMax), AVG(H.hrMin), AVG(H.tmin), AVG(H.tmax), AVG(D.kms)
                        FROM Distancia D
                        JOIN DatosDias H ON D.idema = H.idema
                        WHERE H.fecha = ? AND D.cprov = ? AND D.cmun = ? AND D.kms <= ?
                        ORDER BY D.kms ASC
                        """
                        cur.execute(consulta, (fint, cprov, cmun, dist_max))
                        res = cur.fetchall()

                        if res:
                            print(f"\nDatos para {fint}:")
                            for r in res:
                                print(f"ID: {r[0]}   | Prec: {Formatear(r[1])} | Racha: {Formatear(r[2])}km/h | VMedia: {Formatear(r[3])}km/h | DV: {Formatear(r[4])} | HumMed: {Formatear(r[5])}% | HumMax: {Formatear(r[6])}% | HumMin: {Formatear(r[7])}% | TempMin: {Formatear(r[8])}°C | TempMax: {Formatear(r[9])}°C | Dist: {Formatear(r[10])}km")

                            print(f"Valor medio | Prec: {Formatear(r[11])} | Racha: {Formatear(r[12])}km/h | VMedia: {Formatear(r[13])}km/h | DV: {Formatear(r[14])}km/h| HumMed: {Formatear(r[15])}% | HumMax: {Formatear(r[16])}% | HumMin: {Formatear(r[17])}% | TempMin: {Formatear(r[18])}°C | TempMax: {Formatear(r[19])}°C | Dist: {Formatear(r[20])}km")
                        else:
                            print("No se encontraron datos para los criterios seleccionados.")
                    else:
                        print("Error: La fecha está fuera del rango disponible.")
    
                except ValueError:
                    print("Error: Formato de datos incorrecto.")
            case 4:

                url = f"https://opendata.aemet.es/opendata/api/prediccion/especifica/municipio/diaria/{cmun}"

                fich = open("./datos/keyAEMET.key","r")
                    apikey = fich.read()
                    fich.close()

                querystring = {"api_key":apikey}
                headers = {'cache-control': "no-cache"}

                try:
                    response = requests.get(url, headers=headers, params=querystring)
                    response.raise_for_status()
                except requests.exceptions.RequestException as e:
                    print(f"Error de conexión: {e}")
                    return

                dic1 = response.json() 

                if "datos" not in dic1:
                    raise RuntimeError("La API no devolvió la URL de datos")

                try:
                    r2 = requests.request("GET", dic1['datos'], headers=headers, params=querystring)
                    r2.raise_for_status()
                except requests.exceptions.RequestException as e:
                    print(f"Error de conexión: {e}")
                    return

                dic2=json.loads(r2.text)

                cur.executescript("""
                    PRAGMA foreign_keys = ON;
                    DROP TABLE IF EXISTS PrediccionesDiaMun;        
                    CREATE TABLE PrediccionesDiaMun (
                        cprov INTEGER NOT NULL,
                        cmun INTEGER NOT NULL,
                        fecha DATE NOT NULL,
                        prediccion TEXT,
                        PRIMARY KEY (cprov, cmun, fecha),
                        FOREIGN KEY (cprov, cmun) REFERENCES Municipios(cprov, cmun)
                            ON DELETE CASCADE
                            ON UPDATE CASCADE
                    );
                """)

                con.commit()

                info = []

                print(f"\nPredicción para los próximos {len(dic2[0]['prediccion']['dia'])} días en {nombre_mun}:\n")

                for dia in dic2[0]['prediccion']['dia']:
                    fecha_obj = datetime.fromisoformat(dia['fecha']).date()
                    
                    desc = next((e['descripcion'] for e in dia['estadoCielo'] if e['descripcion']), "Sin datos")
                    t_max = dia['temperatura']['maxima']
                    t_min = dia['temperatura']['minima']
                    prob_lluvia = dia['probPrecipitacion'][0]['value']
                    
                    descripcion = f"Cielo: {desc}. Temp: {t_min}/{t_max}C. Lluvia: {prob_lluvia}%."
                    print(f"\t Día {fecha_obj}:")
                    print(f"\t\t- Se esperan cielos {desc}.")
                    print(f"\t\t- La temperatura máxima será: {t_max}ºC.")
                    print(f"\t\t- La temperatura mínima será: {t_min}ºC.")
                    print(f"\t\t- Habrá una probabilidad del {prob_lluvia}% de lluvia.")

                    info.append((cprov, cmun, fecha_obj,descripcion))
                cur.executemany(
                    "INSERT INTO PrediccionesDiaMun VALUES (?, ?, ?, ?)", info#
                )

                con.commit()
    
def MenuPrincipalApartado10():
    con = sqlite3.connect("./datos/meteo.db")
    cur = con.cursor()
    
    busqueda = input("Introduce parte del nombre del municipio que quiere buscar: ")
    
    if not ExisteTabla(cur, "Municipios"):
        print("La tabla Municipios no existe. Abortando...")
        return

    consulta = "SELECT cprov, cmun, nombre FROM Municipios WHERE nombre LIKE ?"
    cur.execute(consulta, (f'%{busqueda}%',))
    resultados = cur.fetchall()
    
    if not resultados:
        print("No se encontraron municipios que contengan esa cadena de caracteres.")
        return

    print(f"\nSe han encontrado {len(resultados)} municipios que contienen esa cadena de caracteres, elija uno:")
    seleccion = PintarMenu([x[2] for x in resultados])
    
    if 0 <= seleccion < len(resultados):
        municipio_elegido = resultados[seleccion]
        cprov, cmun, nombre_mun = municipio_elegido
        print(f"\nHas seleccionado el municipio {nombre_mun}, con código {cmun}")
        
        MenuSecundarioApartado10(con, cprov, cmun, nombre_mun)
    else:
        print("Selección no válida.")
    con.close()

# %%
# MenuPrincipalApartado10()

# %%
def ExisteBase():
    if os.path.exists("./datos/meteo.db"):
        return(True)
    else:
        return(False)

# %%
def ExisteApiKey():
    try:
        fich=open("./datos/keyAEMET.key","r")
        fich.close()
        return True
    except:
        return False

# %%
def Guardar_API_key():
    API_key = input("Introduce tu API key de AEMET: \n").strip()
    carpeta = "datos/"
    os.makedirs(carpeta, exist_ok=True)  
    ruta_archivo = os.path.join(carpeta, "keyAEMET.key") 
    with open(ruta_archivo, "w") as f: 
        f.write(API_key) 
    print(f"Tu API key ha sido guardada correctamente en {ruta_archivo}")

# %%
def main():
    print("---- P R Á C T I C A  2 :  B A S E  D E  D A T O S,  A N Á L I S I S  D E L  T I E M P O ----\n\nTrabajo realizado por : \n\tAndrea Sayago Butrón")
    print("\n╔══════════════════════════════════╗")
    print("║              MENÚ                ║")
    print("╚══════════════════════════════════╝")
    print("\nNota: la primera opción que debe ejecutar es la 1.")
    seguir = True
    while seguir:
        listaOpciones = ["Salir",
                        "Crear base de datos (necesario para las demás opciones)",
                        "Descargar datos de provincias",
                        "Descargar datos de estaciones", 
                        "Descargar datos de municipios",
                        "Creación de la tabla \"distancia\"",
                        "Nuevos datos horarios de las últimas 12 horas",
                        "Nuevos datos horarios de las últimas 12 horas desde fichero",
                        "Nuevos datos climatológicos diarios",
                        "Nuevos datos horarios de las últimas 12 horas desde fichero",
                        "Buscar municipios",
                        "Introducir API key (necesario para las opciones: 3, 4, 6, 8 y 10)"]
        k = PintarMenu(listaOpciones)
        print(f"\nOpción elegida: {listaOpciones[k]}\n")
        match k:
            case 0:
                print("\nSaliendo...\n")
                seguir = False
            case 1:
                CrearBase()
            case 2:
                if ExisteBase():
                    DescargarProvincias()
                else:
                    print("Antes de ejecutar esta opción debe ejecutar la opción 1 (Crear base de datos)")
            case 3:
                if ExisteBase() and ExisteApiKey():
                    fich = open("./datos/keyAEMET.key","r")
                    apikey = fich.read()
                    fich.close()
                    DescargarEstaciones(apikey)
                else:
                    print("Antes de ejecutar esta opción debe ejecutar la opción 1 y 11 (Crear base de datos e Introducir API key)")    
            case 4:
                if ExisteBase() and ExisteApiKey():
                    fich = open("./datos/keyAEMET.key","r")
                    apikey = fich.read()
                    fich.close()
                    DescargarMunicipios(apikey)
                else:
                    print("Antes de ejecutar esta opción debe ejecutar la opción 1 y 11 (Crear base de datos e Introducir API key)")
                
            case 5:
                if ExisteBase():
                    CrearDistancia()
                else:
                    print("Antes de ejecutar esta opción debe ejecutar la opción 1 (Crear base de datos)")
                
            case 6:
                if ExisteBase() and ExisteApiKey():
                    fich = open("./datos/keyAEMET.key","r")
                    apikey = fich.read()
                    fich.close()
                    DescargarNuevosDatos(apikey)
                else:
                    print("Antes de ejecutar esta opción debe ejecutar la opción 1 y 11 (Crear base de datos e Introducir API key)")
                
            case 7:
                if ExisteBase():
                    DescargarDatosHora()
                else:
                    print("Antes de ejecutar esta opción debe ejecutar la opción 1 (Crear base de datos)")
                
            case 8:
                if ExisteBase() and ExisteApiKey():
                    fich = open("./datos/keyAEMET.key","r")
                    apikey = fich.read()
                    fich.close()
                    DatosEntreFechas(apikey)
                else:
                    print("Antes de ejecutar esta opción debe ejecutar la opción 1 y 11 (Crear base de datos e Introducir API key)")
                
            case 9:
                if ExisteBase():
                    DescargarDatosDias()
                else:
                    print("Antes de ejecutar esta opción debe ejecutar la opción 1 (Crear base de datos)")
                
            case 10:
                if ExisteBase() and ExisteApiKey():
                    MenuPrincipalApartado10()
                else:
                    print("Antes de ejecutar esta opción debe ejecutar la opción 1 y 11 (Crear base de datos e Introducir API key)")
            case 11:
                if ExisteApiKey():
                    seguir2 = True
                    while seguir2:
                        x = input("Ya existe una API key, ¿quieres volver a introducirla?\n 0 - No\n1 - Sí\n")
                        try:
                            x = int(x)
                            if x:
                                Guardar_API_key()
                                seguir2 = False
                            elif x == 0:
                                seguir2 = False
                            else:
                                print("Debes poner 0 o 1")
                        except:
                            print("Debes poner 0 o 1")
                else: 
                    Guardar_API_key()

# %%
if __name__ == '__main__':
    main()

# %%



