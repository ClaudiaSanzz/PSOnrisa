# Importamos las librerias que vamos a usar
import requests                  # para descargar las paginas web
import pandas as pd              # para leer el archivo del HU-1
from bs4 import BeautifulSoup    # para leer el HTML de la pagina


# ----------------------------------------------------------
# Funcion 1: descargar la pagina de la ficha tecnica
# ----------------------------------------------------------
def descargar_pagina(nregistro):

    # Construimos la URL de la ficha tecnica con el numero de registro
    url = "https://cima.aemps.es/cima/dochtml/ft/" + str(nregistro) + "/FT_" + str(nregistro) + ".html"

    # Descargamos la pagina
    respuesta = requests.get(url)
    respuesta.encoding = "utf-8"
    html = respuesta.text

    # Convertimos el HTML en un objeto que podemos recorrer
    soup = BeautifulSoup(html, "html.parser")

    return soup


# ----------------------------------------------------------
# Funcion 2: Volumen Informativo de Seguridad
# Cuenta las palabras de la seccion 4.4
# ----------------------------------------------------------
def contar_palabras(soup):

    inicio = soup.find(id="4.4")   # aqui empieza la seccion 4.4
    fin = soup.find(id="4.5")      # aqui empieza la siguiente (4.5)

    # Si la ficha no tiene seccion 4.4 devolvemos None
    if inicio is None:
        return None

    contador_palabras = 0

    # Recorremos todos los elementos que hay despues de la 4.4
    for elemento in inicio.find_all_next():

        # Si llegamos a la 4.5 paramos, porque ya no es la 4.4
        if elemento == fin:
            break

        # Si el elemento es un parrafo <p>, contamos sus palabras
        if elemento.name == "p":
            texto = elemento.get_text()
            palabras = texto.split()          # separamos el texto en palabras
            contador_palabras = contador_palabras + len(palabras)

    return contador_palabras


# ----------------------------------------------------------
# Funcion 3: contar las tablas de la ficha tecnica
# ----------------------------------------------------------
def contar_tablas(soup):

    # Buscamos todas las etiquetas <table> de la pagina
    tablas = soup.find_all("table")
    numero_tablas = len(tablas)

    return numero_tablas


# ----------------------------------------------------------
# Funcion 4: Indicador de Riesgo Severo
# Cuenta cuantas veces aparece "grave" o "graves" en todo el texto
# ----------------------------------------------------------
def contar_grave(soup):

    # Sacamos todo el texto de la ficha tecnica y lo pasamos a minusculas
    texto = soup.get_text().lower()

    # Separamos el texto en palabras
    palabras = texto.split()

    contador_grave = 0

    # Recorremos las palabras una a una
    for palabra in palabras:

        # Si la palabra es "grave" o "graves" la contamos
        if palabra == "grave" or palabra == "graves":
            contador_grave = contador_grave + 1

    return contador_grave


# ----------------------------------------------------------
# Programa principal
# ----------------------------------------------------------

# Leemos el Excel con los medicamentos
# dtype=str hace que nregistro se lea como texto y no se pierdan los ceros de delante
df = pd.read_excel("../hu_1_extraccion/medicamentos_psoriasis.xlsx", dtype={"nregistro": str})

# Nos quedamos con la columna del numero de registro
lista_nregistros = df["nregistro"]

print("Numero de medicamentos a analizar:", len(lista_nregistros))

# Listas vacias donde iremos guardando los resultados de cada medicamento
lista_contador_palabras = []
lista_numero_tablas = []
lista_contador_grave = []


# Recorremos cada medicamento uno a uno
for nregistro in lista_nregistros:

    # Descargamos su pagina (una sola vez para todas las metricas)
    soup = descargar_pagina(nregistro)

    # Metrica 1: palabras de la seccion 4.4
    palabras = contar_palabras(soup)
    lista_contador_palabras.append(palabras)

    if palabras is None:
        print("No se ha encontrado la seccion 4.4 de", nregistro)
    else:
        print("Numero de palabras en la seccion 4.4. de", nregistro, "=", palabras)

    # Metrica 2: numero de tablas
    numero_tablas = contar_tablas(soup)
    lista_numero_tablas.append(numero_tablas)
    print("Numero de tablas en", nregistro, "=", numero_tablas)

    # Metrica 3: veces que aparece "grave" o "graves"
    graves = contar_grave(soup)
    lista_contador_grave.append(graves)
    print("Numero de veces que aparece 'grave' o 'graves' en", nregistro, "=", graves)

# Metemos las listas como columnas nuevas en la tabla
#df["palabras_seccion_4_4"] = lista_contador_palabras
#df["numero_tablas"] = lista_numero_tablas
#df["contador_grave"] = lista_contador_grave

# Guardamos el resultado en un Excel nuevo
#df.to_excel("medicamentos_con_metricas.xlsx", index=False)

#print("Analisis terminado. Resultados guardados en medicamentos_con_metricas.xlsx")