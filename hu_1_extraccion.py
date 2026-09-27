#Importamos las librerias
import requests #Necesaria para comunicarte con la API (realizar peticiones)
import math #Necesaria para calcular el número de páginas
import pandas as pd #Para crear DataFrames
import numpy as np #Para crear los valores NaN

#Nuestra enfermedad para la búsqueda
enfermedad = "psoriasis"

#Payload que se envía a la API para buscar la enfermedad
payload = [
    {
        "seccion": "4.1",
        "texto": enfermedad,
        "contiene": 1
    }
]


#
# 1) Función para obtener los nregistros de todas las páginas
#
def obtener_todas_las_paginas(payload):
    url = "https://cima.aemps.es/cima/rest/buscarEnFichaTecnica"
    #Primera petición para saber el número total de registros
    respuesta = requests.post(url + "?pagina=1", json=payload) #Haces una petición post
    datos = respuesta.json() #pyhton los convierte en estructuras que puede manejar, normalmente diccionarios y listas
    #Número total de medicamentos encontrados
    medicamentos_total = datos["totalFilas"]
    #Resultados (medicamentos) de la primera páginas
    medicamentos_1pag = datos["resultados"]
    #Calculamos el número total de páginas
    numero = medicamentos_total / len(medicamentos_1pag)
    paginas_total = math.ceil(numero) #Esto te da el numero entero por arriba si es decimal p.e de 1.2 pues 2


    nregistros = []
    #Recorremos todas las páginas que haya
    for pagina in range(1, paginas_total + 1):

        url_pagina = f"{url}?pagina={pagina}" #Va cambiando la página
        respuesta = requests.post(url_pagina, json=payload)
        datos = respuesta.json()
        resultados = datos["resultados"]
        #Mostramos cuantos medicamentos hay por cada pagina
        print("Página:", pagina)
        print("Medicamentos:", len(resultados))

        #Guardamos el nregistro de cada medicamento
        for medicamento in resultados:
            nregistro = medicamento["nregistro"]
            nregistros.append(nregistro)
    
    return nregistros 

nregistros = obtener_todas_las_paginas(payload)
print(nregistros, "Número total de nregristros ", len(nregistros))

#
# 2) Función para obtenet la información de un medicamento
#
def obtener_info_medicamento(nregistro):
    url_base = "https://cima.aemps.es/cima/rest/medicamento?nregistro="
    url_medicamento = url_base + str(nregistro) #Añadimos el nregistro del medicamento a la url
    respuesta = requests.get(url_medicamento)
    datos = respuesta.json()

    #EXTRACCION DE LA INFORMACION
    
    #CN
    presentaciones = datos.get("presentaciones", []) #Da todos los datos que haya dentro y si no existe te crea una lista vacía
    #FormaFarmaceuticaSimplificada
    formaFarmaceuticaSimplificada = datos.get("formaFarmaceuticaSimplificada", {})
    #Estados
    estados = datos.get("estado",{})
    #viasAdministracion
    viasAdministracion = datos.get("viasAdministracion", []) 
    #url HTML   
    docs = datos.get("docs",[])
    #fotos materiales
    fotos = datos.get("fotos", [])
    url_foto_material = np.nan
    for foto in fotos:
        tipo = foto.get("tipo", "")
        url_foto = foto.get("url", "")
        if "material" in str(tipo) and url_foto.endswith(".jpg"): 
            url_foto_material = url_foto
            break #Para quedarnos solo con la primera que encontremos



    medicamento = {
        "nregistro" : datos.get("nregistro"),
        "nombre" : datos.get("nombre"),
        "pactivos" : datos.get("pactivos"),
        "labtitular" : datos.get("labtitular"),
        "labcomercializador" : datos.get("labcomercializador"),
        "cn" : presentaciones[0].get("cn", np.nan) if presentaciones else np.nan,
        "dosis" : datos.get("dosis"),
        "forma_farmaceutica_simplificada" : formaFarmaceuticaSimplificada.get("nombre"),
        "estado_aut" : estados.get("aut"),
        "estado_rev" : estados.get("rev", np.nan), #Pone NaN cuando no exista
        "vias_administracion": ", ".join(v.get("nombre", "") for v in viasAdministracion) if viasAdministracion else np.nan, #Recorre la lista y va juntando los nombres de todas las vías separadas por comas y si esta vacía deja el np.nan
        "comerc" : 1 if datos.get("comerc") else 0 ,
        "requiere_receta" : 1 if datos.get("receta") else 0,
        "generico" : 1 if datos.get("generico") else 0,
        "afecta_conduccion" : 1 if datos.get("conduc") else 0,
        "triangulo_negro" : 1 if datos.get("triangulo") else 0, 
        "medicamento_huerfano" : 1 if datos.get("huerfano") else 0,
        "biosimmilar" :1 if datos.get("biosimilar") else 0,
        "url_html_ficha_tecnica" : docs[0].get("urlHtml", np.nan) if docs else np.nan,
        "url_foto_materiales" : url_foto_material,
        "num_registros_atc" : len(datos.get("atcs" , [])), #ponemos [] para que si no existe, no haya len(None) y no de error
        "num_principios_activos" : len(datos.get("principiosActivos", [])),
        #num_excipientes" No esta dentro de la informacion que da la API cuando haces un GET por lo que voy a poner que no la dan
        "num_excipientes" : "Dato no disponible"
    }
    return medicamento

#
# 3) Función para obtener la información de todos los medicamentos
#

def obtener_info_medicamentos(nregistros):
    medicamentos = []
    for nregistro in nregistros:
        medicamento = obtener_info_medicamento(nregistro)
        medicamentos.append(medicamento)
    return medicamentos


medicamentos = obtener_info_medicamentos(nregistros)
#
# 4) Creamos el DatFrame y lo exportamos a excel
#
df = pd.DataFrame(medicamentos)
df.to_excel("medicamentos_psoriasis.xlsx", index=False) #Guardamos los resultados en Excel
print("Medicamentos obtenidos:", len(medicamentos))



