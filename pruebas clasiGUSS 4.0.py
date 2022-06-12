import pandas as pd
import os
import random


def extraer_planilla_de_archivo_original(path_a_la_planilla):
    # ESTA FUNCIÓN ABRE EL ARCHIVO ORIGINAL,
    # EXTRAE LA INFORMACIÓN QUE SE VA A USAR (NOMBRE, NUMERO FUNCIONAL, DIAS QUE TRABAJA CADA CUAL)
    # Y DEVUELVE UN DATAFRAME QUE CONTIENE ESOS DATOS
    # HE AÑADIDO UN TRY / ELSE PARA QUE EL PROGRAMA SIGA FUNCIONANDO AUNQUE NO SE SELECCIONE UNA RUTA DE ARCHIVO
    # EN LA GUI
    try:
        contenido_crudo = open(path_a_la_planilla, encoding="latin_1")
        lineas = contenido_crudo.readlines()
        with open("archivo temporal de la función extraer_planilla_de_archivo_original.csv", "w",
                  encoding="latin_1") as salida:
            for linea in lineas:
                if "Dias" in linea or "," in linea:
                    salida.write(linea)
        salida.close()
        df = pd.read_csv("archivo temporal de la función extraer_planilla_de_archivo_original.csv",
                         encoding="latin_1", delimiter="\t", header=0)
        del df["Unnamed: 0"]
        df.rename(columns={
            'Dias': 'NOMBRE',
            "Unnamed: 2": "N_FUNCIONAL"
        }, inplace=True)

        if os.path.exists("archivo temporal de la función extraer_planilla_de_archivo_original.csv"):
            os.remove("archivo temporal de la función extraer_planilla_de_archivo_original.csv")
        return df
    except FileNotFoundError:
        df_vacio = pd.DataFrame()
        return df_vacio
    else:
        df_vacio = pd.DataFrame()
        return df_vacio

path_ubicaciones = "/home/portatil/Escritorio/programación/clasiGUSS/ClasiGUSS 3.3/ubiDUEpruebas.xlsx"
df_ubicaciones_M = pd.read_excel(path_ubicaciones, sheet_name=0)
df_ubicaciones_T = pd.read_excel(path_ubicaciones, sheet_name=1)
df_ubicaciones_N = pd.read_excel(path_ubicaciones, sheet_name=2)
path_permisos= "/home/portatil/Escritorio/programación/clasiGUSS/ClasiGUSS 3.3/Permisos DUE 4.0.xlsx"
df_permisos=pd.read_excel(path_permisos)
path_planilla = "/home/portatil/Escritorio/programación/clasiGUSS/ClasiGUSS 3.3/GUSS DIC DUE.xls"
path_planilla_trauma="/home/portatil/Escritorio/programación/clasiGUSS/ClasiGUSS 3.3/TUSS DIC DUE.xls"
path_planilla_osi="/home/portatil/Escritorio/programación/clasiGUSS/ClasiGUSS 3.3/OSI DIC DUE.xls"
df_planilla = extraer_planilla_de_archivo_original(path_planilla)
df_planilla_trauma= extraer_planilla_de_archivo_original(path_planilla_trauma)
df_planilla_osi= extraer_planilla_de_archivo_original(path_planilla_osi)

# AÑADO EN LAS PLANILLAS DE SALAS ESPECIALES UNA COLUMNA QUE SEÑALE QUE ESTAS PERSONAS ESTÁN ROTANDO POR ESAS SALAS
# ASÍ QUEDAN SEÑALADAS AL FUSIONAR ESTOS DF CON EL DF GENERAL
df_planilla_osi["PLANI_OSI"]=True
df_planilla_trauma["PLANI_TRAUMA"]=True

# CREO UN DF CON LAS TRES PLANILLAS UNIDAS
df_plani_unida=pd.concat([df_planilla, df_planilla_trauma, df_planilla_osi])

# EL DF DE PERMISOS NO CONVIERTE BIEN LOS VALORES BOOLEANOS AL CARGAR DESDE EXCEL, LOS VUELVO A FIJAR
df_permisos_2=df_permisos.replace([1,0], [True, False])

# ANTES DE UNIR LOS PERMISOS AL DF DE LAS TRES PLANILLAS ELIMINO LA COLUMNA DE NOMBRES PARA QUE NO DE PROBLEMAS
# EN EL FUTURO
df_permisos_3=df_permisos_2.drop(columns=['NOMBRE'])
# UNO LOS PERMISOS AL DF DE LAS TRES PLANILLAS
df_plani_unida_con_permisos=df_plani_unida.join(df_permisos_3.set_index('N_FUNCIONAL'),
                                                on='N_FUNCIONAL',lsuffix='_left', rsuffix='_right')

# GRABO EL DF UNIFICADO CON LAS TRES PLANILLAS Y LOS PERMISOS EN UN EXCEL
df_plani_unida_con_permisos.to_excel("output.xlsx")


# PARA CADA DIA Y TURNO, EXTRAIGO UN DF DE QUIENES TRABAJAN ESE DÍA Y TURNO
dia= "21"
turno="T"
df_trabajan_ese_dia_y_turno=df_plani_unida_con_permisos[df_plani_unida_con_permisos[dia]== turno]

# A PARTIR DEL DF DE QUIENES TRABAJAN ESE DIA Y TURNO, CREO VARIOS: 1) ROTANTES POR TRAUMA 2) ROTANTES POR GSUC
# 3) CON PERMISO TRAUMA 4) CON PERMISO GSUC, 5) CON PERMISO REA, 6) CON PERMISO TRIAGE
# SALVO LOS DF DE ROTANTES, EL RESTO DEBERÍAN GENERARSE AUTOMÁTICAMENTE A PARTIR DE LOS TIPOS DE UBICACIONES
# QUE APARECEN EN EL EXCEL DE UBICACIONES.
df_rotan_trauma_ese_dia_y_turno=df_trabajan_ese_dia_y_turno[df_trabajan_ese_dia_y_turno.PLANI_TRAUMA == True]
df_rotan_osi_ese_dia_y_turno=df_trabajan_ese_dia_y_turno[df_trabajan_ese_dia_y_turno.PLANI_OSI == True]






df_ubicaciones_M = pd.read_excel(path_ubicaciones, sheet_name=0)
lista_ubicaciones=[]


"""
CÓMO NAVEGAR POR EL DATAFRAME EN FUNCIÓN DEL DÍA, TURNO Y PERMISO
dia= "21"
tipo_ubi= "TUSS"
turno="T"
dfa=df_plani_unida_con_permisos[(df_plani_unida_con_permisos[dia]== turno)
                                & (df_plani_unida_con_permisos[tipo_ubi]== True)]
"""
def obtener_listados(dia, turno, df_plani_unida_con_permisos, path_ubicaciones):
    if turno == "M":
        df_ubicaciones = pd.read_excel(path_ubicaciones, sheet_name=0)
    if turno == "T":
        df_ubicaciones = pd.read_excel(path_ubicaciones, sheet_name=1)
    if turno == "N":
        df_ubicaciones = pd.read_excel(path_ubicaciones, sheet_name=2)

    lista_ubicaciones = []
    dicc_ubicaciones = {}
    for i in df_ubicaciones.index:
        ubicacion = df_ubicaciones["TIPO DE UBICACIÓN"][i]
        if type(ubicacion) == str:
            if ubicacion not in lista_ubicaciones:
                lista_ubicaciones.append(ubicacion)


    df_trabajan_ese_dia_y_turno = df_plani_unida_con_permisos[df_plani_unida_con_permisos[dia] == turno]
    df_rotan_trauma_ese_dia_y_turno = df_trabajan_ese_dia_y_turno[df_trabajan_ese_dia_y_turno.PLANI_TRAUMA == True]
    df_rotan_osi_ese_dia_y_turno = df_trabajan_ese_dia_y_turno[df_trabajan_ese_dia_y_turno.PLANI_OSI == True]



# CREO UN DICCIONARIO DE DF. CON CADA NOMBRE DE UBICACIÓN SE CREA UN DF DE QUIENES TIENEN PERMISO PARA ESA UBICACIÓN
    for tipo_ubicacion in lista_ubicaciones:
        try:
            dicc_ubicaciones[tipo_ubicacion]= df_trabajan_ese_dia_y_turno[df_trabajan_ese_dia_y_turno[tipo_ubicacion] == True]
        except:
            pass
# LA FUNCIÓN DEVUELVE LOS DF DE QUIENES TRABAJAN ESE DÍA Y TURNO Y LOS SUBCONJUNTOS EN FUNCIÓN DE LOS ROTANTES
# Y PERMISOS
    
    return df_trabajan_ese_dia_y_turno, df_rotan_trauma_ese_dia_y_turno, df_rotan_osi_ese_dia_y_turno, dicc_ubicaciones



obtener_listados(dia, turno, df_plani_unida_con_permisos,path_ubicaciones)