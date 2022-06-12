#!/usr/bin/python
# -*- coding: latin-1 -*-
# CLASI GUSS 3.1
# EN LUGAR DE OBTENER LOS LISTADOS DE CADA TRABAJADORA DE UN ARCHIVO CSV, LO HARÁ DESDE UN XLSX
# PORQUE EN LOS ORDENADORES DEL DESPACHO, EXCEL NO ABRIA LOS CSV DE FORMA QUE SE PUDIESEN MODIFICAR CON FACILIDAD

# CLASI GUSS 3.2
# LA EXTENSION DEL ARCHIVO FINAL PASA A SER .XLS EN LUGAR DE .XLSX (LINEA 448)
# PORQUE LOS ORDENADORES DEL DESPACHO NO ABREN BIEN LA EXTENSION .XLSX
# SE ALEATORIZA EL ORDEN DE LA LISTA DE TRABAJADORAS QUE VAN A SER UBICADAS EN LAS SALAS
# QUE NO REQUIEREN PERMISOS ESPECIALES, PORQUE EN VERSIONES ANTERIORES A LA GENTE LE TOCABA A MENUDO EN LA MISMA SALA

# NO BORRAR LOS COMENTARIOS DE LAS FILAS 1 Y 2, SIRVEN PARA QUE EL ARCHIVO PUEDA SER DESCODIFICADO
import sys
import tkinter
import tkinter.filedialog
from tkinter import *
from tkinter.ttk import Combobox
import pandas as pd
import os
import xlwt
import random


#####FUNCIONES###
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


def listado_para_ese_turno_y_dia(dia, turno, dataframe):
    # CREA UN DATAFRAME QUE CONTIENE LOS NOMBRES DE QUIENES TRABAJAN ESE DIA Y TURNO, Y LO ORDENA ALEATORIAMENTE
    # COMO A MENUDO OCURRE QUE LA SEMANA ABARCA DOS MESES, REDIRIJO LOS POSIBLES ERRORES CON TRY/EXCEPT
    try:
        df_no_aleatorio = dataframe[dataframe[str(dia)] == turno].NOMBRE
        return list(df_no_aleatorio.sample(frac=1))
    except:
        return []


def listado_elegidos_trauma_o_gsuc(df, dia, turno, n_puestos):
    # ESTA FUNCION ESCOGE ALEATORIAMENTE TANTAS PERSONAS COMO UBICACIONES HAYA QUE CUBRIR PARA UN PUESTO,
    # EN EL CASO DE LOS PUESTOS QUE TIENEN PLANILLA PROPIA (EN ESTE MOMENTO, TRAUMA Y GSUC).
    # COMO A MENUDO OCURRE QUE LA SEMANA ABARCA DOS MESES, REDIRIJO LOS POSIBLES ERRORES CON TRY/EXCEPT
    try:
        df_dia_turno = df[df[dia] == turno]
        lista_dia_turno = list(df_dia_turno.NOMBRE)
        random.shuffle(lista_dia_turno)
        elegidas = []

        for i in range(0, n_puestos):
            if len(lista_dia_turno) > 0:
                elegidas.append(lista_dia_turno.pop())
        return elegidas, lista_dia_turno
    except:
        return [], []

def listado_elegidos_triage_o_rea(lista_trabajan, df_permisos, nombre_del_puesto, n_puestos):
    # ESTA FUNCION ESCOGE ALEATORIAMENTE TANTAS PERSONAS COMO UBICACIONES HAYA QUE CUBRIR PARA UN PUESTO,
    # EN EL CASO DE LOS PUESTOS PARA LOS QUE HAY QUE TENER PERMISO (EN ESTE MOMENTO, REA Y TRIAGE).

    # EN PRIMER LUGAR CREA UN DATAFRAME CON LA GENTE QUE TRABAJA ESE DIA Y TURNO. LO CREA PARTIENDO DE LA LISTA
    # DE GENTE QUE TRABAJAN ESE DIA Y TURNO, BUSCANDO SU NOMBRE EN EL DATAFRAME CON LOS PERMISOS
    df_trabajan = df_permisos.loc[df_permisos["NOMBRE"].isin(lista_trabajan)]
    # A PARTIR DEL DATAFRAME DE QUIENES TRABAJAN ESE DIA Y TURNO, ESCOGE LOS QUE TIENEN PERMISO PARA ESA UBICACION

    #REDACCIÓN ORIGINAL CON TRUE:
    df_pueden_ir_a_ese_puesto = df_trabajan[df_trabajan[nombre_del_puesto] == True]




    # SI EL DF RESULTANTE TIENE TANTA O MAS GENTE QUE EL NUMERO DE UBICACIONES QUE HACE FALTA CUBRIR,
    # ESCOGE UNA MUESTRA ALEATORIA
    if len(df_pueden_ir_a_ese_puesto) >= n_puestos:
        df_elegidas_para_el_puesto = df_pueden_ir_a_ese_puesto.sample(n_puestos)
    # SI EL DF NO TIENE MAS GENTE QUE EL NUMERO DE PUESTOS A CUBRIR, NO LO ALTERA
    else:
        df_elegidas_para_el_puesto = df_pueden_ir_a_ese_puesto
    # CONVIERTE EL DF EN UNA LISTA DE NOMBRES
    lista_elegidas_para_el_puesto = list(df_elegidas_para_el_puesto.NOMBRE)
    # ELIMINA ESOS NOMBRES DE LA LISTA DE PERSONAS QUE AUN NO TIENEN UBICACION, PARA QUE NO SEAN UBICADAS DOS VECES
    for nombre in lista_elegidas_para_el_puesto:
        lista_trabajan.remove(nombre)
    #ALEATORIZA LA LISTA
    random.shuffle(lista_elegidas_para_el_puesto)
    # DEVUELVE LA LISTA CON EL NUMERO DE NOMBRES NECESARIOS PARA CUBRIR LAS UBICACIONES VACANTES
    return lista_elegidas_para_el_puesto

def crear_y_grabar_dia(dia, turno, df_guss, df_trauma, df_gsuc, df_unida_con_permisos,
                       n_puestos_trauma,n_puestos_gsuc,n_puestos_triage,n_puestos_rea, n_puestos_consultas,
                       columna, ws, path_ubicaciones):


    #ESTAS VARIABLES EXTRAEN LAS UBICACIONES DEL ARCHIVO EXCEL EN EL QUE SE GRABAN.
    #ESTO PERMITE QUE LOS USUARIOS CAMBIEN EL NUMERO Y TIPO DE UBICACIONES
    #ESE EXCEL DE UBICACIONES TIENE UNA HOJA PARA CADA TURNO
    df_ubicaciones_M = pd.read_excel(path_ubicaciones, sheet_name=0)
    df_ubicaciones_T = pd.read_excel(path_ubicaciones, sheet_name=1)
    df_ubicaciones_N = pd.read_excel(path_ubicaciones, sheet_name=2)
    # CREO UN LISTADO DE TODAS LAS PERSONAS QUE FIGURAN EN LA PLANILLA GENERAL PARA TRABAJAR ESE DIA Y TURNO
    lista_trabajan_ese_dia_y_turno = listado_para_ese_turno_y_dia(dia, turno, df_guss)

    # ELIJO LAS PERSONAS QUE VAN A TRAUMA
    elegidas_trauma, sobrantes_trauma = \
        listado_elegidos_trauma_o_gsuc(df_trauma, dia, turno, n_puestos_trauma)
    # SI LE TOCA TRABAJAR A MAS PERSONAS DE LA PLANILLA DE TRAUMA QUE EL NUMERO DE PUESTOS A CUBRIR,
    # EL PERSONAL SOBRANTE PASA A LA LISTA GENERAL PARA LAS UBICACIONES QUE REQUIEREN PERMISO Y LAS GENERALES
    for nombre in sobrantes_trauma:
        lista_trabajan_ese_dia_y_turno.append(nombre)

    # ELIJO LAS PERSONAS PARA LA GSUC
    if categoria == 'TCAE':
        elegidas_gsuc, sobrantes_gsuc = [], []
    else:
        elegidas_gsuc, sobrantes_gsuc = \
            listado_elegidos_trauma_o_gsuc(df_gsuc, dia, turno, n_puestos_gsuc)

    # SI LE TOCA TRABAJAR A MAS PERSONAS DE LA PLANILLA DE LA GSUC QUE EL NUMERO DE PUESTOS A CUBRIR,
    # EL PERSONAL SOBRANTE PASA A LA LISTA GENERAL PARA LAS UBICACIONES QUE REQUIEREN PERMISO Y LAS GENERALES
    for nombre in sobrantes_gsuc:
        lista_trabajan_ese_dia_y_turno.append(nombre)

    # ELIJO LAS PERSONAS PARA EL TRIAGE (ENFERMERAS) O CONSULTAS (TCAES)
    if categoria == 'TCAE':
        lista_elegidas_triage= []
        lista_elegidas_consultas = listado_elegidos_triage_o_rea(
            lista_trabajan_ese_dia_y_turno, df_unida_con_permisos, "CONSULTAS", n_puestos_consultas)
    else:
        lista_elegidas_consultas = []
        lista_elegidas_triage = listado_elegidos_triage_o_rea(
            lista_trabajan_ese_dia_y_turno, df_unida_con_permisos, "MANCHESTER", n_puestos_triage)

    # ELIJO LAS PERSONAS PARA LA REA
    lista_elegidas_rea = listado_elegidos_triage_o_rea(
        lista_trabajan_ese_dia_y_turno, df_unida_con_permisos, "REA", n_puestos_rea)

    # EN EL CASO DE QUE HAYA UBICACIONES VACIAS EN TRAUMA O GSUC,
    # SE BUSCAN PERSONAS DE LA LISTA DE SOBRANTES QUE PUEDAN CUBRIR ESAS UBICACIONES
    if len(elegidas_trauma) < n_puestos_trauma:
        n_puestos_trauma_sin_cubrir = n_puestos_trauma - len(elegidas_trauma)
        lista_para_cubrir_huecos_trauma = listado_elegidos_triage_o_rea\
            (lista_trabajan_ese_dia_y_turno, df_unida_con_permisos, "TUSS", n_puestos_trauma_sin_cubrir)
        for suplente in lista_para_cubrir_huecos_trauma:
            elegidas_trauma.append(suplente)

    if len(elegidas_gsuc) < n_puestos_gsuc:
        n_puestos_gsuc_sin_cubrir = n_puestos_gsuc - len(elegidas_gsuc)
        lista_para_cubrir_huecos_gsuc = listado_elegidos_triage_o_rea\
            (lista_trabajan_ese_dia_y_turno, df_unida_con_permisos, "OSI", n_puestos_gsuc_sin_cubrir)
        for suplente in lista_para_cubrir_huecos_gsuc:
            elegidas_gsuc.append(suplente)





    # ELIJO LAS PERSONAS PARA EL DESLIZANTE
    try:
        df_deslizante = df_guss[df_guss[dia] == "D1"]
        lista_deslizante = list(df_deslizante.NOMBRE)
    except:
        lista_deslizante= []


    # GRABO LOS NOMBRES EN SUS UBICACIONES
    fila = 3
    #EL EXCEL DE LAS UBICACIONES TIENE UNA HOJA PARA CADA TURNO, QUE SON CONVERTIDAS EN UN DATAFRAME CADA UNA
    if turno == "M":
        tipos_ubicaciones = df_ubicaciones_M["TIPO DE UBICACIÓN"]
    if turno == "T":
        tipos_ubicaciones = df_ubicaciones_T["TIPO DE UBICACIÓN"]
    if turno == "N":
        tipos_ubicaciones = df_ubicaciones_N["TIPO DE UBICACIÓN"]
    #ALEATORIZO LA LISTA DE QUIENES TRABAJAN EN UBICACIONES GENERALISTAS. EN LAS PRIMERAS VERSIONES LA GENTE
    #REPETIA UBICACIONES A MENUDO PORQUE ESTA LISTA NO ESTABA ALEATORIZADA
    random.shuffle(lista_trabajan_ese_dia_y_turno)
    #PARA CADA UBICACION, EN FUNCION DEL TIPO QUE SEA (TRAUMA, REA, ETC)
    # BUSCA EN LA LISTA CORRESPONDIENTE Y EXTRAE UN NOMBRE QUE GRABA EN EL ARCHIVO
    for tipo in tipos_ubicaciones:
        if type(tipo) == str:

            if tipo == "TRAUMA":
                if len(elegidas_trauma)>0:
                    nombre=elegidas_trauma.pop(0)
                else: nombre=""
            if tipo == "GSUC":
                if len(elegidas_gsuc) > 0:
                    nombre=elegidas_gsuc.pop(0)
                else: nombre=""
            if tipo == "REA":
                if len(lista_elegidas_rea) > 0:
                    nombre=lista_elegidas_rea.pop(0)
                else: nombre=""
            if tipo == "TRIAGE":
                if len(lista_elegidas_triage) > 0:
                    nombre=lista_elegidas_triage.pop(0)
                else: nombre=""
            if tipo == "CONSULTAS":
                if len(lista_elegidas_consultas) > 0:
                    nombre=lista_elegidas_consultas.pop(0)
                else: nombre=""
            if tipo == "DESLIZANTE":
                if len(lista_deslizante)>0:
                    nombre=lista_deslizante.pop(0)
                else:
                    nombre = ""
            if tipo == "SIN PERMISOS ESPECIALES":
                if len(lista_trabajan_ese_dia_y_turno)>0:
                    nombre=lista_trabajan_ese_dia_y_turno.pop(0)
                else:
                    nombre = ""


            ws.write(fila, columna, nombre)

        else:
            ws.write(fila, columna, "")
        fila += 1
    #SI HAY MAS GENTE TRABAJANDO QUE UBICACIONES, EN EL ARCHIVO APARECE UNA LISTA CON SUS NOMBRES
    if len(lista_trabajan_ese_dia_y_turno)>0:
        fila += 1
        ws.write(fila, columna, "SIN UBICAR", fmt_bordes)
        fila += 1
        for nombre in lista_trabajan_ese_dia_y_turno:
            ws.write(fila, columna, nombre)
            fila += 1
    # SI HAY MAS PERSONAS QUE VIENEN A TRABAJAR EN HORARIO DE DESLIZANTE QUE UBICACIONES PREVISTAS,
    #SE CREA UNA LISTA CON SUS NOMBRES QUE APARECE EN EL ARCHIVO
    if len(lista_deslizante)>0 and turno != "N":
        fila += 1
        ws.write(fila, columna, "DESLIZANTES SIN UBICAR", fmt_bordes)
        fila += 1
        for nombre in lista_deslizante:
            ws.write(fila, columna, nombre)
            fila += 1




 # ESPECIFICO UN ANCHO DE COLUMNAS SUFICIENTE PARA QUE QUEPAN LOS NOMBRES
    columna_0 = ws.col(0)
    columna_1 = ws.col(1)
    columna_2 = ws.col(2)
    columna_3 = ws.col(3)
    columna_4 = ws.col(4)
    columna_5 = ws.col(5)
    columna_6 = ws.col(6)
    columna_7 = ws.col(7)
    columna_8 = ws.col(8)
    columna_9 = ws.col(9)
    columna_10 = ws.col(10)
    columna_11 = ws.col(11)
    columna_12 = ws.col(12)
    columna_13 = ws.col(13)
    columna_14 = ws.col(14)
    columna_15 = ws.col(15)
    columna_0.width = 256 * 40
    columna_1.width = 256 * 40
    columna_2.width = 256 * 40
    columna_3.width = 256 * 40
    columna_4.width = 256 * 40
    columna_5.width = 256 * 40
    columna_6.width = 256 * 40
    columna_7.width = 256 * 40
    columna_8.width = 256 * 40
    columna_9.width = 256 * 40
    columna_10.width = 256 * 40
    columna_11.width = 256 * 40
    columna_12.width = 256 * 40
    columna_13.width = 256 * 40
    columna_14.width = 256 * 40
    columna_15.width = 256 * 40

def n_puestos(df):
    #ESTA FUNCION UTILIZA UN DATAFRAME QUE CONTIENE LAS UBICACIONES PARA UN TURNO (M, T O N)
    #A PARTIR DE LA INFORMACION QUE CONTIENE EL DATAFRAME CALCULA CUANTOS PUESTOS HAY DE CADA TIPO PARA ESE TURNO
    n_puestos_triage = len(df[df["TIPO DE UBICACIÓN"] == "TRIAGE"])
    n_puestos_consultas = len(df[df["TIPO DE UBICACIÓN"] == "CONSULTAS"])
    n_puestos_trauma = len(df[df["TIPO DE UBICACIÓN"] == "TRAUMA"])
    n_puestos_gsuc = len(df[df["TIPO DE UBICACIÓN"] == "GSUC"])
    n_puestos_rea = len(df[df["TIPO DE UBICACIÓN"] == "REA"])
    n_puestos_deslizante = len(df[df["TIPO DE UBICACIÓN"] == "DESLIZANTE"])
    n_puestos_resto = len(df[df["TIPO DE UBICACIÓN"] == "SIN PERMISOS ESPECIALES"])
    return n_puestos_triage, n_puestos_consultas, n_puestos_trauma, n_puestos_gsuc, \
           n_puestos_rea, n_puestos_deslizante, n_puestos_resto


####VARIABLES####
categoria=""


# CREO UN EXCEL LLAMADO WB ON UNA HOJA DENTRO LLAMADA WS
wb = xlwt.Workbook()


# FORMATO (COLOR DE FONDO) DE LAS CELDAS DEL EXCEL
fmt_trauma = xlwt.Style.easyxf("""pattern: pattern solid, fore_colour green""")
fmt_triage = xlwt.Style.easyxf("""pattern: pattern solid, fore_colour blue""")
fmt_rea = xlwt.Style.easyxf("""pattern: pattern solid, fore_colour yellow""")
fmt_gsuc = xlwt.Style.easyxf("""pattern: pattern solid, fore_colour pink""")
fmt_resto = xlwt.Style.easyxf("""pattern: pattern solid, fore_colour orange""")
fmt_deslizante = xlwt.Style.easyxf("""pattern: pattern solid, fore_colour brown""")
fmt_bordes = xlwt.Style.easyxf("""pattern: pattern solid, fore_colour gray25""")


# AL APRETAR EL BOTÓN NEGRO "CREAR TABLA DE CLASIFICACIÓN", LLAMA A ESTA FUNCIÓN QUE PONE EN MARCHA TODO EL PROCESO
def crear_clasi():
    turnos=["M", "T", "N"]
    #path_ubicaciones="ubicaciones.xlsx"

    df_ubicaciones_M = pd.read_excel(path_ubicaciones, sheet_name=0)
    df_ubicaciones_T = pd.read_excel(path_ubicaciones, sheet_name=1)
    df_ubicaciones_N = pd.read_excel(path_ubicaciones, sheet_name=2)

    for turno in turnos:
        if turno == "M":
            n_puestos_triage, n_puestos_consultas, n_puestos_trauma, n_puestos_gsuc, n_puestos_rea,\
            n_puestos_deslizante, n_puestos_resto = n_puestos(df_ubicaciones_M)
            nombre_turno="MAÑANA"
        if turno == "T":
            n_puestos_triage, n_puestos_consultas, n_puestos_trauma, n_puestos_gsuc, n_puestos_rea, \
            n_puestos_deslizante, n_puestos_resto = n_puestos(df_ubicaciones_T)
            nombre_turno="TARDE"
        if turno == "N":
            n_puestos_triage, n_puestos_consultas, n_puestos_trauma, n_puestos_gsuc, n_puestos_rea, \
            n_puestos_deslizante, n_puestos_resto = n_puestos(df_ubicaciones_N)
            nombre_turno = "NOCHE"
        ws = wb.add_sheet(f'Clasificación{turno}', cell_overwrite_ok=True)
        primer_dia = int(combo_dias.get())
        semana = []
        ws.write(1, 0, f"{nombre_turno} {categoria.upper()}")
        ws.write(2, 0, f"Semana desde el {primer_dia}")

        for i in range(int(primer_dia), int(primer_dia) + 7):
            semana.append(str(i))
        nombre_dias_semana = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']


        df_guss = extraer_planilla_de_archivo_original(planilla_guss_path)
        df_trauma = extraer_planilla_de_archivo_original(planilla_trauma_path)
        if categoria == "Enfermeras":
            df_gsuc = extraer_planilla_de_archivo_original(planilla_gsuc_path)
        else:
            df_gsuc = pd.DataFrame()

        #OBTENER PERMISOS DESDE CSV: df_permisos = pd.read_csv(permisos_path, encoding="latin_1")
        #OBTENER PERMISOS DESDE EXCEL:
        df_permisos = pd.read_excel(permisos_path)
       # df_permisos = df_permisos.replace(1, True)
        # df_unida_con_permisos = pd.merge(df_guss, df_trauma, how="outer").merge(df_gsuc, how="outer").merge(df_permisos, how="left")
        df_unida_con_permisos = pd.concat([df_guss, df_trauma, df_gsuc]).merge(df_permisos, how="left", on="NOMBRE")

        # LAS PERSONAS QUE TRABAJAN DE MAÑANA O TARDE A VECES NO APARECEN EN LA PLANILLA CON M O T,
        # SINO CON VALORES COMO MC O TC (MAÑANA COMPLETA O TARDE COMPLETA, CUANDO TIENEN REDUCCION DE JORNADA)
        # MODIFICO ESOS VALORES PARA SIMPLIFICAR SU DETECCION EN EL PROGRAMA
        df_trauma = df_trauma.replace({"MC": "M", "Me": "M", "TC": "T"})
        df_guss = df_guss.replace({"MC": "M", "Me": "M", "TC": "T"})
        df_gsuc = df_gsuc.replace({"MC": "M", "Me": "M", "TC": "T"})
        df_unida_con_permisos = df_unida_con_permisos.replace({"MC": "M", "Me": "M", "TC": "T"})


        # GRABO LOS DIAS (POR EJEMPLO 13,14,15,16,17,18,19) COMO CABECEROS DE COLUMNA EN LA SEGUNDA FILA
        index = 1
        for d in semana:
            ws.write(1, index, d, fmt_bordes)
            index += 1
            ws.write(1, index, "", fmt_bordes)
            index += 1
        index = 1
        # GRABO LOS DIAS DE LA SEMANA (LUNES, MARTES, ETC.) EN LA TERCERA FILA
        for d in nombre_dias_semana:
            ws.write(2, index, d, fmt_bordes)
            index += 1
            ws.write(2, index, "", fmt_bordes)
            index += 1
        # GRABO LOS NOMBRES DE LOS PUESTOS (TRAUMA, TRIAGE, REA, GSUC) EN LA PRIMERA COLUMNA
        fila_puesto = 3
        #CADA TURNO TIENE UNAS UBICACIONES DISTINTAS,
        # QUE SE ENCUENTRAN EN CADA UNA DE LAS TRES HOJAS DEL EXCEL DE UBICACIONES
        if turno == "M":
            nombres_ubicaciones=df_ubicaciones_M["NOMBRE DE LA UBICACIÓN"]
        if turno == "T":
            nombres_ubicaciones=df_ubicaciones_T["NOMBRE DE LA UBICACIÓN"]
        if turno == "N":
            nombres_ubicaciones = df_ubicaciones_N["NOMBRE DE LA UBICACIÓN"]
        #EN LA PRIMERA COLUMNA GRABO EL NOMBRE DE CADA UBICACION
        for nombre in nombres_ubicaciones:
        #LA CONDICION DE QUE SEA UN STRING SIRVE PARA NO PROCESAR LAS FILAS EN BLANCO, CUYO VALOR NO ES STRING
            if type(nombre) == str:
                ws.write(fila_puesto, 0, nombre, fmt_bordes)
            else:
                ws.write(fila_puesto, 0, "")

            fila_puesto += 1
        fila_puesto += 1




        lbl_hecho = Label(window,
                          text=f'Tabla creada para {turnos} '
                               f'desde el día {primer_dia}',
                          bg="black", fg="white")
        lbl_hecho.grid(column=1, row=0)
        counter = 1
        for dia in semana:

            crear_y_grabar_dia(dia, turno, df_guss, df_trauma, df_gsuc, df_unida_con_permisos,
                               n_puestos_trauma, n_puestos_gsuc, n_puestos_triage, n_puestos_rea, n_puestos_consultas,
                               counter, ws, path_ubicaciones)
            counter += 2



#HE CAMBIADO LA EXTENSION EH QUE SE GUARDA EL ARCHIVO DE .XLSX A .XLS PORQUE EN EL DESPACHO SUS ORDENADORES
#NO ABRIAN EL ARCHIVO Y TENIAN QUE CAMBIAR LA EXTENSION MANUALMENTE A .XLS
    wb.save(f'CLASI {categoria} M,T y N desde el día {primer_dia}.xls')



# Para elegir el archivo con la planilla
def examin_plani():
    global planilla_guss_path
    planilla_guss_path = tkinter.filedialog.askopenfilename(title=
                                                            'Seleccione el archivo con la planilla mensual general')
    return planilla_guss_path


def elige_enfermera():
    global categoria
    categoria = 'Enfermeras'
    """
    lbl_triage.grid()
    spin_triage.grid()
    """

def elige_TCAE():
    global categoria
    categoria = 'TCAE'
    """
    lbl_triage.grid_remove()
    spin_triage.grid_remove()
    """

# Para elegir el archivo con la planilla de TRAUMA
def examin_trauma():
    global planilla_trauma_path
    planilla_trauma_path = tkinter.filedialog.askopenfilename(title=
                                                              'Seleccione el archivo con la planilla mensual de TRAUMA')
    return planilla_trauma_path


# Para elegir el archivo con la planilla de OSI
def examin_osi():
    global planilla_gsuc_path
    planilla_gsuc_path = tkinter.filedialog.askopenfilename(title=
                                                            'Seleccione el archivo con la planilla mensual de GSUC')
    return planilla_gsuc_path


# Para elegir el archivo con los permisos
def examin_permisos():
    global permisos_path
    permisos_path = tkinter.filedialog.askopenfilename(
        title='Seleccione el archivo con los permisos para trauma,triage,etc.'
    )
    return permisos_path

def examin_ubicaciones():
    global path_ubicaciones
    path_ubicaciones = tkinter.filedialog.askopenfilename(title=
                                                              'Seleccione el archivo con las ubicaciones')
    return path_ubicaciones

# Cierra la GUI y el programa
def salir():
    sys.exit()


window = Tk()
window.geometry('600x400')
window.title("Clasi GUSS")

# Mensaje que aparece al principio
lbl_inicial = Label(window, text="Bienvenid@ a ClasiGUSS")
lbl_inicial.grid(column=1, row=0)

# Botón para generar la clasi
btn_hacer_clasi = Button(window, text="Crear tabla de clasificación", command=crear_clasi,
                         bg="black", fg="white")
btn_hacer_clasi.grid(column=1, row=13)

# Botón para terminar el programa
btn_plani = Button(window, text="Salir", command=salir)
btn_plani.grid(column=1, row=14)

# Desplegable para elegir el primer día (lunes) de la tabla de la clasi
combo_dias = Combobox(window, width=3)
combo_dias['values'] = (
    "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14",
    "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31")
combo_dias.grid(column=2, row=3)
combo_dias.current(0)

# Texto que acompaña al desplegable para elegir el primer día
lbl_dias = Label(window, text='Seleccione en qué día del mes cae el lunes de la semana:')
lbl_dias.grid(column=1, row=3)

# Texto que acompaña al botón para elegir el archivo con la planilla general
lbl_plani = Label(window, text='Seleccione el archivo con la planilla GENERAL:')
lbl_plani.grid(column=1, row=4)

# Botón para elegir el archivo con la planilla
btn_plani = Button(window, text="Examinar", command=examin_plani)
btn_plani.grid(column=2, row=4)

# Texto que acompaña al botón para elegir el archivo con la planilla de trauma
lbl_plani = Label(window, text='Seleccione el archivo con la planilla de TRAUMA:')
lbl_plani.grid(column=1, row=5)

# Botón para elegir el archivo con la planilla de trauma
btn_plani = Button(window, text="Examinar", command=examin_trauma)
btn_plani.grid(column=2, row=5)

# Texto que acompaña al botón para elegir el archivo con la planilla de OSI
lbl_plani = Label(window, text='Seleccione el archivo con la planilla de GSUC:')
lbl_plani.grid(column=1, row=6)

# Botón para elegir el archivo con la planilla de osi
btn_plani = Button(window, text="Examinar", command=examin_osi)
btn_plani.grid(column=2, row=6)

# Texto que acompaña al botón para elegir el archivo con los permisos
lbl_permisos = Label(window, text='Seleccione el archivo con los permisos:')
lbl_permisos.grid(column=1, row=8)

# Botón para elegir el archivo con los permisos
btn_permisos = Button(window, text="Examinar", command=examin_permisos)
btn_permisos.grid(column=2, row=8)

# Texto que acompaña al botón para elegir el archivo con las ubicaciones
lbl_permisos = Label(window, text='Seleccione el archivo con las ubicaciones:')
lbl_permisos.grid(column=1, row=9)

# Botón para elegir el archivo con las ubicaciones
btn_permisos = Button(window, text="Examinar", command=examin_ubicaciones)
btn_permisos.grid(column=2, row=9)

# Botones para elegir la categoría
btn_enf = Button(window, text="Enfermeras", command=elige_enfermera)
btn_enf.grid(column=1, row=1)
btn_tcae = Button(window, text="TCAE", command=elige_TCAE)
btn_tcae.grid(column=2, row=1)
# Label que acompaña los botones para elegir categoría
lbl_cat = Label(window, text='Categoría: ')
lbl_cat.grid(column=0, row=1)

"""
# SpinBox para elegir el número de puestos de rea
spin_rea = Spinbox(window, from_=0, to=10, width=3)
spin_rea.grid(column=2, row=9)

# Texto que acompaña al spinbox de rea
lbl_rea = Label(window, text='Número de puestos de REANIMACIÓN')
lbl_rea.grid(column=1, row=9)

# SpinBox para elegir el número de puestos de trauma
spin_trauma = Spinbox(window, from_=0, to=10, width=3)
spin_trauma.grid(column=2, row=10)

# Texto que acompaña al spinbox de trauma
lbl_trauma = Label(window, text='Número de puestos de TRAUMATOLOGÍA')
lbl_trauma.grid(column=1, row=10)

# SpinBox para elegir el número de puestos de triage
spin_triage = Spinbox(window, from_=0, to=10, width=3)
spin_triage.grid(column=2, row=11)

# Texto que acompaña al spinbox de triage
lbl_triage = Label(window, text='Número de puestos de TRIAGE')
lbl_triage.grid(column=1, row=11)

# SpinBox para elegir el número de puestos de OSI
spin_osi = Spinbox(window, from_=0, to=10, width=3)
spin_osi.grid(column=2, row=12)

# Texto que acompaña al spinbox de OSI
lbl_osi = Label(window, text='Número de puestos de OSI')
lbl_osi.grid(column=1, row=12)
"""
window.mainloop()
