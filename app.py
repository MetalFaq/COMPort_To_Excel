"""
Requerimientos:
- Mostrar en forma de tabla el error por fase para cada punto de corriente.
ej, FAse 1: 70 mA- error; 5A- error
- Con la tabla anterior hacer graficos del error por fase en funcion de la corriente.

LP:   L1-- 220V 5A
LP:   L1-- 220V 0.04A
LP:   L-2- 220V 5A
LP:   L-2- 220V 0.04A
LP:   L--3 220V 5A
LP:   L--3 220V 0.04A
"""
import re
import os

from dic_to_list import dic_to_list
from write_excel import write_excel_scva, write_excel_norma

# ---------------------------Inicia comunicacion serie-------------------------------------------------------------------
# from com_serie import read_serial_port
# read_serial_port()
# _______________________________________________________________________________________________________________________

# file_text = abrir_archivo()
file_load = [] #contiene cada linea del archivo de datos.

def open_file():
    global file_load
    abs_path = os.path.dirname(__file__)
    relative_path = "/DATA"
    full_path = abs_path + relative_path
    filename = full_path + '/data_reac.txt' #data.txt
    with open(filename, 'r') as file:
        file_load = [lines.strip() for lines in file.readlines()]
    return file_load

# def get_file_name():  # definir como var global a file load.
#     # Ya lo tengo cargado al archivo con el content manager, sino en ésta funcion tengo que hacerlo de nuevo.
#     for line in file_load:  # Una vez cargado debo recorrelo.
#         expression = 'FName:(.*)'  # Regular expression "." letters, numbers,simbols, but not newlines.
#         matches = re.search(expression, line)
#
#         if matches:
#             # print(matches.group(0))
#             # print(matches.group(1))
#             file_name = str(matches.group(1))
#     # print(file_name)
#     return file_name
##########################################################
# def get_info():
#     for info in file_load:  # gruop(1): fase; group(2): tension, group(3): corriente.
#         expression = 'LP:   [A-Z]([0-9]+--|-[0-9]-|--[0-9]|[0-9]+)+ ([0-9]+)V+ ([0-9\.0-9]+)A+ phi=([+-0-9\.0-9]+)'
#         matches = re.search(expression, info)
#
#         if matches:
#             # print(matches.group(3))
#             # print(matches.group(1))
#             fase = str(matches.group(1))
#             tension = float(matches.group(2))
#             corriente = float(matches.group(3))
#             angulo = str(matches.group(4))
#             dic = {'fase': fase, 'tension': tension, 'phi': angulo, 'corriente': corriente, 'error': None}
#             dic_list.append(dic)
#
#     return dic_list
#
#
# def get_errors():
#     i = 0
#     for error in file_load:
#         expression = 'Mid:   ([+-0-9]+[.0-9]+)'
#         matches = re.search(expression, error)
#
#         if matches:
#             # print(matches.group(0))
#             # print(matches.group(1))
#             value_error = float(matches.group(1))
#             dic_list[i]['error'] = value_error
#             i += 1
#
#     return dic_list
####################################################################
# def get_name_for_scva():
#     for line in file_load:
#         expression = 'FName:([A0-Za9-z9]+)'
#         matches = re.search(expression, line)
#
#         if matches:
#             # print(matches.group(0))
#             # print(matches.group(1))
#             file_name_scva = str(matches.group(1))
#
#     # print(file_name_scva)
#     return file_name_scva

def get_tipo_num_serie():

    dic = {'tipo': '', 'num_serie': ''}
    for line in file_load:

        expression = 'Type: ([A-Za-z]+)'
        matches = re.search(expression, line)
        expression1 = 'SerNo:([0-9]+)'
        matches1 = re.search(expression1, line)

        if matches:
            dic['tipo'] = str(matches.group(1))
        if matches1:
            dic['num_serie'] = str(matches1.group(1))
            return dic

def data_from_txt():
    dic_list = []
    i = 0

    for info in file_load:
        expression = 'Mid:   ([+-0-9]+[.0-9]+)'
        matches = re.search(expression, info)
        expression1 = 'LP:   [A-Z]([0-9]+--|-[0-9]-|--[0-9]|[0-9]+)+ ([0-9]+)V+ ([0-9\.0-9]+)A+ phi=([+-0-9\.0-9]+)'
        matches1 = re.search(expression1, info)

        if matches1:
            fase = str(matches1.group(1))
            tension = float(matches1.group(2))
            corriente = float(matches1.group(3))
            angulo = str(matches1.group(4))

            dic = {'fase': fase, 'tension': tension, 'phi': angulo, 'corriente': corriente, 'error': None}
            dic_list.append(dic)

        if matches:
            value_error = float(matches.group(1))
            dic_list[i]['error'] = value_error
            i += 1

    return dic_list

#Abre el archivo de datos que contiene la info del MTE.
open_file()

print("Ingrese:" + '\n' + "1) Para excel SCVA:" + '\n' + "2) Para excel Ensayo por norma:")
input_num = input()

if input_num == '1':
    dic_to_list(data_from_txt())
    write_excel_scva(get_tipo_num_serie())

elif input_num == '2':
    write_excel_norma(data_from_txt(), get_tipo_num_serie())