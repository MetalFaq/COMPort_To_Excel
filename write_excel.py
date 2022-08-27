# import pandas as pd
# import matplotlib.pyplot as plt

import sys
import os
import xlwings as xw
from openpyxl import load_workbook
from datetime import datetime
from dic_to_list import list_e_fase1, list_e_fase2, list_e_fase3,  list_phi_f1

# df1 = []
# df2 = []
# df3 = []
# df4 = []

# def build_dataframe():
#
#     global df1, df2, df3, df4
#     df1 = pd.DataFrame({
#         'Corriente Fase 1': list_i_fase1,
#         'Error Fase 1': list_e_fase1,
#         'Phi': list_phi_f1
#     })
#     df2 = pd.DataFrame({
#         'Corriente Fase 2': list_i_fase2,
#         'Error Fase 2': list_e_fase2,
#         'Phi': list_phi_f2
#     })
#     df3 = pd.DataFrame({
#         'Corriente Fase 3': list_i_fase3,
#         'Error Fase 3': list_e_fase3,
#         'Phi': list_phi_f3
#     })
#     df4 = pd.DataFrame({
#         'Corriente Fase 123': list_i_fase123,
#         'Error Fase 123': list_e_fase123,
#         'Phi': list_phi_f123
#     })

# def write_to_excel_default(file_name):
#
#     name = file_name + '.xlsx'
#     data = pd.ExcelWriter(name, engine = 'xlsxwriter')
#     df1.to_excel(data, sheet_name='Default', index=False)
#     df2.to_excel(data, sheet_name='Default', index = False, startcol = 4)
#     df3.to_excel(data, sheet_name='Default', index = False, startcol = 8)
#     if not df4.empty:                                                     # Si el data frame no es vacio ( si es true)
#         df4.to_excel(data, sheet_name='Default', index=False, startcol=12)
#     else:                                                                 # Sie el df es false (esta vacio)
#         pass
#     data.save()

# def write_to_excel(file_name):
#
#     global df1, df2, df3, df4
#
#     # data_filter1 = df1.groupby(key).filter(lambda x: (x[key] == value).any())
#     # data_filter2 = df2.groupby(key).filter(lambda x: (x[key] == value).any())
#     # data_filter3 = df3.groupby(key).filter(lambda x: (x[key] == value).any())
#     # data_filter4 = df4.groupby(key).filter(lambda x: (x[key] == value).any())
#     # name = file_name +'_'+ key + value + '.xlsx'
#
#     data_filter1 = df1.groupby(by=["Phi", "Corriente Fase 1"]).mean()
#     data_filter2 = df2.groupby(by=["Phi", "Corriente Fase 2"]).mean()
#     data_filter3 = df3.groupby(by=["Phi", "Corriente Fase 3"]).mean()
#     data_filter4 = df4.groupby(by=["Phi", "Corriente Fase 123"]).mean()
#
#     # Escribe Excel con datos por defecto
#     name = file_name + '.xlsx'
#     data = pd.ExcelWriter(name, engine='xlsxwriter')
#     if not df1.empty:
#         df1.to_excel(data, sheet_name='Default', index=False)
#     if not df2.empty:                                                                                                   # Si es monofasico que no escriba los titulos de las columnas de otras fases.
#         df2.to_excel(data, sheet_name='Default', index=False, startcol=4)
#     if not df3.empty:
#         df3.to_excel(data, sheet_name='Default', index=False, startcol=8)
#     if not df4.empty:
#         df4.to_excel(data, sheet_name='Default', index=False, startcol=12)
#
#
#     # Ajusta el ancho de las columnas.
#     worksheet = data.sheets['Default']
#     worksheet.set_column('A:B', 15)
#     worksheet.set_column('E:F', 15)
#     worksheet.set_column('I:J', 15)
#     worksheet.set_column('M:N', 16)
#
#     #Escribe datos fitrados por angulo.
#     if not data_filter1.empty:
#         data_filter1.to_excel(data, sheet_name='Filtrado', index=True)                                                      # Se usa index true para poder escribir en el excel las columnas filtradas con groupby( Phi, Corriente Fase x).
#     if not data_filter2.empty:                                                                                                   # Si es monofasico no escribe, de lo contrario daba error.
#         data_filter2.to_excel(data, sheet_name='Filtrado', index=True, startcol=4)
#     if not data_filter3.empty:
#         data_filter3.to_excel(data, sheet_name='Filtrado', index=True, startcol=8)
#     if not data_filter4.empty:                                                                                                   # Si el data frame no es vacio ( si es true)
#         data_filter4.to_excel(data, sheet_name='Filtrado', index=True, startcol=12)
#
#
#     # Obtiene los objetos xlsxwriter desde el dataframe
#     worksheet = data.sheets['Filtrado']
#     # workbook = data.book
#
#     # Ajusta el ancho de las columnas
#     worksheet.set_column('B:C', 15)
#     worksheet.set_column('F:G', 15)
#     worksheet.set_column('J:K', 15)
#     worksheet.set_column('N:O', 17)
#
#     # chart = workbook.add_chart({'type': 'line'})
#     # chart.add_series({'values': 'Filtrado'})
#
#     data.save()

abs_path = os.path.dirname(__file__)
# Path & name file
relative_path_file = "/resultados"
full_path_file = abs_path + relative_path_file
# Path & name template SCVA
relative_path_template_scva = "/Planillas/Planilla_SCVA.xlsx"
full_path_template = abs_path + relative_path_template_scva
# Path & name template normativa
relative_path_template_norma = "/Planillas/Planilla_Norma.xls"
full_path_template_norma = abs_path + relative_path_template_norma

def write_excel_scva(name_scva):
    name_file = full_path_file + '/SCVA_' + name_scva['num_serie'] + '.xlsx'

    if list_phi_f1[0] == '0': # Para identificar si es un archivo de activa.
        wb = load_workbook(filename=full_path_template)
        ws = wb.worksheets[0]
        ws['F4'] = datetime.now().date()
        ws['E14'] = name_scva['num_serie']
        ws['F14'] = list_e_fase1[0]
        ws['G14'] = list_e_fase1[1]
        ws['H14'] = list_e_fase1[2]
        ws['I14'] = list_e_fase1[3]
        ws['J14'] = list_e_fase1[4]
        ws['K14'] = list_e_fase1[5]                # F14, G14 y H14 carga alta fase 1. I14, J14 y K14 carga baja fase 1.

        if list_e_fase2:                                                       # Si pasa el If, es un medidor trifásico.
            ws['E56'] = name_scva['num_serie']
            ws['F56'] = list_e_fase2[0]
            ws['G56'] = list_e_fase2[1]
            ws['H56'] = list_e_fase2[2]
            ws['I56'] = list_e_fase2[3]
            ws['J56'] = list_e_fase2[4]
            ws['K56'] = list_e_fase2[5]

            ws['E98'] = name_scva['num_serie']
            ws['F98'] = list_e_fase3[0]
            ws['G98'] = list_e_fase3[1]
            ws['H98'] = list_e_fase3[2]
            ws['I98'] = list_e_fase3[3]
            ws['J98'] = list_e_fase3[4]
            ws['K98'] = list_e_fase3[5]
    else:                                   # por defecto escribe en la zona de reactiva. No verifica un PF especifico.
        try:
            wb = load_workbook(filename= name_file)
            ws = wb.worksheets[0]
            ws['L14'] = list_e_fase1[0]
            ws['M14'] = list_e_fase1[1]
            ws['N14'] = list_e_fase1[2]
            ws['O14'] = list_e_fase1[3]
            ws['P14'] = list_e_fase1[4]
            ws['Q14'] = list_e_fase1[5]

            ws['L56'] = list_e_fase2[0]
            ws['M56'] = list_e_fase2[1]
            ws['N56'] = list_e_fase2[2]
            ws['O56'] = list_e_fase2[3]
            ws['P56'] = list_e_fase2[4]
            ws['Q56'] = list_e_fase2[5]

            ws['L98'] = list_e_fase3[0]
            ws['M98'] = list_e_fase3[1]
            ws['N98'] = list_e_fase3[2]
            ws['O98'] = list_e_fase3[3]
            ws['P98'] = list_e_fase3[4]
            ws['Q98'] = list_e_fase3[5]
        except FileNotFoundError:
            print("Por favor, ingrese primero archivo de energia activa")
            sys.exit(1)

    wb.save(filename = name_file)

def write_excel_norma(data, type_num):
    # Problema! ---> Cuando el .txt cargado no tiene = longitud de datos especificada p/hacer el excel
    # no me genera el archivo.

    name_file = full_path_file + '/Planilla_Norma' + type_num['num_serie'] + '.xlsx'
    #name = 'G:/facundo/Escritorio/Herencia centurion/Parser_project_2/resultados/' + 'Planilla_Norma_' + type_num['num_serie'] + '.xlsx'

    if data[0]['phi'] == '0':  # Si es de activa, entra al if.
        wb = xw.Book(full_path_template_norma)
        sheet = wb.sheets['Hoja1']
        #Datos
        sheet.range('J6:N6').value = datetime.now().date()
        sheet.range('J9:K9').value = type_num['tipo']
        sheet.range('M9:N9').value = type_num['num_serie']
        #Valores de errores p/ cargas activas.
        sheet.range('I50:Q50').value = data[0]['error']
        sheet.range('I77:Q77').value = data[1]['error']
        sheet.range('I64:Q64').value = data[2]['error']
        sheet.range('I23:K23').value = data[3]['error']
        sheet.range('I36:K36').value = data[4]['error']
        sheet.range('L23:N23').value = data[5]['error']
        sheet.range('L36:N36').value = data[6]['error']
        sheet.range('O23:Q23').value = data[7]['error']
        sheet.range('O36:Q36').value = data[8]['error']
        sheet.range('I51:Q51').value = data[9]['error']
        sheet.range('I78:Q78').value = data[10]['error']
        sheet.range('I65:Q65').value = data[11]['error']
        sheet.range('I24:K24').value = data[12]['error']
        sheet.range('I37:K37').value = data[13]['error']
        sheet.range('L24:N24').value = data[14]['error']
        sheet.range('L37:N37').value = data[15]['error']
        sheet.range('O24:Q24').value = data[16]['error']
        sheet.range('O37:Q37').value = data[17]['error']
        sheet.range('I52:Q52').value = data[18]['error']
        sheet.range('I79:Q79').value = data[19]['error']
        sheet.range('I66:Q66').value = data[20]['error']
        sheet.range('I25:K25').value = data[21]['error']
        sheet.range('I38:K38').value = data[22]['error']
        sheet.range('L25:N25').value = data[23]['error']
        sheet.range('L38:N38').value = data[24]['error']
        sheet.range('O25:Q25').value = data[25]['error']
        sheet.range('O38:Q38').value = data[26]['error']
        sheet.range('I54:Q54').value = data[27]['error']
        sheet.range('I81:Q81').value = data[28]['error']
        sheet.range('I68:Q68').value = data[29]['error']
        sheet.range('I27:K27').value = data[30]['error']
        sheet.range('I40:K40').value = data[31]['error']
        sheet.range('L27:N27').value = data[32]['error']
        sheet.range('L40:N40').value = data[33]['error']
        sheet.range('O27:Q27').value = data[34]['error']
        sheet.range('O40:Q40').value = data[35]['error']
        sheet.range('I55:Q55').value = data[36]['error']
        sheet.range('I82:Q82').value = data[37]['error']
        sheet.range('I69:Q69').value = data[38]['error']
        sheet.range('I28:K28').value = data[39]['error']
        sheet.range('L28:N28').value = data[40]['error']
        sheet.range('O28:Q28').value = data[41]['error']
        sheet.range('I56:Q56').value = data[42]['error']
    else:
        try:
            wb = xw.Book(name_file)
            #filas de valores de error p/ carga reactiva
            sheet = wb.sheets['Hoja1']
            sheet.range('I111:Q111').value = data[0]['error']
            sheet.range('I131:Q131').value = data[1]['error']
            sheet.range('I122:Q122').value = data[2]['error']
            sheet.range('I90:K90').value = data[3]['error']
            sheet.range('I100:K100').value = data[4]['error']
            sheet.range('L90:N90').value = data[5]['error']
            sheet.range('L100:N100').value = data[6]['error']
            sheet.range('O90:Q90').value = data[7]['error']
            sheet.range('O100:Q100').value = data[8]['error']
            sheet.range('I112:Q112').value = data[9]['error']
            sheet.range('I132:Q132').value = data[10]['error']
            sheet.range('I123:Q123').value = data[11]['error']
            sheet.range('I91:K91').value = data[12]['error']
            sheet.range('I101:K101').value = data[13]['error']
            sheet.range('L91:N91').value = data[14]['error']
            sheet.range('L101:N101').value = data[15]['error']
            sheet.range('O91:Q91').value = data[16]['error']
            sheet.range('O101:Q101').value = data[17]['error']
            sheet.range('I113:Q113').value = data[18]['error']
            sheet.range('I133:Q133').value = data[19]['error']
            sheet.range('I124:Q124').value = data[20]['error']
            sheet.range('I92:K92').value = data[21]['error']
            sheet.range('I102:K102').value = data[22]['error']
            sheet.range('L92:N92').value = data[23]['error']
            sheet.range('L102:N102').value = data[24]['error']
            sheet.range('O92:Q92').value = data[25]['error']
            sheet.range('O102:Q102').value = data[26]['error']
            sheet.range('I115:Q115').value = data[27]['error']
            sheet.range('I135:Q135').value = data[28]['error']
            sheet.range('I126:Q126').value = data[29]['error']
            sheet.range('I94:K94').value = data[30]['error']
            sheet.range('I104:K104').value = data[31]['error']
            sheet.range('L94:N94').value = data[32]['error']
            sheet.range('L104:N104').value = data[33]['error']
            sheet.range('O94:Q94').value = data[34]['error']
            sheet.range('O104:Q104').value = data[35]['error']
            sheet.range('I116:Q116').value = data[36]['error']
            sheet.range('I136:Q136').value = data[37]['error']
            sheet.range('I95:K95').value = data[38]['error']
            sheet.range('L95:N95').value = data[39]['error']
            sheet.range('O95:Q95').value = data[40]['error']
            sheet.range('I117:Q117').value = data[41]['error']
        except FileNotFoundError:
            print('Por favor, ingrese primero archivo de energía activa ')
            sys.exit(1)

    wb.save(name_file)

# def plot_graph(name_file):
#
#     # Crea una lista con valores de phi unicos
#     list_phi = df1['Phi'].unique().tolist()   # Supongo que la fase 1 tendra los mismo angulos que la 2 y 3. En el ensayo por fase.
#     list_phi_4 = df4['Phi'].unique().tolist() # Lista de ang para ensayo energizacion simultanea 123.
#     estilos = ['b.-', 'rd-', 'gv-', 'co-']
#     a = 0
#     for phi in list_phi:
#
#         mask = df1['Phi'] == phi
#         #  DF  con las tres columnas(corriente, error, phi) filtrado por angulo.
#         filtrado_df1 = df1[mask]
#
#         # Sentencias para graficar, dentro del for, para que grafique todas las iteraciones(todos los angulos), y superponga las curvas.
#         ax1 = plt.subplot(2, 2, 1)
#         ax1.plot(filtrado_df1['Corriente Fase 1'], filtrado_df1['Error Fase 1'], estilos[a], label='phi '+ phi)               # (x,y para graficar) se filtra solo la columa del dataframe, con nombredataframe['nombre columna']
#         plt.xlabel('Corriente [A]')
#         plt.ylabel('Error fase 1 [%]')
#         plt.legend()
#
#         if not list_e_fase2:
#             pass
#         else:
#             filtrado_df2 = df2[mask]
#             ax2 = plt.subplot(2, 2, 2)
#             ax2.plot(filtrado_df2['Corriente Fase 2'], filtrado_df2['Error Fase 2'], estilos[a], label='phi ' + phi)
#             plt.xlabel('Corriente [A]')
#             plt.ylabel('Error fase 2 [%]')
#             plt.legend()
#
#             filtrado_df3 = df3[mask]
#             ax3 = plt.subplot(2, 2, 3)
#             ax3.plot(filtrado_df3['Corriente Fase 3'], filtrado_df3['Error Fase 3'], estilos[a], label='phi ' + phi)
#             # plt.title('Fig. 3')
#             plt.xlabel('Corriente [A]')
#             plt.ylabel('Error fase 3 [%]')
#             plt.legend()
#
#         a += 1
#     a = 0
#
#     for phi4 in list_phi_4:
#         mask = df4['Phi'] == phi4
#         filtrado_df4 = df4[mask]
#
#         if not list_e_fase123:
#             pass
#         else:
#             ax4 = plt.subplot(2, 2, 4)
#             ax4.plot(filtrado_df4['Corriente Fase 123'], filtrado_df4['Error Fase 123'], estilos[a], label='phi '+phi4)
#             plt.xlabel('Corriente [A]')
#             plt.ylabel('Error fase 123 [%]')
#             plt.legend()
#         a += 1
#     a = 0
#
#     # plt.figure(figsize= (15, 15))
#     plt.tight_layout()                 # Ajusta los graficos
#     plt.savefig(name_file + '.png')
#     # plt.show()
#     plt.close('all')

