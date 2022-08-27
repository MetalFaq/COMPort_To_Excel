
list_i_fase1 = []
list_e_fase1 = []
list_i_fase2 = []
list_e_fase2 = []
list_i_fase3 = []
list_e_fase3 = []
list_i_fase123 = []
list_e_fase123 = []
list_phi_f1 = []
list_phi_f2 = []
list_phi_f3 = []
list_phi_f123 = []

#Recibe como parámetro una lista de diccionarios
def dic_to_list(list_info):

    for param in list_info:
        if param['fase'] == '1--':
            list_i_fase1.append(param['corriente'])
            list_e_fase1.append(param['error'])
            list_phi_f1.append(param['phi'])

        elif param['fase'] == '-2-':
            list_i_fase2.append(param['corriente'])
            list_e_fase2.append(param['error'])
            list_phi_f2.append(param['phi'])

        elif param['fase'] == '--3':
            list_i_fase3.append(param['corriente'])
            list_e_fase3.append(param['error'])
            list_phi_f3.append(param['phi'])

        elif param['fase'] == '123':
            list_i_fase123.append(param['corriente'])
            list_e_fase123.append(param['error'])
            list_phi_f123.append(param['phi'])

# def dic_to_list(list_info, key, value):
#
#     for param in list_info:
#         if (key(param) == value) and (param['fase'] == '1--'):
#             list_i_fase1.append(param['corriente'])
#             list_e_fase1.append(param['error'])
#
#         elif (key(param) == value) and (param['fase'] == '-2-'):
#             list_i_fase2.append(param['corriente'])
#             list_e_fase2.append(param['error'])
#
#
#         elif (key(param) == value) and (param['fase'] == '--3'):
#             list_i_fase3.append(param['corriente'])
#             list_e_fase3.append(param['error'])
#
#
#         elif key(param) == value and param['fase'] == '123':
#             list_i_fase123.append(param['corriente'])
#             list_e_fase123.append(param['error'])
            



    # print(list_i_fase1)
    # print(list_e_fase1)
    # print(list_fases)