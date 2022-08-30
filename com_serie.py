'''Si tenés el nombre SERIAL por algun lado dando vuelta nombrando alguna carpeta o archivo, reemplazalo por otro.'''

import serial.tools.list_ports as port_lists
import serial
import os
from time import sleep

abs_path = os.path.dirname(__file__)
relative_path = "/DATA"
full_path = abs_path + relative_path

# Si no existe el directorio, lo crea el programa
if not os.path.exists(full_path):
    os.mkdir(full_path)

# Para saber los puertos disponibles:
# 1) En la función "read_serial_port()" colocá un punto de debug
# 2) y sacá los comentarios de la línea 17, 18 y 19. También de la libreria port_lists
# 3) Corré en modo debug y en la consola te va a imprimir los COM disponible.
#
# ports = list(port_lists.comports())
# for p in ports:
#     print(p)

def read_serial_port():

    ports = list(port_lists.comports())
    for p in ports:
        print(p)

    port_com = input("Ingrese Puerto de comunicacion: ")
    ser = serial.Serial(port_com, 115200)
    ser.setDTR(False) # Data Terminal Ready deshabilitado
    sleep(1)
    ser.flushInput()
    ser.setDTR(True)
    # start = time.time()
    dato = []
    cadena = " "

    while True:
        dato = ser.readline()
        dato = dato.decode('UTF-8')      # Por defecto se codifica en byte. Lo convierto a string
        # print(dato)
        cadena = cadena + dato           # concatena cada linea.
        subcadena = cadena[-6:]          # Verifica los ultimos 6 caracteres de la cadena.
        if subcadena == "\r\n\r\n\r\n":
            break

    ser.close()
    filename = full_path + '/data.txt'
    write_file = open(filename, 'w')
    write_file.write(cadena)
    write_file.close()
