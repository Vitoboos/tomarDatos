import sys
import wmi
import platform
from multiprocessing import freeze_support
import cpuinfo
import psutil 
import uuid
import socket

import math
import pandas as pd

print(sys.path)
c = wmi.WMI()

# Marca y Modelo del computador

nombre = c.Win32_ComputerSystem()[0].Name
marca = c.Win32_ComputerSystem()[0].Manufacturer
modelo = c.Win32_computerSystem()[0].Model

# Numero de serie
serial = c.Win32_BIOS()[0].SerialNumber

# Toma de datos con WMIC
print(f"Marca del dispositivo: {marca}")
print(f"Modelo del dispositivo: {modelo}")
print(f"Numero de Serie: {serial}")

# Sistema Operativo

sistema_operativo = f"{platform.system()} {platform.release()}"
print(f"Sistema Operativo: {sistema_operativo}")

# Informacion del procesador

freeze_support()
cpu_info = cpuinfo.get_cpu_info()
print(f"Procesador: {cpu_info['brand_raw']}")

# Informacion de RAM

memoria_ram = math.ceil(psutil.virtual_memory().total  / (1024.0 **  3))
print(f'Memoria RAM: {memoria_ram} GB')

# Informacion del disco

discos = c.Win32_LogicalDisk()
disco_modelo = c.Win32_DiskDrive()
disco_nombre = []
capacidad = []

for disco in disco_modelo:
    disco_nombre.append(f"{disco.Model}") 
    capacidad.append( f"{math.floor(int(disco.Size) / 1000 ** 3)} GB")
    
# Informacion de la red 

# MAC
def tomar_mac():
    node = uuid.getnode()
    mac = ':'.join(['{:02x}'.format((node >> i) &  0xff) for i in range(0,  8*6,  8)][::-1])
    return mac

tomar_mac()

# IP
def tomar_ip():
    hostname = socket.gethostname()
    ip = socket.gethostbyname(hostname)

    print("Nombre de dispositivo en el dominio:", hostname)
    print("IP del dispositivo:", ip)

    return [hostname, ip]
    
# Creacion de Tabla de Excel

data = {
    'nombre_equipo': [],
    'marca': [],
    'modelo': [],
    'serial': [],
    'sistema_operativo': [],
    'procesador': [],
    'ram': [],
    'disco': [],
    'disco_tipo': [],
    'ipv4': [],
    'mac_address': []
}

dataframe = pd.DataFrame(data)

inventario = {
    'nombre_equipo': [tomar_ip()[0]],
    'marca': [marca],
    'modelo': [modelo],
    'serial': [serial],
    'sistema_operativo': [sistema_operativo],
    'procesador': [cpu_info['brand_raw']],
    'ram': [f'{memoria_ram} GB'],
    'disco': [disco_nombre[0]],
    'disco_tipo': [capacidad[0]],
    'ipv4': [tomar_ip()[1]],
    'mac_address': [tomar_mac()]
}

inventario_df = pd.DataFrame([inventario])
dataframe = pd.concat([dataframe, inventario_df], ignore_index=True)

with pd.ExcelWriter(f'{nombre}.xlsx', engine='xlsxwriter') as writer:
    dataframe.to_excel(writer, index=False)
    
