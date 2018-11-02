#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import sys
from pyexcel_xlsx import save_data
from collections import OrderedDict

import xmltodict
import json

data = OrderedDict()

# Obtenemos el ruta
path = os.getcwd()

# Cambiamos al directorio 
os.chdir(path)

# Obtenemos los nombres de los archivos
archivos = os.listdir(path)

lista = [[u'Año','Mes', 'Dia', 'Hora', 'Factura', 'Proveedor', 'Lugar', 'Concepto', 'Total', 'IVA','Archivo']]
to = 0

# Recorremos todos los archivos
for item in archivos:
    
    # Obtenemos la extencion del archivo
    try:
        s = item.split('.')
        extencion = s[len(s)-1]
    except IndexError:
        extencion =''

    # Si el archivo es xml
    if extencion =='xml' or extencion =='XML':
        to=to + 1

        # Leemos los datos
        if sys.version_info[0] < 3:
            xml = open(item,'r')
        else:
            xml = open(item, "rb")

        factura = xmltodict.parse(xml)

        factura = factura['cfdi:Comprobante']

        if factura['@Version']=='3.3':
            # folio
            try:
                folio = factura['@Folio']
            except:
                folio = 'na'
            
            # codigo postal
            try:
                cp = factura['@LugarExpedicion']
            except:
                cp = 'na'

            # fecha
            try:
                fecha = factura['@Fecha']
                f = fecha.split('T')
                fecha = f[0].split('-')
                anio = fecha[0]
                mes = fecha[1]
                dia = fecha[2]
                hora = f[1]
            except:
                anio = 'na'
                mes = 'na'
                dia = 'na'
                hora = 'na'

            # conceptos
            try:
                descripcion = ''
                for concepto in factura['cfdi:Conceptos']['cfdi:Concepto']:
                    descripcion = descripcion + concepto['@Descripcion'] + ', '
            except:
                descripcion = 'na'

            # importe
            try:
                total = factura['@Total']
            except:
                total = 'na'

            # iva
            try:
                iva = factura['cfdi:Impuestos']['@TotalImpuestosTrasladados']
            except:
                iva = '0'

            # emisor
            try:
                emisor = factura['cfdi:Emisor']['@Nombre']
            except:
                emisor = 'na'

            # lo agregamos a la lista de datos
            try:
                lista.append([
                    anio, 
                    mes, 
                    dia, 
                    hora, 
                    folio, 
                    emisor, 
                    cp,
                    descripcion, 
                    float(total), 
                    float(iva), 
                    item ])
            except:
                print(item + ' TypeError')
        else:
            print(item + ' ErrorVersion')


# Agregamos le total de registros
lista.append(['total', to])

# Guardamos los datos
data.update({"MagicAccounting": lista })
save_data("Contabilidad.xlsx", data)

# Mensajes de exito
print("CFDI version 3.3")
print("Se han escrito " + str(to) + " registros")
print("Presione una tecla para continuar")
if sys.version_info[0] < 3:
    raw_input()
else:
    input()