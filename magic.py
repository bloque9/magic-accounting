#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
from xml2obj import xml2obj
from pyexcel_xlsx import save_data
from collections import OrderedDict


data = OrderedDict()

# Obtenemos el ruta
path = os.getcwd()

# Cambiamos al directorio 
os.chdir(path)

# Obtenemos los nombres de los archivos
archivos = os.listdir(path)

lista = [['Año','Mes', 'Dia', 'Factura', 'Proveedor', 'Lugar', 'Concepto', 'Total', 'IVA','Archivo']]
total = 0

# Recorremos todos los archivos
for item in archivos:
    
    # Obtenemos la extencion del archivo
    try:
        extencion = item.split('.')[1]
    except IndexError:
        extencion =''

    # Si el archivo es xml
    if extencion =='xml' or extencion =='XML':
        total=total + 1

        # Leemos los datos
        xml = open(item,'r')
        factura = xml2obj(xml.read())

        # Si tiene serie la obtenemos
        try:
            serie = factura.serie.replace('-','') + '-' + factura.folio
        except (TypeError, AttributeError):
            serie = factura.folio
        
        res = ''
        expedidoen = ''

        # Agregamos los conceptos
        try:
            for concepto in factura.cfdi_Conceptos.cfdi_Concepto:
                res = res + concepto.descripcion + ', '
        except AttributeError:
            print(item + ' Error Atribute')

        # Lugar
        try:
            expedidoen = factura.LugarExpedicion
        except (AttributeError, TypeError):
            print(item + ' error en expedido')

        # lo agregamos a la lista de datos
        try:
            lista.append([factura.fecha[0:4], factura.fecha[5:7], factura.fecha[8:10], serie, factura.cfdi_Emisor.nombre, expedidoen,
             res, float(factura.total), float(factura.cfdi_Impuestos.totalImpuestosTrasladados), item ])
        except (TypeError, AttributeError):
            print(item + ' TypeError')

# Agregamos le total de registros
lista.append(['total', total])

# Guardamos los datos
data.update({"MagicAccounting": lista })
save_data("Contabilidad.xlsx", data)

# Mensajes de exito
print("Se han escrito " + str(total) + " registros")
print("Presione una tecla para continuar")
raw_input()
