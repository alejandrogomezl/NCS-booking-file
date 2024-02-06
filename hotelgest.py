import openpyxl
from openpyxl import Workbook
import os
import shutil
import pandas as pd

class book:
    def __init__(self, fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, cuenta_contable, tipo_Sii):
        self.fecha = fecha
        self.factura = factura
        self.nombre = nombre
        self.NIF = NIF
        self.Base_1 = Base_1
        self.Cuota_1 = Cuota_1
        self.total = total
        self.domicilio = domicilio
        self.cod_postal = cod_postal
        self.pais = pais
        self.CL = CL
        self.observaciones = observaciones
        self.cuenta_contable = cuenta_contable
        self.tipo_Sii = tipo_Sii

    def __str__(self):
        return f"Fecha: {self.fecha}\nFactura: {self.factura}\nNombre: {self.nombre}\nNIF: {self.NIF}\nBase: {self.Base_1}\nCuota: {self.Cuota_1}\nTotal: {self.total}\nDomicilio: {self.domicilio}\nCódigo postal: {self.cod_postal}\nPaís: {self.pais}\nCL: {self.CL}\nObservaciones: {self.observaciones}\nCuenta contable: {self.cuenta_contable}\nTipo SII: {self.tipo_Sii}\n__________________________\n"


books_airbnb = []
books_booking = []
books_web = []
error = []

file_in = "in1.xlsx"
wb = openpyxl.load_workbook(file_in)
hoja = wb.active
last = hoja.max_row
for i in range(1,last):
    fecha = str(hoja.cell(i+1,5).value)
    factura = "MBA24/" + str(hoja.cell(i+1,4).value)
    nombre = hoja.cell(i+1,14).value
    NIF = hoja.cell(i+1,15).value
    Base_1 = hoja.cell(i+1,18).value
    Cuota_1 = hoja.cell(i+1,19).value
    total = hoja.cell(i+1,20).value
    domicilio = hoja.cell(i+1,11).value
    cod_postal = hoja.cell(i+1,12).value
    pais = hoja.cell(i+1,10).value
    CL = 2
    observaciones = hoja.cell(i+1,26).value
    tipo_Sii = 1

    try:
        obs=observaciones.lower()
    except AttributeError:
        continue
    if "airbnb" in obs:
        books_airbnb.append(book(fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, 70500000060, tipo_Sii))
    elif "booking" in obs:
        books_booking.append(book(fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, 70500000057, tipo_Sii))
    elif "mara" in obs:
        books_web.append(book(fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, 70500000053, tipo_Sii))
    else:
        error.append(book(fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, 0, tipo_Sii))

wb.close()

def conv(file):
    df = pd.read_excel(file, engine='openpyxl')
    file_new = file[:-5] + ".xls"
    with pd.ExcelWriter(file_new, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    os.remove(file)

def write (books):
    if books == books_airbnb and books_airbnb != []:
        file_out = "out/airbnb.xlsx"
    elif books == books_booking and books_booking != []:
        file_out = "out/booking.xlsx"
    elif books == books_web and books_web != []:
        file_out = "out/web.xlsx"
    elif books == error and error != []:
        file_out = "out/error.xlsx"
    else:
        return

    try:
        wb = openpyxl.load_workbook(file_out)
    except FileNotFoundError:
        wb = Workbook()

    hoja = wb.active
    last = 0
    for i in books:
        hoja.cell(last+1,1).value = i.fecha
        hoja.cell(last+1,2).value = i.factura
        hoja.cell(last+1,3).value = i.nombre
        hoja.cell(last+1,4).value = i.NIF
        hoja.cell(last+1,5).value = i.Base_1
        hoja.cell(last+1,7).value = i.Cuota_1
        hoja.cell(last+1,8).value = i.total
        hoja.cell(last+1,9).value = i.domicilio
        hoja.cell(last+1,10).value = i.cod_postal
        hoja.cell(last+1,11).value = i.pais
        hoja.cell(last+1,12).value = i.CL
        hoja.cell(last+1,14).value = i.observaciones
        hoja.cell(last+1,15).value = i.cuenta_contable
        hoja.cell(last+1,16).value = i.tipo_Sii
        last += 1
        print(f"Factura {i.factura} añadida")

    wb.save(file_out)
    wb.close()
    print(f"Archivo {file_out} guardado con {last} facturas\n\n")


shutil.rmtree("out")
os.mkdir("out")


write(books_airbnb)
write(books_booking)
write(books_web)
write(error)
for element in os.listdir("out"):
    conv("out/"+element)