from abc import ABC, abstractmethod

#Crea una clase abstracta para las reservas
class book(ABC):
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

    @abstractmethod
    def descripcion(self):
        pass
    def __str__(self):
        return f"Fecha: {self.fecha}\nFactura: {self.factura}\nNombre: {self.nombre}\nNIF: {self.NIF}\nBase: {self.Base_1}\nCuota: {self.Cuota_1}\nTotal: {self.total}\nDomicilio: {self.domicilio}\nCódigo postal: {self.cod_postal}\nPaís: {self.pais}\nCL: {self.CL}\nObservaciones: {self.observaciones}\nCuenta contable: {self.cuenta_contable}\nTipo SII: {self.tipo_Sii}\n__________________________\n"


#Crea una clase para cada una de las distintas plataformas hereadando de la clase abstracta book
class booking(book):
    def __init__(self, fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, cuenta_contable, tipo_Sii):
        super().__init__(fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, cuenta_contable, tipo_Sii)
    def descripcion(self):
        return "Factura Booking"

class airbnb(book):
    def __init__(self, fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, cuenta_contable, tipo_Sii):
        super().__init__(fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, cuenta_contable, tipo_Sii)
    def descripcion(self):
        return "Factura Airbnb"
    
class web(book):
    def __init__(self, fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, cuenta_contable, tipo_Sii):
        super().__init__(fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, cuenta_contable, tipo_Sii)
    def descripcion(self):
        return "Factura Web"
    
class error(book):
    def __init__(self, fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, cuenta_contable, tipo_Sii):
        super().__init__(fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, cuenta_contable, tipo_Sii)
    def descripcion(self):
        return "Error"
