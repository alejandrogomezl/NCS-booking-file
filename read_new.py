import csv
from books import airbnb, booking, web

class read_csv:
    def __init__(self, file_in):
        self.file_in = file_in
        self.platforms = {
            "airbnb": airbnb,
            "booking": booking,
            "mara": web
        }

    def read_csv(self):
        reservas = []
        
        with open(self.file_in, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter=';')
            next(reader)  # Saltar el encabezado

            for row in reader:
                fecha = str(row[0])
                factura = "MBA24/" + str(row[1])
                nombre = row[5]
                NIF = row[7]
                Base_1 = row[12]
                Cuota_1 = row[14]
                total = row[15]
                domicilio = row[8]
                cod_postal = row[10]
                pais = row[11]
                CL = 2
                observaciones = row[4]  # Cambiado a la columna correcta
                tipo_Sii = 1

                # Intentar crear las reservas
                try:
                    obs = observaciones.lower()
                    for key, platform in self.platforms.items():
                        if key in obs:
                            reservas.append(platform(fecha, factura, nombre, NIF, Base_1, Cuota_1, total, domicilio, cod_postal, pais, CL, observaciones, 430000, tipo_Sii))
                except AttributeError:
                    continue
        
        return reservas
    
    def test(self):
        print("hola")

