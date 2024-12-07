import xlsxwriter
libro = xlsxwriter.Workbook('Presupuesto1.xlsx')
hoja = libro.add_worksheet()

 # Add a bold format to use to highlight cells.
bold = libro.add_format({'bold': True})

# Add a number format for cells with money.
money = libro.add_format({'num_format': '$#,##0'})

 # Write some data headers.
#worksheet.write('A1', 'Item', bold)
#worksheet.write('B1', 'Cost', bold)

# Insert an image.
hoja.insert_image("A1", "logo.png")

# El presupuesto que pintaremos en la hoja de cálculo
presupuesto = (
    ['Equipos',     4000],
    ['Cable',        100],
    ['Armario',      200],
    ['Switch',        99],
    ['AP',            50],
    ['Router',       150],
    ['Mano de Obra', 350],
)

# Nos posicionamos en la primera columna de la primera fila
row = 5
col = 0

hoja.write(row, col,     "Descripcion", bold )
hoja.write(row, col + 1, "Precio a Mostrar", bold )
row += 1

# Iteramos los datos para ir pintando fila a fila
for concepto, precio in (presupuesto):
    hoja.write(row, col,     concepto, bold )
    hoja.write(row, col + 1, precio, money )
    row += 1
 
#Pintamos la fila de totales
hoja.write(row, 0, 'Total:', bold)
hoja.write(row, 1, '=SUM(B1:B7)', money)

# Autofit the worksheet.
hoja.autofit()

#Cerramos el libro
libro.close()


libro = xlsxwriter.Workbook('Analisis_Resultado.xlsx')
hoja = libro.add_worksheet("Punto 2")

row = 4

hoja.write( 0, 0, "" )
hoja.write( 1, 0, "Punto 2. Capacidades de los equipos tecnológicos" )
hoja.write( 2, 0, "Capacidad y desempeño del hardware" )

hoja.write( row+1, 0, "Artículo" )
hoja.write( row+2, 0, "" )
hoja.write( row+3, 0, "" )

hoja.write( row+1, 1, "Equipo Producción" )
hoja.write( row+2, 1, "Plataforma central de telecomunicaciones Voz" )
hoja.write( row+3, 1, "Nortel CS1000MG  RLS 7.6" )

hoja.write( row+1, 2, "Servicio/Aplicación Soportada" )
hoja.write( row+2, 2, "Sistemas de grabación voz y Llamadas Predictivas" )
hoja.write( row+3, 2, "Sistema Core Centrales Telecomunicaciones Voz" )

hoja.write( row, 3, "Procesamiento" )
hoja.write( row+1, 3, "Total" )
hoja.write( row+2, 3, "9 Cores INTEL" )
hoja.write( row+3, 3, "2.992 M idle cycles" )

hoja.write( row, 4, "" )
hoja.write( row+1, 4, "Usado %" )
hoja.write( row+2, 4, "4,21%" )
hoja.write( row+3, 4, "0,19%" )

hoja.write( row, 5, "Memoria" )
hoja.write( row+1, 5, "Total" )
hoja.write( row+2, 5, "14,28" )
hoja.write( row+3, 5, "0,22" )

hoja.write( row, 6, "" )
hoja.write( row+1, 6, "Usado %" )
hoja.write( row+2, 6, "45,00%" )
hoja.write( row+3, 6, "23,57%" )

hoja.write( row, 7, "Almacenamiento" )
hoja.write( row+1, 7, "Total" )
hoja.write( row+2, 7, "6249,47 GB" )
hoja.write( row+3, 7, "1599,45 MB" )

hoja.write( row, 8, "" )
hoja.write( row+1, 8, "Usado %" )
hoja.write( row+2, 8, "44,27%" )
hoja.write( row+3, 8, "15,05%" )

hoja.write( 8, 0, "" )
hoja.write( 9, 0, "Líneas comunicacionales" )
hoja.write( 10, 0, "" )

row = 12

hoja.write( row+1, 0, "Artículo" )
hoja.write( row+2, 0, "" )
hoja.write( row+3, 0, "" )

hoja.write( row+1, 1, "Equipo Producción" )
hoja.write( row+2, 1, "Firewall Cluster CPDA" )
hoja.write( row+3, 1, "Firewall Cluster CPDA" )

hoja.write( row+1, 2, "Servicio/Aplicación Soportada" )
hoja.write( row+2, 2, "bmcpdavsxa" )
hoja.write( row+3, 2, "bmcpdavsxb" )

hoja.write( row, 3, "Procesamiento" )
hoja.write( row+1, 3, "Total" )
hoja.write( row+2, 3, "16" )
hoja.write( row+3, 3, "16" )

hoja.write( row, 4, "" )
hoja.write( row+1, 4, "Usado %" )
hoja.write( row+2, 4, "4,21%" )
hoja.write( row+3, 4, "0,19%" )

hoja.write( row, 5, "Memoria" )
hoja.write( row+1, 5, "Total" )
hoja.write( row+2, 5, "14,28" )
hoja.write( row+3, 5, "0,22" )

hoja.write( row, 6, "" )
hoja.write( row+1, 6, "Usado %" )
hoja.write( row+2, 6, "45,00%" )
hoja.write( row+3, 6, "23,57%" )

hoja.write( row, 7, "Almacenamiento" )
hoja.write( row+1, 7, "Total" )
hoja.write( row+2, 7, "6249,47 GB" )
hoja.write( row+3, 7, "1599,45 MB" )

hoja.write( row, 8, "" )
hoja.write( row+1, 8, "Usado %" )
hoja.write( row+2, 8, "44,27%" )
hoja.write( row+3, 8, "15,05%" )


hoja.autofit()

libro.close()


