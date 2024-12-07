
# pip install xlsxwriter

# pip install xlrd
# # pip install pandas
# #pip install cython

# pip install NumPy
# pip install python-dateutil
# pip install pytz

#pip unistall openpixl
#pip uninstall pandas
#pip uninstall aspose-cells
#pip install cython

import xlrd 
#archivo = 'C:\proyectos\excel\cuadro1.xls'
#archivo = 'cuadro1.xls'
archivo = 'cuadro.xlsx'  

# Leer un archivo de Excel & Mostrar un Valor
wb = xlrd.open_workbook(archivo) 
hoja = wb.sheet_by_index(0) 
print(hoja.nrows) 
print(hoja.ncols) 
print(hoja.cell_value(0, 0))

# Mostrar los Datos de la Columna 1 de la HOJA TU
wb = xlrd.open_workbook(archivo) 
hoja = wb.sheet_by_name('TU') 
for i in range(0,hoja.nrows):
    print(hoja.cell_value(i,1))

# Saber los Nombres de las Columnas
wb = xlrd.open_workbook(archivo) 
hoja = wb.sheet_by_index(0) 
nombres = hoja.row(0)  
print(nombres)

hoja = wb.sheet_by_index(1) 
nombres = hoja.row(0)  
print(nombres)

hoja = wb.sheet_by_index(2) 
nombres = hoja.row(0)  
print(nombres)



# Abre el archivo Excel
workbook = xlrd.open_workbook(archivo)

# Obtiene la hoja de trabajo "Hoja1"
hoja = workbook.sheet_by_name("TU")

# Imprime el número de filas y columnas de la hoja de trabajo
print(hoja.nrows, hoja.ncols)

# Imprime los datos de la hoja de trabajo
for fila in range(hoja.nrows):
    for columna in range(hoja.ncols):
        print(hoja.cell_value(fila, columna), end=" ")
    print()


# Leer un rango de Lista
#hoja = wb.sheet_by_index(0) 
#hoja = wb.sheet_by_index(1) 

# # Creamos listas
# filas = []
# for fila in range(1,hoja.nrows):
#     columnas = []
#     for columna in range(0,2):
#         columnas.append(hoja.cell_value(fila,columna))
#     filas.append(columnas)








# # Por consola de comandos Instlar el Excel
# # pip install pandas

#import pandas as pd
#archivo = 'cuadro.xlsx'

#df = pd.read_excel(archivo, sheet_name='TU')

#df.describe()


# # Por consola de comandos Instlar el Excel
# # pip install pandas


# #print( "programa")


# import openpyxl
# excel_document = openpyxl.load_workbook('sample.xlsx')





# # Get worksheets collection
# collection = excel_document.get_sheet_names()
# collectionCount = collection.getCount()

# print( collection )
# print( collectionCount )


# # Loop through all the worksheets
# for worksheetIndex in range(collectionCount):

#     # Get worksheet using its index
#     worksheet = collection.get(worksheetIndex)

#     # Print worksheet name
#     print("Worksheet: " + str(worksheet.getName()))

#     # Get number of rows and columns
#     rows = worksheet.getCells().getMaxDataRow()
#     cols = worksheet.getCells().getMaxDataColumn()

#     # Loop through rows
#     for i in range(rows):

#         # Loop through each column in selected row
#         for j in range(cols):
#             # Print cell value
#             print(worksheet.getCells().get(i, j).getValue(), end =" | ")

#         # Print line break
#         print("\n")
