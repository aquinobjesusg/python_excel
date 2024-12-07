
import pylightxl
archivo = 'cuadro.xlsx'  
db = pylightxl.readxl('cuadro.xlsx')

# request a semi-structured data (ssd) output
ssd = db.ws('TU').ssd(keycols="KEYCOLS", keyrows="KEYROWS")




