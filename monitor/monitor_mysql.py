import sys
import MySQLdb
import xlsxwriter


# destino
servidor = "172.17.63.4"
usuario = "msucre"
password = "S3cur15y"
database = "information_schema"

try:
	db = MySQLdb.connect(servidor,usuario,password,database,3306)
except MySQLdb.Error as e:
	print("No puedo conectar a la base de datos DESTINO:",e)
	sys.exit(1)

libro = xlsxwriter.Workbook('Monitor_MySQL.xlsx')
hoja = libro.add_worksheet("MySQL " + servidor )

hoja.write( 0, 0, "Monitor de MYSQL" )
hoja.write( 1, 0, "Cantidad Tablas & Registros" )

row = 3

hoja.write( row, 0, "TABLE_SCHEMA" )
hoja.write( row, 1, "TABLE_NAME" )
hoja.write( row, 2, "TABLE_ROWS" )
hoja.write( row, 3, "TABLE_COLLATION" )
row = row + 1 

sql="SELECT * FROM TABLES"
cursor = db.cursor(MySQLdb.cursors.DictCursor)
try:
   cursor.execute(sql)
   registros = cursor.fetchall()
   for registro in registros:
      muestra = 0
      if registro["TABLE_SCHEMA"] == "information_schema":
           muestra = 1
      if registro["TABLE_SCHEMA"] == "sys":
           muestra = 1
      if registro["TABLE_SCHEMA"] == "mysql":
           muestra = 1
      if registro["TABLE_SCHEMA"] == "phpmyadmin":
           muestra = 1
      if registro["TABLE_SCHEMA"] == "performance_schema":
           muestra = 1

      if muestra == 0: 
        hoja.write( row, 0, registro["TABLE_SCHEMA"] )
        hoja.write( row, 1, registro["TABLE_NAME"] )
        hoja.write( row, 2, registro["TABLE_ROWS"] )
        hoja.write( row, 3, registro["TABLE_COLLATION"] )

        row = row + 1 


except:
   print("Error en la consulta")
db.close()


# origen
#servidoro = "172.17.63.4"
#usuarioo = "msucre"
#passwordo = "S3cur15y"
#databaseo = "information_schema"

servidoro = "127.0.0.1"
usuarioo = "root"
passwordo = "password2017"
databaseo = "information_schema"

try:
	dbo = MySQLdb.connect(servidoro,usuarioo,passwordo,databaseo,3399)
except MySQLdb.Error as e:
	print("No puedo conectar a la base de datos ORIGEN:",e)
	sys.exit(1)

hoja = libro.add_worksheet( "MySQL " + servidoro )

hoja.write( 0, 0, "Monitor de MYSQL" )
hoja.write( 1, 0, "Cantidad Tablas & Registros" )

row = 3

hoja.write( row, 0, "TABLE_SCHEMA" )
hoja.write( row, 1, "TABLE_NAME" )
hoja.write( row, 2, "TABLE_ROWS" )
hoja.write( row, 3, "TABLE_COLLATION" )
row = row + 1 

sql="SELECT * FROM TABLES"
cursor = dbo.cursor(MySQLdb.cursors.DictCursor)
try:
   cursor.execute(sql)
   registros = cursor.fetchall()
   for registro in registros:
      muestra = 0
      if registro["TABLE_SCHEMA"] == "information_schema":
           muestra = 1
      if registro["TABLE_SCHEMA"] == "sys":
           muestra = 1
      if registro["TABLE_SCHEMA"] == "mysql":
           muestra = 1
      if registro["TABLE_SCHEMA"] == "phpmyadmin":
           muestra = 1
      if registro["TABLE_SCHEMA"] == "performance_schema":
           muestra = 1

      if muestra == 0: 
        hoja.write( row, 0, registro["TABLE_SCHEMA"] )
        hoja.write( row, 1, registro["TABLE_NAME"] )
        hoja.write( row, 2, registro["TABLE_ROWS"] )
        hoja.write( row, 3, registro["TABLE_COLLATION"] )

        row = row + 1 

except:
   print("Error en la consulta")
dbo.close()


import psycopg2

servidor = "172.17.63.4"
usuario = "msucre"
password = "S3cur15y"
database = "information_schema"


# Conexión a la base de datos
connection = psycopg2.connect(
    host="localhost",
    port=5432,
    database="web_nueva",
    user="postgres",
    password="password2017",
)

# Obtención de un cursor
cursor = connection.cursor()

#SELECT * FROM pg_database ;

# Ejecución de una consulta SQL
#cursor.execute("SELECT * FROM pg_database")

cursor.execute("SELECT * FROM pg_tables")

# Obtención de los resultados
results = cursor.fetchall()

# Cerrar el cursor
cursor.close()

# Cerrar la conexión
connection.close()


hoja = libro.add_worksheet( "Potgres " + "localhost" )

hoja.write( 0, 0, "Monitor de POSTGRES" )
hoja.write( 1, 0, "Cantidad Tablas & Registros" )

row = 3

hoja.write( row, 0, "TABLE_SCHEMA" )
hoja.write( row, 1, "TABLE_NAME" )
hoja.write( row, 2, "TABLE_ROWS" )
hoja.write( row, 3, "TABLE_COLLATION" )
row = row + 1 


# Impresión de los resultados
for rowpg in results:
    #print(rowpg)
    hoja.write( row, 0, rowpg[0] )
    hoja.write( row, 1, rowpg[1] )
    hoja.write( row, 2, rowpg[2] )
    row = row + 1 


libro.close()



