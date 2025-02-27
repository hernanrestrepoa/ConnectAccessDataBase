//La librería Pyodbc es una biblioteca de Python que se utiliza para conectar bases de datos ODBC (Open Database Connectivity) a través de una interfaz.

//Pyodbc es un módulo de código abierto de Python que simplifica el acceso a bases de datos ODBC como Oracle, MySQL, Access, Excel, PostgreSQL, SQL Server, etc.

//Pyodbc es compatible con Python 2 y 3, y está disponible para Windows, Linux, OS/X

//El Pyodbc permite conectarnos a bases de datos, ejecutar consultas SQL (Structured Query Language) "Lenguaje de Consulta Estructurado", obtener resultados, administrar transacciones, hacer consultas, etc.

//Importamos la librería pyodbc con las siguientes líneas de código: 

import  pyodbc 

//Creamos una variable llamada “conexion” para conectar el Python con el driver de Microsoft Access a través del argumento DRIVER, usamos el argumento DBQ para indicar la ubicación del archivo de la base de datos de Access. Empleamos
las siguientes líneas de código: 

conexion=pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=C:\Users\herna\Documents\TUTORIAL\tutorial\BaseDeDatos.accdb')

//Creamos la variable llamada “Cursor” para poder ejecutar comandos de SQL en Python como por ejemplo: INSERT, SELECT, WHERE, CREATE, RENAME, AS, con las siguientes líneas de código:

cursor=conexion.cursor()

//Probamos la conexión creando un nuevo registro en el interior de la tabla Categorias de la base de datos Access usando el comando INSERT INTO, confirmamos los cambios realizados con sentencias SQL con el comando “commit”, con las siguientes líneas de código: 

cursor.execute('INSERT INTO Categorias (IdCat,Categoria) VALUES(13,\'Discos Externos\')')
conexion.commit()

//Al abrir la base de datos vemos que se ha creado un nuevo registro en la tabla Categorias. Vamos a imprimir todos los registros de la tabla Categorias, con el comando SQL llamado SELECT * FROM Categorias. Con el comando “fetchall” y un ciclo for obtenemos todos los registros de esta consulta y los imprimimos. Usamos las siguientes líneas de código:

cursor.execute('SELECT * FROM Categorias')
for Fila in cursor.fetchall():
    print(Fila)

//Cerramos la función del cursor asi como la conexión a la base de datos con las siguientes líneas de código:

cursor.close()
conexion.close()
