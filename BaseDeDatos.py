import  pyodbc 
conexion=pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=C:\Users\herna\Documents\TUTORIAL\tutorial\BaseDeDatos.accdb')
cursor=conexion.cursor()
cursor.execute('INSERT INTO Categorias (IdCat,Categoria) VALUES(13,\'Discos Externos\')')
conexion.commit()
cursor.execute('SELECT * FROM Categorias')
for Fila in cursor.fetchall():
    print(Fila)
cursor.close()
conexion.close()
