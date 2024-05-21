import pandas as pd
import pyodbc
import shutil

# Configuración de la base de datos
server = r'nombre del servidor'
database = 'nombre de la base de datos'

# Intenta primero con autenticación de Windows
try:
    conn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;'
    conn = pyodbc.connect(conn_str)
    print("Conexión exitosa con autenticación de Windows.")
except pyodbc.Error as e:
    print(f"Fallo en la autenticación de Windows: {e}")
    # Si falla, intenta con autenticación SQL
    try:
        username = r'nombre del usurio'
        password = 'clave de usuario.'
        conn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        conn = pyodbc.connect(conn_str)
        print("Conexión exitosa con autenticación SQL.")
    except pyodbc.Error as e:
        print(f"Fallo en la autenticación SQL: {e}")
        raise

# Verificar y leer el archivo Excel
try:
    original_file = r'ruta/archivo.xlsx'
    temp_file = r'ruta/archivo_temp.xlsx'
    
    # Copiar el archivo a una ubicación temporal
    shutil.copyfile(original_file, temp_file)
    
    # Leer el archivo de la ubicación temporal
    df = pd.read_excel(temp_file)
    print("Archivo leído exitosamente desde la ubicación temporal")

    # Nombre de la tabla en la base de datos
    table_name = 'nombre_de_la_tabla'

    # Columnas de la tabla
    columns = ','.join(df.columns)

    # Insertar los datos en la tabla
    cursor = conn.cursor()
    for index, row in df.iterrows():
        # Construir la lista de valores
        values = []
        for value in row:
            if pd.isna(value):
                values.append('NULL')
            elif isinstance(value, str):
                values.append(f"'{value.replace('\'', '\'\'')}'")
            elif isinstance(value, (int, float)):
                values.append(str(value))
            elif isinstance(value, pd.Timestamp):
                values.append(f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'")
            else:
                values.append(f"'{str(value).replace('\'', '\'\'')}'")

        values_str = ','.join(values)
        query = f"INSERT INTO {table_name} ({columns}) VALUES ({values_str})"
        cursor.execute(query)

    # Confirmar los cambios en la base de datos
    conn.commit()
    print("Datos insertados correctamente.")
except Exception as e:
    print(f"Ocurrió un error durante la inserción de datos: {e}")
finally:
    # Cerrar la conexión
    if 'conn' in locals():
        conn.close()
