# Excel to SQL Importer

Excel to SQL Importer es una aplicación de interfaz gráfica de usuario (GUI) desarrollada en Python para importar datos desde un archivo Excel a una base de datos SQL Server. La aplicación permite conectar a un servidor SQL, mapear columnas del archivo Excel a columnas de la base de datos, y realizar la inserción de datos en tablas específicas.

## Características

- Conexión a servidores SQL Server utilizando ODBC.
- Importación de datos desde archivos Excel (.xlsx, .xls).
- Mapeo de columnas del archivo Excel a columnas de la base de datos.
- Inserción de datos en tablas "Personas" y "Proveedores" en la base de datos seleccionada.
- Interfaz gráfica de usuario fácil de usar.

## Requisitos

- Python 3.7 o superior
- Paquetes de Python: `pyodbc`, `pandas`, `yaml`, `tkinter`

## Instalación

1. **Clonar el repositorio**:  
   ```bash
   git clone https://github.com/tu-usuario/excel-to-sql-importer.git
   cd excel-to-sql-importer

2. **Instalar dependencias**:
Ejecuta el siguiente comando para instalar los paquetes necesarios:

```bash
pip install -r requirements.txt
```

3. **Configurar archivos YAML** (opcional):
Crea un archivo config.yaml para almacenar la configuración de conexión a los servidores SQL (en caso de no ser configurado tambien se pueden introducir los datos de conexion en la pantalla de la aplicacion):

 ```yaml
servidor-ejemplo:
  user: "nombre_usuario"
  password: "contraseña"
```
Crea un archivo defaultvalues.yaml para definir los valores por defecto para las columnas (en caso de no ser creado se preguntara por cada columna que valor debe tener por defecto. Si hay nuevas columnas en la tabla tambien se preguntara el valor por defecto a introducir de estas:

```yaml
default_values_personas:
  numIdPaisSii: 0
  strTitular: ""
  # Agrega otros valores por defecto según sea necesario
 
default_values_proveedores:
  numIdPaisSii: 0
  # Agrega otros valores por defecto según sea necesario
```

4. **Ejecutar la aplicación**:
Ejecuta el script principal para iniciar la aplicación:
```bash
python excel_to_sql_importer.py
```

5. **Conectar al servidor SQL**:
Selecciona el servidor de la lista desplegable o ingresa manualmente las credenciales.
Haz clic en "Conectar" para establecer la conexión.

6. **Seleccionar la base de datos y tipo de datos**:
Elige la base de datos y el tipo de datos ("Personas" o "Proveedores") desde los menús desplegables.

7. **Importar archivo Excel**:
Haz clic en "Importar Excel" y selecciona el archivo Excel desde el que deseas importar los datos.

8. **Mapear columnas**:
Mapea las columnas del archivo Excel a las columnas de la base de datos mediante la interfaz de mapeo.
Confirma el mapeo para continuar.

9. **Insertar datos**:
Los datos serán insertados en la tabla seleccionada, mostrando una barra de progreso durante la operación.

## Estructura del Proyecto
.
├── config.yaml
├── defaultvalues.yaml
├── excel_to_sql_importer.py
└── icono.ico

## Contribuciones
Las contribuciones son bienvenidas. Por favor, realiza un fork del repositorio y envía un pull request con tus cambios.
