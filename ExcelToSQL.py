import pyodbc
import pandas as pd
import yaml
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import difflib

# Ruta del archivo YAML
CONFIG_FILE = 'config.yaml'
SIMILAR = 0.8

class ExcelToSQLImporter:
    def __init__(self):
        self.config = self.load_config()
        self.connection = None
        self.excel_data = None
        self.column_mappings = {}

        # Inicializar la interfaz gráfica
        self.root = tk.Tk()
        self.root.resizable(False, False)

        self.root.title("Importar Excel")

        self.root.iconbitmap('icono.ico')


        self.setup_ui()
        self.root.mainloop()
    
    def load_default_values(self):
        # Cargar los valores por defecto desde el archivo YAML
        with open('defaultvalues.yaml', 'r', encoding='utf-8') as file:
            return yaml.safe_load(file)

    def load_config(self):
        try:
            with open(CONFIG_FILE, 'r') as file:
                return yaml.safe_load(file) or {}
        except FileNotFoundError:
            return {}

    def save_config(self):
        with open(CONFIG_FILE, 'w') as file:
            yaml.dump(self.config, file)

    def connect_to_server(self, server, user, password):
        try:
            connection_string = (
                f'DRIVER={{ODBC Driver 17 for SQL Server}};'
                f'SERVER={server};'
                f'UID={user};'
                f'PWD={password};'
            )
            self.connection = pyodbc.connect(connection_string)
        except pyodbc.Error as e:
            messagebox.showerror("Error de conexión", f"No se pudo conectar al servidor: {e}")
            self.connection = None

    def update_config(self, server, user, password):
        if server not in self.config or self.config[server]['user'] != user or self.config[server]['password'] != password:
            self.config[server] = {'user': user, 'password': password}
            self.save_config()

    def get_databases(self):
        query = "SELECT name FROM sys.databases"
        cursor = self.connection.cursor()
        cursor.execute(query)
        return [row[0] for row in cursor.fetchall()]

    def get_columns(self, database, table):
        query = f"SELECT column_name FROM {database}.information_schema.columns WHERE table_name = '{table}'"
        cursor = self.connection.cursor()
        cursor.execute(query)
        return [row[0] for row in cursor.fetchall()]

    def import_excel(self, file_path):
        self.excel_data = pd.read_excel(file_path)

    def get_next_numCodigoCliente(self, database):
        cursor = self.connection.cursor()
        query = f"SELECT MAX(numCodigoCliente) FROM [{database}].dbo.Personas"
        cursor.execute(query)
        last_num_codigo_cliente = cursor.fetchone()[0] or 0
        return last_num_codigo_cliente + 1
    
    def get_next_numIdProveedor(self, database):
        cursor = self.connection.cursor()
        query = f"SELECT MAX(numIdProveedor) FROM [{database}].dbo.Proveedores"
        cursor.execute(query)
        last_num_id_proveedor = cursor.fetchone()[0] or 0
        return last_num_id_proveedor + 1
    
    def set_nuevas_columnas(self, db_columns, yaml_config):
        self.actualizar_config(db_columns, yaml_config)

    def actualizar_config(self, db_columns, yaml_config):
        with open('defaultvalues.yaml', 'r', encoding='utf-8') as file:
                yaml_data  = yaml.safe_load(file)
            
        config = yaml_data[yaml_config]

        for key in db_columns:
            if key == "numIdPersona" or key == "numIdProveedor": continue
            if key not in config:
                input_dialog = tk.Toplevel(self.root)
                input_dialog.title(f"Nueva columna: {key}")
                input_dialog.iconbitmap('icono.ico')
                # Configurar el tamaño de la ventana de diálogo
                input_dialog.geometry("450x200")
                input_dialog.resizable(False, False)
                
                # Añadir una etiqueta para instrucción
                label = ttk.Label(input_dialog, text=f"¿Que valor tiene '{key}' por defecto?\nEn caso de ser '' no introducir texto\nEn caso de ser NULL introducir NULL\nNo introducir un valor que no permita la columna para no tener errores")
                label.pack(pady=10)
                
                # Campo de entrada
                entry = ttk.Entry(input_dialog, width=30)
                entry.pack(pady=10)
                self.user_input = ''
                def on_ok():
                    self.user_input = entry.get()
                    input_dialog.destroy()

                ok_button = ttk.Button(input_dialog, text="OK", command=on_ok)
                ok_button.pack(pady=10)
                
                # Centrar la ventana de diálogo sobre la ventana principal
                input_dialog.transient(self.root)
                input_dialog.grab_set()
                self.root.wait_window(input_dialog)

                if self.user_input == "null" or self.user_input == "NULL": 
                    self.user_input=None
                config[key] = self.user_input  # Puedes definir un valor por defecto si es necesario

        # Guardar cambios si se realizaron modificaciones
        with open('defaultvalues.yaml', 'w') as file:
            yaml.safe_dump(yaml_data, file)

    def insert_personas(self, database, column_mapping):
        cursor = self.connection.cursor()

        # Crear ventana de progreso
        progress_window = tk.Toplevel(self.root)
        progress_window.iconbitmap('icono.ico')
        progress_window.title("Progreso de Inserción")
        ttk.Label(progress_window, text="Insertando datos, por favor espere...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
        progress_bar.pack(pady=10)
        progress_window.update()

        # Obtener el número total de filas
        total_rows = len(self.excel_data)
        self.all_values = self.load_default_values()
        # Definir valores por defecto para Personas
        default_values = self.all_values['default_values_personas']
        
        # Comprobar si hay nuevas columnas en la tabla
        if len(column_mapping.keys()) != len(default_values):
            self.set_nuevas_columnas(column_mapping.keys(), 'default_values_personas')
            self.all_values = self.load_default_values()
            default_values = self.all_values['default_values_personas']
        # Obtener el número de dígitos de subcuenta desde la base de datos
        cursor.execute(f"SELECT numDigitosSubcuenta FROM [{database}].dbo.ConfiguracionAgencia")
        num_digitos_subcuenta = cursor.fetchone()[0]

        # Obtener el próximo numCodigoCliente
        next_num_codigo_cliente = self.get_next_numCodigoCliente(database)

        # Construir los valores a insertar
        for index, row in self.excel_data.iterrows():
            values = {}
            for db_column in column_mapping.keys():
                if column_mapping[db_column]:
                    # Si la columna está mapeada, purgar el valor del Excel
                    excel_value = row[column_mapping[db_column]]
                    if db_column == "sPais":
                    # Intentar obtener el numIdPais para el país especificado
                        try:
                            cursor.execute(f"SELECT numidpais FROM [{database}].dbo.Paises WHERE strPais LIKE '{excel_value}'")
                            num_id_pais_sii = cursor.fetchone()
                            if num_id_pais_sii:
                                values["numIdPaisSii"] = num_id_pais_sii[0]
                                values[db_column] = excel_value
                            else:
                                values[db_column] = excel_value
                                values["numIdPaisSii"] = default_values["numIdPaisSii"]
                        except pyodbc.Error as e:
                            print(f"Error al obtener numIdPaisSii para {excel_value}: {e}")
                            values[db_column] = default_values[db_column]
                            values["numIdPaisSii"] = default_values["numIdPaisSii"]
                    elif db_column in default_values:
                        expected_type = type(default_values[db_column])
                        values[db_column] = self.purga_valores(excel_value, expected_type)
                    else:
                        # Si la columna no existe en los valores por defecto, continuar
                        print(f"Columna {db_column} no reconocida en Personas, se omite.")
                elif db_column == "numCodigoCliente":
                    values[db_column] = next_num_codigo_cliente
                    next_num_codigo_cliente += 1  # Incrementar para el siguiente registro
                elif db_column == "strSubcuenta":
                    values[db_column] = f"4300{str(values['numCodigoCliente']).zfill(num_digitos_subcuenta - 4)}"
                elif db_column == "sAmexClientNumber":
                    values[db_column] = values["numCodigoCliente"]
                elif db_column == "datAlta" or db_column == "datFechaHoraActualizacion":
                    values[db_column] = pd.Timestamp.today().strftime('%Y%m%d')
                elif db_column in ["sRazonSocialCliente", "sTitularCuentaBancaria", "sPersonaContacto"]:
                    # Usar el nombre del cliente (strTitular)
                    values[db_column] = values.get("strTitular", default_values["strTitular"])
                elif db_column == "strNombre":
                    values[db_column] = values.get("strTitular", default_values["strTitular"])
                elif db_column != "numIdPersona" and db_column != "numIdPaisSii":
                    # Usar valor por defecto para otros campos
                    values[db_column] = default_values.get(db_column, None)

            # Excluir numIdPersona del inserto
            if "numIdPersona" in values:
                del values["numIdPersona"]

            # Insertar el registro
            columns = ', '.join(f'[{col}]' for col in values.keys())
            placeholders = ', '.join('?' for _ in values)
            insert_query = f"INSERT INTO [{database}].dbo.Personas ({columns}) VALUES ({placeholders})"
            
            try:
                cursor.execute(insert_query, list(values.values()))
            except pyodbc.Error as e:
                messagebox.showerror("Error de inserción", f"No se pudo insertar la fila: {index}\n {e}")

            # Actualizar barra de progreso
            progress_bar['value'] = (index + 1) / total_rows * 100
            progress_window.update_idletasks()

        self.connection.commit()
        progress_window.destroy()
        messagebox.showinfo("Éxito", f"Los datos han sido insertados con éxito.\n")

    def on_server_select(self, event):
        selected_server = self.server_combobox.get()
        if selected_server in self.config:
            self.user_entry.delete(0, tk.END)
            self.user_entry.insert(0, self.config[selected_server]['user'])
            self.password_entry.delete(0, tk.END)
            self.password_entry.insert(0, self.config[selected_server]['password'])

    def on_connect(self):
        server = self.server_combobox.get()
        user = self.user_entry.get()
        password = self.password_entry.get()
        self.connect_to_server(server, user, password)

        if self.connection:
            self.update_config(server, user, password)
            db_list = self.get_databases()
            self.database_combobox['values'] = db_list

    def on_import(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        self.import_excel(file_path)

        # Llamar a la función para mapear las columnas
        self.map_columns()

    def insert_proveedores(self, database, column_mapping):
        cursor = self.connection.cursor()

        # Crear ventana de progreso
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Progreso de Inserción")
        ttk.Label(progress_window, text="Insertando datos, por favor espere...").pack(pady=10)
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
        progress_bar.pack(pady=10)
        progress_window.update()

        # Obtener el número total de filas
        total_rows = len(self.excel_data)

        # Definir valores por defecto para Proveedores
        self.all_values = self.load_default_values()
        default_values = self.all_values['default_values_proveedores']
        
        # Comprobar si hay nuevas columnas en la tabla
        if len(column_mapping.keys()) != len(default_values):
            self.set_nuevas_columnas(column_mapping.keys(), 'default_values_proveedores')
            self.all_values = self.load_default_values()
            default_values = self.all_values['default_values_proveedores']

        cursor.execute(f"SELECT numDigitosSubcuenta FROM [{database}].dbo.ConfiguracionAgencia")
        num_digitos_subcuenta = cursor.fetchone()[0]

        next_num_id_prov = self.get_next_numIdProveedor(database)

        # Construir los valores a insertar
        for index, row in self.excel_data.iterrows():
            values = {}
            for db_column in column_mapping.keys():
                if column_mapping[db_column]:
                    # Si la columna está mapeada, purgar el valor del Excel
                    excel_value = row[column_mapping[db_column]]
                    if db_column == "strPais":
                    # Intentar obtener el numIdPais para el país especificado
                        try:
                            cursor.execute(f"SELECT numidpais FROM [{database}].dbo.Paises WHERE strPais LIKE '{excel_value}'")
                            num_id_pais_sii = cursor.fetchone()
                            if num_id_pais_sii:
                                values["numIdPaisSii"] = num_id_pais_sii[0]
                                values[db_column] = excel_value
                            else:
                                values[db_column] = excel_value
                                values["numIdPaisSii"] = default_values["numIdPaisSii"]
                        except pyodbc.Error as e:
                            print(f"Error al obtener numIdPaisSii para {excel_value}: {e}")
                            values[db_column] = default_values[db_column]
                    elif db_column in default_values:
                        expected_type = type(default_values[db_column])
                        values[db_column] = self.purga_valores(excel_value, expected_type)
                    else:
                        # Si la columna no existe en los valores por defecto, continuar
                        print(f"Columna {db_column} no reconocida en Proveedores, se omite.")
                elif db_column == "strSubcuenta":
                    # Construir el valor de subcuenta con el número de dígitos correcto
                    values[db_column] = f"4000{str(next_num_id_prov + index).zfill(num_digitos_subcuenta - 4)}"
                elif db_column == "datFechaHoraActualizacion" or db_column == "datAlta":
                    values[db_column] = pd.Timestamp.today().strftime('%Y%m%d')
                elif db_column != "numIdPaisSii" and db_column != "numIdProveedor":
                    # Usar valor por defecto
                    values[db_column] = default_values.get(db_column, None)

            # Excluir numIdProveedor del inserto
            if "numIdProveedor" in values:
                del values["numIdProveedor"]

            # Insertar el registro
            columns = ', '.join(f'[{col}]' for col in values.keys())
            placeholders = ', '.join('?' for _ in values)
            insert_query = f"INSERT INTO [{database}].dbo.Proveedores ({columns}) VALUES ({placeholders})"
            
            try:
                cursor.execute(insert_query, list(values.values()))
            except pyodbc.Error as e:
                messagebox.showerror("Error de inserción", f"No se pudo insertar la fila: {index}\n {e}")

            # Actualizar barra de progreso
            progress_bar['value'] = (index + 1) / total_rows * 100
            progress_window.update_idletasks()

        self.connection.commit()
        progress_window.destroy()
        messagebox.showinfo("Éxito", "Los datos han sido insertados con éxito.")

    def purga_valores(self, value, expected_type):
        """
        Limpia y valida los valores según el tipo esperado.
        """
        try:
            if expected_type == int:
                return int(value) if value is not None else None
            elif expected_type == float:
                return float(value) if value is not None else None
            elif expected_type == str:
                return str(value).strip() if value is not None else ""
            elif expected_type == bool:
                # Considera valores comunes para booleanos
                return value.lower() in ['true', '1', 'yes'] if isinstance(value, str) else bool(value)
            elif expected_type == 'date':
                # Formatear fechas al formato AAAAMMDD
                return pd.to_datetime(value).strftime('%Y%m%d') if value is not None else None
            else:
                return value
        except Exception as e:
            print(f"Error al purgar valor: {value}, {e}")
            return None

    def confirm_mapping(self, map_window, column_mappings):
        for db_column, combo in column_mappings.items():
            selected_column = combo.get()
            if selected_column != "<Omitir>":
                self.column_mappings[db_column] = selected_column
            else:
                self.column_mappings[db_column] = None  # Usar None para indicar que se omite

        map_window.destroy()
        # Llamar a la función para insertar los datos
        self.on_map_columns()

    def on_map_columns(self):
        server = self.server_combobox.get()
        database = self.database_combobox.get()
        table_type = self.data_type_combobox.get()

        if not self.connection:
            self.connect_to_server(server, self.config[server]['user'], self.config[server]['password'])

        if not self.column_mappings:
            messagebox.showerror("Error", "Debe mapear las columnas antes de insertar los datos.")
            return

        if table_type == "Personas":
            self.insert_personas(database, self.column_mappings)
        elif table_type == "Proveedores":
            self.insert_proveedores(database, self.column_mappings)

    def setup_ui(self):
        # Configura el tamaño de la ventana
        self.root.geometry("300x400")

        # Configura la expansión para que ocupe todo el espacio disponible
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Aplicar un estilo moderno usando ttk
        # style = ttk.Style(self.root)
        # style.theme_use("clam") 

        style = ttk.Style()
        self.set_theme(style)

        # Crear un marco para contener los widgets con padding para un mejor espaciado
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # Configurar la rejilla de main_frame para que los widgets se expandan correctamente
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Configurar filas del marco para que los widgets se expandan verticalmente
        for i in range(7):  # Número de filas de widgets
            main_frame.rowconfigure(i, weight=1)

        # Etiqueta y combobox para el servidor
        ttk.Label(main_frame, text="Servidor:").grid(row=0, column=0, sticky="ew", pady=5)
        self.server_combobox = ttk.Combobox(main_frame)
        self.server_combobox.grid(row=0, column=1, sticky="ew", pady=5)
        self.server_combobox['values'] = list(self.config.keys())
        self.server_combobox.bind("<<ComboboxSelected>>", self.on_server_select)

        # Etiqueta y campo de entrada para el usuario
        ttk.Label(main_frame, text="Usuario:").grid(row=1, column=0, sticky="ew", pady=5)
        self.user_entry = ttk.Entry(main_frame)
        self.user_entry.grid(row=1, column=1, sticky="ew", pady=5)

        # Etiqueta y campo de entrada para la contraseña
        ttk.Label(main_frame, text="Contraseña:").grid(row=2, column=0, sticky="ew", pady=5)
        self.password_entry = ttk.Entry(main_frame, show='*')
        self.password_entry.grid(row=2, column=1, sticky="ew", pady=5)

        # Botón para conectar
        connect_button = ttk.Button(main_frame, text="Conectar", command=self.on_connect)
        connect_button.grid(row=3, column=0, columnspan=2, pady=10, sticky="ew")

        # Etiqueta y combobox para la base de datos
        ttk.Label(main_frame, text="Base de datos:").grid(row=4, column=0, sticky="ew", pady=5)
        self.database_combobox = ttk.Combobox(main_frame)
        self.database_combobox.grid(row=4, column=1, sticky="ew", pady=5)

        # Etiqueta y combobox para el tipo de datos
        ttk.Label(main_frame, text="Tipo de datos:").grid(row=5, column=0, sticky="ew", pady=5)
        self.data_type_combobox = ttk.Combobox(main_frame, values=["Personas", "Proveedores"])
        self.data_type_combobox.grid(row=5, column=1, sticky="ew", pady=5)

        # Botón para importar Excel
        import_button = ttk.Button(main_frame, text="Importar Excel", command=self.on_import)
        import_button.grid(row=6, column=0, columnspan=2, pady=10, sticky="ew")

    def set_theme(self, style):
        """
        Configura un estilo moderno para los widgets de ttk.
        """
        style.theme_use('clam')  # Usar 'clam' como tema base

        # Configuración de estilo para botones
        style.configure('TButton',
                        font=('Segoe UI', 9, 'bold'),  # Tamaño de fuente reducido
                        foreground='#ffffff',
                        background='#0078D7',
                        borderwidth=0,
                        focuscolor='none',
                        padding=5)  # Acolchado reducido
        style.map('TButton',
                background=[('active', '#005A9E'), ('pressed', '#004578')],
                foreground=[('disabled', '#888888')])

        # Configuración de estilo para cuadros de texto
        style.configure('TEntry',
                        font=('Segoe UI', 10),
                        foreground='#333333',
                        fieldbackground='#ffffff',
                        bordercolor='#cccccc',
                        borderwidth=1,
                        focusthickness=2,
                        focuscolor='#0078D7',
                        padding=5)
        style.map('TEntry',
                fieldbackground=[('readonly', '#f7f7f7')],
                foreground=[('readonly', '#888888')])

        # Configuración de estilo para etiquetas
        style.configure('TLabel',
                        font=('Segoe UI', 10),
                        foreground='#333333',
                        background='#f7f7f7')

        # Configuración de estilo para Combobox
        style.configure('TCombobox',
                        font=('Segoe UI', 10),
                        foreground='#333333',
                        fieldbackground='#ffffff',
                        bordercolor='#cccccc',
                        borderwidth=1,
                        padding=5)
        style.map('TCombobox',
                fieldbackground=[('readonly', '#f7f7f7')],
                foreground=[('readonly', '#888888')],
                arrowcolor=[('active', '#0078D7')])

        # Configuración de estilo para Frames
        style.configure('TFrame',
                        background='#f0f0f0')

        # Configuración de estilo para etiquetas y combobox seleccionados
        style.configure('Highlighted.TLabel',
                        font=('Segoe UI', 10, 'bold'),
                        foreground='#ffffff',
                        background='#0078D7')

        style.configure('Highlighted.TCombobox',
                        font=('Segoe UI', 10, 'bold'),
                        foreground='#ffffff',
                        fieldbackground='#0078D7',
                        bordercolor='#005A9E',
                        arrowcolor='#ffffff')

    def map_columns(self):
        table = self.data_type_combobox.get()
        database = self.database_combobox.get()

        if not self.connection:
            messagebox.showerror("Error", "No se ha conectado a la base de datos.")
            return

        db_columns = self.get_columns(database, table)
        excel_columns = self.excel_data.columns.tolist()

        map_window = tk.Toplevel(self.root)
        map_window.iconbitmap('icono.ico')
        map_window.title("Mapear Columnas")
        map_window.geometry("500x600")
        map_window.resizable(False, False)
        map_window.minsize(500, 500)

        # Main frame for organizing components
        main_frame = ttk.Frame(map_window)
        main_frame.pack(fill="both", expand=True)

        # Frame for canvas with scrollbar
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.pack(side="top", fill="both", expand=True)

        # Creating canvas and scrollbar
        column_canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=column_canvas.yview)

        # Create a frame within the canvas
        scrollable_frame = ttk.Frame(column_canvas)

        # Create window inside the canvas
        window_id = column_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # Configure the canvas and scroll
        column_canvas.configure(yscrollcommand=scrollbar.set)

        # Ensure the scrollable_frame takes full width of the canvas
        def on_frame_configure(event=None):
            column_canvas.configure(scrollregion=column_canvas.bbox("all"))

        def on_canvas_resize(event=None):
            # Set canvas width to the width of the canvas_frame minus the scrollbar
            canvas_width = event.width - scrollbar.winfo_width()
            column_canvas.itemconfig(window_id, width=canvas_width)

        scrollable_frame.bind("<Configure>", on_frame_configure)
        canvas_frame.bind("<Configure>", on_canvas_resize)

        # Pack the canvas and scrollbar
        column_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Mapping columns
        column_mappings = {}

        for i, db_column in enumerate(db_columns):
            label = ttk.Label(scrollable_frame, text=db_column)
            label.grid(row=i, column=0, padx=5, pady=5, sticky="ew")
            
            options = ["<Omitir>"] + excel_columns
            combo = ttk.Combobox(scrollable_frame, values=options, width=30)

            best_match = "<Omitir>"
            highest_similarity = 0.8

            for excel_column in excel_columns:
                similarity = difflib.SequenceMatcher(None, db_column.lower(), excel_column.lower()).ratio()
                if similarity > highest_similarity:
                    highest_similarity = similarity
                    best_match = excel_column

            combo.set(best_match)
            combo.grid(row=i, column=1, padx=(5, 15), pady=5, sticky="w")

            def on_combobox_change(event=None, label=label, combo=combo):
                if combo.get() != "<Omitir>":
                    label.configure(style='Highlighted.TLabel')
                    combo.configure(style='Highlighted.TCombobox')
                else:
                    label.configure(style='TLabel')
                    combo.configure(style='TCombobox')

            # Check the initial state of the combobox
            on_combobox_change()

            combo.bind("<<ComboboxSelected>>", on_combobox_change)
            combo.bind("<MouseWheel>", lambda event: "break")
            column_mappings[db_column] = combo

        # Make the columns in the grid expand with specific proportions
        scrollable_frame.columnconfigure(0, weight=4)  # Label occupies 80% of the space
        scrollable_frame.columnconfigure(1, weight=1)  # Combobox occupies 20% of the space
        scrollable_frame.columnconfigure(2, weight=0, minsize=10)  # Add extra margin

        # Canvas for search at the bottom
        search_canvas = ttk.Frame(main_frame)
        search_canvas.pack(side="bottom", fill="x", pady=5)

        search_label = ttk.Label(search_canvas, text="Buscar columna:")
        search_label.pack(side="left", padx=5)
        search_entry = ttk.Entry(search_canvas)
        search_entry.pack(side="left", fill="x", expand=True, padx=5)

        # Search buttons
        prev_button = ttk.Button(search_canvas, text="Anterior", command=lambda: previous_match())
        prev_button.pack(side="left", padx=5)

        next_button = ttk.Button(search_canvas, text="Siguiente", command=lambda: next_match())
        next_button.pack(side="left", padx=5)

        # Canvas for confirming at the bottom
        confirm_canvas = ttk.Frame(main_frame)
        confirm_canvas.pack(side="bottom", fill="x", pady=5)

        confirm_button = ttk.Button(confirm_canvas, text="Confirmar Mapeo", command=lambda: self.confirm_mapping(map_window, column_mappings))
        confirm_button.pack(pady=10)

        # Search logic
        matching_indices = []
        current_match_index = 0

        def search_column(event=None):
            nonlocal current_match_index
            query = search_entry.get().strip().lower()
            if not query:
                messagebox.showwarning("Error", "Por favor, ingrese un término de búsqueda.")
                return

            matching_indices.clear()

            for idx, col in enumerate(db_columns):
                if query in col.lower():
                    matching_indices.append(idx)

            if not matching_indices:
                messagebox.showwarning("No encontrado", f"No se encontraron columnas que coincidan con '{query}'.")
                return

            # Move to the first match when a new search is done
            current_match_index = 0
            go_to_match(matching_indices[current_match_index])

        def go_to_match(index):
            position = index / len(db_columns)
            column_canvas.yview_moveto(position)

        def next_match(event=None):
            nonlocal current_match_index
            if not matching_indices:
                search_column()
            else:
                current_match_index = (current_match_index + 1) % len(matching_indices)
                go_to_match(matching_indices[current_match_index])

        def previous_match():
            nonlocal current_match_index
            if not matching_indices:
                search_column()
            else:
                current_match_index = (current_match_index - 1) % len(matching_indices)
                go_to_match(matching_indices[current_match_index])

        # Bind the Enter key to go to the next match
        search_entry.bind("<Return>", next_match)

        # Mouse wheel event for scroll
        def on_mouse_wheel(event):
            column_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        map_window.bind_all("<MouseWheel>", on_mouse_wheel)

if __name__ == "__main__":
    ExcelToSQLImporter()

