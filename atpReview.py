

from pandas import read_excel, to_datetime, isna
from matplotlib.pyplot import subplots
from tkinter import Tk, Frame, Button, Label, Entry
from tkinter import TOP, X, BOTTOM, BOTH, END
from tkinter import BooleanVar, filedialog, ttk, messagebox
from tkinter.ttk import Combobox, Checkbutton, Treeview, Scrollbar
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from matplotlib.dates import num2date
from datetime import datetime
from sys import exit
from numpy import datetime64
from PIL import Image, ImageTk  # Para icono en PNG
from os import getcwd, path

class ExcelApp:
    def __init__(self):
        self.root = Tk()
        self.root.title("ATP Reviewer")
        self.root.geometry("900x900")
        self.lines = {}
        self.cursors = {}
        # Variables
        self.data = None
        self.columns = []
        self.plots = []
        self.current_plot_index = 0
        self.time_markers_enabled = BooleanVar(value=False)
        self.marker1 = None
        self.marker2 = None
        self.time_diff_label = None

        # Secciones de la interfaz
        # Lista de marcadores
        self.marcadores = ['+', 'x', '*', '^', 'o', 's', 'D',  'v', '<', '>', 'p', 'h',  ]
        self.marcador_index = 0  # Índice inicial de marcador

        # Variable to store difference plot
        self.difference_line = None

        self.create_widgets()

    def create_widgets(self):

        # ancho común para todos los botones
        button_width = 20  # 

        # Frame de control
        control_frame = Frame(self.root)
        control_frame.pack(side=TOP, fill=X, padx=10, pady=10)

        # Botón para abrir archivo Excel
        self.open_button = Button(control_frame, text="Abrir archivo Excel", command=self.open_excel, width=button_width)
        #self.open_button.pack(pady=5)
        self.open_button.grid(row=0, column=0, padx=5)

        self.label_filename = Label(control_frame, text="                 ")
        self.label_filename.grid(row=0, column=1, padx=5)

         # Add time markers checkbox
        self.time_markers_check = Checkbutton(
            control_frame, 
            text="Mostrar marcadores", 
            variable=self.time_markers_enabled, 
            command=self.toggle_time_markers
        )
        self.time_markers_check.grid(row=5, column=1, padx=5)

        # Label to show time difference
        self.time_diff_label = Label(control_frame, text="Distancia entre marcas: ")
        self.time_diff_label.grid(row=5, column=2, padx=5)

        
        # Combobox para seleccionar columna de fecha
        self.label_ejeX = Label(control_frame, text="Eje X (tiempo)")
        self.label_ejeX.grid(row=0, column=2, padx=5)
        self.combo_fecha = Combobox(control_frame, state="readonly")
        self.combo_fecha.grid(row=1, column=2, padx=5)

        # Combobox para seleccionar columna de valores
        self.label_ejeY = Label(control_frame, text="Eje Y (gráfico) ")
        self.label_ejeY.grid(row=2, column=2, padx=5)
        self.combo_valor = Combobox(control_frame, state="readonly")
        #self.combo_valor.pack(pady=5)
        self.combo_valor.grid(row=3, column=2, padx=5)

        # Entrada para rango de fechas
        self.label_fecha_inicio = Label(control_frame, text="Hora de inicio (YYYY-MM-DD HH:MM:SS)")
        self.label_fecha_inicio.grid(row=0, column=3, padx=5)
        self.entry_fecha_inicio = Entry(control_frame)
        self.entry_fecha_inicio.grid(row=1, column=3, padx=5)

        self.label_fecha_fin = Label(control_frame, text="Hora de fin (YYYY-MM-DD HH:MM:SS)")
        self.label_fecha_fin.grid(row=2, column=3, padx=5)
        self.entry_fecha_fin = Entry(control_frame)
        self.entry_fecha_fin.grid(row=3, column=3, padx=5)

        self.grid_enable = BooleanVar(value=True)  # Inicialmente activado
        self.grid_check = Checkbutton(control_frame, text="Cuadrícula", variable=self.grid_enable) #,command=self.actualizar_grafica)
        self.grid_check.grid(row=5, column=0, padx=5)

        # Botón para graficar datos
        self.graph_button = Button(control_frame, text="Agregar gráfico", command=self.add_plot, width=button_width)
        self.graph_button.grid(row=1, column=4, padx=5)
        # Botón para eliminar gráficos seleccionados
        self.delete_button = Button(control_frame, text="Eliminar gráfico", command=self.delete_plot, width=button_width)
        self.delete_button.grid(row=2, column=4, padx=5)
        # Botón para eliminar todos los gráficos
        self.delete_all_button = Button(control_frame, text="Eliminar todos", command=self.delete_all_plots, width=button_width)
        self.delete_all_button.grid(row=3, column=4, padx=5)

        # Botón para gráficos predefinidos
        self.predefined_graphs_button = Button(control_frame, text="Gráficos Predefinidos", command=self.add_predefined_graphs, width=button_width)
        self.predefined_graphs_button.grid(row=2, column=0, padx=5)

        # Add a button for plot difference calculation  
        self.difference_button = Button(control_frame, text="Calcular Diferencia", command=self.calculate_plot_difference, width=button_width)
        self.difference_button.grid(row=3, column=0, padx=5)

        # Botón para eliminar gráfico de diferencia
        self.delete_difference_button = Button(control_frame, text="Eliminar Diferencia", command=self.delete_plot_difference, width=button_width)
        self.delete_difference_button.grid(row=3, column=1, padx=5)

        # Tabla para mostrar los datos del Excel (omitida)
        self.table_frame = Frame(self.root)
        
        self.data_table = Treeview(self.table_frame, columns=(), show='headings')
        self.vsb = Scrollbar(self.table_frame, orient="vertical", command=self.data_table.yview)
        self.hsb = Scrollbar(self.table_frame, orient="horizontal", command=self.data_table.xview)
        self.data_table.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        #self.data_table.pack(pady=1, fill='x')
        #self.hsb.pack(pady=1, fill='x')
        
        self.data_table.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")

        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

        #self.table_frame.pack(fill=X, padx=2, pady=5)

        # Lienzo para mostrar gráficos
        self.figure, self.ax = subplots(figsize=(8, 4))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=True)

        self.plot_toolbar = NavigationToolbar2Tk(self.canvas, self.root)
        self.plot_toolbar.update()
        # Colocar la barra de herramientas en la ventana
        self.canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=True)


        

        # Empaquetar la barra de herramientas abajo del gráfico
        self.plot_toolbar.pack(side=BOTTOM, fill='x')

    def open_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                #print(file_path.split("/")[-1])
                self.data = read_excel(file_path, engine='openpyxl')
                self.columns = self.data.columns.tolist()
                self.combo_fecha['values'] = self.columns
                self.combo_valor['values'] = self.columns
                self.data['Timestamp'] = self.data['Timestamp'].dt.tz_localize(None)
                self.filename = file_path.split("/")[-1]
                self.label_filename.config(text = self.filename)
                #print(self.data['Timestamp'].iloc[0])     
                messagebox.showinfo("Éxito", "Archivo cargado correctamente.")
                self.entry_fecha_inicio.delete(0, END)
                self.entry_fecha_fin.delete(0, END)
                self.entry_fecha_inicio.insert(0, str(self.data['Timestamp'].iloc[0]))
                self.entry_fecha_fin.insert(0, str(self.data['Timestamp'].iloc[-1]))
                self.delete_all_plots()

            except Exception as e:
                messagebox.showerror("Error", f"Error al cargar archivo: {e}")
        
        

    def update_table(self):
        # Limpiar la tabla existente
        self.data_table.delete(*self.data_table.get_children())
        # Configurar encabezados
        self.data_table['columns'] = self.columns
        for col in self.columns:
            self.data_table.heading(col, text=col)
        # Insertar datos
        for _, row in self.data.iterrows():
            row.fillna("", inplace=True)
            self.data_table.insert('', 'end', values=list(row))

    def add_plot(self):
        self.update_table()
        if self.data is not None:
            try:
                # Obtener columnas seleccionadas
                col_fecha = self.combo_fecha.get()
                col_valor = self.combo_valor.get()
                
                if not col_fecha or not col_valor:
                    messagebox.showerror("Error", "Debe seleccionar las columnas de Fecha y Valor.")
                    return

                # Convertir columna de fecha a formato datetime
                self.data[col_fecha] = to_datetime(self.data[col_fecha], errors='coerce')
                
                # Filtrar por rango de fechas
                fecha_inicio = self.entry_fecha_inicio.get()
                fecha_fin = self.entry_fecha_fin.get()

                filtered_data = self.data.copy()
                if fecha_inicio and fecha_fin:
                    fecha_inicio = datetime.strptime(fecha_inicio, "%Y-%m-%d %H:%M:%S")
                    fecha_fin = datetime.strptime(fecha_fin, "%Y-%m-%d %H:%M:%S")
                    filtered_data = filtered_data[(filtered_data[col_fecha] >= fecha_inicio) & (filtered_data[col_fecha] <= fecha_fin)]

                # Graficar en el mismo gráfico
                if col_valor == "Current_Speed(MPerSec)" or col_valor == "Permitted_Speed(MPerSec)" or col_valor == "Current_Odometer(CM)" or col_valor == "ALU_Alive_Byte" :
                    line, = self.ax.plot(filtered_data[col_fecha], filtered_data[col_valor], label=col_valor)
                else:
                    marcador = self.marcadores[self.marcador_index % len(self.marcadores)]
                    line, = self.ax.plot(filtered_data[col_fecha], filtered_data[col_valor], marker=marcador, label=col_valor)
                    self.marcador_index += 1
                #line, = self.ax.plot(filtered_data[col_fecha], filtered_data[col_valor], label=col_valor)
                self.ax.legend(loc='upper left')
                self.lines[col_valor] = line
                #print("Added : ", col_valor )

                #self.cursors[col_valor] = mplcursors.cursor(self.ax, hover=True)
                self.ax.grid(self.grid_enable)
                # Actualizar la interfaz
                self.canvas.draw()

                
            except Exception as e:
                messagebox.showerror("Error", f"Error al graficar los datos: {e}")
        else:
            messagebox.showerror("Error", "No se ha cargado ningún archivo.")

    def toggle_time_markers(self):
        if self.time_markers_enabled.get():
            # Connect click event to add time markers
            self.canvas.mpl_connect('button_press_event', self.add_time_marker)
        else:
            # Remove markers and disconnect event
            if self.marker1:
                self.marker1.remove()
                self.marker1 = None
            if self.marker2:
                self.marker2.remove()
                self.marker2 = None
            self.time_diff_label.config(text="Distancia entre marcas: ")
            self.canvas.draw()
            self.canvas.mpl_disconnect(self.canvas.mpl_connect('button_press_event', self.add_time_marker))

    def add_time_marker(self, event):
        # Only add markers if the click is on the axes
        if event.inaxes != self.ax:
            return

        # If no first marker, add first marker
        if not self.marker1:
            self.marker1 = self.ax.axvline(x=event.xdata, color='r', linestyle='--')
            self.canvas.draw()
            return

        # If first marker exists but second doesn't, add second marker and calculate difference
        if self.marker1 and not self.marker2:
            self.marker2 = self.ax.axvline(x=event.xdata, color='g', linestyle='--')
            
            # Calculate time difference
            try:
                # Find the nearest points to marker locations
                marker1_time = self.convert_xdata_to_distance(self.marker1.get_xdata()[0])
                marker2_time = self.convert_xdata_to_distance(self.marker2.get_xdata()[0])
                
                # Calculate time difference
                distance_diff = abs(marker2_time - marker1_time)
                
                # Update label with time difference
                self.time_diff_label.config(text=f" Δ Distancia: {distance_diff:.2f} metros")
            except Exception as e:
                messagebox.showerror("Error", f"Error calculando distancia según tiempo: {e}")
            
            self.canvas.draw()

    def add_predefined_graphs(self):
        """
        Automatically add two predefined graphs:
        1. Current Speed 
        2. Permitted Speed
        """
        if self.data is None:
            messagebox.showerror("Error", "No se ha cargado ningún archivo.")
            return

        try:
            # Set predefined columns
            self.combo_fecha.set('Timestamp')
            
            # First graph - Current Speed
            self.combo_valor.set('Current_Speed(MPerSec)')
            self.add_plot()

            # Second graph - Permitted Speed
            self.combo_valor.set('Permitted_Speed(MPerSec)')
            self.add_plot()

            # Reset combo box to avoid confusion
            self.combo_valor.set('')
            
            messagebox.showinfo("Éxito", "Gráficos predefinidos añadidos.")

        except Exception as e:
            messagebox.showerror("Error", f"Error al añadir gráficos predefinidos: {e}")

    def convert_xdata_to_distance(self, x_value):
        if self.data is not None:
            col_fecha = self.combo_fecha.get()
            x_value_2 = datetime64(num2date(x_value), 'ns')
            # Find the nearest index
            nearest_index = (self.data[col_fecha] - x_value_2).abs().argmin()
            # Search backwards for a valid odometer value
            for i in range(nearest_index, -1, -1):
                odometer_value = self.data["Current_Odometer(CM)"].iloc[i]
                if not isna(odometer_value):
                    return abs(odometer_value) / 100   
        return None

    def delete_plot(self):
        col_valor = self.combo_valor.get()
        self.lines[col_valor].remove()
        del self.lines[col_valor]
        self.ax.legend()
        self.canvas.draw()

    def delete_all_plots(self):
        #Eliminar todos los gráficos del gráfico actual.
        try:
            if self.lines:
                for line in self.lines:
                    self.lines[line].remove()
            self.ax.clear()
            self.canvas.draw()
        except: 
            self.root.destroy()
            main()

    def calculate_plot_difference(self):
        if len(self.lines) < 2:
            messagebox.showerror("Error", "Se necesitan al menos dos gráficas para calcular la diferencia.")
            return

        try:
            # Get the names of the two most recently added lines
            line_names = list(self.lines.keys())[-2:]
            
            # Get the data for these lines
            col_fecha = self.combo_fecha.get()
            
            # Filter data based on current date range
            fecha_inicio = self.entry_fecha_inicio.get()
            fecha_fin = self.entry_fecha_fin.get()
            
            filtered_data = self.data.copy()
            if fecha_inicio and fecha_fin:
                fecha_inicio = datetime.strptime(fecha_inicio, "%Y-%m-%d %H:%M:%S")
                fecha_fin = datetime.strptime(fecha_fin, "%Y-%m-%d %H:%M:%S")
                filtered_data = filtered_data[(filtered_data[col_fecha] >= fecha_inicio) & (filtered_data[col_fecha] <= fecha_fin)]
            
            # Calculate difference
            #diff_values = abs(filtered_data[line_names[0]] - filtered_data[line_names[1]])
            diff_values = filtered_data[line_names[0]] - filtered_data[line_names[1]]
            # Plot difference
            if self.difference_line:
                self.difference_line.remove()
            
            diff_label = f"Diferencia ({line_names[0]} - {line_names[1]})"
            self.difference_line, = self.ax.plot(filtered_data[col_fecha], diff_values, 
                                                 label=diff_label, 
                                                 linestyle='--', 
                                                 color='purple')
            
            self.ax.legend(loc='upper left')
            self.canvas.draw()
            
            #messagebox.showinfo("Éxito", f"Diferencia calculada entre {line_names[0]} y {line_names[1]}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al calcular diferencia: {e}")

    def delete_plot_difference(self):
        """
        Remove the difference plot if it exists
        """
        try:
            if self.difference_line:
                self.difference_line.remove()
                self.difference_line = None
                self.ax.legend(loc='upper left')
                self.canvas.draw()
            else:
                messagebox.showinfo("Información", "No hay gráfico de diferencia para eliminar.")
        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar gráfico de diferencia: {e}")


def main():
    app = ExcelApp()
    def close_window():
        app.root.quit()  # Cierra el loop de la GUI
        app.root.destroy()  # Destruye la ventana
        exit() 
    try:
        iconPath = path.join(getcwd() , "_internal/SopIcon2.ico")
        app.root.iconbitmap(iconPath)
    except:
        try:
            iconPath = path.join(getcwd() , "SopIcon2.ico")
            app.root.iconbitmap(iconPath)
        except:
            messagebox.showerror("Error", "No se encontró: ", iconPath)

    app.root.protocol("WM_DELETE_WINDOW", close_window)
    
    app.root.mainloop()

    

if __name__ == "__main__":
    main()
    


#EXEC : pyinstaller --onefile --windowed --clean --strip --noconsole --exclude-module tensorflow --exclude-module scipy --icon="SopIcon2.ico"  atpReview.py 