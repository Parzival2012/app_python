import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import openpyxl


class ExcelDataApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Guardar datos en Excel')
        self.root.config(bg='black')
        self.root.geometry('560x380')
        self.root.resizable(0, 0)

        # Estructura de datos
        self.datos = {
            'nombres': [],
            'apellidos': [],
            'edades': [],
            'correos': [],
            'telefonos': []
        }

        self.setup_ui()

    def setup_ui(self):

        # Frame principal
        self.main_frame = tk.Frame(self.root, bg='gray15')
        self.main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Frame de entrada de datos
        self.data_frame = tk.Frame(self.main_frame, bg='gray15')
        self.data_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))

        # Frame de control
        self.control_frame = tk.Frame(self.main_frame, bg='gray16')
        self.control_frame.pack(side='right', fill='both', expand=True)

        self.create_data_widgets()
        self.create_control_widgets()

    def create_data_widgets(self):

        # Configurar grid
        self.data_frame.columnconfigure(1, weight=1)

        # Campos de entrada
        campos = [
            ('Nombre', 'nombre'),
            ('Apellido', 'apellido'),
            ('Edad', 'edad'),
            ('Correo', 'correo'),
            ('Teléfono', 'telefono')
        ]

        self.entries = {}
        for i, (label_text, key) in enumerate(campos):
            # Label
            label = tk.Label(
                self.data_frame,
                text=label_text,
                width=10,
                bg='gray15',
                fg='white'
            )
            label.grid(column=0, row=i, pady=10, padx=5, sticky='e')

            # Entry
            entry = tk.Entry(
                self.data_frame,
                width=20,
                font=('Arial', 12)
            )
            entry.grid(column=1, row=i, pady=10, padx=5, sticky='ew')
            self.entries[key] = entry

        # Botón Agregar
        self.btn_agregar = tk.Button(
            self.data_frame,
            text='Agregar',
            width=20,
            font=('Arial', 12, 'bold'),
            bg='orange',
            bd=5,
            command=self.agregar_datos
        )
        self.btn_agregar.grid(column=0, columnspan=2, row=5, pady=20)

    def create_control_widgets(self):

        # Label para nombre de archivo
        lbl_archivo = tk.Label(
            self.control_frame,
            text='Nombre del archivo:',
            bg='gray16',
            fg='white',
            font=('Arial', 12, 'bold')
        )
        lbl_archivo.pack(pady=(20, 5))

        # Entry para nombre de archivo
        self.entry_archivo = tk.Entry(
            self.control_frame,
            width=23,
            font=('Arial', 12),
            highlightbackground='green',
            highlightthickness=4
        )
        self.entry_archivo.pack(pady=5)

        # Botón Guardar
        self.btn_guardar = tk.Button(
            self.control_frame,
            text='Guardar en Excel',
            width=20,
            font=('Arial', 12, 'bold'),
            bg='green2',
            bd=5,
            command=self.guardar_datos
        )
        self.btn_guardar.pack(pady=20)

        # Listbox para ver datos ingresados
        self.lista_datos = tk.Listbox(
            self.control_frame,
            width=30,
            height=10,
            bg='white',
            fg='black'
        )
        self.lista_datos.pack(pady=10, fill='both', expand=True)

    def validar_datos(self):

        nombre = self.entries['nombre'].get().strip()
        apellido = self.entries['apellido'].get().strip()

        if not nombre or not apellido:
            messagebox.showerror("Error", "Nombre y Apellido son obligatorios")
            return False

        # Validar edad numérica
        edad = self.entries['edad'].get().strip()
        if edad and not edad.isdigit():
            messagebox.showerror("Error", "La edad debe ser un número")
            return False

        return True

    def limpiar_campos(self):

        for entry in self.entries.values():
            entry.delete(0, tk.END)

    def agregar_datos(self):

        if not self.validar_datos():
            return

        # Recoger datos
        self.datos['nombres'].append(self.entries['nombre'].get().strip())
        self.datos['apellidos'].append(self.entries['apellido'].get().strip())
        self.datos['edades'].append(self.entries['edad'].get().strip() or 'N/A')
        self.datos['correos'].append(self.entries['correo'].get().strip() or 'N/A')
        self.datos['telefonos'].append(self.entries['telefono'].get().strip() or 'N/A')

        # Actualizar listbox
        self.lista_datos.insert(
            tk.END,
            f"{self.datos['nombres'][-1]} {self.datos['apellidos'][-1]}"
        )

        self.limpiar_campos()
        messagebox.showinfo("Éxito", "Datos agregados correctamente")

    def guardar_datos(self):

        if not any(self.datos.values()):
            messagebox.showerror("Error", "No hay datos para guardar")
            return

        nombre_archivo = self.entry_archivo.get().strip()
        if not nombre_archivo:
            messagebox.showerror("Error", "Ingrese un nombre para el archivo")
            return

        try:
            # Crear DataFrame
            df = pd.DataFrame(self.datos)

            # Generar nombre de archivo con timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            nombre_completo = f"{nombre_archivo}_{timestamp}.xlsx"

            # Guardar
            df.to_excel(nombre_completo, index=False)

            messagebox.showinfo(
                "Éxito",
                f"Datos guardados en: {nombre_completo}\n"
                f"Total de registros: {len(self.datos['nombres'])}"
            )

            # Limpiar después de guardar
            self.datos = {key: [] for key in self.datos}
            self.lista_datos.delete(0, tk.END)
            self.entry_archivo.delete(0, tk.END)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {str(e)}")


def main():
    root = tk.Tk()
    app = ExcelDataApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()