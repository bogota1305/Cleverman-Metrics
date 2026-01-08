
from tkinter import filedialog, messagebox, simpledialog, Label, Button, StringVar
import pandas as pd
import tkinter as tk

def seleccionar_archivos_para_casos():
    # Diccionario para almacenar las rutas de los archivos seleccionados
    archivos_seleccionados = {
        "Customized Kit - Funnel": None,
        "All In One - Funnel": None,
        "Shop - Funnel": None,
        "My Account - Funnel": None,
        "Buy Again - Funnel": None,
        "Buy Again Reactivate - Funnel": None,
        "Buy Again Without Sub - Funnel": None,
        "My Subscriptions - Funnel": None,
        "NPD account - Funnel": None,
        "NPD mail - Funnel": None,
    }

    def seleccionar_archivo(caso, label):
        archivo = filedialog.askopenfilename(title=f"Seleccionar: {caso}",
                                             filetypes=[("CSV files", "*.csv")])
        if archivo:
            archivos_seleccionados[caso] = archivo
            label.config(text=f"Seleccionado: {archivo}")

    # Crear ventana principal
    root = tk.Tk()
    root.title("Seleccionar archivos para cada caso")

    # Crear botones y etiquetas para cada caso
    for caso in archivos_seleccionados.keys():
        frame = tk.Frame(root)
        frame.pack(pady=5, padx=10, anchor="w")
        
        label = Label(frame, text=f"Selecciona: {caso}", wraplength=500, justify="left")
        label.pack(side="left", padx=10)
        
        boton = Button(frame, text=f"Seleccionar archivo para {caso}",
                       command=lambda c=caso, l=label: seleccionar_archivo(c, l))
        boton.pack(side="right")

    # Espaciado entre secciones
    tk.Label(root, text="").pack()

    # Botón para confirmar selección
    confirmar = Button(root, text="Confirmar selección", command=root.quit)
    confirmar.pack(pady=20)

    # Mostrar ventana
    root.mainloop()

    # Cerrar ventana
    root.destroy()

    # Retornar los archivos seleccionados y el mes elegido
    return archivos_seleccionados

def seleccionar_archivos_stripe():
    # Diccionario para almacenar las rutas de los archivos seleccionados
    archivos_seleccionados = {
        "Blocked Payments": None,
        "All Payments": None,
    }

    def seleccionar_archivo(caso, label):
        archivo = filedialog.askopenfilename(title=f"Seleccionar: {caso}",
                                             filetypes=[("CSV files", "*.csv")])
        if archivo:
            archivos_seleccionados[caso] = archivo
            label.config(text=f"Seleccionado: {archivo}")

    # Crear ventana principal
    root = tk.Tk()
    root.title("Seleccionar archivo")
    
    # Crear botones y etiquetas para cada caso
    for caso in archivos_seleccionados.keys():
        frame = tk.Frame(root)
        frame.pack(pady=5, padx=10, anchor="w")
        
        label = Label(frame, text=f"Selecciona: {caso}", wraplength=500, justify="left")
        label.pack(side="left", padx=10)
        
        boton = Button(frame, text=f"Seleccionar archivo para {caso}",
                       command=lambda c=caso, l=label: seleccionar_archivo(c, l))
        boton.pack(side="right")

    # Espaciado entre secciones
    tk.Label(root, text="").pack()

    # Botón para confirmar selección
    confirmar = Button(root, text="Confirmar selección", command=root.quit)
    confirmar.pack(pady=20)

    # Mostrar ventana
    root.mainloop()

    # Cerrar ventana
    root.destroy()

    # Retornar los archivos seleccionados y el mes elegido
    return archivos_seleccionados
