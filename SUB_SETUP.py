import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

def seleccionar_archivo():
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo",
        filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
    )

    if not ruta:
        return

    try:
        # -------------------------------------------------
        # 1️ Leer el archivo línea por línea
        # -------------------------------------------------
        with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
            lineas = f.readlines()

        # Buscar JobFolder y JobName en la primera línea de datos (línea 2, índice 1)
        if len(lineas) > 1:
            primer_dato = lineas[1].strip().split(",")
            JobFolder = primer_dato[0].strip('"\'\\')
            JobName = primer_dato[1].strip('"\'\\') if len(primer_dato) > 1 else "No encontrado"
        else:
            JobFolder = "No encontrado"
            JobName = "No encontrado"
        
        label_JobFolder.config(text=f"JobFolder: {JobFolder}")

        # -------------------------------------------------
        # 2️ Buscar la fila de encabezados con las columnas deseadas
        # -------------------------------------------------
        encabezados_idx = -1
        for i, linea in enumerate(lineas):
            if "ModuleNumber" in linea and "PartNumber" in linea and "Location" in linea and "FeederID" in linea:
                encabezados_idx = i
                break

        if encabezados_idx == -1:
            messagebox.showerror("Error", "No se encontraron las columnas necesarias.")
            return

        # Parsear los encabezados
        encabezados = [col.strip().strip('"\'') for col in lineas[encabezados_idx].split(",")]
        
        # Encontrar índices de las columnas deseadas
        idx_ModuleNumber = encabezados.index("ModuleNumber")
        idx_PartNumber = encabezados.index("PartNumber")
        idx_Location = encabezados.index("Location")
        idx_FeederID = encabezados.index("FeederID")

        # -------------------------------------------------
        # 3️ Leer datos desde la siguiente línea
        # -------------------------------------------------
        for item in tree.get_children():
            tree.delete(item)

        for i in range(encabezados_idx + 1, len(lineas)):
            linea = lineas[i].strip()
            if not linea:  # Saltar líneas vacías
                continue
            
            cols = [col.strip().strip('"\'') for col in linea.split(",")]
            
            # Validar que la línea tenga suficientes columnas
            if len(cols) <= max(idx_ModuleNumber, idx_PartNumber, idx_Location, idx_FeederID):
                continue
            
            values = [
                JobName,
                cols[idx_ModuleNumber],
                cols[idx_PartNumber],
                cols[idx_Location],
                cols[idx_FeederID]
            ]
            tree.insert("", tk.END, values=values)

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")



ventana = tk.Tk()
ventana.title("Extractor FeederSetUp")
ventana.geometry("900x500")

btn_seleccionar = tk.Button(ventana, text="Seleccionar Archivo CSV", command=seleccionar_archivo)
btn_seleccionar.pack(pady=10)

label_JobFolder = tk.Label(ventana, text="JobFolder: ", font=("Arial", 12, "bold"))
label_JobFolder.pack(pady=5)

columnas = ("JobFolder", "ModuleNumber", "PartNumber", "Location", "FeederID")
tree = ttk.Treeview(ventana, columns=columnas, show="headings")

for col in columnas:
    tree.heading(col, text=col)
    tree.column(col, width=150)

tree.pack(expand=True, fill="both", padx=10, pady=10)

ventana.mainloop()
