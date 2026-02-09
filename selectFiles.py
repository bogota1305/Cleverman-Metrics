from tkinter import filedialog, Label, Button, messagebox
import tkinter as tk


def _build_scrollable_window(title: str):
    root = tk.Tk()
    root.title(title)
    root.geometry("900x600")

    container = tk.Frame(root)
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(container)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    scrollable_frame = tk.Frame(canvas)
    window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    def _on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    scrollable_frame.bind("<Configure>", _on_frame_configure)

    def _on_canvas_configure(event):
        canvas.itemconfig(window_id, width=event.width)

    canvas.bind("<Configure>", _on_canvas_configure)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux_up(event):
        canvas.yview_scroll(-1, "units")

    def _on_mousewheel_linux_down(event):
        canvas.yview_scroll(1, "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Button-4>", _on_mousewheel_linux_up)
    canvas.bind_all("<Button-5>", _on_mousewheel_linux_down)

    return root, scrollable_frame


def seleccionar_archivos_para_casos():
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

        "Beard - JetBlack": None,
        "Beard - AfricanAmerican": None,
        "Beard - Black": None,
        "Beard - Blond": None,
        "Beard - Red": None,
        "Beard - Brown": None,
        "Beard - MediumDarkBrown": None,

        "Hair - AfricanAmerican": None,
        "Hair - Red": None,
        "Hair - Black": None,
        "Hair - Brown": None,
        "Hair - LightBrown": None,
        "Hair - Blond": None,

        "My instructions": None,
        "Inmediate coverage": None,
        "Referral": None,
        "Grey hair color touch up": None,
        "Salt and peper": None,
        "Full coverage": None,
        "Customized beard": None,
        "Best beard color": None,
        "Best hair color one time": None,
        "Best hair color sub": None,
        "Customized beard sub": None,
    }

    # Guardamos labels por key para actualizar texto cuando asignemos en batch
    labels_por_caso = {}

    def seleccionar_archivo(caso):
        archivo = filedialog.askopenfilename(
            title=f"Seleccionar: {caso}",
            filetypes=[("CSV files", "*.csv")]
        )
        if archivo:
            archivos_seleccionados[caso] = archivo
            labels_por_caso[caso].config(text=f"Seleccionado: {archivo}")

    def seleccionar_todos_en_orden():
        paths = filedialog.askopenfilenames(
            title="Selecciona TODOS los CSV (en el mismo orden del script)",
            filetypes=[("CSV files", "*.csv")]
        )
        if not paths:
            return

        keys = list(archivos_seleccionados.keys())

        if len(paths) != len(keys):
            messagebox.showerror(
                "Cantidad incorrecta",
                f"Seleccionaste {len(paths)} archivos, pero se esperaban {len(keys)}.\n\n"
                "Selecciona todos los CSV de una sola vez."
            )
            return

        # Asignación por orden
        for k, p in zip(keys, paths):
            archivos_seleccionados[k] = p
            labels_por_caso[k].config(text=f"Seleccionado: {p}")

        messagebox.showinfo("Listo", "Se asignaron todos los archivos por orden.")

    root, content = _build_scrollable_window("Seleccionar archivos para cada caso")

    # Barra superior con selección masiva
    top_bar = tk.Frame(content)
    top_bar.pack(pady=10, padx=10, fill="x")

    btn_all = Button(top_bar, text="Seleccionar TODOS (en orden)", command=seleccionar_todos_en_orden)
    btn_all.pack(side="left")

    hint = Label(
        top_bar,
        text="Tip: en el explorador ordena por nombre y selecciona en bloque (Shift) para mantener el orden.",
        wraplength=650,
        justify="left"
    )
    hint.pack(side="left", padx=15)

    # Lista individual (por si falta alguno)
    for caso in archivos_seleccionados.keys():
        frame = tk.Frame(content)
        frame.pack(pady=5, padx=10, fill="x")

        label = Label(frame, text=f"Selecciona: {caso}", wraplength=520, justify="left")
        label.pack(side="left", padx=10)
        labels_por_caso[caso] = label

        boton = Button(
            frame,
            text="Seleccionar archivo",
            command=lambda c=caso: seleccionar_archivo(c)
        )
        boton.pack(side="right", padx=10)

    tk.Label(content, text="").pack()
    confirmar = Button(content, text="Confirmar selección", command=root.quit)
    confirmar.pack(pady=20)

    root.mainloop()
    root.destroy()
    return archivos_seleccionados


def seleccionar_archivos_stripe():
    archivos_seleccionados = {
        "Blocked Payments": None,
        "All Payments": None,
    }

    labels_por_caso = {}

    def seleccionar_archivo(caso):
        archivo = filedialog.askopenfilename(
            title=f"Seleccionar: {caso}",
            filetypes=[("CSV files", "*.csv")]
        )
        if archivo:
            archivos_seleccionados[caso] = archivo
            labels_por_caso[caso].config(text=f"Seleccionado: {archivo}")

    def seleccionar_todos_en_orden():
        paths = filedialog.askopenfilenames(
            title="Selecciona los 2 CSV (en orden)",
            filetypes=[("CSV files", "*.csv")]
        )
        if not paths:
            return

        keys = list(archivos_seleccionados.keys())

        if len(paths) != len(keys):
            messagebox.showerror(
                "Cantidad incorrecta",
                f"Seleccionaste {len(paths)} archivos, pero se esperaban {len(keys)}."
            )
            return

        for k, p in zip(keys, paths):
            archivos_seleccionados[k] = p
            labels_por_caso[k].config(text=f"Seleccionado: {p}")

        messagebox.showinfo("Listo", "Se asignaron ambos archivos por orden.")

    root, content = _build_scrollable_window("Seleccionar archivo")

    top_bar = tk.Frame(content)
    top_bar.pack(pady=10, padx=10, fill="x")

    btn_all = Button(top_bar, text="Seleccionar los 2 (en orden)", command=seleccionar_todos_en_orden)
    btn_all.pack(side="left")

    for caso in archivos_seleccionados.keys():
        frame = tk.Frame(content)
        frame.pack(pady=5, padx=10, fill="x")

        label = Label(frame, text=f"Selecciona: {caso}", wraplength=520, justify="left")
        label.pack(side="left", padx=10)
        labels_por_caso[caso] = label

        boton = Button(frame, text="Seleccionar archivo", command=lambda c=caso: seleccionar_archivo(c))
        boton.pack(side="right", padx=10)

    tk.Label(content, text="").pack()
    confirmar = Button(content, text="Confirmar selección", command=root.quit)
    confirmar.pack(pady=20)

    root.mainloop()
    root.destroy()
    return archivos_seleccionados
