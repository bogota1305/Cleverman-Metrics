from tkinter import Tk, Label, Button, Entry, Checkbutton, IntVar, messagebox, Frame, Canvas, Scrollbar
from tkcalendar import Calendar
import pandas as pd

def open_date_selector():
    def get_dates():
        nonlocal start_date, end_date, output_file
        start_date = cal_start.get_date()
        end_date = cal_end.get_date()
        output_file = entry_name.get()

        if not output_file:
            messagebox.showerror("Error", "Por favor ingresa un nombre para el archivo.")
            return

        start_date = pd.to_datetime(start_date).strftime('%Y-%m-%d')
        end_date = pd.to_datetime(end_date).strftime('%Y-%m-%d')

        root.quit()
        root.destroy()

    def toggle_all(state):
        # for var in [all_var] + unique_orders_var + [orders_var, sales_var, payment_errors_var, expected_renewals_var, frequency_var, full_control_var, subs_var, refill_var, upsize_var, hear_var]:
        for var in [all_var] + unique_orders_var + [orders_var, sales_var, payment_errors_var, subs_var, refill_var, upsize_var, hear_var]:
            var.set(state)

    def toggle_section_a(state):
        for var in unique_orders_var:
            var.set(state)

    def on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # Variables de control
    start_date = None
    end_date = None
    output_file = None

    root = Tk()
    root.title("Seleccionar Fechas, Nombre de Archivo y Variables")
    root.geometry("420x650")
    root.resizable(False, True)

    # Botón de confirmación (declarado antes del canvas para que side=bottom funcione)
    btn_frame = Frame(root)
    btn_frame.pack(side="bottom", fill="x", pady=10)
    Button(btn_frame, text="Generar Reporte", command=get_dates).pack(pady=5)

    # --- Canvas + Scrollbar ---
    canvas = Canvas(root, borderwidth=0)
    scrollbar = Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    # Frame interior que vive dentro del canvas
    inner_frame = Frame(canvas)
    canvas_window = canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def on_canvas_configure(event):
        canvas.itemconfig(canvas_window, width=event.width)

    inner_frame.bind("<Configure>", on_frame_configure)
    canvas.bind("<Configure>", on_canvas_configure)
    canvas.bind_all("<MouseWheel>", on_mousewheel)

    # Sección de selección de fechas
    Label(inner_frame, text="Fecha de inicio:").pack(pady=5)
    cal_start = Calendar(inner_frame, selectmode='day', date_pattern='yyyy-mm-dd')
    cal_start.pack(pady=5)

    Label(inner_frame, text="Fecha de fin:").pack(pady=5)
    cal_end = Calendar(inner_frame, selectmode='day', date_pattern='yyyy-mm-dd')
    cal_end.pack(pady=5)

    Label(inner_frame, text="Name of the folder").pack(pady=5)
    entry_name = Entry(inner_frame)
    entry_name.pack(pady=5)

    # Sección de selección de variables
    Label(inner_frame, text="Select the reports:").pack(pady=10)

    variables_frame = Frame(inner_frame)
    variables_frame.pack(pady=5, anchor="w")

    all_var = IntVar()
    Checkbutton(variables_frame, text="All", variable=all_var, 
                command=lambda: toggle_all(all_var.get())).grid(row=0, column=0, sticky="w", padx=10)

    orders_var = IntVar()
    Checkbutton(variables_frame, text="All Orders", variable=orders_var, 
                command=lambda: toggle_section_a(orders_var.get())).grid(row=1, column=0, sticky="w", padx=20)

    orders_names = [
        'New - SUBS',
        'New - OTO',
        'New - MIX',
        'New - ALL',
        'Existing - SUBS',
        'Existing - OTO',
        'Existing - MIX',
        'Existing - ALL',
        'Recurrent Orders - ALL',
    ]

    unique_orders_var = []
    for i in range(1, 10):
        var = IntVar()
        unique_orders_var.append(var)
        Checkbutton(variables_frame, text= orders_names[i-1], variable=var).grid(row=i+1, column=0, sticky="w", padx=40)

    sales_var = IntVar()
    Checkbutton(variables_frame, text="Sales", variable=sales_var).grid(row=11, column=0, sticky="w", padx=20)

    payment_errors_var = IntVar()
    Checkbutton(variables_frame, text="Payment Errors", variable=payment_errors_var).grid(row=12, column=0, sticky="w", padx=20)

    expected_renewals_var = IntVar()
    Checkbutton(variables_frame, text="Expected Renewals", variable=expected_renewals_var).grid(row=13, column=0, sticky="w", padx=20)

    frequency_var = IntVar()
    Checkbutton(variables_frame, text="Real Frequency", variable=frequency_var).grid(row=14, column=0, sticky="w", padx=20)

    full_control_var = IntVar()
    Checkbutton(variables_frame, text="Full Control", variable=full_control_var).grid(row=15, column=0, sticky="w", padx=20)

    subs_var = IntVar()
    Checkbutton(variables_frame, text="Subscriptions", variable=subs_var).grid(row=16, column=0, sticky="w", padx=20)

    refill_var = IntVar()
    Checkbutton(variables_frame, text="Refill", variable=refill_var).grid(row=17, column=0, sticky="w", padx=20)
    
    upsize_var = IntVar()
    Checkbutton(variables_frame, text="Upsize", variable=upsize_var).grid(row=18, column=0, sticky="w", padx=20)

    hear_var = IntVar()
    Checkbutton(variables_frame, text="Hear from us?", variable=hear_var).grid(row=19, column=0, sticky="w", padx=20)

    root.mainloop()

    unique_orders_var = [ 
        unique_orders_var[0].get(),
        unique_orders_var[1].get(),
        unique_orders_var[2].get(),
        unique_orders_var[3].get(),
        unique_orders_var[4].get(),
        unique_orders_var[5].get(),
        unique_orders_var[6].get(),
        unique_orders_var[7].get(),
        unique_orders_var[8].get()
    ]

    if start_date and end_date and output_file:
        return start_date, end_date, output_file, all_var.get(), orders_var.get(), unique_orders_var, sales_var.get(), payment_errors_var.get(), expected_renewals_var.get(), frequency_var.get(), full_control_var.get(), subs_var.get(), refill_var.get(), upsize_var.get(), hear_var.get()
    else:
        messagebox.showerror("Error", "Por favor completa todos los campos.")
        return None, None, None, None, None, None, None, None, None