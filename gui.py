
import tkinter as tk
from tkinter import messagebox
from sap import connect_sap, sap_transaction


def submit():
    cliente = entry_cliente.get().strip()

    if not cliente:
        messagebox.showwarning("Warning", "Informe o cliente.")
        return

    try:
        session = connect_sap()
        sap_transaction(session, cliente)

        messagebox.showinfo("Success", "Processo finalizado com sucesso!")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# ---- GUI ----
root = tk.Tk()
root.title("SAP BP Tool")
root.geometry("300x140")
root.resizable(False, False)

label = tk.Label(root, text="Cliente:")
label.pack(pady=(15, 5))

entry_cliente = tk.Entry(root, width=30)
entry_cliente.pack()

btn = tk.Button(root, text="Submit", command=submit)
btn.pack(pady=15)

root.mainloop()
