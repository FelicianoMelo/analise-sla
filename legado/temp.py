import os
import tkinter as tk
from tkinter import messagebox, ttk
import openpyxl 


def initialize_excel(filename="dados.xlsx"):
    """Cria o arquivo Excel com os cabeçalhos se ele ainda não existir."""
    if not os.path.exists(filename):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Registros"
        sheet.append(
            ["Nome", "Idade", "Departamento", "Cargo", "Status de Emprego"]
        )
        wb.save(filename)

def save_data():
    """Lê os dados da interface, valida e salva no Excel."""
    nome = name_entry.get().strip()
    idade = age_spinbox.get()
    dept = dept_combobox.get()
    cargo = title_entry.get().strip()
    status = status_var.get()

    # Validação simples
    if not nome or not cargo:
        messagebox.showwarning(
            "Aviso de Validação", "Por favor, preencha os campos Nome e Cargo."
        )
        return

    filename = "dados.xlsx"
    initialize_excel(filename)

    # Abre o arquivo e adiciona a nova linha
    wb = openpyxl.load_workbook(filename)
    sheet = wb["Registros"]
    sheet.append([nome, int(idade), dept, cargo, status])
    wb.save(filename)

    messagebox.showinfo("Sucesso", "Dados registrados com sucesso no Excel!")
    clear_form()

def clear_form():
    """Limpa todos os campos da interface."""
    name_entry.delete(0, tk.END)
    age_spinbox.delete(0, tk.END)
    age_spinbox.insert(0, "18")
    dept_combobox.set("TI")
    title_entry.delete(0, tk.END)
    status_var.set("Ativo")

# --- Configuração da Janela Principal ---
root = tk.Tk()
root.title("Formulário de Entrada de Dados")
#root.geometry("450 x 380")
root.geometry("450x380+100+50")
root.resizable(False, False)

# Estilo
style = ttk.Style()
style.theme_use("clam")

frame = ttk.LabelFrame(root, text=" Informações do Funcional ", padding=15)
frame.pack(fill="both", expand=True, padx=15, pady=15)

# Campo: Nome
ttk.Label(frame, text="Nome Completo:").grid(
    row=0, column=0, sticky="w", pady=5
)
name_entry = ttk.Entry(frame, width=30)
name_entry.grid(row=0, column=1, pady=5)

# Campo: Idade
ttk.Label(frame, text="Idade:").grid(row=1, column=0, sticky="w", pady=5)
age_spinbox = ttk.Spinbox(frame, from_=18, to=100, width=28)
age_spinbox.insert(0, "18")
age_spinbox.grid(row=1, column=1, pady=5)

# Campo: Departamento
ttk.Label(frame, text="Departamento:").grid(
    row=2, column=0, sticky="w", pady=5
)
dept_combobox = ttk.Combobox(
    frame,
    values=["TI", "Recursos Humanos", "Financeiro", "Vendas", "Marketing"],
    state="readonly",
    width=28,
)
dept_combobox.set("TI")
dept_combobox.grid(row=2, column=1, pady=5)

# Campo: Cargo
ttk.Label(frame, text="Cargo:").grid(row=3, column=0, sticky="w", pady=5)
title_entry = ttk.Entry(frame, width=30)
title_entry.grid(row=3, column=1, pady=5)

# Campo: Status (Radiobuttons)
ttk.Label(frame, text="Status:").grid(row=4, column=0, sticky="w", pady=5)
status_var = tk.StringVar(value="Ativo")
status_frame = ttk.Frame(frame)
status_frame.grid(row=4, column=1, sticky="w", pady=5)

ttk.Radiobutton(
    status_frame, text="Ativo", value="Ativo", variable=status_var
).pack(side="left", padx=5)
ttk.Radiobutton(
    status_frame, text="Inativo", value="Inativo", variable=status_var
).pack(side="left", padx=5)

# Botões
button_frame = ttk.Frame(frame)
button_frame.grid(row=5, column=0, columnspan=2, pady=20)

submit_btn = ttk.Button(button_frame, text="Salvar Dados", command=save_data)
submit_btn.pack(side="left", padx=10)

clear_btn = ttk.Button(button_frame, text="Limpar", command=clear_form)
clear_btn.pack(side="left", padx=10)

root.mainloop()