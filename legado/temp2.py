import os
import tkinter as tk
from tkinter import messagebox, ttk
import openpyxl

# Constantes globais do sistema
FILE_PATH = "dados.xlsx"
DEPARTMENTS = ["TI", "Recursos Humanos", "Financeiro", "Vendas", "Marketing"]
STATUS_OPTIONS = ["Ativo", "Inativo"]


class DataEntryApp(tk.Tk):
    """Classe principal do aplicativo de formulário de cadastro."""

    def __init__(self, filename=FILE_PATH):
        super().__init__()

        self.filename = filename
        self.title("Formulário de Entrada de Dados")
        self.geometry("460x400")
        self.resizable(False, False)

        self._configure_styles()
        self._create_widgets()
        self._ensure_excel_file()

    def _configure_styles(self):
        """Aplica estilos visuais aos widgets do ttk."""
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

    def _ensure_excel_file(self):
        """Garante a existência do arquivo Excel de destino com cabeçalhos."""
        if not os.path.exists(self.filename):
            try:
                wb = openpyxl.Workbook()
                sheet = wb.active
                sheet.title = "Registros"
                sheet.append(
                    [
                        "Nome",
                        "Idade",
                        "Departamento",
                        "Cargo",
                        "Status de Emprego",
                    ]
                )
                wb.save(self.filename)
            except Exception as e:
                messagebox.showerror(
                    "Erro de Arquivo",
                    f"Não foi possível criar a planilha inicial:\n{e}",
                )

    def _create_widgets(self):
        """Constrói e posiciona todos os elementos visuais do formulário."""
        frame = ttk.LabelFrame(self, text=" Informações do Funcional ", padding=15)
        frame.pack(fill="both", expand=True, padx=15, pady=15)
        frame.columnconfigure(1, weight=1)

        # Nome Completo
        ttk.Label(frame, text="Nome Completo:").grid(
            row=0, column=0, sticky="w", pady=6
        )
        self.name_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.name_var).grid(
            row=0, column=1, sticky="ew", pady=6
        )

        # Idade
        ttk.Label(frame, text="Idade:").grid(
            row=1, column=0, sticky="w", pady=6
        )
        self.age_var = tk.StringVar(value="18")
        ttk.Spinbox(
            frame, from_=18, to=100, textvariable=self.age_var, width=28
        ).grid(row=1, column=1, sticky="ew", pady=6)

        # Departamento
        ttk.Label(frame, text="Departamento:").grid(
            row=2, column=0, sticky="w", pady=6
        )
        self.dept_var = tk.StringVar(value=DEPARTMENTS[0])
        ttk.Combobox(
            frame,
            textvariable=self.dept_var,
            values=DEPARTMENTS,
            state="readonly",
        ).grid(row=2, column=1, sticky="ew", pady=6)

        # Cargo
        ttk.Label(frame, text="Cargo:").grid(
            row=3, column=0, sticky="w", pady=6
        )
        self.title_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.title_var).grid(
            row=3, column=1, sticky="ew", pady=6
        )

        # Status de Emprego
        ttk.Label(frame, text="Status:").grid(
            row=4, column=0, sticky="w", pady=6
        )
        self.status_var = tk.StringVar(value=STATUS_OPTIONS[0])
        status_frame = ttk.Frame(frame)
        status_frame.grid(row=4, column=1, sticky="w", pady=6)

        for status in STATUS_OPTIONS:
            ttk.Radiobutton(
                status_frame,
                text=status,
                value=status,
                variable=self.status_var,
            ).pack(side="left", padx=5)

        # Área dos Botões
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=20)

        ttk.Button(
            button_frame, text="Salvar Dados", command=self.save_data
        ).pack(side="left", padx=10)

        ttk.Button(
            button_frame, text="Limpar", command=self.clear_form
        ).pack(side="left", padx=10)

    def _validate_inputs(self):
        """Valida se os campos obrigatórios e os tipos de dados estão corretos."""
        nome = self.name_var.get().strip()
        cargo = self.title_var.get().strip()

        if not nome or not cargo:
            messagebox.showwarning(
                "Aviso de Validação", "Os campos Nome e Cargo são obrigatórios."
            )
            return None

        try:
            idade = int(self.age_var.get())
            if not (18 <= idade <= 100):
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Aviso de Validação",
                "Insira um número inteiro válido entre 18 e 100 no campo Idade.",
            )
            return None

        return {
            "nome": nome,
            "idade": idade,
            "dept": self.dept_var.get(),
            "cargo": cargo,
            "status": self.status_var.get(),
        }

    def save_data(self):
        """Executa a persistência dos dados validados no arquivo Excel."""
        data = self._validate_inputs()
        if not data:
            return

        try:
            wb = openpyxl.load_workbook(self.filename)
            sheet = wb["Registros"]
            sheet.append(
                [
                    data["nome"],
                    data["idade"],
                    data["dept"],
                    data["cargo"],
                    data["status"],
                ]
            )
            wb.save(self.filename)

            messagebox.showinfo(
                "Sucesso", "Dados registrados com sucesso no Excel!"
            )
            self.clear_form()

        except PermissionError:
            messagebox.showerror(
                "Erro de Permissão",
                f"O arquivo '{self.filename}' está aberto em outro programa.\n"
                "Feche o arquivo e tente salvar novamente.",
            )
        except Exception as e:
            messagebox.showerror(
                "Erro Inesperado", f"Ocorreu um erro ao salvar os dados:\n{e}"
            )

    def clear_form(self):
        """Reseta todos os elementos de entrada para o estado inicial."""
        self.name_var.set("")
        self.age_var.set("18")
        self.dept_var.set(DEPARTMENTS[0])
        self.title_var.set("")
        self.status_var.set(STATUS_OPTIONS[0])


if __name__ == "__main__":
    app = DataEntryApp()
    app.mainloop()