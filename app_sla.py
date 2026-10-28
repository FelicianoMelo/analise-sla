"""Interface gráfica para listar planilhas Excel de uma pasta."""

from pathlib import Path
from tkinter import filedialog, messagebox
import customtkinter as ctk


class App(ctk.CTk):
    """Janela principal da aplicação de consulta de relatórios SLA."""

    def __init__(self) -> None:
        super().__init__()

        self.title("Consulta de Relatórios SLA")
        self.geometry("900x600")

        self.btn_carregar = ctk.CTkButton(
            self,
            text="Carregar Pasta",
            command=self.carregar_pasta,
        )
        self.btn_carregar.pack(pady=10)

        self.lista = ctk.CTkTextbox(self, width=800, height=400)
        self.lista.pack(padx=10, pady=10)

    def carregar_pasta(self) -> None:
        """Solicita uma pasta e mostra as planilhas Excel encontradas nela."""
        pasta_selecionada = filedialog.askdirectory(
            title="Selecione a pasta com os relatórios",
        )

        # Cancelar o diálogo não deve apagar a lista já exibida.
        if not pasta_selecionada:
            return

        pasta = Path(pasta_selecionada)

        try:
            arquivos_excel = sorted(
                (
                    arquivo.name
                    for arquivo in pasta.iterdir()
                    if arquivo.is_file() and arquivo.suffix.lower() == ".xlsx"
                ),
                key=str.casefold,
            )
        except OSError as erro:
            messagebox.showerror(
                "Erro ao carregar pasta",
                f"Não foi possível ler a pasta selecionada:\n{erro}",
                parent=self,
            )
            return

        self.lista.delete("1.0", "end")
        if arquivos_excel:
            self.lista.insert("end", "\n".join(arquivos_excel))


def main() -> None:
    """Configura o tema e inicia o loop de eventos da interface."""
    ctk.set_appearance_mode("dark")
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
