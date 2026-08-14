import customtkinter as ctk
from tkinter import filedialog
import os

ctk.set_appearance_mode("dark")

class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Consulta de Relatórios SLA")
        self.geometry("900x600")

        self.btn_carregar = ctk.CTkButton(
            self,

            text="Carregar Pasta",
            command=self.carregar_pasta
        )
        self.btn_carregar.pack(pady=10)

        self.lista = ctk.CTkTextbox(self, width=800, height=400)
        self.lista.pack(padx=10, pady=10)

    def carregar_pasta(self):
        pasta = filedialog.askdirectory()

        if not pasta:
            return

        self.lista.delete("1.0", "end")

        for arquivo in os.listdir(pasta):
            if arquivo.endswith(".xlsx"):
                self.lista.insert("end", f"{arquivo}\n")

app = App()
app.mainloop()