import tkinter as tk

class FormMain:

    def __init__(self):
        self.build_controls()
    
    def build_controls(self):
        self.form_main = tk.Tk()
        self.form_main.title("SLA manager")

        # Criar controles
        frame_top_left = tk.Frame(self.form_main,relief="solid",borderwidth=2)
        butto_edit = tk.Button(frame_top_left,text="Edit",width=10)
        butto_delete = tk.Button(frame_top_left,text="Delete",width=10)
            

        # Posicionar Controles
        frame_top_left.grid(row=1,column=1)
        butto_edit.grid(row=1,column=1)
        butto_delete.grid(row=1,column=2)
        self.form_main.mainloop()

form=FormMain() 