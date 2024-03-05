import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Checkbutton

class App:
    def __init__(self, root):
        self.root = root
        self.root.title('Excel Cleaner')
        self.filename = None
        self.df = None
        self.columns = []
        self.checkbuttons = []

        # Botão para selecionar o arquivo
        self.select_button = tk.Button(root, text='Selecionar arquivo', command=self.select_file)
        self.select_button.pack()

        # Frame para os checkboxes
        self.frame = tk.Frame(root)
        self.frame.pack()

        # Campo para o nome do novo arquivo
        self.new_file_label = tk.Label(root, text='Nome do novo arquivo:')
        self.new_file_label.pack()
        self.new_file_entry = tk.Entry(root)
        self.new_file_entry.pack()

        # Botão para limpar a planilha
        self.clean_button = tk.Button(root, text='Limpar planilha', command=self.clean_excel, state=tk.DISABLED)
        self.clean_button.pack()

    def select_file(self):
        self.filename = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xls *.xlsx *.xml')])
        if self.filename:
            self.df = pd.read_excel(self.filename)
            self.columns = self.df.columns.tolist()

            # Limpar os checkboxes antigos
            for cb in self.checkbuttons:
                cb.destroy()
            self.checkbuttons.clear()

            # Criar novos checkboxes
            for column in self.columns:
                var = tk.BooleanVar(value=True)
                cb = Checkbutton(self.frame, text=column, variable=var)
                cb.var = var
                cb.pack(side='left')
                self.checkbuttons.append(cb)

            self.clean_button.config(state=tk.NORMAL)

    def clean_excel(self):
        new_filename = self.new_file_entry.get()
        if not new_filename:
            messagebox.showerror('Erro', 'Por favor, insira o nome do novo arquivo.')
            return

        # Verificar quais colunas foram selecionadas
        selected_columns = [cb.cget('text') for cb in self.checkbuttons if cb.var.get()]

        # Criar um novo DataFrame com apenas as colunas selecionadas
        new_df = self.df[selected_columns]

        # Salvar o novo DataFrame em um novo arquivo Excel
        new_df.to_excel(new_filename, index=False)
        messagebox.showinfo('Sucesso', 'Planilha limpa com sucesso!')

root = tk.Tk()
app = App(root)
root.mainloop()
