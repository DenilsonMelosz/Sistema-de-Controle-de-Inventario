import tkinter as tk
from tkinter import messagebox
import sqlite3
import pandas as pd
from tkinter import filedialog
import hashlib
import os  # Importe o módulo os para usar os.urandom(32)
from tkinter import Scrollbar


# Database connection
conn = sqlite3.connect('inventario.db')
c = conn.cursor()
conn = sqlite3.connect('users.db')
c = conn.cursor()


# Users table
c.execute("""
    CREATE TABLE IF NOT EXISTS users  
    (id INTEGER PRIMARY KEY AUTOINCREMENT,
     username TEXT UNIQUE,
     password_hash TEXT)
""")


# Inventory table
c.execute("""
    CREATE TABLE IF NOT EXISTS inventario  
    (wks TEXT PRIMARY KEY, 
     nome TEXT, 
     quantidade INTEGER,
     descricao TEXT,
     status TEXT,
     setor TEXT,
     responsavel TEXT)
""")

conn.commit()

# Movements table
c.execute("""
    CREATE TABLE IF NOT EXISTS movimentacoes  
    (id INTEGER PRIMARY KEY AUTOINCREMENT,
     wks TEXT,
     setor_anterior TEXT,
     responsavel_anterior TEXT,
     status_anterior TEXT,
     descricao_anterior TEXT,
     quantidade_anterior INTEGER,
     data_movimentacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP)
""")

conn.commit()

# Add default users if they don't exist
users = [
    ("Denilson", "123456789"),
    ("Victor", "123456789"),
    ("Felipe", "123456789")
]

for user in users:
    username = user[0]
    password = user[1]
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    c.execute("INSERT OR IGNORE INTO users (username, password_hash) VALUES (?, ?)", (username, password_hash))

conn.commit()

class LoginSystem:
    def __init__(self):
        # Create the login window
        self.login_window = tk.Tk()
        self.login_window.title('Login')
        self.login_window.geometry('450x250')



        self.username_label = tk.Label(self.login_window, text='USUARIO:', fg='white', bg='Black', font=('Arial', 15))
        self.login_window.configure(bg='black')  # Cor de fundo para a janela principal
        self.username_label.pack(pady=8)
        self.username_entry = tk.Entry(self.login_window, width=30,font=('Arial', 13))
        self.username_entry.pack()

        self.password_label = tk.Label(self.login_window, text='SENHA:', fg='white', bg='Black', font=('Arial',15))
        self.password_label.pack(pady=14)
        self.password_entry = tk.Entry(self.login_window, show='*', width=30,font=('Arial', 13))
        self.password_entry.pack()

        self.login_button = tk.Button(self.login_window, text='ENTRAR', command=self.login, font=('Arial bold', 9), width=14, height=2, bg='black', fg='white', bd=5)
        self.login_button.pack(pady=14)

    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        # Hash the entered password
        password_hash = hashlib.sha256(password.encode()).hexdigest()

        c.execute("SELECT * FROM users WHERE username=? AND password_hash=?", (username, password_hash))
        user = c.fetchone()

        if user:
            # Prompt for additional confirmation
            confirm_login = messagebox.askyesno('Confirm Login', 'Você deseja ingressar no Controle de Inventario?')
            if confirm_login:
                self.logged_in = True  # Set login status to True
                self.login_window.destroy()
        else:
            messagebox.showerror('Login Failed', 'Usuario ou Senha incorreto!!')
            self.username_entry.delete(0, tk.END)
            self.password_entry.delete(0, tk.END)

    def run(self):
        self.login_window.mainloop()
        if not self.logged_in:
            self.login_window.quit()  # If not logged in, quit the entire application


class InventorySystem:

    def __init__(self):
        # Create the main window
        self.root = tk.Tk()
        self.root.title('Controle de Inventário')

        # Configurar as cores de fundo e de texto para os diferentes widgets
        self.root.configure(bg='lightcyan')  # Cor de fundo para a janela principal

        # Frame de entrada
        self.input_frame = tk.LabelFrame(self.root, text='Registrar Equipamento')
        self.input_frame.pack(fill='x', padx=10, pady=10)

        # Rótulos e campos de entrada
        self.name_label = tk.Label(self.input_frame, text='Nome:')
        self.name_label.grid(row=0, column=0)
        self.name_entry = tk.Entry(self.input_frame)
        self.name_entry.grid(row=0, column=1)

        self.quantity_label = tk.Label(self.input_frame, text='Quantidade:')
        self.quantity_label.grid(row=1, column=0)
        self.quantity_spinbox = tk.Spinbox(self.input_frame, from_=0, to=100)
        self.quantity_spinbox.grid(row=1, column=1)

        self.description_label = tk.Label(self.input_frame, text='Descrição:')
        self.description_label.grid(row=2, column=0)
        self.selected_description = tk.StringVar()  # Variável para armazenar a descrição selecionada
        self.description_dropdown = tk.OptionMenu(self.input_frame, self.selected_description, *self.get_descriptions())
        self.description_dropdown.grid(row=2, column=1)

        self.status_label = tk.Label(self.input_frame, text='Status:')
        self.status_label.grid(row=3, column=0)
        self.selected_status = tk.StringVar()  # Variável para armazenar o status selecionado
        self.status_dropdown = tk.OptionMenu(self.input_frame, self.selected_status, *self.get_status())
        self.status_dropdown.grid(row=3, column=1)

        # Dropdown de seleção de setor
        self.sector_label = tk.Label(self.input_frame, text='Setor:')
        self.sector_label.grid(row=4, column=0)
        self.selected_sector = tk.StringVar()  # Variável para armazenar o setor selecionado
        self.sector_dropdown = tk.OptionMenu(self.input_frame, self.selected_sector, *self.get_sectors())
        self.sector_dropdown.grid(row=4, column=1)

        self.responsible_label = tk.Label(self.input_frame, text='Responsável:')
        self.responsible_label.grid(row=5, column=0)
        self.responsible_entry = tk.Entry(self.input_frame)
        self.responsible_entry.grid(row=5, column=1)

        self.wks_label = tk.Label(self.input_frame, text='WKS:')
        self.wks_label.grid(row=6, column=0)
        self.wks_entry = tk.Entry(self.input_frame)
        self.wks_entry.grid(row=6, column=1)

        # Botão Registrar
        self.register_button = tk.Button(self.input_frame, text='Registrar', bg='black', fg='white')  # Alterando as cores do botão 'Registrar'
        self.register_button.grid(row=7, column=0, pady=10)

        # Botões Limpar e Excluir
        self.clear_button = tk.Button(self.input_frame, text='Limpar',  bg='black', fg='white', command=self.clear_fields)
        self.clear_button.grid(row=7, column=1)


        self.delete_button = tk.Button(self.input_frame, text='Excluir', bg='red', fg='white', command=self.delete_equipment)
        self.delete_button.grid(row=7, column=2, pady=10)

        # Criar e exibir um botão para iniciar a pesquisa
        self.search_button = tk.Button(self.input_frame, text='Pesquisar', command=self.search_equipment)
        self.search_button.grid(row=8, column=0, pady=10)

        ###############vvvv################ 27/09/2023 #####################vvvvvv##########

        # Botão Mover Equipamento
        self.move_button = tk.Button(self.input_frame, text='Mover Equipamento', command=self.move_equipment)
        self.move_button.grid(row=8, column=1, pady=10)  # Coloque o botão na posição desejada

        # Botão Consultar Histórico
        self.history_button = tk.Button(self.input_frame, text='Consultar Histórico',
                                        command=self.show_movement_history)
        self.history_button.grid(row=8, column=2, pady=10)  # Coloque o botão na posição desejada


        ###############^^^^^################ 27/09/2023 ######################^^^^^##########

        ###############vvvvv################ 05/10/2023 ######################vvvvv##########
        # Botão para carregar arquivo de planilha
        self.load_button = tk.Button(self.input_frame, text='Carregar Planilha', command=self.load_excel)
        self.load_button.grid(row=8, column=3, pady=10)
        ###############^^^^^################ 05/10/2023 ######################^^^^^##########

        ###############vvvvv################ 11/10/2023 ######################vvvvv##########

        self.delete_all_button = tk.Button(self.input_frame, text='Apagar Todos', command=self.delete_all_records)
        self.delete_all_button.grid(row=8, column=4, pady=12)  # Coloque o botão na posição desejada

        ###############^^^^^################ 05/10/2023 ######################^^^^^##########

        # Frame de saída
        self.output_frame = tk.LabelFrame(self.root, text='Relatórios')
        self.output_frame.pack(fill='x', padx=10, pady=10)

        # Rótulo de contagem de inventário
        self.inventory_count = tk.StringVar()
        self.count_label = tk.Label(self.output_frame, textvariable=self.inventory_count)
        self.count_label.grid(row=1, column=0, columnspan=3)

        # Inicializar a contagem do inventário
        self.update_inventory_count()

        # Botões de saída
        self.list_button = tk.Button(self.output_frame, text='Listar Equipamentos', command=self.list_equipment)
        self.list_button.grid(row=0, column=0, padx=20)

        self.count_button = tk.Button(self.output_frame, text='Quantidade Status', command=self.count_inventory)
        self.count_button.grid(row=0, column=1, padx=10)

        self.export_button = tk.Button(self.output_frame, text='Exportar Planilha', command=self.export_report)
        self.export_button.grid(row=0, column=2, padx=10)
#################################vv 05/10/23 vv#################################
    def get_movement_history(self, wks):
        # Consulta o histórico de movimentações para o equipamento com o número de série (WKS) fornecido
        c.execute("SELECT * FROM movimentacoes WHERE wks=?", (wks,))
        history = c.fetchall()
        return history

#################################^^ 05/10/23 ^^#################################

    def get_sectors(self):
        # Modifique este método para retornar a lista de setores que você deseja exibir no menu suspenso
        sectors = ["","RECEPÇÃO-A", "RECEPÇÃO-O", "RECEPÇÃO-T", "CONSULTÓRIO3-A", "CONSULTÓRIO4-A", "CONSULTÓRIO1-O", "CONSULTÓRIO2-O", "CONSULTÓRIO1-T", "CONSULTÓRIO2-T", "ENFERMAGEM-A", "ENFERMAGEM-O", "ENFERMAGEM-T", "REEMBOLSO", "ALMOXARIFADO-O", "ALMOXARIFADO-T", "COORD-SUPRIMENTO", "COORD-ENFERMAGEM", "COORD-ATENDIMENTO", "CPD-O", "CPD-A", "CPD-T", "CPD-ADM", "DIRETORIA-ADM", "DIRETORIA-FINANCEIRO", "RH", "CALLCENTER", "FINANCEIRO", "AUTORIZAÇÃO", "COMERCIAL", "FATURAMENTO", "SALA-REUNIÃO","ARQUIVO", "T.I.", "DEPOSITO-O"]
        return sectors

    def get_status(self):
        # Modifique este método para retornar a lista de setores que você deseja exibir no menu suspenso
        status = ["","EM USO", "BACKUP", "INUTILIZÁVEL", "MANUTENÇÃO-SSIT", "MANUTENÇÃO-A.TÉCNICA", "DEPÓSITO"]
        return status

    def get_descriptions(self):
        descriptions = ["", "NOTEBOOK", "DESKTOP", "TABLET", "CELULAR", "CABO VGA", "CABO HDMI", "CABO FORÇA", "CABO DE REDE", "IMPRESSORA", "AP", "NOBREAK", "ESTABILIZADOR", "TECLADO", "MOUSE"]
        return descriptions

    def register_equipment(self):
        wks = self.wks_entry.get()
        name = self.name_entry.get()
        quantity = int(self.quantity_spinbox.get())
        description = self.selected_description.get()  # Obtenha a descrição selecionada no menu suspenso
        status = self.selected_status.get()  # Obtenha o status selecionado no menu suspenso
        sector = self.selected_sector.get()  # Obtenha o setor selecionado no menu suspenso
        responsible = self.responsible_entry.get()

        # Validação de dados
        if not wks or not name or not quantity or quantity == 0 or not description or not status or not sector:
            messagebox.showerror(title='Erro', message='Preencha todos os campos obrigatórios')
            return

        # Verifique se o equipamento com o mesmo WKS já existe
        c.execute("SELECT * FROM inventario WHERE wks=?", (wks,))
        existing_equipment = c.fetchone()

        if existing_equipment:
            # O equipamento com o mesmo WKS já existe, exiba uma mensagem de erro
            messagebox.showerror(title='Erro', message='Equipamento com o mesmo WKS já existe')
        else:
            # O equipamento não existe, insira um novo registro
            c.execute("""INSERT INTO inventario (wks, nome, quantidade, descricao, status, setor, responsavel)
                         VALUES (?, ?, ?, ?, ?, ?, ?)""",
                      (wks, name, quantity, description, status, sector, responsible))
            conn.commit()
            messagebox.showinfo(title='Sucesso', message='Equipamento registrado com sucesso!')

        # Limpar campos
        self.clear_fields()

        # Atualizar a contagem de inventário
        self.update_inventory_count()

    def clear_fields(self):
        self.name_entry.delete(0, tk.END)
        self.quantity_spinbox.delete(0, tk.END)
        self.selected_description.set('')  # Limpe a descrição selecionada
        self.selected_status.set('')  # Limpe o status selecionado
        self.selected_sector.set('')  # Limpe o setor selecionado
        self.responsible_entry.delete(0, tk.END)
        self.wks_entry.delete(0, tk.END)

    def delete_equipment(self):
        wks = self.wks_entry.get()  # Altere a variável "id" para "wks" para corresponder à nova coluna

        if wks == '':
            messagebox.showerror(title='Erro', message='Insira o WKS do equipamento')
            return

        query = "SELECT * FROM inventario WHERE wks = ?"  # Use a coluna "wks" corretamente
        result = c.execute(query, (wks,)).fetchone()

        if not result:
            messagebox.showerror(title='Erro', message='Equipamento não encontrado')
            return

        confirm = messagebox.askyesno(title='Confirmar', message='Excluir equipamento?')
        if confirm:
            c.execute("DELETE FROM inventario WHERE wks=?", (wks,))  # Use a coluna "wks" corretamente
            conn.commit()
            messagebox.showinfo(title='Sucesso', message='Equipamento excluído')

        self.clear_fields()
        self.update_inventory_count()

    def list_equipment(self):
        # Consultar todos os equipamentos do banco de dados
        c.execute("SELECT * FROM inventario")
        records = c.fetchall()

        # Criar uma nova janela para exibir a lista de equipamentos
        list_window = tk.Toplevel(self.root)
        list_window.title('Lista de Equipamentos')

        # Frame para conter o widget Text e a barra de rolagem
        list_frame = tk.Frame(list_window)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # Barra de rolagem vertical
        y_scrollbar = tk.Scrollbar(list_frame)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Barra de rolagem horizontal
        x_scrollbar = tk.Scrollbar(list_frame, orient=tk.HORIZONTAL)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Widget Text para exibir a lista de equipamentos
        text_area = tk.Text(list_frame, wrap=tk.NONE, yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Alterar a fonte e o tamanho da letra na lista de equipamentos
        text_font = ("Arial", 14)  # Substitua "Arial" pela fonte desejada e 10 pelo tamanho da fonte
        text_area.config(font=text_font)

        # Configurar as barras de rolagem para controlar o widget Text
        y_scrollbar.config(command=text_area.yview)
        x_scrollbar.config(command=text_area.xview)

        # Preencher o widget Text com os dados da lista de equipamentos
        for record in records:
            text_area.insert(tk.END, f"ID: {record[0]}, Nome: {record[1]}, Quantidade: {record[2]}\n")
            text_area.insert(tk.END, f"Responsável: {record[6]}\n\n")  # Incluindo o responsável


        # Desativar a edição no widget Text
        text_area.config(state=tk.DISABLED)

        # Configurar o tamanho da janela para se ajustar ao conteúdo
        list_window.update()
        list_window_width = text_area.winfo_reqwidth() + y_scrollbar.winfo_reqwidth()
        list_window_height = text_area.winfo_reqheight() + x_scrollbar.winfo_reqheight()
        list_window.minsize(list_window_width, list_window_height)

    def count_inventory(self):
        selected_description = self.selected_description.get()
        selected_status = self.selected_status.get()

        if not selected_description and not selected_status:
            messagebox.showerror(title='Erro', message='Selecione uma descrição ou status')
            return

        sql = "SELECT SUM(quantidade) FROM inventario WHERE "
        conditions = []
        parameters = []

        if selected_description:
            conditions.append("descricao = ?")
            parameters.append(selected_description)

        if selected_status:
            conditions.append("status = ?")
            parameters.append(selected_status)

        if conditions:
            sql += " AND ".join(conditions)

        c.execute(sql, parameters)
        total_quantity = c.fetchone()[0] or 0

        # Atualize a variável de classe total_quantity
        self.total_quantity = total_quantity

        messagebox.showinfo(title='Quantidade em Estoque', message=f'Total de equipamentos: {self.total_quantity}')

    def update_inventory_count(self):
        c.execute("SELECT SUM(quantidade) FROM inventario")
        total_quantity = c.fetchone()[0] or 0

        # Atualize a variável de classe total_quantity
        self.total_quantity = total_quantity

        self.inventory_count.set(f'Total de equipamentos: {self.total_quantity}')

    def export_report(self):
        path = filedialog.askdirectory()

        c.execute("SELECT * FROM inventario")
        rows = c.fetchall()

        # Criar um DataFrame
        df = pd.DataFrame(rows, columns=['ID', 'Nome', 'Quantidade', 'Descrição', 'Status', 'Setor', 'Responsável'])

        # Exportar para Excel
        filename = path + '/relatorio_inventario.xlsx'
        df.to_excel(filename, index=False)

        messagebox.showinfo(title='Exportado', message='Relatório salvo em ' + filename)

    def update_inventory_count(self):
        c.execute("SELECT SUM(quantidade) FROM inventario")
        total_quantity = c.fetchone()[0] or 0
        self.inventory_count.set(f'Total de equipamentos: {total_quantity}')
####################################################################################################################
    def move_equipment(self):
        wks = self.wks_entry.get()  # Obtenha o número de série do equipamento a ser movido
        new_sector = self.selected_sector.get()  # Obtenha o novo setor para o equipamento
        new_responsible = self.responsible_entry.get()  # Obtenha o novo responsável para o equipamento
        new_status = self.selected_status.get()  # Obtenha o novo status para o equipamento
        new_description = self.selected_description.get()  # Obtenha a nova descrição para o equipamento
        new_quantity = int(self.quantity_spinbox.get())  # Obtenha a nova quantidade para o equipamento

        # Verifique se o número de série (WKS) foi fornecido
        if not wks:
            messagebox.showerror(title='Erro', message='Insira o WKS do equipamento a ser movido')
            return

        # Verifique se o equipamento existe no banco de dados
        c.execute("SELECT * FROM inventario WHERE wks=?", (wks,))
        existing_equipment = c.fetchone()

        if not existing_equipment:
            messagebox.showerror(title='Erro', message='Equipamento não encontrado')
            return

        # Obtenha informações atuais do equipamento
        current_sector = existing_equipment[5]
        current_responsible = existing_equipment[6]
        current_status = existing_equipment[4]
        current_description = existing_equipment[3]
        current_quantity = existing_equipment[2]

        # Registre as informações anteriores na tabela de movimentações
        c.execute("""
            INSERT INTO movimentacoes (wks, setor_anterior, responsavel_anterior, status_anterior, descricao_anterior, quantidade_anterior)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (wks, current_sector, current_responsible, current_status, current_description, current_quantity))

        # Atualize todas as informações do equipamento no banco de dados
        c.execute("UPDATE inventario SET setor=?, responsavel=?, status=?, descricao=?, quantidade=? WHERE wks=?",
                  (new_sector, new_responsible, new_status, new_description, new_quantity, wks))
        conn.commit()

        messagebox.showinfo(title='Sucesso', message='Informações do equipamento atualizadas com sucesso')

        # Limpe os campos após a atualização
        self.clear_fields()

        # Atualize a contagem de inventário após a atualização
        self.update_inventory_count()

    def show_movement_history(self):
        # Obtenha o número de série (WKS) do equipamento a ser consultado
        wks = self.wks_entry.get()

        # Verifique se o número de série (WKS) foi fornecido
        if not wks:
            messagebox.showerror(title='Erro', message='Insira o WKS do equipamento para consultar o histórico')
            return

        # Use a função get_movement_history para recuperar o histórico de movimentações
        history = self.get_movement_history(wks)

        # Exiba o histórico de movimentações em uma nova janela
        self.show_history_popup(history)

    def show_history_popup(self, history):
        # Criar uma nova janela para exibir o histórico de movimentações
        history_window = tk.Toplevel(self.root)
        history_window.title('Histórico de Movimentações')

        # Criar uma área de texto para exibir o histórico
        history_text = tk.Text(history_window, wrap=tk.WORD)
        history_text.pack()

        # Preencher a área de texto com as informações do histórico
        if history:
            for entry in history:
                history_text.insert(tk.END,
                                    f'Data: {entry[7]}, Setor Anterior: {entry[2]}, Responsável Anterior: {entry[3]}, Status Anterior: {entry[4]}, Descrição Anterior: {entry[5]}, Quantidade Anterior: {entry[6]}\n\n')
        else:
            history_text.insert(tk.END, 'Nenhuma movimentação encontrada para este equipamento.')

        history_text.config(state=tk.DISABLED)  # Impede a edição do texto

        # Fechar a janela popup quando o usuário clicar em "Fechar"
        close_button = tk.Button(history_window, text='Fechar', command=history_window.destroy)
        close_button.pack()

    ####################################################################################################################


    def search_equipment(self):
        # Criar uma nova janela para os resultados da pesquisa
        search_results_window = tk.Toplevel(self.root)
        search_results_window.title('Resultados da Pesquisa')

        # Criar e exibir rótulos e campos de entrada para os critérios de pesquisa
        tk.Label(search_results_window, text='Responsável:').grid(row=0, column=0)
        tk.Label(search_results_window, text='Setor:').grid(row=1, column=0)
        tk.Label(search_results_window, text='Descrição:').grid(row=2, column=0)
        tk.Label(search_results_window, text='WKS:').grid(row=3, column=0)

        search_responsible_entry = tk.Entry(search_results_window)
        search_responsible_entry.grid(row=0, column=1)

        search_sector_entry = tk.Entry(search_results_window)
        search_sector_entry.grid(row=1, column=1)

        search_description_entry = tk.Entry(search_results_window)
        search_description_entry.grid(row=2, column=1)

        search_wks_entry = tk.Entry(search_results_window)
        search_wks_entry.grid(row=3, column=1)

        # Criar e exibir um botão para iniciar a pesquisa
        search_button = tk.Button(search_results_window, text='Pesquisar', command=lambda: self.perform_search(
            search_responsible_entry.get(), search_sector_entry.get(), search_description_entry.get(),
            search_wks_entry.get()))
        search_button.grid(row=4, column=0, columnspan=2, pady=10)

    def perform_search(self, responsible, sector, description, wks):
        # Construir e executar a consulta SQL
        sql = "SELECT * FROM inventario WHERE "

        conditions = []
        parameters = []

        if responsible:
            conditions.append("responsavel = ?")
            parameters.append(responsible)

        if sector:
            conditions.append("setor = ?")
            parameters.append(sector)

        if description:
            conditions.append("descricao LIKE ?")
            parameters.append('%' + description + '%')

        if wks:
            conditions.append("wks = ?")
            parameters.append(wks)

        if conditions:
            sql += " AND ".join(conditions)

        c.execute(sql, parameters)

        # Exibir resultados em uma nova janela
        records = c.fetchall()
        search_results_window = tk.Toplevel(self.root)
        search_results_window.title('Resultados da Pesquisa')

        # Crie um rótulo ou uma área de texto para exibir os resultados
        result_text = tk.Text(search_results_window, wrap=tk.WORD)
        result_text.pack()

        # Adicione os resultados ao rótulo ou área de texto
        for record in records:
            result_text.insert(tk.END,
                               f'WKS: {record[0]}, Nome: {record[1]}, Quantidade: {record[2]}, Descrição: {record[3]}, Status: {record[4]}, Setor: {record[5]}, Responsável: {record[6]}')
#########################################vvvvv 05/10/2023 vvvvv#####################################
    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[('Planilhas Excel', '*.xlsx')])

        if not file_path:
            return  # O usuário cancelou a seleção de arquivo

        try:
            # Lê o arquivo Excel
            df = pd.read_excel(file_path)

            # Itera sobre as linhas do DataFrame e insere os equipamentos no banco de dados
            for index, row in df.iterrows():
                wks = str(row['WKS'])
                name = str(row['Nome'])
                quantity = int(row['Quantidade'])
                description = str(row['Descrição'])
                status = str(row['Status'])
                sector = str(row['Setor'])
                responsible = str(row['Responsável'])

                # Validação de dados
                if not wks or not name or not quantity or quantity == 0 or not description or not status or not sector:
                    continue  # Pula linhas com dados incompletos

                # Verifica se o equipamento com o mesmo WKS já existe
                c.execute("SELECT * FROM inventario WHERE wks=?", (wks,))
                existing_equipment = c.fetchone()

                if not existing_equipment:
                    # O equipamento não existe, insere um novo registro
                    c.execute("""INSERT INTO inventario (wks, nome, quantidade, descricao, status, setor, responsavel)
                                 VALUES (?, ?, ?, ?, ?, ?, ?)""",
                              (wks, name, quantity, description, status, sector, responsible))

            conn.commit()
            messagebox.showinfo(title='Sucesso', message='Equipamentos cadastrados com sucesso!')

            # Atualiza a contagem de inventário após o carregamento em massa
            self.update_inventory_count()

        except Exception as e:
            messagebox.showerror(title='Erro', message=f"Erro ao carregar a planilha: {str(e)}")

    #########################################^^^^^ 05/10/2023 ^^^^^#####################################

    ###############vvvvv################ 11/10/2023 ######################vvvvv##########
    def delete_all_records(self):
        confirm = messagebox.askyesno(title='Confirmar',
                                      message='Tem certeza de que deseja apagar todos os registros? Esta ação é irreversível.')
        if confirm:
            c.execute("DELETE FROM inventario")  # Apaga todos os registros da tabela 'inventario'
            conn.commit()
            self.update_inventory_count()  # Atualize a contagem de inventário
            messagebox.showinfo(title='Sucesso', message='Todos os registros foram apagados com sucesso.')



    ###############^^^^^################ 11/10/2023 ######################^^^^^##########


    def run(self):
        if login_system.logged_in:  # Check if logged in before running the inventory system
            self.root.mainloop()
        else:
            self.root.destroy()  # Destroy the inventory system if not logged in



# Run the login system
login_system = LoginSystem()
login_system.run()

# Create GUI if logged in successfully
if login_system.logged_in:
    inventory_system = InventorySystem()
    inventory_system.run()