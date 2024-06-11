import sqlite3
import pandas as pd
import datetime
import sqlite3 
import json
import openpyxl
#import warnings
#import os
#import sys
from tkinter import *
from tkinter import simpledialog
import tkinter as tk
#from tkinter import ttk
from tkinter import messagebox, filedialog, ttk
from package import email

# Lista mensagens
def add_mess_list():
    # Verifica se existe uma coluna selecionanda na listbox colunas e se o registo não existe na listbox das mensagens
    if columns_list.get(ANCHOR) and columns_list.get(ANCHOR) not in list(mess_list.get(0, END)):

        selected_index = columns_list.curselection()
        #print(selected_index )
        if selected_index:
            selected_item = columns_list.get(selected_index)
            new_name = simpledialog.askstring("Atribuir nome", f"Nome para coluna {selected_item}:")

            if new_name:
                mess_list.insert(END, columns_list.get(ANCHOR))
                mess_list_names.delete(selected_index)
                mess_list_names.insert(selected_index, f"{new_name}")

def delete_mess_list():
    # Para agapar convem ter seleção ativa para não gerar erro
    if mess_list.curselection():
        index = mess_list.curselection()
        mess_list.delete(ANCHOR)
        
        mess_list_names.delete(index)
        #print(index)

def change_name_mess_list():
    # Obter o índice do item selecionado
    selected_index = mess_list_names.curselection()
    if selected_index:
        selected_item = mess_list_names.get(selected_index)
        new_name = simpledialog.askstring("Atribuir nome", f"Nome para coluna {selected_item}:")
        if new_name:
            mess_list_names.delete(selected_index)
            mess_list_names.insert(selected_index, f"{new_name}")

def change_name_date_list():
    # Obter o índice do item selecionado
    selected_index = date_list_names.curselection()
    if selected_index:
        selected_item = date_list_names.get(selected_index)
        new_name = simpledialog.askstring("Atribuir nome", f"Nome para coluna {selected_item}:")
        if new_name:
            date_list_names.delete(selected_index)
            date_list_names.insert(selected_index, f"{new_name}")

    
# Lista datas
def add_date_list():
    # Verifica se existe uma coluna selecionanda na listbox colunas e se o registo não existe na listbox das datas
    if columns_list.get(ANCHOR) and columns_list.get(ANCHOR) not in list(date_list.get(0, END)):

        selected_index = columns_list.curselection()
        #print(selected_index )
        if selected_index:
            selected_item = columns_list.get(selected_index)
            new_name = simpledialog.askstring("Atribuir nome", f"Nome para coluna {selected_item}:")

            if new_name:
                date_list.insert(END, columns_list.get(ANCHOR))

                date_list_names.delete(selected_index)
                date_list_names.insert(selected_index, f"{new_name}")

def delete_date_list():
    # Para agapar convem ter seleção ativa para não gerar erro
    if date_list.curselection():
        index = date_list.curselection()
        date_list.delete(ANCHOR)
        
        date_list_names.delete(index)
        #print(index)

# Função de atualiza
def update(id_registro):
 
    if not mess_list.size():
        messagebox.showerror("Erro", "Lista de mensagens vazia.")
        return
    
    if not date_list.size():
        messagebox.showerror("Erro", "Lista de datas vazia.")
        return
    
    if not columns_list.size():
        messagebox.showerror("Erro", "Lista de colunas vazia.")
        return
    
    if not txt_delta.get():
        messagebox.showerror("Erro", "Campo 'Delta (dias)' vazio.")
        return

    # Codificar listas para JSON
    serialized_mess = json.dumps(list(mess_list.get(0, END)))
    serialized_mess_names = json.dumps(list(mess_list_names.get(0, END)))
    serialized_dates = json.dumps(list(date_list.get(0, END)))
    serialized_dates_names = json.dumps(list(date_list_names.get(0, END)))
    serialized_sheets =  json.dumps(list(sheet_combobox["values"]), ensure_ascii=False, default=str)
    #serialized_selected_sheet = sheet_combobox.get()
   
   
    # Create a database or connect to one
    conn = sqlite3.connect('database.db')
    # Create cursor
    c = conn.cursor() 

    c.execute("""UPDATE registros 
                SET 
                        ficheiro = :filename,
                        mess_fields = :mess,
                        mess_fields_names = :mess_names,
                        date_fields = :dates,
                        date_fields_names = :dates_names,
                        columns = :columns,
                        path = :directorio,
                        sheets = :serialized_sheets,
                        sheet = :selected_sheet,
                        delta = :txt_delta,
                        check_alert = :checkbox_value
                WHERE 
                        id=:id""",
                {
                        'filename': filename.get(),
                        'mess': serialized_mess,
                        'mess_names': serialized_mess_names, 
                        'dates_names': serialized_dates_names,
                        'dates': serialized_dates,
                        'columns': columns_list.size(),
                        'directorio': filepath.get(),
                        'txt_delta': txt_delta.get(),
                        'serialized_sheets': serialized_sheets,
                        'selected_sheet': sheet_combobox.get(),
                        'checkbox_value': checkbox_var.get(),
                        'id': id_registro
                } 
                )

   # Commit Changes
    conn.commit()

    # Close Connection
    conn.close() 

    messagebox.showinfo("Success", "Dados submetidos com sucesso!")

    app.carregar_registros()

def open_file():

    file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsm;*.xlsx;*.xls")])
    
    if file_path:
        try:
            workbook = openpyxl.load_workbook(file_path)
            # Access the sheets, cells, or perform other operations as needed
            sheet_names = workbook.sheetnames
            sheet_combobox["values"] = []
            sheet_combobox.set("")  
            sheet_combobox["values"] = sheet_names
            sheet_combobox.set(sheet_combobox['values'][0])
                       
            df = pd.read_excel(file_path, sheet_name=sheet_combobox.get())

            columns_list.delete(0, END)
            for col in range(len(df.columns)):
                 columns_list.insert(END, col+1)

            print(f"Opened Excel file with sheet names: {sheet_names}")
        except Exception as e:
            print(f"Error opening Excel file: {e}")

        # Limpa campo
        filepath.delete(0, END)
        # Insere caminho
        filepath.insert(0, file_path)

def on_checkbox_checked():
   # global checkbox_value
    # if checkbox_var.get() == 1:
    #     checkbox_value = 1
    # else:
    #     checkbox_value = 0
    return

def load_selected_sheet(event):
        df = pd.read_excel(filepath.get(), sheet_name=sheet_combobox.get())

        columns_list.delete(0, END)
        for col in range(len(df.columns)):
            columns_list.insert(END, col+1)

def excel_viewer():

    #new_window = Toplevel(root)
    app = ExcelViewerApp(root)

class ExcelViewerApp:
    def __init__(self, root):
        # Pandas DataFrame
        self.df = pd.DataFrame()  # DataFrame to store Excel data
        self.selected_columns = []  # List to store selected columns
        # File path
        if  filepath.get():
                # Valor selecionado na sheet
                self.df = pd.read_excel(filepath.get(), sheet_name=sheet_combobox.get())

                # Clear and update the listbox with column names
                columns_list.delete(0, END)
                for col in range(len(self.df.columns)):
                    columns_list.insert(END, col+1)

                # Clean list
                selected_indices = [] 

                selected_indices = mess_list.get(0, "end")
                adjusted_indices = [int(item) - 1 for item in selected_indices]

                try:
                    self.mess_columns = [self.df.columns[i] for i in adjusted_indices]
                except IndexError as e:
                    messagebox.showerror("Erro", "Vefique se as colunas selecionadas existem.")
                    self.top.destroy()
                    return

                color = "#0ad1c6"

                self.top = Toplevel(root)
                self.top.title("Selected Columns")    
                self.top.grab_set()  # Make the window modal   
                # Bind the on_close_top function to the "WM_DELETE_WINDOW" event of the Toplevel window
                self.top.protocol("WM_DELETE_WINDOW", self.on_close_top)    

                # Display selected mess columns 
                if self.mess_columns:

                    for col in self.mess_columns:
                        frame = Frame(self.top, bg=color)
                        frame.pack(side=LEFT, padx=5, pady=5, fill=BOTH, expand=True)

                        label = Label(frame, text=col, bg=color, padx=5, pady=5)
                        label.pack()

                        text_widget = Text(frame, height=len(self.df), width=40)
                        text_widget.insert(END, self.df[col].to_string(index=False))
                        text_widget.pack()

      
                # Clean list
                selected_indices = []
                # Mensagens
                selected_indices = date_list.get(0, "end")
                adjusted_indices = [int(item) - 1 for item in selected_indices]

                try:
                    self.dates_columns = [self.df.columns[i] for i in adjusted_indices]
                except IndexError as e:
                    messagebox.showerror("Erro", "Vefique se as colunas selecionadas existem.")
                    self.top.destroy()
                    return

                color = "#ffd866"
        
                # Display selected dates  columns 
                if self.dates_columns:

                    for col in self.dates_columns:
                        #print(col)
                        frame = Frame(self.top, bg=color)
                        frame.pack(side=LEFT, padx=5, pady=5, fill=BOTH, expand=True)

                        label = Label(frame, text=col, bg=color, padx=5, pady=5)
                        label.pack()

                        text_widget = Text(frame, height=len(self.df), width=40)
                        text_widget.insert(END, self.df[col].to_string(index=False))
                        text_widget.pack()

    # Define on_close_top as an instance method
    def on_close_top(self):
        self.top.grab_release()  # Release the grab to allow interaction with other windows
        self.top.destroy()  # Destroy the Toplevel window
        self.top = None  # Reset the reference
        form_frame.grab_set() 

def create_html_table(data):
    html = "<table border='1'>\n"

    # Adiciona os cabeçalhos de coluna
    html += "<tr>\n"
    html += "<th></th>\n"  # Add line number column header
    for column_title in data[0]:
        html += f"<th>{column_title}</th>\n"
    html += "</tr>\n"

    # Adiciona os dados restantes com números de linha
    for line_number, row in enumerate(data[1:], start=1):
        html += "<tr>\n"
        html += f"<td>{line_number}</td>\n"  # Add line number
        for col in row:
            html += f"<td>{col}</td>\n"
        html += "</tr>\n"

    html += "</table>"

    return html

def execute():

    if not mess_list.size():
        messagebox.showerror("Erro", "Lista de mensagens vazia.")
        return
    
    if not date_list.size():
        messagebox.showerror("Erro", "Lista de datas vazia.")
        return
    
    if not columns_list.size():
        messagebox.showerror("Erro", "Lista de colunas vazia.")
        return
    
    if not txt_delta.get():
        messagebox.showerror("Erro", "Campo 'Delta (dias)' vazio.")
        return

    matriz = []

    #print(filename.get())

    # Read the Excel file into a DataFrame
    df_base = pd.read_excel(filepath.get(), sheet_name=sheet_combobox.get())
    # Sem duplicadas
    df = df_base.drop_duplicates()

    actual_date = datetime.datetime.now().date()
    #print(date_list.get(0, "end"))
    count_line = 0
    for i in date_list.get(0, "end"):
        #print(type(date_list.get(0, "end")))

        for index, row in df.iterrows():
            #print(type(row[df.columns[i-1]]))
            if not pd.isnull(row[df.columns[i-1]]) and (type(row[df.columns[i-1]]) == datetime.datetime  or isinstance(row[df.columns[i-1]], pd.Timestamp)): # Em alguns casos aparece formatos diferentes??????
                date_time = pd.to_datetime(row[df.columns[i-1]])  # Convert the column to datetime format
                # Data da validade
                duedate = date_time.date()
                
                # Calcula diferença
                differece = (duedate - actual_date).days

                # Seleciona itens existentes na listbox das mensagens
                selected_indices = mess_list.get(0, "end")
                lst_to_add = []
                list_header = []

                if differece < int(txt_delta.get()):

                    # Para adicionar o none das colunas
                    for index, item in enumerate(selected_indices):
                        list_header.append(mess_list_names.get(index))

                        #print(type(row.iloc[item-1]))
                        # Algumas colunas de data são lidas como livraria dos pandas e outras data como livraria standard do Python
                        if isinstance(row.iloc[item-1], pd.Timestamp):
                           
                            # Convertendo o Timestamp para uma string no formato desejado
                            timestamp = pd.Timestamp(row.iloc[item-1])
                            dt_without_timestamp = timestamp.strftime("%Y-%m-%d")
                            lst_to_add.append(dt_without_timestamp)
                            #print(dt_without_timestamp)

                        elif type(row.iloc[item-1]) == datetime.datetime:

                            dt_without_timestamp = row.iloc[item-1].date()
                            lst_to_add.append(dt_without_timestamp)
                            #print(dt_without_timestamp)
                        else:
                            lst_to_add.append(row.iloc[item-1])
                
                    if count_line == 0:
                        matriz.append(list_header)
                        #print(list_header)

                    count_line += 1
                    matriz.append(lst_to_add)
            else:
                continue
               
    if matriz:
        messagebox.showinfo("Information", f"{len(matriz)-1} linhas enviadas.")
        html_table = create_html_table(matriz)

        # Envio de email
        email.send_message(html_table, filename.get())

def quit_application():
    result = messagebox.askokcancel("Close", "Deseja fechar aplicação?")
    if result:
        root.destroy()

def on_close():
    result = messagebox.askokcancel("Close", "Deseja fechar aplicação?")
    if result:
        root.destroy()

#class NovoRegistroPopup:

def form_new(parent, registro):
    # SET FRAMES
    # Frame for the form
    global form_frame
    form_frame = tk.Toplevel(parent)
    form_frame.title("Novo registo")
    form_frame.grab_set()  # Make the window modal

    # Top Frame 
    top_frame = Frame(form_frame, padx=10, pady=10)
    top_frame.grid(row=0, column=0, sticky="w")

    # Botton Frame 
    botton_frame = Frame(form_frame, padx=10, pady=10)
    botton_frame.grid(row=1, column=0, sticky="w")

    # Text box
    global filepath
    filepath = Entry(top_frame, width=72)
    filepath.grid(row=0, column=1, padx=10, sticky="w")

    filepath_label = Label(top_frame, text='Directório:')
    filepath_label.grid(row=0, column=0, padx=10, sticky="w")

    global filename
    filename = Entry(top_frame, width=40)
    filename.grid(row=1, column=1, padx=10, pady=10, sticky="w")

    filename_label = Label(top_frame, text='Ficheiro:')
    filename_label.grid(row=1, column=0, padx=10, sticky="w")

    # Botões
    btn_openfile = Button(top_frame, text="...", width=2 ,command=open_file)
    btn_openfile.grid(row=0, column=2,  padx=10, sticky="")

    global sheet_combobox
    sheet_combobox = ttk.Combobox(botton_frame, state='readonly')
    sheet_combobox.grid(row=0, column=1, padx=18, sticky="wn")

    sheet_combobox_label = Label(botton_frame, text='Sheet:')
    sheet_combobox_label.grid(row=0, column=0, padx=17, sticky="wn")

    # Set an event handler for Combobox selection change
    sheet_combobox.bind("<<ComboboxSelected>>", load_selected_sheet)

    # List Boxes
    global columns_list
    columns_list_label = Label(botton_frame, text="Colunas do ficheiro:")
    columns_list_label.grid(row=0, column=2, padx=20, sticky="wn")

    columns_list = Listbox(botton_frame, width=20, height=10)
    columns_list.grid(row=0, column=3, sticky="w")


    btn_update = Button(form_frame, width=12, text="Gravar", command=adicionar_registo)
    btn_update.grid(row=3, column=0, padx=10, pady=10)

def adicionar_registo():

    if not filepath.get():
        messagebox.showerror("Erro", "Campo do directório vazio.")
        return
    
    if not filename.get():
        messagebox.showerror("Erro", "Campo do ficheiro vazio.")
        return
    
    if not columns_list.size():
        messagebox.showerror("Erro", "Lista de colunas vazia.")
        return
    
    if not sheet_combobox.get():
        messagebox.showerror("Erro", "Sheet vazia.")
        return   
    
    app.adicionar_registo(filename.get(), sheet_combobox["values"], sheet_combobox.get(), columns_list.size(), filepath.get())
    #form_frame.destroy()

def form_details(parent, registro):
    id_registro = registro[0]

    # Verifica se já existe uma janela aberta para este registro
    if id_registro in janelas_abertas:
        return

    # SET FRAMES
    # Frame for the form
    global form_frame
    form_frame = tk.Toplevel(parent)
    form_frame.title("Detalhes do Registro")
    form_frame.grab_set()  # Make the window modal

    # Frame Direito
    right_frame = Frame(form_frame, padx=10, pady=0)
    right_frame.grid(row=0, column=1, sticky="n")

    # Frame Esquerdo
    left_frame = LabelFrame(form_frame, padx=10, pady=10)
    left_frame.grid(row=0, column=0, sticky="w")

    # Top Frame (Main Left Frame)
    top_frame = Frame(left_frame, padx=10, pady=10)
    top_frame.grid(row=0, column=0, sticky="w")

    # Botton Frame (Main Left Frame)
    botton_frame = Frame(left_frame)
    botton_frame.grid(row=1, column=0)

    botton_frame_left = Frame(botton_frame, padx=10, pady=10)
    botton_frame_left.grid(row=0, column=0, sticky="w")

    botton_frame_right = Frame(botton_frame, padx=10, pady=10)
    botton_frame_right.grid(row=0, column=1)

###################################### Top frame #########################################

    # Text box
    global filepath
    filepath = Entry(top_frame, width=75)
    filepath.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

    filepath_label = Label(top_frame, text='Directório:')
    filepath_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    # Botões
    btn_openfile = Button(top_frame, text="...", width=2 ,command=open_file)
    btn_openfile.grid(row=0, column=2,  padx=10, sticky="")

    global filename
    filename = Entry(top_frame, width=40)
    filename.grid(row=1, column=1, padx=10, pady=10, sticky="w")

    filename_label = Label(top_frame, text='Ficheiro:')
    filename_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

    # Variável para armazenar o estado da caixa de seleção
    global checkbox_var
    checkbox_var = IntVar()
    # Criar a caixa de seleção
    checkbox = Checkbutton(top_frame, text="Alerta", variable=checkbox_var, command=on_checkbox_checked)
    checkbox.grid(row=1, column=2, pady=10, sticky="w")

    # Sheet
    # global txt_sheet
    # txt_sheet = Entry(top_frame, width=40)
    # txt_sheet.grid(row=2, column=1, padx=10, sticky="w")

    # txt_sheet_label = Label(top_frame, text='Sheet:')
    # txt_sheet_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")

    global sheet_combobox
    sheet_combobox = ttk.Combobox(top_frame, state='readonly')
    sheet_combobox.grid(row=2, column=1, padx=10, sticky="w")

    sheet_combobox_label = Label(top_frame, text='Sheet:')
    sheet_combobox_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")

     # Set an event handler for Combobox selection change
    sheet_combobox.bind("<<ComboboxSelected>>", load_selected_sheet)

    # Campo delta (dias)
    global txt_delta
    txt_delta = Entry(top_frame, width=10)
    txt_delta.grid(row=3, column=1, padx=10, pady=10, sticky="w")

    txt_delta_label = Label(top_frame, text='Delta (dias):')
    txt_delta_label.grid(row=3, column=0, padx=10, pady=10)

    ################################### Botton Left Frame ####################################

    global mess_list
    mess_list = Listbox(botton_frame_left, width=10)
    mess_list.grid(row=1, column=1, padx=10,  sticky="w")

    mess_fields_label = Label(botton_frame_left, text="Nome das colunas - Mensagens (azul):")
    mess_fields_label.grid(row=0, column=0, padx=5, sticky="w")

    global mess_list_names
    mess_list_names = Listbox(botton_frame_left, width=30)
    mess_list_names.grid(row=1, column=0, padx=10,  sticky="w")

    # Button to change the name of the selected item
    btn_change_mess_list = Button(botton_frame_left, text="Alterar nome",  command=change_name_mess_list)
    btn_change_mess_list.grid(row=2, column=0, padx=10,  sticky="n")

    # Botões
    btn_add_mess = Button(botton_frame_left, text="<<", command=add_mess_list)
    btn_add_mess.grid(row=1, column=2, padx=10, pady=10)

    btn_del_mess = Button(botton_frame_left, text="Eliminar", command=delete_mess_list)
    btn_del_mess.grid(row=1, column=3, padx=10, pady=10)

    # List boxes
    date_fields_label = Label(botton_frame_left, text="Nome das colunas - Datas (amarelo):")
    date_fields_label.grid(row=3, column=0, padx=5, sticky="ws")

    global date_list
    date_list = Listbox(botton_frame_left, width=10)
    date_list.grid(row=4, column=1, padx=10,  sticky="w")

    global date_list_names
    date_list_names = Listbox(botton_frame_left, width=30)
    date_list_names.grid(row=4, column=0, padx=10,  sticky="wn")


    # Button to change the name of the selected item
    btn_change_date_list = Button(botton_frame_left, text="Alterar nome",  command=change_name_date_list)
    btn_change_date_list.grid(row=5, column=0, padx=10,  sticky="n")

    # Botões
    btn_add_date = Button(botton_frame_left, text="<<", command=add_date_list)
    btn_add_date.grid(row=4, column=2, padx=10, pady=10)

    btn_del_date = Button(botton_frame_left, text="Eliminar", command=delete_date_list)
    btn_del_date.grid(row=4, column=3, padx=10, pady=10)

    ################################### Botton Right Frame ####################################

    # List Boxes
    global columns_list
    columns_list_label = Label(botton_frame_right, text="Colunas do ficheiro:")
    columns_list_label.grid(row=0, column=0, padx=5, sticky="w")

    columns_list = Listbox(botton_frame_right, width=20, height=25)
    columns_list.grid(row=1, column=0, padx=10, sticky="sn")

    ####################################### Right Frame #######################################

    btn_update = Button(right_frame, width=12, text="Atualizar dados", command=lambda: update(id_registro))
    btn_update.grid(row=0, column=0, padx=10, pady=10)

    btn_display = Button(right_frame, width=12, text="Mostrar seleção", command=excel_viewer)
    btn_display.grid(row=1, column=0, padx=10, pady=10)

    btn_execute = Button(right_frame, width=12, text="Executar", command=execute)
    btn_execute.grid(row=2, column=0, padx=10, pady=10)

    # btn_close = Button(right_frame, width=12, text="Fechar", command=quit_application)
    # btn_close.grid(row=3, column=0, padx=10, pady=10)

    # Insere os valores
    query(registro)

    # Adiciona esta janela ao dicionário de janelas abertas
    janelas_abertas[id_registro] = form_frame

    def on_close_form():
        janelas_abertas.pop(id_registro, None)  # Remove a janela fechada do dicionário
        form_frame.destroy()  # Fecha a janela do formulário

    form_frame.protocol("WM_DELETE_WINDOW", on_close_form)  # Define o evento de fechamento

# Create Query function
def query(registro):
    # Inserir dados nos form fields 
    filename.insert(0, str(registro[1]))
    #filename.configure(state='readonly')
    filepath.insert(0, str(registro[9]))

    # Campo Sheet
    # txt_sheet.insert(0, str(registro[2]))
    # txt_sheet.configure(state='readonly')
    # Campos da combobox com sheets
    stored_sheets = json.loads(registro[2])
    sheet_combobox["values"] = stored_sheets
    sheet_combobox.set(registro[3])
    checkbox_var.set(registro[11])

    # global checkbox_value
    # checkbox_value = registro[11]

    # if checkbox_value is None:   
    #     checkbox_value = 0


    # Campos de mensagem
    if registro[4]:
        stored_mess = json.loads(registro[4])
        for item in stored_mess:
            mess_list.insert(END, item)

    # Nomes da coluna mensagem
    if registro[5]:        
        stored_mess_names = json.loads(registro[5])
        for item in stored_mess_names:
            mess_list_names.insert(END, item)

    # Campos de data
    if registro[6]:  
        stored_dates = json.loads(registro[6])
        for item in stored_dates:
            date_list.insert(END, item)

    # Nomes daa coluna data
    if registro[7]:
        stored_dates_names = json.loads(registro[7])
        for item in stored_dates_names:
            date_list_names.insert(END, item)
    
    # O número de coluna começa no index 0
    number_columns = registro[8]
    for index in range(number_columns):
        columns_list.insert(END, index+1)
    
    # Campo Delta (dias)
    if registro[10]:
        txt_delta.insert(0, str(registro[10]))

    # Campo directório do ficheiro
    # global file_path
    # file_path = registro[8]

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Visualizador de Registros")

        self.conn = sqlite3.connect("database.db")
        self.cursor = self.conn.cursor()

        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=10, pady=10, fill="both", expand=True)
        #self.frame.pack_propagate(False)  # Disable frame resizing based on its content

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(self.frame, textvariable=self.search_var, width=53)
        self.search_entry.insert(0, "Procurar...")  # Insert placeholder text
        self.search_entry.pack(side="top", padx=5, pady=5, anchor="w")
        self.search_entry.bind("<KeyRelease>", self.search)
        self.search_entry.bind("<FocusIn>", self.clear_placeholder)
        self.search_entry.bind("<FocusOut>", self.restore_placeholder)

        self.search_coluna_var = tk.StringVar()
        self.search_coluna_combobox = ttk.Combobox(self.frame, textvariable=self.search_coluna_var, width=50, state='readonly')
        self.search_coluna_combobox.pack(side="top", padx=5, pady=5, anchor="w")
        self.search_coluna_combobox['values'] = ('Ficheiro', 'Sheet')  # Colunas disponíveis para pesquisa
        self.search_coluna_combobox.set('Ficheiro')  # Coluna padrão selecionada

        self.tree = ttk.Treeview(self.frame, columns=("Ficheiro", "Sheet", "Columns", "Path", "Check_alert"), show="headings")
        self.tree.heading("Ficheiro", text="Ficheiro")
        self.tree.heading("Sheet", text="Sheet")
        self.tree.heading("Columns", text="Colunas")
        self.tree.heading("Path", text="Directório")
        self.tree.heading("Check_alert", text="Alerta")
        self.tree.column("Ficheiro", width=300)     # Adjust width as needed
        self.tree.column("Sheet", width=300)        # Adjust width as needed
        self.tree.column("Columns", width=100)      # Adjust width as needed
        self.tree.column("Path", width=600)         # Adjust width as needed
        self.tree.column("Check_alert", width=50)   # Adjust width as needed
        self.tree.pack(side="top", fill="both", expand=True, anchor="w")

        # self.scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
        # self.scrollbar.pack(side="right", fill="y")
        # self.tree.configure(yscrollcommand=self.scrollbar.set)

        # self.scrollbar_horizontal = ttk.Scrollbar(self.frame, orient="horizontal", command=self.tree.xview)
        # self.scrollbar_horizontal.pack(side="bottom", fill="x")
        # self.tree.configure(xscrollcommand=self.scrollbar_horizontal.set)

        self.botao_novo_registro = tk.Button(self.root, text="Novo registo", width=15, command=self.mostrar_novo_registro)
        #self.botao_novo_registro.pack(pady=10)
        self.botao_novo_registro.pack(side="left", padx=(10,5), pady=10)

        self.botao_deletar_registro = tk.Button(self.root, text="Apagar registo", width=15, command=self.delete_selected_registro)
        #self.botao_deletar_registro.pack(pady=10)
        self.botao_deletar_registro.pack(side="left", padx=(5,5), pady=10)

        #self.botao_execute_all = ttk.Button(self.root, text="Executar", width=15, command=self.execute_all)
        command_executed = True
        self.botao_execute_all = tk.Button(self.root, text="Executar", width=15, command=lambda: self.execute_all(command_executed))

        self.botao_execute_all.pack(side="right", padx=(5,10), pady=10)

        self.carregar_registros()

        self.tree.bind("<Double-1>", self.mostrar_detalhes)

    def clear_placeholder(self, event):
        if self.search_var.get() == "Procurar...":
            self.search_entry.delete(0, tk.END)

    def restore_placeholder(self, event):
        if not self.search_var.get():
            self.search_entry.insert(0, "Procurar...")

    def carregar_registros(self):
        self.tree.delete(*self.tree.get_children())
        self.cursor.execute("SELECT * FROM registros")
        registros = self.cursor.fetchall()

        for registro in registros:
            self.tree.insert("", "end", text=registro[0], values=(registro[1], registro[3], registro[8], registro[9], registro[11]))

    def mostrar_detalhes(self, event):
        item = self.tree.selection()[0]
        id_registro = self.tree.item(item, "text")
        self.cursor.execute("SELECT * FROM registros WHERE id=?", (id_registro,))
        registro = self.cursor.fetchone()
        popup = form_details(root, registro)

    def mostrar_novo_registro(self):
        popup = form_new(self.root, self)

    def adicionar_registo(self, filename, sheets_combobox, selected_sheet, columns_list, filepath):  
        serialized_sheets =  json.dumps(list(sheet_combobox["values"]), ensure_ascii=False, default=str)
        self.cursor.execute("INSERT INTO registros (ficheiro, sheets, sheet, columns, path, check_alert) VALUES (?, ?, ?, ?, ?, ?)", (filename, serialized_sheets, selected_sheet, columns_list, filepath, 0))
        self.conn.commit()
        self.carregar_registros()

    def deletar_registro(self, id_registro):
        self.cursor.execute("DELETE FROM registros WHERE id=?", (id_registro,))
        self.conn.commit()
        self.carregar_registros()

    def delete_selected_registro(self):
        item = self.tree.selection()[0]
        id_registro = self.tree.item(item, "text")
        self.deletar_registro(id_registro)

    def search(self, event=None):
        search_term = self.search_var.get()
        coluna = self.search_coluna_var.get()
        self.tree.delete(*self.tree.get_children())
        self.cursor.execute(f"SELECT * FROM registros WHERE {coluna} LIKE ?", ('%' + search_term + '%',))
        registros = self.cursor.fetchall()
        for registro in registros:
            self.tree.insert("", "end", values=(registro[1], registro[3]))

    # Executa para todos os registos os alertas
    def execute_all(self, command_executed):

        #self.cursor.execute("SELECT * FROM registros")
        self.cursor.execute("SELECT * FROM registros WHERE check_alert = 1")
        registros = self.cursor.fetchall()

        for registro in registros:
            matriz = []
            # Directório
            path_file = registro[9]
            filename = registro[1]
            sheet = registro[3]
            date_list =  json.loads(registro[6])
            mess_list = json.loads(registro[4])
            mess_list_names = json.loads(registro[5])
            delta = registro[10]
            actual_date = datetime.datetime.now().date()

            # Read the Excel file into a DataFrame
            df_base = pd.read_excel(path_file, sheet_name=sheet)
            # Sem duplicadas
            df = df_base.drop_duplicates()
             
            count_line = 0
            for i in date_list:
                for index, row in df.iterrows():

                    if not pd.isnull(row[df.columns[i-1]]) and (type(row[df.columns[i-1]]) == datetime.datetime  or isinstance(row[df.columns[i-1]], pd.Timestamp)): # Em alguns casos aparece formatos diferentes??????
                        date_time = pd.to_datetime(row[df.columns[i-1]])  # Convert the column to datetime format
                        # Data da validade
                        duedate = date_time.date()
                        
                        # Calcula diferença
                        differece = (duedate - actual_date).days

                        # Seleciona itens existentes na listbox das mensagens (variável)
                        selected_indices = mess_list
                        lst_to_add = []
                        list_header = []

                        if differece < delta:

                            # Para adicionar o none das colunas
                            for index, item in enumerate(selected_indices):
                                list_header.append(mess_list_names[index])

                                # Algumas colunas de data são lidas como livraria dos pandas e outras data como livraria standard do Python
                                if isinstance(row.iloc[item-1], pd.Timestamp):     
                                    # Convertendo o Timestamp para uma string no formato desejado
                                    timestamp = pd.Timestamp(row.iloc[item-1])
                                    dt_without_timestamp = timestamp.strftime("%Y-%m-%d")
                                    lst_to_add.append(dt_without_timestamp)

                                elif type(row.iloc[item-1]) == datetime.datetime:
                                    dt_without_timestamp = row.iloc[item-1].date()
                                    lst_to_add.append(dt_without_timestamp)
                                else:
                                    lst_to_add.append(row.iloc[item-1])
                        
                            if count_line == 0:
                                matriz.append(list_header)
                                #print(list_header)

                            count_line += 1
                            matriz.append(lst_to_add)

                    else:
                        continue
                    
            if matriz:
                #messagebox.showinfo("Information", f"{len(matriz)-1} linhas enviadas.")
                html_table = create_html_table(matriz)
                
                #print(html_table)

                # Envio de email
                email.send_message(html_table, filename)

        if command_executed == True:
           messagebox.showinfo("Informação", "Análise de alertas executado.")

if __name__ == "__main__":
    #command_executed = True
    global janelas_abertas
    janelas_abertas = {}
    # Flag para indicar que botão do executar foi pressinado
    root = tk.Tk()
    app = App(root)
    root.mainloop()