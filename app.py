import os
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import ttkbootstrap as tb
from ttkbootstrap.widgets import DateEntry
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io
from PIL import Image, ImageTk

import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd  # Importando a biblioteca pandas

def carregar_tabelas():
    banco = "dados.db"
    conexao = sqlite3.connect(banco)
    cursor = conexao.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tabelas = [row[0] for row in cursor.fetchall()]
    conexao.close()
    return tabelas

def carregar_dados(tabela):
    banco = "dados.db"
    conexao = sqlite3.connect(banco)
    cursor = conexao.cursor()
    cursor.execute(f"SELECT * FROM {tabela};")
    dados = cursor.fetchall()
    colunas = [description[0] for description in cursor.description]
    conexao.close()
    return colunas, dados

def exportar_para_excel(tabela):
    colunas, dados = carregar_dados(tabela)
    
    # Criar um DataFrame do pandas
    df = pd.DataFrame(dados, columns=colunas)
    
    # Salvar o DataFrame em um arquivo Excel
    try:
        df.to_excel(f"{tabela}.xlsx", index=False)
        messagebox.showinfo("Sucesso", f"Tabela '{tabela}' exportada com sucesso para Excel!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar para Excel: {e}")

def abrir_tela_edicao():
    global colunas, tabela_selecionada  # Declarar variáveis globais
    tabela_selecionada = None

    def carregar_e_exibir():
        global tabela_selecionada, colunas  # Usar variáveis globais
        tabela_selecionada = combobox_tabelas.get()
        if tabela_selecionada:
            colunas, dados = carregar_dados(tabela_selecionada)
            
            # Limpar colunas e dados do Treeview
            treeview["columns"] = colunas
            treeview.delete(*treeview.get_children())
            
            # Configurar cabeçalhos
            for col in colunas:
                treeview.heading(col, text=col)
                
            # Inserir dados no Treeview
            for linha in dados:
                treeview.insert("", "end", values=linha)

    def habilitar_edicao(event):
        # Ativa o botão de editar registro quando uma linha é selecionada
        if treeview.selection():
            botao_editar_registro.config(state=tk.NORMAL)
            botao_exportar.config(state=tk.NORMAL)  # Habilitar o botão de exportação
        else:
            botao_editar_registro.config(state=tk.DISABLED)
            botao_exportar.config(state=tk.DISABLED)  # Desabilitar o botão de exportação

    def editar_registro():
        item = treeview.selection()[0]  # Obtém o item selecionado
        valores = treeview.item(item, 'values')
        janela_principal.iconify()

        janela_edicao = tk.Toplevel()
        largura = 600
        altura = 700
        largura_tela = janela_edicao.winfo_screenwidth()
        altura_tela = janela_edicao.winfo_screenheight()
        x = (largura_tela - largura) // 2
        y = (altura_tela - altura) // 2 - 30
        janela_edicao.geometry(f"{largura}x{altura}+{x}+{y}")


        try:
            janela_edicao.iconbitmap('Logo/logo.ico')
        except:
            pass
        janela_edicao.title("Gerenciador de Anexos")
        # Tentar carregar e redimensionar a imagem para exibição
        try:
            imagem = Image.open('Logo/logo.ico')  # Tente abrir o arquivo de imagem
            imagem = imagem.resize((200, 200), Image.LANCZOS)  # Redimensione conforme necessário
            imagem_tk = ImageTk.PhotoImage(imagem)  # Converta para PhotoImage

                    # Criando um label para exibir a imagem
            label_imagem = tk.Label(janela_edicao, image=imagem_tk)
            label_imagem.image = imagem_tk  # Mantendo a referência da imagem
            label_imagem.pack(pady=(10, 5))  
  # Adicionando um espaço ao redor da imagem

        except FileNotFoundError:
            print("Arquivo logo.ico não encontrado. A imagem não será exibida.")
        
        label_coluna = ttk.Label(janela_edicao, text="Selecione a coluna:")
        label_coluna.pack(pady=(10, 0))

        combobox_colunas = ttk.Combobox(janela_edicao, values=colunas)
        combobox_colunas.pack(fill='x', padx=10, pady=(0, 10))

        label_novo_valor = ttk.Label(janela_edicao, text="Novo Valor:")
        label_novo_valor.pack(pady=(10, 0))

        entry_novo_valor = ttk.Entry(janela_edicao)
        entry_novo_valor.pack(fill='x', padx=10, pady=(0, 10))
            
       

        
        


        def salvar_novo_valor():
            coluna_selecionada = combobox_colunas.get()
            novo_valor = entry_novo_valor.get()

            # Captura o valor atual da coluna que está sendo editada
            item = treeview.selection()[0]  # Obtém o item selecionado
            valores = treeview.item(item, 'values')
            valor_antigo = valores[colunas.index(coluna_selecionada)]  # Captura o valor atual

            banco = "dados.db"
            conexao = sqlite3.connect(banco)
            cursor = conexao.cursor()
            try:
                # Atualiza o registro com o novo valor
                cursor.execute(f"UPDATE {tabela_selecionada} SET {coluna_selecionada} = ? WHERE {coluna_selecionada} = ?", (novo_valor, valor_antigo))
                conexao.commit()
                messagebox.showinfo("Sucesso", "Registro atualizado com sucesso!")
            except sqlite3.OperationalError as e:
                messagebox.showerror("Erro", f"Erro ao atualizar o banco de dados: {e}")
            finally:
                conexao.close()

            # Fecha a janela de edição
            janela_edicao.destroy()

        botao_salvar = ttk.Button(janela_edicao, text="Salvar", command=salvar_novo_valor)
        botao_salvar.pack(pady=(10, 10))

    janela_edicao = tk.Toplevel()
    largura = 600
    altura = 700
    largura_tela = janela_edicao.winfo_screenwidth()
    altura_tela = janela_edicao.winfo_screenheight()
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2 - 30
    janela_edicao.geometry(f"{largura}x{altura}+{x}+{y}")

    try:
        janela_edicao.iconbitmap('Logo/logo.ico')
    except:
            pass
        # Tentar carregar e redimensionar a imagem para exibição
    try:
            imagem = Image.open('Logo/logo.ico')  # Tente abrir o arquivo de imagem
            imagem = imagem.resize((200, 200), Image.LANCZOS)  # Redimensione conforme necessário
            imagem_tk = ImageTk.PhotoImage(imagem)  # Converta para PhotoImage


                    # Criando um label para exibir a imagem
            label_imagem = tk.Label(janela_edicao, image=imagem_tk)
            label_imagem.image = imagem_tk  # Mantendo a referência da imagem
            label_imagem.pack(pady=(10, 5))  

            # Adicionando um espaço ao redor da imagem

    except FileNotFoundError:
        print("Arquivo logo.ico não encontrado. A imagem não será exibida.")
    # Combobox para selecionar a tabela
    label_tabela = ttk.Label(janela_edicao, text="Selecione a Tabela:")
    label_tabela.pack(pady=(10, 0))
    
    combobox_tabelas = ttk.Combobox(janela_edicao, values=carregar_tabelas())
    combobox_tabelas.pack(fill='x', padx=50, pady=(0, 10))
    
    botao_carregar = ttk.Button(janela_edicao, text="Carregar Dados", command=carregar_e_exibir)
    botao_carregar.pack(pady=(5, 10))
    
    # Treeview para exibir os dados
    treeview = ttk.Treeview(janela_edicao, show='headings')
    treeview.pack(fill='both', expand=True, padx=50, pady=(0, 10))

    # Bind para permitir a edição
    treeview.bind("<ButtonRelease-1>", habilitar_edicao)

    # Botão para editar registro
    botao_editar_registro = ttk.Button(janela_edicao, text="Editar Registro", command=editar_registro, state=tk.DISABLED)
    botao_editar_registro.pack(pady=(10, 10))

    # Botão para exportar para Excel
    botao_exportar = ttk.Button(janela_edicao, text="Exportar para Excel", command=lambda: exportar_para_excel(tabela_selecionada), state=tk.DISABLED)
    botao_exportar.pack(pady=(10, 10))

    # Scrollbars
    scrollbar_vertical = ttk.Scrollbar(janela_edicao, orient="vertical", command=treeview.yview)
    treeview.configure(yscroll=scrollbar_vertical.set)
    scrollbar_vertical.pack(side='right', fill='y')

    scrollbar_horizontal = ttk.Scrollbar(janela_edicao, orient="horizontal", command=treeview.xview)
    treeview.configure(xscroll=scrollbar_horizontal.set)
    scrollbar_horizontal.pack(side='bottom', fill='x')


# Função para conectar ao banco de dados
def export_conectar_db():
    return sqlite3.connect('dados.db')

# # Função para criar a tabela, se não existir
# def export_criar_tabela():
#     conn = export_conectar_db()
#     cursor = conn.cursor()
#     cursor.execute('''
#         CREATE TABLE IF NOT EXISTS anexos (
#             id INTEGER PRIMARY KEY AUTOINCREMENT,
#             nome TEXT,
#             tipo TEXT,
#             arquivo BLOB,
#             extensao TEXT,
#             task TEXT,
#             evento TEXT
#         )
#     ''')
#     conn.commit()
#     conn.close()

# Função para converter e salvar o arquivo no banco de dados
def export_inserir_arquivo(nome_arquivo, tipo, task, evento):
    conn = export_conectar_db()
    cursor = conn.cursor()
    
    if tipo in ['jpg', 'jpeg', 'png']:
        with Image.open(nome_arquivo) as img:
            img = img.convert('RGB')
            img_byte_array = io.BytesIO()
            img.save(img_byte_array, format='JPEG')
            blob_data = img_byte_array.getvalue()
    elif tipo == 'txt':
        pdf_byte_array = io.BytesIO()
        c = canvas.Canvas(pdf_byte_array, pagesize=letter)
        with open(nome_arquivo, 'r') as file:
            text = file.read()
            c.drawString(100, 750, text)
        c.save()
        blob_data = pdf_byte_array.getvalue()
    else:
        messagebox.showerror("Erro", "Formato não suportado.")
        return
    
    extensao = os.path.splitext(nome_arquivo)[1]
    cursor.execute('INSERT INTO anexos (nome, tipo, arquivo, extensao, task, evento) VALUES (?, ?, ?, ?, ?, ?)', 
                   (os.path.basename(nome_arquivo), tipo, blob_data, extensao, task, evento))
    conn.commit()
    conn.close()
    messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")

# Função para recuperar um arquivo do banco de dados
def export_recuperar_arquivo(id_arquivo):
    conn = export_conectar_db()
    cursor = conn.cursor()
    cursor.execute('SELECT tipo, arquivo, extensao FROM anexos WHERE id = ?', (id_arquivo,))
    result = cursor.fetchone()
    conn.close()
    return result

# Função para exportar todos os arquivos
def exportar_todos_arquivos():
    pasta_destino = filedialog.askdirectory(title="Escolha a pasta para exportar os anexos")
    if not pasta_destino:
        return  # Se o usuário cancelar, não faz nada

    conn = export_conectar_db()
    cursor = conn.cursor()
    cursor.execute('SELECT id, nome, extensao FROM anexos')
    arquivos = cursor.fetchall()
    
    for id_arquivo, nome_original, extensao in arquivos:
        tipo, blob_data, _ = export_recuperar_arquivo(id_arquivo)
        if blob_data:
            nome_arquivo_com_extensao = os.path.join(pasta_destino, nome_original)
            with open(nome_arquivo_com_extensao, 'wb') as file:
                file.write(blob_data)
    
    messagebox.showinfo("Sucesso", "Todos os arquivos foram exportados com sucesso!")
    conn.close()

# Função para exportar arquivos selecionados
def export_exportar_arquivos_selecionados():
    try:
        # Obter os índices dos itens selecionados na listbox
        selecionados = listbox.curselection()
        if not selecionados:
            messagebox.showwarning("Seleção inválida", "Por favor, selecione pelo menos um arquivo para exportar.")
            return
        
        # Solicitar a pasta de destino
        pasta_destino = filedialog.askdirectory(title="Escolha a pasta para exportar os anexos")
        if not pasta_destino:
            return  # Se o usuário cancelar, não faz nada

        for index in selecionados:
            # Obter o ID do arquivo selecionado
            id_arquivo = int(listbox.get(index).split(",")[0].split(":")[1].strip())
            tipo, blob_data, _ = export_recuperar_arquivo(id_arquivo)

            if blob_data:
                nome_original = listbox.get(index).split(",")[1].split(":")[1].strip()
                nome_arquivo_com_extensao = os.path.join(pasta_destino, nome_original)
                with open(nome_arquivo_com_extensao, 'wb') as file:
                    file.write(blob_data)

        messagebox.showinfo("Sucesso", "Arquivos exportados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", str(e))

# Função para listar arquivos salvos com base nos filtros
def export_listar_arquivos():
    conn = export_conectar_db()
    cursor = conn.cursor()
    
    # Filtrar por evento e task
    evento_selecionado = evento_combobox.get()
    task_selecionada = task_combobox.get()

    query = 'SELECT id, nome FROM anexos WHERE 1=1'
    params = []

    if evento_selecionado:
        query += ' AND evento = ?'
        params.append(evento_selecionado)

    if task_selecionada:
        query += ' AND task = ?'
        params.append(task_selecionada)

    cursor.execute(query, params)
    arquivos = cursor.fetchall()
    conn.close()
    
    # Limpar a lista existente
    listbox.delete(0, tk.END)
    
    # Adicionar arquivos à listbox
    if arquivos:  # Verifica se há arquivos
        for arquivo in arquivos:
            listbox.insert(tk.END, f"ID: {arquivo[0]}, Nome: {arquivo[1]}")
    else:
        messagebox.showinfo("Informação", "Nenhum arquivo encontrado.")

# Função para preencher o combobox de eventos
def export_preencher_eventos():
    conn = export_conectar_db()
    cursor = conn.cursor()
    cursor.execute('SELECT DISTINCT evento FROM anexos')
    eventos = cursor.fetchall()
    conn.close()

    # Adicionar eventos ao combobox
    evento_combobox['values'] = [evento[0] for evento in eventos]

# Função para preencher o combobox de tasks
def export_preencher_tasks():
    conn = export_conectar_db()
    cursor = conn.cursor()
    cursor.execute('SELECT DISTINCT task FROM anexos')
    tasks = cursor.fetchall()
    conn.close()

    # Adicionar tasks ao combobox
    task_combobox['values'] = [task[0] for task in tasks]

# Criar a interface gráfica
def export_criar_interface():

    janela_principal.iconify()
    # export_criar_tabela()
    
    root_export = tk.Toplevel()
    largura = 600
    altura = 700
    largura_tela = root_export.winfo_screenwidth()
    altura_tela = root_export.winfo_screenheight()
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2 - 30
    root_export.geometry(f"{largura}x{altura}+{x}+{y}")


    try:
        root_export.iconbitmap('Logo/logo.ico')
    except:
        pass
    root_export.title("Gerenciador de Anexos")
    # Tentar carregar e redimensionar a imagem para exibição
    try:
        imagem = Image.open('Logo/logo.ico')  # Tente abrir o arquivo de imagem
        imagem = imagem.resize((200, 200), Image.LANCZOS)  # Redimensione conforme necessário
        imagem_tk = ImageTk.PhotoImage(imagem)  # Converta para PhotoImage

        # Criando um label para exibir a imagem
        label_imagem = tk.Label(root_export, image=imagem_tk)
        label_imagem.pack(pady=(10, 5))  # Adicionando um espaço ao redor da imagem

    except FileNotFoundError:
        print("Arquivo logo.ico não encontrado. A imagem não será exibida.")
    

    # Seção de filtros
    tk.Label(root_export, text="Evento:").pack(pady=(0,5))
    global evento_combobox
    evento_combobox = ttk.Combobox(root_export)
    evento_combobox.pack(pady=5)
    export_preencher_eventos()  # Preencher combobox com eventos existentes

    tk.Label(root_export, text="Task:").pack(pady=5)
    global task_combobox
    task_combobox = ttk.Combobox(root_export)
    task_combobox.pack(pady=5)
    export_preencher_tasks()  # Preencher combobox com tasks existentes

    btn_listar = tk.Button(root_export, text="Listar Arquivos Filtrados", command=export_listar_arquivos)
    btn_listar.pack(pady=10)

    btn_exportar_selecionados = tk.Button(root_export, text="Exportar Arquivos Selecionados", command=export_exportar_arquivos_selecionados)
    btn_exportar_selecionados.pack(pady=10)

    global listbox
    listbox = tk.Listbox(root_export, width=50, selectmode=tk.MULTIPLE)  # Permitir seleção múltipla
    listbox.pack(pady=10)

    root_export.mainloop()

# Executar a interface

def eventos_carregar_banco():
    banco = "dados.db"
    conexao = sqlite3.connect(banco)
    cursor = conexao.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS eventos (
            evento TEXT,
            area TEXT,
            status TEXT,
            data_conclusao TEXT,
            descricao TEXT,
            stakeholder TEXT,
            impacto TEXT
        )
    ''')
    conexao.commit()
    conexao.close()
    return banco

def eventos_registrar_dados():
    evento = entry_evento.get()
    area = entry_area.get()
    status = combobox_status.get()
    data_conclusao = entry_data_conclusao.get()
    descricao = entry_descricao.get("1.0", tk.END).strip()
    stakeholder = entry_stakeholder.get()
    impacto = combobox_impacto.get()

    if not evento or not area or not status or not data_conclusao or not descricao or not stakeholder or not impacto:
        messagebox.showerror("Erro", "Todos os campos devem ser preenchidos!")
        return

    banco = eventos_carregar_banco()
    conexao = sqlite3.connect(banco)
    cursor = conexao.cursor()

    cursor.execute("""
        INSERT INTO eventos (evento, area, status, data_conclusao, descricao, stakeholder, impacto)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (evento, area, status, data_conclusao, descricao, stakeholder, impacto))
    conexao.commit()
    conexao.close()

    messagebox.showinfo("Sucesso", "Dados registrados com sucesso!")

def eventos_abrir_janela_registro():
    janela_principal.iconify()
    janela = tk.Toplevel()
    janela.title("Registrar evento")

    largura = 600
    altura = 700
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2 - 30
    janela.geometry(f"{largura}x{altura}+{x}+{y}")
    try:
        janela.iconbitmap('Logo/logo.ico')
    except:
            pass
        # Tentar carregar e redimensionar a imagem para exibição
    try:
            imagem = Image.open('Logo/logo.ico')  # Tente abrir o arquivo de imagem
            imagem = imagem.resize((85, 85), Image.LANCZOS)  # Redimensione conforme necessário
            imagem_tk = ImageTk.PhotoImage(imagem)  # Converta para PhotoImage


                    # Criando um label para exibir a imagem
            label_imagem = tk.Label(janela, image=imagem_tk)
            label_imagem.image = imagem_tk  # Mantendo a referência da imagem
            label_imagem.pack(pady=(10))  

            # Adicionando um espaço ao redor da imagem

    except FileNotFoundError:
        print("Arquivo logo.ico não encontrado. A imagem não será exibida.")


    frame_principal = ttk.Frame(janela)
    frame_principal.pack(fill="both", expand=True)

    # evento
    label_evento = ttk.Label(frame_principal, text="evento:")
    label_evento.pack(pady=(10, 0))
    global entry_evento
    entry_evento = ttk.Entry(frame_principal)
    entry_evento.pack(fill='x', padx=50, pady=(0, 10))

    # Área
    label_area = ttk.Label(frame_principal, text="Área:")
    label_area.pack(pady=(10, 0))
    global entry_area
    entry_area = ttk.Entry(frame_principal)
    entry_area.pack(fill='x', padx=50, pady=(0, 10))

    # status
    label_status = ttk.Label(frame_principal, text="status:")
    label_status.pack(pady=(10, 0))
    global combobox_status
    combobox_status = ttk.Combobox(frame_principal, values=["Pendente", "Em andamento", "Concluído"])
    combobox_status.pack(fill='x', padx=50, pady=(0, 10))

    # Data de Conclusão
    label_data_conclusao = ttk.Label(frame_principal, text="Data de Conclusão:")
    label_data_conclusao.pack(pady=(10, 0))
    global entry_data_conclusao
    entry_data_conclusao = ttk.Entry(frame_principal, bootstyle="primary")
    entry_data_conclusao.pack(fill='x', padx=50, pady=(0, 10))

    # Descrição
    label_descricao = ttk.Label(frame_principal, text="Descrição:")
    label_descricao.pack(pady=(10, 0))
    global entry_descricao
    entry_descricao = tk.Text(frame_principal, height=5, width=20)
    entry_descricao.pack(fill='x', padx=50, pady=(0, 10))

    # stakeholder
    label_stakeholder = ttk.Label(frame_principal, text="stakeholder:")
    label_stakeholder.pack(pady=(10, 0))
    global entry_stakeholder
    entry_stakeholder = ttk.Entry(frame_principal)
    entry_stakeholder.pack(fill='x', padx=50, pady=(0, 10))

    # impacto
    label_impacto = ttk.Label(frame_principal, text="impacto:")
    label_impacto.pack(pady=(10, 0))
    global combobox_impacto
    combobox_impacto = ttk.Combobox(frame_principal, values=['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'])
    combobox_impacto.pack(fill='x', padx=50, pady=(0, 10))

    botao_registrar = ttk.Button(frame_principal, text="Registrar", command=eventos_registrar_dados)
    botao_registrar.pack(fill='x', padx=50, pady=(20, 5))


# Variável global para armazenar anexos
anexos = []

def task_carregar_banco():
    banco = "dados.db"
    conexao = sqlite3.connect(banco)
    cursor = conexao.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Task (
            evento TEXT,
            task TEXT,
            categoria TEXT,
            descricao TEXT,
            inicio TEXT,
            fim TEXT,
            tempo TEXT,
            observacoes TEXT
        )
    ''')
    cursor.execute('''
         CREATE TABLE IF NOT EXISTS anexos (
            evento,
            task,
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            tipo TEXT,
            arquivo BLOB,
            extensao TEXT
        )
    ''')
    conexao.commit()
    conexao.close()
    return banco

def task_atualizar_eventos():
    banco = task_carregar_banco()
    conexao = sqlite3.connect(banco)
    cursor = conexao.cursor()
    
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='eventos'")
    existe = cursor.fetchone() is not None
    
    if not existe:
        messagebox.showinfo("Informação", "Por favor, registre um evento antes de continuar.")
        conexao.close()
        return

    cursor.execute("SELECT DISTINCT evento FROM eventos WHERE Evento IS NOT NULL")
    eventos = [row[0] for row in cursor.fetchall()]
    conexao.close()
    combobox_evento["values"] = eventos

def task_registrar_dados(evento, task, categoria, descricao, inicio, fim, tempo, observacoes):
    if not evento or not task or not categoria or not descricao or not inicio or not fim or not tempo:
        messagebox.showerror("Erro", "Todos os campos devem ser preenchidos!")
        return

    banco = task_carregar_banco()
    conexao = sqlite3.connect(banco)
    cursor = conexao.cursor()

    cursor.execute("""
        INSERT INTO Task (evento, task, categoria, descricao, inicio, fim, tempo, observacoes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (evento, task, categoria, descricao, inicio, fim, tempo, observacoes))
    conexao.commit()

    for anexo in anexos:
        blob_data = None
        tipo = anexo.split('.')[-1].lower()

        if tipo in ['jpg', 'jpeg', 'png']:
            with Image.open(anexo) as img:
                img = img.convert('RGB')
                img_byte_array = io.BytesIO()
                img.save(img_byte_array, format='JPEG')
                blob_data = img_byte_array.getvalue()
        elif tipo == 'txt':
            pdf_byte_array = io.BytesIO()
            c = canvas.Canvas(pdf_byte_array, pagesize=letter)
            with open(anexo, 'r') as file:
                text = file.read()
                c.drawString(100, 750, text)
            c.save()
            blob_data = pdf_byte_array.getvalue()
        else:
            with open(anexo, 'rb') as file:
                blob_data = file.read()

        extensao = os.path.splitext(anexo)[1]
        cursor.execute('INSERT INTO anexos (evento, task, nome, tipo, arquivo, extensao) VALUES (?, ?, ?, ?, ?, ?)', 
                       (evento, task, os.path.basename(anexo), tipo, blob_data, extensao))
    
    conexao.commit()
    conexao.close()
    messagebox.showinfo("Sucesso", "Dados registrados com sucesso!")
    task_atualizar_eventos()

def task_anexar_arquivo():
    global anexos
    escolha = messagebox.askquestion("Escolher Anexos", "Você deseja selecionar arquivos individuais (Sim) ou uma pasta (Não)?")
    
    if escolha == 'yes':
        arquivos = filedialog.askopenfilenames(
            title="Selecionar Anexos",
            filetypes=[("Todos os Arquivos", "*.*"), ("PDF", "*.pdf"), ("Imagens", "*.png;*.jpg;*.jpeg"), ("Documentos", "*.docx")]
        )
        if arquivos:
            anexos.extend(arquivos)  # Adiciona os arquivos à lista de anexos
    else:
        pasta = filedialog.askdirectory(title="Selecionar Pasta")
        if pasta:
            for arquivo in os.listdir(pasta):
                caminho_arquivo = os.path.join(pasta, arquivo)
                if os.path.isfile(caminho_arquivo):
                    anexos.append(caminho_arquivo)  # Adiciona o arquivo à lista de anexos

    # Preencher o campo entry_anexo com os anexos selecionados
    entry_anexo.delete(0, tk.END)  # Limpa o campo antes de adicionar
    entry_anexo.insert(0, ', '.join(anexos))  # Adiciona os anexos como uma string separada por vírgulas

def task_abrir_janela_registro():
    janela_principal.iconify()
    janela = tk.Toplevel()
    janela.title("Registrar Tarefa")
    try:
        janela.iconbitmap('Logo/logo.ico')
    except:
        pass
    largura = 600
    altura = 700
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2 - 30
    janela.geometry(f"{largura}x{altura}+{x}+{y}")
    try:
        janela.iconbitmap('Logo/logo.ico')
    except:
            pass
        # Tentar carregar e redimensionar a imagem para exibição
    # try:
    #         imagem = Image.open('Logo/logo.ico')  # Tente abrir o arquivo de imagem
    #         imagem = imagem.resize((85, 85), Image.LANCZOS)  # Redimensione conforme necessário
    #         imagem_tk = ImageTk.PhotoImage(imagem)  # Converta para PhotoImage


    #                 # Criando um label para exibir a imagem
    #         label_imagem = tk.Label(janela, image=imagem_tk)
    #         label_imagem.image = imagem_tk  # Mantendo a referência da imagem
    #         label_imagem.pack(pady=(10,5))  

    #         # Adicionando um espaço ao redor da imagem

    # except FileNotFoundError:
    #     print("Arquivo logo.ico não encontrado. A imagem não será exibida.")



    frame_principal = ttk.Frame(janela)
    frame_principal.pack(fill="both", expand=True)

    label_evento = ttk.Label(frame_principal, text="Evento:")
    label_evento.pack(pady=(10, 0))
    global combobox_evento
    combobox_evento = ttk.Combobox(frame_principal)
    combobox_evento.pack(fill='x', padx=50, pady=(0, 10))

    label_titulo = ttk.Label(frame_principal, text="Título da Task:")
    label_titulo.pack(pady=(10, 0))
    global entry_titulo
    entry_titulo = ttk.Entry(frame_principal)
    entry_titulo.pack(fill='x', padx=50, pady=(0, 10))

    frame_descricao_categoria = ttk.Frame(frame_principal)
    frame_descricao_categoria.pack(fill='x', padx=50, pady=(10, 0))

    frame_categoria = ttk.Frame(frame_descricao_categoria)
    frame_categoria.pack(side='left', fill='both', padx=(0,5), expand=True)

    label_categoria = ttk.Label(frame_categoria, text="Categoria da Tarefa:")
    label_categoria.pack(anchor='nw', pady=(5, 0))

    global combobox_categoria
    combobox_categoria = ttk.Combobox(frame_categoria, values=["Planejamento", "Logística", "Comunicação", "Suporte no dia do evento", "Pós-evento", "Administrativo"])
    combobox_categoria.pack(fill='x', expand=False)

    frame_descricao = ttk.Frame(frame_descricao_categoria)
    frame_descricao.pack(side='left', fill='x', padx=(5, 0), expand=True)

    label_descricao = ttk.Label(frame_descricao, text="Descrição:")
    label_descricao.pack(pady=(5, 0))
    global entry_descricao
    entry_descricao = tk.Text(frame_descricao, height=5, width=20)
    entry_descricao.pack(fill='x', expand=True)

    frame_datas = ttk.Frame(frame_principal)
    frame_datas.pack(fill='x', padx=50, pady=(10, 0))

    frame_inicio = ttk.Frame(frame_datas)
    frame_inicio.pack(side='left', fill='x', expand=True)

    label_inicio = ttk.Label(frame_inicio, text="Início:")
    label_inicio.pack(pady=(5, 0))
    global entry_inicio
    entry_inicio = ttk.Entry(frame_inicio)
    entry_inicio.pack(fill='x', padx=(0,5), expand=True)

    frame_fim = ttk.Frame(frame_datas)
    frame_fim.pack(side='left', fill='x', expand=True)

    label_fim = ttk.Label(frame_fim, text="Fim:")
    label_fim.pack(pady=(5, 0))
    global entry_fim
    entry_fim = ttk.Entry(frame_fim)
    entry_fim.pack(fill='x', padx=(5,0), expand=True)

    frame_temporizador = ttk.Frame(frame_principal)
    frame_temporizador.pack(fill='x', padx=50, pady=(10, 0))

    label_temporizador = ttk.Label(frame_temporizador, text="Tempo dedicado")
    label_temporizador.pack(fill='x', pady=(5, 0), anchor='nw')

    label_meses = ttk.Label(frame_temporizador, text="Meses:")
    label_meses.pack(side='left', padx=(0, 5))
    global spinbox_meses
    spinbox_meses = ttk.Spinbox(frame_temporizador, from_=0, to=11, width=5)
    spinbox_meses.pack(side='left', padx=(0, 5))

    label_dias = ttk.Label(frame_temporizador, text="Dias:")
    label_dias.pack(side='left', padx=(0, 5))
    global spinbox_dias
    spinbox_dias = ttk.Spinbox(frame_temporizador, from_=0, to=30, width=5)
    spinbox_dias.pack(side='left', padx=(0, 5))

    label_horas = ttk.Label(frame_temporizador, text="Horas:")
    label_horas.pack(side='left', padx=(0, 5))
    global spinbox_horas
    spinbox_horas = ttk.Spinbox(frame_temporizador, from_=0, to=23, width=5)
    spinbox_horas.pack(side='left', padx=(0, 5))

    label_observacoes = ttk.Label(frame_principal, text="Observações:")
    label_observacoes.pack(pady=(10, 0))
    global entry_observacoes
    entry_observacoes = tk.Text(frame_principal, height=5)
    entry_observacoes.pack(fill='x', padx=50, pady=(0, 10))

    label_anexo = ttk.Label(frame_principal, text="Anexo:")
    label_anexo.pack(pady=(10, 0))
    global entry_anexo
    entry_anexo = ttk.Entry(frame_principal)
    entry_anexo.pack(fill='x', padx=50, pady=(0, 10))

    botao_anexo = ttk.Button(frame_principal, text="Selecionar Anexo", command=task_anexar_arquivo)
    botao_anexo.pack(pady=(5, 10))

    botao_registrar = ttk.Button(frame_principal, text="Registrar", command=lambda: task_registrar_dados(
        combobox_evento.get(),
        entry_titulo.get(),
        combobox_categoria.get(),
        entry_descricao.get("1.0", tk.END).strip(),
        entry_inicio.get(),
        entry_fim.get(),
        f"{spinbox_meses.get().zfill(2)}/{spinbox_dias.get().zfill(2)}{spinbox_horas.get().zfill(2)}",
        entry_observacoes.get("1.0", tk.END).strip(),
    ))
    botao_registrar.pack(fill='x', padx=50, pady=(20, 5))

    task_atualizar_eventos()



janela_principal = tk.Tk()
janela_principal.title("Sistema de tarefas")

# Definindo o ícone da janela
try:
    janela_principal.iconbitmap('Logo/logo.ico')  # Tente abrir o arquivo de ícone
except Exception as e:
    print(f"Erro ao carregar o ícone: {e}")

largura = 600
altura = 700
largura_tela = janela_principal.winfo_screenwidth()
altura_tela = janela_principal.winfo_screenheight()
x = (largura_tela - largura) // 2
y = (altura_tela - altura) // 2 - 30
janela_principal.geometry(f"{largura}x{altura}+{x}+{y}")

# Tentar carregar e redimensionar a imagem para exibição
try:
    imagem = Image.open('Logo/logo.ico')  # Tente abrir o arquivo de imagem
    imagem = imagem.resize((200, 200), Image.LANCZOS)  # Redimensione conforme necessário
    imagem_tk = ImageTk.PhotoImage(imagem)  # Converta para PhotoImage

    # Criando um label para exibir a imagem
    label_imagem = tk.Label(janela_principal, image=imagem_tk)
    label_imagem.pack(pady=(10, 5))  # Adicionando um espaço ao redor da imagem

except FileNotFoundError:
    print("Arquivo logo.ico não encontrado. A imagem não será exibida.")


# Criando um Frame para centralizar os botões
frame_botao = ttk.Frame(janela_principal)
frame_botao.pack(expand=True)

# Centralizando os botões e diminuindo o espaçamento
botao_registrar_evento = ttk.Button(frame_botao, text="Registrar evento", command=eventos_abrir_janela_registro, width=30)
botao_registrar_evento.pack(pady=(10, 5), fill='x')

botao_registrar_tarefa = ttk.Button(frame_botao, text="Registrar tarefa", command=task_abrir_janela_registro, width=30)
botao_registrar_tarefa.pack(pady=(10, 5), fill='x')

botao_exportar_anexo = ttk.Button(frame_botao, text="Exportar anexos", command=export_criar_interface, width=30)
botao_exportar_anexo.pack(pady=(10, 5), fill='x')

# Para abrir a tela de edição a partir da janela principal
botao_editar_tabela = ttk.Button(frame_botao, text="Editar Tabela", command=abrir_tela_edicao)
botao_editar_tabela.pack(pady=(10, 5), fill="x")

janela_principal.mainloop()