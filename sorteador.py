import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.constants import *
import pandas as pd
import random
from tkinter import ttk
import json
import os
import pygame

CONFIG_FILE = "config_estilo.json"
LOGO_FILE = "logo_selecionada.xlsx"  # Nome do arquivo Excel para salvar a logo

def carregar_imagens_do_diretorio(diretorio):
    """Carrega os arquivos PNG do diretório especificado."""
    return [f for f in os.listdir(diretorio) if f.endswith('.png')]

def atualizar_logo():
    """Atualiza o logo com a imagem selecionada no combobox e salva a escolha."""
    nome_logo = combobox_imagens.get()
    if nome_logo:
        caminho_logo = os.path.join("Ittruck/Sorteador", nome_logo)  # Ajuste para o diretório correto
        logo_original = tk.PhotoImage(file=caminho_logo)
        logo = logo_original.subsample(logo_original.width() // 180, logo_original.height() // 120)  # Redimensiona para 20x20

        # Atualiza o label do logo
        logo_label.config(image=logo)
        logo_label.image = logo  # Mantém uma referência da imagem

        # Salvar a escolha da logo em um arquivo Excel
        df = pd.DataFrame({'Logo': [nome_logo]})
        df.to_excel(LOGO_FILE, index=False)

def carregar_logo_salva():
    """Carrega a logo salva do arquivo Excel ao iniciar o aplicativo."""
    if os.path.exists(LOGO_FILE):
        df = pd.read_excel(LOGO_FILE)
        if not df.empty:
            nome_logo = df['Logo'].iloc[0]
            if nome_logo in imagens:
                combobox_imagens.set(nome_logo)
                atualizar_logo()

def aplicar_estilo():
    estilo_selecionado = combobox_estilos.get()
    style.theme_use(estilo_selecionado)
    with open(CONFIG_FILE, "w") as config_file:
        json.dump({"estilo": estilo_selecionado}, config_file)

def carregar_estilo_salvo():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as config_file:
            config = json.load(config_file)
            estilo_salvo = config.get("estilo")
            if estilo_salvo in style.theme_names():
                style.theme_use(estilo_salvo)
                combobox_estilos.set(estilo_salvo)

def selecionar_arquivo_excel():
    global df
    arquivo_excel = filedialog.askopenfilename(
        title="Selecione um arquivo Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if arquivo_excel:
        df = pd.read_excel(arquivo_excel)
        janela.iconify()  # Minimizar a janela principal
        abrir_janela_colunas()
    else:
        messagebox.showwarning("Arquivo não selecionado", "Nenhum arquivo foi selecionado.")

def abrir_janela_colunas():
    janela.iconify()
    janela_colunas = tk.Toplevel(janela)
    janela_colunas.title("Selecionar Coluna para Sortear")
    sincronizar_tela_cheia(janela_colunas)

    botao_voltar = tk.Button(janela_colunas, text="Voltar", command=lambda: voltar(janela_colunas))
    botao_voltar.pack(side='top', anchor='nw', padx=10, pady=10)

     # Adicionando o logo na janela de colunas
    logo_label_colunas = tk.Label(janela_colunas, image=logo_label.image)
    logo_label_colunas.pack(pady=(10, 20))

    frame_coluna = tk.Frame(janela_colunas)
    frame_coluna.pack(pady=(20, 10), padx=50, fill='x')

    label_coluna = tk.ttk.Label(frame_coluna, text="Sortear com base na coluna:")
    label_coluna.pack(side='left', padx=(0, 10))

    combobox_colunas = tk.ttk.Combobox(frame_coluna, values=df.columns.tolist())
    combobox_colunas.pack(side='left', expand=True, fill='x')

    botao_sortear = tk.ttk.Button(janela_colunas, text="Sortear Valor", command=lambda: sortear_valor(combobox_colunas.get()))
    botao_sortear.pack(fill='x', padx=50, pady=(10, 20))

    tree = ttk.Treeview(janela_colunas, columns=df.columns.tolist(), show='headings')
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, anchor='center')

    for index, row in df.iterrows():
        tree.insert("", "end", values=list(row))

    tree.pack(expand=True, fill='both', padx=50, pady=(0, 20))

   

def sortear_valor(coluna_selecionada):
    if coluna_selecionada and coluna_selecionada in df.columns:
        # Escolher uma linha aleatória
        linha_sorteada = df.dropna(subset=[coluna_selecionada]).sample(n=1)
        
        # Obter o valor sorteado da coluna selecionada
        valor_sorteado = linha_sorteada[coluna_selecionada].values[0]
        
        # Obter o valor correspondente da coluna "Participante"
        valor_participante = linha_sorteada['Participante'].values[0] if 'Participante' in df.columns else "Participante não encontrado"
        
 # Passar ambos os valores para exibir a notificação
        exibir_notificacao_animada(valor_sorteado, valor_participante)

def exibir_notificacao_animada(valor, participante):
    try:
        # Inicializar o mixer do pygame para tocar o som
        pygame.mixer.init()
        
        # Carregar o arquivo de som MP3 (certifique-se de que o caminho está correto)
        pygame.mixer.music.load("Ittruck\Sorteador\som.mp3")
        
        # Tocar o som
        # pygame.mixer.music.play()
    except pygame.error as e:
        print(f"Erro ao carregar ou tocar o som: {e}")
    
    janela.iconify()
    janela_notificacao = tk.Toplevel(janela)
    janela_notificacao.title("IT truck")
    sincronizar_tela_cheia(janela_notificacao)

    botao_voltar = tk.Button(janela_notificacao, text="Voltar", command=lambda: voltar(janela_notificacao))
    botao_voltar.pack(side='top', anchor='nw', padx=10, pady=10)

    logo_label_notificacao = tk.Label(janela_notificacao, image=logo_label.image, bg="#006181")
    logo_label_notificacao.place(relx=0.5, y=100, anchor='center')

    frame_central = tk.Frame(janela_notificacao, bg="#006181")
    frame_central.place(relx=0.5, rely=0.5, anchor='center')

    # label_titulo = tk.Label(frame_central, text="IT truck", font=("Helvetica", 40, "bold"), fg="#006181", bg="#006181")
    # label_titulo.pack(pady=(20, 10))

    label_parabens = tk.Label(frame_central, text="Parabéns!!!", font=("Helvetica", 45), fg="#006181", bg="#006181")
    label_parabens.pack(pady=(10, 5))

    label_resultado = tk.Label(frame_central, text=valor, font=("Helvetica", 48, "bold"), fg="red", bg="#006181")
    label_resultado.pack(pady=(5, 20))

    # Adicionar o valor do participante
    label_participante = tk.Label(frame_central, text=participante, font=("Helvetica", 36, "bold"), fg="white", bg="#006181")
    label_participante.pack(pady=(5, 20))

    def flash_effect():
        current_color = label_parabens.cget('fg')
        new_color = '#006181' if current_color == 'white' else 'white'
        label_parabens.configure(fg=new_color)
        janela_notificacao.after(500, flash_effect)

    flash_effect()

def adicionar_checkbox_tela_cheia(janela):
    def toggle_fullscreen():
        fullscreen = var_tela_cheia.get()
        janela.attributes('-fullscreen', fullscreen)
        if fullscreen:
            janela.geometry("")
        atualizar_estado_tela_cheia(fullscreen)

    var_tela_cheia = tk.BooleanVar(value=estado_tela_cheia)
    checkbox_tela_cheia = tk.Checkbutton(janela, text="Tela cheia", variable=var_tela_cheia, command=toggle_fullscreen)
    checkbox_tela_cheia.pack(side='left', anchor='sw', padx=10, pady=10)

def sincronizar_tela_cheia(janela):
    adicionar_checkbox_tela_cheia(janela)
    janela.attributes('-fullscreen', estado_tela_cheia)

def atualizar_estado_tela_cheia(fullscreen):
    global estado_tela_cheia
    estado_tela_cheia = fullscreen

def voltar(janela_atual):
    janela_atual.destroy()
    janela.deiconify()  # Restaurar a janela principal

estado_tela_cheia = False

janela = tk.Tk()
janela.title("Sorteio IT TRUCK")
sincronizar_tela_cheia(janela)

# Carregando o logo e redimensionando para 20x20
logo_original = tk.PhotoImage("ittruck_negativo.png")  # Altere para o caminho correto do seu logo
logo = logo_original.subsample(logo_original.width() // 180, logo_original.height() // 120)  # Redimensiona para 20x20

# Adicionando o logo
logo_label = tk.Label(janela, image=logo)
logo_label.pack(pady=(10, 20))  # Adicionando espaço acima e abaixo do logo

style = Style(theme='superhero')

frameTema = tk.Frame(janela)
frameTema.pack(pady=20, padx=40, side=tk.BOTTOM)

estilos_disponiveis = style.theme_names()
combobox_estilos = tk.ttk.Combobox(frameTema, values=estilos_disponiveis)
combobox_estilos.set(style.theme_use())
combobox_estilos.pack(expand=True, fill='x', side='left')

carregar_estilo_salvo()

botao_aplicar = tk.ttk.Button(frameTema, text="Aplicar tema", command=aplicar_estilo)
botao_aplicar.pack(pady=(10), side='left')

# Carregar imagens do diretório
imagens = carregar_imagens_do_diretorio("Ittruck/Sorteador")  # Altere para o diretório correto

# Combobox para selecionar o logo
combobox_imagens = tk.ttk.Combobox(frameTema, values=imagens)
combobox_imagens.set("Selecionar Logo")
combobox_imagens.pack(pady=(10), side='left')

# Botão para atualizar o logo
botao_atualizar_logo = tk.ttk.Button(frameTema, text="Atualizar Logo", command=atualizar_logo)
botao_atualizar_logo.pack(pady=(10), side='left')

# Adicionando título acima do botão "Selecionar arquivo Excel"
titulo_label = ttk.Label(janela, text="Sorteio IT TRUCK", font=("Arial", 16))
titulo_label.pack(pady=(50, 5))  # Adicionando espaço acima e abaixo do título

botao_selecionar_excel = tk.ttk.Button(janela, text="Selecionar arquivo Excel", command=selecionar_arquivo_excel)
botao_selecionar_excel.pack(fill='x', padx=50, pady=(100, 5))

# Carregar a logo salva ao iniciar
carregar_logo_salva()

janela.mainloop()