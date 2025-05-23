import os
import re
import zipfile
import cv2
import pandas as pd
import pytesseract
from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import ttkbootstrap as ttk
from ttkbootstrap.style import Style
from ttkbootstrap.constants import *
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook


def tratamento_arquivos(zip_file_path, cfop_file_path, monitor_file_path):
    # Diretório onde os arquivos Excel serão extraídos
    extract_dir = 'extracted_files'
    
    # Lê o arquivo CFOP1.xlsx e a coluna CFOP NORMAL
    cfop_df = pd.read_excel(cfop_file_path)

    # Cria o diretório se não existir
    os.makedirs(extract_dir, exist_ok=True)

    # Extrai o arquivo ZIP
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    # Lista para armazenar os DataFrames e a contagem total de linhas
    dataframes = []
    total_lines_count = 0

    # Lê todos os arquivos Excel extraídos
    for file in os.listdir(extract_dir):
        if file.endswith('.xlsx'):  # Verifica se o arquivo é um Excel
            file_path = os.path.join(extract_dir, file)
            
            # Extrai o número da nota do nome do arquivo
            nota_number = file.split('-')[0]  # Pega a parte antes do primeiro '-'
            
            df = pd.read_excel(file_path)
            
            # Remove caracteres não numéricos das colunas relevantes
            df['Chave de acesso'] = df['Chave de acesso'].astype(str).str.replace(r'\D', '', regex=True)
            df['Emitente CNPJ/CPF'] = df['Emitente CNPJ/CPF'].astype(str).str.replace(r'\D', '', regex=True)
            df['Emitente I.E'] = df['Emitente I.E'].astype(str).str.replace(r'\D', '', regex=True)
            df['Destintario I.E'] = df['Destintario I.E'].astype(str).str.replace(r'\D', '', regex=True)
            df['Destinatario CNPJ/CPF'] = df['Destinatario CNPJ/CPF'].astype(str).str.replace(r'\D', '', regex=True)
            
            # Adiciona a coluna com o número da nota como a primeira coluna
            df.insert(0, 'Numero_NFE', nota_number)  # Insere a coluna na posição 0
            dataframes.append(df)
            total_lines_count += len(df)  # Soma o número de linhas do DataFrame

    # Concatena todos os DataFrames em um único
    combined_df = pd.concat(dataframes, ignore_index=True)

    # Mescla o DataFrame combinado com o CFOP1 com base na coluna CFOP
    merged_df = pd.merge(combined_df, cfop_df[['CFOP NORMAL']], left_on='[CFOP]', right_on='CFOP NORMAL', how='inner')

    # Carregar o arquivo Excel do Monitor
    df_monitor = pd.read_excel(monitor_file_path)

    # Definir as colunas que devem ser concatenadas na ordem correta e suas respectivas quantidades de dígitos
    columns_to_concatenate = {
        'Region of Issuer': 2,
        'Year of Document Date': 2,
        'Month of Document Date': 2,
        'CNPJ/CPF Number of Issuer': 14,
        'Nota Fiscal model': 2,
        'Series': 3,
        'Nine-Digit Document Number': 9,
        'Random Number in Access Key': 9,
        'Check Digit in Access Key': 1
    }

    # Função para remover caracteres não numéricos e adicionar zeros à esquerda
    def format_value(value, target_length):
        numeric_value = ''.join(filter(str.isdigit, str(value)))
        if len(numeric_value) < target_length:
            return numeric_value.zfill(target_length)
        return numeric_value[:target_length]

    # Concatenar as colunas na ordem especificada e aplicar a formatação
    def concatenate_values(row):
        formatted_values = []
        for col, length in columns_to_concatenate.items():
            if col == 'Nine-Digit Document Number':
                nine_digit_value = format_value(row[col], length)
                if nine_digit_value.endswith('0'):
                    nine_digit_value = '0' + nine_digit_value[:-1]
                formatted_values.append(nine_digit_value)
            else:
                formatted_values.append(format_value(row[col], length))
        return ''.join(formatted_values)

    df_monitor['chave_de_acesso'] = df_monitor.apply(concatenate_values, axis=1)
    df_monitor['chave_de_acesso_valida'] = df_monitor['chave_de_acesso'].apply(lambda x: len(x) == 44)

    # Remove espaços em branco dos nomes das colunas
    df_monitor.columns = df_monitor.columns.str.strip()
    merged_df.columns = merged_df.columns.str.strip()

    # Realiza o merge com base na coluna 'chave_de_acesso' do df_monitor e 'Chave de acesso' do merged_df
    final_merged_df = pd.merge(df_monitor, merged_df, left_on='chave_de_acesso', right_on='Chave de acesso', how='inner')

    # Abre a caixa de diálogo para salvar o arquivo Excel
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                      filetypes=[("Excel files", "*.xlsx"),
                                                                 ("All files", "*.*")])
    if output_file_path:  # Verifica se o usuário não cancelou a operação
        final_merged_df.to_excel(output_file_path, index=False)
        messagebox.showinfo("Sucesso", f"Arquivo salvo como '{output_file_path}' com as correspondências encontradas.")

    

def selecionar_arquivo_tratamento(entry):
    file_path = filedialog.askopenfilename()
    if file_path:
        entry.delete(0, tk.END)  # Limpa o campo de entrada
        entry.insert(0, file_path)  # Insere o caminho do arquivo

def abrir_tela_tratamento():
    janela.iconify()
    top = tk.Toplevel()
    top.title("Seleção de Arquivos")

    # Centraliza os componentes na tela
    for i in range(4):  # 4 linhas
        top.grid_rowconfigure(i, weight=1)
    
    tk.Label(top, text="Caminho para o arquivo ZIP:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
    zip_entry = tk.Entry(top, width=100)
    zip_entry.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(top, text="Selecionar", command=lambda: selecionar_arquivo_tratamento(zip_entry)).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(top, text="Caminho para o arquivo CFOP de referência:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
    cfop_entry = tk.Entry(top, width=100)
    cfop_entry.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(top, text="Selecionar", command=lambda: selecionar_arquivo_tratamento(cfop_entry)).grid(row=1, column=2, padx=5, pady=5)

    tk.Label(top, text="Caminho para o arquivo MONITOR:").grid(row=2, column=0, sticky='e', padx=5, pady=5)
    monitor_entry = tk.Entry(top, width=100)
    monitor_entry.grid(row=2, column=1, padx=5, pady=5)
    tk.Button(top, text="Selecionar", command=lambda: selecionar_arquivo_tratamento(monitor_entry)).grid(row=2, column=2, padx=5, pady=5)

    tk.Button(top, text="Executar Tratamento", command=lambda: tratamento_arquivos(zip_entry.get(), cfop_entry.get(), monitor_entry.get())).grid(row=3, columnspan=3, pady=10)



def extract_data_from_image(image_path, progress_callback=None):
    # Definir o caminho para o Tesseract
    # Obter o caminho do diretório atual do script
    pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"

    # Verificar se o caminho do arquivo existe
    if not os.path.isfile(image_path):
        print(f"Arquivo não encontrado: {image_path}")
        return pd.DataFrame()  # Retorna um DataFrame vazio

    # Variáveis para armazenar os dados gerais da nota
    # Variáveis para armazenar os dados gerais da nota
    barcode = None
    number = None
    bayer = None
    endereco = None
    numero_endereco = None
    endereco_bayer_jurisdiction_code = None
    endereco_bayer_jurisdiction_name = None
    endereco_bayer_state = None
    CNPJ_BAYER = None
    reg_number = None
    random_number = None
    CFOP_natOp = None
    invoice_type = None
    serie_number = None
    nfe_number = None
    barcode_identification = None
    total_amount = None
    additional_information = None
    additional_info_for_material = None
    jurisdiction_code = None
    vendor = None
    vendor_state = None
    vendor_jurisdiction_name = None
    material_service = None
    qty = None
    unit = None
    price_per_unit = None
    amount_type = None
    CFOP_material = None
    vendor_cnpj = None
    
    # Lista para armazenar dados de materiais
    materiais = []

    bayer_found = False
    vendor_found = False

    # Abrir a imagem usando Pillow
    try:
        with Image.open(image_path) as img:
            # Se o arquivo TIFF tiver várias páginas
            for page in range(img.n_frames):
                img.seek(page)  # Mover para a página atual
                texto = pytesseract.image_to_string(img)
            
                # Exibir texto extraído da página
                print(f"\nTexto extraído da página {page + 1} da imagem {image_path}:\n{texto}")

                # Usar expressão regular para encontrar o Barcode
                # Usar expressão regular para encontrar o Barcode
                match_barcode = re.search(r'Barcode:\s*(\d+)', texto)
                if match_barcode:
                    barcode = match_barcode.group(1).strip()
                    print(f"Barcode encontrado: {barcode}")

            

                # Usar expressão regular para encontrar o Buyer ou Bayer
                match_bayer = re.search(r'(Buyer|Bayer)\s*([\w\s\.]+)', texto)
                if match_bayer and not bayer_found:
                    bayer = match_bayer.group(2).strip().split('\n')[0]
                    print(f"Bayer encontrado: {bayer}")
                    bayer_found = True  # Atualizar controle para indicar que o Buyer foi encontrado

                # Se o Buyer já foi encontrado, extrair os outros campos
                if bayer_found and not vendor_found:
                    # Usar expressão regular para encontrar o endereço
                    match_endereco = re.search(r'RUA\s*([\w\s\.]+)', texto)
                    if match_endereco:
                        endereco = match_endereco.group(1).strip().split('\n')[0]
                        print(f"Endereço encontrado: {endereco}")

                    # Usar expressão regular para encontrar o número do endereço
                    match_numero_endereco = re.search(r'Number:\s*(\d+)', texto)
                    if match_numero_endereco:
                        numero_endereco = match_numero_endereco.group(1).strip()
                        print(f"Número do endereço encontrado: {numero_endereco}")

                    # Usar expressões regulares para encontrar Jurisdiction Code, Jurisdiction Name e State
                    match_jurisdiction_code = re.search(r'Jurisdiction Code:\s*(\d+)', texto)
                    if match_jurisdiction_code:
                        endereco_bayer_jurisdiction_code = match_jurisdiction_code.group(1).strip()
                        print(f"Jurisdiction Code encontrado: {endereco_bayer_jurisdiction_code}")

                    match_jurisdiction_name = re.search(r'Jurisdiction Name:\s*([\w\s]+)', texto)
                    if match_jurisdiction_name:
                        endereco_bayer_jurisdiction_name = match_jurisdiction_name.group(1).strip().split('\n')[0]
                        print(f"Jurisdiction Name encontrado: {endereco_bayer_jurisdiction_name}")

                    match_state = re.search(r'State:\s*([A-Z]{2})', texto)
                    if match_state:
                        endereco_bayer_state = match_state.group(1).strip()
                        print(f"State encontrado: {endereco_bayer_state}")

                    # Usar expressões regulares para encontrar Reg Number e CNPJ
                    match_reg_number = re.search(r'Reg Number \(IE\):\s*(\d+)', texto)
                    if match_reg_number:
                        reg_number = match_reg_number.group(1).strip()
                        print(f"Reg Number encontrado: {reg_number}")

                    match_cnpj = re.search(r'(CNPJ|CNPUJ|CNPU):\s*([\d\s]+)', texto)
                    if match_cnpj:
                        CNPJ_BAYER = match_cnpj.group(2).strip()
                        print(f"CNPJ encontrado: {CNPJ_BAYER}")

                # Usar expressão regular para encontrar o Vendor
                match_vendor = re.search(r'Vendor\s*(.*)', texto)
                if match_vendor and not vendor_found:
                    vendor = match_vendor.group(1).strip().split('\n')[0]
                    print(f"Vendor encontrado: {vendor}")
                    vendor_found = True  # Atualizar controle para indicar que o Vendor foi encontrado

                # Se o Vendor já foi encontrado, extrair os outros campos
                if vendor_found:
                    # Usar expressão regular para encontrar o endereço do Vendor
                    match_vendor_endereco = re.search(r'Av\.?\s*([\w\s\.]+)', texto)
                    if match_vendor_endereco:
                        vendor_numero_endereco = match_vendor_endereco.group(1).strip().split('\n')[0]
                        print(f"Vendor Endereço encontrado: {vendor_numero_endereco}")

                    # Usar expressão regular para encontrar o número do endereço do Vendor
                    match_vendor_numero_endereco = re.search(r'Number:\s*(\d+)', texto)
                    if match_vendor_numero_endereco:
                        vendor_numero_endereco = match_vendor_numero_endereco.group(1).strip()
                        print(f"Vendor Número do endereço encontrado: {vendor_numero_endereco}")

                    # Usar expressões regulares para encontrar Jurisdiction Code, Jurisdiction Name e State do Vendor
                    match_vendor_jurisdiction_code = re.search(r'Jurisd\. Gode:\s*(\d+)', texto)
                    if match_vendor_jurisdiction_code:
                        jurisdiction_code = match_vendor_jurisdiction_code.group(1).strip()
                        print(f"Vendor Jurisdiction Code encontrado: {jurisdiction_code}")

                    match_vendor_jurisdiction_name = re.search(r'Jurisd\. Name:\s*([\w\s]+)', texto)
                    if match_vendor_jurisdiction_name:
                        vendor_jurisdiction_name = match_vendor_jurisdiction_name.group(1).strip().split('\n')[0]
                        print(f"Vendor Jurisdiction Name encontrado: {vendor_jurisdiction_name}")

                    match_vendor_state = re.search(r'State:\s*([A-Z]{2})', texto)
                    if match_vendor_state:
                        vendor_state = match_vendor_state.group(1).strip()
                        print(f"Vendor State encontrado: {vendor_state}")

                    # Usar expressões regulares para encontrar Reg Number e CNPJ do Vendor
                    match_vendor_reg_number = re.search(r'Reg Number \(IE\):\s*(\d+)', texto)
                    if match_vendor_reg_number:
                        match_vendor_reg_number = match_vendor_reg_number.group(1).strip()
                        print(f"Vendor Reg Number encontrado: {match_vendor_reg_number}")

                    match_vendor_inscricao_muni = re.search(r'Inscricao Muni\. \(IM\):\s*(\d+)', texto)
                    if match_vendor_inscricao_muni:
                        match_vendor_inscricao_muni = match_vendor_inscricao_muni.group(1).strip()
                        print(f"Vendor Inscricao Muni encontrado: {match_vendor_inscricao_muni}")

                    # Encontrar o CNPJ do Vendor após o texto "Vendor"
                    vendor_text = texto[texto.find("Vendor"):]
                    match_vendor_cnpj = re.search(r'(CNPJ|CNPUJ|CNPU):\s*([\d\s]+)', vendor_text)
                    if match_vendor_cnpj:
                        vendor_cnpj = match_vendor_cnpj.group(2).strip()
                        print(f"Vendor CNPJ encontrado: {vendor_cnpj}")

                # Usar expressões regulares para encontrar informações adicionais
                match_random_number = re.search(r'Random Number \(cNF\):\s*(\d+)', texto)
                if match_random_number:
                    random_number = match_random_number.group(1).strip()
                    print(f"Random Number encontrado: {random_number}")

                match_CFOP_natOp = re.search(r'CFOP \(natOp\):\s*([\w\s\/]+)', texto)
                if match_CFOP_natOp:
                    CFOP_natOp = match_CFOP_natOp.group(1).strip()
                    print(f"CFOP (natOp) encontrado: {CFOP_natOp}")

                match_invoice_type = re.search(r'Invoice type \(mod\):\s*(\d+)', texto)
                if match_invoice_type:
                    invoice_type = match_invoice_type.group(1).strip()
                    print(f"Invoice type encontrado: {invoice_type}")

                match_serie_number = re.search(r'Serie number \(serie\):\s*(\d+)', texto)
                if match_serie_number:
                    serie_number = match_serie_number.group(1).strip()
                    print(f"Série number encontrado: {serie_number}")

                match_nfe_number = re.search(r'NFe number \(nNF\):\s*(\d+)', texto)
                if match_nfe_number:
                    nfe_number= match_nfe_number.group(1).strip()
                    print(f"NFe number encontrado: {nfe_number}")

                match_barcode_identification = re.search(r'Barcode identification \(chNFe\):\s*([\d\s]+)', texto)
                if match_barcode_identification:
                    barcode_identification = match_barcode_identification.group(1).strip()
                    print(f"Barcode identification encontrado: {barcode_identification }")

                match_total_amount = re.search(r'Total amount:\s*([\d,\.]+)', texto)
                if match_total_amount:
                    total_amount = match_total_amount.group(1).strip()
                    print(f"Total amount encontrado: {total_amount}")
               
                # Usar expressões regulares para encontrar informações adicionais
                matches_additional_information = re.finditer(r'Additional Information:\s*(.*?)(?=\n\s*Additional Information:|\Z)', texto, re.DOTALL)
                # Pegar a terceira ocorrência de Additional Information
                additional_info_list = [match.group(1).strip() for match in matches_additional_information]
                if len(additional_info_list) >= 3:
                    additional_information = additional_info_list[2]
                    print(f"additional_information {additional_information}")

                # Usar expressão regular para encontrar materiais e CFOP
                matches_materiais = re.findall(r'(\d+)\s+([\d\.]+)\s+(\w+)\s+([\d\.]+)\s+([\d\.]+)', texto)
                for match_material in matches_materiais:
                    material_service, qty, unit, price_per_unit, amount_type = match_material
                    print(f"Material encontrado: {material_service}, Qty: {qty}, Unit: {unit}, PricePerUnit: {price_per_unit}, Amount/Type: {amount_type}")
                    # Converter material_service para inteiro
                    material_service = int(material_service)
                    # Encontrar CFOP específico para o material
                    match_CFOP_material = re.search(r'CFOP:\s*([\w\s\/]+)', texto)
                    CFOP_material = match_CFOP_material.group(1).strip() if match_CFOP_material else None
                    print(f"CFOP para material encontrado: {CFOP_material}")

                    # Encontrar a primeira informação adicional relacionada ao material
                    
                    if additional_info_list:
                        additional_info_for_material = additional_info_list[0]  # Captura a primeira informação adicional
                        print(f"Additional Information para material {material_service}: {additional_info_for_material}")

                    

                                        # Adicionar material ao DataFrame
                    materiais.append({
                        'imagem': image_path,
                        'NFe Number': nfe_number,
                        'Barcode': barcode,
                        'Bayer': bayer,
                        'Bayer Endereco': endereco,
                        # 'Número Endereço': numero_endereco,
                        'Bayer Jurisdiction Code': endereco_bayer_jurisdiction_code,
                        'Bayer Jurisdiction Name': endereco_bayer_jurisdiction_name,
                        'Bayer State': endereco_bayer_state,
                        'CNPJ Bayer': CNPJ_BAYER,
                        'Vendor': vendor,
                        'CNPJ Vendor': vendor_cnpj,
                        # 'Vendor Numero Endereco':vendor_numero_endereco,
                        'Vendor State': vendor_state,
                        'Vendor Jurisdiction Name': vendor_jurisdiction_name,
                        'Vendor Jurisdiction Code': jurisdiction_code,
                        'Reg Number': reg_number,
                        'Random Number': random_number,
                        'CFOP (natOp)': CFOP_natOp,
                        'Invoice Type': invoice_type,
                        'Série Number': serie_number,
                        'Barcode Identification': barcode_identification,
                        'Total Amount': total_amount,
                        'Informações Adicionais': additional_information,
                        'Material/Service': material_service,
                        'Qty': qty,
                        'Unit': unit,
                        'PricePerUnit': price_per_unit,
                        'Amount/Type': amount_type,
                        'CFOP': CFOP_material,
                        'additional information material': additional_info_for_material
                    })

    except Exception as e:
        print(f"Erro ao processar a imagem: {e}")

    # Criar DataFrame com os dados dos materiais
    df_materiais = pd.DataFrame(materiais)

    # Verificar se há valores diferentes de 0 na coluna 'Material/Service'
    if (df_materiais['Material/Service'] != 0).any():
        # Se houver valores diferentes de 0, remover as linhas com valor 0
        df_materiais = df_materiais[df_materiais['Material/Service'] != 0]
    else:
        # Se houver apenas valores 0, manter apenas uma linha com valor 0
        df_materiais = df_materiais[df_materiais['Material/Service'] == 0].head(1)
    # Remover caracteres não numéricos da coluna "CFOP"
    df_materiais['CFOP'] = df_materiais['CFOP'].str.replace(r'\D', '', regex=True)

    if progress_callback:
        progress_callback() 
    return df_materiais

def processar_pasta(pasta):
    all_dataframes = []
    total_files = len([f for f in os.listdir(pasta) if f.endswith(('.tif', '.tiff', '.png', '.jpg', '.jpeg'))])
    current_file = 0

    # Criar uma nova janela para exibir o progresso
    progress_window = tk.Toplevel()
    progress_window.title("Processando Imagens")
    progress_label = tk.Label(progress_window, text="Processando imagens...")
    progress_label.pack(padx=10, pady=10)
    progress_message = tk.Label(progress_window, text="")
    progress_message.pack(padx=10, pady=10)

    # Função de callback para atualizar o progresso
    def update_progress():
        nonlocal current_file
        current_file += 1
        progress_message.config(text=f"{current_file} de {total_files} imagens processadas")
        progress_window.update()  # Atualiza a janela para mostrar a mensagem

    # Percorrer todos os arquivos na pasta
    for filename in os.listdir(pasta):
        if filename.endswith(('.tif', '.tiff', '.png', '.jpg', '.jpeg')):  # Adicione as extensões desejadas
            caminho_completo = os.path.join(pasta, filename)
            print(f"Processando: {caminho_completo}")
            df = extract_data_from_image(caminho_completo, progress_callback=update_progress)
            all_dataframes.append(df)

    # Fecha a janela de progresso após o processamento
    progress_window.destroy()

    # Concatenar todos os DataFrames em um único DataFrame
    if all_dataframes:
        final_df = pd.concat(all_dataframes, ignore_index=True)
        return final_df
    else:
        return pd.DataFrame() 
    
def selecionar_pasta():
    pasta = filedialog.askdirectory()
    if pasta:
        print(f"Pasta selecionada: {pasta}")
        df_final = processar_pasta(pasta)

        # Exibir mensagem de conclusão
        if not df_final.empty:
            df_final.to_excel('dados_extraidos.xlsx', index=False)
            messagebox.showinfo("Sucesso", "Os dados foram extraídos e salvos em 'dados_extraidos.xlsx'.")
        else:
            messagebox.showwarning("Aviso", "Nenhum dado foi extraído.")




# Criar janela principal
janela = tk.Tk()
janela.title("Escolher Arquivo Python")

style = Style(theme='cosmo')

# Titulo da janela
janela.title('BayFlow')

# Definir tamanho
largura = 400
altura = 350
# Obter a largura e a altura da tela
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()
x = (largura_tela - largura) // 8
y = 0
janela.geometry(f"{largura}x{altura}+{x}+{y}")

# Adicionar botão para executar "Ler imagens"
botao_requisicao = tk.Button(janela, text="Ler imagens", command=selecionar_pasta)
botao_requisicao.pack(fill='x', padx=50, pady=(100, 5))   

botao_tratamento = tk.Button(janela, text="Operação triangular", command=abrir_tela_tratamento)
botao_tratamento.pack(fill='x', padx=50, pady=(20, 5))

janela.mainloop()



# frameTema = ttk.Frame(janela)
# frameTema.pack(pady=20, padx=40, side=BOTTOM)
# estilos_disponiveis = style.theme_names()
# combobox_estilos = ttk.Combobox(frameTema, values=estilos_disponiveis)
# combobox_estilos.set(style.theme_use())  # Definir o valor inicial para o estilo atual
# combobox_estilos.pack(expand=True, fill='x', side='left')
# botao_aplicar = ttk.Button(frameTema, text="Aplicar tema", command=aplicar_estilo)
# botao_aplicar.pack(pady=(10),side='left')
# try:
#     with open("config_estilo.json", "r") as config_file:
#         config = json.load(config_file)
#         estilo_salvo = config.get("estilo", None)

#         if estilo_salvo and estilo_salvo in estilos_disponiveis:
#             style.theme_use(estilo_salvo)
#             combobox_estilos.set(estilo_salvo)
# except FileNotFoundError:
#     pass  # O arquivo de configuração ainda não existe

# Rodar aplicação
janela.mainloop()
