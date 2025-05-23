import pandas as pd

# Carregar os dados do Excel
arquivo_jogos = r"C:\Users\GOWTX\OneDrive - Bayer\Desktop\Projetos desenvolvimento\Teste\Banco de Dados - BoardGames.xlsx"
arquivo_jogadores = r"C:\Users\GOWTX\OneDrive - Bayer\Desktop\Projetos desenvolvimento\tabela_jogadores.xlsx"
aba_jogos = "Tabela13"

# Ler a tabela de jogos
df_jogos = pd.read_excel(arquivo_jogos, sheet_name=aba_jogos)
df_jogos = df_jogos[['Jogos', 'Fabricante', 'Mínimo de Jogadores', 'Máximo de Jogadores', 
                     'Tempo de Jogo', 'Base de Jogo', 'Jogabilidade', 'Estilo', 
                     'Tema', 'Dificuldade', 'Classif. Indicativa', 
                     'Habilidades Cognitivas', 'Preço']]

# Ler a tabela de jogadores
df_jogadores = pd.read_excel(arquivo_jogadores)

# Função para calcular a similaridade
def calcular_similaridade(jogo, preferencias):
    similaridade = 0
    pesos = [3, 1, 1, 1, 1, 1, 2, 3, 1, 2, 4, 3]

    # Atributos que são categóricos e binários (Fabricante, Base de Jogo, etc.)
    similaridade += pesos[0] * (1 if jogo["Fabricante"] == preferencias["Fabricante"] else 0)
    similaridade += pesos[4] * (1 if jogo["Base de Jogo"] == preferencias["Base de Jogo"] else 0)
    similaridade += pesos[5] * (1 if jogo["Jogabilidade"] == preferencias["Jogabilidade"] else 0)
    similaridade += pesos[6] * (1 if jogo["Estilo"] == preferencias["Estilo"] else 0)
    similaridade += pesos[7] * (1 if jogo["Tema"] == preferencias["Tema"] else 0)
    similaridade += pesos[9] * (1 if jogo["Classif. Indicativa"] == preferencias["Classif. Indicativa"] else 0)
    similaridade += pesos[10] * (1 if jogo["Habilidades Cognitivas"] == preferencias["Habilidades Cognitivas"] else 0)

    # Atributos numéricos (Mínimo de Jogadores, Tempo de Jogo, Dificuldade)
    similaridade += pesos[1] * (1 - abs(jogo["Mínimo de Jogadores"] - preferencias["Mínimo de Jogadores"]) / max(jogo["Mínimo de Jogadores"], preferencias["Mínimo de Jogadores"]))
    similaridade += pesos[2] * (1 - abs(jogo["Máximo de Jogadores"] - preferencias["Máximo de Jogadores"]) / max(jogo["Máximo de Jogadores"], preferencias["Máximo de Jogadores"]))
    similaridade += pesos[3] * (1 - abs(jogo["Tempo de Jogo"] - preferencias["Tempo de Jogo"]) / max(jogo["Tempo de Jogo"], preferencias["Tempo de Jogo"]))
    similaridade += pesos[8] * (1 - abs(jogo["Dificuldade"] - preferencias["Dificuldade"]) / max(jogo["Dificuldade"], preferencias["Dificuldade"]))

    # Normalizar a pontuação para garantir que a similaridade total não ultrapasse 1
    total_pesos = sum(pesos)
    similaridade = similaridade / total_pesos  # Normalizar para um valor entre 0 e 1

    return similaridade * 100  # Converter para porcentagem

# Criar tabela de similaridade
resultados = []
for _, jogador in df_jogadores.iterrows():
    preferencias = jogador  # Característicasas do jogador
    for _, jogo in df_jogos.iterrows():
        sim = calcular_similaridade(jogo, preferencias)
        resultados.append([jogador["Nome"], jogo["Jogos"], sim])

# Criar DataFrame com os resultados
tabela_similaridade = pd.DataFrame(resultados, columns=["Jogador", "Jogo", "Porcentagem de Similaridade"])

# Exibir a tabela
print(tabela_similaridade)

# Salvar a tabela em um arquivo Excel
tabela_similaridade.to_excel("tabela_similaridade_jogadores.xlsx", index=False)
