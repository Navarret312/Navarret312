import pandas as pd
import os


# Carregar os dados do Excel
arquivo_excel = "Teste\Banco de Dados - BoardGames.xlsx"
aba = "Tabela13"

# Ler a tabela do excel
df = pd.read_excel(arquivo_excel, sheet_name=aba)
# Colunas selecionadas para o calculo
df_jogos = df[['Jogos', 'Fabricante', 'Mínimo de Jogadores', 'Máximo de Jogadores', 
                'Tempo de Jogo', 'Base de Jogo', 'Jogabilidade', 'Estilo', 
                'Tema', 'Dificuldade', 'Classif. Indicativa', 
                'Habilidades Cognitivas', 'Preço']]

# Função de calculo de similaridade
def calcular_similaridade(jogo1, jogo2):
    pesos = [3, 1, 1, 1, 1, 1, 2, 3, 1, 2, 4, 3]
    similaridade = 0

    # Calculo das similaridades
    similaridade += pesos[0] * (1 if jogo1["Fabricante"] == jogo2["Fabricante"] else 0)
    similaridade += pesos[1] * (1 - abs(jogo1["Mínimo de Jogadores"] - jogo2["Mínimo de Jogadores"]) / max(jogo1["Mínimo de Jogadores"], jogo2["Mínimo de Jogadores"]))
    similaridade += pesos[2] * (1 - abs(jogo1["Máximo de Jogadores"] - jogo2["Máximo de Jogadores"]) / max(jogo1["Máximo de Jogadores"], jogo2["Máximo de Jogadores"]))
    similaridade += pesos[3] * (1 - abs(jogo1["Tempo de Jogo"] - jogo2["Tempo de Jogo"]) / max(jogo1["Tempo de Jogo"], jogo2["Tempo de Jogo"]))
    similaridade += pesos[4] * (1 if jogo1["Base de Jogo"] == jogo2["Base de Jogo"] else 0)
    similaridade += pesos[5] * (1 if jogo1["Jogabilidade"] == jogo2["Jogabilidade"] else 0)
    similaridade += pesos[6] * (1 if jogo1["Estilo"] == jogo2["Estilo"] else 0)
    similaridade += pesos[7] * (1 if jogo1["Tema"] == jogo2["Tema"] else 0)
    similaridade += pesos[8] * (1 - abs(jogo1["Dificuldade"] - jogo2["Dificuldade"]) / max(jogo1["Dificuldade"], jogo2["Dificuldade"]))
    similaridade += pesos[9] * (1 if jogo1["Classif. Indicativa"] == jogo2["Classif. Indicativa"] else 0)
    similaridade += pesos[10] * (1 if jogo1["Habilidades Cognitivas"] == jogo2["Habilidades Cognitivas"] else 0)
    similaridade += pesos[11] * (1 if jogo1["Preço"] == jogo2["Preço"] else 0)

    # Normalizar a pontuação para garantir que a similaridade total não ultrapasse 1
    total_pesos = sum(pesos)
    similaridade = similaridade / total_pesos  # Normalizar para um valor entre 0 e 1
    return similaridade * 100  # Para porcentagem

# Criar tabela de similaridade
resultados = []
for i in range(len(df_jogos)):
    for j in range(len(df_jogos)):
        if i < j:  # Para evitar comparações duplicadas e comparação consigo mesmo
            sim = calcular_similaridade(df_jogos.iloc[i], df_jogos.iloc[j])
            resultados.append([df_jogos.iloc[i]["Jogos"], df_jogos.iloc[j]["Jogos"], sim])

# criando dataframe com os resultados
tabela_similaridade = pd.DataFrame(resultados, columns=["Jogo 1", "Jogo 2", "Porcentagem de Similaridade"])

# Exibir a tabela
print(tabela_similaridade)

# Salvar o dataframe como um excel
tabela_similaridade.to_excel("tabela_similaridade.xlsx", index=False)
