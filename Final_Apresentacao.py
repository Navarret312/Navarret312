import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import networkx as nx
import pandas as pd
import plotly.graph_objs as go
import community as community_louvain

tabela_similaridade_jogadores = pd.read_excel("tabela_similaridade_jogadores.xlsx")
tabela_similaridade_jogos = pd.read_excel("tabela_similaridade.xlsx")

def criar_grafo_jogadores(limiar, jogador_selecionado=None):
    G = nx.Graph()
    if jogador_selecionado:
        similaridades = tabela_similaridade_jogadores[tabela_similaridade_jogadores["Jogador"] == jogador_selecionado]
    else:
        similaridades = tabela_similaridade_jogadores

    for index, row in similaridades.iterrows():
        if row['Porcentagem de Similaridade'] >= limiar:
            G.add_edge(row['Jogador'], row['Jogo'], weight=row['Porcentagem de Similaridade'])

    # Se o grafo não tiver arestas, retorne um grafo vazio
    if G.number_of_edges() == 0:
        return nx.Graph()  # Retorna um grafo vazio

    return G

def criar_grafo_jogos(limiar, jogo_selecionado=None):
    G = nx.Graph()
    if jogo_selecionado:
        similaridades = tabela_similaridade_jogos[(tabela_similaridade_jogos["Jogo 1"] == jogo_selecionado) | (tabela_similaridade_jogos["Jogo 2"] == jogo_selecionado)]
    else:
        similaridades = tabela_similaridade_jogos

    for index, row in similaridades.iterrows():
        if row['Porcentagem de Similaridade'] >= limiar:
            G.add_edge(row['Jogo 1'], row['Jogo 2'], weight=row['Porcentagem de Similaridade'])

    # Se o grafo não tiver arestas, retorne um grafo vazio
    if G.number_of_edges() == 0:
        return nx.Graph()  # Retorna um grafo vazio

    return G

def gerar_figura_grafo(grafo):
    # Calcular a posição dos nós
    pos = nx.spring_layout(grafo, seed=42)
    
    # Detectar comunidades
    partition = community_louvain.best_partition(grafo)
    num_comunidades = max(partition.values()) + 1
    
    # Atribuir cores às comunidades
    cores = [partition[node] for node in grafo.nodes()]
    
    edge_trace = [
        go.Scatter(
            x=[pos[edge[0]][0], pos[edge[1]][0], None],
            y=[pos[edge[0]][1], pos[edge[1]][1], None],
            line=dict(width=0.5, color='#888'),
            hoverinfo='none',
            mode='lines'
        ) for edge in grafo.edges(data=True)
    ]
    
    node_trace = go.Scatter(
        x=[pos[node][0] for node in grafo.nodes()],
        y=[pos[node][1] for node in grafo.nodes()],
        text=list(grafo.nodes()),
        mode='markers+text',
        hoverinfo='text',
        marker=dict(
            showscale=True,
            colorscale='Viridis',  # Escolha um esquema de cores
            color=cores,  # A cor depende da comunidade
            size=10,
            colorbar=dict(
                thickness=15,
                title='Comunidades',
                xanchor='left',
                titleside='right'
            ),
        )
    )
    
    fig = go.Figure(
        data=edge_trace + [node_trace],
        layout=go.Layout(
            showlegend=False,
            hovermode='closest',
            margin=dict(b=20, l=5, r=5, t=40),
            xaxis=dict(showgrid=False, zeroline=False),
            yaxis=dict(showgrid=False, zeroline=False)
        )
    )
    return fig

app = dash.Dash(__name__)

app.layout = html.Div([
    html.Div([
        html.H1("Visualizador de Grafo de", className='title'),
        html.H1("Jogos e Jogadores", className='title2'),

        html.Div([
            html.Div([
                html.H3("Tipo de Grafo", className='card-title'),
                html.P("Escolha o tipo de grafo que deseja visualizar.", className='card-text'),
                dcc.Dropdown(
                    id='tipo-grafo',
                    options=[
                        {'label': 'Jogadores', 'value': 'Jogadores'},
                        {'label': 'Jogos', 'value': 'Jogos'}
                    ],
                    value='Jogadores',
                    className='dropdown'
                )
            ], className='card'),

            html.Div([
                html.H3("Limiar de Similaridade", className='card-title'),
                html.P("Ajuste o limiar para exibir similaridades acima do valor especificado.", className='card-text'),
                dcc.Input(
                    id='limiar',
                    type='number',
                    value=80,
                    min=0,
                    max=100,
                    step=1,
                    className='input'
                )
            ], className='card'),

            html.Div([
                html.H3("Seleção", className='card-title'),
                html.P("Selecione um elemento para destacar no grafo.", className='card-text'),
                dcc.Dropdown(
                    id='selecao',
                    options=[],
                    value=None,
                    className='dropdown'
                )
            ], className='card'),
        ], className='cards'),

        html.Div([
            html.Div([
                html.Div([
                    html.H3("Visualização do Grafo", className='card-title'),
                    html.P("Aqui está a representação visual do grafo com base nos parâmetros escolhidos.", className='card-text'),
                    dcc.Graph(id='grafo', className='graph')
                ], className='card-grafo'),
                
                html.Div([
                    html.H4("Parâmetros do Grafo", className='param-title'),
                    html.Ul([
                        html.Li(html.Strong("Centralidade:"), id='centralidade', className='param-item'),
                        html.Li(html.Strong("Modularidade:"), id='modularidade', className='param-item'),
                        html.Li(html.Strong("Grau Médio:"), id='grau-medio', className='param-item'),
                        html.Li(html.Strong("Diâmetro:"), id='diametro', className='param-item'),
                        html.Li(html.Strong("Densidade:"), id='densidade', className='param-item'),
                        html.Li(html.Strong("Comprimento Médio do Caminho:"), id='comprimento-medio-caminho', className='param-item'),
                        html.Li(html.Strong("Closeness:"), id='closeness', className='param-item'),
                        html.Li(html.Strong("Betweenness:"), id='betweenness', className='param-item'),
                    ], className='param-list')
                ], className='card-info')
            ], className='grafo-container')
        ], className='cards'),

        html.Div([
            html.Div([
                html.H3("Formulário de Informações", className='card-title'),
                html.P("Preencha os dados necessários para análise.", className='card-text'),
                html.Div([
                    dcc.Input(id='jogo1', type='text', placeholder='Jogo 1', className='input'),
                    dcc.Input(id='jogo2', type='text', placeholder='Jogo 2', className='input'),
                    dcc.Input(id='sim-jogos', type='number', placeholder='Porcentagem de Similaridade', className='input'),
                    dcc.Input(id='jogador', type='text', placeholder='Jogador', className='input'),
                    dcc.Input(id='jogo', type='text', placeholder='Jogo', className='input'),
                    dcc.Input(id='sim-jogador', type='number', placeholder='Porcentagem de Similaridade', className='input'),
                ], className='form-group'),
                html.Div([
                    html.Button("Enviar", id='botao-enviar', className='botao')
                ], className='botao-container')
            ], className='card'),
        ], className='cards')

    ], className='container')
], className='main', style={
    'backgroundColor': '#1e1e1e',
    'color': '#ffffff',
    'fontFamily': 'Arial, sans-serif',
    'backgroundImage': 'url("https://www.transparenttextures.com/patterns/dark-mosaic.png")',
    'padding': '20px'
})

app.index_string = '''
<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Visualizador de Grafo</title>
        <link href="https://fonts.googleapis.com/css2?family=Jaro:wght@400;700&family=Rubik:wght@400;500;700&display=swap" rel="stylesheet">
        <link rel="stylesheet" href="https://codepen.io/chriddyp/pen/bWLwgP.css">
        <style>
            body {
                background-color: #F7F7F7;
                color: #ffffff;
                font-family: 'Rubik', sans-serif;
                background-image: url('https://www.transparenttextures.com/patterns/dark-mosaic.png');
                margin: 0;
                padding: 0;
            }
            .main {
                padding: 20px;
            }
            .container {
                background-color: #2e2e2e;
                border-radius: 50px;
                padding: 50px;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
                margin: 10px auto;
                max-width: 1200px;
            }
            .title {
                text-align: center;
                color: #8e44ad;
                font-size: 48px;
                font-weight: 700;
                font-family: 'Jaro', sans-serif;
            }
            .title2 {
                text-align: center;
                color: white;
                font-size: 48px;
                font-weight: 700;
                font-family: 'Jaro', sans-serif;
            }
            .label {
                font-weight: bold;
                color: #8e44ad;
            }
            .form-group {
                display: flex;
                flex-direction: column;
                gap: 15px;
                margin-top: 20px;
            }
            .botao-container {
                display: flex;
                justify-content: center;
                margin-top: 20px;
            }
            .botao {
                font-size: 16px;
                background-color: #8e44ad;
                color: #ffffff;
                border: none;
                border-radius: 5px;
                cursor: pointer;
                text-align: center;
            }
            .botao:hover {
                background-color: #732d91;
            }
            .botao:focus {
                outline: none;
                box-shadow: 0 0 5px #ffcc00;
            }
            .cards {
                display: flex;
                justify-content: center;
                gap: 20px;
                margin-top: 20px;
            }
            .card {
                background-color: #383838;
                border-radius: 8px;
                padding: 20px;
                width: 300px;
            }
            .card-grafo {
                background-color: #383838;
                border-radius: 8px;
                padding: 20px;
                width: 80%;
            }
            .card-info {
                margin-top: 0px;
                background-color: #383838;
                border-radius: 8px;
                padding: 20px;
                width: 30%;
            }
            .grafo-container {
                display: flex;
                justify-content: space-between;
                gap: 20px;
                width: 100%;
            }
            .input {
                padding: 10px;
                border-radius: 8px;
                border: none;
                margin-bottom: 10px;
                width: 100%;
                background-color: #f7f7f7;
                color: #333333;
                font-size: 14px;
            }
            .dropdown {
                padding: 10px;
                border-radius: 8px;
                border: none;
                width: 100%;
                background-color: #f7f7f7;
                color: #333333;
                font-size: 14px;
            }
            .graph {
                width: 100%;
                height: 500px;
            }
            .param-title {
                font-size: 20px;
                font-weight: 700;
                color: #ffffff;
                margin-bottom: 15px;
                text-align: center;
            }
            .param-list {
                list-style-type: none;
                padding: 0;
                margin: 0;
                color: #ffffff;
                font-size: 14px;
            }
            .param-item {
                margin-bottom: 15px;
            }
            .param-item strong {
                color: #8e44ad;
                font-weight: 700;
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        {%config%}
        {%scripts%}
        {%renderer%}
    </body>
</html>
'''

@app.callback(
    Output('grafo', 'figure'),
    Output('selecao', 'options'),
    Output('centralidade', 'children'),
    Output('modularidade', 'children'),
    Output('grau-medio', 'children'),
    Output('diametro', 'children'),
    Output('densidade', 'children'),
    Output('comprimento-medio-caminho', 'children'),
    Output('closeness', 'children'),
    Output('betweenness', 'children'),
    Input('tipo-grafo', 'value'),
    Input('limiar', 'value'),
    Input('selecao', 'value')
)
def update_graph(tipo_grafo, limiar, selecao):
    try:
        limiar = float(limiar)  # Tenta converter para float
    except (TypeError, ValueError):
        limiar = 80  # Define um valor padrão se não for possível converter

    # Criar o grafo com base no tipo e seleção
    if tipo_grafo == 'Jogadores':
        grafo = criar_grafo_jogadores(limiar, selecao)
        options = [{'label': jogador, 'value': jogador} for jogador in tabela_similaridade_jogadores['Jogador'].unique()]
    else:
        grafo = criar_grafo_jogos(limiar, selecao)
        options = [{'label': jogo, 'value': jogo} for jogo in tabela_similaridade_jogos['Jogo 1'].unique()]

    # Verificar se o grafo é vazio (sem arestas)
    if grafo.number_of_edges() == 0:
        return go.Figure(), options, "Centralidade: N/A", "Modularidade: N/A", "Grau Médio: N/A", "Diâmetro: N/A", "Densidade: N/A", "Comprimento Médio do Caminho: N/A", "Closeness: N/A", "Betweenness: N/A"

    # Cálculo da centralidade ponderada (degree centrality)
    centralidade = nx.degree_centrality(grafo)
    centralidade_max_node = max(centralidade, key=centralidade.get, default="N/A")

    # Modularidade (somente se o grafo não estiver vazio)
    try:
        partition = community_louvain.best_partition(grafo)
        modularidade = community_louvain.modularity(partition, grafo)
    except ValueError:  # Se ocorrer um erro, definir modularidade como N/A
        modularidade = "N/A"

    # Grau médio ponderado
    grau_medio = sum(dict(grafo.degree(weight='weight')).values()) / len(grafo.nodes())

    # Diâmetro ponderado
    diametro = nx.diameter(grafo) if nx.is_connected(grafo) else float('inf')

    # Densidade ponderada
    densidade = nx.density(grafo)

    # Comprimento médio do caminho ponderado
    comprimento_medio_caminho = nx.average_shortest_path_length(grafo, weight='weight') if nx.is_connected(grafo) else float('inf')

    # Closeness ponderado
    closeness = nx.closeness_centrality(grafo, distance='weight')
    if closeness:
        closeness_max_node = max(closeness, key=closeness.get, default="N/A")
        closeness_value = closeness[closeness_max_node]
    else:
        closeness_max_node = "N/A"
        closeness_value = 0

    # Betweenness ponderado
    betweenness = nx.betweenness_centrality(grafo, weight='weight')
    if betweenness:
        betweenness_max_node = max(betweenness, key=betweenness.get, default="N/A")
        betweenness_value = betweenness[betweenness_max_node]
    else:
        betweenness_max_node = "N/A"
        betweenness_value = 0

    # Retornando os resultados para o callback
    return (
        gerar_figura_grafo(grafo),
        options,
        f"Centralidade de grau: {centralidade_max_node}",
        f"Centralidade Betweenness: {betweenness_max_node} (Valor: {betweenness_value:.2f})",
        f"Centralidade Closeness: {closeness_max_node} (Valor: {closeness_value:.2f})",
        f"Modularidade: {modularidade}",
        f"Grau Médio: {grau_medio:.2f}",
        f"Diâmetro: {diametro}",
        f"Densidade: {densidade:.2f}",
        f"Comprimento Médio do Caminho: {comprimento_medio_caminho:.2f}"
    )

if __name__ == '__main__':
    app.run_server(debug=True)