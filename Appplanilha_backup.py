import customtkinter as ctk
import pandas as pd
from tkinter import filedialog

# Variáveis globais para paginação
pagina_atual = 0
itens_por_pagina = 20  # Número de itens a serem exibidos por página

# Função para carregar o arquivo Excel e retornar as abas disponíveis
def carregar_planilha():
    try:
        # Abre a janela para o usuário escolher o arquivo
        arquivo = filedialog.askopenfilename(title="Selecionar Planilha", filetypes=(("Arquivos Excel", "*.xlsx;*.xls"), ("Todos os arquivos", "*.*")))

        # Se o usuário não selecionar nenhum arquivo, retorna
        if not arquivo:
            return pd.DataFrame(), []

        # Lê o arquivo Excel e obtém as abas (sheets) disponíveis
        excel_file = pd.ExcelFile(arquivo)
        abas = excel_file.sheet_names  # Lista de abas (planilhas) no arquivo
        print(f"Planilha carregada: {arquivo}")
        return excel_file, abas
    except Exception as e:
        print("Erro ao carregar planilha:", e)
        return pd.DataFrame(), []

# Função para carregar os dados de uma aba específica
def carregar_dados_aba(aba):
    try:
        # Carrega os dados da aba selecionada
        dados = excel_file.parse(aba)
        dados.columns = dados.columns.str.strip().str.upper()  # Padroniza os nomes das colunas
        return dados
    except Exception as e:
        print(f"Erro ao carregar dados da aba {aba}: {e}")
        return pd.DataFrame()

# Função para filtrar os dados com base na pesquisa
def filtrar_dados():
    categoria = filtro_categoria.get()
    termo_pesquisa = entrada_pesquisa.get().strip()  # Limpa espaços antes e depois

    if termo_pesquisa:  # Se um termo de pesquisa foi inserido
        if categoria:  # Filtra pela categoria (coluna) selecionada
            if categoria in df.columns:
                # Acessa a coluna de dados
                coluna = df[categoria]

                # Garantir que a coluna seja tratada como string
                coluna = coluna.astype(str)

                # Aplicar a busca usando str.contains()
                resultado = df[coluna.str.contains(termo_pesquisa, case=False, na=False)]
                atualizar_lista(resultado)
            else:
                print(f"Categoria '{categoria}' não encontrada.")
        else:  # Se nenhuma categoria for selecionada, busca por todas as colunas
            # Filtra por todas as colunas, convertendo para string antes de aplicar o filtro
            resultado = df[df.apply(lambda row: row.astype(str).str.contains(termo_pesquisa, case=False, na=False).any(), axis=1)]
            atualizar_lista(resultado)
    else:
        print("Insira um termo de pesquisa")
        atualizar_lista(df)  # Se não houver filtro, exibe todos os dados

# Função para atualizar a lista de resultados
def atualizar_lista(dados_filtrados):
    global pagina_atual

    # Limpa os widgets antigos
    for widget in frame_lista.winfo_children():
        widget.destroy()

    # Calcular os índices de início e fim para a página atual
    inicio = pagina_atual * itens_por_pagina
    fim = inicio + itens_por_pagina
    dados_pag = dados_filtrados.iloc[inicio:fim]

    if not dados_pag.empty:
        # Exibe as colunas na interface gráfica
        for i, column in enumerate(dados_pag.columns):
            ctk.CTkLabel(frame_lista, text=column, width=130, anchor="w", font=("Arial",10)).grid(row=0, column=i, padx=10, pady=5)
        
        # Exibe as linhas de dados filtrados
        for i, row in dados_pag.iterrows():
            for j, column in enumerate(dados_pag.columns):
                ctk.CTkLabel(frame_lista, text=row[column], width=130, anchor="w", font=("Arial",10)).grid(row=i+1, column=j, padx=10, pady=5)
    else:
        ctk.CTkLabel(frame_lista, text="Nenhum resultado encontrado", width=130, anchor="w", font=("Arial",10)).grid(row=0, column=0, padx=10, pady=5)

    # Atualiza os botões de navegação de páginas
    atualizar_botoes_paginacao(dados_filtrados)

# Função para atualizar os botões de navegação de páginas
def atualizar_botoes_paginacao(dados_filtrados):
    global pagina_atual

    total_itens = len(dados_filtrados)
    total_paginas = (total_itens // itens_por_pagina) + (1 if total_itens % itens_por_pagina > 0 else 0)

    # Desabilitar o botão de "Página Anterior" se estamos na primeira página
    if pagina_atual == 0:
        botao_anterior.configure(state=ctk.DISABLED)
    else:
        botao_anterior.configure(state=ctk.NORMAL)

    # Desabilitar o botão de "Próxima Página" se estamos na última página
    if pagina_atual == total_paginas - 1:
        botao_proxima.configure(state=ctk.DISABLED)
    else:
        botao_proxima.configure(state=ctk.NORMAL)

# Função para exibir a página anterior
def pagina_anterior():
    global pagina_atual
    if pagina_atual > 0:
        pagina_atual -= 1
        atualizar_lista(df)  # Atualiza a lista com a página anterior

# Função para exibir a próxima página
def pagina_proxima():
    global pagina_atual
    pagina_atual += 1
    atualizar_lista(df)  # Atualiza a lista com a próxima página

# Função para atualizar as categorias de filtro disponíveis
def atualizar_categorias():
    if df.empty:
        return
    categorias = df.columns.tolist()
    filtro_categoria.configure(values=categorias)  # Atualiza o menu de seleção de categorias

# Função para abrir o arquivo e carregar as abas
def abrir_planilha():
    global excel_file, df, abas
    excel_file, abas = carregar_planilha()
    if abas:
        filtro_aba.configure(values=abas)  # Atualiza o menu de seleção de abas
        filtro_aba.set(abas[0])  # Seleciona a primeira aba por padrão

# Função para carregar e exibir os dados da aba selecionada
def exibir_aba_selecionada():
    global df
    aba_selecionada = filtro_aba.get()
    df = carregar_dados_aba(aba_selecionada)
    if not df.empty:
        atualizar_categorias()  # Atualiza as categorias no menu
        atualizar_lista(df)  # Exibe os dados na interface

# Função para redefinir os filtros e mostrar todos os dados novamente
def redefinir_filtro():
    entrada_pesquisa.delete(0, ctk.END)  # Limpa o campo de pesquisa
    filtro_categoria.set('')  # Limpa a seleção da categoria
    atualizar_lista(df)  # Mostra todos os dados novamente

# Criando a interface gráfica
app = ctk.CTk()

# Variáveis globais
excel_file = None
df = pd.DataFrame()
abas = []

# Configuração do layout para expandir widgets
app.grid_columnconfigure(0, weight=1)
app.grid_columnconfigure(1, weight=3)
app.grid_columnconfigure(2, weight=1)

app.grid_rowconfigure(0, weight=1)
app.grid_rowconfigure(1, weight=1)
app.grid_rowconfigure(2, weight=1)
app.grid_rowconfigure(3, weight=4)
app.grid_rowconfigure(4, weight=1)

# Campo de entrada de pesquisa
entrada_pesquisa = ctk.CTkEntry(app, placeholder_text="Pesquisar...")
entrada_pesquisa.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

# Dropdown para filtrar categorias
filtro_categoria = ctk.CTkOptionMenu(app)
filtro_categoria.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

# Dropdown para selecionar a aba
filtro_aba = ctk.CTkOptionMenu(app)
filtro_aba.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

# Botão de filtro
botao_filtrar = ctk.CTkButton(app, text="Filtrar", command=filtrar_dados)
botao_filtrar.grid(row=0, column=2, padx=10, pady=10, sticky="ew")

# Botão para abrir a planilha
botao_abrir = ctk.CTkButton(app, text="Abrir Planilha", command=abrir_planilha)
botao_abrir.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

# Botão para redefinir os filtros
botao_redefinir = ctk.CTkButton(app, text="Redefinir Filtro", command=redefinir_filtro)
botao_redefinir.grid(row=1, column=2, padx=10, pady=10, sticky="ew")

# Botão para exibir os dados da aba selecionada
botao_exibir_aba = ctk.CTkButton(app, text="Exibir Aba Selecionada", command=exibir_aba_selecionada)
botao_exibir_aba.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

# Frame para exibir os dados
frame_lista = ctk.CTkFrame(app)
frame_lista.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# Botões de navegação de página
botao_anterior = ctk.CTkButton(app, text="Página Anterior", command=pagina_anterior)
botao_anterior.grid(row=4, column=0, padx=10, pady=10, sticky="ew")

botao_proxima = ctk.CTkButton(app, text="Próxima Página", command=pagina_proxima)
botao_proxima.grid(row=4, column=2, padx=10, pady=10, sticky="ew")

# Iniciar a aplicação
app.mainloop()
