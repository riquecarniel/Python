# Bibliotecas utilizadas
import geopandas as gpd
import pandas as pd
from shapely.geometry import Point, LineString, Polygon
import numpy as np
import fiona
import time
import customtkinter as ctk
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter import messagebox
from pathlib import Path
import webbrowser


# Tema para a janela criada
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Interface gráfica para iteração com o usuário
janela = ctk.CTk()
janela.title("Estudo de Visibilidade")
janela.geometry("325x340")
janela.rowconfigure(0, weight=1)
janela.columnconfigure([0, 1], weight=1)
janela.resizable(width=False, height=False)

# Função para criar entry
def criar_entry(janela, placeholder, x, y):
    valor = ctk.CTkEntry(
        janela,
        placeholder_text=placeholder,
        placeholder_text_color="white",
        text_color="white",
        border_width=3,
        corner_radius=10,
        width=305)
    valor.place(x=x, y=y)
    return valor

# Caixas para inserção de valores
valor_velocidade = criar_entry(janela, "Inf. a velocidade em km/h", 10, 120)
valor_proibicao = criar_entry(janela, "Inf. a distância mínima de proibição em metros", 10, 155)
valor_faixa = criar_entry(janela, "Inf. a largura da faixa em metros (Usar Ponto)", 10, 190)
valor_acost = criar_entry(janela, "Inf. a largura do Acost. em metros (Usar Ponto)", 10, 225)
valor_nome_arquivo = criar_entry(janela, "Inf. o nome do excel a ser gerado", 10, 260)

# Funções para selecionar os arquivos e salvar
def importar_alt():
    global caminho_alt
    caminho_alt = askopenfilename(title="Selecione o excel correspondente ao relatório de Altimetria!")

def importar_plan():
    global caminho_plan
    caminho_plan = askopenfilename(title="Selecione o excel correspondente ao relatório de Planimetria!")

def selecionar_pasta():
    global caminho_salvar
    caminho_salvar = askdirectory(title="Selecione uma pasta para salvar")

# Função para abrir url/site
def abrir_nuvem():
    # URL do site que você deseja abrir
    url_do_site = 'https://drive.google.com/drive/folders/11YA_bQfIZ8XH2y49A1WZpkIsiTD97JnV?usp=sharing'  # Substitua pelo URL real

    # Abrir o site usando a biblioteca webbrowser
    webbrowser.open(url_do_site)

caminho_salvar = None
caminho_plan = None
caminho_alt = None

# Função que vai verificar as planilhas e gerar o resultado
def verificar():
    try:
        velocidade = int(valor_velocidade.get())
        proibicao = int(valor_proibicao.get())
        faixa = float(valor_faixa.get())
        acost = float(valor_acost.get())
    except ValueError as e:
        Janela = ctk.CTk()
        messagebox.showerror("Erro", f"Erro: {e}. Certifique-se de inserir valores numéricos corretos e usar ponto quando nescessário.")
        Janela.destroy()
        return
    nome_arquivo = valor_nome_arquivo.get()
    largura_total = faixa + acost

    # INICIO - Verifica se o usuario selecionou uma pasta para salvar o arquivo
    global caminho_alt
    if caminho_alt is None or not caminho_alt:
        Janela = ctk.CTk()
        messagebox.showerror("Erro", "Selecione um arquivo de Altimetria.")
        Janela.destroy()
        return

    global caminho_plan
    if caminho_plan is None or not caminho_plan:
        Janela = ctk.CTk()
        messagebox.showerror("Erro", "Selecione um arquivo de Planimetria.")
        Janela.destroy()
        return  # Encerra a função se a pasta não estiver selecionada

    global caminho_salvar
    if caminho_salvar is None or not caminho_salvar:
        Janela = ctk.CTk()
        messagebox.showerror("Erro", "Selecione uma pasta para salvar Excel.")
        Janela.destroy()
        return
    # FIM - Verifica se o usuario selecionou uma pasta para salvar o arquivo

    Janela = ctk.CTk()
    messagebox.showinfo(
        "Estudo de Visibilidade de Ultrapassagem",
        "Análise iniciada, aguarde a conclusão!")
    Janela.destroy()

    # Inicio contagem do tempo rodando o programa
    ti = time.time()

    df_planimetria = pd.read_excel(caminho_plan) # Le a planilha
    df_planimetria = df_planimetria.rename(columns={df_planimetria.columns[0]: 'Estaca'}) # Modifica o nome da coluna 1 para Estaca na planilha planimetria

    # INICIO - ALGORITMO QUE REMOVE A LETRA M E CONVERTE EM FLOAT - Iterar sobre as linhas e colunas
    for index, row in df_planimetria.iterrows():
        for coluna in df_planimetria.columns:
            # Verificar se a célula contém a letra 'm'
            if 'm' in str(row[coluna]):
                # Remover a letra 'm', substituir vírgula por ponto e converter para float
                valor_sem_m = str(row[coluna]).replace('m', '').replace(',', '.')
                df_planimetria.at[index, coluna] = float(valor_sem_m)

    # Converter todas as colunas para float
    df_planimetria = df_planimetria.astype(float)
    # FIM - ALGORITMO QUE REMOVE A LETRA M E CONVERTE EM FLOAT

    df_altimetria = pd.read_excel(caminho_alt)
    df_altimetria = df_altimetria.rename(columns={df_altimetria.columns[0]: 'Estaca'}) # Modifica o nome da coluna 1 para Estaca na planilha altimetria

    # INICIO - ALGORITMO QUE REMOVE A LETRA M E CONVERTE EM FLOAT - Iterar sobre as linhas e colunas
    for index, row in df_altimetria.iterrows():
        for coluna in df_altimetria.columns:
            # Verificar se a célula contém a letra 'm'
            if 'm' in str(row[coluna]):
                # Remover a letra 'm', substituir vírgula por ponto e converter para float
                valor_sem_m = str(row[coluna]).replace('m', '').replace(',', '.')
                df_altimetria.at[index, coluna] = float(valor_sem_m)

    # Converter todas as colunas para float
    df_altimetria = df_altimetria.astype(float)
    # FIM - ALGORITMO QUE REMOVE A LETRA M E CONVERTE EM FLOAT

    ponto_inicial = df_altimetria["Estaca"].min() # Station - Significa Estaca no excel
    ponto_final = df_altimetria["Estaca"].max()
    #pontos_analise = list(range(ponto_inicial, ponto_final + 20, 20)) # Obrigatorio as estacas estar em 20 em 20 metros no excel / codigo antigo
    pontos_analise = list(range(int(ponto_inicial), int(ponto_final) + 20, 20)) # Converte os pontos em int
    df_analise = pd.DataFrame(pontos_analise, columns=["Estaca"])

    df_planimetria["Ponto"] = [Point(round(row[3], 4), round(row[2], 4)) for row in df_planimetria.itertuples()]
    df_altimetria["Ponto"] = [Point(row[1], row[2]) for row in df_altimetria.itertuples()]

    eixo = LineString(df_planimetria["Ponto"])
    eixo_esq = eixo.parallel_offset(largura_total, "left")
    eixo_dir = eixo.parallel_offset(largura_total, "right")

    linha_alt = LineString(df_altimetria["Ponto"])

    gdf_planimetria = gpd.GeoDataFrame(df_planimetria, geometry="Ponto")
    gdf_altimetria = gpd.GeoDataFrame(df_altimetria, geometry="Ponto")

    # Cria um dicionario com as velocidades regulamentadas e distancia minima de visibilidade
    dict_velocidades = {
        40: 140,
        50: 160,
        60: 180,
        70: 210,
        80: 245,
        90: 280,
        100: 320,
        110: 355,
    }

    df_velocidades = pd.DataFrame.from_dict(dict_velocidades, orient="index", columns=["Distância"]).reset_index(names=["Velocidade"])

    def verificar_alt(pto, sentido):
        if pto in df_altimetria["Estaca"].values:
            ponto = df_altimetria[df_altimetria["Estaca"] == pto]["Ponto"].values[0]

            raio = df_velocidades[df_velocidades["Velocidade"] == velocidade][
                "Distância"
            ].values[0]
            circle = ponto.buffer(raio)

            intersection = linha_alt.intersection(circle)
            x, y = intersection.xy

            if sentido == "C":
                ponto_sentido = Point(x[-1], y[-1])
                diagonal = LineString((ponto, ponto_sentido))
            else:
                ponto_sentido = Point(x[0], y[0])
                diagonal = LineString((ponto_sentido, ponto))

            diagonal_offset = diagonal.parallel_offset(1.2, "left")

            if diagonal_offset.intersection(intersection):
                # poly = linha_verificada.intersection(linha_regua)
                typeline = 1
            else:
                typeline = 2

            return typeline
        else:
            return None

    cache = {}

    def verificar_plan(pto, sentido):
        try:
            if pto in df_planimetria["Estaca"].values:
                ponto = df_planimetria[df_planimetria["Estaca"] == pto]["Ponto"].values[0]
                raio = df_velocidades[df_velocidades["Velocidade"] == velocidade]["Distância"].values[0]
                circle = ponto.buffer(raio)

                intersection = eixo.intersection(circle)
                x, y = intersection.xy

                if sentido == "C":
                    ponto_sentido = Point(x[-1], y[-1])
                else:
                    ponto_sentido = Point(x[0], y[0])

                diagonal = LineString((ponto, ponto_sentido))

                if diagonal.intersection(eixo_dir) or diagonal.intersection(eixo_esq):
                    typeline = 1
                else:
                    typeline = 2

                cache[pto] = typeline
                return typeline
            else:
                return None
        except:
            return cache[pto - 20]

    dict_analise = {}
    dict_sh = {}
    typeline_ref = 0
    indice = 0

    for row in df_analise.itertuples():
        Station = row[1]

        alt_c = verificar_alt(Station, "C")
        alt_d = verificar_alt(Station, "D")
        plan_c = verificar_plan(Station, "C")
        plan_d = verificar_plan(Station, "D")

        typeline_c = min(alt_c, plan_c)
        typeline_d = min(alt_d, plan_d)

        dict_analise[Station] = [alt_c, plan_c, typeline_c, alt_d, plan_d, typeline_d]

        if typeline_c == 1 and typeline_d == 1:
            typeline = 3
            layer = "LFO-3 (Cont./Cont.)"
        elif typeline_c == 1 and typeline_d == 2:
            typeline = 4
            layer = "LFO-4D (Descont./Cont.)"
        elif typeline_c == 2 and typeline_d == 1:
            typeline = 5
            layer = "LFO-4C (Cont./Descont.)"
        elif typeline_c == 2 and typeline_d == 2:
            typeline = 2
            layer = "LFO-2 (Descont.)"

        if typeline_ref == 0:
            typeline_ref = typeline
            Station_ref = Station
            layer_ref = layer
        elif typeline_ref != typeline:
            dict_sh[indice] = [Station_ref, Station, typeline_ref, layer_ref]
            indice += 1
            Station_ref = Station
            typeline_ref = typeline
            layer_ref = layer

    df_excel = pd.DataFrame.from_dict(
        dict_analise,
        orient="index",
        columns=[
            "Altimetria Crescente",
            "Planimetria Crescente",
            "Typeline Crescente",
            "Altimetria Decrescente",
            "Planimetria Decrescente",
            "Typeline Decrescente",
        ],
    ).reset_index(names="Station")

    #file_name = 'dataframe.xlsx' # Salvar dataframe em excel
    #df_excel.to_excel(file_name)
    df_dwg = pd.DataFrame.from_dict(dict_sh, orient="index", columns=["Km Inicial", "Km Final", "Typeline", "Layer"])
    df_dwg["Extensão"] = df_dwg["Km Final"] - df_dwg["Km Inicial"]
    nome_excel = caminho_salvar + "\\" + nome_arquivo + ".xlsx"
    df_dwg.to_excel(nome_excel, index=False)

    tf = time.time() # Final da contagem do programa rodando
    tempo_total = tf - ti

    Janela = ctk.CTk()
    messagebox.showinfo("Estudo de Visibilidade de Ultrapassagem",
                        "Análise concluída com sucesso! Tempo de execução: {:.2f} segundos".format(tempo_total))
    Janela.destroy()

def fechar():
    janela.destroy()

def criar_button(janela, texto, comando, x, y, fg_cor, hover_cor, height, width, corner_radius):
        botao = ctk.CTkButton(
        janela,
        text=texto,
        font=("Helvetica", 13, "bold"),
        command=comando,
        height=height,
        width=width,
        corner_radius=corner_radius,
        fg_color=fg_cor,
        hover_color=hover_cor,)
        botao.place(x=x, y=y)
        return botao

# Botões
but_abrir_alt = criar_button(janela, "Excel Altimetria", importar_alt, 10, 10, None, None, height=35, width=150, corner_radius=200)
but_abrir_plan = criar_button(janela, "Excel Planimetria", importar_plan, 165, 10, None, None, height=35, width=150, corner_radius=200)
but_abrir_nuvem = criar_button(janela, "Ex. Planilhas Planimetria e Altimetria", abrir_nuvem, 10, 50, None, None, height=25, width=305, corner_radius=200)
but_selecionar_pasta = criar_button(janela, 'Selecionar pasta para salvar Excel', selecionar_pasta, 10, 80, None, None, height=35, width=305, corner_radius=200)
but_verificar = criar_button(janela, "Verificar", verificar, 10, 295, None, None, height=35, width=150, corner_radius=200)
but_fechar = criar_button(janela, "Fechar", fechar, 165, 295, "red", "#960000", height=35, width=150, corner_radius=200)


janela.mainloop()