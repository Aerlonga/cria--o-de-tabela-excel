import os
import pandas as pd

def extrair_informacoes(planilha_original_path, planilha_nova_path):
    try:
        df_original = pd.read_excel(planilha_original_path, header=None)
        destino = df_original.iloc[2, 0]
        patrimonio = df_original.iloc[7:9, 2].tolist()
        serie = df_original.iloc[7:9, 3].tolist()

        # Verificando se as listas têm o mesmo comprimento
        if len(patrimonio) != len(serie):
            raise ValueError("As listas 'patrimonio' e 'serie' devem ter o mesmo comprimento.")

        df_nova = pd.DataFrame({
            'Destino': [destino] * len(patrimonio),
            'Patrimônio': patrimonio,
            'Série': serie,
        })

        df_nova.to_excel(planilha_nova_path, index=False, engine='openpyxl')
        print(f"Nova planilha gerada com sucesso: {planilha_nova_path}")

    except Exception as e:
        print(f"Erro ao processar a planilha: {e}")

def processar_planilhas_no_diretorio(diretorio_origem, diretorio_destino):
    os.makedirs(diretorio_destino, exist_ok=True)

    for arquivo in os.listdir(diretorio_origem):
        if arquivo.endswith(".xlsx"):
            caminho_origem = os.path.join(diretorio_origem, arquivo)
            caminho_destino = os.path.join(diretorio_destino, arquivo)
            extrair_informacoes(caminho_origem, caminho_destino)

diretorio_origem = 'C:/Users/aerlon.alves/Desktop/DFI-DESKTOPS' #origem dos arquivos excel
diretorio_destino = 'C:/Users/aerlon.alves/Desktop/DFI-DESKTOPS/PlanilhasExtraidas' #pasta criada e local onde as planilhas serão salvas

processar_planilhas_no_diretorio(diretorio_origem, diretorio_destino)

import os
import pandas as pd

def unir_planilhas(diretorio_origem, diretorio_destino, nome_arquivo_destino):
    dfs = []

    for arquivo in os.listdir(diretorio_origem):
        if arquivo.endswith(".xlsx"):
            caminho_planilha = os.path.join(diretorio_origem, arquivo)
            df = pd.read_excel(caminho_planilha)
            dfs.append(df)

    # Unir os DataFrames em um único DataFrame
    df_final = pd.concat(dfs, ignore_index=True)

    # Salvar o DataFrame final em uma nova planilha
    caminho_destino = os.path.join(diretorio_destino, nome_arquivo_destino)
    df_final.to_excel(caminho_destino, index=False, engine='openpyxl')
    print(f"Planilhas unidas e salvas em: {caminho_destino}")

# Exemplo de uso
diretorio_origem_planilhas = 'C:/Users/aerlon.alves/Desktop/DFI-DESKTOPS/PlanilhasExtraidas'
diretorio_destino_uniao = 'C:/Users/aerlon.alves/Desktop/DFI-DESKTOPS'
nome_arquivo_destino = 'PlanilhaUnida.xlsx'

unir_planilhas(diretorio_origem_planilhas, diretorio_destino_uniao, nome_arquivo_destino)

