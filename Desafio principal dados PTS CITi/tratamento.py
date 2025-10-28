import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

try:
    # --- FASE 1: CARREGAMENTO DOS DADOS ---
    df = pd.read_csv('Base_Membros_Desempenho.csv')
    print("Dados carregados. Iniciando tratamento...")
    # --- FASE 2: TRATAMENTO DOS DADOS ---
    # 1. Padronizar Nivel_Senioridade
    mapa_senioridade = {
        'Jr': 'Júnior',
        'JR': 'Júnior',
        'P': 'Pleno',
        'pleno': 'Pleno',
        'Senior': 'Sênior', 
        'senior': 'Sênior',
        'N/D': np.nan 
    }
    df['Nivel_Senioridade'] = df['Nivel_Senioridade'].replace(mapa_senioridade)
    # Preencher os nulos (que agora incluem os 'N/D') com a moda
    moda_senioridade = df['Nivel_Senioridade'].mode()[0]
    df['Nivel_Senioridade'] = df['Nivel_Senioridade'].fillna(moda_senioridade)
    # 2. Padronizar Avaliacao_Tecnica e Avaliacao_Comportamental
    df['Avaliacao_Tecnica'] = pd.to_numeric(df['Avaliacao_Tecnica'], errors='coerce')
    df['Avaliacao_Comportamental'] = pd.to_numeric(df['Avaliacao_Comportamental'], errors='coerce')
    # Preencher nulos de Avaliacao_Tecnica com a média
    media_tecnica = df['Avaliacao_Tecnica'].mean()
    df['Avaliacao_Tecnica'] = df['Avaliacao_Tecnica'].fillna(media_tecnica)
    # Preencher nulos de Avaliacao_Comportamental com a média
    media_comportamental = df['Avaliacao_Comportamental'].mean()
    df['Avaliacao_Comportamental'] = df['Avaliacao_Comportamental'].fillna(media_comportamental)
    # 3. Tratar Engajamento_PIGs
    df['Engajamento_PIGs'] = df['Engajamento_PIGs'].replace({'%': '', 'N/A': np.nan}, regex=True)
    df['Engajamento_PIGs'] = pd.to_numeric(df['Engajamento_PIGs'], errors='coerce')
    df['Engajamento_PIGs'] = df['Engajamento_PIGs'] / 100
    media_engajamento = df['Engajamento_PIGs'].mean()
    df['Engajamento_PIGs'] = df['Engajamento_PIGs'].fillna(media_engajamento)
    print("--- FASE 2 CONCLUÍDA: Dados limpos e padronizados ---")
    # --- FASE 3: ENRIQUECIMENTO DOS DADOS ---
    # 4. Calcular Score_Desempenho
    df['Score_Desempenho'] = (df['Avaliacao_Tecnica'] * 0.5) + (df['Avaliacao_Comportamental'] * 0.5)
    # 5. Criar Status_Membro
    condicao_score = df['Score_Desempenho'] >= 7.0
    condicao_engajamento = df['Engajamento_PIGs'] >= 0.8
    df['Status_Membro'] = np.where(condicao_score & condicao_engajamento, 'Em Destaque', 'Padrão')
    print("--- FASE 3 CONCLUÍDA: Novas colunas criadas ---")
    # --- FASE 4: GERAÇÃO DOS ENTREGÁVEIS ---
    # --- SALVANDO EM EXCEL (.xlsx) PARA A SIDEQUEST ---
    # Vamos salvar o dataframe 'df' aqui, pois ele está com os números
    # puros (ex: 7.5), que é o ideal para o Google Planilhas e Looker Studio.
    try:
        # Arredondando os números antes de salvar
        df_para_excel = df.copy()
        colunas_num = ['Avaliacao_Tecnica', 'Avaliacao_Comportamental', 'Engajamento_PIGs', 'Score_Desempenho']
        df_para_excel[colunas_num] = df_para_excel[colunas_num].round(2)
        df_para_excel.to_excel("Base_Membros_Tratada.xlsx", index=False, engine='openpyxl')
        print("--- SUCESSO: 'Base_Membros_Tratada.xlsx' gerado com sucesso! ---")
    except Exception as e:
        print(f"Ocorreu um erro ao salvar o arquivo Excel: {e}")
    print("Iniciando formatação para o PDF...")
    # Criar um DataFrame final para exportação com vírgula decimal
    df_final_pdf = df.copy()
    colunas_numericas = ['Avaliacao_Tecnica', 'Avaliacao_Comportamental', 'Engajamento_PIGs', 'Score_Desempenho']
    for col in colunas_numericas:
        # Arredonda para 2 casas, converte para string e troca ponto por vírgula
        df_final_pdf[col] = df_final_pdf[col].round(2).astype(str).str.replace('.', ',')
    # --- Geração do PDF ---
    fig, ax = plt.subplots(figsize=(16, 10)) # Aumentei o tamanho para caber
    ax.axis('tight')
    ax.axis('off')
    # Criar a tabela
    tabela = ax.table(cellText=df_final_pdf.values, colLabels=df_final_pdf.columns, loc='center', cellLoc='center')
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(8) # Diminuí a fonte para caber mais dados
    tabela.scale(1.2, 1.2)
    # Salvar a figura como PDF
    plt.savefig('Base_Membros_Tratada.pdf', bbox_inches='tight', pad_inches=0.1)
    print("---------------------------------------------------------")
    print("SUCESSO! Arquivo 'Base_Membros_Tratada.pdf' gerado!")
    print("O desafio foi concluído.")
except FileNotFoundError:
    print("ERRO: O arquivo 'Base_Membros_Desempenho.csv' não foi encontrado.")
    print("Por favor, baixe o arquivo .csv do desafio e coloque-o na mesma pasta do script.")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")