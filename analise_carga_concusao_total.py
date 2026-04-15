import pandas as pd
import numpy as np

print("🔄 Carregando as bases de dados para análise de Carga Horária (Panorama Geral)...")
df_participacoes = pd.read_excel('Base People Analytics Filtrada.xlsx',sheet_name = 'Participações Filtrada')
df_capacitacoes =  pd.read_excel('Base People Analytics Filtrada.xlsx',sheet_name = 'Capacitações')

# ==========================================
# 1. CRUZAMENTO E LIMPEZA
# ==========================================
# Fazemos o Join para trazer a Carga Horária de cada curso
df_final = pd.merge(
    df_participacoes,
    df_capacitacoes[['Código da Capacitação', 'Carga Horária']],
    on='Código da Capacitação',
    how='left'
)

# Transforma a carga horária em número (ex: "16,00" vira 16.0)
def limpar_carga(valor):
    if pd.isna(valor):
        return np.nan
    valor_str = str(valor).replace(',', '.') 
    try:
        return float(valor_str)
    except:
        return np.nan

df_final['Carga_Num'] = df_final['Carga Horária'].apply(limpar_carga)

# Transforma o Status em Binário (1 = SIM, 0 = NÃO/Vazio)
df_final['Concluiu'] = df_final['Certificado'].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)

# ==========================================
# 2. CATEGORIZAR POR DURAÇÃO
# ==========================================
# Cria as "caixinhas" de horas
bins = [0, 8, 20, 1000]
labels = ['1. Curta (até 8h)', '2. Média (9h-20h)', '3. Longa/Extensa (+20h)']
df_final['Categoria de Duração'] = pd.cut(df_final['Carga_Num'], bins=bins, labels=labels)

# ==========================================
# 3. CALCULAR TAXA E GERAR RELATÓRIO
# ==========================================
resumo_carga = df_final.groupby('Categoria de Duração', observed=False).agg(
    Total_Participacoes=('Concluiu', 'count'),
    Total_Concluidos=('Concluiu', 'sum'),
    Taxa_Conclusao=('Concluiu', 'mean')
).reset_index()

# Formata para porcentagem
resumo_carga['Taxa_Conclusao'] = (resumo_carga['Taxa_Conclusao'] * 100).round(1)

print("\n" + "="*60)
print("⏳ TAXA DE CONCLUSÃO GERAL POR TAMANHO DO CURSO")
print("="*60)

for index, row in resumo_carga.dropna(subset=['Categoria de Duração']).iterrows():
    categoria = row['Categoria de Duração']
    taxa = row['Taxa_Conclusao']
    total = int(row['Total_Participacoes'])
    concluidos = int(row['Total_Concluidos'])
    
    print(f"   - {categoria}: {taxa}% (Concluiu {concluidos} de {total} participações)")

print("="*60 + "\n")