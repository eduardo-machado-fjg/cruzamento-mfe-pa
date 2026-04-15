import pandas as pd

print("🔄 Carregando as bases de dados para analisar os Órgãos...")
df_participacoes = pd.read_excel('Base People Analytics Filtrada.xlsx',sheet_name='Participações Filtrada')

# 1. Limpeza do nome da Unidade para extrair apenas a Sigla (ex: "Secretaria de Saúde / SMS" vira "SMS")
df_participacoes['Unidade_Clean'] = df_participacoes['Unidade Administrativa'].astype(str).str.strip().str.upper()
df_participacoes['Órgão'] = df_participacoes['Unidade_Clean'].apply(lambda x: x.split('/')[-1].strip() if '/' in x else x)

# 2. Contagem do Top 15 Órgãos
print("\n" + "="*50)
print("🏢 RANKING: TOP 15 ÓRGÃOS COM MAIS PARTICIPAÇÕES")
print("="*50)

# Pega o total geral de participações na base
total_geral = len(df_participacoes)

ranking_orgaos = df_participacoes['Órgão'].value_counts().head(15)

for i, (orgao, quantidade) in enumerate(ranking_orgaos.items(), start=1):
    # Calcula qual porcentagem do programa aquele órgão representa
    porcentagem = (quantidade / total_geral) * 100
    print(f"{i}º LUGAR: {orgao} com {quantidade} participações ({porcentagem:.1f}%)")
    
print("\n" + "="*50)