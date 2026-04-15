import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap

print("🔄 Preparando o Mapa de Calor Limpo (Para Canva)...")

# ==========================================
# 1. CARREGAR E PREPARAR OS DADOS (Mesmo processo)
# ==========================================
df_participacoes = pd.read_excel('Base People Analytics Filtrada.xlsx', sheet_name='Participações Filtrada')
df_capacitacoes = pd.read_excel('Base People Analytics Filtrada.xlsx', sheet_name='Capacitações')
df_mfe = pd.read_excel('MFE.xlsx', sheet_name='Cálculo')

df_mfe_orgaos = df_mfe[['Órgão', 'Macro Área\n(referente ao órgão)']].dropna().drop_duplicates()
df_mfe_orgaos['Órgão'] = df_mfe_orgaos['Órgão'].astype(str).str.strip().str.upper()
dicionario_macro_area = dict(zip(df_mfe_orgaos['Órgão'], df_mfe_orgaos['Macro Área\n(referente ao órgão)']))

df_participacoes['Unidade_Clean'] = df_participacoes['Unidade Administrativa'].astype(str).str.strip().str.upper()
df_participacoes['Unidade_Clean'] = df_participacoes['Unidade_Clean'].apply(lambda x: x.split('/')[-1].strip() if '/' in x else x)
df_participacoes['Macro Área'] = df_participacoes['Unidade_Clean'].map(dicionario_macro_area)

competencias = ['Inovação', 'Liderança Colaborativa', 'Compromisso Público', 'Resiliência', 'Visão Estratégica']
colunas_capacitacoes = ['Código da Capacitação'] + competencias

df_final = pd.merge(df_participacoes, df_capacitacoes[colunas_capacitacoes], on='Código da Capacitação', how='left')

for comp in competencias:
    df_final[comp] = df_final[comp].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)

df_heatmap = (df_final.groupby('Macro Área')[competencias].mean() * 100).T 

# ==========================================
# 2. CONFIGURAR A PALETA DE CORES
# ==========================================
cores_paleta = [
    "#F0F4F8", # Cor 1 (Mais clara)
    "#9CB4CC", # Cor 2 
    "#4B779E", # Cor 3 
    "#1A426B"  # Cor 4 (Mais escura)
]
meu_cmap = LinearSegmentedColormap.from_list("MinhaPaleta", cores_paleta)

# ==========================================
# 3. PLOTAR GRÁFICO "LIMPO" PARA O CANVA
# ==========================================
# Tiramos a borda padrão criando a figura com frameon=False
fig, ax = plt.subplots(figsize=(10, 6))

sns.heatmap(
    df_heatmap, 
    cmap=meu_cmap,          
    annot=True,             # Mude para False se quiser remover os números também
    fmt=".1f",              
    linewidths=3,           # Espaçamento um pouco maior para ficar elegante
    xticklabels=False,      # REMOVE os textos da Macro Área (Eixo X)
    yticklabels=False,      # REMOVE os textos das Competências (Eixo Y)
    cbar_kws={'label': ''}, # REMOVE o título da barra lateral de cores
    annot_kws={"size": 14, "weight": "bold", "family": "sans-serif"} # Deixa os números neutros
)

# Remove os títulos dos eixos
plt.title('')
plt.ylabel('')
plt.xlabel('')

# Remove os "tracinhos" (ticks) dos eixos que ficam sobrando
ax.tick_params(axis='both', which='both', length=0)

plt.tight_layout()

# ==========================================
# 4. SALVAR COM FUNDO TRANSPARENTE
# ==========================================
nome_arquivo = 'heatmap_canva.png'
# O argumento transparent=True é a mágica para o Canva!
plt.savefig(nome_arquivo, dpi=300, bbox_inches='tight', transparent=True)
print(f"✅ Gráfico limpo gerado e salvo como '{nome_arquivo}' (com fundo transparente!).")