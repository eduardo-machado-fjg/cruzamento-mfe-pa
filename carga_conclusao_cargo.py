import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt

# ==========================================
# 1. CARREGAR E MESCLAR OS DADOS
# ==========================================
print("🔄 Carregando as bases de dados...")
df_participacoes = pd.read_excel('Base People Analytics Filtrada.xlsx',sheet_name = 'Participações Filtrada')
df_capacitacoes =  pd.read_excel('Base People Analytics Filtrada.xlsx',sheet_name = 'Capacitações')

# Mesclamos trazendo apenas a Carga Horária da base de Capacitações
df_final = pd.merge(
    df_participacoes,
    df_capacitacoes[['Código da Capacitação', 'Carga Horária']],
    on='Código da Capacitação',
    how='left'
)

# ==========================================
# 2. LIMPEZA DOS DADOS E TRATAMENTO DE HORAS
# ==========================================
# A Carga Horária precisa virar número (pode vir como '16,00' no Excel)
def limpar_carga(valor):
    if pd.isna(valor):
        return np.nan
    valor_str = str(valor).replace(',', '.') 
    try:
        return float(valor_str)
    except:
        return np.nan

df_final['Carga_Num'] = df_final['Carga Horária'].apply(limpar_carga)

# Convertendo o Status de Certificação em 1 (Concluiu) e 0 (Evadiu)
df_final['Concluiu'] = df_final['Certificado'].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)

# Limpando Cargo
df_final['Cargo_Clean'] = df_final['Cargo'].astype(str).str.strip().str.upper()

# ==========================================
# 3. ANÁLISE 1: CORRELAÇÃO GERAL
# ==========================================
print("\n" + "="*50)
print("⏳ CORRELAÇÃO: CARGA HORÁRIA x CONCLUSÃO")
print("="*50)

# Correlação de Pearson (-1 a 1)
correlacao = df_final['Carga_Num'].corr(df_final['Concluiu'])
print(f"📊 Correlação Global (Pearson): {correlacao:.3f}")
if -0.1 < correlacao < 0.1:
    print("💡 Insight: A correlação é praticamente zero (estatisticamente nula). Cursos mais longos NÃO estão causando mais abandono!")

# ==========================================
# 4. ANÁLISE 2: TAXA DE CONCLUSÃO POR CARGO E DURAÇÃO
# ==========================================
# Categorizando a carga horária em 3 tamanhos
bins = [0, 8, 20, 1000] # Até 8h, de 9 a 20h, e mais de 20h
labels = ['1. Curta (até 8h)', '2. Média (9h-20h)', '3. Longa/Extensa (+20h)']
df_final['Categoria de Duração'] = pd.cut(df_final['Carga_Num'], bins=bins, labels=labels)

# Selecionamos os 5 cargos mais frequentes para evitar amostras muito pequenas
top_5_cargos = df_final['Cargo_Clean'].value_counts().head(5).index
df_top_cargos = df_final[df_final['Cargo_Clean'].isin(top_5_cargos)]

# Calculando a Taxa (Média) e Quantidade (Count)
resumo = df_top_cargos.groupby(['Cargo_Clean', 'Categoria de Duração'], observed=False)['Concluiu'].agg(['mean', 'count']).dropna()
resumo = resumo[resumo['count'] > 0] # Remove categorias vazias
resumo['Taxa de Conclusão'] = (resumo['mean'] * 100).round(1).astype(str) + '%'

print("\n🏢 Taxa de Conclusão nos Top 5 Cargos por Tamanho do Curso:")
print(resumo[['count', 'Taxa de Conclusão']].rename(columns={'count': 'Total Matriculados'}).to_string())

# ==========================================
# 5. GERANDO UM GRÁFICO (Opcional)
# ==========================================
try:
    print("\n🎨 Gerando gráfico ('conclusao_carga_cargo.png')...")
    dados_grafico = resumo.reset_index()
    dados_grafico['Taxa (%)'] = dados_grafico['mean'] * 100
    
    plt.figure(figsize=(10, 6))
    # O barplot vai cruzar o Cargo, a Taxa e pintar a barra pelo Tamanho do Curso
    sns.barplot(data=dados_grafico, x='Cargo_Clean', y='Taxa (%)', hue='Categoria de Duração', palette='Blues')
    
    plt.title('Taxa de Conclusão por Cargo e Carga Horária do Treinamento', pad=15)
    plt.xlabel('Cargo / Posição', fontsize=12)
    plt.ylabel('Taxa de Certificação (%)', fontsize=12)
    plt.ylim(0, 100) # Eixo Y até 100%
    plt.legend(title='Carga Horária')
    plt.tight_layout()
    plt.savefig('conclusao_carga_cargo.png', dpi=300)
    print("✅ Imagem salva na pasta do script!")
except Exception as e:
    print(f"⚠️ Erro ao gerar gráfico: {e}")