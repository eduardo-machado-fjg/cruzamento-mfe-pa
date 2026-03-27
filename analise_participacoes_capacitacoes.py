import pandas as pd

# ==========================================
# 1. CARREGAR OS DADOS
# ==========================================
print("🔄 Carregando as bases de dados...")
# Substitua por pd.read_excel se estiver usando os arquivos .xlsx diretamente
df_participacoes = pd.read_excel('Base People Analytics Filtrada.xlsx', sheet_name='Participações Filtrada')
df_capacitacoes = pd.read_excel('Base People Analytics Filtrada.xlsx', sheet_name='Capacitações')
df_mfe = pd.read_excel('MFE.xlsx', sheet_name='Cálculo')

# ==========================================
# 2. CRIAR O DE-PARA DE MACRO ÁREAS (Usando o MFE)
# ==========================================
# Pegamos apenas valores únicos de Órgão e Macro Área do MFE
df_mfe_orgaos = df_mfe[['Órgão', 'Macro Área\n(referente ao órgão)']].dropna().drop_duplicates()
df_mfe_orgaos['Órgão'] = df_mfe_orgaos['Órgão'].astype(str).str.strip().str.upper()

# Cria um dicionário no formato {'SMS': 'Social', 'SMF': 'Gestão'}
dicionario_macro_area = dict(zip(df_mfe_orgaos['Órgão'], df_mfe_orgaos['Macro Área\n(referente ao órgão)']))

# Tratando a Unidade Administrativa nas Participações para garantir o cruzamento
df_participacoes['Unidade_Clean'] = df_participacoes['Unidade Administrativa'].astype(str).str.strip().str.upper()
# Caso algum órgão venha no formato "Secretaria Municipal de Saúde / SMS", extraímos apenas a sigla "SMS"
df_participacoes['Unidade_Clean'] = df_participacoes['Unidade_Clean'].apply(lambda x: x.split('/')[-1].strip() if '/' in x else x)

# Aplicando a Macro Área
df_participacoes['Macro Área'] = df_participacoes['Unidade_Clean'].map(dicionario_macro_area)

# Verificando se algum órgão ficou sem Macro Área mapeada
orgaos_sem_macro = df_participacoes[df_participacoes['Macro Área'].isna()]['Unidade_Clean'].unique()
if len(orgaos_sem_macro) > 0:
    print(f"⚠️ Aviso: Algumas siglas não foram encontradas no MFE e ficarão sem Macro Área: {orgaos_sem_macro}")

# ==========================================
# 3. O GRANDE JOIN: PARTICIPAÇÕES + CAPACITAÇÕES
# ==========================================
# Selecionamos apenas as colunas de interesse de Capacitações para não duplicar dados desnecessários (como 'Ano' e 'Capacitação')
colunas_capacitacoes = [
    'Código da Capacitação', 'Tipo de Capacitação', 'Carga Horária', 
    'Modalidade', 'Portfólio', 'Liderança Colaborativa', 'Inovação', 
    'Compromisso Público', 'Resiliência', 'Visão Estratégica'
]

# Realiza o PROCV/Merge pelo Código da Capacitação
df_final = pd.merge(
    df_participacoes,
    df_capacitacoes[colunas_capacitacoes],
    on='Código da Capacitação',
    how='left'
)

print(f"✅ Join finalizado! A base agora possui {len(df_final)} linhas cruzadas.")

# Opcional: Salvar a base final super completa para usar no Excel ou Power BI depois
# df_final.to_excel('Participacoes_Enriquecidas.xlsx', index=False)

# ==========================================
# 4. ANÁLISES ESTRATÉGICAS SUGERIDAS
# ==========================================

print("\n" + "="*50)
print("🎯 ANÁLISE 1: ENGAJAMENTO POR MACRO ÁREA")
print("="*50)
# Mostra qual proporção das capacitações foi feita por qual Macro Área
engajamento_macro = df_final['Macro Área'].value_counts(normalize=True) * 100
print(engajamento_macro.round(1).astype(str) + '% das participações')

print("\n" + "="*50)
print("🎯 ANÁLISE 2: PREFERÊNCIA DE MODALIDADE (ONLINE VS PRESENCIAL)")
print("="*50)
# Cruza a Macro Área com o formato do curso (Online/Presencial/Híbrido)
if 'Modalidade' in df_final.columns:
    modalidade_macro = pd.crosstab(df_final['Macro Área'], df_final['Modalidade'], normalize='index') * 100
    print("Distribuição percentual da escolha de modalidade por área:")
    print(modalidade_macro.round(1).astype(str) + '%')
else:
    print("Coluna 'Modalidade' não encontrada no cruzamento.")

print("\n" + "="*50)
print("🎯 ANÁLISE 3: DESENVOLVIMENTO DE COMPETÊNCIAS POR MACRO ÁREA")
print("="*50)
# Vamos ver quantas vezes a competência "Inovação" vs "Resiliência" foi trabalhada por Macro Área
competencias = ['Liderança Colaborativa', 'Inovação', 'Compromisso Público', 'Resiliência', 'Visão Estratégica']

print("Qual porcentagem das participações de cada Macro Área incluiu a competência de INOVAÇÃO?")
# Mapeia Sim = 1 e Não = 0
df_final['Inovação_Flag'] = df_final['Inovação'].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)
taxa_inovacao = df_final.groupby('Macro Área')['Inovação_Flag'].mean() * 100
print(taxa_inovacao.round(1).astype(str) + '%\n')

print("Qual porcentagem das participações de cada Macro Área incluiu a competência de RESILIÊNCIA?")
df_final['Resiliencia_Flag'] = df_final['Resiliência'].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)
taxa_resiliencia = df_final.groupby('Macro Área')['Resiliencia_Flag'].mean() * 100
print(taxa_resiliencia.round(1).astype(str) + '%')

print("Qual porcentagem das participações de cada Macro Área incluiu a competência de Liderança Colaborativa?")
df_final['Liderança_Flag'] = df_final['Liderança Colaborativa'].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)
taxa_lideranca = df_final.groupby('Macro Área')['Liderança_Flag'].mean() * 100
print(taxa_lideranca.round(1).astype(str) + '%')

print("Qual porcentagem das participações de cada Macro Área incluiu a competência de Compromisso Público?")
df_final['Compromisso_Flag'] = df_final['Compromisso Público'].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)
taxa_compromisso = df_final.groupby('Macro Área')['Compromisso_Flag'].mean() * 100
print(taxa_compromisso.round(1).astype(str) + '%')

print("Qual porcentagem das participações de cada Macro Área incluiu a competência de Visão Estratégica?")
df_final['Visao_Flag'] = df_final['Visão Estratégica'].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)
taxa_visao = df_final.groupby('Macro Área')['Visao_Flag'].mean() * 100
print(taxa_visao.round(1).astype(str) + '%')

# ---------------------------------------------------------
# NOVAS ANÁLISES DE DIVERSIDADE (GÊNERO E RAÇA)
# ---------------------------------------------------------

print("\n" + "="*50)
print("🌍 ANÁLISE 4: PERFIL DE GÊNERO NAS CAPACITAÇÕES")
print("="*50)
# 4.1 Visão Geral
genero_total = df_final['Gênero'].value_counts(normalize=True) * 100
print("--- Distribuição Geral de Gênero na Prefeitura ---")
print(genero_total.round(1).astype(str) + '%\n')

# 4.2 Gênero x Macro Área
genero_macro = pd.crosstab(df_final['Macro Área'], df_final['Gênero'], normalize='index') * 100
print("--- Distribuição de Gênero por Macro Área ---")
print(genero_macro.round(1).astype(str) + '%')


print("\n" + "="*50)
print("✊ ANÁLISE 5: PERFIL DE RAÇA/COR NAS CAPACITAÇÕES")
print("="*50)
# 5.1 Visão Geral
raca_total = df_final['Raça/Cor'].value_counts(normalize=True) * 100
print("--- Distribuição Geral de Raça/Cor na Prefeitura ---")
print(raca_total.round(1).astype(str) + '%\n')

# 5.2 Raça/Cor x Macro Área
raca_macro = pd.crosstab(df_final['Macro Área'], df_final['Raça/Cor'], normalize='index') * 100
print("--- Distribuição de Raça/Cor por Macro Área ---")
print(raca_macro.round(1).astype(str) + '%')