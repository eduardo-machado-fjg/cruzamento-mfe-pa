import pandas as pd
import numpy as np

# ==========================================
# 1. CARREGAR E MESCLAR OS DADOS
# ==========================================
print("🔄 Carregando as bases de dados para análise de níveis...")
df_participacoes = pd.read_excel('Base People Analytics Filtrada.xlsx',sheet_name='Participações Filtrada')
df_capacitacoes = pd.read_excel('Base People Analytics Filtrada.xlsx',sheet_name = 'Capacitações')

# Selecionamos apenas as colunas que importam para não pesar a memória
colunas_capacitacoes = [
    'Código da Capacitação', 'Carga Horária', 'Modalidade', 
    'Liderança Colaborativa', 'Inovação', 'Compromisso Público', 
    'Resiliência', 'Visão Estratégica'
]

# Realizamos o Join (PROCV) pelo Código da Capacitação
df_final = pd.merge(
    df_participacoes,
    df_capacitacoes[colunas_capacitacoes],
    on='Código da Capacitação',
    how='left'
)

# ==========================================
# 2. LIMPEZA ESPECÍFICA PARA ESTA ANÁLISE
# ==========================================
# 2.1 Tratamento do Nível do Cargo
df_final['Nível do Cargo'] = df_final['Nível do Cargo'].astype(str).str.strip().str.title()
niveis_validos = ['Estratégico', 'Tático', 'Operacional']
df_niveis = df_final[df_final['Nível do Cargo'].isin(niveis_validos)].copy()

# 2.2 Tratamento da Carga Horária para formato numérico
def limpar_carga(valor):
    if pd.isna(valor):
        return np.nan
    valor_str = str(valor).replace(',', '.') 
    try:
        return float(valor_str)
    except:
        return np.nan

df_niveis['Carga_Num'] = df_niveis['Carga Horária'].apply(limpar_carga)

# ==========================================
# 3. RELATÓRIO: PERFIL DE APRENDIZAGEM POR NÍVEL
# ==========================================
print("\n" + "="*60)
print("🎯 RELATÓRIO: PERFIL DE APRENDIZAGEM POR NÍVEL DA PIRÂMIDE")
print("="*60)

if df_niveis.empty:
    print("⚠️ Não foram encontrados dados válidos de Nível do Cargo (Estratégico, Tático, Operacional).")
else:
    # ---------------------------------------------------------
    # Quantidade de Participações por Nível
    # ---------------------------------------------------------
    print("\n📊 VOLUME DE PARTICIPAÇÕES POR NÍVEL:")
    contagem_niveis = df_niveis['Nível do Cargo'].value_counts()
    for nivel, quantidade in contagem_niveis.items():
        print(f"   - {nivel}: {quantidade} participações")

    # ---------------------------------------------------------
    # ANÁLISE A: Carga Horária Média
    # ---------------------------------------------------------
    
    print("\n⏳ A. CARGA HORÁRIA MÉDIA POR NÍVEL:")
    
    carga_por_nivel = df_niveis.groupby('Nível do Cargo', observed=False)['Carga_Num'].mean()
    for nivel, horas in carga_por_nivel.items():
        print(f"   - {nivel}: {horas:.1f} horas em média por curso")

    # ---------------------------------------------------------
    # ANÁLISE B: Modalidade
    # ---------------------------------------------------------
    print("\n🏢 B. PREFERÊNCIA DE MODALIDADE (%):")
    modalidade_nivel = pd.crosstab(df_niveis['Nível do Cargo'], df_niveis['Modalidade'], normalize='index') * 100
    print(modalidade_nivel.round(1).astype(str) + '%')

    # ---------------------------------------------------------
    # ANÁLISE C: Competências
    # ---------------------------------------------------------
    print("\n🧠 C. BUSCA POR COMPETÊNCIAS (% das participações que trabalharam o tema):")
    
    competencias = ['Liderança Colaborativa', 'Inovação', 'Compromisso Público', 'Resiliência', 'Visão Estratégica']
    
    # Criamos as flags binárias (1 e 0) para conseguir calcular a porcentagem
    for comp in competencias:
        nome_flag = f"{comp}_Flag"
        df_niveis[nome_flag] = df_niveis[comp].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)
    
    flags_cols = [f"{comp}_Flag" for comp in competencias]
    comp_por_nivel = df_niveis.groupby('Nível do Cargo')[flags_cols].mean() * 100
    comp_por_nivel.columns = competencias # Renomeia para impressão limpa
    
    print(comp_por_nivel.round(1).astype(str) + '%')
    print("\n" + "="*60)

# ==========================================
# 4. LIMPEZA ESPECÍFICA PARA ESTA ANÁLISE
# ==========================================
# 4.1 Tratamento do Nível do Cargo
# Limpeza básica (tudo em maiúsculas e sem espaços extras)

df_final['Símbolo do Cargo (ou equivalente)'] = df_final['Símbolo do Cargo (ou equivalente)'].astype(str).str.strip().str.upper()
niveis_validos = ['DAS08', 'DAS09', 'DAS10']
df_niveis = df_final[df_final['Símbolo do Cargo (ou equivalente)'].isin(niveis_validos)].copy()

df_niveis['Carga_Num'] = df_niveis['Carga Horária'].apply(limpar_carga)

# ==========================================
# 5. RELATÓRIO: PERFIL DE APRENDIZAGEM POR NÍVEL
# ==========================================
print("\n" + "="*60)
print("🎯 RELATÓRIO: PERFIL DE APRENDIZAGEM POR SÍMBOLO DO CARGO")
print("="*60)

if df_niveis.empty:
    print("⚠️ Não foram encontrados dados válidos de Símbolo do Cargo (DAS10,DAS09,DAS08).")
else:
    # ---------------------------------------------------------
    # Quantidade de Participações por Nível
    # ---------------------------------------------------------
    print("\n📊 VOLUME DE PARTICIPAÇÕES POR SÍMBOLO:")
    contagem_niveis = df_niveis['Símbolo do Cargo (ou equivalente)'].value_counts()
    for nivel, quantidade in contagem_niveis.items():
        print(f"   - {nivel}: {quantidade} participações")
        
    # ---------------------------------------------------------
    # ANÁLISE A: Carga Horária Média
    # ---------------------------------------------------------
    
    print("\n⏳ A. CARGA HORÁRIA MÉDIA POR SÍMBOLO:")
    
    carga_por_nivel = df_niveis.groupby('Símbolo do Cargo (ou equivalente)', observed=False)['Carga_Num'].mean()
    for nivel, horas in carga_por_nivel.items():
        print(f"   - {nivel}: {horas:.1f} horas em média por curso")

    # ---------------------------------------------------------
    # ANÁLISE B: Modalidade
    # ---------------------------------------------------------
    print("\n🏢 B. PREFERÊNCIA DE MODALIDADE (%):")
    modalidade_nivel = pd.crosstab(df_niveis['Símbolo do Cargo (ou equivalente)'], df_niveis['Modalidade'], normalize='index') * 100
    print(modalidade_nivel.round(1).astype(str) + '%')

    # ---------------------------------------------------------
    # ANÁLISE C: Competências
    # ---------------------------------------------------------
    print("\n🧠 C. BUSCA POR COMPETÊNCIAS (% das participações que trabalharam o tema):")
    
    competencias = ['Liderança Colaborativa', 'Inovação', 'Compromisso Público', 'Resiliência', 'Visão Estratégica']
    
    # Criamos as flags binárias (1 e 0) para conseguir calcular a porcentagem
    for comp in competencias:
        nome_flag = f"{comp}_Flag"
        df_niveis[nome_flag] = df_niveis[comp].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)
    
    flags_cols = [f"{comp}_Flag" for comp in competencias]
    comp_por_nivel = df_niveis.groupby('Símbolo do Cargo (ou equivalente)')[flags_cols].mean() * 100
    comp_por_nivel.columns = competencias # Renomeia para impressão limpa
    
    print(comp_por_nivel.round(1).astype(str) + '%')
    print("\n" + "="*60)


    print('CATÁLOGO DE PARTICIPAÇÕES')
    total_cursos = len(df_capacitacoes)

    # ==========================================
# 2. RELATÓRIO DA OFERTA (PRATELEIRA DE CURSOS)
# ==========================================
print("\n" + "="*60)
print("📚 VISÃO DO CATÁLOGO DE CURSOS (OFERTA)")
print("="*60)
print(f"Total de cursos únicos mapeados na base: {total_cursos}")

# ---------------------------------------------------------
# A. Oferta por Modalidade
# ---------------------------------------------------------
print("\n🏢 TOTAL DE CURSOS OFERTADOS POR MODALIDADE:")
if 'Modalidade' in df_capacitacoes.columns:
    contagem_modalidade = df_capacitacoes['Modalidade'].value_counts()
    for modalidade, qtd in contagem_modalidade.items():
        porcentagem = (qtd / total_cursos) * 100
        print(f"   - {modalidade}: {qtd} cursos ({porcentagem:.1f}% do portfólio)")
else:
    print("   ⚠️ Coluna 'Modalidade' não encontrada.")

# ---------------------------------------------------------
# B. Oferta de Competências
# ---------------------------------------------------------
print("\n🧠 TOTAL DE CURSOS QUE DESENVOLVEM CADA COMPETÊNCIA:")
competencias = ['Liderança Colaborativa', 'Inovação', 'Compromisso Público', 'Resiliência', 'Visão Estratégica']

contagem_competencias = {}

# Mapeia cada competência e soma os "SIM"
for comp in competencias:
    if comp in df_capacitacoes.columns:
        qtd_sim = df_capacitacoes[comp].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0).sum()
        contagem_competencias[comp] = qtd_sim

# Ordena da competência mais ofertada para a menos ofertada
contagem_competencias = dict(sorted(contagem_competencias.items(), key=lambda item: item[1], reverse=True))

for comp, qtd in contagem_competencias.items():
    porcentagem = (qtd / total_cursos) * 100
    print(f"   - {comp}: {qtd} cursos (presente em {porcentagem:.1f}% da prateleira)")
    
print("\n" + "="*60)