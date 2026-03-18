import pandas as pd
from thefuzz import process 
import seaborn as sns
import matplotlib.pyplot as plt

# ==========================================
# 1. CARREGAR OS DADOS (Lendo abas do Excel)
# ==========================================
# Substitua pelo caminho completo caso o terminal não esteja na mesma pasta dos arquivos
df_participacoes = pd.read_excel('Base People Analytics.xlsx', sheet_name='Participações')
df_mfe = pd.read_excel('MFE.xlsx', sheet_name='Cálculo')

# Lendo a planilha De-Para (ignorando possíveis linhas vazias no cabeçalho. 
# Confirme se a aba se chama 'Cargos' ou ajuste o sheet_name)
df_depara = pd.read_excel('De-Para Cargos.xlsx', sheet_name='Cargos') 

# ==========================================
# 2. IA DE MAPEAMENTO DE CARGOS (DE-PARA)
# ==========================================
# Limpamos as colunas do De-Para para ter apenas valores únicos
cargos_pa = df_depara['Cargos People Analytics'].dropna().str.strip().str.upper().unique()
cargos_mfe = df_depara['Cargos MFE'].dropna().str.strip().str.upper().unique()

print("🤖 Gerando correspondência inteligente de cargos...")
dicionario_depara = {}

for cargo_pa in cargos_pa:
    # A IA busca o cargo do MFE que mais se parece com o cargo do People Analytics
    melhor_match, pontuacao = process.extractOne(cargo_pa, cargos_mfe)
    
    # Definimos um nível de confiança mínimo de 80% para aceitar o casamento
    if pontuacao >= 90:
        dicionario_depara[cargo_pa] = melhor_match
    else:
        dicionario_depara[cargo_pa] = None # Sem correspondência segura

pd.DataFrame(list(dicionario_depara.items()), columns=['Cargo PA', 'Cargo MFE']).to_excel('Auditoria_Mapeamento.xlsx', index=False)

# ==========================================
# 3. APLICAR O MAPEAMENTO E PREPARAR O CRUZAMENTO
# ==========================================
# Converte tudo para string para evitar erros se houver números ou valores vazios (NaN)
df_participacoes['Unidade Administrativa'] = df_participacoes['Unidade Administrativa'].astype(str).str.strip().str.upper()
df_participacoes['Cargo_Clean'] = df_participacoes['Cargo'].astype(str).str.strip().str.upper()

# Traduzindo o cargo do PA para o MFE usando a IA
df_participacoes['Cargo_Traduzido_MFE'] = df_participacoes['Cargo_Clean'].map(dicionario_depara)

df_mfe['Órgão_Clean'] = df_mfe['Órgão'].astype(str).str.strip().str.upper()
df_mfe['Nome do Cargo_Clean'] = df_mfe['Nome do Cargo\n(como está no SICI)'].astype(str).str.strip().str.upper()

# ==========================================
# 4. CRUZAMENTO DOS DADOS (MERGE)
# ==========================================
df_cruzado = pd.merge(
    df_participacoes, 
    df_mfe, 
    left_on=['Unidade Administrativa', 'Cargo_Traduzido_MFE'], 
    right_on=['Órgão_Clean', 'Nome do Cargo_Clean'], 
    how='inner' 
)

print(f"✅ Cruzamento finalizado! Linhas com correspondência exata: {len(df_cruzado)}")

# ==========================================
# 5. ANÁLISES
# ==========================================
if len(df_cruzado) > 0:
    print("\n--- 📊 Análise 1: Onde estamos investindo capacitação? (Por Nível Estratégico) ---")
    participacao_estrategica = df_cruzado['Nível estratégico'].value_counts(normalize=True) * 100
    print(participacao_estrategica.round(2).astype(str) + '%')

    print("\n--- 📉 Análise 2: Quem desiste mais? (Taxa de Conclusão por Escalão) ---")
    df_cruzado['Concluiu'] = df_cruzado['Certificado'].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)
    taxa_conclusao_escalao = df_cruzado.groupby('Escalão\n(como está na árvore do SICI)')['Concluiu'].mean() * 100
    print(taxa_conclusao_escalao.round(2).astype(str) + '%')

    print("\n--- 🎯 Análise 3: Cursos vs Competência de Alta Gestão ---")
    poder_orcamento_mask = df_cruzado['Poder de Decisão sobre o Orçamento'] != '1 - Não possui poder de decisão sobre alocação de recursos orçamentários no órgão'
    cursos_alta_gestao = df_cruzado[poder_orcamento_mask]['Capacitação'].value_counts().head(5)
    print("Top 5 cursos feitos por líderes com poder de decisão orçamentária:")
    print(cursos_alta_gestao)

    print("\n" + "="*50)
    print("🎯 EXPANSÃO DA ANÁLISE 3: ALINHAMENTO ESTRATÉGICO POR ESCALÃO")
    print("="*50)

    escaloes = ['1º', '2º', '3º']

    # Vamos analisar tudo separadamente POR ESCALÃO
    for escalao in escaloes:
        print(f"\n" + "-"*50)
        
        print(f"🏢 [ {escalao} ESCALÃO ]")
        
        # Descobre o total de cadeiras mapeadas no MFE para esse escalão
        total_cadeiras_mfe = len(df_mfe[df_mfe['Escalão\n(como está na árvore do SICI)'] == escalao])
        
        # Filtramos a base cruzada para olhar APENAS para o escalão da vez
        df_escalao = df_cruzado[df_cruzado['Escalão\n(como está na árvore do SICI)'] == escalao]
        # DIAGNÓSTICO: Quantos cargos desse escalão exigem habilidade gerencial?
        contagem_gerencial = df_escalao['Requer Habilidade Técnica'].str.strip().str.upper().value_counts()
        print(f"📊 Exigência Habilidade Técnica neste escalão:\n{contagem_gerencial.to_string()}\n")
        # Imprime o contexto populacional
        print(f"Total de cadeiras mapeadas na prefeitura (MFE): {total_cadeiras_mfe}")
        print(f"Total de participações em cursos neste escalão: {len(df_escalao)}")
        print("-" * 50)
        
        if df_escalao.empty:
            print("⚠️ Sem dados de capacitação cruzados para este escalão.")
            continue # Pula para o próximo escalão se não tiver ninguém
            
        # Top Cursos Gerais daquele escalão
        print("\n🏆 Top 3 Cursos Gerais:")
        print(df_escalao['Capacitação'].value_counts().head(3).to_string())
        
        # Top Cursos (Habilidade Política) daquele escalão
        print("\n🗣️ Top 3 Cursos (Exige Alta HABILIDADE POLÍTICA):")
        mask_politica = df_escalao['Requer Habilidade Política'].str.strip().str.upper() == 'SIM'
        cursos_politica = df_escalao[mask_politica]['Capacitação'].value_counts().head(10)
        if not cursos_politica.empty:
            print(cursos_politica.to_string())
        else:
            print("Nenhum registro encontrado.")

        # Top Cursos (Habilidade Técnica) daquele escalão
        print("\n👥 Top 5 Cursos (Exige Habilidade Técnica):")
        mask_gerencial = df_escalao['Requer Habilidade Técnica'].str.strip().str.upper() == 'SIM'
        cursos_gerencial = df_escalao[mask_gerencial]['Capacitação'].value_counts().head(10)
        if not cursos_gerencial.empty:
            print(cursos_gerencial.to_string())
        else:
            print("Nenhum registro encontrado.")

        # Top Cursos (Habilidade Gerencial) daquele escalão
        print("\n👥 Top 5 Cursos (Exige Habilidade Gerencial):")
        mask_gerencial = df_escalao['Requer Habilidade Gerencial'].str.strip().str.upper() == 'SIM'
        cursos_gerencial = df_escalao[mask_gerencial]['Capacitação'].value_counts().head(10)
        if not cursos_gerencial.empty:
            print(cursos_gerencial.to_string())
        else:
            print("Nenhum registro encontrado.")

        # Top Cursos (Policy making) daquele escalão
        print("\n👥 Top 5 Cursos (Exige Policy Making):")
        mask_gerencial = df_escalao['Requer Policy Making'].str.strip().str.upper() == 'SIM'
        cursos_gerencial = df_escalao[mask_gerencial]['Capacitação'].value_counts().head(10)
        if not cursos_gerencial.empty:
            print(cursos_gerencial.to_string())
        else:
            print("Nenhum registro encontrado.")
else:
    print("\n⚠️ Nenhuma linha cruzou. Considere baixar o limite de pontuação da IA (de 75 para 60) ou exportar a planilha de auditoria para verificar o mapeamento do de-para de órgãos.")





print("\n" + "="*50)
print("🌍 ANÁLISE 4: DIVERSIDADE E INCLUSÃO NA PIRÂMIDE ESTRATÉGICA")
print("="*50)

# 4.1. Gênero x Escalão
print("\n--- 🚺 Distribuição de Gênero por Escalão (%) ---")
# pd.crosstab cruza as duas colunas. O normalize='index' garante que a soma da linha dê 100%
genero_escalao = pd.crosstab(df_cruzado['Escalão\n(como está na árvore do SICI)'], df_cruzado['Sexo'], normalize='index') * 100
print(genero_escalao.round(1).astype(str) + '%')

# 4.2. Raça/Cor x Escalão
print("\n--- ✊ Distribuição de Raça/Cor por Escalão (%) ---")
raca_escalao = pd.crosstab(df_cruzado['Escalão\n(como está na árvore do SICI)'], df_cruzado['Raça/Cor'], normalize='index') * 100
print(raca_escalao.round(1).astype(str) + '%')

print("\n" + "="*50)
print("🗺️ ANÁLISE 5: MAPA DE CAPACITAÇÃO POR MACRO ÁREA")
print("="*50)

# Para não poluir a tela com dezenas de cursos, vamos isolar os 5 cursos mais procurados na prefeitura toda
top_5_cursos = df_cruzado['Capacitação'].value_counts().head(5).index
df_top_cursos = df_cruzado[df_cruzado['Capacitação'].isin(top_5_cursos)]

# 5.1. Visão Tabular (No Terminal)
print("\n--- 📊 Onde está o engajamento dos Top 5 Cursos? (Qtd de Participações por Macro Área) ---")
mapa_texto = pd.crosstab(df_top_cursos['Macro Área\n(referente ao órgão)'], df_top_cursos['Capacitação'])
print(mapa_texto.to_string())

# 5.2. Gerando uma Imagem do Mapa de Calor (Heatmap)
try:
    print("\n🎨 Gerando imagem do Mapa de Calor ('mapa_calor_areas.png')...")
    plt.figure(figsize=(12, 6))
    
    # Criamos o gráfico usando a biblioteca seaborn
    sns.heatmap(mapa_texto, annot=True, cmap="YlGnBu", fmt='g', linewidths=.5)
    
    plt.title('Mapa de Calor: Concentração dos Top 5 Cursos por Macro Área', fontsize=14, pad=15)
    plt.xlabel('Cursos', fontsize=12)
    plt.ylabel('Macro Área', fontsize=12)
    plt.xticks(rotation=45, ha='right') # Inclina o nome dos cursos para caber na tela
    plt.tight_layout() # Ajusta as margens para não cortar texto
    
    # Salva o gráfico na mesma pasta onde seu script está rodando
    plt.savefig('mapa_calor_areas.png', dpi=300)
    print("✅ Imagem salva com sucesso! Procure pelo arquivo 'mapa_calor_areas.png' na sua pasta.")
except Exception as e:
    print(f"⚠️ Não foi possível gerar o gráfico (verifique se matplotlib e seaborn estão instalados). Erro: {e}")