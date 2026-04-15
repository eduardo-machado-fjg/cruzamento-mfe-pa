import pandas as pd

print("🔄 Carregando a base de dados para análise de conclusão por sexo...")
df_participacoes = pd.read_excel('Base People Analytics Filtrada.xlsx', sheet_name='Participações Filtrada')

# ==========================================
# 1. LIMPEZA DOS DADOS
# ==========================================
# Padroniza a coluna Sexo
df_participacoes['Sexo'] = df_participacoes['Sexo'].astype(str).str.strip().str.title()

# Cria uma coluna numérica 'Concluiu' onde SIM = 1 e qualquer outra coisa (NÃO, vazio) = 0
df_participacoes['Concluiu'] = df_participacoes['Certificado'].apply(lambda x: 1 if str(x).strip().upper() == 'SIM' else 0)

# ==========================================
# 2. CALCULAR TAXA DE CONCLUSÃO
# ==========================================
# Agrupamos por sexo e pedimos 3 cálculos: Total de inscritos (count), Total de certificados (sum) e a Taxa (mean)
resumo_sexo = df_participacoes.groupby('Sexo').agg(
    Total_Participacoes=('Concluiu', 'count'),
    Total_Concluidos=('Concluiu', 'sum'),
    Taxa_Conclusao=('Concluiu', 'mean')
).reset_index()

# Formata a taxa para virar porcentagem bonita
resumo_sexo['Taxa_Conclusao'] = (resumo_sexo['Taxa_Conclusao'] * 100).round(1)

# Ordena do maior pro menor
resumo_sexo = resumo_sexo.sort_values(by='Taxa_Conclusao', ascending=False)

# ==========================================
# 3. IMPRIMIR RELATÓRIO
# ==========================================
print("\n" + "="*60)
print("🎓 TAXA DE CONCLUSÃO DE CAPACITAÇÕES POR SEXO")
print("="*60)

for index, row in resumo_sexo.iterrows():
    sexo = row['Sexo']
    # Ignora valores em branco ou "Sem Informação" para focar no comparativo direto
    if sexo not in ['Nan', 'Sem Informação']: 
        taxa = row['Taxa_Conclusao']
        total = row['Total_Participacoes']
        concluidos = row['Total_Concluidos']
        print(f"   - {sexo}: {taxa}% (Concluiu {concluidos} de {total} participações)")

print("="*60 + "\n")