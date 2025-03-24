# Dashboard-de-vendas-
import openpyxl
import matplotlib.pyplot as plt
import pandas as pd

# Carregar a planilha
caminho_arquivo = r"\\caminho=dashboard_vendas.xlsx"
workbook = openpyxl.load_workbook(caminho_arquivo)
sheet = workbook.active  # Seleciona a planilha ativa

# Ler os dados da planilha para um DataFrame do pandas
dados = pd.DataFrame(sheet.values)
dados.columns = dados.iloc[0]  # Define a primeira linha como cabeçalho
dados = dados[1:]  # Remove a linha do cabeçalho

# Verificar os nomes das colunas
print("Colunas disponíveis:", dados.columns)

# Renomear as colunas conforme necessário
dados.rename(columns={
    'Nome': 'Nome',
    'Plan': 'Plan',
    'data': 'Data',
    'Auto': 'Auto',
    'preço': 'Valor',
    'plano': 'Tipo de Assinatura',
    'Pass': 'EA Play Season Pass',
    '"EA Play Season Pass Price"': 'EA Play Season Pass Price',
    'Minecraft Season Pass': 'Minecraft Season Pass',
    'Minecraft Season Pass Price': 'Minecraft Season Pass Price',
    'Coupon Value': 'Coupon Value',
    'Total': 'Total Value'
}, inplace=True)

# Verificar novamente os nomes das colunas após a renomeação
print("Colunas após renomeação:", dados.columns)

# Converter colunas relevantes para o tipo correto
if 'Data' in dados.columns:
    dados['Data'] = pd.to_datetime(dados['Data'], errors='coerce')  # Converter para datetime
else:
    print("A coluna 'Data' não foi encontrada após a renomeação!")

if 'Valor' in dados.columns:
    dados['Valor'] = pd.to_numeric(dados['Valor'], errors='coerce')  # Converter para numérico
else:
    print("A coluna 'Valor' não foi encontrada!")

# 1. Título do Dashboard
titulo_dashboard = "Vendas de Assinaturas do Xbox Game Pass"

# 2. Resumo de Vendas
if 'Data' in dados.columns and 'Valor' in dados.columns:
    total_vendas = dados['Valor'].sum()
    vendas_mensais = dados.groupby(dados['Data'].dt.to_period('M'))['Valor'].sum()

    # 3. Tipo de Assinatura mais Vendido
    tipo_assinatura_mais_vendido = dados['Tipo de Assinatura'].value_counts(normalize=True) * 100

    # 5. Gráficos
    # Gráfico de Linhas: Faturamento Mensal
    plt.figure(figsize=(10, 6))
    vendas_mensais.plot(kind='line', marker='o')
    plt.title('Faturamento Mensal ao longo do Ano')
    plt.xlabel('Meses')
    plt.ylabel('Faturamento')
    plt.grid()
    plt.savefig('faturamento_mensal.png')
    plt.close()

    # Gráfico de Barras: Comparação de Vendas entre Tipos de Assinaturas
    plt.figure(figsize=(10, 6))
    vendas_por_tipo = dados['Tipo de Assinatura'].value_counts()
    vendas_por_tipo.plot(kind='bar', color='skyblue')
    plt.title('Comparação de Vendas entre Tipos de Assinaturas')
    plt.xlabel('Tipos de Assinaturas')
    plt.ylabel('Número de Vendas')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('vendas_por_tipo.png')
    plt.close()

    # Gráfico de Pizza: Tipo de Assinatura mais Vendido
    plt.figure(figsize=(8, 8))
    tipo_assinatura_mais_vendido.plot(kind='pie', autopct='%1.1f%%', startangle=90)
    plt.title('Tipo de Assinatura mais Vendido')
    plt.ylabel('')
    plt.savefig('tipo_assinatura_mais_vendido.png')
    plt.close()

    # Adicionar gráficos à planilha
    img_faturamento = openpyxl.drawing.image.Image('faturamento_mensal.png')
    img_vendas_tipo = openpyxl.drawing.image.Image('vendas_por_tipo.png')
    img_tipo_assinatura = openpyxl.drawing.image.Image('tipo_assinatura_mais_vendido.png')

    # Adicionando as imagens às células desejadas
    sheet.add_image(img_faturamento, 'E5')
    sheet.add_image(img_vendas_tipo, 'E20')
    sheet.add_image(img_tipo_assinatura, 'E35')

    # Adicionar resumo de vendas
    sheet['A1'] = titulo_dashboard
    sheet['A3'] = f'Total de Vendas de Assinaturas: {total_vendas}'
    sheet['A4'] = 'Vendas Mensais:'
    for i, (mes, valor) in enumerate(vendas_mensais.items()):  # Alterado para 'items()'
        sheet[f'A{i + 5}'] = f'{mes}: {valor}'

    # Salvar a planilha atualizada
    workbook.save(r"\\caminho=dashboard_vendas.xlsx")
else:
    print("Dados insuficientes para gerar o dashboard.")
