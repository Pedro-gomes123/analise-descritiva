import pandas as pd

dados = pd.read_excel('Vendas.xlsx')

if 'Data' in dados.columns:
    dados['Data'] = pd.to_datetime(dados['Data'])
    dados['Mes'] = dados['Data'].dt.to_period('M').astype(str)
else:
    raise Exception('Coluna "Data" não encontrada no arquivo Vendas.xlsx')

dados['Preço Unitário'] = pd.to_numeric(dados['Preço Unitário'])
dados['Quantidade'] = pd.to_numeric(dados['Quantidade'])

dados['Desconto (%)'] = dados['Desconto (%)'].astype(str).str.replace('%', '', regex=True).astype(float) / 100

dados['Valor Bruto'] = dados['Preço Unitário'] * dados['Quantidade']
dados['Valor Desconto'] = dados['Valor Bruto'] * dados['Desconto (%)']
dados['Faturamento Liquido'] = dados['Valor Bruto'] - dados['Valor Desconto']

faturamento_bruto_loja = dados.groupby(['Loja', 'Mes'])['Valor Bruto'].sum()
faturamento_liquido_loja = dados.groupby(['Loja', 'Mes'])['Faturamento Liquido'].sum()
total_vendas_loja = dados.groupby(['Loja', 'Mes']).size()
ticket_loja = faturamento_liquido_loja / total_vendas_loja
Produto_loja = dados.groupby(['Loja', 'Mes', 'Produto'])['Quantidade'].sum()
Produto_loja_ordenado = Produto_loja.sort_values(ascending=False)
top_produto_por_loja = Produto_loja_ordenado.groupby(['Loja', 'Mes']).head(1)

df_bruto = faturamento_bruto_loja.rename('Faturamento Bruto').reset_index()
df_liquido = faturamento_liquido_loja.rename('Faturamento Liquido').reset_index()
df_vendas = total_vendas_loja.rename('Qtd Vendas').reset_index()
df_ticket = ticket_loja.rename('Ticket Medio').reset_index()
df_top = top_produto_por_loja.rename('Produto Mais Vendido').reset_index()

df_loja_final = df_bruto.merge(df_liquido, on=['Loja', 'Mes']) \
                        .merge(df_vendas, on=['Loja', 'Mes']) \
                        .merge(df_ticket, on=['Loja', 'Mes']) \
                        .merge(df_top, on=['Loja', 'Mes'])

df_loja_final.to_excel('resumo_lojas.xlsx', index=False)

categoria_bruto = dados.groupby(['Categoria', 'Mes'])['Valor Bruto'].sum()
categoria_liquido = dados.groupby(['Categoria', 'Mes'])['Faturamento Liquido'].sum()
Categoria_vendas = dados.groupby(['Categoria', 'Mes'])['Quantidade'].sum()
ticket_categoria = categoria_liquido / Categoria_vendas
desconto_categoria = dados.groupby(['Categoria', 'Mes'])['Valor Desconto'].mean()

df_cat_bruto = categoria_bruto.rename('Faturamento Bruto').reset_index()
df_cat_liquido = categoria_liquido.rename('Faturamento Liquido').reset_index()
df_cat_vendas = Categoria_vendas.rename('Qtd Vendas').reset_index()
df_cat_ticket = ticket_categoria.rename('Ticket Medio').reset_index()
df_cat_desc = desconto_categoria.rename('Desconto Medio').reset_index()

df_categoria_final = df_cat_bruto.merge(df_cat_liquido, on=['Categoria', 'Mes']) \
                                 .merge(df_cat_vendas, on=['Categoria', 'Mes']) \
                                 .merge(df_cat_ticket, on=['Categoria', 'Mes']) \
                                 .merge(df_cat_desc, on=['Categoria', 'Mes'])

df_categoria_final.to_excel('resumo_categorias.xlsx', index=False)
