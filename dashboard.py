from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import pandas as pd
import io
import streamlit as st
import plotly.express as px

# Detalhes da autenticação e URL do SharePoint
url_shrpt = "https://zigpay.sharepoint.com/sites/Plan.Comercial-2023/"  # Substitua pelo URL correto do seu SharePoint
username_shrpt = "arthur.ramos@zig.fun"
password_shrpt = "Project@64"
file_url_shrpt = "/sites/Plan.Comercial-2023/Shared Documents/Projetos/Novo Relatório Cesta de Produtos/Relatório Cesta de Produtos.xlsx"

# Autenticação
ctx_auth = AuthenticationContext(url_shrpt)
if ctx_auth.acquire_token_for_user(username_shrpt, password_shrpt):
    ctx = ClientContext(url_shrpt, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("Autenticação bem-sucedida!")
else:
    print("Erro de autenticação:", ctx_auth.get_last_error())
    exit()

# Acessar o arquivo no SharePoint
response = File.open_binary(ctx, file_url_shrpt)

# Carregar o conteúdo do arquivo em um objeto BytesIO
bytes_file = io.BytesIO(response.content)

# Ler a aba "GeralEventos" da planilha
df = pd.read_excel(bytes_file, sheet_name="GeralEventos")

# Redefinir o índice para remover a coluna de índice original
df.reset_index(drop=True, inplace=True)

# Excluir a linha onde o "Comercial" é "Total Geral"
df = df[df['Comercial'] != 'Total Geral']

# Certificar-se de que a coluna '%' é numérica e multiplicar por 100 para converter em porcentagem
df['%'] = pd.to_numeric(df['%'], errors='coerce') * 100

# Ordenar o dataframe pela coluna 'Pontuação Total' em ordem decrescente
df = df.sort_values(by='Pontuação Total', ascending=False)

st.title('Relatório de Produtos:')
st.subheader('Pontuação Geral Eventos')

# Criar um filtro para selecionar o comercial
comercial_selecionado = st.selectbox('Selecione o Comercial', df['Comercial'].unique())

# Filtrar o dataframe para o comercial selecionado (somente para o gráfico)
df_filtrado = df[df['Comercial'] == comercial_selecionado]

# Criar o gráfico usando Plotly, com rótulos de valor nas barras
fig = px.bar(df_filtrado, x='Comercial', y=['Pontuação Total', 'Meta', '%'],
             title=f'Dados do Comercial: {comercial_selecionado}',
             labels={'value': 'Valores', 'variable': 'Indicadores'},
             barmode='group',
             text_auto=True)

# Atualizar o trace apenas para a coluna '%', formatando-a como porcentagem
fig.for_each_trace(lambda t: t.update(texttemplate='%{y:.2f}%', textposition='outside') if t.name == '%' else t)

# Exibir o gráfico no Streamlit
st.plotly_chart(fig)

# Formatar a coluna 'Meta' como porcentagem no DataFrame
df['%'] = df['%'].apply(lambda x: "{:.2f}%".format(x))

# Exibir o dataframe completo abaixo do gráfico
st.title('Dados completos')
st.dataframe(df)
