import streamlit as st
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
#from io import BytesIO

# Configurações de email
EMAIL = 'daniel.feitosa.mis@gmail.com'
SENHA = 'eefr tlum huof spwh'
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
DESTINATARIO = 'daniel.feitosa.mis@gmail.com'

# Tela de login
def login():
    st.title("Login")
    usuario = st.selectbox("Usuário", ["Escritório A", "Escritório B", "Escritório C"])
    email = st.text_input("Email")
    if st.button("Entrar"):
        if usuario and email:
            st.session_state['logged_in'] = True
            st.session_state['usuario'] = usuario
            st.session_state['user_email'] = email
        else:
            st.error("Por favor, selecione um usuário e digite o email.")

# Função para validar os campos
def validar_campos(numero_nota, empresa_pagadora, quantidade_itens, items):
    erros = []

    # Verificação dos campos principais
    if not numero_nota:
        erros.append("Número da Nota")
    if not empresa_pagadora:
        erros.append("Empresa Pagadora")
    if quantidade_itens < 1:
        erros.append("Quantidade de Itens")

    # Verificação dos campos dos itens
    for i, item in enumerate(items):
        if not item["Código da Causa"]:
            erros.append(f"Código da Causa do Item {i+1}")
        if not item["Número do Processo"]:
            erros.append(f"Número do Processo do Item {i+1}")
        if not item["Tipo de Despesa"]:
            erros.append(f"Tipo de Despesa do Item {i+1}")
        if item["Valor do Item"] <= 0:
            erros.append(f"Valor do Item {i+1}")

    return erros

# Tela de cadastro de Nota Fiscal
def cadastro_nf():
    # Saudação personalizada
    st.title("Cadastro de Nota Fiscal")
    st.write(f"Olá, {st.session_state['usuario']}")

    numero_nota = st.text_input("Número da Nota")
    empresa_pagadora = st.selectbox("Empresa Pagadora", ["Banco xpto", "Banco ABC", "Financeira X"])
    quantidade_itens = st.number_input("Quantidade de Itens da Nota", min_value=1, step=1)

    # Configurando cabeçalhos para a tabela de itens
    st.subheader("Detalhes dos Itens")

    # Exibe o cabeçalho da tabela apenas uma vez
    colunas = st.columns([4, 4, 5, 3, 5, 1])
    with colunas[0]: st.write("Código da Causa")
    with colunas[1]: st.write("Número do Processo")
    with colunas[2]: st.write("Tipo de Despesa")
    with colunas[3]: st.write("Valor do Item")
    with colunas[4]: st.write("Observação")

    # Lista para armazenar os itens
    items = []
    valor_total = 0

    # Exibindo os itens em uma tabela dinâmica
    for i in range(int(quantidade_itens)):
        colunas = st.columns([4, 4, 5, 3, 5, 1])
        with colunas[0]:
            codigo_causa = st.text_input("", key=f'causa_{i}')
        with colunas[1]:
            numero_processo = st.text_input("", key=f'processo_{i}')
        with colunas[2]:
            tipo_despesa = st.selectbox(
                "", 
                ["Pro-Labore", "Exito", "Recuperação Judicial"], 
                key=f'tipo_{i}'
            )
        with colunas[3]:
            valor_item = st.number_input("", min_value=0.0, step=0.01, key=f'valor_{i}')
        with colunas[4]:
            observacao = st.text_input("", key=f'observacao_{i}')
        
        # Somando o valor dos itens para o total
        valor_total += valor_item
        items.append({
            "Código da Causa": codigo_causa,
            "Número do Processo": numero_processo,
            "Tipo de Despesa": tipo_despesa,
            "Valor do Item": valor_item,
            "Observação": observacao
        })

    # Campo de valor total da nota (desabilitado)
    st.text_input("Valor Total da Nota", value=valor_total, disabled=True)

    # Campo para anexar múltiplos arquivos
    arquivos_anexos = st.file_uploader("Anexar arquivos", accept_multiple_files=True)

    # Função para enviar email com anexo
    if st.button("Enviar Email"):
        # Validar campos antes de enviar
        erros = validar_campos(numero_nota, empresa_pagadora, quantidade_itens, items)
        if erros:
            st.error("Erro: Os seguintes campos estão vazios ou inválidos: " + ", ".join(erros))
        else:
            enviar_email(numero_nota, empresa_pagadora, valor_total, items, arquivos_anexos, st.session_state['usuario'], st.session_state['user_email'])
            st.success("Email enviado com sucesso!")
            limpar_campos()

# Função para enviar e-mail com os dados e gerar o arquivo .xlsx
def enviar_email(numero_nota, empresa_pagadora, valor_total, items, arquivos_anexos, usuario, user_email):
    # Gerando o arquivo .xlsx
    df = pd.DataFrame(items)
    df['Número da Nota'] = numero_nota
    df['Empresa Pagadora'] = empresa_pagadora
    df['Valor Total da Nota'] = valor_total
    df['Usuário'] = usuario  # Adiciona o usuário como uma coluna no Excel

    # Salvar o Excel em um buffer de memória
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    excel_buffer.seek(0)

    # Nome do arquivo com o formato desejado
    nome_arquivo = f"notafiscal - {numero_nota}.xlsx"

    # Configuração do e-mail
    msg = MIMEMultipart()
    msg['From'] = EMAIL
    msg['To'] = DESTINATARIO
    msg['Cc'] = user_email  # E-mail do usuário será adicionado como cópia
    msg['Subject'] = f"Nota Fiscal - {numero_nota}"
    
    body = f"""
    Número da Nota: {numero_nota}
    Empresa Pagadora: {empresa_pagadora}
    Valor Total: {valor_total}
    Usuário: {usuario}
    """
    msg.attach(MIMEText(body, 'plain'))
    
    # Anexar o arquivo .xlsx
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(excel_buffer.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{nome_arquivo}"')
    msg.attach(part)
    
    # Anexar os arquivos selecionados, se houver
    for arquivo in arquivos_anexos:
        part_anexo = MIMEBase('application', 'octet-stream')
        part_anexo.set_payload(arquivo.read())
        encoders.encode_base64(part_anexo)
        part_anexo.add_header('Content-Disposition', f'attachment; filename="{arquivo.name}"')
        msg.attach(part_anexo)

    # Enviar o e-mail
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL, SENHA)
            server.sendmail(EMAIL, [DESTINATARIO, user_email], msg.as_string())
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {e}")

# Função para limpar os campos do formulário
def limpar_campos():
    #st.session_state.clear() ----anterior bkp
    st.session_state["numero_nota"] = ""
    st.session_state["quantidade_itens"] = 1
    #for key in list(st.session_state.keys()):
     #   if key.startswith("causa_") or key.startswith("processo_") or key.stouartswith("tipo_") or key.startswith("valor_") or key.startswith("observacao_"):
      #      st.session_state[key] = ""

# Verifica se o usuário está logado
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

if st.session_state['logged_in']:
    cadastro_nf()
else:
    login()
