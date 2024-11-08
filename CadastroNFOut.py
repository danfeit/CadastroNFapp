import streamlit as st
import os
import pandas as pd
from io import BytesIO
import platform

if platform.system() == "Windows":
    import win32com.client
else:
    print("win32com.client is not available on non-Windows platforms.")
# import win32com.client as win32
import pythoncom  # Biblioteca para controle do ambiente COM
from tempfile import NamedTemporaryFile  # Para criar arquivos temporários


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

    # Campos principais
    numero_nota = st.text_input("Número da Nota", key="numero_nota")
    empresa_pagadora = st.selectbox("Empresa Pagadora", ["Banco xpto", "Banco ABC", "Financeira X"])
    quantidade_itens = st.number_input("Quantidade de Itens da Nota", min_value=1, step=1, key="quantidade_itens")

    # Configurando cabeçalhos para a tabela de itens
    st.subheader("Detalhes dos Itens")
    colunas = st.columns([5, 5, 6, 3, 6, 1])
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
        colunas = st.columns([5, 5, 6, 3, 6, 1])
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
            #limpar_campos()

# Função para enviar e-mail com os dados e gerar o arquivo .xlsx
def enviar_email(numero_nota, empresa_pagadora, valor_total, items, arquivos_anexos, usuario, user_email):
    # Gerando o arquivo .xlsx
    df = pd.DataFrame(items)
    df['Número da Nota'] = numero_nota
    df['Empresa Pagadora'] = empresa_pagadora
    df['Valor Total da Nota'] = valor_total
    df['Usuário'] = usuario  # Adiciona o usuário como uma coluna no Excel

    # Salvar o Excel em um arquivo temporário
    nome_arquivo = f"{usuario}_notafiscal_{numero_nota}.xlsx"
    caminho_arquivo = os.path.join(os.getcwd(), nome_arquivo)
    df.to_excel(caminho_arquivo, index=False)

    # Inicializando o COM e o Outlook
    pythoncom.CoInitialize()  # Inicializa o COM
    outlook = win32.Dispatch("Outlook.Application")
    email = outlook.CreateItem(0)
    email.To = "daniel.feitosa.mis@gmail.com"
    email.CC = user_email
    email.Subject = f"Nota Fiscal - {numero_nota}"
    email.Body = f"""
    Número da Nota: {numero_nota}
    Empresa Pagadora: {empresa_pagadora}
    Valor Total: {valor_total}
    Usuário: {usuario}
    """

    # Anexando o arquivo .xlsx pelo caminho temporário
    email.Attachments.Add(caminho_arquivo)

    # Anexar outros arquivos, se houver
    temp_files = []  # Lista para armazenar os arquivos temporários
    for arquivo in arquivos_anexos:
        # Usar o nome original do arquivo para salvar o temporário
        arquivo_nome = arquivo.name  # Nome original do arquivo
        temp_file_path = os.path.join(os.getcwd(), arquivo_nome)

        # Salvar o conteúdo do arquivo UploadedFile no caminho temporário
        with open(temp_file_path, "wb") as temp_file:
            temp_file.write(arquivo.getvalue())

        # Adicionar o caminho à lista de arquivos temporários e anexar ao e-mail
        temp_files.append(temp_file_path)
        email.Attachments.Add(temp_file_path)

    # Enviar o e-mail
    email.Send()
    pythoncom.CoUninitialize()  # Finaliza o COM

    # Remover o arquivo temporário do Excel e anexos após o envio
    os.remove(caminho_arquivo)
    for temp_file_path in temp_files:
        os.remove(temp_file_path)


# Função para limpar apenas os campos específicos do formulário
#def limpar_campos():
    #st.session_state["numero_nota"] = ""
    #st.session_state["quantidade_itens"] = 1
    #for key in list(st.session_state.keys()):
        #if key.startswith("causa_") or key.startswith("processo_") or key.startswith("tipo_") or key.startswith("valor_") or key.startswith("observacao_"):
            #st.session_state[key] = ""

# Verifica se o usuário está logado
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

if st.session_state['logged_in']:
    cadastro_nf()
else:
    login()
