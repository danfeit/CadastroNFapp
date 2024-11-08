import streamlit as st
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from io import BytesIO

# Configurações de email (recomendado usar st.secrets para produção)
EMAIL = st.secrets["EMAIL"]
SENHA = st.secrets["SENHA"]
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
DESTINATARIO = st.secrets["DESTINATARIO"]

# Tela de login
def login():
    st.title("Login")
    with st.form("login_form"):
        usuario = st.selectbox("Usuário", ["Escritório A", "Escritório B", "Escritório C"])
        email = st.text_input("Email")
        submit_button = st.form_submit_button("Entrar")
        
        if submit_button:
            if usuario and email:
                st.session_state['logged_in'] = True
                st.session_state['usuario'] = usuario
                st.session_state['user_email'] = email
                st.rerun()
            else:
                st.error("Por favor, selecione um usuário e digite o email.")

# Função para validar os campos
def validar_campos(numero_nota, empresa_pagadora, quantidade_itens, items):
    erros = []
    
    if not numero_nota:
        erros.append("Número da Nota")
    if not empresa_pagadora:
        erros.append("Empresa Pagadora")
    if quantidade_itens < 1:
        erros.append("Quantidade de Itens")
        
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

# Função para enviar e-mail com os dados e gerar o arquivo .xlsx
def enviar_email(numero_nota, empresa_pagadora, valor_total, items, arquivos_anexos, usuario, user_email):
    try:
        # Gerando o arquivo .xlsx
        df = pd.DataFrame(items)
        df['Número da Nota'] = numero_nota
        df['Empresa Pagadora'] = empresa_pagadora
        df['Valor Total da Nota'] = valor_total
        df['Usuário'] = usuario

        # Salvar o Excel em um buffer de memória
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        excel_buffer.seek(0)

        # Configuração do e-mail
        msg = MIMEMultipart()
        msg['From'] = EMAIL
        msg['To'] = DESTINATARIO
        msg['Cc'] = user_email
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
        part.set_payload(excel_buffer.getvalue())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="notafiscal - {numero_nota}.xlsx"')
        msg.attach(part)
        
        # Anexar os arquivos selecionados
        for arquivo in arquivos_anexos:
            part_anexo = MIMEBase('application', 'octet-stream')
            part_anexo.set_payload(arquivo.getvalue())
            encoders.encode_base64(part_anexo)
            part_anexo.add_header('Content-Disposition', f'attachment; filename="{arquivo.name}"')
            msg.attach(part_anexo)

        # Enviar o e-mail
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL, SENHA)
            todos_destinatarios = [DESTINATARIO, user_email]
            server.sendmail(EMAIL, todos_destinatarios, msg.as_string())
            
        return True
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {str(e)}")
        return False

# Tela de cadastro de Nota Fiscal
def cadastro_nf():
    st.title("Cadastro de Nota Fiscal")
    st.write(f"Olá, {st.session_state['usuario']}")
    
    with st.form("nota_fiscal_form"):
        numero_nota = st.text_input("Número da Nota")
        empresa_pagadora = st.selectbox("Empresa Pagadora", ["Banco xpto", "Banco ABC", "Financeira X"])
        quantidade_itens = st.number_input("Quantidade de Itens da Nota", min_value=1, step=1)

        st.subheader("Detalhes dos Itens")
        
        items = []
        valor_total = 0

        for i in range(int(quantidade_itens)):
            st.markdown(f"**Item {i+1}**")
            cols = st.columns([2, 2, 2, 1, 2])
            
            codigo_causa = cols[0].text_input("Código da Causa", key=f'causa_{i}')
            numero_processo = cols[1].text_input("Número do Processo", key=f'processo_{i}')
            tipo_despesa = cols[2].selectbox(
                "Tipo de Despesa",
                ["Pro-Labore", "Exito", "Recuperação Judicial"],
                key=f'tipo_{i}'
            )
            valor_item = cols[3].number_input("Valor", min_value=0.0, step=0.01, key=f'valor_{i}')
            observacao = cols[4].text_input("Observação", key=f'observacao_{i}')
            
            valor_total += valor_item
            items.append({
                "Código da Causa": codigo_causa,
                "Número do Processo": numero_processo,
                "Tipo de Despesa": tipo_despesa,
                "Valor do Item": valor_item,
                "Observação": observacao
            })

        st.text_input("Valor Total da Nota", value=f"R$ {valor_total:.2f}", disabled=True)
        arquivos_anexos = st.file_uploader("Anexar arquivos", accept_multiple_files=True)
        
        submitted = st.form_submit_button("Enviar")
        
        if submitted:
            erros = validar_campos(numero_nota, empresa_pagadora, quantidade_itens, items)
            if erros:
                st.error("Erro: Os seguintes campos estão vazios ou inválidos: " + ", ".join(erros))
            else:
                if enviar_email(numero_nota, empresa_pagadora, valor_total, items, arquivos_anexos, 
                              st.session_state['usuario'], st.session_state['user_email']):
                    st.success("Email enviado com sucesso!")
                    st.rerun()  # Limpa o formulário após envio bem-sucedido

# Inicialização do app
if __name__ == "__main__":
    st.set_page_config(page_title="Sistema de Notas Fiscais", layout="wide")
    
    if 'logged_in' not in st.session_state:
        st.session_state['logged_in'] = False

    if st.session_state['logged_in']:
        cadastro_nf()
    else:
        login()
