import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import time
import re
import os
from PIL import Image
import io
import json
import base64
from datetime import datetime
import streamlit.components.v1 as components

# Configuração da página
st.set_page_config(
    page_title="Email Sender Pro",
    page_icon="✉️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado
st.markdown("""
    <style>
    .stTextInput > label {
        font-size: 20px;
        color: #31333F;
        font-weight: 500;
    }
    .stTextArea > label {
        font-size: 20px;
        color: #31333F;
        font-weight: 500;
    }
    .main {
        background-color: #F5F7F9;
        border-radius: 20px;
        padding: 20px;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        padding: 15px 30px;
        border-radius: 10px;
        border: none;
        font-size: 16px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #45a049;
        transform: translateY(-2px);
    }
    .template-card {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        margin: 10px 0;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .editor-container {
        background-color: white;
        border-radius: 10px;
        padding: 10px;
        margin: 10px 0;
    }
    </style>
    """, unsafe_allow_html=True)

# Inicialização de estados da sessão
if 'signatures' not in st.session_state:
    st.session_state.signatures = {}

if 'email_history' not in st.session_state:
    st.session_state.email_history = []

if 'scheduled_emails' not in st.session_state:
    st.session_state.scheduled_emails = []

if 'templates' not in st.session_state:
    st.session_state.templates = {
        'Template Formal': """
Prezado(a) [Nome],

Espero que esta mensagem o(a) encontre bem.

[Seu conteúdo aqui]

Atenciosamente,
[Seu nome]
        """,
        'Template Informal': """
Olá [Nome]!

[Seu conteúdo aqui]

Abraços,
[Seu nome]
        """,
        'Template Newsletter': """
Olá!

Confira as novidades desta semana:

• [Item 1]
• [Item 2]
• [Item 3]

Para mais informações, entre em contato.

Atenciosamente,
[Seu nome]
        """
    }

def validate_email(email):
    """Valida o formato do email."""
    if not email:
        return False
    email_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return bool(re.match(email_pattern, email))

def parse_recipients(recipients_text):
    """
    Analisa o texto de destinatários no formato 'Nome,Email' ou apenas 'Email'
    Retorna uma lista de tuplas (nome, email)
    """
    recipients_list = []
    if not recipients_text:
        return recipients_list
    
    lines = recipients_text.strip().split('\n')
    for line in lines:
        line = line.strip()
        if ',' in line:  # Formato Nome,Email
            name, email = line.split(',', 1)
            name = name.strip()
            email = email.strip()
        else:  # Apenas email
            email = line
            name = "Destinatário"  # Nome padrão
            
        if validate_email(email):
            recipients_list.append((name, email))
    return recipients_list

def replace_placeholders(message, name):
    """
    Substitui placeholders no texto da mensagem
    """
    replacements = {
        '[Nome]': name,
        '[nome]': name,
        '[NOME]': name,
        '{nome}': name,
        '{Nome}': name,
        '{NOME}': name
    }
    
    for placeholder, value in replacements.items():
        message = message.replace(placeholder, value)
    
    return message

def save_signature(name, image_file):
    """Salva uma nova assinatura."""
    if image_file:
        st.session_state.signatures[name] = image_file
        return True
    return False

def schedule_email(email_data):
    """Agenda um novo email."""
    st.session_state.scheduled_emails.append(email_data)

def add_to_history(recipient, subject, status, timestamp):
    """Adiciona um email ao histórico."""
    st.session_state.email_history.append({
        'timestamp': timestamp,
        'recipient': recipient,
        'subject': subject,
        'status': status
    })

def get_file_type(filename):
    """Retorna o tipo do arquivo baseado na extensão."""
    ext = filename.split('.')[-1].lower()
    mime_types = {
        'pdf': 'application/pdf',
        'doc': 'application/msword',
        'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'xls': 'application/vnd.ms-excel',
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'ppt': 'application/vnd.ms-powerpoint',
        'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'zip': 'application/zip',
        'rar': 'application/x-rar-compressed',
        'txt': 'text/plain',
    }
    return mime_types.get(ext, 'application/octet-stream')

def send_email(sender_email, sender_password, recipient, subject, message, attachments=None, signature_image=None):
    """Envia um único email com suporte a múltiplos anexos."""
    try:
        smtp_server = "smtp.titan.email"
        port = 465

        msg = MIMEMultipart('related')
        msg['From'] = sender_email
        msg['To'] = recipient
        msg['Subject'] = subject

        html_message = f"""
        <html>
            <body>
                {message}
                {f'<img src="cid:signature" width="200"/>' if signature_image else ''}
            </body>
        </html>
        """
        
        msg_alternative = MIMEMultipart('alternative')
        msg.attach(msg_alternative)
        
        html_part = MIMEText(html_message, 'html')
        msg_alternative.attach(html_part)

        if signature_image:
            img = MIMEImage(signature_image.read())
            img.add_header('Content-ID', '<signature>')
            img.add_header('Content-Disposition', 'inline')
            msg.attach(img)

        if attachments:
            for attachment in attachments:
                if attachment is not None:
                    file_type = get_file_type(attachment.name)
                    attachment_part = MIMEApplication(
                        attachment.read(),
                        _subtype=file_type.split('/')[-1]
                    )
                    attachment_part.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=attachment.name
                    )
                    msg.attach(attachment_part)

        with smtplib.SMTP_SSL(smtp_server, port) as server:
            server.login(sender_email, sender_password)
            server.send_message(msg)
        return True, "Email enviado com sucesso!"
    except Exception as e:
        return False, f"Erro ao enviar email: {str(e)}"

# Interface principal
st.title("✉️ Sistema de Envio de Emails Pro")

# Barra lateral com todas as funcionalidades
with st.sidebar:
    tab1, tab2, tab3 = st.tabs(["📝 Templates", "✍️ Assinaturas", "📅 Agendados"])

    with tab1:
        st.header("📝 Templates")
        selected_template = st.selectbox(
            "Escolha um template",
            ["Sem template"] + list(st.session_state.templates.keys())
        )
        
        with st.expander("➕ Adicionar Novo Template"):
            new_template_name = st.text_input("Nome do Template")
            new_template_content = st.text_area("Conteúdo do Template")
            if st.button("Salvar Template"):
                if new_template_name and new_template_content:
                    st.session_state.templates[new_template_name] = new_template_content
                    st.success("Template salvo com sucesso!")
                    st.experimental_rerun()

    with tab2:
        st.header("✍️ Assinaturas")
        new_signature_name = st.text_input("Nome da Assinatura")
        new_signature_file = st.file_uploader(
            "Imagem da Assinatura",
            type=['png', 'jpg', 'jpeg'],
            key="new_signature"
        )
        if st.button("Salvar Assinatura"):
            if new_signature_name and new_signature_file:
                if save_signature(new_signature_name, new_signature_file):
                    st.success("Assinatura salva com sucesso!")
                    st.experimental_rerun()

        if st.session_state.signatures:
            st.subheader("Assinaturas Salvas")
            for name in st.session_state.signatures:
                with st.expander(name):
                    st.image(st.session_state.signatures[name], width=200)
                    if st.button("Excluir", key=f"del_{name}"):
                        del st.session_state.signatures[name]
                        st.experimental_rerun()

    with tab3:
        st.header("📅 Emails Agendados")
        if st.session_state.scheduled_emails:
            for idx, email in enumerate(st.session_state.scheduled_emails):
                with st.expander(f"Email {idx + 1} - {email['subject']}"):
                    st.write(f"Para: {email['recipient']}")
                    st.write(f"Data: {email['schedule_time']}")
                    if st.button("Cancelar Agendamento", key=f"cancel_{idx}"):
                        st.session_state.scheduled_emails.pop(idx)
                        st.experimental_rerun()
        else:
            st.info("Não há emails agendados no momento.")

# Seção de configuração
with st.expander("⚙️ Configurações de Email", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        sender_email = st.text_input("Email do Remetente (Titan)")
    with col2:
        sender_password = st.text_input("Senha", type="password")

# Seção de composição do email
st.subheader("📝 Composição do Email")

subject = st.text_input("Assunto do Email")

# Editor de texto rico usando HTML
st.write("Mensagem do Email")
message_content = st.session_state.templates.get(selected_template, "") if selected_template != "Sem template" else ""

# Componente de editor rico
components.html(
    f"""
    <div class="editor-container">
        <link href="https://cdn.quilljs.com/1.3.6/quill.snow.css" rel="stylesheet">
        <script src="https://cdn.quilljs.com/1.3.6/quill.js"></script>
        <div id="editor" style="height: 300px;">{message_content}</div>
        <script>
            var quill = new Quill('#editor', {{
                theme: 'snow',
                modules: {{
                    toolbar: [
                        ['bold', 'italic', 'underline', 'strike'],
                        ['blockquote', 'code-block'],
                        [{{ 'header': 1 }}, {{ 'header': 2 }}],
                        [{{ 'list': 'ordered' }}, {{ 'list': 'bullet' }}],
                        [{{ 'color': [] }}, {{ 'background': [] }}],
                        ['clean']
                    ]
                }}
            }});
        </script>
    </div>
    """,
    height=400,
)

# Seção de destinatários
st.subheader("📧 Destinatários")
recipient_method = st.radio(
    "Escolha o método de entrada dos destinatários",
    ["Digitar manualmente", "Importar CSV"]
)

recipients_data = []

if recipient_method == "Digitar manualmente":
    st.markdown("""
    **Formatos aceitos:**
    ```
    # Com nome:
    Nome,Email
    João Silva,joao@email.com
    Maria Santos,maria@email.com
    
    # Apenas email:
    joao@email.com
    maria@email.com
    ```
    """)
    
    recipients_text = st.text_area(
        "Lista de Destinatários (Um por linha, formato: Nome,Email ou apenas Email)"
    )
    if recipients_text:
        recipients_data = parse_recipients(recipients_text)
        if recipients_data:
            st.write("✅ Destinatários identificados:")
            for name, email in recipients_data:
                st.write(f"- {name}: {email}")
        else:
            st.warning("Nenhum destinatário válido encontrado. Verifique o formato.")

else:
    st.markdown("""
    **Formatos aceitos para o arquivo CSV:**
    1. Com nome e email:
       - Coluna 'Nome': Nome do destinatário
       - Coluna 'Email': Email do destinatário
       
    2. Apenas email:
       - Coluna 'Email': Email do destinatário
    """)
    
    uploaded_file = st.file_uploader(
        "Upload do arquivo CSV",
        type=['csv']
    )
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file)
            if 'Email' in df.columns:
                for _, row in df.iterrows():
                    email = row['Email'].strip()
                    if validate_email(email):
                        name = row['Nome'].strip() if 'Nome' in df.columns else "Destinatário"
                        recipients_data.append((name, email))
                if recipients_data:
                    st.write("✅ Destinatários identificados do CSV:")
                    for name, email in recipients_data:
                        st.write(f"- {name}: {email}")
                else:
                    st.warning("Nenhum email válido encontrado no arquivo.")
            else:
                st.error("O arquivo CSV deve conter pelo menos a coluna 'Email'")
        except Exception as e:
            st.error(f"Erro ao ler o arquivo CSV: {str(e)}")

# Seção de anexos e assinatura
st.subheader("📎 Anexos e Assinatura")
col1, col2 = st.columns(2)

with col1:
    st.write("Anexos (Arraste ou selecione múltiplos arquivos)")
    allowed_types = ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'zip', 'rar']
    attachments = st.file_uploader(
        "Escolha os arquivos para anexar",
        type=allowed_types,
        accept_multiple_files=True
    )
    if attachments:
        for attachment in attachments:
            st.write(f"📎 {attachment.name}")

with col2:
    st.write("Assinatura")
    signature_file = None
    if st.session_state.signatures:
        selected_signature = st.selectbox(
            "Selecionar Assinatura",
            ["Sem assinatura"] + list(st.session_state.signatures.keys())
        )
        if selected_signature != "Sem assinatura":
            signature_file = st.session_state.signatures[selected_signature]
            st.image(signature_file, width=200)
    else:
        signature_file = st.file_uploader(
            "Upload da imagem da assinatura",
            type=['png', 'jpg', 'jpeg']
        )
        if signature_file:
            st.image(signature_file, width=200)

# Configurações de envio
interval = st.number_input("Intervalo entre envios (segundos)", min_value=1, value=5)

# Opção de agendamento
with st.expander("🕐 Agendar Envio", expanded=False):
    schedule_email_enabled = st.checkbox("Habilitar agendamento")
    if schedule_email_enabled:
        col1, col2 = st.columns(2)
        with col1:
            schedule_date = st.date_input("Data de envio")
        with col2:
            schedule_time = st.time_input("Hora de envio")
        schedule_datetime = datetime.combine(schedule_date, schedule_time)

# Botões de envio e histórico
col1, col2 = st.columns([2, 1])
with col1:
    if st.button("Enviar/Agendar Emails"):
        if not sender_email or not sender_password or not subject or not recipients_data:
            st.error("Por favor, preencha todos os campos obrigatórios!")
        else:
            if schedule_email_enabled:
                for name, email in recipients_data:
                    personalized_message = replace_placeholders(message_content, name)
                    email_data = {
                        'recipient': email,
                        'recipient_name': name,
                        'subject': subject,
                        'message': personalized_message,
                        'attachments': attachments,
                        'signature': signature_file,
                        'schedule_time': schedule_datetime
                    }
                    schedule_email(email_data)
                st.success(f"✅ {len(recipients_data)} emails agendados para {schedule_datetime}")
                st.experimental_rerun()
            
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                st.subheader("📋 Log de Envios")
                log_container = st.container()
                
                for idx, (name, email) in enumerate(recipients_data):
                    progress = (idx + 1) / len(recipients_data)
                    progress_bar.progress(progress)
                    status_text.text(f"Processando {idx + 1} de {len(recipients_data)} emails...")
                    
                    # Personaliza a mensagem para cada destinatário
                    personalized_message = replace_placeholders(message_content, name)
                    
                    # Resetar posição dos arquivos
                    if signature_file:
                        signature_file.seek(0)
                    if attachments:
                        for attachment in attachments:
                            attachment.seek(0)
                    
                    success, message_status = send_email(
                        sender_email,
                        sender_password,
                        email,
                        subject,
                        personalized_message,
                        attachments,
                        signature_file
                    )
                    
                    # Adiciona ao histórico
                    add_to_history(
                        f"{name} <{email}>",
                        subject,
                        "Sucesso" if success else "Falha",
                        datetime.now()
                    )
                    
                    if success:
                        log_container.success(f"✅ {name} ({email}): Email enviado com sucesso!")
                    else:
                        log_container.error(f"❌ {name} ({email}): {message_status}")
                    
                    if idx < len(recipients_data) - 1:
                        time.sleep(interval)
                
                status_text.text("✨ Processo finalizado!")
                progress_bar.progress(1.0)

with col2:
    if st.button("Ver Histórico"):
        st.session_state.show_history = True

if st.session_state.get('show_history', False):
    with st.expander("📋 Histórico de Envios", expanded=True):
        if st.session_state.email_history:
            df = pd.DataFrame(st.session_state.email_history)
            st.dataframe(
                df.sort_values('timestamp', ascending=False),
                column_config={
                    "timestamp": st.column_config.DatetimeColumn("Data/Hora"),
                    "recipient": "Destinatário",
                    "subject": "Assunto",
                    "status": "Status"
                }
            )
        else:
            st.info("Nenhum email enviado ainda.")
        if st.button("Fechar Histórico"):
            st.session_state.show_history = False

# Rodapé
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center'>
        <p>Desenvolvido com ❤️ usando Streamlit</p>
    </div>
    """,
    unsafe_allow_html=True
)
