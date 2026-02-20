# -*- coding: utf-8 -*-
import os
import json
import base64
import requests
import webbrowser
import time
import re
import subprocess
import fitz  # PyMuPDF
from datetime import datetime, timedelta
from requests.exceptions import RequestException
from urllib.parse import urlparse, parse_qs


# Configurações de autenticação (estas precisarão ser obtidas do Azure AD)
TENANT_ID = "COLOQUE_SEU_TENANT_ID"
CLIENT_ID = "COLOQUE_SEU_CLIENT_ID"
CLIENT_SECRET = "COLOQUE_SEU_CLIENT_SECRET"

# ID da pasta #NFE PLATINUM
NFE_PLATINUM_FOLDER_ID = "COLOQUE_SEU_FOLDER_ID"


# Configurações Globais
SENDER_EMAIL = "noreply@omie.com.br"
SUBJECT_CONTAINS = "PLATINUM TELEINFORMATICA LTDA - Nota Fiscal Eletrônica - "
BODY_CONTAINS = ["CLARO ", "TELMEX "]
SCOPES = ["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Mail.ReadWrite", "offline_access"]

# Caminhos Relativos (dentro da pasta do projeto)
import sys
if getattr(sys, 'frozen', False):
    # Se estiver rodando como executável (.exe)
    base_path = os.path.dirname(sys.executable)
    # Se o executável estiver dentro da pasta 'dist', subir um nível para a raiz do projeto
    if os.path.basename(base_path).lower() == "dist":
        SCRIPT_DIR = os.path.dirname(base_path)
    else:
        SCRIPT_DIR = base_path
else:
    # Se estiver rodando como script (.py)
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

SAVE_DIRECTORY = os.path.join(SCRIPT_DIR, "Notas-Salvas-xml-email")
CARREGADAS_DIR = os.path.join(SAVE_DIRECTORY, "Carregadas ao portal")

# Controle p/ GUI
STOP_MONITOR = False
LOG_CALLBACK = None
TIMER_CALLBACK = None

def log(message):
    print(message)
    if LOG_CALLBACK:
        LOG_CALLBACK(message)

def update_timer(seconds):
    if TIMER_CALLBACK:
        TIMER_CALLBACK(seconds)

def get_authorization_url(tenant_id, client_id, scopes):
    scope_string = " ".join(scopes)
    redirect_uri = "http://localhost:8080/callback"
    auth_url = (
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/authorize?"
        f"client_id={client_id}&response_type=code&redirect_uri={redirect_uri}&"
        f"scope={scope_string}&response_mode=query"
    )
    return auth_url

def get_access_token_from_code(tenant_id, client_id, authorization_code):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": client_id,
        "client_secret": CLIENT_SECRET,
        "scope": " ".join(SCOPES),
        "code": authorization_code,
        "redirect_uri": "http://localhost:8080/callback",
        "grant_type": "authorization_code"
    }
    try:
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()
        return response.json() # Retorna o json completo p/ pegar o refresh_token tambem
    except RequestException as e:
        log(f"Erro ao obter token: {e}")
        return None


def save_tokens(access_token, refresh_token, expires_in):
    expiration_time = (datetime.now() + timedelta(seconds=expires_in)).isoformat()
    data = {
        "access_token": access_token,
        "refresh_token": refresh_token,
        "expires_at": expiration_time
    }
    with open("token.json", "w") as f:
        json.dump(data, f)

def load_tokens():
    if not os.path.exists("token.json"):
        return None, None, None
    with open("token.json", "r") as f:
        data = json.load(f)
    return data["access_token"], data["refresh_token"], datetime.fromisoformat(data["expires_at"])

def refresh_access_token(tenant_id, client_id, client_secret, refresh_token):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": " ".join(SCOPES)
    }
    try:
        response = requests.post(url, data=data)
        response.raise_for_status()
        tokens = response.json()
        return tokens["access_token"], tokens["refresh_token"], tokens["expires_in"]
    except Exception as e:
        print(f"Erro ao renovar token: {e}")
        return None, None, None


def get_nfe_emails(access_token, folder_id, sender, subject_contains, body_contains_list):
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    
    # Criar filtro para múltiplos BODY_CONTAINS
    body_filters = " or ".join([f"contains(body/content, '{body}')" for body in body_contains_list])
    filter_query = (
        f"from/emailAddress/address eq '{sender}' and "
        f"isRead eq false and "
        f"contains(subject, '{subject_contains}') and "
        f"({body_filters})"
    )
    encoded_filter = requests.utils.quote(filter_query)
    url = f"https://graph.microsoft.com/v1.0/me/mailFolders('{folder_id}')/messages?$filter={encoded_filter}&$expand=attachments"
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()["value"]
    except RequestException as e:
        print(f"Erro ao buscar e-mails: {e}")
        return []

def extract_nf_data(pdf_content):
    # Extrair número NF, pedido e protocolo do PDF
    nf_match = re.search(r"Nº\s*(\d+(?:\.\d+)*)\s*Série", pdf_content, re.IGNORECASE)
    nf_number = nf_match.group(1).replace(".", "") if nf_match else None
    
    pedido_protocolo_match = re.search(r"PEDIDO\s*(\d+)\s*/\s*PROTOCOLO\s*(\d+)", pdf_content, re.IGNORECASE)
    pedido = pedido_protocolo_match.group(1) if pedido_protocolo_match else None
    protocolo = pedido_protocolo_match.group(2) if pedido_protocolo_match else None
    
    return nf_number, pedido, protocolo

def create_or_rename_folder(nf_number, pedido, protocolo, is_cancelled=False):
    prefix = "Cancelada - " if is_cancelled else ""
    folder_name = f"{prefix}Nota Fiscal Eletrônica - {nf_number} {pedido} {protocolo}"
    folder_path = os.path.join(SAVE_DIRECTORY, folder_name)
    
    # Se é cancelamento, procurar pasta existente para renomear
    if is_cancelled:
        original_name = f"Nota Fiscal Eletrônica - {nf_number} {pedido} {protocolo}"
        original_path = os.path.join(SAVE_DIRECTORY, original_name)
        if os.path.exists(original_path):
            os.rename(original_path, folder_path)
            print(f"Pasta renomeada para: {folder_name}")
            return folder_path
    
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def extract_text_fallback_with_pymupdf(file_path):
    """Tenta extrair texto do PDF usando PyMuPDF (fitz)"""
    try:
        with fitz.open(file_path) as pdf:
            text = ""
            for page in pdf:
                text += page.get_text()
        return text
    except Exception as e:
        print(f"Erro ao extrair texto com PyMuPDF: {e}")
        return None

def download_attachments(attachments, save_dir):
    pdf_content = None
    os.makedirs(save_dir, exist_ok=True)

    for attachment in attachments:
        if attachment["@odata.type"] == "#microsoft.graph.fileAttachment":
            file_name = attachment["name"]
            file_content_bytes = base64.b64decode(attachment["contentBytes"])
            file_path = os.path.join(save_dir, file_name)

            with open(file_path, "wb") as f:
                f.write(file_content_bytes)
            print(f"Anexo salvo: {file_name}")

            if file_name.lower().endswith(".pdf"):
                try:
                    result = subprocess.run(
                        ["pdftotext", file_path, "-"],
                        capture_output=True,
                        text=True,
                        encoding="utf-8"
                    )
                    if result.returncode == 0:
                        pdf_content = result.stdout
                    else:
                        print("pdftotext não conseguiu extrair texto, tentando PyMuPDF...")
                        pdf_content = extract_text_fallback_with_pymupdf(file_path)
                except Exception as e:
                    print(f"Erro com pdftotext: {e}. Tentando PyMuPDF...")
                    pdf_content = extract_text_fallback_with_pymupdf(file_path)

    return pdf_content


def get_or_create_draft(access_token):
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    subject = "NOTAS CARREGADAS NO PORTAL"
    
    # Buscar rascunho existente
    url = f"https://graph.microsoft.com/v1.0/me/mailFolders/drafts/messages?$filter=subject eq '{subject}'"
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        drafts = response.json()["value"]
        if drafts:
            return drafts[0]["id"]
    
    # Criar novo rascunho
    initial_body = "Boa tarde. Tudo bom?\nSegue abaixo, relação de notas carregadas no portal.\n\n"
    draft_data = {
        "subject": subject,
        "body": {"contentType": "Text", "content": initial_body},
        "toRecipients": []
    }
    
    response = requests.post("https://graph.microsoft.com/v1.0/me/messages", headers=headers, json=draft_data)
    if response.status_code == 201:
        return response.json()["id"]
    return None

def append_text_to_draft(access_token, draft_id, text_line):
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    
    # Obter corpo atual
    url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}"
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        current_body = response.json().get("body", {}).get("content", "")
        # Verifica se a linha já existe para evitar duplicatas em caso de re-processamento no mesmo loop
        if text_line not in current_body:
            new_body = current_body + text_line + "\n"
            update_data = {
                "body": {"contentType": "Text", "content": new_body}
            }
            requests.patch(url, headers=headers, json=update_data)
            print(f"✓ Texto adicionado ao corpo do rascunho: {text_line}")

def attach_email_to_draft(access_token, draft_id, eml_path):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    with open(eml_path, "rb") as f:
        eml_bytes = f.read()
        eml_base64 = base64.b64encode(eml_bytes).decode("utf-8")

    file_name = os.path.basename(eml_path)

    file_attachment = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": file_name,
        "contentType": "message/rfc822",
        "contentBytes": eml_base64
    }

    url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}/attachments"
    response = requests.post(url, headers=headers, json=file_attachment)

    if response.status_code == 201:
        print(f"✓ Anexo .eml enviado para o rascunho: {file_name}")
        return True
    else:
        print(f"❌ Erro ao anexar .eml: {response.status_code} - {response.text}")
        return False



def remove_email_from_draft(access_token, draft_id, nf_number, pedido, protocolo):
    headers = {"Authorization": f"Bearer {access_token}"}
    
    # Buscar anexos do rascunho
    url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}/attachments"
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        attachments = response.json()["value"]
        for attachment in attachments:
            # Verificar se o anexo corresponde à NF cancelada
            if f"{nf_number}" in attachment["name"] and f"{pedido}" in attachment["name"]:
                delete_url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}/attachments/{attachment['id']}"
                requests.delete(delete_url, headers=headers)
                print(f"Anexo removido do rascunho: {attachment['name']}")

def save_email_as_eml(access_token, email_id, save_path):
    """Salva o e-mail completo como arquivo .eml"""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}/$value"

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        with open(save_path, "wb") as f:
            f.write(response.content)
        log(f"E-mail salvo como: {save_path}")
    except RequestException as e:
        log(f"Erro ao salvar e-mail como .eml: {e}")


def mark_email_as_read(message_id, access_token):
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}"
    data = {"isRead": True}
    requests.patch(url, headers=headers, json=data)


def process_emails(callback=None, timer_callback=None):
    global STOP_MONITOR, LOG_CALLBACK, TIMER_CALLBACK
    STOP_MONITOR = False
    LOG_CALLBACK = callback
    TIMER_CALLBACK = timer_callback
    
    log("=== Iniciando monitoramento de e-mails ===")
    
    # Criar pastas se não existirem
    os.makedirs(SAVE_DIRECTORY, exist_ok=True)
    os.makedirs(CARREGADAS_DIR, exist_ok=True)
    
    access_token, refresh_token, expires_at = load_tokens()

    if access_token and datetime.now() < expires_at:
        log("✓ Token válido carregado.")
    else:
        log("⚠️ Token expirado ou inexistente.")
        return "AUTH_REQUIRED"

    draft_id = None
    
    try:
        while not STOP_MONITOR:
            log(f"=== Verificando e-mails - {datetime.now().strftime('%H:%M:%S')} ===")
            
            # Verifica se está prestes a expirar
            _, _, expires_at = load_tokens()
            if datetime.now() > expires_at - timedelta(minutes=5):
                log("🔄 Renovando token...")
                access_token, refresh_token, expires_in = refresh_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET, refresh_token)
                if access_token:
                    save_tokens(access_token, refresh_token, expires_in)
                else:
                    log("Erro ao renovar token. Encerrando.")
                    break


            emails = get_nfe_emails(access_token, NFE_PLATINUM_FOLDER_ID, SENDER_EMAIL, SUBJECT_CONTAINS, BODY_CONTAINS)
            
            if emails:
                # Busca ou cria o rascunho apenas se houver e-mails para processar
                draft_id = get_or_create_draft(access_token)
                
                for email in emails:
                    subject = email["subject"]
                    is_cancelled = "Cancelamento da Nota Fiscal Eletrônica" in subject
                    
                    print(f"Processando: {subject[:50]}...")
                    
                    if email.get("attachments"):
                        # Criar pasta temporária para processar anexos
                        temp_dir = os.path.join(SAVE_DIRECTORY, "temp")
                        os.makedirs(temp_dir, exist_ok=True)
                        
                        pdf_content = download_attachments(email["attachments"], temp_dir)
                        
                        if pdf_content:
                            nf_number, pedido, protocolo = extract_nf_data(pdf_content)
                            
                            if nf_number and pedido and protocolo:
                                # Criar/renomear pasta
                                folder_path = create_or_rename_folder(nf_number, pedido, protocolo, is_cancelled)
                                
                                # Mover arquivos da pasta temp para pasta final
                                for file in os.listdir(temp_dir):
                                    src = os.path.join(temp_dir, file)
                                    dst = os.path.join(folder_path, file)
                                    if os.path.exists(dst):
                                        os.remove(dst)
                                    os.rename(src, dst)
                                
                                # Salvar e-mail como .eml
                                safe_subject = re.sub(r'[\\/*?:"<>|]', "", email['subject'])[:100]  # limita tamanho e limpa caracteres inválidos para nome de arquivo
                                email_path = os.path.join(folder_path, f"{safe_subject}.eml")
                                save_email_as_eml(access_token, email["id"], email_path)
                                
                                # Gerenciar rascunho
                                if is_cancelled:
                                    remove_email_from_draft(access_token, draft_id, nf_number, pedido, protocolo)
                                else:
                                    attach_email_to_draft(access_token, draft_id, email_path)
                                    
                                    # Adicionar lista ao corpo do e-mail
                                    nf_padded = str(nf_number).zfill(8)
                                    text_line = f"PLATINUM TELEINFORMATICA LTDA - Nota Fiscal Eletrônica - {nf_padded} {pedido} {protocolo}"
                                    append_text_to_draft(access_token, draft_id, text_line)
                                
                                log(f"✓ Processado: NF {nf_number} - {pedido} - {protocolo}")
                        
                        # Limpar pasta temp
                        for file in os.listdir(temp_dir):
                            os.remove(os.path.join(temp_dir, file))
                        os.rmdir(temp_dir)
                    
                    mark_email_as_read(email["id"], access_token) #marca e-mail como lido
            
            log(f"Processados {len(emails)} e-mails. Aguardando 5 minutos...")
            # Sleep em pequenos intervalos para permitir interrupção rápida
            for i in range(300, 0, -1):
                if STOP_MONITOR: 
                    update_timer(0)
                    break
                update_timer(i)
                time.sleep(1)
            update_timer(0)
    except Exception as e:
        log(f"Erro crítico no monitor: {e}")
    finally:
        log("=== Monitoramento encerrado ===")

if __name__ == "__main__":
    process_emails()

# Este script monitora a pasta #NFE PLATINUM no Outlook, processa e-mails com anexos de Nota Fiscal Eletrônica,
# extrai dados relevantes, organiza em pastas e salva os e-mails como .eml
# Além disso, gerencia rascunhos para anexar e-mails processados.
# Certifique-se de ter as bibliotecas necessárias instaladas:
# pip install requests PyMuPDF
# Além disso, o pdftotext deve estar instalado e acessível no PATH do sistema
# Para Windows, você pode baixar o pdftotext do site do Xpdf: https

