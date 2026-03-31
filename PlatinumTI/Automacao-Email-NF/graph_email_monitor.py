# -*- coding: utf-8 -*-
"""
Projeto: Platinum NF Automation
Desenvolvido por: Mr Ludolff (Bruno Ludolff)
Descrição: Monitoramento e processamento de NF-e via Microsoft Graph API.
Licença: Este software é de propriedade intelectual de Bruno Ludolff. 
Uso restrito e autorizado apenas para fins específicos de automação interna.
Todos os direitos reservados © 2026.
"""
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
from dotenv import load_dotenv

# Carrega variáveis do arquivo .env
load_dotenv()

# Configurações de autenticação (obtidas do arquivo .env)
TENANT_ID = os.getenv("TENANT_ID")  # ID do Tenant Azure AD
CLIENT_ID = os.getenv("CLIENT_ID")  # ID do Aplicativo Registrado no Azure AD
CLIENT_SECRET = os.getenv("CLIENT_SECRET") # Segredo do Cliente do Aplicativo Registrado

# ID da pasta #NFE PLATINUM (obtido do Graph Explorer)
NFE_PLATINUM_FOLDER_ID = "AAMkAGU5NjJkNGFiLWI5MjItNDU2NS1hZjA0LWY3NDZiNGI4YjYyOAAuAAAAAAC_L7sBebFGTqj2golQ1YDeAQB0ejjyOe8HQLqgrMerz56wAAK16f7eAAA="


# Configurações Globais
SENDER_EMAIL = "noreply@omie.com.br"
SUBJECT_CONTAINS = "PLATINUM TELEINFORMATICA LTDA - Nota Fiscal Eletrônica - "
BODY_CONTAINS = ["CLARO S.A.", "CLARO NXT", "TELMEX DO BRASIL", "CLARO SA"]
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


def find_folder_id_by_name(access_token, folder_name):
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    url = "https://graph.microsoft.com/v1.0/me/mailFolders?$top=100"
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        folders = response.json().get("value", [])
        for folder in folders:
            if folder["displayName"].strip() == folder_name.strip():
                return folder["id"]
        return None
    except Exception as e:
        log(f"❌ Erro ao localizar pasta '{folder_name}': {e}")
        return None

def get_nfe_emails(access_token, folder_id, sender, subject_contains, body_contains_list):
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    
    # Validação dinâmica da pasta se houver erro ou se necessário
    # (Opcional: podemos buscar pela ID fornecida e se falhar tentar pelo nome)
    target_folder_id = folder_id
    
    # Vamos buscar os e-mails não lidos do remetente com o assunto base
    filter_query = (
        f"from/emailAddress/address eq '{sender}' and "
        f"isRead eq false and "
        f"contains(subject, '{subject_contains}')"
    )
    encoded_filter = requests.utils.quote(filter_query)
    url = f"https://graph.microsoft.com/v1.0/me/mailFolders('{target_folder_id}')/messages?$filter={encoded_filter}&$select=subject,body,id,receivedDateTime&$expand=attachments&$top=50"
    
    try:
        response = requests.get(url, headers=headers)
        
        # Se der 404, a ID da pasta mudou. Vamos tentar achar pelo nome.
        if response.status_code == 404:
            log("ℹ️ ID da pasta desatualizado. Tentando localizar por nome...")
            new_id = find_folder_id_by_name(access_token, "#NFE PLATINUM")
            if new_id:
                log(f"✅ Pasta localizada! Nova ID obtida.")
                url = f"https://graph.microsoft.com/v1.0/me/mailFolders('{new_id}')/messages?$filter={encoded_filter}&$select=subject,body,id,receivedDateTime&$expand=attachments&$top=50"
                response = requests.get(url, headers=headers)
            else:
                log("❌ Não foi possível encontrar a pasta '#NFE PLATINUM' pelo nome.")
                return []

        response.raise_for_status()
        all_emails = response.json().get("value", [])
        
        if not all_emails:
            # Check de diagnóstico: ver se existem e-mails que já estão LIDOS
            try:
                diag_filter = f"from/emailAddress/address eq '{sender}' and contains(subject, '{subject_contains}')"
                diag_url = f"https://graph.microsoft.com/v1.0/me/mailFolders('{folder_id}')/messages?$filter={requests.utils.quote(diag_filter)}&$top=5"
                diag_res = requests.get(diag_url, headers=headers).json().get("value", [])
                if diag_res:
                    log(f"ℹ️ Diagnóstico: Encontrei {len(diag_res)}+ e-mails, mas parecem estar todos como LIDOS.")
            except:
                pass
            return []

        filtered_emails = []
        for email in all_emails:
            body_content = email.get("body", {}).get("content", "")
            subject = email.get("subject", "")
            
            # Verificar se contém alguma das palavras-chave ou se o assunto já é muito específico
            match = any(body.upper() in body_content.upper() for body in body_contains_list)
            
            # Fallback: Se o assunto sugerir que é uma nota, prossegue para validar no PDF depois
            if not match and "PLATINUM" in subject.upper() and "NOTA FISCAL" in subject.upper():
                match = True
                log(f"ℹ️ Aguardando validação do PDF (Corpo sem palavras-chave): {subject[:50]}...")

            if match:
                log(f"✅ E-mail compatível encontrado: {subject[:50]}...")
                filtered_emails.append(email)
            else:
                # Log de debug para ajudar a entender o porquê de não estar pegando
                log(f"ℹ️ E-mail ignorado: {subject[:50]}...")
        
        return filtered_emails

    except RequestException as e:
        log(f"❌ Erro ao buscar e-mails: {e}")
        return []

def extract_nf_data(pdf_content):
    # Verificação de Cliente (Claro/NXT) no PDF
    # Se não encontrar as palavras-chave no PDF, ignora o processamento
    if "CLARO" not in pdf_content.upper() and "NXT" not in pdf_content.upper() and "TELMEX" not in pdf_content.upper():
        log("⚠️ Ignorando: Nota Fiscal pertence a outro cliente (não é Claro/NXT).")
        return None, None, None

    # Extrair número NF, pedido e protocolo do PDF
    # Log para debug (opcional, mostra os primeiros 500 chars se falhar)
    
    # Busca Número NF: Tenta padrão "Nº 123.456 Série" ou apenas "Nº 123456"
    nf_match = re.search(r"N[oº]\.?\s*(\d+(?:\.\d+)*)", pdf_content, re.IGNORECASE)
    nf_number = nf_match.group(1).replace(".", "") if nf_match else None
    
    # Busca Pedido e Protocolo: Tenta versões com e sem barra, com e sem espaços
    # Padrão flexível: PEDIDO [números] ... PROTOCOLO [números]
    pedido_match = re.search(r"PEDIDO\s*:?\s*(\d+)", pdf_content, re.IGNORECASE)
    protocolo_match = re.search(r"PROTOCOLO\s*:?\s*(\d+)", pdf_content, re.IGNORECASE)
    
    pedido = pedido_match.group(1) if pedido_match else None
    protocolo = protocolo_match.group(1) if protocolo_match else None
    
    # Log de diagnóstico removido (limpeza final)
    # log(f"📊 Extração: NF={nf_number}, Pedido={pedido}, Protocolo={protocolo}")
    
    if not (nf_number and pedido and protocolo):
        log(f"⚠️ Dados incompletos no PDF da NF {nf_number if nf_number else 'desconhecida'}. Ignorando.")
        
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
        log(f"❌ Erro ao extrair texto com PyMuPDF: {e}")
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
            log(f"📎 Anexo salvo: {file_name}")

            if file_name.lower().endswith(".pdf"):
                try:
                    # Tenta localizar o pdftotext no projeto se não estiver no PATH
                    pdftotext_cmd = "pdftotext"
                    local_xpdf = os.path.join(SCRIPT_DIR, "xpdf-tools-win-4.06", "bin64", "pdftotext.exe")
                    if not subprocess.run(["where.exe", "pdftotext"], capture_output=True).returncode == 0:
                        if os.path.exists(local_xpdf):
                            pdftotext_cmd = local_xpdf
                            log(f"ℹ️ Usando pdftotext local: {pdftotext_cmd}")
                    
                    result = subprocess.run(
                        [pdftotext_cmd, file_path, "-"],
                        capture_output=True,
                        text=True,
                        encoding="utf-8"
                    )
                    if result.returncode == 0 and result.stdout and len(result.stdout.strip()) > 0:
                        pdf_content = result.stdout
                        log("✅ Texto extraído com sucesso.")
                    else:
                        log("⚠️ pdftotext retornou vazio ou falhou, tentando PyMuPDF...")
                        pdf_content = extract_text_fallback_with_pymupdf(file_path)
                except Exception as e:
                    log(f"⚠️ Erro ao executar pdftotext: {e}. Tentando PyMuPDF...")
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

def sort_draft_content(access_token, draft_id):
    """
    Ordena as linhas do rascunho pelo número da Nota Fiscal.
    Mantém o cabeçalho original e ordena apenas as entradas de notas.
    """
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}"
    
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        return

    content = response.json().get("body", {}).get("content", "")
    if not content:
        return

    # Separar cabeçalho das linhas de notas
    header_end_marker = "relação de notas carregadas no portal."
    parts = content.split(header_end_marker)
    
    if len(parts) < 2:
        return

    header = parts[0] + header_end_marker + "\n\n"
    # Filtrar apenas as linhas de Notas Fiscais para ignorar a lista de protocolos do final
    lines = [l.strip() for l in parts[1].split("\n") if "Nota Fiscal Eletrônica" in l]
    
    if not lines:
        return

    # Função para extrair o número da NF para ordenação
    def get_nf_number(line):
        match = re.search(r"Nota Fiscal Eletrônica - (\d+)", line)
        return int(match.group(1)) if match else 0

    # Ordenar linhas
    sorted_lines = sorted(lines, key=get_nf_number)
    
    # Gerar lista de protocolos para o final
    protocols = []
    for line in sorted_lines:
        match_p = re.search(r"(\d+)$", line.strip()) # Pega o último grupo de números (protocolo)
        if match_p:
            protocols.append(match_p.group(1))
            
    protocol_section = "\n".join(protocols)
    new_body = header + "\n".join(sorted_lines) + "\n\n\n" + protocol_section + "\n"

    # Atualizar rascunho
    update_data = {
        "body": {"contentType": "Text", "content": new_body}
    }
    requests.patch(url, headers=headers, json=update_data)
    print("✓ Rascunho ordenado por número de NF.")

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

    if access_token and expires_at and datetime.now() < expires_at:
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
            if expires_at and datetime.now() > expires_at - timedelta(minutes=5):
                log("🔄 Renovando token...")
                access_token, refresh_token, expires_in = refresh_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET, refresh_token)
                if access_token:
                    save_tokens(access_token, refresh_token, expires_in)
                else:
                    log("Erro ao renovar token. Encerrando.")
                    break


            emails = get_nfe_emails(access_token, NFE_PLATINUM_FOLDER_ID, SENDER_EMAIL, SUBJECT_CONTAINS, BODY_CONTAINS)
            log(f"🔎 Foram encontrados {len(emails)} e-mails compatíveis para processar.")
            
            if emails:
                # Busca ou cria o rascunho apenas se houver e-mails para processar
                draft_id = get_or_create_draft(access_token)
                
                for email in emails:
                    subject = email["subject"]
                    is_cancelled = "Cancelamento da Nota Fiscal Eletrônica" in subject
                    
                    log(f"⏳ Processando: {subject[:50]}...")
                    
                    if email.get("attachments"):
                        # Criar pasta temporária para processar anexos
                        temp_dir = os.path.join(SAVE_DIRECTORY, "temp")
                        os.makedirs(temp_dir, exist_ok=True)
                        
                        pdf_content = download_attachments(email["attachments"], temp_dir)
                        
                        if pdf_content:
                            nf_number, pedido, protocolo = extract_nf_data(pdf_content)
                        else:
                            nf_number, pedido, protocolo = None, None, None
                        
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
                
                # Ordenar o rascunho após processar o lote atual
                if draft_id:
                    sort_draft_content(access_token, draft_id)
            
            log(f"Processados {len(emails)} e-mails. Aguardando 1 minuto...")
            # Sleep em pequenos intervalos para permitir interrupção rápida
            for i in range(60, 0, -1):
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

