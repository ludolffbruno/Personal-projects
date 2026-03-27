"""
graph/email_reader.py
Lê e-mails da caixa de entrada via Microsoft Graph API.
Suporta download de anexos (PDF, Word, imagens).
"""
import base64
import datetime
import requests
from typing import Optional

GRAPH_BASE = "https://graph.microsoft.com/v1.0"



def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def get_inbox_emails(token: str, count: int = 10) -> list[dict]:
    """
    Retorna os e-mails mais recentes da caixa de entrada limito a 2 dias uteis e nao lidos.
    Cada item contém: id, subject, from, body, hasAttachments, receivedDateTime.
    """
    hoje = datetime.datetime.utcnow()
    dias_voltar = 4 if hoje.weekday() in (0, 1) else 2
    data_limite = (hoje - datetime.timedelta(days=dias_voltar)).strftime("%Y-%m-%dT%H:%M:%SZ")
    
    filter_query = f"isRead eq false and receivedDateTime ge {data_limite}"

    url = (
        f"{GRAPH_BASE}/me/mailFolders/inbox/messages"
        f"?$top={count}"
        f"&$filter={filter_query}"
        f"&$select=id,subject,from,body,hasAttachments,receivedDateTime,categories"
        f"&$orderby=receivedDateTime desc"
    )
    resp = requests.get(url, headers=_headers(token))
    resp.raise_for_status()
    return resp.json().get("value", [])


def get_sent_emails(token: str, count: int = 50) -> list[dict]:
    """
    Retorna e-mails enviados recentemente (para treinar a memória de estilo).
    """
    url = (
        f"{GRAPH_BASE}/me/mailFolders/sentitems/messages"
        f"?$top={count}"
        f"&$select=id,subject,body,toRecipients,sentDateTime"
        f"&$orderby=sentDateTime desc"
    )
    resp = requests.get(url, headers=_headers(token))
    resp.raise_for_status()
    return resp.json().get("value", [])


def get_email_attachments(token: str, message_id: str) -> list[dict]:
    """
    Retorna lista de anexos de um e-mail.
    Cada item: name, contentType, contentBytes (base64).
    """
    url = f"{GRAPH_BASE}/me/messages/{message_id}/attachments"
    resp = requests.get(url, headers=_headers(token))
    resp.raise_for_status()
    return resp.json().get("value", [])


def decode_attachment(attachment: dict) -> bytes:
    """Decodifica o conteúdo base64 de um anexo para bytes."""
    content_bytes = attachment.get("contentBytes", "")
    return base64.b64decode(content_bytes)


def extract_body_text(email: dict) -> str:
    """Extrai o texto limpo do corpo do e-mail (HTML ou texto)."""
    body = email.get("body", {})
    content = body.get("content", "")
    # Remove tags HTML básicas se necessário
    if body.get("contentType", "").lower() == "html":
        import re
        content = re.sub(r"<[^>]+>", " ", content)
        content = re.sub(r"\s+", " ", content).strip()
    return content
