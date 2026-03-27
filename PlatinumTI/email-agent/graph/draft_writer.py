"""
graph/draft_writer.py
Cria rascunhos de e-mail no Outlook via Microsoft Graph API.
NUNCA envia o e-mail — apenas salva na pasta Rascunhos.
"""
import requests
from typing import Optional


GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def create_draft(
    token: str,
    to_address: str,
    subject: str,
    body_html: str,
    in_reply_to_id: Optional[str] = None,
) -> dict:
    """
    Cria um rascunho de resposta na pasta Drafts do Outlook.
    Retorna o objeto do draft criado (com id).
    
    Parâmetros:
        - to_address: endereço do destinatário
        - subject: assunto do e-mail (geralmente Re: X)
        - body_html: corpo em HTML
        - in_reply_to_id: ID do e-mail original (para criar como reply)
    """
    if in_reply_to_id:
        # Cria como reply ao e-mail original
        url = f"{GRAPH_BASE}/me/messages/{in_reply_to_id}/createReply"
        resp = requests.post(url, headers=_headers(token), json={})
        resp.raise_for_status()
        draft = resp.json()
        draft_id = draft["id"]
        
        # Recupera o corpo original gerado pelo Outlook (com o historico)
        get_url = f"{GRAPH_BASE}/me/messages/{draft_id}?$select=body"
        get_resp = requests.get(get_url, headers=_headers(token))
        if get_resp.status_code == 200:
            existing_body = get_resp.json().get("body", {}).get("content", "")
            # Insere a nova resposta acima do historico
            # O Outlook usa tags especificas, mas quebras basicas funcionam
            full_html = f"{body_html}<br><hr><br>{existing_body}"
        else:
            full_html = body_html

        # Atualiza o corpo do rascunho gerado
        patch_url = f"{GRAPH_BASE}/me/messages/{draft_id}"
        patch_body = {
            "body": {
                "contentType": "HTML",
                "content": full_html,
            }
        }
        patch_resp = requests.patch(patch_url, headers=_headers(token), json=patch_body)
        patch_resp.raise_for_status()
        return patch_resp.json()
    else:
        # Cria um novo draft do zero
        url = f"{GRAPH_BASE}/me/messages"
        payload = {
            "subject": subject,
            "isDraft": True,
            "body": {
                "contentType": "HTML",
                "content": body_html,
            },
            "toRecipients": [
                {"emailAddress": {"address": to_address}}
            ],
        }
        resp = requests.post(url, headers=_headers(token), json=payload)
        resp.raise_for_status()
        return resp.json()


def get_drafts(token: str, count: int = 20) -> list[dict]:
    """Lista os rascunhos mais recentes."""
    url = (
        f"{GRAPH_BASE}/me/mailFolders/drafts/messages"
        f"?$top={count}"
        f"&$select=id,subject,toRecipients,createdDateTime,body"
        f"&$orderby=createdDateTime desc"
    )
    resp = requests.get(url, headers=_headers(token))
    resp.raise_for_status()
    return resp.json().get("value", [])
