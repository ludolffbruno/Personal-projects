"""
graph/categorizer.py
Aplica a categoria "Platinum Sales" a e-mails identificados como cotações.
A categoria aparece visualmente no Outlook como uma etiqueta colorida.
"""
import requests


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
PLATINUM_CATEGORY = "Platinum Sales"


def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def ensure_category_exists(token: str) -> bool:
    """
    Verifica se a categoria 'Platinum Sales' existe na caixa.
    Se não existir, cria com a cor laranja (preset8).
    """
    url = f"{GRAPH_BASE}/me/outlook/masterCategories"
    resp = requests.get(url, headers=_headers(token))
    resp.raise_for_status()
    categories = resp.json().get("value", [])

    existing = [c for c in categories if c.get("displayName") == PLATINUM_CATEGORY]
    if existing:
        return True

    # Cria a categoria com cor laranja (preset8 = laranja dourado)
    create_url = f"{GRAPH_BASE}/me/outlook/masterCategories"
    payload = {
        "displayName": PLATINUM_CATEGORY,
        "color": "preset8",  # Laranja dourado — cor premium
    }
    create_resp = requests.post(create_url, headers=_headers(token), json=payload)
    return create_resp.status_code in (200, 201)


def tag_email_as_platinum(token: str, message_id: str) -> bool:
    """
    Adiciona a categoria 'Platinum Sales' ao e-mail especificado.
    Preserva categorias existentes.
    """
    # Busca categorias atuais do e-mail
    get_url = f"{GRAPH_BASE}/me/messages/{message_id}?$select=categories"
    get_resp = requests.get(get_url, headers=_headers(token))
    get_resp.raise_for_status()
    current_categories = get_resp.json().get("categories", [])

    if PLATINUM_CATEGORY in current_categories:
        return True  # Já está marcado

    updated_categories = list(set(current_categories + [PLATINUM_CATEGORY]))

    patch_url = f"{GRAPH_BASE}/me/messages/{message_id}"
    payload = {"categories": updated_categories}
    patch_resp = requests.patch(patch_url, headers=_headers(token), json=payload)
    return patch_resp.status_code == 200


def remove_platinum_tag(token: str, message_id: str) -> bool:
    """Remove a categoria 'Platinum Sales' de um e-mail."""
    get_url = f"{GRAPH_BASE}/me/messages/{message_id}?$select=categories"
    get_resp = requests.get(get_url, headers=_headers(token))
    get_resp.raise_for_status()
    current_categories = get_resp.json().get("categories", [])

    updated_categories = [c for c in current_categories if c != PLATINUM_CATEGORY]

    patch_url = f"{GRAPH_BASE}/me/messages/{message_id}"
    payload = {"categories": updated_categories}
    patch_resp = requests.patch(patch_url, headers=_headers(token), json=payload)
    return patch_resp.status_code == 200
