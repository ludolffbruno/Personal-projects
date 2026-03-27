"""
auth/msal_client.py
Fluxo OAuth2 Authorization Code com redirect URL.
O usuário faz login normalmente no browser Microsoft e é redirecionado de volta.
Usa ConfidentialClientApplication (com client_secret).
"""
import os
import uuid
import msal
from pathlib import Path


SCOPES = [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Mail.ReadWrite",
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/MailboxSettings.Read",
]

REDIRECT_URI = "http://localhost:5001/auth/callback"
TOKEN_CACHE_FILE = Path(__file__).parent.parent / ".token_cache.json"


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text())
    return cache


def _save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        TOKEN_CACHE_FILE.write_text(cache.serialize())


def _build_app(cache: msal.SerializableTokenCache) -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        client_id=os.environ["AZURE_CLIENT_ID"],
        client_credential=os.environ["AZURE_CLIENT_SECRET"],
        authority=f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}",
        token_cache=cache,
    )


def get_login_url() -> tuple[str, str]:
    """
    Gera a URL de login Microsoft e um state aleatório para CSRF.
    Retorna (auth_url, state).
    """
    cache = _load_cache()
    app = _build_app(cache)
    state = str(uuid.uuid4())
    auth_url = app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        state=state,
    )
    return auth_url, state


def exchange_code_for_token(code: str) -> str:
    """
    Troca o authorization code recebido no callback por um access token.
    Salva o token no cache local.
    """
    cache = _load_cache()
    app = _build_app(cache)
    result = app.acquire_token_by_authorization_code(
        code=code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    _save_cache(cache)
    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Erro desconhecido"))
        raise ValueError(f"Falha na autenticacao: {error}")
    return result["access_token"]


def get_access_token() -> str | None:
    """
    Retorna um access token valido do cache (silencioso).
    Retorna None se o usuario precisar fazer login novamente.
    """
    cache = _load_cache()
    app = _build_app(cache)
    accounts = app.get_accounts()
    if not accounts:
        return None
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    _save_cache(cache)
    if result and "access_token" in result:
        return result["access_token"]
    return None


def is_authenticated() -> bool:
    return get_access_token() is not None


def logout():
    if TOKEN_CACHE_FILE.exists():
        TOKEN_CACHE_FILE.unlink()
