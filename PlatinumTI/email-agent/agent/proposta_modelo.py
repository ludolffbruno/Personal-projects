"""
agent/proposta_modelo.py
Renderizacao do template padrao de proposta (corpo de e-mail).
"""
from pathlib import Path
from typing import Any

CAMINHO_TEMPLATE_PROPOSTA_MODELO = (
    Path(__file__).resolve().parent.parent
    / "proposta-modelo"
    / "proposta_modelo_email_body_template.html"
)


def proposta_modelo_disponivel() -> bool:
    return CAMINHO_TEMPLATE_PROPOSTA_MODELO.exists()


def carregar_proposta_modelo() -> str:
    if not CAMINHO_TEMPLATE_PROPOSTA_MODELO.exists():
        raise FileNotFoundError(f"Template nao encontrado: {CAMINHO_TEMPLATE_PROPOSTA_MODELO}")
    return CAMINHO_TEMPLATE_PROPOSTA_MODELO.read_text(encoding="utf-8")


def renderizar_proposta_modelo(payload: dict[str, Any]) -> str:
    """Renderiza template Mustache para HTML final."""
    try:
        import pystache
    except Exception as e:
        raise RuntimeError(f"pystache nao instalado: {e}")

    template = carregar_proposta_modelo()
    renderer = pystache.Renderer(escape=lambda u: u)
    return renderer.render(template, payload)
