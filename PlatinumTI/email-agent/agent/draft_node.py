"""
agent/draft_node.py
Gera o rascunho de e-mail de resposta.
Fluxo principal: extrai campos estruturados -> renderiza template padrao.
Fallback preferencial: manter template padrao mesmo sem extracao completa.
"""
import json
import os
from datetime import datetime

from agent.config import get_llm_response
from agent.proposta_modelo import renderizar_proposta_modelo, proposta_modelo_disponivel

USE_PROPOSTA_MODELO = os.environ.get("USE_PROPOSTA_MODELO", "true").lower() == "true"

DEFAULT_COMPANY = {
    "empresa_nome": "Platinum Teleinformática Ltda.",
    "empresa_nome_curto": "Platinum",
    "empresa_cnpj": "01.278.897/0001-82",
    "empresa_ie": "85.918.435",
}

DEFAULT_ITEM = {
    "item_quantidade": "[A Confirmar]",
    "item_valor_unitario_bruto": "[A Confirmar]",
    "item_valor_total_bruto": "[A Confirmar]",
}

EXTRACTION_PROMPT = """Voce e um extrator de dados comerciais B2B.
Extraia os campos para preencher um template de proposta.

Retorne APENAS JSON valido (sem markdown) com este formato:
{
  "cliente_empresa": "",
  "cliente_contato_nome": "",
  "processo_descricao": "",
  "itens": [
    {
      "item_descricao": "",
      "item_detalhes": "",
      "item_quantidade": "",
      "item_valor_unitario_bruto": "",
      "item_valor_total_bruto": "",
      "item_observacao": ""
    }
  ],
  "valor_global": "",
  "condicao_entrega": "",
  "condicao_pagamento": "",
  "validade_proposta": "",
  "missing_fields": []
}

Regras:
- LEIA O E-MAIL LINHA POR LINHA. Siga ESTREITAMENTE o detalhe do e-mail do cliente.
- MUITA ATENCAO AS QUANTIDADES: procure palavras como "Quantidade:", "Qtd:" logo abaixo do nome do item. Copie exatamente (ex: "1 unidade", "12 unidades").
- Se a quantidade nao for mencionada no email, deixe empty string `""`.
- Valores monetarios (unitario, total, global) que nao estiverem claros no email devem retornar empty string `""`.
- "valor_global" eh a soma total da proposta.
- "item_detalhes" deve ser usado para variacoes (ex: cores).
- "item_observacao" deve ser usado para avisos.
- SUA RESPOSTA DEVE SER EXCLUSIVAMENTE UM OBJETO JSON VALIDO. NAO envolva a resposta em blocos de codigo (```json). Inicie com { e termine com }.

E-MAIL:
De: [[SENDER_NAME]] <[[SENDER_EMAIL]]>
Assunto: [[SUBJECT]]
Corpo:
[[EMAIL_BODY]]

ANEXOS (texto extraido):
[[ATTACHMENT_TEXT]]
"""

FREEFORM_DRAFT_PROMPT = """Voce e o assistente pessoal de um vendedor de alta performance.
Sua tarefa e redigir uma resposta profissional e personalizada para o e-mail de cotacao abaixo.

E-MAIL DO CLIENTE:
De: {sender_name} <{sender_email}>
Assunto: {subject}

{email_body}

{attachment_section}

ITENS DETECTADOS NO PEDIDO:
{detected_items}

ESTILO DE ESCRITA DO VENDEDOR:
{style_context}

INSTRUCOES PARA O RASCUNHO:
1. Comece com saudacao profissional.
2. Agradeca pelo contato.
3. Liste os itens solicitados.
4. Use placeholders de valor quando faltar dado.
5. Finalize com chamada para acao.

IMPORTANTE: Responda APENAS com o corpo do e-mail em HTML."""


def _strip_code_fences(raw: str) -> str:
    text = (raw or "").strip()
    if text.startswith("```"):
        parts = text.split("```")
        if len(parts) >= 2:
            text = parts[1]
            if text.startswith("json"):
                text = text[4:]
    return text.strip()


def _to_ptbr_date(dt: datetime) -> str:
    months = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
    return f"{dt.day} de {months[dt.month - 1]} de {dt.year}"

def _get_user_info(token: str) -> dict:
    import requests
    try:
        url = "https://graph.microsoft.com/v1.0/me"
        resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=5)
        if resp.status_code == 200:
            return resp.json()
    except Exception:
        pass
    return {}


def _item_from_detected(item_desc: str) -> dict:
    item = {**DEFAULT_ITEM}
    item["item_descricao"] = item_desc or "[ITEM A CONFIRMAR]"
    return item


def _normalize_item(raw_item: dict) -> dict:
    normalized = {**DEFAULT_ITEM}
    
    # Tolerancia a chaves diferentes que o LLM possa inventar
    desc = raw_item.get("item_descricao") or raw_item.get("descricao") or raw_item.get("nome") or ""
    normalized["item_descricao"] = str(desc).strip() or "[ITEM A CONFIRMAR]"

    qtd = raw_item.get("item_quantidade") or raw_item.get("quantidade") or raw_item.get("qtd") or ""
    if str(qtd).strip():
        normalized["item_quantidade"] = str(qtd).strip()

    val_un = raw_item.get("item_valor_unitario_bruto") or raw_item.get("valor_unitario", "")
    if str(val_un).strip():
        normalized["item_valor_unitario_bruto"] = str(val_un).strip()
        
    val_tot = raw_item.get("item_valor_total_bruto") or raw_item.get("valor_total", "")
    if str(val_tot).strip():
        normalized["item_valor_total_bruto"] = str(val_tot).strip()

    for key in ["item_detalhes", "detalhes", "item_observacao", "observacao"]:
        val = str(raw_item.get(key, "")).strip()
        if val:
            normalized[key if key.startswith("item_") else f"item_{key}"] = val

    return normalized


import re

def _parse_raw_llm_json(raw: str) -> dict:
    # 1. Tentar parse direto (com strip fences)
    try:
        return json.loads(_strip_code_fences(raw))
    except Exception:
        pass

    # 2. Tentar buscar apenas o bloco JSON caso haja texto antes/depois
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except Exception:
            pass

    # Se falhou, logar no console para debug
    print(f"\\n[ERRO FATAL] Falha no Parse JSON. RAW OUTPUT:\\n{raw}\\n---\\n")
    raise ValueError("O LLM não retornou um JSON válido.")

def _extract_structured_fields(
    sender_name: str,
    sender_email: str,
    subject: str,
    email_body: str,
    attachment_text: str,
) -> dict:
    prompt = EXTRACTION_PROMPT
    prompt = prompt.replace("[[SENDER_NAME]]", str(sender_name))
    prompt = prompt.replace("[[SENDER_EMAIL]]", str(sender_email))
    prompt = prompt.replace("[[SUBJECT]]", str(subject))
    prompt = prompt.replace("[[EMAIL_BODY]]", str(email_body)[:3000])
    prompt = prompt.replace("[[ATTACHMENT_TEXT]]", str(attachment_text)[:3000])
    
    if attachment_text:
        print(f"\\n[DEBUG DRAFT] Texto extraido dos anexos que vai para a IA:\\n{str(attachment_text)[:1500]}...", flush=True)
    
    # Roda a IA
    raw = get_llm_response(prompt)
    parsed = _parse_raw_llm_json(raw)
    return parsed if isinstance(parsed, dict) else {}


def _build_payload(state: dict, extracted: dict) -> tuple[dict, dict]:
    email_data = state.get("email_data", {})
    sender_info = email_data.get("from", {}).get("emailAddress", {})
    sender_name = sender_info.get("name", "Cliente")
    sender_email = sender_info.get("address", "")
    subject = email_data.get("subject", "Cotacao")
    detected_items = state.get("detected_items", []) or []

    itens = []
    for i, raw_item in enumerate(extracted.get("itens", []) or [], start=1):
        if isinstance(raw_item, dict):
            normalized = _normalize_item(raw_item)
            normalized["item_indice"] = f"{i:02d}"
            itens.append(normalized)

    if not itens:
        if detected_items:
            for i, item_desc in enumerate(detected_items, start=1):
                item = _item_from_detected(item_desc)
                item["item_indice"] = f"{i:02d}"
                itens.append(item)
        else:
            item = _item_from_detected("[ITEM A CONFIRMAR]")
            item["item_indice"] = "01"
            itens.append(item)

    cliente_empresa = (extracted.get("cliente_empresa") or "").strip()
    if not cliente_empresa:
        cliente_empresa = sender_email.split("@")[0].upper() if "@" in sender_email else "CLIENTE"

    contato_nome = (extracted.get("cliente_contato_nome") or sender_name or "Contato").strip()
    if not contato_nome.lower().startswith("sra") and not contato_nome.lower().startswith("sr"):
        contato_nome = "Sra/Sr " + contato_nome

    token = state.get("access_token", "")
    user_info = _get_user_info(token)
    empresa_contato = user_info.get("displayName", "Alexandro Xisto")
    empresa_email = user_info.get("mail") or user_info.get("userPrincipalName", "alex@platinumti.com.br")

    payload = {
        "cliente_empresa": cliente_empresa,
        "cliente_contato_nome": contato_nome,
        "processo_descricao": (extracted.get("processo_descricao") or subject).strip(),
        "itens": itens,
        "valor_global": (extracted.get("valor_global") or "[A Confirmar]").strip(),
        "condicao_entrega": (extracted.get("condicao_entrega") or "5 Dias úteis.").strip(),
        "condicao_pagamento": (extracted.get("condicao_pagamento") or "30 dias.").strip(),
        "validade_proposta": (extracted.get("validade_proposta") or "enquanto durarem os estoques.").strip(),
        "empresa_contato": empresa_contato,
        "empresa_email": empresa_email,
        "empresa_dados_bancarios": "Banco Santander AG: 3848 C/C: 13000420-7",
        "data_atual": _to_ptbr_date(datetime.now()),
        **DEFAULT_COMPANY,
    }

    summary = {
        "cliente_empresa": payload["cliente_empresa"],
        "cliente_contato_nome": payload["cliente_contato_nome"],
        "processo_descricao": payload["processo_descricao"],
        "itens_count": len(payload["itens"]),
        "missing_fields": extracted.get("missing_fields", []),
    }
    return payload, summary


def _generate_freeform_fallback(state: dict) -> tuple[str, str]:
    email_data = state.get("email_data", {})
    sender_info = email_data.get("from", {}).get("emailAddress", {})
    sender_name = sender_info.get("name", "Cliente")
    sender_email = sender_info.get("address", "")
    subject = email_data.get("subject", "Cotacao")
    email_body = state.get("email_body", "")
    attachment_text = state.get("attachment_text", "")
    detected_items = state.get("detected_items", [])
    style_context = state.get("style_context", "")

    attachment_section = f"CONTEUDO DOS ANEXOS:\n{attachment_text}" if attachment_text else ""
    items_text = "\n".join(f"- {item}" for item in detected_items) if detected_items else "(sem itens claros)"

    prompt = FREEFORM_DRAFT_PROMPT.format(
        sender_name=sender_name,
        sender_email=sender_email,
        subject=subject,
        email_body=str(email_body)[:2000],
        attachment_section=str(attachment_section)[:1500],
        detected_items=items_text,
        style_context=str(style_context)[:2000],
    )

    raw = get_llm_response(prompt)
    draft_html = _strip_code_fences(raw)
    if "```html" in draft_html:
        draft_html = draft_html.split("```html")[1].split("```")[0].strip()
    if "<" not in draft_html:
        draft_html = draft_html.replace("\n", "<br>")

    draft_subject = f"Re: {subject}" if not subject.startswith("Re:") else subject
    return draft_html, draft_subject


def generate_draft(state: dict) -> dict:
    """Gera HTML do rascunho (template padrao + fallback seguro)."""
    # [PLAN] Futuramente alterar fluxo para compor arquivo PDF com cabeçalho timbrado e anexar no draft.
    email_data = state.get("email_data", {})
    sender_info = email_data.get("from", {}).get("emailAddress", {})
    sender_email = sender_info.get("address", "")
    subject = email_data.get("subject", "Cotacao")

    # Fallback default
    draft_subject = f"Re: {subject}" if not subject.startswith("Re:") else subject

    if USE_PROPOSTA_MODELO and proposta_modelo_disponivel():
        try:
            extracted = _extract_structured_fields(
                sender_name=sender_info.get("name", "Cliente"),
                sender_email=sender_email,
                subject=subject,
                email_body=state.get("email_body", ""),
                attachment_text=state.get("attachment_text", ""),
            )
            payload, summary = _build_payload(state, extracted)
            draft_html = renderizar_proposta_modelo(payload)

            return {
                **state,
                "draft_html": draft_html,
                "draft_subject": draft_subject,
                "draft_to": sender_email,
                "template_applied": True,
                "structured_fields": summary,
            }
        except Exception as e:
            # Mantem o padrao do template mesmo se a extracao LLM falhar
            try:
                payload, summary = _build_payload(state, {})
                summary["error"] = f"extracao_falhou_usando_template_minimo: {e}"
                draft_html = renderizar_proposta_modelo(payload)
                return {
                    **state,
                    "draft_html": draft_html,
                    "draft_subject": draft_subject,
                    "draft_to": sender_email,
                    "template_applied": True,
                    "structured_fields": summary,
                }
            except Exception as e2:
                # Ultimo fallback: modo antigo
                try:
                    fallback_html, fallback_subject = _generate_freeform_fallback(state)
                    return {
                        **state,
                        "draft_html": fallback_html,
                        "draft_subject": fallback_subject,
                        "draft_to": sender_email,
                        "template_applied": False,
                        "structured_fields": {"error": f"template_e_minimo_falharam: {e2}"},
                    }
                except Exception as e3:
                    return {
                        **state,
                        "draft_html": f"<p>[Erro ao gerar rascunho: {e3}]</p>",
                        "draft_subject": draft_subject,
                        "draft_to": sender_email,
                        "template_applied": False,
                        "structured_fields": {"error": f"falha_total: {e3}"},
                    }

    # Template desativado/indisponivel -> modo antigo
    try:
        fallback_html, fallback_subject = _generate_freeform_fallback(state)
        return {
            **state,
            "draft_html": fallback_html,
            "draft_subject": fallback_subject,
            "draft_to": sender_email,
            "template_applied": False,
            "structured_fields": {"mode": "freeform"},
        }
    except Exception as e:
        return {
            **state,
            "draft_html": f"<p>[Erro ao gerar rascunho: {e}]</p>",
            "draft_subject": draft_subject,
            "draft_to": sender_email,
            "template_applied": False,
            "structured_fields": {"error": str(e)},
        }


