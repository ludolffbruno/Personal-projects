"""
agent/triage_node.py
Nó LangGraph: Classifica o e-mail como pedido de cotação ou não.
Usa a bridge configurada (Gemini ou Groq) para análise semântica.
"""
import json
from agent.config import get_llm_response

TRIAGE_PROMPT = """Você é um assistente especializado em vendas B2B. Analise o e-mail abaixo e classifique-o.

E-MAIL:
De: {sender}
Assunto: {subject}
Corpo:
{body}

Responda APENAS com um JSON válido no seguinte formato (sem markdown, sem explicação extra):
{{
  "is_quote_request": true ou false,
  "confidence": número de 0 a 100,
  "reason": "explicação curta em português de por que é ou não é uma solicitação de cotação",
  "detected_items": ["item1", "item2"] ou [] se não houver itens detectados
}}

CRITÉRIOS DE COTAÇÃO (true):
- Pedidos de preço, orçamento, proposta comercial
- Solicitações de quantidade + produto/serviço
- Manifestações de interesse em comprar com pedido de valores
- E-mails com lista de itens para cotar

NÃO É COTAÇÃO (false):
- Newsletters, e-mails de marketing
- Confirmações de pedidos já feitos
- E-mails administrativos, financeiros, de RH
- Spam ou e-mails não relacionados a vendas
- Agradecimentos sem solicitação
- Propostas, orçamentos, contratos ou documentação (ex: API) ENVIADOS por parceiros ou fornecedores PARA A NOSSA EMPRESA.
- Regra de Ouro: O remetente DEVE ser um cliente querendo COMPRAR algo. Se ele está VENDENDO ou apresentando um produto para nós, retorne false."""

def triage_email(state: dict) -> dict:
    """
    Nó LangGraph: recebe o estado com email_data e retorna is_quote, reason, detected_items.
    """
    email_data = state.get("email_data", {})
    sender = email_data.get("from", {}).get("emailAddress", {}).get("address", "desconhecido")
    subject = email_data.get("subject", "(sem assunto)")
    body = state.get("email_body", "")

    prompt = TRIAGE_PROMPT.format(sender=sender, subject=subject, body=body[:3000])

    try:
        raw = get_llm_response(prompt)
        
        # Remove markdown code blocks se presentes
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        raw = raw.strip()

        result = json.loads(raw)
    except (json.JSONDecodeError, Exception) as e:
        result = {
            "is_quote_request": False,
            "confidence": 0,
            "reason": f"Erro ao processar resposta da IA: {str(e)}",
            "detected_items": [],
        }

    return {
        **state,
        "is_quote": result.get("is_quote_request", False),
        "triage_confidence": result.get("confidence", 0),
        "triage_reason": result.get("reason", ""),
        "detected_items": result.get("detected_items", []),
    }
