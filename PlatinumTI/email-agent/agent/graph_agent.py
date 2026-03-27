"""
agent/graph_agent.py
Orquestrador LangGraph para o Vendedor Imortal.
Define o grafo de nos: triage -> (se nao for cotacao: fim) -> attachments -> context -> draft -> save.
"""
from typing import TypedDict, Optional, List, Dict, Any
from langgraph.graph import StateGraph, END
from dotenv import load_dotenv

from agent.triage_node import triage_email
from agent.attachment_node import process_attachments
from agent.context_node import retrieve_style_context
from agent.draft_node import generate_draft
from graph.draft_writer import create_draft
from graph.categorizer import tag_email_as_platinum, ensure_category_exists
from graph.email_reader import extract_body_text

load_dotenv()


class AgentState(TypedDict):
    # Input
    access_token: str
    email_data: dict

    # Processamento
    email_body: str
    is_quote: bool
    triage_confidence: int
    triage_reason: str
    detected_items: List[str]
    attachment_text: str
    style_context: str
    style_examples_count: int

    # Output
    draft_html: str
    draft_subject: str
    draft_to: str
    draft_id: Optional[str]
    error: Optional[str]
    template_applied: bool
    structured_fields: Dict[str, Any]


def prepare_email(state: AgentState) -> AgentState:
    """Extrai o corpo de texto do e-mail para uso nos nos seguintes."""
    email_data = state.get("email_data", {})
    body = extract_body_text(email_data)
    return {**state, "email_body": body}


def save_draft_to_outlook(state: AgentState) -> AgentState:
    """Salva o rascunho gerado no Outlook e aplica a categoria Platinum Sales."""
    token = state.get("access_token", "")
    email_data = state.get("email_data", {})
    message_id = email_data.get("id", "")
    draft_html = state.get("draft_html", "")
    draft_subject = state.get("draft_subject", "")
    draft_to = state.get("draft_to", "")

    try:
        ensure_category_exists(token)

        draft = create_draft(
            token=token,
            to_address=draft_to,
            subject=draft_subject,
            body_html=draft_html,
            in_reply_to_id=message_id,
        )
        draft_id = draft.get("id", "")

        tag_email_as_platinum(token, message_id)

        return {**state, "draft_id": draft_id, "error": None}
    except Exception as e:
        return {**state, "draft_id": None, "error": str(e)}


def should_proceed(state: AgentState) -> str:
    """Roteador condicional: so processa se for uma cotacao."""
    return "process" if state.get("is_quote", False) else "skip"


def build_graph() -> StateGraph:
    graph = StateGraph(AgentState)

    graph.add_node("prepare", prepare_email)
    graph.add_node("triage", triage_email)
    graph.add_node("attachments", process_attachments)
    graph.add_node("context", retrieve_style_context)
    graph.add_node("draft", generate_draft)
    graph.add_node("save", save_draft_to_outlook)

    graph.set_entry_point("prepare")
    graph.add_edge("prepare", "triage")
    graph.add_conditional_edges(
        "triage",
        should_proceed,
        {
            "process": "attachments",
            "skip": END,
        },
    )
    graph.add_edge("attachments", "context")
    graph.add_edge("context", "draft")
    graph.add_edge("draft", "save")
    graph.add_edge("save", END)

    return graph.compile()


def process_email(access_token: str, email_data: dict) -> dict:
    """Processa um unico e-mail pelo grafo do agente."""
    app = build_graph()

    initial_state: AgentState = {
        "access_token": access_token,
        "email_data": email_data,
        "email_body": "",
        "is_quote": False,
        "triage_confidence": 0,
        "triage_reason": "",
        "detected_items": [],
        "attachment_text": "",
        "style_context": "",
        "style_examples_count": 0,
        "draft_html": "",
        "draft_subject": "",
        "draft_to": "",
        "draft_id": None,
        "error": None,
        "template_applied": False,
        "structured_fields": {},
    }

    return app.invoke(initial_state)
