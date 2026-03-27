"""
agent/context_node.py
Nó LangGraph: Recupera exemplos de e-mails similares do ChromaDB.
PRIORIDADE: Exemplos para aquele cliente específico → depois estilo geral do vendedor.
"""
from memory.chroma_client import query_similar_emails, query_by_client


def retrieve_style_context(state: dict) -> dict:
    """
    Nó LangGraph: Busca e-mails similares priorizando o histórico com aquele cliente.
    """
    email_body = state.get("email_body", "")
    email_data = state.get("email_data", {})
    subject = email_data.get("subject", "")
    sender_email = email_data.get("from", {}).get("emailAddress", {}).get("address", "")
    detected_items = state.get("detected_items", [])

    # Query text combinando assunto, corpo e itens
    query_parts = [subject, email_body[:500]]
    if detected_items:
        query_parts.append("Itens: " + ", ".join(detected_items))
    query_text = "\n".join(query_parts)

    client_examples = []
    general_examples = []

    # 1️⃣ Primeiro: busca respostas anteriores para ESTE cliente específico
    if sender_email:
        client_examples = query_by_client(sender_email, n_results=2)

    # 2️⃣ Depois: complementa com estilo geral do vendedor
    general_examples = query_similar_emails(query_text, n_results=3)

    # Monta o contexto priorizando o cliente
    sections = []

    if client_examples:
        examples_text = []
        for i, item in enumerate(client_examples, 1):
            doc = item["document"][:600] + "..." if len(item["document"]) > 600 else item["document"]
            examples_text.append(f"  Resposta anterior {i} para este cliente:\n  {doc}")
        sections.append(
            "HISTORICO COM ESTE CLIENTE (prioridade maxima):\n"
            + "\n\n".join(examples_text)
        )

    if general_examples:
        examples_text = []
        for i, item in enumerate(general_examples, 1):
            sim = item.get("similarity", 0)
            doc = item["document"][:500] + "..." if len(item["document"]) > 500 else item["document"]
            examples_text.append(f"  Exemplo {i} ({sim:.0%} similar):\n  {doc}")
        sections.append(
            "ESTILO GERAL DO VENDEDOR:\n"
            + "\n\n".join(examples_text)
        )

    if not sections:
        style_context = (
            "Nenhum exemplo disponivel ainda. "
            "Use tom profissional, cordial e objetivo. "
            "Cumprimente pelo nome se disponivel, seja direto."
        )
    else:
        style_context = "\n\n" + "\n\n---\n\n".join(sections)

    return {
        **state,
        "style_context": style_context,
        "style_examples_count": len(client_examples) + len(general_examples),
        "has_client_history": len(client_examples) > 0,
    }
