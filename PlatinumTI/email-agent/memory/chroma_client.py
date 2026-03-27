"""
memory/chroma_client.py
Gerencia o banco vetorial ChromaDB local para memória de estilo do vendedor.
Armazena e-mails enviados e recupera exemplos similares para manter o tom de voz.
"""
import chromadb
from chromadb.config import Settings
from pathlib import Path
from typing import Optional
import hashlib


CHROMA_PATH = Path(__file__).parent.parent / "chroma_data"
COLLECTION_NAME = "seller_style"


def _get_client() -> chromadb.PersistentClient:
    return chromadb.PersistentClient(path=str(CHROMA_PATH))


def _get_collection() -> chromadb.Collection:
    client = _get_client()
    return client.get_or_create_collection(
        name=COLLECTION_NAME,
        metadata={"hnsw:space": "cosine"},
    )


def add_sent_email(
    email_id: str,
    subject: str,
    body: str,
    to_address: str,
    sent_date: str,
) -> bool:
    """
    Adiciona um e-mail enviado à memória de estilo.
    Usa o ID do e-mail como chave única para evitar duplicatas.
    """
    collection = _get_collection()

    # Verifica se já existe
    existing = collection.get(ids=[email_id])
    if existing["ids"]:
        return False  # Já indexado

    full_text = f"Assunto: {subject}\n\n{body}"

    collection.add(
        documents=[full_text],
        metadatas=[{
            "email_id": email_id,
            "subject": subject,
            "to_address": to_address,
            "sent_date": sent_date,
            "type": "sent_email",
        }],
        ids=[email_id],
    )
    return True


def query_similar_emails(query_text: str, n_results: int = 3) -> list[dict]:
    """
    Recupera os N e-mails mais similares ao pedido do cliente.
    Retorna uma lista de dicts com document e metadata.
    """
    collection = _get_collection()

    count = collection.count()
    if count == 0:
        return []

    actual_n = min(n_results, count)
    results = collection.query(
        query_texts=[query_text],
        n_results=actual_n,
        include=["documents", "metadatas", "distances"],
    )

    output = []
    if results and results["documents"]:
        for doc, meta, dist in zip(
            results["documents"][0],
            results["metadatas"][0],
            results["distances"][0],
        ):
            output.append({
                "document": doc,
                "metadata": meta,
                "similarity": round(1 - dist, 3),
            })
    return output



def query_by_client(client_email: str, n_results: int = 2) -> list[dict]:
    """
    Recupera e-mails enviados para um cliente específico.
    Usado para personalizar respostas com base no histórico com aquele contato.
    """
    collection = _get_collection()

    count = collection.count()
    if count == 0:
        return []

    # Filtra por to_address usando where clause do ChromaDB
    try:
        results = collection.get(
            where={"to_address": client_email},
            include=["documents", "metadatas"],
            limit=n_results,
        )
        output = []
        if results and results["documents"]:
            for doc, meta in zip(results["documents"], results["metadatas"]):
                output.append({"document": doc, "metadata": meta, "similarity": 1.0})
        return output
    except Exception:
        return []


def get_collection_stats() -> dict:
    """Retorna estatísticas da coleção de memória."""
    collection = _get_collection()
    return {
        "total_emails": collection.count(),
        "collection_name": COLLECTION_NAME,
        "storage_path": str(CHROMA_PATH),
    }
