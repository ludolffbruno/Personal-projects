"""
memory/style_updater.py
Aprende com e-mails enviados pelo vendedor.
Monitora a pasta SentItems e sincroniza com o ChromaDB.
"""
from graph.email_reader import get_sent_emails, extract_body_text
from memory.chroma_client import add_sent_email, get_collection_stats


def sync_sent_emails(token: str, count: int = 50) -> dict:
    """
    Busca os e-mails enviados recentes e os adiciona ao ChromaDB.
    Retorna um resumo da sincronização.
    """
    print(f"[StyleUpdater] Buscando últimos {count} e-mails enviados...")
    sent_emails = get_sent_emails(token, count=count)

    added = 0
    skipped = 0

    for email in sent_emails:
        email_id = email.get("id", "")
        subject = email.get("subject", "(sem assunto)")
        body = extract_body_text(email)
        to_recipients = email.get("toRecipients", [])
        to_address = ""
        if to_recipients:
            to_address = to_recipients[0].get("emailAddress", {}).get("address", "")
        sent_date = email.get("sentDateTime", "")

        # Ignora e-mails sem corpo
        if len(body.strip()) < 20:
            skipped += 1
            continue

        was_added = add_sent_email(
            email_id=email_id,
            subject=subject,
            body=body,
            to_address=to_address,
            sent_date=sent_date,
        )
        if was_added:
            added += 1
        else:
            skipped += 1

    stats = get_collection_stats()
    result = {
        "emails_processed": len(sent_emails),
        "new_added": added,
        "already_indexed": skipped,
        "total_in_memory": stats["total_emails"],
    }
    print(f"[StyleUpdater] Sync concluído: {result}")
    return result
