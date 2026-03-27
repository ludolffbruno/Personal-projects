"""
ui_server.py
Servidor Flask — Vendedor Imortal.
Gerencia autenticação, polling automático (APScheduler) e dashboard.
"""
import logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler("vendedor.log", encoding="utf-8"), logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

import os
import time
import threading
from datetime import datetime
from flask import (
    Flask, render_template, redirect, url_for,
    session, request, jsonify
)
from dotenv import load_dotenv

load_dotenv()

from auth.msal_client import (
    get_access_token, get_login_url,
    exchange_code_for_token, is_authenticated, logout
)
from graph.email_reader import get_inbox_emails
from agent.graph_agent import process_email
from memory.style_updater import sync_sent_emails
from memory.chroma_client import get_collection_stats
from memory.seen_tracker import is_seen, mark_as_seen, get_stats as get_seen_stats

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "vendedor-imortal-secret-2024")

# Intervalo de polling em minutos (respeitando rate limit do Gemini Free: 15 req/min)
POLL_INTERVAL_MINUTES = int(os.environ.get("POLL_INTERVAL_MINUTES", 5))
# Delay entre chamadas Gemini (evita 429 no Free Tier)
GEMINI_DELAY_SECONDS = float(os.environ.get("GEMINI_DELAY_SECONDS", 5))
# Quantos e-mails buscar por rodada (Free Tier: 15 req/min; 1 email = ~2 chamadas)
EMAILS_PER_RUN = int(os.environ.get("EMAILS_PER_RUN", 5))
DASHBOARD_ONLY_QUOTES = os.environ.get("DASHBOARD_ONLY_QUOTES", "false").lower() == "true"

agent_state = {
    "is_running": False,
    "last_run": None,
    "next_run": None,
    "results": [],          # resultados da ÚLTIMA rodada
    "all_results": [],      # histórico de todas as rodadas (últimos 50)
    "error": None,
    "scheduler_active": False,
}

@app.before_request
def log_request():
    logger.info(f">>> [{datetime.now().strftime('%H:%M:%S')}] {request.method} {request.path}")


# ─────────────────────────────────────────────
# Auth
# ─────────────────────────────────────────────
@app.route("/")
def index():
    if is_authenticated():
        return redirect(url_for("dashboard"))
    return render_template("index.html")


@app.route("/auth/login")
def auth_login():
    """Redireciona o usuario para a pagina de login da Microsoft."""
    logger.info(">>> [DEBUG] Entrou na rota /auth/login")
    try:
        logger.info(">>> [DEBUG] Chamando get_login_url()...")
        auth_url, state = get_login_url()
        session["auth_state"] = state
        logger.info(f">>> [DEBUG] Sucesso! Redirecionando para: {auth_url[:50]}...")
        return redirect(auth_url)
    except Exception as e:
        import traceback
        logger.info("!!! [ERRO] Falha critica em auth_login:")
        traceback.print_exc()
        return render_template("index.html", error=str(e))


@app.route("/auth/callback")
def auth_callback():
    """Microsoft redireciona aqui apos o login com o authorization code."""
    logger.info(">>> [GET] /auth/callback - Recebido retorno da Microsoft")
    error = request.args.get("error")
    if error:
        desc = request.args.get("error_description", error)
        logger.info(f"!!! [ERRO] Microsoft retornou erro: {desc}")
        return render_template("index.html", error=desc)

    # Valida o state para prevenir CSRF
    returned_state = request.args.get("state", "")
    expected_state = session.pop("auth_state", None)
    logger.info(f">>> Validando state: {returned_state} vs {expected_state}")
    if returned_state != expected_state:
        logger.info("!!! [ERRO] State mismatch (CSRF?)")
        return render_template("index.html", error="State invalido. Tente fazer login novamente.")

    code = request.args.get("code")
    if not code:
        logger.info("!!! [ERRO] No code in request")
        return render_template("index.html", error="Codigo de autorizacao nao recebido.")

    try:
        logger.info(">>> Trocando code por token...")
        exchange_code_for_token(code)
        logger.info(">>> Auth concluida com sucesso!")
        _start_scheduler()
        return redirect(url_for("dashboard"))
    except Exception as e:
        import traceback
        logger.info("!!! [ERRO] Falha ao trocar code por token:")
        traceback.print_exc()
        return render_template("index.html", error=str(e))


@app.route("/auth/logout")
def auth_logout():
    logout()
    session.clear()
    agent_state["scheduler_active"] = False
    return redirect(url_for("index"))


# ─────────────────────────────────────────────
# Dashboard
# ─────────────────────────────────────────────
@app.route("/dashboard")
def dashboard():
    if not is_authenticated():
        return redirect(url_for("index"))
    stats = get_collection_stats()
    seen_stats = get_seen_stats()
    # Mostra todos os e-mails e inverte a lista para o mais recente ficar no topo
    dashboard_results = agent_state["all_results"][::-1]
    
    return render_template(
        "dashboard.html",
        is_running=agent_state["is_running"],
        last_run=agent_state["last_run"],
        next_run=agent_state["next_run"],
        results=dashboard_results[-20:],  # últimos 20 processados (filtrados)
        error=agent_state["error"],
        memory_stats=stats,
        seen_stats=seen_stats,
        scheduler_active=agent_state["scheduler_active"],
        poll_interval=POLL_INTERVAL_MINUTES,
    )


# ─────────────────────────────────────────────
# API
# ─────────────────────────────────────────────
@app.route("/api/run-agent", methods=["POST"])
def api_run_agent():
    """Dispara o agente manualmente (uma rodada)."""
    if not is_authenticated():
        return jsonify({"error": "Nao autenticado"}), 401
    if agent_state["is_running"]:
        return jsonify({"error": "Agente ja esta rodando"}), 409

    thread = threading.Thread(target=_run_agent_task, daemon=True)
    thread.start()
    return jsonify({"status": "started"})


@app.route("/api/status")
def api_status():
    return jsonify({
        "is_running": agent_state["is_running"],
        "last_run": agent_state["last_run"],
        "next_run": agent_state["next_run"],
        "results_count": len(agent_state["all_results"]),
        "error": agent_state["error"],
        "scheduler_active": agent_state["scheduler_active"],
    })


@app.route("/api/results")
def api_results():
    return jsonify(agent_state["all_results"][-20:])


@app.route("/api/sync-memory", methods=["POST"])
def api_sync_memory():
    """Sincroniza manualmente a memória de estilo com e-mails enviados."""
    if not is_authenticated():
        return jsonify({"error": "Nao autenticado"}), 401
    try:
        token = get_access_token()
        result = sync_sent_emails(token, count=50)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/scheduler/toggle", methods=["POST"])
def api_toggle_scheduler():
    """Liga/desliga o scheduler automático."""
    if not is_authenticated():
        return jsonify({"error": "Nao autenticado"}), 401
    if agent_state["scheduler_active"]:
        agent_state["scheduler_active"] = False
        agent_state["next_run"] = None
        return jsonify({"active": False})
    else:
        _start_scheduler()
        return jsonify({"active": True})


# ─────────────────────────────────────────────
# Scheduler (loop em thread separada)
# ─────────────────────────────────────────────
def _start_scheduler():
    """Inicia o loop de polling automático em background."""
    if agent_state["scheduler_active"]:
        return
    agent_state["scheduler_active"] = True
    thread = threading.Thread(target=_scheduler_loop, daemon=True)
    thread.start()
    logger.info(f"[Scheduler] Iniciado — verificando novos e-mails a cada {POLL_INTERVAL_MINUTES} minutos")


def _scheduler_loop():
    """Loop que executa o agente periodicamente."""
    while agent_state["scheduler_active"]:
        if is_authenticated():
            _run_agent_task()
        interval = POLL_INTERVAL_MINUTES * 60
        next_run_time = datetime.now()
        import datetime as dt
        next_run_time = datetime.now() + dt.timedelta(minutes=POLL_INTERVAL_MINUTES)
        agent_state["next_run"] = next_run_time.strftime("%H:%M:%S")
        # Espera o intervalo, verificando o flag a cada 10s
        elapsed = 0
        while elapsed < interval and agent_state["scheduler_active"]:
            time.sleep(10)
            elapsed += 10


def _run_agent_task():
    """
    Núcleo do agente: busca novos e-mails, processa somente os NÃO VISTOS.
    Rate-limit respeitado com delay entre chamadas Gemini.
    """
    if agent_state["is_running"]:
        return

    agent_state["is_running"] = True
    agent_state["error"] = None

    try:
        token = get_access_token()
        if not token:
            raise ValueError("Token expirado. Faca login novamente.")

        # Busca e-mails recentes (Aumentado para 100 para pegar e-mails antigos que chegaram enquanto o script estava offline)
        emails = get_inbox_emails(token, count=100)

        # Filtra apenas os NÃO VISTOS
        new_emails = [e for e in emails if not is_seen(e.get("id", ""))]

        if not new_emails:
            logger.info("[Agent] Nenhum e-mail novo para processar.")
            agent_state["last_run"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            return

        # Limita ao batch configurado por rodada para respeitar rate limits do Gemini
        new_emails = new_emails[:EMAILS_PER_RUN]
        logger.info(f"[Agent] Processando {len(new_emails)} e-mail(s) novo(s)...")

        batch_results = []
        for i, email in enumerate(new_emails):
            subject = email.get("subject", "(sem assunto)")
            sender = email.get("from", {}).get("emailAddress", {}).get("address", "")
            email_id = email.get("id", "")

            logger.info(f"[Agent] [{i+1}/{len(new_emails)}] {subject}")

            final_state = process_email(access_token=token, email_data=email)

            result = {
                "email_id": email_id,
                "subject": subject,
                "sender": sender,
                "received": email.get("receivedDateTime", ""),
                "is_quote": final_state.get("is_quote", False),
                "confidence": final_state.get("triage_confidence", 0),
                "reason": final_state.get("triage_reason", ""),
                "detected_items": final_state.get("detected_items", []),
                "has_client_history": final_state.get("has_client_history", False),
                "draft_created": bool(final_state.get("draft_id")),
                "draft_subject": final_state.get("draft_subject", ""),
                "draft_preview": _get_text_preview(final_state.get("draft_html", "")),
                "error": final_state.get("error"),
                "template_applied": final_state.get("template_applied", False),
                "structured_fields": final_state.get("structured_fields", {}),
                "processed_at": datetime.now().strftime("%d/%m %H:%M"),
            }
            batch_results.append(result)

            # Marca como visto independente de ser cotação ou não
            mark_as_seen(
                email_id,
                subject=subject,
                sender=sender,
                is_quote=result["is_quote"],
                draft_created=result["draft_created"],
            )

            # Rate-limit: delay entre chamadas Gemini (exceto no último)
            if i < len(new_emails) - 1:
                logger.info(f"[Agent] Aguardando {GEMINI_DELAY_SECONDS}s (rate limit)...")
                time.sleep(GEMINI_DELAY_SECONDS)

        # Atualiza estado global
        agent_state["results"] = batch_results
        # Akumulace no histórico (máx 50)
        agent_state["all_results"] = (agent_state["all_results"] + batch_results)[-50:]
        agent_state["last_run"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        quotes_found = sum(1 for r in batch_results if r["is_quote"])
        logger.info(f"[Agent] Concluido. {quotes_found} cotacao(oes) encontrada(s) de {len(batch_results)} novos e-mail(s).")

    except Exception as e:
        import traceback
        agent_state["error"] = str(e)
        logger.info(f"[Agent] Erro: {e}\n{traceback.format_exc()}")
    finally:
        agent_state["is_running"] = False


def _get_text_preview(html: str, max_len: int = 300) -> str:
    import re
    text = re.sub(r"<[^>]+>", " ", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:max_len] + "..." if len(text) > max_len else text


# ─────────────────────────────────────────────
# Startup
# ─────────────────────────────────────────────
if __name__ == "__main__":
    try:
        logger.info("="*50)
        logger.info("VENDEDOR IMORTAL -- INICIANDO")
        logger.info(f"URL: http://localhost:5001")
        logger.info("="*50)
        
        # Check ENV
        missing = [v for v in ["AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "GEMINI_API_KEY"] if not os.environ.get(v)]
        if missing:
            logger.info(f"!!! AVISO: Variaveis faltando no .env: {', '.join(missing)}")

        if is_authenticated():
            logger.info(">>> Usuario ja autenticado, ligando scheduler...")
            _start_scheduler()
        else:
            logger.info(">>> Aguardando login do usuario...")

        from waitress import serve
        serve(app, host='0.0.0.0', port=5001)
    except Exception as e:
        import traceback
        logger.info("!!! [ERRO FATAL] O servidor parou abruptamente:")
        traceback.print_exc()






