# 🧠 Memória Técnica: Projeto "Vendedor Imortal" (email-agent)

Este documento contém toda a base técnica, lógica e arquitetural necessária para a continuidade do desenvolvimento do **Vendedor Imortal**. Use este arquivo para contextualizar qualquer nova instância de IA.

---

## 🏗️ 1. Arquitetura do Sistema

### 📦 Componentes Core
*   **Servidor UI (`ui_server.py`):** Flask (Porta 5001). Gerencia rotas, autenticação OAuth2 (MSAL/Microsoft Graph) e o agendador de tarefas (`APScheduler`).
*   **Orquestrador de IA (`agent/graph_agent.py`):** Utiliza **LangGraph** para manter o estado e fluxo de decisão.
*   **Banco de Dados:**
    *   `ChromaDB` (Local em `chroma_data/`): Armazena vetores de e-mails enviados historicamente para busca por similaridade de estilo.
    *   `seen_emails.db` (SQLite): Rastreador de duplicatas para evitar processar o mesmo e-mail duas vezes.
*   **Bridge de LLM (`agent/config.py`):** Implementa um sistema de **Failover Circular**. Ordem de tentativa: `Gemini 2.5 Flash Lite` -> `Groq (Llama 3.3 / Llama 3.2 Vision)` -> `OpenAI (GPT-4o)`.
    *   **Failover de Visão:** Groq está configurado com `llama-3.2-11b-vision-preview` para garantir leitura de imagens mesmo sem Gemini.

---

## ⚙️ 2. O Fluxo do Agente (LangGraph Nodes)

Todo e-mail passa pelo grafo definido em `build_graph()`:

1.  **`prepare`**: Extrai o texto limpo do corpo do e-mail.
2.  **`triage`**: 
    *   **Lógica:** Classifica se é `is_quote` (true/false).
    *   **Filtro Anti-Vendedor:** Configurado para ignorar e-mails onde a nossa empresa é o cliente (ex: parceiros vendendo APIs/serviços para nós).
3.  **`attachments`**:
    *   **Word:** Suporte total a **tabelas** (iterando `doc.tables`), essencial para capturar itens em grades.
    *   **Image Hunter:** Bypass da flag `hasAttachments` do Outlook se `cid:` for detectado.
    *   **Base64 Support:** Extração via Regex de imagens `data:image` embutidas no HTML do corpo.
4.  **`context`**: Busca no ChromaDB exemplos de e-mails enviados anteriormente para clonar o tom de voz do vendedor.
5.  **`draft`**: 
    *   Utiliza placeholders `[[CLIENTE]]` e substituição manual para evitar erros de JSON (`KeyError`).
    *   Prompt otimizado para extrair quantidades de forma rígida.
6.  **`save`**: Salva no Outlook na pasta **Rascunhos** como Resposta (Reply) para manter histórico, e aplica a categoria **"Platinum Sales"**.

---

## 🚦 3. Configurações de Performance e Cotas (.env)

Devido às limitações da **Free Tier do Gemini** (20 RPM), as seguintes configurações foram otimizadas e devem ser mantidas:

*   `POLL_INTERVAL_MINUTES=5`: O scheduler verifica a caixa a cada 5 minutos.
*   `EMAILS_PER_RUN=2`: Processa no máximo 2 e-mails por ciclo para não estourar o limite de tokens/minuto da IA.
*   `GEMINI_DELAY_SECONDS=5`: Pausa entre chamadas para evitar o erro `429 RESOURCE_EXHAUSTED`.

---

## 🛠️ 4. Regras de Interface (dashboard.html)

O Dashboard foi restaurado para o padrão "Premium Dark":
*   **Ordenação:** Lista de e-mails sempre `all_results[::-1]` (Mais recentes no topo).
*   **Visibilidade:** Exibe todos os e-mails processados (Cotações com badge laranja, Não-cotações com badge cinza).
*   **Interação:** Cards expansíveis exibindo o "Raciocínio da IA", "Itens Detectados" e "Prévia do Rascunho".

---

## 🎯 5. Roadmap e Próximas Implementações

1.  **PDF Timbrado (Próximo Passo):**
    *   Atualmente o sistema gera apenas o HTML para o corpo do e-mail.
    *   **Objetivo:** Implementar geração de PDF (usando `ReportLab` ou `FPDF`) com o papel timbrado da Platinum TI. 
2.  **Consulta de Estoque Real:** Integrar com o banco de dados de produtos da Platinum para validar quantidades.
3.  **Integração com ERP (Futuro):** Consultar estoque real no Tiny/Bling antes de gerar o rascunho.

---
**Atualizado em:** 27 de Março de 2024
**Responsável:** Antigravity Agent (Google Deepmind)
