# 🌬️🦾 Vendedor Imortal — Outlook AI Agent

Agente de IA que vive dentro do Microsoft Outlook e age como o "braço direito" de um vendedor de alta performance. Lê pedidos de orçamento, gera rascunhos personalizados e marca e-mails como "Platinum Sales" — automaticamente.

## ✨ Funcionalidades

- **Triagem Cirúrgica**: Distingue cotações reais de spam/newsletters com 95%+ de precisão
- **Memória de Estilo**: Aprende como o vendedor escreve e replica seu tom de voz (ChromaDB)
- **Visão Multimodal Avançada**: Lê pedidos em PDFs, Word (incluindo tabelas!), Imagens em anexo e até imagens coladas no corpo (Base64/Inline)
- **Drafts Automáticos**: Salva rascunhos no Outlook preservando o histórico da conversa — o vendedor só clica em Enviar
- **Failover Inteligente**: Se o Gemini atingir o limite de cota, o Groq assume a triagem e a visão (OCR) automaticamente
- **Dashboard Premium**: Interface Dark Mode para monitorar o raciocínio da IA e logs de extração em tempo real

## 🏗️ Arquitetura

```
Outlook (MS Graph) ──► Flask Dashboard ──► LangGraph Agent
                                               ├── Gemini 2.5 Flash / Groq Llama 3.2 Vision
                                               ├── ChromaDB (Memória de Estilo)
                                               └── Parsers (PDF / Word Tables / Base64 Images)
                        ◄────────────────── Drafts + Categorias
```

## 🚀 Setup

### 1. Pré-requisitos
- Python 3.11+
- Conta Microsoft 365 (Outlook)
- App Registration no Azure (já criado: `Outlook-AI-Agent`)
- Chave da API Gemini (gratuita em [aistudio.google.com](https://aistudio.google.com))

### 2. Instalação

```bash
# Clone / navegue até o projeto
cd email-agent

# Crie e ative o ambiente virtual
python -m venv venv
venv\Scripts\activate   # Windows

# Instale as dependências
pip install -r requirements.txt
```

### 3. Configuração

```bash
# Copie o template de ambiente
copy .env.example .env
```

Edite o `.env` e preencha:
- `AZURE_CLIENT_SECRET` — Crie em: Azure Portal > App registrations > Outlook-AI-Agent > Certificates & secrets
- `GEMINI_API_KEY` — Obtenha em: https://aistudio.google.com/app/apikey
- `FLASK_SECRET_KEY` — Qualquer string longa e aleatória

### 4. Permissões Azure (Obrigatório)

No Portal Azure, vá em **App registrations > Outlook-AI-Agent > API permissions** e adicione:

| Permission | Type | Admin Consent |
|---|---|---|
| `Mail.Read` | Delegated | ✅ |
| `Mail.ReadWrite` | Delegated | ✅ |
| `Mail.Send` | Delegated | ✅ (para ler Sent Items) |
| `MailboxSettings.Read` | Delegated | ✅ |

### 5. Executar

```bash
python ui_server.py
```

Acesse: **http://localhost:5001**

## 💰 Custo Operacional

| Componente | Plano | Custo |
|---|---|---|
| Gemini 2.5 Flash | Free Tier (15 req/min) | **R$ 0** |
| Groq (Llama 3.3) | Free Tier (High Speed) | **R$ 0** |
| ChromaDB | Local (em disco) | **R$ 0** |
| Microsoft Graph | Incluído no M365 | **R$ 0** |
| Flask | Open Source | **R$ 0** |

**Custo total: R$ 0/mês** 🎉 (Open Source & Free Tier Focused)
