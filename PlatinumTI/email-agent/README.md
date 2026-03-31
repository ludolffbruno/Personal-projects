# 🌬️🦾 Vendedor Imortal — Outlook AI Agent

![Python](https://img.shields.io/badge/Python-3.11%2B-blue?logo=python)
![Gemini](https://img.shields.io/badge/Gemini-8E75C2?style=flat&logo=googlegemini&logoColor=white)
![LangGraph](https://img.shields.io/badge/LangGraph-000?style=flat&logo=langchain&logoColor=white)

## 🎯 O Problema de Negócio | Business Problem
Vendedores de alta performance perdem horas valiosas triando e-mails, lendo anexos complexos e respondendo cotações repetitivas. O desafio é automatizar a triagem e o rascunho de respostas sem perder o "toque humano" e a personalização necessária para fechar vendas.

## 💡 A Solução | The Solution
Um agente de IA sofisticado que vive dentro do Microsoft Outlook. Ele age como um "braço direito", lendo pedidos de orçamento (mesmo em formatos complexos como imagens ou tabelas em PDFs), comparando com o histórico do vendedor para manter o tom de voz e gerando rascunhos automáticos. O vendedor apenas revisa e clica em enviar.

## 🚀 Principais Funcionalidades | Key Features
- **Triagem Cirúrgica**: Distingue cotações reais de spam/newsletters com alta precisão.
- **Visão Multimodal Avançada**: OCR integrado para ler pedidos em PDFs, Word, Imagens e até prints colados no corpo do e-mail.
- **Memória de Estilo (RAG)**: Utiliza **ChromaDB** para aprender o padrão de escrita do usuário e replicá-lo nos rascunhos.
- **Failover Inteligente**: Arquitetura resiliente que alterna entre **Gemini 2.5 Flash** e **Groq Llama 3.2** para garantir disponibilidade.
- **Drafts Automáticos**: Salva rascunhos diretamente no Outlook via **MS Graph API**, preservando a thread original.
- **Dashboard Premium**: Interface Dark Mode (Flask) para monitorar o "raciocínio" da IA e logs de extração.

## 🛠️ Tecnologias Utilizadas | Tech Stack
- **Python & LangGraph**: Orquestração de agentes inteligentes.
- **Microsoft Graph API**: Integração profunda com o ecossistema M365.
- **ChromaDB**: Banco de dados vetorial para memória de longo prazo.
- **Flask**: Dashboard de monitoramento e controle.
- **Gemini / Groq**: Motores de IA generativa de última geração.

## 💰 Custo Operacional: R$ 0/mês
Projetado para rodar inteiramente em **Free Tiers** e soluções **Open Source**, permitindo alta performance sem custos fixos de API.

---
> [!IMPORTANT]
> Este projeto demonstra a aplicação prática de IA Generativa para ganho de escala em processos de vendas B2B.
