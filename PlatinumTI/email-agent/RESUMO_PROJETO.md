# Projeto: Vendedor Imortal — Agente de Vendas com IA

## Resumo para o Gestor (Visão Estratégica)

O **Vendedor Imortal** é uma solução de automação inteligente projetada para otimizar o ciclo de vendas B2B, eliminando o trabalho manual de triagem e primeira resposta a cotações.

### 1. Objetivos e Benefícios
*   **Aumento de Eficiência**: Reduz o tempo de resposta a novos leads em até 90%, garantindo que pedidos de cotação nunca sejam ignorados.
*   **Qualidade e Personalização**: Utiliza **Gemini 1.5 Flash** para aprender o estilo de escrita do vendedor através de e-mails enviados anteriormente, gerando respostas que parecem humanas e autênticas.
*   **Escalabilidade**: Processa dezenas de e-mails, anexos (PDF, Word, Imagens) e listas de itens automaticamente, 24/7.

### 2. Destaques Técnicos
*   **Inteligência Artificial Multimodal**: Processamento de texto e visão para extrair itens de pedidos até em fotos, PDFs ou **tabelas complexas em arquivos Word**.
*   **Caçador de Imagens**: Identifica automaticamente imagens coladas direto no corpo do e-mail (Base64) que outros sistemas ignorariam.
*   **Failover de Cota**: Sistema redundante que alterna entre Gemini e Groq Vision para garantir que o agente nunca "fique cego" por falta de cota de API.
*   **Fluxo de Trabalho Determinístico**: Implementado via **LangGraph**, garantindo que a IA siga regras de negócio rígidas (só responder se for cotação).

### 3. Impacto no Negócio
Com o agente cuidando da triagem e preparando o rascunho, o vendedor foca exclusivamente na **negociação e fechamento**, aumentando a taxa de conversão do time comercial.

---

## Resumo para o Usuário (Guia do Operador)

O **Vendedor Imortal** é o seu assistente pessoal no Outlook que trabalha enquanto você foca no que importa: vender.

### Como ele te ajuda na rotina:
1.  **Triagem Automática**: Ele lê sua Caixa de Entrada e identifica o que é um pedido de cotação real e o que é apenas e-mail administrativo ou spam.
2.  **Preparação de Respostas**: Para cada cotação encontrada, ele cria um rascunho (Draft) pronto para você no seu Outlook. O rascunho já vem com os itens listados e usa o seu jeito de escrever.
3.  **Leitura de Anexos**: Não precisa abrir PDFs ou fotos para saber o que o cliente quer. O agente já extrai os itens para você.

### Como Usar:
*   **Dashboard**: Acompanhe em tempo real o que o agente está fazendo.
*   **Verificar Agora**: Clique neste botão para processar e-mails imediatamente.
*   **Sincronizar Memória**: Use este botão quando quiser que o agente "aprenda" com as suas respostas recentes. Ele vai ler seus e-mails enviados para se ajustar ao seu tom de voz.
*   **Categoria "Platinum Sales"**: Todo e-mail que o agente identificar como cotação valiosa será marcado automaticamente no seu Outlook com uma categoria dourada para fácil visualização.

**Seu trabalho**: Abrir a pasta de Rascunhos, preencher os preços onde estiver "[VALOR A CONFIRMAR]" e clicar em enviar!
