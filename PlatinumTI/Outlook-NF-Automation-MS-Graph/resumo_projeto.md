# Resumo do Projeto: Automação E-mail Alex

Este sistema moderniza e automatiza o fluxo de recebimento, extração e organização de Notas Fiscais Eletrônicas (NF-e) diretamente da conta de e-mail do usuário.

## 🚀 Funcionalidades Principais

1.  **Monitoramento Digital**: Utiliza a API Microsoft Graph para monitorar a pasta `#NFE PLATINUM` no Outlook sem precisar do Outlook Desktop aberto.
2.  **Extração de Dados Inteligente**: Analisa anexos PDF para extrair automaticamente **Número da NF**, **Pedido** e **Protocolo**.
3.  **Organização Automática**:
    *   Cria pastas dentro do projeto: `Notas-Salvas-xml-email\[Dados da NF]`.
    *   Cria uma pasta `Carregadas ao portal` para controle manual do usuário.
4.  **Gestão de Fluxo**:
    *   Marca e-mails como **lidos** após o processamento.
    *   Mantém um rascunho vivo no Outlook (**NOTAS CARREGADAS NO PORTAL**) com a lista detalhada de todas as NFs processadas no dia.
5.  **Interface Gráfica (GUI V2)**:
    *   Painel moderno com log em tempo real.
    *   Botões de **Iniciar/Pausar** sincronização.
    *   Botão de atalho para abrir a pasta de notas.
    *   Fluxo de autenticação simplificado direto na tela.

## 🏗️ Estrutura Final do Projeto

O projeto foi consolidado em uma versão estável que pode ser distribuída como um único arquivo executável.

| Componente | Função |
| :--- | :--- |
| **`automacao-email-alex.exe`** | Executável principal (unidade final de uso). |
| `nf_automation_gui.py` | Código da interface gráfica. |
| `graph_email_monitor.py` | Motor de conexão com Microsoft Graph e lógica de arquivos. |

## 📦 Entrega
*   O programa final está localizado na pasta `dist` após o build.
*   Toda a configuração de chaves do Azure (Client ID/Secret) já está injetada no motor do sistema.

> [!NOTE]
> Para o funcionamento correto da leitura de PDFs, o software **pdftotext** deve estar instalado e configurado no PATH do Windows.
