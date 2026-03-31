# 🌬️🦾 Platinum NF Automation (Outlook Agent)

## 💡 Problema
O processamento manual de Notas Fiscais Eletrônicas (NF-e) que chegam por e-mail é um trabalho repetitivo e sujeito a falhas. Separar anexos, validar se a nota pertence ao cliente correto e organizar os arquivos para faturamento consome horas produtivas da equipe administrativa.

## 🚀 Solução
Um agente de automação robusto que utiliza a **Microsoft Graph API** para monitorar pastas específicas do Outlook. O sistema realiza a extração inteligente de dados de PDFs (Número da NF, Pedido, Protocolo), valida a origem dos documentos e automatiza toda a organização de pastas e registros de processamento.

## 🛠️ Tecnologias
- **Python 3**: Core do sistema.
- **Microsoft Graph API**: Conexão segura e moderna com o ecossistema Office 365.
- **Tkinter**: Interface gráfica (GUI) para controle e acompanhamento em tempo real.
- **PyPDF / RE**: Extração e tratamento de dados complexos de documentos fiscais.

## ⚙️ Como executar
1. Configure as credenciais do seu App no Azure no arquivo `.env`.
2. Certifique-se de que o **pdftotext** está configurado no seu PATH.
3. Instale as dependências: `pip install requests PyMuPDF python-dotenv`
4. Execute o script: `python nf_automation_gui.py`

## 📈 Resultado
- **100% de Precisão**: Validação dupla (corpo do e-mail + PDF) para garantir o cliente correto.
- **Organização Instantânea**: Notas categorizadas e protocolos extraídos automaticamente.
- **Agilidade**: Redução drástica no tempo entre o recebimento do e-mail e a disponibilidade para faturamento.

---
> [!TIP]
> **Prova de Execução**: Veja o log em tempo real e a extração automática de protocolos na interface.
> ![Automation-Demo](https://via.placeholder.com/800x450.png?text=Inserir+GIF+da+Interface+de+Automação)
