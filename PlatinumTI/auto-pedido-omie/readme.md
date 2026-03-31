# 🛒 Automação de Pedidos de Venda - Omie Integration

![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python)
![ERP](https://img.shields.io/badge/ERP-Omie-orange)

## 🎯 O Problema de Negócio
Receber "espelhos de notas fiscais" em PDF e redigitar manualmente cada item no ERP Omie é um processo moroso e propenso a erros de digitação (SKU, Quantidade, Preço). Esse gargalo operacional atrasa a expedição e pode causar divergências de estoque.

## 💡 A Solução
Desenvolvi um pipeline de automação que realiza o parsing inteligente de PDFs, extraindo dados críticos (CNPJs, Itens, Valores) e gerando planilhas estruturadas prontas para importação ou integração direta com o ERP Omie. A solução elimina a redigitação de dados, garantindo 100% de precisão na transferência de informações.

## 🚀 Principais Funcionalidades
- **PDF Data Extraction**: Uso de Expressões Regulares (`re`) e `pypdf` para extrair dados complexos de espelhos de notas.
- **Detecção Inteligente de Itens**: Parsing do corpo da nota para identificar nomes de produtos, quantidades e valores unitários.
- **Normalização de CNPJs e Protocolos**: Limpeza e validação automática de campos cadastrais.
- **Integração Omie-Ready**: Geração de arquivos `.xlsx` formatados conforme o template padrão de pedidos de venda do Omie.
- **Log de Execução**: Monitoramento em tempo real de cada arquivo processado para auditoria e tratamento de erros.

## 🛠️ Tecnologias Utilizadas
- **Python 3**: Core da automação.
- **PyPDF**: Leitura e extração de texto de documentos PDF.
- **OpenPyXL**: Geração e manipulação de planilhas de pedidos.
- **Logging**: Sistema de rastreabilidade de processamento.

## 📂 Estrutura do Projeto
- `extract_pdf_data.py`: Módulo de inteligência para leitura de PDFs.
- `populate_excel.py`: Motor de geração de planilhas formatadas.
- `main_automation_script.py`: Orquestrador de fluxo de trabalho.
- `Notas/`: Diretório de entrada para arquivos PDF.
- `Planilhas/`: Diretório de saída com os pedidos estruturados.

---
> [!TIP]
> Esta ferramenta faz parte de um ecossistema de automação voltado para a escala operacional, reduzindo drasticamente o tempo de processamento administrativo.
