# Automação de Pedidos de Venda para Omie

Este projeto automatiza a criação de planilhas de pedidos de venda no formato Excel, com base em arquivos PDF de espelho de nota fiscal, para importação no sistema ERP Omie.

## Estrutura de Pastas

O projeto deve ser organizado da seguinte forma:

```
Auto-pedido-omie-manus/
├── gui.py  <-- NOVO! Interface gráfica para o usuário
├── main_automation_script.py
├── extract_pdf_data.py
├── populate_excel.py
├── README.md  <-- Este arquivo
├── Notas/
│   ├── EspelhoNotaFiscal1.pdf
│   ├── EspelhoNotaFiscal2.pdf
│   └── ... (outros PDFs)
└── Planilhas/
├── Planilha_Modelo_Pedido_de_Venda.xlsx
├── Pedido_5500520611.xlsx (gerado automaticamente)
└── ... (outras planilhas geradas)
```

## Funcionalidades

O programa agora é executado via interface gráfica e realiza as seguintes operações:

1.  **Processamento em Lote**: Processa automaticamente todos os arquivos PDF encontrados na pasta `Notas/`.
2.  **Extração de Dados do PDF**: Lê cada arquivo PDF de espelho de nota fiscal e extrai informações cruciais como CNPJ, número do pedido, número do protocolo, UF do cliente, valor total dos produtos e valor total da nota.
3.  **Preenchimento da Planilha Excel**: Utiliza o modelo de planilha Excel (`Planilhas/Planilha_Modelo_Pedido_de_Venda.xlsx`) e preenche automaticamente os campos. A formatação de números agora segue o padrão brasileiro, com ponto como separador de milhar e vírgula como separador decimal.
4.  **Interface Gráfica**: Oferece uma interface simples para iniciar a automação e visualizar o progresso em tempo real, sem a necessidade de janelas de pop-up.

## Tecnologias Utilizadas

* **Python 3.13**: Linguagem de programação.
* **tkinter**: Biblioteca padrão do Python para criação da interface gráfica.
* **threading**: Usado para executar a automação em segundo plano, mantendo a interface responsiva.
* **PyMuPDF (fitz)**: Leitura precisa de PDFs.
* **OpenPyXL**: Leitura e escrita em planilhas `.xlsx`.
* **re (regex)**: Extração de padrões no texto.
* **os**: Manipulação de arquivos e diretórios.
* **logging**: Geração de logs detalhados.

## Execução do Programa

Para usar a automação, siga os passos abaixo:

1.  Coloque os arquivos PDF de espelho de nota fiscal na pasta `Notas/`.
2.  Execute o script da interface gráfica via terminal:
    ```bash
    python gui.py
    ```
3.  Na janela que se abrirá, clique no botão "Iniciar Automação".
4.  O progresso será exibido na área de log na parte inferior da tela.

## Logs e Monitoramento

O script gera logs detalhados que são exibidos em tempo real na área de texto da interface, incluindo:

* Informações sobre cada PDF processado.
* Dados extraídos de cada arquivo.
* Erros encontrados e suas possíveis causas.
* Confirmação de planilhas geradas com sucesso.

## Solução de Problemas Comuns

### "Nenhum arquivo PDF encontrado"
* Verifique se os PDFs estão na pasta `Notas/`.
* Certifique-se de que os arquivos têm a extensão `.pdf`.

### "Arquivo modelo não encontrado"
* Verifique se `Planilha_Modelo_Pedido_de_Venda.xlsx` está na pasta `Planilhas/`.
* Certifique-se de que o nome do arquivo está correto.

### "Número do pedido não encontrado"
* Verifique se o PDF contém o texto "Pedido:" seguido de um número.
* O PDF pode estar corrompido ou ter um formato diferente do esperado.

### "Erro ao salvar planilha"
* Verifique se você tem permissão de escrita na pasta `Planilhas/`.
* Certifique-se de que não há uma planilha com o mesmo nome já aberta no Excel.

---

**Autor**: Bruno Ludolff
**Data**: Agosto de 2025
**Versão**: 2.0 - Processamento em lote com tratamento de erros

