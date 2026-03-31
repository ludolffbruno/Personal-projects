# 📊 Cálculo de Substituição Tributária (ICMS-ST) Platinum

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python)
![License](https://img.shields.io/badge/License-Proprietary-red)

## 🎯 O Problema de Negócio
O cálculo de Substituição Tributária (ST) no Brasil é um dos processos mais complexos da área fiscal, envolvendo variáveis como NCM, MVA, DIFAL, e alíquotas que variam conforme a origem e o destino da mercadoria. O processo manual é lento, suscetível a erros humanos e pode resultar em prejuízos financeiros ou passivos fiscais.

## 💡 A Solução
Este sistema automatiza o cálculo do custo final de aquisição de produtos, integrando uma base de dados robusta em Excel com uma interface gráfica intuitiva (GUI). Ele fornece resultados instantâneos e precisos, permitindo que a equipe de compras ou fiscal tome decisões baseadas em dados reais.

## 🚀 Principais Funcionalidades
- **Busca Global**: Filtragem instantânea por NCM ou descrição do produto.
- **Lógica Tributária Dinâmica**: Diferenciação automática entre itens com ST (`C/ST`) e sem ST (`S/ST`).
- **Configuração de Origem**: Suporte a alíquotas de 4%, 12% e 22% (RJ) conforme a origem nacional ou importada.
- **Interface GUI (Tkinter)**: Design moderno com tema `clam`, janelas maximizadas e tutorial integrado.
- **Integração Excel**: Sincronização em tempo real com `CALCULO_ST_DIFAL_COMPLETA.xlsx`.
- **Produtividade**: Botão de cópia rápida para o clipboard e validação rigorosa de entradas.

## 🛠️ Tecnologias Utilizadas
- **Python 3.12**: Core do sistema.
- **Tkinter**: Interface gráfica rica e responsiva.
- **OpenPyXL**: Manipulação avançada de matrizes de dados em Excel.
- **Pyperclip**: Gestão de área de transferência para agilizar processos.

## 📂 Estrutura do Projeto
- `Calc_ST_Platinum_ver15_OK.py`: Script principal com a lógica de negócio.
- `CALCULO_ST_DIFAL_COMPLETA.xlsx`: Banco de dados com parâmetros tributários.
- `Documentação_Técnica.txt`: Guia para desenvolvedores e manutenção.
- `Manual de Uso.txt`: Guia instrucional para o usuário final.

---
> [!NOTE]
> Este projeto foi desenvolvido para atender uma necessidade crítica de precisão fiscal, transformando regras complexas em uma ferramenta de "um clique".
