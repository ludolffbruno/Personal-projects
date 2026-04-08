# 🧮 Calculadora ICMS-ST e DIFAL (Fase 2 - Saídas)

![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python)
![Tributário](https://img.shields.io/badge/Fiscal-ICMS--ST-green)
![DIFAL](https://img.shields.io/badge/Fiscal-DIFAL-yellow)

## 🎯 O Problema de Negócio
O cálculo de impostos estaduais em vendas interestaduais (ICMS-ST e DIFAL) envolve regras fiscais complexas, como Diferencial de Alíquotas com Base Dupla e Base Única (Inclusa), além da Margem de Valor Agregado (MVA) para cenários de revenda. Fazer esses cálculos manualmente ou usando apenas planilhas soltas gera um alto risco de multas fiscais e precificação incorreta dos produtos nas operações de saída.

## 💡 A Solução
Desenvolvi um software desktop interativo que automatiza o cálculo de Substituição Tributária (ST) e DIFAL para operações de saída. O sistema cruza os dados do produto (NCM, Origem) com Matrizes Fiscais atualizadas em Excel, determinando instantaneamente a modalidade correta (Revenda, Base Dupla ou Base Única) e exibindo a memória de cálculo detalhada.

## 🚀 Principais Funcionalidades
- **Cálculo de ICMS-ST (Revenda)**: Aplicação automática da Margem de Valor Agregado (MVA) baseada na tabela oficial.
- **DIFAL Base Dupla (CFOP 6403)**: Cálculo "grossed up" para vendas a Consumidores Finais que possuem Inscrição Estadual (Ribeiro/Contribuinte).
- **DIFAL Base Única/Inclusa (CFOP 6108)**: Embutimento do ICMS de destino para Consumidores Finais Não Contribuintes.
- **Leitura Inteligente de Matrizes Excel**: Sincronização em tempo real com `MATRIZ_ST_SAIDA.xlsx` usando a biblioteca Pandas.
- **Interface Gráfica Interativa**: Inputs simplificados para o usuário selecionar UF de destino, NCM, tipo de aquisição e perfil do cliente (Consumidor Final, Contribuinte).

## 🛠️ Tecnologias Utilizadas
- **Python 3**: Linguagem principal do motor de cálculo.
- **CustomTkinter / Tkinter**: Desenvolvimento da Interface Gráfica de Usuário (GUI) fluida e amigável.
- **Pandas & OpenPyXL**: Manipulação e extração de parâmetros das planilhas de matriz fiscal (alíquotas e MVA).

## 📂 Estrutura do Projeto
- `Calc_ST_Fase2_Saidas_v3.py`: Script principal rodando o motor de cálculos e a interface da fase 2.
- `MATRIZ_ST_SAIDA.xlsx`: Banco de dados contendo o de-para de NCMs, regras de MVA e perfis tributários.
- `Calculadora Base dupla ICMS - CALCULO DE ST.xlsx`: Simuladores fiscais validados pela contabilidade e usados como prova real.
- `Fluxograma_Fase2_COLE_AQUI.mmd`: Arquitetura visual das lógicas de decisão da aplicação (Mermaid).

---
> [!TIP]
> Este projeto representa a **Fase 2** de um ecossistema fiscal, adaptando o motor MVA previamente criado para suportar com precisão cirúrgica operações faturadas contra Consumidores Finais (B2B e B2C).
