# Manual do Usuário - Sistema "Cálculo ST/DIFAL" Expandido

## Introdução

O Sistema "Cálculo ST/DIFAL" Expandido é uma ferramenta desenvolvida para auxiliar na gestão fiscal de operações de compra e venda, considerando os impostos ST (Substituição Tributária) e DIFAL (Diferencial de Alíquota) aplicáveis no Brasil. Esta versão expandida inclui dois modos de operação (Compra e Venda) e um sistema básico de usuários.

## Requisitos do Sistema

- Python 3.6 ou superior
- Bibliotecas: tkinter, openpyxl, pyperclip
- Planilha Excel: CALCULO_ST_DIFAL_COMPLETA.xlsx (deve estar no mesmo diretório do programa)

## Instalação

1. Certifique-se de que o Python está instalado em seu computador
2. Instale as bibliotecas necessárias:
   ```
   pip install openpyxl pyperclip
   ```
3. Coloque os arquivos `Calc_ST_Platinum_ver16_Expandido.py` e `CALCULO_ST_DIFAL_COMPLETA.xlsx` no mesmo diretório
4. Execute o programa:
   ```
   python Calc_ST_Platinum_ver16_Expandido.py
   ```

## Funcionalidades Principais

### Sistema de Usuários
- **Comprador**: Acesso apenas ao Modo Compra
- **Vendedor**: Acesso apenas ao Modo Venda

### Modo Compra
- Cálculo do custo final de produtos considerando impostos (ST, DIFAL, ICMS)
- Seleção do estado de origem da compra
- Configuração de ST/DIFAL com base no produto e estado

### Modo Venda
- Cálculo do preço de venda sugerido com base no custo final
- Configuração de margem de lucro (%)
- Seleção do estado de destino da venda
- Ajuste automático dos impostos aplicáveis ao estado de destino

## Guia de Uso

### Para Compradores

1. Selecione o tipo de usuário "Comprador"
2. Verifique se o "Modo Compra" está selecionado
3. Busque o produto por NCM ou descrição
4. Selecione o produto na lista de resultados
5. Escolha o estado de origem da compra
6. Se o produto for "C/ST", configure se a ST já está inclusa na compra
7. Se necessário, selecione a origem do ICMS (0%, 4%, 12% ou 22%)
8. Digite o custo unitário do produto
9. Clique em "Calcular" para obter o custo final
10. Use o botão "Copiar" para copiar o resultado para a área de transferência

### Para Vendedores

1. Selecione o tipo de usuário "Vendedor"
2. Verifique se o "Modo Venda" está selecionado
3. Busque o produto por NCM ou descrição
4. Selecione o produto na lista de resultados
5. Escolha o estado de destino da venda
6. Digite o custo final do produto (já com impostos de compra)
7. Digite a margem de lucro desejada (%)
8. Clique em "Calcular" para obter o preço de venda sugerido
9. Use o botão "Copiar" para copiar o resultado para a área de transferência

## Detalhes Técnicos

### Cálculos Fiscais

#### Modo Compra
- **Custo Final** = Custo Unitário + ICMS + DIFAL (se aplicável) + ST (se aplicável)
- **DIFAL** = (Alíquota Interna do Estado de Destino - Alíquota Interestadual) × Custo Unitário

#### Modo Venda
- **Preço de Venda** = Custo Final / (1 - Margem de Lucro/100 - Alíquota ICMS/100)
- A alíquota de ICMS varia conforme o estado de destino

### Alíquotas de ICMS
- **Internas**: Variam por estado (ex.: 18% no RJ)
- **Interestaduais**: 
  - 12% para operações de Sul/Sudeste para Sul/Sudeste (exceto ES)
  - 7% para operações de Sul/Sudeste para Norte/Nordeste/Centro-Oeste/ES
  - 12% para operações entre estados do Norte/Nordeste/Centro-Oeste/ES

## Suporte e Manutenção

Para atualizar as alíquotas ou regras fiscais, você pode:
1. Editar diretamente a planilha Excel CALCULO_ST_DIFAL_COMPLETA.xlsx
2. Atualizar os dicionários de alíquotas no código Python (para alíquotas internas e interestaduais)

## Limitações Conhecidas

- O sistema não implementa autenticação com senhas
- Algumas regras fiscais específicas podem requerer ajustes manuais
- A planilha Excel deve ser mantida no mesmo formato para garantir a compatibilidade

## Contato

Para suporte ou dúvidas sobre o sistema, entre em contato com o desenvolvedor.
