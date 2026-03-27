# Placeholders - Template de Corpo de E-mail

Template: `proposta_modelo_email_body_template.html`

## Destinatario
- `{{cliente_empresa}}`
- `{{cliente_contato_nome}}`

## Emitente
- `{{empresa_nome}}`
- `{{empresa_cnpj}}`
- `{{empresa_ie}}`
- `{{empresa_contato}}`
- `{{empresa_email}}`
- `{{empresa_dados_bancarios}}`

## Processo
- `{{processo_descricao}}`

## Itens (bloco repetivel)
- `{{#itens}} ... {{/itens}}`

Campos por item:
- `{{item_descricao}}`
- `{{item_marca}}`
- `{{item_valor_unitario_bruto}}`
- `{{item_quantidade}}`
- `{{item_valor_total_bruto}}`
- `{{item_icms}}`
- `{{item_ncm}}`
- `{{item_origem}}`
- `{{item_st}}`
- `{{item_pis}}`
- `{{item_cofins}}`
- `{{item_entrega}}`
- `{{item_frete}}`
- `{{item_faturamento}}`
- `{{item_validade_proposta}}`
- `{{item_aviso_validade}}`

## Fechamento
- `{{condicao_pagamento}}`
- `{{local_data}}`
- `{{assinatura_cargo}}`

