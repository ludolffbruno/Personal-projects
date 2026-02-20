import openpyxl
from datetime import datetime, timedelta
import logging
import os
import pandas as pd
import re
import locale 

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def carregar_mapeamento_produtos(mapping_path):
    try:
        df = pd.read_excel(mapping_path, dtype=str, usecols="A:B", header=0)
        df = df.dropna(subset=[df.columns[0], df.columns[1]])
        df[df.columns[0]] = df[df.columns[0]].str.upper().str.strip()
        df[df.columns[1]] = df[df.columns[1]].str.strip()
        mapping = dict(zip(df[df.columns[0]], df[df.columns[1]]))
        return mapping
    except Exception as e:
        logging.error(f"Erro ao carregar mapeamento de produtos: {e}")
        return {}


def populate_excel_sheet(excel_template_path, data, output_path):
    try:
        if not os.path.exists(excel_template_path):
            logging.error(f"Arquivo modelo não encontrado: {excel_template_path}")
            return False

        workbook = openpyxl.load_workbook(excel_template_path)

        if "Omie_Pedido_Venda" not in workbook.sheetnames:
            logging.error("Aba 'Omie_Pedido_Venda' não encontrada na planilha modelo.")
            return False

        sheet = workbook["Omie_Pedido_Venda"]

        if 'numero_pedido' not in data or not data['numero_pedido']:
            logging.error("Número do pedido é obrigatório mas não foi fornecido.")
            return False

        # 1. CNPJ do Cliente (D7)
        if 'cnpj_cliente' in data and data['cnpj_cliente']:
            sheet["D7"] = data["cnpj_cliente"]

        # 2. Previsão de Faturamento (E7)
        sheet["E7"] = (datetime.now() + timedelta(days=1)).strftime("%d/%m/%Y")

        # 3. Nº do Pedido do Cliente (K7)
        sheet["K7"] = data["numero_pedido"]

        # 4. Dados Adicionais (E12)
        logging.info(f"Valores recebidos: UF={data.get('uf')}, ALIQ={data.get('aliq_icms')}, NOTA={data.get('valor_total_nota')}, PRODUTOS={data.get('valor_total_produtos')}")

        mensagem = generate_additional_data_message(
            uf=data.get("uf", ""),
            aliq_icms=data.get("aliq_icms", 0.0),
            valor_total_nota=data.get("valor_total_nota", 0.0),
            valor_total_produtos=data.get("valor_total_produtos", 0.0),
            pedido=data.get("numero_pedido", ""),
            protocolo=data.get("numero_protocolo", "")
        )
        sheet["E12"] = mensagem

        # 5. Campos fixos
        sheet["C12"] = "Sim"
        sheet["H17"] = "1"
        sheet["I17"] = "CAIXA"
        sheet["J7"] = "Santander"
        sheet["F7"] = "Receita de Venda de Mercadoria"
        sheet["G7"] = "Para 60 dias"

        # 6. Preencher Produtos (D22 em diante)
        mapeamento_path = "Relação produtos EspelhoxOmie.xlsx"
        mapeamento = carregar_mapeamento_produtos(mapeamento_path)
        produtos = data.get("produtos", [])
        linha_base = 22

        for i, produto in enumerate(produtos):
            nome_original = str(produto.get("nome", "")).upper().strip()
            nome_original = re.sub(r"\b\d{4}\.\d{2}\.\d{2}\b", "", nome_original).strip()

            print(f"Procurando produto: '{nome_original}'")
            if nome_original in mapeamento:
                print(f"Produto '{nome_original}' encontrado no mapeamento.")
            else:
                print(f"Produto '{nome_original}' NÃO encontrado no mapeamento.")

            nome_map = mapeamento.get(nome_original)
            if not nome_map:
                logging.warning(f"Produto não mapeado: '{nome_original}'")
                continue  # Pula se não houver mapeamento

            linha = linha_base + i

            # Usar locale para conversão correta do valor no formato brasileiro
            valor_unitario_raw = str(produto.get("valor_unitario", "0")).strip()

            # Corrigir valores no formato americano, tipo "1,079.56"
            # Remove vírgula (milhar), mantém o ponto decimal
            try:
                valor_convertido_float = float(valor_unitario_raw.replace(",", ""))
            except ValueError:
                logging.warning(f"Erro ao converter valor: '{valor_unitario_raw}', usando 0.0")
                valor_convertido_float = 0.0

            sheet[f"G{linha}"] = valor_convertido_float
            print(f"Valor unitário convertido: {valor_convertido_float}")

            sheet[f"D{linha}"] = nome_map
            sheet[f"E{linha}"] = "UNICO"
            sheet[f"F{linha}"] = produto.get("quantidade", 0)
            
            # Preencher item do pedido sequencial em J{linha}
            sheet[f"K{linha}"] = i + 1


        # Preencher número do pedido na célula J22 (só na primeira linha)
        sheet["J22"] = data["numero_pedido"]

        # Salvar
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        workbook.save(output_path)
        logging.info(f"Planilha salva com sucesso: {output_path}")
        return True

    except Exception as e:
        logging.error(f"Erro inesperado: {e}")
        return False


def generate_additional_data_message(uf, aliq_icms, valor_total_nota, valor_total_produtos, pedido, protocolo):

    mensagem_procon = (
        "PROCON RJ - TEL 151 RUA DA AJUDA 5 SUBSOLO CENTRAL DO BRASIL CEP 20040-000/ "
        "PRAÇA CRISTIANO OTTONI S/N SUBSOLO CENTRO DO RJ CEP 20221-250/ "
        "ALOALERJ 0800-0220008 PALACIO TIRANDENTES - RUA PRIMEIRO DE MARÇO S/N PRAÇA XV RJ CEP 20010-090  / "
        f"PEDIDO {pedido} / PROTOCOLO {protocolo} / A/C MUCIO 2121-3885"
    )

    mensagem_st_aliq_zero = (
        "ICMS RETIDO ANTERIORMENTE POR SUBSTITUIÇÃO TRIBUTÁRIA/MERCADORIA SUJEITA AO REGIME DE "
        "SUBSTITUICAO TRIBUTARIA CONF ANEXO I DO LIVRO II DO RICMS/RJ E PROTOCOLOS/ " +
        mensagem_procon
    )

    mensagem_st_valores_diferentes = (
        "MERCADORIA SUJEITA AO REGIME DE SUBSTITUICAO TRIBUTARIA CONF ANEXO I DO LIVRO II DO RICMS/RJ "
        "E PROTOCOLOS ICMS ST REFERENTE A DIFERENCIAL DE ALIQUOTA / " + mensagem_procon
    )

    # Lógica:
    if uf == "RJ":
        if aliq_icms == 0:
            return mensagem_st_aliq_zero
        else:
            return mensagem_procon
    else:
        if abs(valor_total_nota - valor_total_produtos) > 0.01: 
            return mensagem_st_valores_diferentes
        else:
            return mensagem_procon
