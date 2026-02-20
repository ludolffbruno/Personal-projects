from pypdf import PdfReader
import re
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def extract_data_from_pdf(pdf_path):
    """
    Extrai dados específicos de um arquivo PDF de espelho de nota fiscal.

    Args:
        pdf_path (str): Caminho para o arquivo PDF

    Returns:
        dict: Dicionário com os dados extraídos ou None em caso de erro
    """
    try:
        reader = PdfReader(pdf_path)

        if len(reader.pages) == 0:
            logging.error(f"PDF {pdf_path} não contém páginas.")
            return None

        page = reader.pages[0]
        text = page.extract_text()

        if not text or len(text.strip()) == 0:
            logging.error(f"Não foi possível extrair texto do PDF {pdf_path}.")
            return None

        lines = text.splitlines()
        data = {}

        # Extrair CNPJ do Emitente
        cnpj_emitente_match = re.search(r"CNPJ\s+(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", text)
        if cnpj_emitente_match:
            data["cnpj_emitente"] = cnpj_emitente_match.group(1)

        # Extrair CNPJ do Cliente
        cnpj_cliente_match = re.search(r"CPF / CNPJ\s*([\d.]{2,}\.\d{3}\.\d{3}/\d{4}-\d{2})", text)
        if cnpj_cliente_match:
            data["cnpj_cliente"] = cnpj_cliente_match.group(1)

        # Extrair Número do Pedido
        pedido_match = re.search(r"Pedido:\s*(\d+)", text)
        if pedido_match:
            data["numero_pedido"] = pedido_match.group(1)
        else:
            logging.error("Número do pedido não encontrado no PDF.")
            return None

        # Extrair Número do Protocolo
        protocolo_match = re.search(r"Protocolo:\s*(\d+)", text)
        if protocolo_match:
            data["numero_protocolo"] = protocolo_match.group(1)

        # Extrair UF
        uf_match = re.search(r"\s([A-Z]{2})\s+INSCRIÇÃO ESTADUAL", text)
        if uf_match:
            data["uf"] = uf_match.group(1)

        # Extrair Valor Total dos Produtos
        valor_produtos_match = re.search(r"VALOR TOTAL DOS PRODUTOS\s+([\d.,]+)", text)
        if valor_produtos_match:
            try:
                valor_str = valor_produtos_match.group(1).replace(".", "").replace(",", ".")
                data["valor_total_produtos"] = float(valor_str)
            except ValueError:
                logging.error(f"Erro ao converter valor total dos produtos: {valor_produtos_match.group(1)}")

        # Extrair Valor Total da Nota
        valor_nota_match = re.search(r"VALOR TOTAL DA NOTA\s+([\d.,]+)", text)
        if valor_nota_match:
            try:
                valor_str = valor_nota_match.group(1).replace(".", "").replace(",", ".")
                data["valor_total_nota"] = float(valor_str)
            except ValueError:
                logging.error(f"Erro ao converter valor total da nota: {valor_nota_match.group(1)}")

        # === Extrair ALIQ ICMS da linha do item ===
        aliq_icms_encontrado = False
        for line in text.splitlines():
            if "000" in line and "UN" in line:
                logging.debug(f"Linha candidata para ALIQ ICMS: {line}")
                # Extrair todos os números no formato decimal, ex: 45.71, 0.00, 22.00
                matches = re.findall(r"\d+\.\d{2}", line)
                if matches:
                    try:
                        # Pela estrutura do seu PDF, ALIQ ICMS parece ser o 13º ou penúltimo número
                        # Ajuste esse índice conforme necessário com base nas observações
                        aliq_icms_str = matches[9]  # ou outro índice se necessário
                        data["aliq_icms"] = float(aliq_icms_str)
                        logging.info(f"ALIQ ICMS extraído da linha do item: {data['aliq_icms']}")
                        aliq_icms_encontrado = True
                        logging.debug(f"Valores decimais encontrados: {matches}")
                        break
                    except ValueError:
                        logging.warning(f"Erro ao converter ALIQ ICMS da linha: {aliq_icms_str}")
                        logging.debug(f"Valores decimais encontrados: {matches}")
                        break
        
        if not aliq_icms_encontrado:
            data["aliq_icms"] = 0.0
            logging.warning("Não foi possível extrair ALIQ ICMS da linha do item.")


        # ==========================
        # NOVA LÓGICA PARA PRODUTOS
        # ==========================
        produtos_extraidos = []
        for line in lines:
            if re.search(r'\d{4}\.\d{2}\.\d{2}(UNID\.?|UN)', line):
                # Substituir qualquer variação de UNID colada ao NCM por ' UN'
                line = re.sub(r'(\d{4}\.\d{2}\.\d{2})\s?(UNID\.?|UN)', r'\1 UN', line)
                partes = line.split()
                try:
                    idx_un = partes.index("UN")
                    if idx_un >= 2:
                        # Extrair campos com base nas posições relativas
                        descricao = " ".join(partes[2:idx_un])
                        quantidade_str = partes[idx_un + 1].replace(",", ".")
                        valor_unitario_str = partes[idx_un + 2].replace(",", "")

                        quantidade = float(quantidade_str)
                        valor_unitario = float(valor_unitario_str)

                        produtos_extraidos.append({
                            "nome": descricao.strip(), # No populate_excel, usamos 'nome'
                            "quantidade": quantidade,
                            "valor_unitario": valor_unitario
                        })
                except Exception as e:
                    logging.warning(f"Erro ao tentar extrair produto da linha: '{line}' - {e}")

        data["produtos"] = produtos_extraidos
        logging.info(f"{len(produtos_extraidos)} produtos extraídos com sucesso.")

        return data

    except FileNotFoundError:
        logging.error(f"Arquivo PDF não encontrado: {pdf_path}")
        return None
    except Exception as e:
        logging.error(f"Erro inesperado ao processar PDF {pdf_path}: {e}")
        return None

# Teste direto (opcional)
if __name__ == "__main__":
    test_pdf = "Notas/EspelhoNotaFiscal.pdf"
    extracted_data = extract_data_from_pdf(test_pdf)

    if extracted_data:
        from pprint import pprint
        pprint(extracted_data)
    else:
        print("Falha na extração de dados.")
