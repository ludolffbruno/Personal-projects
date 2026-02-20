import os
from extract_pdf_data import extract_data_from_pdf
from populate_excel import populate_excel_sheet

def run_automation():
    pdf_input_dir = "Notas"
    excel_output_dir = "Planilhas"
    excel_template_path = os.path.join(excel_output_dir, "Planilha_Modelo_Pedido_de_Venda.xlsx")

    # Criar diretórios se não existirem
    os.makedirs(pdf_input_dir, exist_ok=True)
    os.makedirs(excel_output_dir, exist_ok=True)

    # Listar todos os arquivos PDF na pasta de entrada
    pdf_files = [f for f in os.listdir(pdf_input_dir) if f.lower().endswith(".pdf")]

    if not pdf_files:
        print(f"Nenhum arquivo PDF encontrado na pasta: {pdf_input_dir}")
        return

    for pdf_file_name in pdf_files:
        pdf_path = os.path.join(pdf_input_dir, pdf_file_name)
        print(f"\nProcessando PDF: {pdf_file_name}")
        
        try:
            # Extrair dados do PDF
            extracted_data = extract_data_from_pdf(pdf_path)
            
            if not extracted_data or 'numero_pedido' not in extracted_data:
                print(f"Erro: Não foi possível extrair dados essenciais do PDF {pdf_file_name}. Pulando este arquivo.")
                continue

            # Gerar nome inteligente para a planilha de saída
            output_excel_name = f"Pedido_{extracted_data['numero_pedido']}.xlsx"
            output_excel_path = os.path.join(excel_output_dir, output_excel_name)

            # Preencher a planilha Excel
            populate_excel_sheet(excel_template_path, extracted_data, output_excel_path)
            print(f"Planilha gerada com sucesso: {output_excel_name}")

        except Exception as e:
            print(f"Erro ao processar o arquivo {pdf_file_name}: {e}")

    print("\nAutomação concluída para todos os PDFs.")

if __name__ == "__main__":
    run_automation()


