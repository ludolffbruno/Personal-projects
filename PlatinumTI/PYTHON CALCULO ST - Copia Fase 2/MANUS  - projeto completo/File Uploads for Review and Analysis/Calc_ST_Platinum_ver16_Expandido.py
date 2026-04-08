# ------------------------------------------------------------
# IMPORTAÇÕES NECESSÁRIAS
# ------------------------------------------------------------
import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import pyperclip

# Função para validar entrada no "Custo Unitário" e "Margem de Lucro"
def validar_numero(P):
    if P == "" or P.replace(",", "").replace(".", "").isdigit():
        return True
    return False

# Variáveis globais
dados = []
dados_planilha = []  # Variável para armazenar os dados completos da planilha
estados_brasileiros = [
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", 
    "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", 
    "RS", "RO", "RR", "SC", "SP", "SE", "TO"
]

# Dicionário de alíquotas internas por estado
aliquotas_internas = {
    "AC": 17, "AL": 17, "AP": 18, "AM": 18, "BA": 18, "CE": 18, 
    "DF": 18, "ES": 17, "GO": 17, "MA": 18, "MT": 17, "MS": 17, 
    "MG": 18, "PA": 17, "PB": 18, "PR": 18, "PE": 18, "PI": 18, 
    "RJ": 18, "RN": 18, "RS": 17, "RO": 17.5, "RR": 17, "SC": 17, 
    "SP": 18, "SE": 18, "TO": 18
}

# Dicionário de alíquotas interestaduais
# Formato: origem -> destino -> alíquota
aliquotas_interestaduais = {
    # Estados do Sul e Sudeste (exceto ES)
    "SP": {"SP": 18, "RJ": 12, "MG": 12, "ES": 12, "RS": 12, "SC": 12, "PR": 12,
           "AC": 7, "AL": 7, "AP": 7, "AM": 7, "BA": 7, "CE": 7, "DF": 7, "GO": 7,
           "MA": 7, "MT": 7, "MS": 7, "PA": 7, "PB": 7, "PE": 7, "PI": 7, "RN": 7,
           "RO": 7, "RR": 7, "SE": 7, "TO": 7},
    "RJ": {"SP": 12, "RJ": 18, "MG": 12, "ES": 12, "RS": 12, "SC": 12, "PR": 12,
           "AC": 7, "AL": 7, "AP": 7, "AM": 7, "BA": 7, "CE": 7, "DF": 7, "GO": 7,
           "MA": 7, "MT": 7, "MS": 7, "PA": 7, "PB": 7, "PE": 7, "PI": 7, "RN": 7,
           "RO": 7, "RR": 7, "SE": 7, "TO": 7},
    "MG": {"SP": 12, "RJ": 12, "MG": 18, "ES": 12, "RS": 12, "SC": 12, "PR": 12,
           "AC": 7, "AL": 7, "AP": 7, "AM": 7, "BA": 7, "CE": 7, "DF": 7, "GO": 7,
           "MA": 7, "MT": 7, "MS": 7, "PA": 7, "PB": 7, "PE": 7, "PI": 7, "RN": 7,
           "RO": 7, "RR": 7, "SE": 7, "TO": 7},
    "RS": {"SP": 12, "RJ": 12, "MG": 12, "ES": 12, "RS": 17, "SC": 12, "PR": 12,
           "AC": 7, "AL": 7, "AP": 7, "AM": 7, "BA": 7, "CE": 7, "DF": 7, "GO": 7,
           "MA": 7, "MT": 7, "MS": 7, "PA": 7, "PB": 7, "PE": 7, "PI": 7, "RN": 7,
           "RO": 7, "RR": 7, "SE": 7, "TO": 7},
    "SC": {"SP": 12, "RJ": 12, "MG": 12, "ES": 12, "RS": 12, "SC": 17, "PR": 12,
           "AC": 7, "AL": 7, "AP": 7, "AM": 7, "BA": 7, "CE": 7, "DF": 7, "GO": 7,
           "MA": 7, "MT": 7, "MS": 7, "PA": 7, "PB": 7, "PE": 7, "PI": 7, "RN": 7,
           "RO": 7, "RR": 7, "SE": 7, "TO": 7},
    "PR": {"SP": 12, "RJ": 12, "MG": 12, "ES": 12, "RS": 12, "SC": 12, "PR": 18,
           "AC": 7, "AL": 7, "AP": 7, "AM": 7, "BA": 7, "CE": 7, "DF": 7, "GO": 7,
           "MA": 7, "MT": 7, "MS": 7, "PA": 7, "PB": 7, "PE": 7, "PI": 7, "RN": 7,
           "RO": 7, "RR": 7, "SE": 7, "TO": 7},
    
    # Demais estados
    "ES": {"SP": 12, "RJ": 12, "MG": 12, "ES": 17, "RS": 12, "SC": 12, "PR": 12,
           "AC": 7, "AL": 7, "AP": 7, "AM": 7, "BA": 7, "CE": 7, "DF": 7, "GO": 7,
           "MA": 7, "MT": 7, "MS": 7, "PA": 7, "PB": 7, "PE": 7, "PI": 7, "RN": 7,
           "RO": 7, "RR": 7, "SE": 7, "TO": 7},
    "AC": {"SP": 7, "RJ": 7, "MG": 7, "ES": 7, "RS": 7, "SC": 7, "PR": 7,
           "AC": 17, "AL": 12, "AP": 12, "AM": 12, "BA": 12, "CE": 12, "DF": 12, "GO": 12,
           "MA": 12, "MT": 12, "MS": 12, "PA": 12, "PB": 12, "PE": 12, "PI": 12, "RN": 12,
           "RO": 12, "RR": 12, "SE": 12, "TO": 12},
    # Padrão para os demais estados do Norte, Nordeste e Centro-Oeste
    "OUTROS": {"SP": 7, "RJ": 7, "MG": 7, "ES": 7, "RS": 7, "SC": 7, "PR": 7,
               "AC": 12, "AL": 12, "AP": 12, "AM": 12, "BA": 12, "CE": 12, "DF": 12, "GO": 12,
               "MA": 12, "MT": 12, "MS": 12, "PA": 12, "PB": 12, "PE": 12, "PI": 12, "RN": 12,
               "RO": 12, "RR": 12, "SE": 12, "TO": 12}
}

# Função para obter a alíquota interestadual
def obter_aliquota_interestadual(origem, destino):
    if origem in aliquotas_interestaduais:
        return aliquotas_interestaduais[origem][destino] / 100
    else:
        # Usar o padrão para estados não listados explicitamente
        return aliquotas_interestaduais["OUTROS"][destino] / 100

# Função para obter a alíquota interna do estado
def obter_aliquota_interna(estado):
    return aliquotas_internas[estado] / 100

# Função para carregar os dados da planilha
def carregar_dados():
    global dados, dados_planilha
    try:
        workbook = openpyxl.load_workbook("CALCULO_ST_DIFAL_COMPLETA.xlsx", data_only=True)
        sheet = workbook["COMPLETA"]
        dados.clear()
        dados_planilha.clear()
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True):
            ncm = str(row[0])
            produto = row[1]
            st_difal = row[2]
            aquisicao = row[3]  # Coluna D
            origem = row[4]     # Coluna E
            if ncm and produto:
                dados.append({"NCM": ncm, "Produto": produto, "ST / DIFAL": st_difal, "Aquisição": aquisicao, "Origem": origem})
                dados_planilha.append(row)
    except FileNotFoundError:
        resumo_var.set("Arquivo CALCULO_ST_DIFAL_COMPLETA.xlsx não encontrado!")
    except Exception as e:
        resumo_var.set(f"Erro ao carregar dados: {str(e)}")

# Função para filtrar os dados
def filtrar_dados(event=None):
    global resultados_tree
    consulta = busca_var.get().lower().replace(".", "")
    resultados_tree.delete(*resultados_tree.get_children())
    for item in dados:
        ncm = item["NCM"].replace(".", "").lower()
        produto = item["Produto"].lower()
        if consulta in ncm or consulta in produto:
            resultados_tree.insert("", "end", values=(item["NCM"], item["Produto"]))

# Função para exibir detalhes do item selecionado e atualizar campos dependentes
def selecionar_item(event):
    global resultados_tree, dados_planilha
    item_selecionado = resultados_tree.focus()
    if item_selecionado:
        valores = resultados_tree.item(item_selecionado, "values")
        ncm_var.set(valores[0])  # Preencher o NCM
        resumo_var.set(f"NCM {valores[0]} | Produto: {valores[1]}")
        
        # Usar os dados já carregados em dados_planilha
        ncm_input = valores[0].replace(".", "")
        for row in dados_planilha:
            ncm_planilha = str(row[0]).replace(".", "")
            if ncm_planilha == ncm_input:
                st_difal = row[2] or "N/A"  # Coluna C (índice 2)
                st_difal_var.set(str(st_difal))
                atualizar_campos_dependentes()
                break
        else:
            st_difal_var.set("N/A")
            atualizar_campos_dependentes()

# Função para atualizar os campos dependentes com base em ST / DIFAL
def atualizar_campos_dependentes():
    # Verificar o modo atual
    if modo_var.get() == "Compra":
        if st_difal_var.get() == "C/ST":
            st_inclusa_checkbutton.config(state="normal")
            if st_inclusa_var.get():
                origem_icms_combobox.config(state="disabled")
            else:
                origem_icms_combobox.config(state="normal")  # Desbloqueado para seleção do usuário
        else:  # S/ST ou outro valor
            st_inclusa_var.set(False)  # Desmarcar o checkbox
            st_inclusa_checkbutton.config(state="disabled")
            origem_icms_combobox.config(state="normal")  # Habilitar Origem ICMS para S/ST
    else:  # Modo Venda
        # No modo venda, esses campos não são relevantes da mesma forma
        st_inclusa_var.set(False)
        st_inclusa_checkbutton.config(state="disabled")
        origem_icms_combobox.config(state="disabled")

# Função para atualizar Origem ICMS com base em ST Inclusa na compra?
def atualizar_origem_icms(*args):
    if modo_var.get() == "Compra":
        if st_difal_var.get() == "C/ST":
            if st_inclusa_var.get():  # Se o checkbox está marcado
                origem_icms_var.set("0%")  # Valor padrão quando desabilitado
                origem_icms_combobox.config(state="disabled")
            else:
                origem_icms_combobox.config(state="normal")  # Desbloqueado para seleção do usuário
        else:
            origem_icms_combobox.config(state="normal")  # Sempre habilitado para S/ST

# Função para calcular o custo final (Modo Compra)
def calcular_custo():
    global dados_planilha
    try:
        ncm_input = ncm_var.get().replace(".", "")
        custo_unitario_str = custo_var.get().replace(",", ".")
        custo_unitario = float(custo_unitario_str)

        linha_ncm = None
        for row in dados_planilha:
            ncm_planilha = str(row[0]).replace(".", "")
            if ncm_planilha == ncm_input:
                linha_ncm = row
                break

        if not linha_ncm:
            resumo_var.set("NCM não encontrado na planilha.")
            return

        # Extrair dados da planilha
        aquisicao = linha_ncm[3]  # Coluna D: INTERNA ou EXTERNA
        origem = linha_ncm[4]     # Coluna E: NACIONAL ou IMPORTADA
        st_inclusa = linha_ncm[5] # Coluna F: SIM, NÃO ou ICMS-ST
        origem_importada = linha_ncm[6] or 0  # Coluna G (4%)
        origem_nacional = linha_ncm[7] or 0   # Coluna H (12%)
        origem_rj = linha_ncm[8] or 0         # Coluna I (22%)

        # Obter o estado de origem selecionado
        estado_origem = estado_origem_var.get()
        
        # Determinar a alíquota com base no estado de origem
        if estado_origem == "RJ":
            # Compra interna no RJ
            aliquota_origem = obter_aliquota_interna("RJ")
        else:
            # Compra interestadual
            aliquota_origem = obter_aliquota_interestadual(estado_origem, "RJ")
        
        # Calcular DIFAL (se aplicável)
        aliquota_interna_rj = obter_aliquota_interna("RJ")
        difal = 0
        
        if estado_origem != "RJ":  # Se for compra interestadual
            difal = (aliquota_interna_rj - aliquota_origem) * custo_unitario

        # Determinar origem_icms_value para C/ST com ST não inclusa
        if st_difal_var.get() == "C/ST" and not st_inclusa_var.get():
            origem_icms_percent = origem_icms_var.get()
            if origem_icms_percent == "4%":
                origem_icms_value = origem_importada / 100
            elif origem_icms_percent == "12%":
                origem_icms_value = origem_nacional / 100
            elif origem_icms_percent == "22%":
                origem_icms_value = origem_rj / 100
            else:
                origem_icms_value = 0
        else:
            # Para C/ST com ST inclusa ou S/ST, determinamos automaticamente
            if origem == "IMPORTADA":
                origem_icms_value = origem_importada / 100
            elif origem == "NACIONAL":
                origem_icms_value = origem_nacional / 100
            else:
                origem_icms_value = 0

        if st_difal_var.get() == "C/ST":
            if st_inclusa_var.get():  # ST já recolhido
                custo_st_interes_estadual = 0
                custo_st_interna = 0
            else:  # ST não inclusa
                custo_st_interes_estadual = custo_unitario * origem_icms_value
                custo_st_interna = 0  # Removido o cálculo de ST interna
        else:  # S/ST
            origem_icms_percent = origem_icms_var.get()
            if origem_icms_percent == "4%":
                origem_icms_value = origem_importada / 100
            elif origem_icms_percent == "12%":
                origem_icms_value = origem_nacional / 100
            elif origem_icms_percent == "22%":
                origem_icms_value = origem_rj / 100
            else:
                origem_icms_value = 0
            custo_st_interes_estadual = custo_unitario * origem_icms_value
            custo_st_interna = 0

        # Adicionar DIFAL ao custo final para compras interestaduais
        custo_final = custo_unitario + custo_st_interes_estadual + custo_st_interna + difal
        custo_final_formatado = f"{custo_final:.2f}".replace('.', ',')
        custo_calculado_var.set(custo_final_formatado)
        
    except ValueError:
        resumo_var.set("Erro nos dados de entrada. Verifique os valores!")
    except Exception as e:
        resumo_var.set(f"Erro: {str(e)}")

# Função para calcular o preço de venda (Modo Venda)
def calcular_preco_venda():
    global dados_planilha
    try:
        ncm_input = ncm_var.get().replace(".", "")
        custo_unitario_str = custo_var.get().replace(",", ".")
        custo_unitario = float(custo_unitario_str)
        margem_str = margem_var.get().replace(",", ".")
        margem = float(margem_str) / 100  # Converter percentual para decimal

        linha_ncm = None
        for row in dados_planilha:
            ncm_planilha = str(row[0]).replace(".", "")
            if ncm_planilha == ncm_input:
                linha_ncm = row
                break

        if not linha_ncm:
            resumo_var.set("NCM não encontrado na planilha.")
            return

        # Obter o estado de destino selecionado
        estado_destino = estado_destino_var.get()
        
        # Determinar a alíquota com base no estado de destino
        if estado_destino == "RJ":
            # Venda interna no RJ
            aliquota_destino = obter_aliquota_interna("RJ")
        else:
            # Venda interestadual
            aliquota_destino = obter_aliquota_interestadual("RJ", estado_destino)
        
        # Calcular o preço de venda
        # Fórmula: Preço = Custo / (1 - Margem - Alíquota ICMS)
        preco_venda = custo_unitario / (1 - margem - aliquota_destino)
        
        # Ajustar para ST se aplicável
        if st_difal_var.get() == "C/ST":
            # Se o produto está sujeito a ST, pode ser necessário ajustar o preço
            # Simplificação: apenas informamos que o preço pode precisar de ajuste
            resumo_var.set(f"Produto sujeito a ST no estado de destino ({estado_destino}). O preço sugerido pode precisar de ajustes adicionais.")
        
        preco_venda_formatado = f"{preco_venda:.2f}".replace('.', ',')
        preco_venda_var.set(preco_venda_formatado)
        
    except ValueError:
        resumo_var.set("Erro nos dados de entrada. Verifique os valores!")
    except Exception as e:
        resumo_var.set(f"Erro: {str(e)}")

# Função para calcular com base no modo atual
def calcular():
    if modo_var.get() == "Compra":
        calcular_custo()
    else:  # Modo Venda
        calcular_preco_venda()

# Função para copiar o resultado
def copiar_resultado():
    try:
        root.clipboard_clear()
        if modo_var.get() == "Compra":
            root.clipboard_append(custo_calculado_var.get())
            resumo_var.set("Custo final copiado para a área de transferência!")
        else:  # Modo Venda
            root.clipboard_append(preco_venda_var.get())
            resumo_var.set("Preço de venda copiado para a área de transferência!")
        root.update()
    except Exception as e:
        resumo_var.set(f"Erro ao copiar: {str(e)}")

# Função para limpar os campos
def limpar_campos():
    ncm_var.set("")
    st_difal_var.set("N/A")
    st_inclusa_var.set(False)  # Desmarcar o checkbox
    origem_icms_var.set("0%")
    custo_var.set("")
    custo_calculado_var.set("")
    preco_venda_var.set("")
    margem_var.set("")
    resumo_var.set("")
    resultados_tree.selection_remove(resultados_tree.selection())
    st_inclusa_checkbutton.config(state="disabled")
    origem_icms_combobox.config(state="normal")  # Habilitar Origem ICMS por padrão
    resultados_tree.delete(*resultados_tree.get_children())
    for item in dados:
        resultados_tree.insert("", "end", values=(item["NCM"], item["Produto"]))
    atualizar_interface_modo()

# Função para alternar entre os modos Compra e Venda
def alternar_modo():
    # Atualizar a interface com base no modo selecionado
    atualizar_interface_modo()
    limpar_campos()

# Função para atualizar a interface com base no modo selecionado
def atualizar_interface_modo():
    modo_atual = modo_var.get()
    
    # Atualizar título
    if modo_atual == "Compra":
        titulo_label.config(text="CÁLCULO ST/DIFAL - MODO COMPRA")
        # Mostrar elementos do Modo Compra
        estado_origem_label.grid(column=0, row=6, sticky="W", padx=5, pady=2)
        estado_origem_combobox.grid(column=1, row=6, sticky="EW", columnspan=3, padx=5, pady=2)
        st_inclusa_label.grid(column=0, row=7, sticky="W", padx=5, pady=2)
        st_inclusa_checkbutton.grid(column=1, row=7, sticky="W", padx=5, pady=2)
        origem_icms_label.grid(column=0, row=8, sticky="W", padx=5, pady=2)
        origem_icms_combobox.grid(column=1, row=8, sticky="EW", columnspan=3, padx=5, pady=2)
        custo_label.grid(column=0, row=9, sticky="W", padx=5, pady=2)
        custo_entry.grid(column=1, row=9, sticky="EW", columnspan=3, padx=5, pady=2)
        calcular_button.grid(column=0, row=10, sticky="W", padx=5, pady=2)
        copiar_button.grid(column=1, row=10, sticky="EW", padx=5, pady=2)
        limpar_button.grid(column=2, row=10, sticky="E", padx=5, pady=2)
        resultado_label.config(text="Custo Final:")
        resultado_label.grid(column=0, row=13, sticky="W", padx=5, pady=2)
        resultado_value_label.grid(column=1, row=13, sticky="W", columnspan=3, padx=5, pady=2)
        
        # Esconder elementos do Modo Venda
        estado_destino_label.grid_remove()
        estado_destino_combobox.grid_remove()
        margem_label.grid_remove()
        margem_entry.grid_remove()
        preco_venda_label.grid_remove()
        preco_venda_value_label.grid_remove()
        
    else:  # Modo Venda
        titulo_label.config(text="CÁLCULO ST/DIFAL - MODO VENDA")
        # Esconder elementos do Modo Compra
        estado_origem_label.grid_remove()
        estado_origem_combobox.grid_remove()
        st_inclusa_label.grid_remove()
        st_inclusa_checkbutton.grid_remove()
        origem_icms_label.grid_remove()
        origem_icms_combobox.grid_remove()
        
        # Mostrar elementos do Modo Venda
        estado_destino_label.grid(column=0, row=6, sticky="W", padx=5, pady=2)
        estado_destino_combobox.grid(column=1, row=6, sticky="EW", columnspan=3, padx=5, pady=2)
        custo_label.grid(column=0, row=7, sticky="W", padx=5, pady=2)
        custo_entry.grid(column=1, row=7, sticky="EW", columnspan=3, padx=5, pady=2)
        margem_label.grid(column=0, row=8, sticky="W", padx=5, pady=2)
        margem_entry.grid(column=1, row=8, sticky="EW", columnspan=3, padx=5, pady=2)
        calcular_button.grid(column=0, row=9, sticky="W", padx=5, pady=2)
        copiar_button.grid(column=1, row=9, sticky="EW", padx=5, pady=2)
        limpar_button.grid(column=2, row=9, sticky="E", padx=5, pady=2)
        preco_venda_label.grid(column=0, row=13, sticky="W", padx=5, pady=2)
        preco_venda_value_label.grid(column=1, row=13, sticky="W", columnspan=3, padx=5, pady=2)
        
        # Esconder label de custo final
        resultado_label.grid_remove()
        resultado_value_label.grid_remove()

# Função para alternar entre os tipos de usuário
def alternar_usuario():
    tipo_usuario = usuario_var.get()
    
    if tipo_usuario == "Comprador":
        # Comprador só pode acessar Modo Compra
        modo_var.set("Compra")
        modo_compra_radio.config(state="normal")
        modo_venda_radio.config(state="disabled")
    elif tipo_usuario == "Vendedor":
        # Vendedor só pode acessar Modo Venda
        modo_var.set("Venda")
        modo_compra_radio.config(state="disabled")
        modo_venda_radio.config(state="normal")
    
    alternar_modo()

# Função para exibir o tutorial
def mostrar_tutorial():
    modo_atual = modo_var.get()
    
    if modo_atual == "Compra":
        tutorial = (
            "Tutorial: Como Usar o Programa de Cálculo ST/DIFAL para Compras\n\n"
            "1. Selecione seu Tipo de Usuário\n"
            "- Escolha 'Comprador' para acessar o Modo Compra.\n\n"
            "2. Busque o NCM ou Produto\n"
            "- No campo 'Buscar (NCM ou Produto)', digite o NCM ou nome do produto (ex.: '7605' ou 'fios').\n"
            "- Clique no item desejado na tabela que aparece abaixo.\n\n"
            "3. Verifique os Dados Automáticos\n"
            "- 'NCM Selecionado' e 'ST / DIFAL' (C/ST ou S/ST) serão preenchidos automaticamente.\n"
            "- Se 'ST / DIFAL' for 'C/ST', o campo 'ST Inclusa na compra?' será ativado.\n\n"
            "4. Selecione o Estado de Origem\n"
            "- Escolha o estado de origem da compra no campo 'Estado de Origem'.\n\n"
            "5. Configure 'ST Inclusa na compra?' (se aplicável)\n"
            "- Para 'C/ST': Marque a caixa se o ST interestadual já foi recolhido.\n"
            "  - Se marcada, 'Origem ICMS' será desativado e o custo final será o custo unitário.\n"
            "  - Se não marcada, escolha a 'Origem ICMS' (0%, 4%, 12% ou 22%) para calcular o ST interestadual.\n\n"
            "6. Escolha 'Origem ICMS' (se aplicável)\n"
            "- Para 'S/ST' ou 'C/ST' com 'ST Inclusa' desmarcada: Selecione o percentual de 'Origem ICMS' (0%, 4%, 12% ou 22%).\n\n"
            "7. Insira o Custo Unitário\n"
            "- Digite o valor do produto no campo 'Custo Unitário' (ex.: '1200,50'). Use vírgula para decimais.\n\n"
            "8. Calcule o Custo Final\n"
            "- Clique em 'Calcular' ou pressione 'Enter' no campo 'Custo Unitário'.\n"
            "- O 'Custo Final' será exibido, refletindo o ST interestadual e DIFAL conforme a configuração.\n\n"
            "9. Copie o Resultado (Opcional)\n"
            "- Clique em 'Copiar' para copiar o custo final para a área de transferência.\n\n"
            "10. Limpe os Campos (Opcional)\n"
            "- Clique em 'Limpar' para reiniciar todos os campos e a tabela.\n\n"
            "Dica\n"
            "- Certifique-se de que a planilha 'CALCULO_ST_DIFAL_COMPLETA.xlsx' está na mesma pasta do programa."
        )
    else:  # Modo Venda
        tutorial = (
            "Tutorial: Como Usar o Programa de Cálculo ST/DIFAL para Vendas\n\n"
            "1. Selecione seu Tipo de Usuário\n"
            "- Escolha 'Vendedor' para acessar o Modo Venda.\n\n"
            "2. Busque o NCM ou Produto\n"
            "- No campo 'Buscar (NCM ou Produto)', digite o NCM ou nome do produto (ex.: '7605' ou 'fios').\n"
            "- Clique no item desejado na tabela que aparece abaixo.\n\n"
            "3. Verifique os Dados Automáticos\n"
            "- 'NCM Selecionado' e 'ST / DIFAL' (C/ST ou S/ST) serão preenchidos automaticamente.\n\n"
            "4. Selecione o Estado de Destino\n"
            "- Escolha o estado de destino da venda no campo 'Estado de Destino'.\n\n"
            "5. Insira o Custo Final\n"
            "- Digite o custo final do produto no campo 'Custo Final' (ex.: '1200,50'). Use vírgula para decimais.\n"
            "- Este valor deve ser o custo final já calculado, incluindo impostos de compra.\n\n"
            "6. Insira a Margem de Lucro\n"
            "- Digite a margem de lucro desejada em percentual (ex.: '30' para 30%).\n\n"
            "7. Calcule o Preço de Venda\n"
            "- Clique em 'Calcular' ou pressione 'Enter' no campo 'Margem de Lucro'.\n"
            "- O 'Preço de Venda Sugerido' será exibido, considerando a margem e os impostos aplicáveis.\n\n"
            "8. Copie o Resultado (Opcional)\n"
            "- Clique em 'Copiar' para copiar o preço de venda para a área de transferência.\n\n"
            "9. Limpe os Campos (Opcional)\n"
            "- Clique em 'Limpar' para reiniciar todos os campos e a tabela.\n\n"
            "Dica\n"
            "- Certifique-se de que a planilha 'CALCULO_ST_DIFAL_COMPLETA.xlsx' está na mesma pasta do programa."
        )
    
    tk.messagebox.showinfo("Tutorial", tutorial)

# ------------------------------------------------------------
# VARIÁVEIS GLOBAIS E INTERFACE GRÁFICA
# ------------------------------------------------------------
root = tk.Tk()
root.title("CÁLCULO ST/DIFAL")
root.option_add("*Font", ("Arial", 16))  # Reduz a fonte global para 16

# Maximizar a janela automaticamente
root.state('zoomed')  # Abre maximizada no Windows
root.minsize(900, 700)  # Tamanho mínimo

# Variáveis para controle de usuário e modo
usuario_var = tk.StringVar(value="Comprador")  # Padrão: Comprador
modo_var = tk.StringVar(value="Compra")  # Padrão: Modo Compra

# Variáveis para campos de entrada
busca_var = tk.StringVar()
ncm_var = tk.StringVar()
custo_var = tk.StringVar()
margem_var = tk.StringVar()
resumo_var = tk.StringVar()
custo_calculado_var = tk.StringVar(value="")
preco_venda_var = tk.StringVar(value="")
st_difal_var = tk.StringVar(value="N/A")
st_inclusa_var = tk.BooleanVar(value=False)  # Booleano para o checkbox
origem_icms_var = tk.StringVar(value="0%")
estado_origem_var = tk.StringVar(value="RJ")  # Padrão: RJ
estado_destino_var = tk.StringVar(value="RJ")  # Padrão: RJ

# Validação de números
vcmd = (root.register(validar_numero), '%P')

frame = ttk.Frame(root, padding="10")
frame.grid(sticky="NSEW")

# Configurar expansão do frame
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# Título
titulo_label = ttk.Label(frame, text="CÁLCULO ST/DIFAL - MODO COMPRA", font=("Arial", 18, "bold"), foreground="blue")
titulo_label.grid(column=0, row=0, columnspan=4, pady=5, sticky="EW")

# Botão Tutorial no canto superior direito
ttk.Button(frame, text="Tutorial", command=mostrar_tutorial).grid(column=3, row=0, sticky="NE", padx=5, pady=5)

# Frame para seleção de usuário
usuario_frame = ttk.LabelFrame(frame, text="Tipo de Usuário", padding="5")
usuario_frame.grid(column=0, row=1, columnspan=4, sticky="EW", padx=5, pady=5)

ttk.Radiobutton(usuario_frame, text="Comprador", variable=usuario_var, value="Comprador", command=alternar_usuario).grid(column=0, row=0, padx=10, pady=2)
ttk.Radiobutton(usuario_frame, text="Vendedor", variable=usuario_var, value="Vendedor", command=alternar_usuario).grid(column=1, row=0, padx=10, pady=2)

# Frame para seleção de modo
modo_frame = ttk.LabelFrame(frame, text="Modo de Operação", padding="5")
modo_frame.grid(column=0, row=2, columnspan=4, sticky="EW", padx=5, pady=5)

modo_compra_radio = ttk.Radiobutton(modo_frame, text="Modo Compra", variable=modo_var, value="Compra", command=alternar_modo)
modo_compra_radio.grid(column=0, row=0, padx=10, pady=2)

modo_venda_radio = ttk.Radiobutton(modo_frame, text="Modo Venda", variable=modo_var, value="Venda", command=alternar_modo)
modo_venda_radio.grid(column=1, row=0, padx=10, pady=2)

# Buscar
ttk.Label(frame, text="Buscar (NCM ou Produto):").grid(column=0, row=3, sticky="W", padx=5, pady=2)
entry_busca = ttk.Entry(frame, textvariable=busca_var)
entry_busca.grid(column=1, row=3, sticky="EW", columnspan=3, padx=5, pady=2)
entry_busca.bind("<KeyRelease>", filtrar_dados)

style = ttk.Style()
style.theme_use("clam")  # Tema mais moderno
style.configure("Treeview", font=("Arial", 14), rowheight=30)
style.configure("Treeview.Heading", font=("Arial", 14, "bold"))
style.configure("TButton", font=("Arial", 14))
style.configure("Large.TCheckbutton", font=("Arial", 10)) #botão flag

# Criar Treeview com scrollbar
colunas = ("NCM", "Produto")
resultados_tree = ttk.Treeview(frame, columns=colunas, show="headings", height=8)  # Reduzido para 8 linhas
resultados_tree.heading("NCM", text="NCM")
resultados_tree.heading("Produto", text="Produto")
resultados_tree.column("NCM", width=260, minwidth=80, stretch=False)  # Reduzido para 80 pixels
resultados_tree.column("Produto", width=550, minwidth=550)
resultados_tree.grid(column=0, row=4, columnspan=4, padx=5, pady=5, sticky="NSEW")

scrollbar = ttk.Scrollbar(frame, orient="vertical", command=resultados_tree.yview)  # Depois cria o Scrollbar
scrollbar.grid(column=4, row=4, sticky="NS", padx=5, pady=5)
resultados_tree.configure(yscrollcommand=scrollbar.set)  # Vincula o yscrollcommand após criar o Treeview

resultados_tree.bind("<<TreeviewSelect>>", selecionar_item)

# NCM Selecionado
ttk.Label(frame, text="NCM Selecionado:").grid(column=0, row=5, sticky="W", padx=5, pady=2)
ttk.Entry(frame, textvariable=ncm_var, state="readonly").grid(column=1, row=5, sticky="EW", columnspan=3, padx=5, pady=2)

# ST / DIFAL
ttk.Label(frame, text="ST / DIFAL:").grid(column=2, row=5, sticky="E", padx=5, pady=2)
ttk.Entry(frame, textvariable=st_difal_var, state="readonly", width=10).grid(column=3, row=5, sticky="W", padx=5, pady=2)

# Elementos do Modo Compra
# Estado de Origem
estado_origem_label = ttk.Label(frame, text="Estado de Origem:")
estado_origem_combobox = ttk.Combobox(frame, textvariable=estado_origem_var, values=estados_brasileiros, state="readonly", font=("Arial", 14))

# ST Inclusa na compra? (checkbox)
st_inclusa_label = ttk.Label(frame, text="ST Inclusa na compra?:")
st_inclusa_checkbutton = ttk.Checkbutton(frame, variable=st_inclusa_var, state="disabled", style="Large.TCheckbutton")
st_inclusa_var.trace("w", atualizar_origem_icms)

# Origem ICMS (atualizado para 4%, 12%, 22%)
origem_icms_label = ttk.Label(frame, text="Origem ICMS:")
origem_icms_combobox = ttk.Combobox(frame, textvariable=origem_icms_var, values=["0%", "4%", "12%", "22%"], state="normal", font=("Arial", 14))

# Elementos do Modo Venda
# Estado de Destino
estado_destino_label = ttk.Label(frame, text="Estado de Destino:")
estado_destino_combobox = ttk.Combobox(frame, textvariable=estado_destino_var, values=estados_brasileiros, state="readonly", font=("Arial", 14))

# Margem de Lucro
margem_label = ttk.Label(frame, text="Margem de Lucro (%):")
margem_entry = ttk.Entry(frame, textvariable=margem_var, font=("Arial", 14), validate="key", validatecommand=vcmd)
margem_entry.bind("<Return>", lambda event: calcular())

# Custo Unitário/Final
custo_label = ttk.Label(frame, text="Custo Unitário:")
custo_entry = ttk.Entry(frame, textvariable=custo_var, font=("Arial", 14), validate="key", validatecommand=vcmd)
custo_entry.bind("<Return>", lambda event: calcular())

# Botões
calcular_button = ttk.Button(frame, text="Calcular", command=calcular)
copiar_button = ttk.Button(frame, text="Copiar", command=copiar_resultado)
limpar_button = ttk.Button(frame, text="Limpar", command=limpar_campos)

# Selecionados
ttk.Label(frame, text="Selecionados:", foreground="blue", font=("Arial", 14)).grid(column=0, row=11, columnspan=4, sticky="W", padx=5, pady=2)
ttk.Label(frame, textvariable=resumo_var, font=("Arial", 14), wraplength=700).grid(column=0, row=12, columnspan=4, sticky="NSEW", padx=5, pady=2)

# Resultado (Custo Final para Compra)
resultado_label = ttk.Label(frame, text="Custo Final:", foreground="green", font=("Arial", 14))
resultado_value_label = ttk.Label(frame, textvariable=custo_calculado_var, foreground="green", font=("Arial", 14))

# Resultado (Preço de Venda para Venda)
preco_venda_label = ttk.Label(frame, text="Preço de Venda Sugerido:", foreground="green", font=("Arial", 14))
preco_venda_value_label = ttk.Label(frame, textvariable=preco_venda_var, foreground="green", font=("Arial", 14))

carregar_dados()
filtrar_dados()

# Configurar expansão do frame e widgets
frame.columnconfigure(1, weight=1)  # Expande a segunda coluna (widgets com fundo branco)
frame.columnconfigure(2, weight=1)  # Expande a terceira coluna (mais espaço para os widgets)
frame.columnconfigure(3, weight=1)  # Expande a quarta coluna (para alinhamento)
frame.rowconfigure(4, weight=1)    # Faz o Treeview expandir verticalmente
frame.rowconfigure(12, weight=1)   # Faz o label de "Selecionados" expandir verticalmente

# Inicializar a interface com base no modo padrão
alternar_usuario()  # Isso também chamará alternar_modo()

root.mainloop()
