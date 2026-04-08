import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import pyperclip

# --- Variáveis Globais ---
dados_matriz = []
ncm_selecionado_dados = {}

# --- Funções de Dados ---
def carregar_dados_matriz():
    global dados_matriz
    try:
        workbook = openpyxl.load_workbook("MATRIZ_ST_SAIDA.xlsx", data_only=True)
        sheet = workbook["MATRIZ_SAIDA"]
        dados_matriz.clear()
        
        # Ignora cabeçalho (linha 1)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            ncm = str(row[0]).strip()
            produto = row[1]
            uf_destino = row[2]
            tem_st = row[3]
            alq_inter = row[4]
            alq_intra = row[5]
            mva = row[6]
            
            if ncm and uf_destino:
                dados_matriz.append({
                    "NCM": ncm,
                    "Produto": produto,
                    "UF": uf_destino,
                    "ST": tem_st,
                    "Alq_Inter": float(alq_inter) if alq_inter else 0.0,
                    "Alq_Intra": float(alq_intra) if alq_intra else 0.0,
                    "MVA": float(mva) if mva else 0.0
                })
        print(f"Carregados {len(dados_matriz)} registros da matriz de saídas.")
    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo 'MATRIZ_ST_SAIDA.xlsx' não encontrado.\nCertifique-se de que a matriz ST está na mesma pasta que o programa.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro crítico ao carregar a matriz ST: {e}")

# --- Funções de Interface ---
def buscar_ncm_estado(event=None):
    ncm_busca = busca_var.get().strip().replace(".", "")
    uf_busca = uf_var.get().strip().upper()
    
    resultados_tree.delete(*resultados_tree.get_children())
    
    for item in dados_matriz:
        ncm_item = item["NCM"].replace(".", "")
        if (ncm_busca in ncm_item or ncm_busca == "") and (uf_busca == "" or uf_busca == item["UF"]):
            resultados_tree.insert("", "end", values=(
                item["NCM"], item["Produto"], item["UF"], item["ST"], 
                f"{item['Alq_Inter']}%", f"{item['Alq_Intra']}%", f"{item['MVA']}%"
            ))

def selecionar_item(event):
    global ncm_selecionado_dados
    selecao = resultados_tree.focus()
    if selecao:
        valores = resultados_tree.item(selecao, "values")
        
        # Guardar valores brutos para cálculo
        ncm_selecionado_dados = {
            "NCM": valores[0],
            "UF": valores[2],
            "ST": valores[3],
            "Alq_Inter": float(valores[4].replace("%", "")),
            "Alq_Intra": float(valores[5].replace("%", "")),
            "MVA": float(valores[6].replace("%", ""))
        }
        
        info_var.set(f"Selecionado: {valores[0]} | Destino: {valores[2]} | ST? {valores[3]}")
        
        # Habilitar/Desabilitar cálculo com base na ST
        if valores[3].upper() == "NÃO":
            btn_calcular.config(state="disabled")
            messagebox.showinfo("Regra", "Este Produto/UF não possui acordo de ST. Aplica-se a regra geral do ICMS.")
            limpar_resultados()
        else:
            btn_calcular.config(state="normal")

def calcular_st():
    if not ncm_selecionado_dados or ncm_selecionado_dados.get("ST") == "NÃO":
        return
        
    try:
        valor_produto_str = valor_var.get().replace(",", ".")
        valor_produto = float(valor_produto_str)
        
        frete_str = frete_var.get().replace(",", ".") if frete_var.get() else "0"
        frete = float(frete_str)
        
        # Lógica ST Saída
        alq_inter = ncm_selecionado_dados["Alq_Inter"] / 100.0
        alq_intra = ncm_selecionado_dados["Alq_Intra"] / 100.0
        mva_ajustada = ncm_selecionado_dados["MVA"] / 100.0
        
        valor_total_nota = valor_produto + frete
        
        # 1. ICMS Próprio (Origem)
        icms_proprio = valor_total_nota * alq_inter
        
        # 2. Base ST
        base_st = valor_total_nota * (1 + mva_ajustada)
        
        # 3. ICMS ST a reter
        icms_st = (base_st * alq_intra) - icms_proprio
        if icms_st < 0:
            icms_st = 0 # Não existe ST negativa
            
        custo_final_cliente = valor_total_nota + icms_st
        
        # Atualizar Interface
        res_icms_proprio_var.set(f"R$ {icms_proprio:.2f}".replace(".", ","))
        res_base_st_var.set(f"R$ {base_st:.2f}".replace(".", ","))
        res_icms_st_var.set(f"R$ {icms_st:.2f}".replace(".", ","))
        res_total_nf_var.set(f"R$ {custo_final_cliente:.2f}".replace(".", ","))
        
    except ValueError:
        messagebox.showerror("Erro de Formato", "Por favor, digite valores numéricos válidos (ex: 1500,50)")

def limpar_resultados():
    res_icms_proprio_var.set("")
    res_base_st_var.set("")
    res_icms_st_var.set("")
    res_total_nf_var.set("")

def copiar_calculo():
    if res_icms_st_var.get():
        texto = f"Base ST: {res_base_st_var.get()} | ICMS ST Retido: {res_icms_st_var.get()} | Total com ST: {res_total_nf_var.get()}"
        root.clipboard_clear()
        root.clipboard_append(texto)
        messagebox.showinfo("Copiado", "Resumo do cálculo copiado para a área de transferência!")

# --- Configuração Tkinter ---
root = tk.Tk()
root.title("Cálculo ST: Fase 2 - Saída (Venda)")
root.state('zoomed')
root.option_add("*Font", ("Arial", 14))

# Variáveis String
busca_var = tk.StringVar()
uf_var = tk.StringVar(value="SP")
info_var = tk.StringVar(value="Selecione um produto na lista.")
valor_var = tk.StringVar()
frete_var = tk.StringVar()

# Variáveis de Resultado
res_icms_proprio_var = tk.StringVar()
res_base_st_var = tk.StringVar()
res_icms_st_var = tk.StringVar()
res_total_nf_var = tk.StringVar()

frame = ttk.Frame(root, padding=20)
frame.pack(fill="both", expand=True)

# --- Cabeçalho ---
lbl_titulo = ttk.Label(frame, text="MÓDULO DE VENDAS : CÁLCULO ST INTERESTADUAL", font=("Arial", 18, "bold"), foreground="#B22222")
lbl_titulo.grid(row=0, column=0, columnspan=4, pady=10, sticky="w")

# --- Linha 1: Buscas ---
ttk.Label(frame, text="Buscar NCM/Prod:").grid(row=1, column=0, sticky="e", pady=5)
entry_busca = ttk.Entry(frame, textvariable=busca_var, width=30)
entry_busca.grid(row=1, column=1, sticky="w", padx=5)
entry_busca.bind("<KeyRelease>", buscar_ncm_estado)

ttk.Label(frame, text="Filtrar UF:").grid(row=1, column=2, sticky="e")
combo_uf = ttk.Combobox(frame, textvariable=uf_var, values=["SP", "MG", "ES", "BA", "PR", "RJ Todos..."], state="normal", width=10)
combo_uf.grid(row=1, column=3, sticky="w", padx=5)
combo_uf.bind("<<ComboboxSelected>>", buscar_ncm_estado)

# --- Linha 2: Tabela Principal ---
colunas = ("NCM", "Produto", "UF", "ST", "Inter", "Intra", "MVA")
resultados_tree = ttk.Treeview(frame, columns=colunas, show="headings", height=8)

resultados_tree.heading("NCM", text="NCM")
resultados_tree.heading("Produto", text="Produto")
resultados_tree.heading("UF", text="UF Destino")
resultados_tree.heading("ST", text="Tem ST?")
resultados_tree.heading("Inter", text="Alíq. Interestadual")
resultados_tree.heading("Intra", text="Alíq. Interna")
resultados_tree.heading("MVA", text="MVA Ajustada")

resultados_tree.column("NCM", width=120)
resultados_tree.column("Produto", width=350)
resultados_tree.column("UF", width=100)
resultados_tree.column("ST", width=80)
resultados_tree.column("Inter", width=150)
resultados_tree.column("Intra", width=120)
resultados_tree.column("MVA", width=120)

resultados_tree.grid(row=2, column=0, columnspan=4, pady=15, sticky="nsew")
resultados_tree.bind("<<TreeviewSelect>>", selecionar_item)

# Expansão
frame.rowconfigure(2, weight=1)
frame.columnconfigure(1, weight=1)

# --- Linha 3: Info ---
lbl_info = ttk.Label(frame, textvariable=info_var, font=("Arial", 14, "bold"), foreground="blue")
lbl_info.grid(row=3, column=0, columnspan=4, sticky="w")

separator = ttk.Separator(frame, orient='horizontal')
separator.grid(row=4, column=0, columnspan=4, sticky="ew", pady=15)

# --- Linha 4: Inputs de Finanças ---
frm_valores = ttk.Frame(frame)
frm_valores.grid(row=5, column=0, columnspan=4, sticky="w")

ttk.Label(frm_valores, text="Valor da Mercadoria (R$):").grid(row=0, column=0, padx=5, sticky="e")
entry_valor = ttk.Entry(frm_valores, textvariable=valor_var, width=15)
entry_valor.grid(row=0, column=1)

ttk.Label(frm_valores, text="Frete / Despesas (R$):").grid(row=0, column=2, padx=(20, 5), sticky="e")
entry_frete = ttk.Entry(frm_valores, textvariable=frete_var, width=15)
entry_frete.grid(row=0, column=3)
frete_var.set("0,00")

btn_calcular = ttk.Button(frm_valores, text="💲 CALCULAR ST", command=calcular_st, state="disabled")
btn_calcular.grid(row=0, column=4, padx=20)

btn_copiar = ttk.Button(frm_valores, text="📋 Copiar", command=copiar_calculo)
btn_copiar.grid(row=0, column=5)

# --- Linha 5: Resultados da ST ---
frm_resultados = ttk.LabelFrame(frame, text=" Demonstrativo do Diferencial de Substituição ", padding=15)
frm_resultados.grid(row=6, column=0, columnspan=4, pady=20, sticky="ew")

ttk.Label(frm_resultados, text="ICMS Próprio (Origem):").grid(row=0, column=0, sticky="w")
ttk.Label(frm_resultados, textvariable=res_icms_proprio_var, font=("Arial", 14, "bold")).grid(row=0, column=1, sticky="w", padx=10)

ttk.Label(frm_resultados, text="Base de Cálculo ST:").grid(row=1, column=0, sticky="w", pady=10)
ttk.Label(frm_resultados, textvariable=res_base_st_var, font=("Arial", 14, "bold")).grid(row=1, column=1, sticky="w", padx=10)

ttk.Label(frm_resultados, text="ICMS-ST A RETER:").grid(row=2, column=0, sticky="w", pady=10)
ttk.Label(frm_resultados, textvariable=res_icms_st_var, font=("Arial", 14, "bold"), foreground="red").grid(row=2, column=1, sticky="w", padx=10)

separator_res = ttk.Separator(frm_resultados, orient='horizontal')
separator_res.grid(row=3, column=0, columnspan=2, sticky="ew", pady=10)

ttk.Label(frm_resultados, text="TOTAL DA NOTA FISCAL:").grid(row=4, column=0, sticky="w")
ttk.Label(frm_resultados, textvariable=res_total_nf_var, font=("Arial", 16, "bold"), foreground="green").grid(row=4, column=1, sticky="w", padx=10)

# Inicializar
carregar_dados_matriz()
buscar_ncm_estado()

root.mainloop()
