"""
============================================================
  CÁLCULO ST — FASE 2: SAÍDAS INTERESTADUAIS
  Versão: 2.1
  Origem: Rio de Janeiro (RJ)
  Tabela ICMS 2026 — válida a partir de 01/04/2026
============================================================
  Fluxo:
    - Excel (contadora preenche): NCM | Produto | Acordo ST | MVA
    - Interface (usuário seleciona por operação): UF destino + Importado?
  Dependências: openpyxl, tkinter (padrão)
  Planilha: MATRIZ_ST_SAIDA_ver1.xlsx (mesma pasta)
============================================================
"""

import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import os

# ============================================================
# TABELA ICMS 2026 — hardcoded (Origem: RJ)
# ============================================================

TODOS_ESTADOS = [
    "AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO",
    "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI",
    "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"
]

ALIQ_INTERNA = {
    "AC": 19.0,  "AL": 20.0,  "AM": 20.0,  "AP": 18.0,
    "BA": 20.5,  "CE": 20.0,  "DF": 20.0,  "ES": 17.0,
    "GO": 19.0,  "MA": 23.0,  "MT": 17.0,  "MS": 17.0,
    "MG": 18.0,  "PA": 19.0,  "PB": 20.0,  "PR": 19.5,
    "PE": 20.5,  "PI": 22.5,  "RN": 20.0,  "RS": 17.0,
    "RO": 19.5,  "RR": 20.0,  "SC": 17.0,  "SP": 18.0,
    "SE": 20.0,  "TO": 20.0,
}

UFS_7_PORCENTO = {"MG", "SP", "PR", "RS", "SC", "ES"}

NOME_PLANILHA = "MATRIZ_ST_SAIDA_ver2.xlsx"
ABA_PLANILHA  = "MATRIZ_SAIDA"

# ============================================================
# LÓGICA DE NEGÓCIO
# ============================================================

def obter_aliq_interestadual(uf_destino, importado):
    if importado:
        return 4.0
    return 7.0 if uf_destino in UFS_7_PORCENTO else 12.0


def obter_aliq_interna(uf_destino):
    """Retorna a alíquota interna do estado de destino (tabela 2026)."""
    return ALIQ_INTERNA.get(uf_destino, 0.0)


def calcular_difal_base_dupla(valor_merc, frete, aliq_inter_pct, aliq_intra_pct):
    """
    Calculo Base Dupla (CFOP 6403) - Consumidor Final Contribuinte
    Fórmula: (Valor - ICMS_Inter) / (1 - Aliq_Intra)
    """
    vt = valor_merc + frete
    ai = aliq_inter_pct / 100.0
    at = aliq_intra_pct / 100.0

    icms_inter = vt * ai
    base_dupla = (vt - icms_inter) / (1.0 - at)
    icms_st_difal = (base_dupla * at) - icms_inter

    if icms_st_difal < 0: icms_st_difal = 0.0

    return {
        "valor_total_nf": vt,
        "icms_proprio": icms_inter,
        "base_st": base_dupla,
        "icms_st": icms_st_difal,
        "total_nf": vt + icms_st_difal,
        "aliq_inter": aliq_inter_pct,
        "aliq_intra": aliq_intra_pct,
        "mva": 0.0,
        "modalidade": "DIFAL - BASE DUPLA (Contribuinte)"
    }


def calcular_difal_base_unica(valor_merc, frete, aliq_inter_pct, aliq_intra_pct):
    """
    Calculo Base Unica/Inclusa (CFOP 6108) - Consumidor Final Nao Contribuinte
    Fórmula: Valor / (1 - Aliq_Intra)
    """
    vt_original = valor_merc + frete
    ai = aliq_inter_pct / 100.0
    at = aliq_intra_pct / 100.0

    # Novo valor com imposto embutido
    novo_valor_prod = vt_original / (1.0 - at)
    icms_inter = novo_valor_prod * ai
    difal = (novo_valor_prod * at) - icms_inter

    if difal < 0: difal = 0.0

    return {
        "valor_total_nf": vt_original,
        "icms_proprio": icms_inter,
        "base_st": novo_valor_prod,
        "icms_st": difal,
        "total_nf": novo_valor_prod, # Na base unica, o total da NF e o proprio valor bruto
        "aliq_inter": aliq_inter_pct,
        "aliq_intra": aliq_intra_pct,
        "mva": 0.0,
        "modalidade": "DIFAL - BASE UNICA (Não Contribuinte)"
    }


def calcular_st(valor_merc, frete, mva_percent, aliq_inter_pct, aliq_intra_pct, cons_final=False, contribuinte=False):
    """
    Decide a modalidade de calculo e executa.
    """
    if cons_final:
        if contribuinte:
            return calcular_difal_base_dupla(valor_merc, frete, aliq_inter_pct, aliq_intra_pct)
        else:
            return calcular_difal_base_unica(valor_merc, frete, aliq_inter_pct, aliq_intra_pct)

    # Calculo original com MVA (Revenda)
    vt = valor_merc + frete
    ai = aliq_inter_pct / 100.0
    at = aliq_intra_pct / 100.0
    mva = mva_percent / 100.0

    icms_proprio = vt * ai
    base_st = vt * (1.0 + mva)
    icms_st = (base_st * at) - icms_proprio

    if icms_st < 0: icms_st = 0.0

    return {
        "valor_total_nf": vt,
        "icms_proprio": icms_proprio,
        "base_st": base_st,
        "icms_st": icms_st,
        "total_nf": vt + icms_st,
        "aliq_inter": aliq_inter_pct,
        "aliq_intra": aliq_intra_pct,
        "mva": mva_percent,
        "modalidade": "ICMS-ST (Revenda - MVA)"
    }

# ============================================================
# CARREGAMENTO DE DADOS
# Nova estrutura da planilha:
#   Col A: NCM | Col B: Produto | Col C: UF | Col D: ConsFinal | Col E: Contrib | Col F: Acordo ST | Col G: Importado | Col H: MVA
# ============================================================

dados_matriz = []

def carregar_dados():
    global dados_matriz
    dados_matriz.clear()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(script_dir, NOME_PLANILHA)

    if not os.path.exists(path):
        messagebox.showerror(
            "Planilha não encontrada",
            f"'{NOME_PLANILHA}' não encontrado em:\n{script_dir}\n\n"
            "Certifique-se de que a planilha está na mesma pasta que o programa."
        )
        return 0

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[ABA_PLANILHA]

        for row in ws.iter_rows(min_row=4, values_only=True):
            ncm       = row[0]   # Col A
            produto   = row[1]   # Col B
            uf        = row[2]   # Col C
            cons_fin  = row[3]   # Col D
            contrib   = row[4]   # Col E
            acordo    = row[5]   # Col F
            importado = row[6]   # Col G
            mva       = row[7]   # Col H

            if not ncm:
                continue

            try:
                mva_val = float(str(mva).replace(",", ".")) if mva is not None else None
            except (ValueError, TypeError):
                mva_val = None

            uf_str     = str(uf).strip().upper() if uf else ""
            cons_str   = str(cons_fin).strip().upper() if cons_fin else "NAO"
            cont_str   = str(contrib).strip().upper() if contrib else "NAO"
            acordo_str = str(acordo).strip().upper() if acordo else ""
            imp_str    = str(importado).strip().upper() if importado else "NAO"

            dados_matriz.append({
                "NCM":       str(ncm).strip(),
                "Produto":   str(produto).strip() if produto else "",
                "UF":        uf_str,
                "ConsFinal": cons_str,
                "Contribuinte": cont_str,
                "AcordoST":  acordo_str,
                "Importado": imp_str,
                "MVA":       mva_val,
            })

        wb.close()
        return len(dados_matriz)

    except KeyError:
        messagebox.showerror("Erro", f"Aba '{ABA_PLANILHA}' não encontrada.")
        return 0
    except Exception as e:
        messagebox.showerror("Erro ao carregar", str(e))
        return 0

# ============================================================
# FORMATAÇÃO
# ============================================================

def fmt_brl(v):
    return "R$ {:,.2f}".format(v).replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_pct(v):
    return "{:.2f}%".format(v).replace(".", ",")

# ============================================================
# INTERFACE
# ============================================================

class AppST:
    def __init__(self, root):
        self.root = root
        self.root.title("Cálculo ST — Fase 2: Saídas Interestaduais | Origem: RJ")
        self.root.state("zoomed")
        self.root.minsize(1000, 680)

        self.ncm_selecionado = {}
        self.dados_filtrados = []
        self._ultimo_resultado = None
        self._ultimo_item = None

        self._estilo()
        self._interface()

        total = carregar_dados()
        preenchidos = sum(1 for d in dados_matriz if d["MVA"] is not None and d["AcordoST"])
        self.status_var.set(
            "Planilha carregada: {} NCMs ({} com ST e MVA preenchidos). "
            "Selecione UF, Importado e clique no produto.".format(total, preenchidos)
        )
        self._refresh()

    # ----------------------------------------------------------
    def _estilo(self):
        s = ttk.Style()
        s.theme_use("clam")
        COR_AZUL  = "#1F3864"
        COR_MED   = "#2E75B6"
        COR_VERDE = "#217346"
        COR_BG    = "#F5F5F5"
        COR_W     = "#FFFFFF"

        self.root.configure(bg=COR_BG)
        s.configure("TFrame",    background=COR_BG)
        s.configure("TLabel",    background=COR_BG, font=("Calibri", 11))
        s.configure("TButton",   font=("Calibri", 11, "bold"), padding=6)
        s.configure("TEntry",    font=("Calibri", 11))
        s.configure("TCombobox", font=("Calibri", 11))
        s.configure("TCheckbutton", background=COR_BG, font=("Calibri", 11))

        s.configure("Main.Treeview",
                    font=("Calibri", 10), rowheight=24,
                    background=COR_W, fieldbackground=COR_W)
        s.configure("Main.Treeview.Heading",
                    font=("Calibri", 10, "bold"),
                    background=COR_AZUL, foreground=COR_W)
        s.map("Main.Treeview.Heading", background=[("active", COR_MED)])
        s.map("Main.Treeview",         background=[("selected", COR_MED)])

        s.configure("Titulo.TLabel",
                    font=("Calibri", 16, "bold"), foreground=COR_AZUL, background=COR_BG)
        s.configure("SubTit.TLabel",
                    font=("Calibri", 11), foreground="#666666", background=COR_BG)
        s.configure("Res.TLabel",
                    font=("Calibri", 13, "bold"), foreground=COR_AZUL, background=COR_BG)
        s.configure("ResST.TLabel",
                    font=("Calibri", 15, "bold"), foreground="#C00000", background=COR_BG)
        s.configure("ResTotal.TLabel",
                    font=("Calibri", 16, "bold"), foreground=COR_VERDE, background=COR_BG)
        s.configure("Calc.TButton",
                    font=("Calibri", 12, "bold"), foreground=COR_W,
                    background=COR_VERDE, padding=8)
        s.map("Calc.TButton",
              background=[("active", "#185A34"), ("disabled", "#AAAAAA")])
        s.configure("Btn.TButton", font=("Calibri", 11), padding=6)

        # Tags treeview
        self.TAG_SIM   = "sim"
        self.TAG_NAO   = "nao"
        self.TAG_VAZIO = "vazio"

    # ----------------------------------------------------------
    def _interface(self):
        main = ttk.Frame(self.root, padding="15 10 15 10")
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        # ── CABEÇALHO ──────────────────────────────────────────
        frm_cab = ttk.Frame(main)
        frm_cab.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        frm_cab.columnconfigure(0, weight=1)

        ttk.Label(frm_cab,
                  text="MÓDULO DE VENDAS — CÁLCULO ST INTERESTADUAL",
                  style="Titulo.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(frm_cab,
                  text="Fase 2: Saídas | Origem: Rio de Janeiro (RJ) | Tabela ICMS 2026",
                  style="SubTit.TLabel").grid(row=1, column=0, sticky="w")

        ttk.Button(frm_cab, text="📋 Tutorial",
                   command=self._tutorial, style="Btn.TButton").grid(
            row=0, column=1, padx=(10, 0))
        ttk.Button(frm_cab, text="🔄 Recarregar",
                   command=self._recarregar, style="Btn.TButton").grid(
            row=0, column=2, padx=5)

        # ── PAINEL DE OPERAÇÃO (UF + Importado) ───────────────
        frm_op = ttk.LabelFrame(
            main,
            text=" ① Selecione a UF de Destino e tipo de Mercadoria ",
            padding="10 6"
        )
        frm_op.grid(row=1, column=0, sticky="ew", pady=(0, 4))

        ttk.Label(frm_op, text="UF Destino:",
                  font=("Calibri", 12, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.uf_var = tk.StringVar(value="SP")
        combo_uf = ttk.Combobox(frm_op, textvariable=self.uf_var,
                                 values=TODOS_ESTADOS,
                                 state="readonly", width=6,
                                 font=("Calibri", 12))
        combo_uf.grid(row=0, column=1, sticky="w")
        combo_uf.bind("<<ComboboxSelected>>", lambda e: self._on_uf_changed())

        # Informação auto da UF
        self.uf_info_var = tk.StringVar()
        ttk.Label(frm_op, textvariable=self.uf_info_var,
                  font=("Calibri", 11), foreground="#2E75B6").grid(
            row=0, column=2, sticky="w", padx=(15, 40))

        # Importado checkbox
        self.importado_var = tk.BooleanVar(value=False)
        chk = ttk.Checkbutton(
            frm_op,
            text="Mercadoria IMPORTADA  (alíq. interestadual = 4%)",
            variable=self.importado_var,
            command=self._on_uf_changed,
            style="TCheckbutton"
        )
        chk.grid(row=0, column=3, sticky="w", padx=(10, 0))

        # Consumidor Final / Contribuinte
        ttk.Label(frm_op, text="Cons. Final:").grid(row=0, column=5, sticky="w", padx=(15, 6))
        self.cons_final_var = tk.StringVar(value="NAO")
        combo_cf = ttk.Combobox(frm_op, textvariable=self.cons_final_var, values=["SIM", "NAO"], state="readonly", width=5)
        combo_cf.grid(row=0, column=6, sticky="w")
        combo_cf.bind("<<ComboboxSelected>>", lambda e: self._on_uf_changed())

        ttk.Label(frm_op, text="Contrib. (IE):").grid(row=0, column=7, sticky="w", padx=(10, 6))
        self.contribuinte_var = tk.StringVar(value="NAO")
        combo_ie = ttk.Combobox(frm_op, textvariable=self.contribuinte_var, values=["SIM", "NAO"], state="readonly", width=5)
        combo_ie.grid(row=0, column=8, sticky="w")
        combo_ie.bind("<<ComboboxSelected>>", lambda e: self._on_uf_changed())

        # Separador
        ttk.Separator(frm_op, orient="vertical").grid(row=0, column=9, sticky="ns", padx=20)

        # Filtro busca
        ttk.Label(frm_op, text="Buscar:").grid(row=0, column=10, sticky="w", padx=(0, 6))
        self.busca_var = tk.StringVar()
        entry_busca = ttk.Entry(frm_op, textvariable=self.busca_var, width=15)
        entry_busca.grid(row=0, column=11, sticky="w")
        entry_busca.bind("<KeyRelease>", lambda e: self._refresh())

        # Filtro Acordo ST
        ttk.Label(frm_op, text="Acordo ST:").grid(
            row=0, column=12, sticky="w", padx=(15, 6))
        self.st_filtro_var = tk.StringVar(value="TODOS")
        ttk.Combobox(frm_op, textvariable=self.st_filtro_var,
                     values=["TODOS", "SIM", "NAO", "(vazio)"],
                     state="readonly", width=9).grid(row=0, column=13, sticky="w")
        self.st_filtro_var.trace("w", lambda *_: self._refresh())

        self._on_uf_changed()  # inicializar info UF

        # ── TREEVIEW ② ──────────────────────────────────────────
        frm_tree = ttk.LabelFrame(
            main, text=" ② Selecione o Produto / NCM ", padding="4 2")
        frm_tree.grid(row=2, column=0, sticky="nsew", pady=(0, 4))
        frm_tree.columnconfigure(0, weight=1)
        frm_tree.rowconfigure(0, weight=1)

        cols = ("NCM", "Produto", "UF", "ConsF", "IE", "ST", "MVA")
        self.tree = ttk.Treeview(frm_tree, columns=cols, show="headings",
                                  height=12, style="Main.Treeview",
                                  selectmode="browse")

        specs = [
            ("NCM",    "NCM",              90,  True),
            ("Produto","Produto / Descrição",330, False),
            ("UF",     "UF",              50,  True),
            ("ConsF",  "Cons.F?",         65,  True),
            ("IE",     "IE?",             45,  True),
            ("ST",     "Acordo ST?",       90,  True),
            ("MVA",    "MVA (%)",          80,  True),
        ]
        for col, head, w, stretch in specs:
            self.tree.heading(col, text=head)
            self.tree.column(col, width=w, minwidth=40,
                             anchor="center" if col != "Produto" else "w",
                             stretch=stretch)

        sb_y = ttk.Scrollbar(frm_tree, orient="vertical",   command=self.tree.yview)
        sb_x = ttk.Scrollbar(frm_tree, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        sb_y.grid(row=0, column=1, sticky="ns")
        sb_x.grid(row=1, column=0, sticky="ew")

        self.tree.tag_configure(self.TAG_SIM,   background="#E8F5E9")  # verde
        self.tree.tag_configure(self.TAG_NAO,   background="#FDECEA")  # vermelho
        self.tree.tag_configure(self.TAG_VAZIO, background="#FFF8E1")  # amarelo

        self.tree.bind("<<TreeviewSelect>>", self._on_selecionar)

        # ── INFO SELEÇÃO ────────────────────────────────────────
        self.info_var = tk.StringVar(value="⬆  Selecione UF, marque se importado e clique no produto.")
        ttk.Label(main, textvariable=self.info_var,
                  font=("Calibri", 11, "bold"), foreground="#2E75B6").grid(
            row=3, column=0, sticky="w", pady=(0, 2))

        ttk.Separator(main, orient="horizontal").grid(
            row=4, column=0, sticky="ew", pady=4)

        # ── PAINEL CÁLCULO ③ ────────────────────────────────────
        frm_calc = ttk.LabelFrame(
            main, text=" ③ Informe os Valores e Calcule ", padding="10 6")
        frm_calc.grid(row=5, column=0, sticky="ew")

        # Valor mercadoria
        ttk.Label(frm_calc, text="Valor Mercadoria (R$):").grid(
            row=0, column=0, sticky="e", padx=(0, 6))
        self.valor_var = tk.StringVar()
        self.entry_valor = ttk.Entry(frm_calc, textvariable=self.valor_var, width=14,
                                      font=("Calibri", 11))
        self.entry_valor.grid(row=0, column=1, sticky="w")
        self.entry_valor.bind("<Return>", lambda e: self._calcular())

        # Frete
        ttk.Label(frm_calc, text="Frete / Despesas (R$):").grid(
            row=0, column=2, sticky="e", padx=(18, 6))
        self.frete_var = tk.StringVar(value="0,00")
        ttk.Entry(frm_calc, textvariable=self.frete_var, width=12,
                  font=("Calibri", 11)).grid(row=0, column=3, sticky="w")

        # MVA
        ttk.Label(frm_calc, text="MVA (%):").grid(
            row=0, column=4, sticky="e", padx=(18, 6))
        self.mva_var = tk.StringVar()
        self.entry_mva = ttk.Entry(frm_calc, textvariable=self.mva_var, width=10,
                                    font=("Calibri", 11))
        self.entry_mva.grid(row=0, column=5, sticky="w")
        ttk.Label(frm_calc, text="(da planilha — editável)",
                  foreground="#888888",
                  font=("Calibri", 9, "italic")).grid(row=0, column=6, sticky="w", padx=3)

        # Botões
        self.btn_calc = ttk.Button(frm_calc, text="💲  CALCULAR ST",
                                    command=self._calcular,
                                    style="Calc.TButton", state="disabled")
        self.btn_calc.grid(row=0, column=7, padx=(20, 5))

        ttk.Button(frm_calc, text="🗑  Limpar",
                   command=self._limpar, style="Btn.TButton").grid(
            row=0, column=8, padx=5)

        # ── DEMONSTRATIVO ───────────────────────────────────────
        frm_demo = ttk.LabelFrame(
            main, text=" 📊 Demonstrativo do Cálculo ST ", padding="12 8")
        frm_demo.grid(row=6, column=0, sticky="ew", pady=(8, 4))
        frm_demo.columnconfigure(1, weight=1)
        frm_demo.columnconfigure(3, weight=1)

        # Coluna esquerda
        esq = [
            ("Alíq. Interestadual (%):", "v_inter"),
            ("Alíq. Interna Destino (%):", "v_intra"),
            ("MVA Ajustada (%):", "v_mva"),
            ("Valor Total NF:", "v_nf_base"),
            ("ICMS Próprio (Origem RJ):", "v_icms_prop"),
        ]
        for i, (lbl, attr) in enumerate(esq):
            ttk.Label(frm_demo, text=lbl).grid(row=i, column=0, sticky="w", pady=2)
            var = tk.StringVar()
            setattr(self, attr, var)
            ttk.Label(frm_demo, textvariable=var, style="Res.TLabel",
                      width=16, anchor="w").grid(row=i, column=1, sticky="w", padx=(8, 30))

        # Linha divisória
        ttk.Separator(frm_demo, orient="vertical").grid(
            row=0, column=2, rowspan=5, sticky="ns", padx=10)

        # Coluna direita
        ttk.Label(frm_demo, text="Base de Cálculo ST:").grid(row=0, column=3, sticky="w")
        self.v_base_st = tk.StringVar()
        ttk.Label(frm_demo, textvariable=self.v_base_st,
                  style="Res.TLabel", width=18, anchor="w").grid(
            row=0, column=4, sticky="w", padx=8)

        ttk.Separator(frm_demo, orient="horizontal").grid(
            row=1, column=3, columnspan=2, sticky="ew", pady=6)

        ttk.Label(frm_demo, text="ICMS-ST A RETER:",
                  font=("Calibri", 13, "bold")).grid(row=2, column=3, sticky="w")
        self.v_icms_st = tk.StringVar()
        ttk.Label(frm_demo, textvariable=self.v_icms_st,
                  style="ResST.TLabel").grid(row=2, column=4, sticky="w", padx=8)

        ttk.Separator(frm_demo, orient="horizontal").grid(
            row=3, column=3, columnspan=2, sticky="ew", pady=6)

        ttk.Label(frm_demo, text="TOTAL DA NOTA FISCAL:",
                  font=("Calibri", 14, "bold")).grid(row=4, column=3, sticky="w")
        self.v_total_nf = tk.StringVar()
        ttk.Label(frm_demo, textvariable=self.v_total_nf,
                  style="ResTotal.TLabel").grid(row=4, column=4, sticky="w", padx=8)

        ttk.Button(frm_demo, text="📋  Copiar Resumo",
                   command=self._copiar, style="Btn.TButton").grid(
            row=4, column=5, padx=(20, 0))

        # ── STATUS BAR ──────────────────────────────────────────
        self.status_var = tk.StringVar()
        ttk.Label(main, textvariable=self.status_var,
                  font=("Calibri", 9), foreground="#666666").grid(
            row=7, column=0, sticky="ew", pady=(4, 0))

    # ----------------------------------------------------------
    # ATUALIZAR INFO DA UF SELECIONADA
    # ----------------------------------------------------------
    def _on_uf_changed(self, *_):
        uf = self.uf_var.get()
        importado = self.importado_var.get()
        aliq_inter = obter_aliq_interestadual(uf, importado)
        aliq_intra = obter_aliq_interna(uf)
        self.uf_info_var.set(
            "→  {uf}:  Interestadual {inter:.0f}%  |  Interna {intra:.1f}%".format(
                uf=uf, inter=aliq_inter, intra=aliq_intra
            ).replace(".", ",")
        )
        # Se já há NCM selecionado, atualiza o btn
        if self.ncm_selecionado.get("AcordoST") == "SIM":
            self.btn_calc.config(state="normal")

    # ----------------------------------------------------------
    # LISTA
    # ----------------------------------------------------------
    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        busca  = self.busca_var.get().strip().lower().replace(".", "")
        st_f   = self.st_filtro_var.get().strip().upper()
        self.dados_filtrados = []

        for item in dados_matriz:
            ncm_n  = item["NCM"].replace(".", "").lower()
            prod_n = item["Produto"].lower()

            if busca and busca not in ncm_n and busca not in prod_n:
                continue
            if st_f == "SIM"     and item["AcordoST"] != "SIM":
                continue
            if st_f == "NAO"     and item["AcordoST"] != "NAO":
                continue
            if st_f == "(VAZIO)" and item["AcordoST"]:
                continue

            self.dados_filtrados.append(item)

        for item in self.dados_filtrados:
            mva_d = "{:.1f}".format(item["MVA"]).replace(".", ",") if item["MVA"] is not None else "—"
            tag   = (self.TAG_SIM   if item["AcordoST"] == "SIM"
                     else self.TAG_NAO   if item["AcordoST"] == "NAO"
                     else self.TAG_VAZIO)

            self.tree.insert("", "end", values=(
                item["NCM"],
                item["Produto"][:75] + ("..." if len(item["Produto"]) > 75 else ""),
                item["UF"] or "—",
                item["ConsFinal"] or "—",
                item["Contribuinte"] or "—",
                item["AcordoST"] or "—",
                mva_d,
            ), tags=(tag,))

        n = len(self.dados_filtrados)
        self.status_var.set(
            "Exibindo {}/{} registros  |  UF: {}  |  Clique no produto para selecionar".format(
                n, len(dados_matriz), self.uf_var.get()))

    # ----------------------------------------------------------
    # SELEÇÃO
    # ----------------------------------------------------------
    def _on_selecionar(self, event=None):
        sel = self.tree.focus()
        if not sel:
            return
        idx = self.tree.index(sel)
        if idx >= len(self.dados_filtrados):
            return

        item = self.dados_filtrados[idx]
        self.ncm_selecionado = item

        # Sincronizar UI com dados da planilha (Redundância)
        if item["UF"] in TODOS_ESTADOS:
            self.uf_var.set(item["UF"])
        
        self.importado_var.set(item["Importado"] == "SIM")
        self.cons_final_var.set(item["ConsFinal"] or "NAO")
        self.contribuinte_var.set(item["Contribuinte"] or "NAO")
        self._on_uf_changed() # atualiza info da UF

        acordo = item["AcordoST"] or "não informado"
        mva_txt = ("{:.2f}%".format(item["MVA"]).replace(".", ",")
                   if item["MVA"] is not None else "não informada")

        self.info_var.set(
            "✅ NCM {ncm} | UF: {uf} | Cons.Final: {cf} | Contrib: {ie} | MVA: {mva}".format(
                ncm=item["NCM"], uf=item["UF"] or "—", cf=item["ConsFinal"] or "—",
                ie=item["Contribuinte"] or "—", mva=mva_txt)
        )

        # Preencher MVA no campo
        if item["MVA"] is not None:
            self.mva_var.set(str(item["MVA"]).replace(".", ","))
        else:
            self.mva_var.set("")

        # Habilitar calcular somente se tem acordo ST
        if item["AcordoST"] == "SIM":
            self.btn_calc.config(state="normal")
        elif item["AcordoST"] == "NAO":
            self.btn_calc.config(state="normal") # Habilitado para DIFAL
        else:
            # Acordo ST não preenchido — habilita mas MVA pode estar vazia
            self.btn_calc.config(state="normal")
            self.info_var.set(
                "⚠  NCM {}  |  Acordo ST não preenchido na planilha. "
                "Preencha antes de calcular.".format(item["NCM"])
            )

    # ----------------------------------------------------------
    # CÁLCULO
    # ----------------------------------------------------------
    def _calcular(self):
        if not self.ncm_selecionado:
            messagebox.showwarning("Atenção", "Selecione um produto na lista primeiro.")
            return

        # Valor mercadoria
        try:
            valor = float(self.valor_var.get().replace(",", "."))
            if valor <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Valor inválido",
                                  "Digite um valor de mercadoria válido (ex: 1500,00).")
            self.entry_valor.focus()
            return

        # Frete
        try:
            frete = float(self.frete_var.get().replace(",", ".") or "0")
        except (ValueError, TypeError):
            frete = 0.0

        # MVA
        try:
            mva = float(self.mva_var.get().replace(",", "."))
            if mva < 0:
                raise ValueError
        except (ValueError, TypeError):
            mva = 0.0

        uf        = self.uf_var.get()
        importado = self.importado_var.get()
        cons_fin  = self.cons_final_var.get() == "SIM"
        contrib   = self.contribuinte_var.get() == "SIM"

        aliq_inter = obter_aliq_interestadual(uf, importado)
        aliq_intra = obter_aliq_interna(uf)

        res = calcular_st(valor, frete, mva, aliq_inter, aliq_intra, cons_fin, contrib)
        self._exibir(res, uf, importado)

    def _exibir(self, res, uf, importado):
        self.v_inter.set(fmt_pct(res["aliq_inter"]))
        self.v_intra.set(fmt_pct(res["aliq_intra"]))
        self.v_mva.set(fmt_pct(res["mva"]))
        self.v_nf_base.set(fmt_brl(res["valor_total_nf"]))
        self.v_icms_prop.set(fmt_brl(res["icms_proprio"]))
        self.v_base_st.set(fmt_brl(res["base_st"]))
        self.v_icms_st.set(fmt_brl(res["icms_st"]))
        self.v_total_nf.set(fmt_brl(res["total_nf"]))

        self.status_var.set(
            "✅  NCM {ncm} → {uf}  |  ST a Reter: {st}  |  Total NF: {tot}".format(
                ncm=self.ncm_selecionado["NCM"], uf=uf,
                st=fmt_brl(res["icms_st"]), tot=fmt_brl(res["total_nf"]))
        )
        self._ultimo_resultado = res
        self._ultimo_item = {"uf": uf, "importado": importado,
                             **self.ncm_selecionado}

    def _limpar_resultado(self):
        for attr in ["v_inter","v_intra","v_mva","v_nf_base",
                     "v_icms_prop","v_base_st","v_icms_st","v_total_nf"]:
            getattr(self, attr, tk.StringVar()).set("")
        self._ultimo_resultado = None
        self._ultimo_item = None

    def _limpar(self):
        self.ncm_selecionado = {}
        self.valor_var.set("")
        self.frete_var.set("0,00")
        self.mva_var.set("")
        self.info_var.set("⬆  Selecione UF, marque se importado e clique no produto.")
        self.btn_calc.config(state="disabled")
        self._limpar_resultado()
        self.busca_var.set("")
        self.st_filtro_var.set("TODOS")
        self._refresh()

    # ----------------------------------------------------------
    # COPIAR
    # ----------------------------------------------------------
    def _copiar(self):
        res  = self._ultimo_resultado
        item = self._ultimo_item
        if not res or not item:
            messagebox.showinfo("Copiar", "Realize o cálculo antes de copiar.")
            return

        imp_txt = "SIM (4%)" if item.get("importado") else "NÃO"
        texto = (
            "=== CÁLCULO ST/DIFAL — SAÍDA INTERESTADUAL ===\n"
            "NCM: {ncm}\n"
            "Produto: {prod}\n"
            "UF Destino: {uf} | Importado: {imp}\n"
            "Modalidade: {modal}\n"
            "-------------------------------------------\n"
            "Alíq. Interestadual: {inter}\n"
            "Alíq. Interna Destino: {intra}\n"
            "MVA Usada (se aplicável): {mva}\n"
            "-------------------------------------------\n"
            "Valor Total NF (Mercadoria): {nf_base}\n"
            "ICMS Próprio (Origem): {icms_prop}\n"
            "Base de Cálculo ST/DIFAL: {base_st}\n"
            "ICMS-ST / DIFAL A RETER: {icms_st}\n"
            "VALOR TOTAL DA NOTA FISCAL: {total}\n"
            "===========================================\n"
            "(Tabela ICMS 2026 | Origem: RJ)"
        ).format(
            ncm=item["NCM"], prod=item["Produto"][:80],
            uf=item["uf"], imp=imp_txt,
            modal=res.get("modalidade", "N/A"),
            inter=fmt_pct(res["aliq_inter"]),
            intra=fmt_pct(res["aliq_intra"]),
            mva=fmt_pct(res["mva"]),
            nf_base=fmt_brl(res["valor_total_nf"]),
            icms_prop=fmt_brl(res["icms_proprio"]),
            base_st=fmt_brl(res["base_st"]),
            icms_st=fmt_brl(res["icms_st"]),
            total=fmt_brl(res["total_nf"]),
        )
        self.root.clipboard_clear()
        self.root.clipboard_append(texto)
        self.root.update()
        self.status_var.set("✅  Resumo copiado para a área de transferência!")

    # ----------------------------------------------------------
    # RECARREGAR
    # ----------------------------------------------------------
    def _recarregar(self):
        self._limpar()
        total = carregar_dados()
        preench = sum(1 for d in dados_matriz if d["MVA"] is not None and d["AcordoST"])
        self.status_var.set(
            "Recarregado: {} NCMs ({} com ST e MVA preenchidos).".format(total, preench)
        )
        self._refresh()

    # ----------------------------------------------------------
    # TUTORIAL
    # ----------------------------------------------------------
    def _tutorial(self):
        messagebox.showinfo("Tutorial", (
            "TUTORIAL — Cálculo ST: Fase 2 (Saídas)\n"
            "========================================\n\n"
            "PASSO ①  —  Selecione a UF de Destino\n"
            "  • Escolha o estado para onde vai a mercadoria\n"
            "  • Marque 'Importado' se a merc. é importada\n"
            "  • As alíquotas são exibidas automaticamente\n\n"
            "PASSO ②  —  Selecione o Produto / NCM\n"
            "  • Use a busca ou role a lista\n"
            "  • Verde = tem Acordo ST | Vermelho = sem ST\n"
            "  • A MVA da planilha é preenchida automaticamente\n\n"
            "PASSO ③  —  Informe os Valores e Calcule\n"
            "  • Digite o Valor da Mercadoria (R$)\n"
            "  • Informe o Frete se houver\n"
            "  • Ajuste a MVA se necessário\n"
            "  • Clique em CALCULAR ST\n\n"
            "PLANILHA EXCEL (contadora preenche):\n"
            "  Col C — Acordo ST?: SIM ou NAO\n"
            "  Col D — MVA (%): ex: 45,00\n"
            "  (UF e Importado são selecionados aqui no programa)\n\n"
            "ALÍQUOTAS (Origem RJ — Tabela 2026):\n"
            "  4%  → Importado (qualquer UF)\n"
            "  7%  → MG, SP, PR, RS, SC, ES\n"
            "  12% → Demais estados\n\n"
            "FÓRMULA:\n"
            "  ICMS Próprio = Total NF × Alíq. Inter.\n"
            "  Base ST      = Total NF × (1 + MVA/100)\n"
            "  ICMS-ST      = Base ST × Alíq. Interna − ICMS Próprio\n"
            "  Total NF     = Mercadoria + Frete + ICMS-ST"
        ))


# ============================================================
# INICIALIZAÇÃO
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    AppST(root)
    root.mainloop()
