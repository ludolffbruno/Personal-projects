"""
Recria MATRIZ_ST_SAIDA_ver1.xlsx — v1 (SIMPLES / REDUNDANTE)
- Colunas: NCM, Produto, UF Destino, Acordo ST?, Importado?, MVA, Inter, Intra, Obs.
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# === DADOS ===
TODOS_ESTADOS = ["AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"]

# Carregar NCMs
wb_src = openpyxl.load_workbook("Fase 2 - ST Saida/MATRIZ_ST_SAIDA-j.xlsx", data_only=True)
ws_src = wb_src.active

ncm_dados = []
for row in ws_src.iter_rows(min_row=3, values_only=True):
    if row[0]: ncm_dados.append((str(row[0]).strip(), str(row[1]).strip()))

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "MATRIZ_SAIDA"

# Styles
azul_esc = "1F3864"; azul_med = "2E75B6"; amarelo = "FFF2CC"; verde_cl = "E2EFDA"; azul_cl = "DDEEFF"
f_titulo = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
f_ama = Font(name="Calibri", size=10, bold=True, color="1F3864")
borda = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

# Hidden lists
for i, uf in enumerate(TODOS_ESTADOS): ws.cell(4+i, 26, uf)
ws.cell(4, 27, "SIM"); ws.cell(5, 27, "NAO")
ws.column_dimensions["Z"].hidden = True; ws.column_dimensions["AA"].hidden = True

ws.merge_cells("A1:I1"); ws["A1"] = "MATRIZ ST SAIDAS - v1 SIMPLES"; ws["A1"].font = f_titulo; ws["A1"].fill = PatternFill("solid", fgColor=azul_esc)

cols = [("A","NCM",12,azul_esc,f_titulo), ("B","Produto",55,azul_esc,f_titulo), ("C","UF Destino",12,amarelo,f_ama), ("D","Acordo ST?",12,verde_cl,f_ama), ("E","Importado?",12,amarelo,f_ama), ("F","MVA (%)",15,amarelo,f_ama), ("G","Inter (%)",14,azul_med,f_titulo), ("H","Intra (%)",14,azul_med,f_titulo), ("I","Obs",20,azul_esc,f_titulo)]
for l, h, w, c, f in cols:
    ws[l+"3"] = h; ws[l+"3"].font = f; ws[l+"3"].fill = PatternFill("solid", fgColor=c); ws.column_dimensions[l].width = w; ws[l+"3"].border = borda

dv_uf = DataValidation(type="list", formula1="MATRIZ_SAIDA!$Z$4:$Z$29"); ws.add_data_validation(dv_uf); dv_uf.sqref = "C4:C5000"
dv_sn = DataValidation(type="list", formula1="MATRIZ_SAIDA!$AA$4:$AA$5"); ws.add_data_validation(dv_sn); dv_sn.sqref = "D4:E5000"

f_g = '=IF(E{r}="SIM",4,IF(OR(C{r}="MG",C{r}="SP",C{r}="PR",C{r}="RS",C{r}="SC",C{r}="ES"),7,12))'
U_L = '"AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RN","RS","RO","RR","SC","SP","SE","TO"'
V_L = "19,20,20,18,20.5,20,20,17,19,23,17,17,18,19,20,19.5,20.5,22.5,20,17,19.5,20,17,18,20,20"
f_h = '=IF(C{r}="","",IFERROR(CHOOSE(MATCH(C{r},{{'+U_L+'}},0),'+V_L+'),"?"))'

for i, (n, p) in enumerate(ncm_dados):
    r = i+4
    ws.cell(r, 1, n); ws.cell(r, 2, p)
    for c in [3,5,6]: ws.cell(r, c).fill = PatternFill("solid", fgColor=amarelo)
    ws.cell(r, 4).fill = PatternFill("solid", fgColor=verde_cl)
    ws.cell(r, 7, f_g.format(r=r)).fill = PatternFill("solid", fgColor=azul_cl)
    ws.cell(r, 8, f_h.format(r=r)).fill = PatternFill("solid", fgColor=azul_cl)
    for c in range(1,10): ws.cell(r, c).border = borda

ws.freeze_panes = "A4"; ws.auto_filter.ref = "A3:I3"
wb.save("Fase 2 - ST Saida/MATRIZ_ST_SAIDA_ver1.xlsx")
print("VER1 GERADA!")
