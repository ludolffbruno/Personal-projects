from django.shortcuts import render

from django.shortcuts import render
from .models import NCM

def calcular_custo(request):
    ncm_list = NCM.objects.all()
    resultado = None
    resumo = ""

    if request.method == 'POST':
        ncm_input = request.POST.get('ncm', '').replace('.', '')
        custo_unitario = float(request.POST.get('custo_unitario', '0').replace(',', '.'))
        st_inclusa = request.POST.get('st_inclusa') == 'on'
        origem_icms = request.POST.get('origem_icms', '0%')

        try:
            ncm = NCM.objects.get(ncm__replace('.', '') == ncm_input)
            origem_icms_value = 0

            if ncm.st_difal == "C/ST" and not st_inclusa:
                origem_icms_value = {
                    "4%": ncm.origem_importada / 100,
                    "12%": ncm.origem_nacional / 100,
                    "22%": ncm.origem_rj / 100
                }.get(origem_icms, 0)
            else:
                origem_icms_value = {
                    "IMPORTADA": ncm.origem_importada / 100,
                    "NACIONAL": ncm.origem_nacional / 100
                }.get(ncm.origem, 0)

            if ncm.st_difal == "C/ST":
                custo_st_interes_estadual = 0 if st_inclusa else custo_unitario * origem_icms_value
                custo_st_interna = custo_unitario * ncm.origem_rj / 100 if ncm.aquisicao == "INTERNA" and not st_inclusa else 0
            else:  # S/ST
                origem_icms_value = {
                    "4%": ncm.origem_importada / 100,
                    "12%": ncm.origem_nacional / 100,
                    "22%": ncm.origem_rj / 100
                }.get(origem_icms, 0)
                custo_st_interes_estadual = custo_unitario * origem_icms_value
                custo_st_interna = 0

            custo_final = custo_unitario + custo_st_interes_estadual + custo_st_interna
            if ncm.reducao == "SIM":
                custo_final -= ncm.valor_reducao

            resultado = f"{custo_final:.2f}".replace('.', ',')
            resumo = f"NCM {ncm.ncm} | Produto: {ncm.produto}"
        except NCM.DoesNotExist:
            resumo = "NCM não encontrado."
        except ValueError:
            resumo = "Erro nos dados de entrada."

    return render(request, 'calculo/calcular.html', {
        'ncm_list': ncm_list,
        'resultado': resultado,
        'resumo': resumo
    })