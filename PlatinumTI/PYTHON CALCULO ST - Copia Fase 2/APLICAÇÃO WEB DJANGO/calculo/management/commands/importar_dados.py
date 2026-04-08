 
from django.core.management.base import BaseCommand
from calculo.models import NCM
import openpyxl

class Command(BaseCommand):
    help = 'Importa dados da planilha para o modelo NCM'

    def handle(self, *args, **kwargs):
        workbook = openpyxl.load_workbook("CALCULO_ST_DIFAL_COMPLETA.xlsx", data_only=True)
        sheet = workbook["COMPLETA"]
        for row in sheet.iter_rows(min_row=2, values_only=True):
            NCM.objects.create(
                ncm=str(row[0]),
                produto=row[1],
                st_difal=row[2] or "N/A",
                aquisicao=row[3],
                origem=row[4],
                st_inclusa=row[5],
                origem_importada=row[6] or 0,
                origem_nacional=row[7] or 0,
                origem_rj=row[8] or 0,
                reducao=row[13] or "NÃO",
                valor_reducao=row[14] or 0
            )
        self.stdout.write(self.style.SUCCESS('Dados importados com sucesso!'))