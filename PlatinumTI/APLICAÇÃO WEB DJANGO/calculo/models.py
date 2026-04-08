from django.db import models

class NCM(models.Model):
    ncm = models.CharField(max_length=20)
    produto = models.CharField(max_length=255)
    st_difal = models.CharField(max_length=10)
    aquisicao = models.CharField(max_length=10)  # INTERNA ou EXTERNA
    origem = models.CharField(max_length=20)     # NACIONAL ou IMPORTADA
    st_inclusa = models.CharField(max_length=10) # SIM, NÃO ou ICMS-ST
    origem_importada = models.FloatField(default=0)  # 4%
    origem_nacional = models.FloatField(default=0)   # 12%
    origem_rj = models.FloatField(default=0)         # 22%
    reducao = models.CharField(max_length=3, default="NÃO")  # SIM ou NÃO
    valor_reducao = models.FloatField(default=0)

    def __str__(self):
        return self.ncm