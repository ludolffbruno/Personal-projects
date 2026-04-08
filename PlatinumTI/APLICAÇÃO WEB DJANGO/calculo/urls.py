from django.urls import path
from . import views

urlpatterns = [
    path('', views.calcular_custo, name='calcular_custo'),
]