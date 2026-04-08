import os
import sys
import tkinter as tk
from tkinter import messagebox

def verificar_arquivos():
    """Verifica se os arquivos necessários estão presentes"""
    arquivos_necessarios = ["CALCULO_ST_DIFAL_COMPLETA.xlsx"]
    arquivos_faltantes = []
    
    for arquivo in arquivos_necessarios:
        if not os.path.exists(arquivo):
            arquivos_faltantes.append(arquivo)
    
    return arquivos_faltantes

def main():
    """Função principal para testar o programa"""
    # Verificar se os arquivos necessários estão presentes
    arquivos_faltantes = verificar_arquivos()
    
    if arquivos_faltantes:
        print(f"ERRO: Os seguintes arquivos estão faltando: {', '.join(arquivos_faltantes)}")
        print("Por favor, certifique-se de que estes arquivos estão no mesmo diretório do programa.")
        return
    
    # Tentar importar o módulo principal
    try:
        # Importar o módulo principal
        print("Iniciando o programa Calc_ST_Platinum_ver16_Expandido.py...")
        
        # Executar o programa
        os.system("python3 Calc_ST_Platinum_ver16_Expandido.py")
        
    except Exception as e:
        print(f"ERRO ao executar o programa: {str(e)}")
        return

if __name__ == "__main__":
    main()
