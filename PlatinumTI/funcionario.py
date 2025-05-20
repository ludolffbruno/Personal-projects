
import csv
import openpyxl
import tkinter as tk
import time

# Constantes
ARQUIVO_XLSX = '3-funcionario.xlsx'
COLUNA_DESCRICAO = 2
COLUNA_PRECO = 3
COLUNA_CODIGO = 4

class ConsultaPrecosGUI:
    def __init__(self):
        self.carrinho = 0

        self.janela = tk.Tk()
        self.janela.title("Consulta de Preços - By Mr.Ludolff")
        self.janela.geometry(f"{self.janela.winfo_screenwidth()}x{self.janela.winfo_screenheight()}")
        self.janela.resizable(True, True)

        tk.Label(self.janela, text="Código de Barras:", font=("Arial", 14)).pack(pady=5)

        self.codigo_barras_entry = tk.Entry(self.janela, font=("Arial", 14), bg="#F2D0E0", borderwidth=1, relief="solid")
        self.codigo_barras_entry.pack()
        self.codigo_barras_entry.focus_set()

        self.carrinho_label = tk.Label(self.janela, text="Carrinho: R$ 0.00", font=("Arial", 14), wraplength=1000)
        self.carrinho_label.pack(pady=5)

        self.resultados_text = tk.Text(self.janela, font=("Arial", 14), wrap=tk.WORD, width=90, height=19, bg="#F2D0E0", borderwidth=1, relief="solid", padx=30, pady=20, state=tk.DISABLED)
        self.resultados_text.pack(pady=5)

        btn_zerar_carrinho = tk.Button(self.janela, text="Zerar Carrinho", font=("Arial", 9), command=self.zerar_carrinho)
        btn_zerar_carrinho.pack(side=tk.LEFT, pady=10, padx=10)

        btn_sair = tk.Button(self.janela, text="Sair", font=("Arial", 9), command=self.sair)
        btn_sair.pack(side=tk.LEFT, pady=10)

        scrollbar = tk.Scrollbar(self.janela)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar.config(width=30)

        btn_zerar_carrinho.place(relx=0.5, rely=0.89, anchor=tk.W)
        btn_sair.place(relx=0.5, rely=0.89, anchor=tk.E)

        self.aviso_label = tk.Label(self.janela, text="", font=("Arial", 14), wraplength=1000)
        self.aviso_label.pack(pady=5)
#
        self.tempo_inatividade = 1200  # Tempo em segundos
        self.tempo_ultima_consulta = time.time()
#
        self.codigo_barras_entry.bind("<Return>", self.consultar_preco)


    def mostrar_mensagem_apresentacao(self):
        mensagem = """
        Bem-vindo(a) à consulta de preço BELEZA DIVINA FUNCIONÁRIO!

        
        -Use o leitor de código de barras no campo acima.
        
        -Os resultados da consulta serão exibidos nesta caixa de texto.

        -Para zerar o carrinho, leia o código de barras que está fixado no monitor.

        -Para sair do programa, clique no botão "Sair" ou feche a janela.
        """
        self.resultados_text.config(state=tk.NORMAL)
        self.resultados_text.delete('1.0', tk.END)
        self.resultados_text.insert(tk.END, mensagem)
        self.resultados_text.config(state=tk.DISABLED)
        

    def sair(self):
        print("Fechando programa.")
        self.janela.destroy()


    def zerar_carrinho(self):
        self.carrinho = 0
        print("Zerando carrinho.")
        self.carrinho_label.config(text="Carrinho: R$ 0.00")


    def consultar_preco(self, event=None):
        codigo = self.codigo_barras_entry.get()
        self.codigo_barras_entry.delete(0, tk.END)

        # Atualiza o tempo da última consulta
        self.tempo_ultima_consulta = time.time()

        if codigo == "7898357417892" or codigo == "1":
            self.zerar_carrinho()
            return

        codigo_encontrado = False
        resultados = []

        try:
            if ARQUIVO_XLSX.endswith(".xlsx"):
                resultados = self.consultar_preco_xlsx(codigo)
                print("Arquivo .XLSX")
        except FileNotFoundError:
            self.aviso_label.config(text="Arquivo não encontrado.")
        except Exception as e:
            self.aviso_label.config(text="Erro ao ler o arquivo.")


        if codigo_encontrado:
            self.aviso_label.config(text="")
        else:
            self.aviso_label.config(text="Código de barras não encontrado.")

        self.resultados_text.config(state=tk.NORMAL)
        self.resultados_text.delete('1.0', tk.END)

        if resultados:
            for r in resultados:
                descricao_preco = f"{r['descricao']} - R$ {r['preco']}\n\n"
                self.resultados_text.insert(tk.END, descricao_preco)
                
        else:
            self.resultados_text.insert(tk.END, "Nenhum resultado encontrado.")

        self.aviso_label.config(text="")
        self.resultados_text.config(state=tk.DISABLED)
        self.carrinho_label.config(text="Carrinho: R$ {:.2f}".format(self.carrinho))


    def verificar_inatividade(self):
        if time.time() - self.tempo_ultima_consulta >= self.tempo_inatividade:
            self.mostrar_mensagem_apresentacao()

        self.janela.after(1000, self.verificar_inatividade)  # Verifica a inatividade a cada 1 segundo

    def iniciar_verificacao_inatividade(self):
        self.janela.after(1000, self.verificar_inatividade)



    def consultar_preco_xlsx(self, codigo):
        codigo_encontrado = False
        resultados = []

        try:
            workbook = openpyxl.load_workbook(ARQUIVO_XLSX)
            worksheet = workbook.active

            for linha in worksheet.iter_rows(min_row=2, values_only=True):
                if str(linha[COLUNA_CODIGO - 1]) == codigo:
                    codigo_encontrado = True
                    self.carrinho += float(str(linha[COLUNA_PRECO - 1]).replace(',', '.'))
                    resultados.append({'descricao': str(linha[COLUNA_DESCRICAO - 1]), 'preco': str(linha[COLUNA_PRECO - 1])})

            workbook.close()
        except Exception as e:
            print("Erro ao ler o arquivo:", str(e))

        return resultados

gui = ConsultaPrecosGUI()
gui.mostrar_mensagem_apresentacao()
gui.iniciar_verificacao_inatividade()
gui.janela.mainloop()

