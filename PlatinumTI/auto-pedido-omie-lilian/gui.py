import tkinter as tk
from tkinter import scrolledtext
import threading
import sys
import logging
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl') # Ignora avisos do openpyxl  

logging.basicConfig(
    filename="log_execucao.txt",  # Nome do arquivo de log
    filemode="a",  # 'a' para anexar, 'w' para sobrescrever a cada execução
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.DEBUG  # ou INFO se quiser menos detalhes
)

# Importa o script principal da sua automação
# Certifique-se de que main_automation_script.py está na mesma pasta
try:
    from main_automation_script import run_automation
except ImportError:
    # Caso o import falhe, exibe um erro
    def run_automation():
        print("Erro: O arquivo 'main_automation_script.py' não foi encontrado.")
        print("Certifique-se de que ele está na mesma pasta que este arquivo.")


class TextRedirector(object):
    """
    Classe para redirecionar a saída do console (print)
    para o widget de texto da interface.
    """
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str_text):
        self.widget.insert(tk.END, str_text, (self.tag,))
        self.widget.see(tk.END)  # Rola automaticamente para o final

    def flush(self):
        pass


class AutomationGUI(tk.Tk):
    """
    Interface Gráfica para a Automação de Pedidos Omie.
    """
    def __init__(self):
        super().__init__()
        self.title("Automação de Pedidos de Venda Omie - By Mr.Ludolff")
        self.geometry("600x450")
        self.configure(padx=10, pady=10)

        # Frame principal
        main_frame = tk.Frame(self, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = tk.Label(main_frame, text="Automação PDF em Planilha para Omie", font=("Helvetica", 16, "bold"), bg="#f0f0f0")
        title_label.pack(pady=(0, 20))

        # Caminhos das pastas
        paths_frame = tk.Frame(main_frame, bg="#f0f0f0")
        paths_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(paths_frame, text=f"Pasta de Entrada (PDFs): Notas/", bg="#f0f0f0").pack(anchor="w")
        tk.Label(paths_frame, text=f"Pasta de Saída (Planilhas): Planilhas/", bg="#f0f0f0").pack(anchor="w")
        
        # Botão de Iniciar
        self.start_button = tk.Button(main_frame, text="Iniciar Automação",
                                      command=self.start_automation_thread,
                                      font=("Helvetica", 12, "bold"),
                                      bg="#4CAF50", fg="white",
                                      activebackground="#45a049",
                                      activeforeground="white",
                                      relief=tk.RAISED,
                                      bd=3,
                                      padx=20, pady=10)
        self.start_button.pack(pady=20, fill=tk.X)

        # Área de log
        log_label = tk.Label(main_frame, text="Log e Status:", font=("Helvetica", 10, "bold"), bg="#f0f0f0")
        log_label.pack(pady=(10, 5), anchor="w")
        
        self.log_text = scrolledtext.ScrolledText(main_frame, height=10, state=tk.DISABLED, wrap=tk.WORD, bg="#ffffff")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Redireciona a saída do console para a área de log
        sys.stdout = TextRedirector(self.log_text)

    def start_automation_thread(self):
        """
        Inicia a automação em uma thread separada para não travar a UI.
        """
        self.start_button.config(state=tk.DISABLED, text="Processando...")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        print("Iniciando o processo de automação...")

        # Cria e inicia a thread para a automação
        automation_thread = threading.Thread(target=self.run_automation_and_finish)
        automation_thread.daemon = True # Permite que a thread seja fechada com a janela principal
        automation_thread.start()

    def run_automation_and_finish(self):
        """
        Executa a função de automação e habilita o botão novamente.
        """
        try:
            run_automation()
        except Exception as e:
            print(f"\n[ERRO FATAL] Ocorreu um erro na automação: {e}")
        finally:
            print("\nAutomação finalizada.")
            self.start_button.config(state=tk.NORMAL, text="Iniciar Automação")
            self.log_text.config(state=tk.DISABLED)


if __name__ == "__main__":
    app = AutomationGUI()
    app.mainloop()