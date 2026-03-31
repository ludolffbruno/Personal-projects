# -*- coding: utf-8 -*-
"""
Projeto: Platinum NF Automation
Desenvolvido por: Mr Ludolff (Bruno Ludolff)
Descrição: Interface Gráfica para monitoramento e processamento de NF-e via Microsoft Graph API.
Licença: Este software é de propriedade intelectual de Bruno Ludolff. 
Uso restrito e autorizado apenas para fins específicos de automação interna.
Todos os direitos reservados © 2026.
"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import os
import subprocess
import webbrowser
from urllib.parse import urlparse, parse_qs
import graph_email_monitor as monitor

class NFAutomationGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Platinum NF Automation - By Mr.Ludolff")
        self.geometry("1100x800")
        self.configure(bg="#f0f2f5")
        
        self.monitor_thread = None
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.create_widgets()
        self.update_log("Sistema iniciado. Aguardando comando...")

    def create_widgets(self):
        # Header
        header = tk.Frame(self, bg="#1a73e8", height=60)
        header.pack(fill="x")
        tk.Label(header, text="Platinum NF Automation", fg="white", bg="#1a73e8", font=("Segoe UI", 16, "bold")).pack(pady=15)

        # Control Panel
        ctrl_frame = tk.LabelFrame(self, text=" Painel de Controle ", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        ctrl_frame.pack(fill="x", padx=20, pady=10)

        self.btn_sync = ttk.Button(ctrl_frame, text="Iniciar Sincronização", command=self.toggle_sync)
        self.btn_sync.pack(side="left", padx=5)

        ttk.Button(ctrl_frame, text="Abrir Pasta de Notas", command=self.open_folder).pack(side="left", padx=5)
        ttk.Button(ctrl_frame, text="Limpar Log", command=self.clear_log).pack(side="right", padx=5)

        # Timer Display
        self.timer_var = tk.StringVar(value="")
        tk.Label(ctrl_frame, textvariable=self.timer_var, font=("Segoe UI", 9, "italic"), fg="#666").pack(side="left", padx=20)

        # Auth Section (Hidden by default, shown if needed)
        self.auth_frame = tk.Frame(self, padx=20)
        self.auth_label = tk.Label(self.auth_frame, text="Autenticação Necessária!", fg="red", font=("Segoe UI", 10, "bold"))
        self.auth_label.pack(pady=5)
        
        self.url_entry = ttk.Entry(self.auth_frame, width=110) # Aumentado para melhor visibilidade
        self.url_entry.pack(pady=5)
        self.url_entry.insert(0, "Cole a URL de callback aqui...")
        self.url_entry.bind("<FocusIn>", lambda e: self.url_entry.delete(0, tk.END))
        
        ttk.Button(self.auth_frame, text="Confirmar URL", command=self.submit_auth).pack(pady=5)

        # Log Area
        log_frame = tk.LabelFrame(self, text=" Logs do Sistema ", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 9))
        self.log_area.pack(fill="both", expand=True)

        # Footer/Status
        self.status_var = tk.StringVar(value="Status: Parado")
        status_bar = tk.Label(self, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W, padx=10)
        status_bar.pack(fill="x", side="bottom")

    def update_log(self, message):
        def _update():
            self.log_area.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
            self.log_area.see(tk.END)
        self.after(0, _update)

    def update_timer_gui(self, seconds):
        def _update():
            if seconds > 0:
                mins, secs = divmod(seconds, 60)
                self.timer_var.set(f"Próxima sincronização em: {mins:02d}:{secs:02d}")
            else:
                self.timer_var.set("")
        self.after(0, _update)

    def toggle_sync(self):
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.stop_sync()
        else:
            self.start_sync()

    def start_sync(self):
        monitor.STOP_MONITOR = False
        self.btn_sync.configure(text="Pausar Sincronização")
        self.status_var.set("Status: Monitorando...")
        self.auth_frame.pack_forget()
        
        self.monitor_thread = threading.Thread(target=self.run_monitor, daemon=True)
        self.monitor_thread.start()

    def stop_sync(self):
        monitor.STOP_MONITOR = True
        self.btn_sync.configure(text="Iniciar Sincronização")
        self.status_var.set("Status: Parando...")
        self.timer_var.set("")

    def run_monitor(self):
        res = monitor.process_emails(callback=self.update_log, timer_callback=self.update_timer_gui)
        if res == "AUTH_REQUIRED":
            self.show_auth_ui()
        self.status_var.set("Status: Parado")
        self.btn_sync.configure(text="Iniciar Sincronização")
        self.timer_var.set("")

    def show_auth_ui(self):
        self.auth_frame.pack(padx=20, pady=5)
        auth_url = monitor.get_authorization_url(monitor.TENANT_ID, monitor.CLIENT_ID, monitor.SCOPES)
        webbrowser.open(auth_url)
        self.update_log("⚠️ Autenticação necessária.")
        self.update_log(f"🔗 LINK PARA COPIAR: {auth_url}")
        self.update_log("💡 DICA: Copie o link acima e cole em uma aba ANÔNIMA para logar na conta correta.")
        self.update_log("👉 Após o login, a página dará erro. Cole a URL final no campo acima.")

    def submit_auth(self):
        url = self.url_entry.get().strip()
        try:
            parsed_url = urlparse(url)
            query_params = parse_qs(parsed_url.query)
            if 'code' not in query_params:
                messagebox.showerror("Erro", "URL inválida. O código não foi encontrado.")
                return
            
            code = query_params['code'][0]
            tokens = monitor.get_access_token_from_code(monitor.TENANT_ID, monitor.CLIENT_ID, code)
            
            if tokens:
                monitor.save_tokens(tokens["access_token"], tokens["refresh_token"], tokens["expires_in"])
                self.update_log("✅ Autenticação realizada com sucesso!")
                self.auth_frame.pack_forget()
                self.start_sync()
            else:
                messagebox.showerror("Erro", "Falha ao obter tokens. Tente novamente.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar URL: {e}")

    def open_folder(self):
        if not os.path.exists(monitor.SAVE_DIRECTORY):
            os.makedirs(monitor.SAVE_DIRECTORY, exist_ok=True)
        os.startfile(monitor.SAVE_DIRECTORY)

    def clear_log(self):
        self.log_area.delete(1.0, tk.END)

if __name__ == "__main__":
    from datetime import datetime
    app = NFAutomationGUI()
    app.mainloop()
