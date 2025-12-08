import tkinter as tk
import ttkbootstrap as tb
import subprocess
import os
import sys
import threading
from tkinter import ttk, messagebox
from PIL import Image, ImageTk

# Diretório base do script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Encontrar pythonw.exe
PYTHON_DIR = os.path.dirname(sys.executable)
PYTHONW = os.path.join(PYTHON_DIR, "pythonw.exe")
if not os.path.exists(PYTHONW):
    PYTHONW = sys.executable

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("DocsGen - Vestas - v2.050625")
        self.root.geometry("1280x720")
        self.root.resizable(False, False)

        # Estilo do ttkbootstrap
        style = tb.Style(theme="darkly")

        # Canvas com imagem de fundo
        self.canvas = tk.Canvas(self.root, width=1280, height=720)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        bkg_img = Image.open(os.path.join(BASE_DIR, "bkg_docsgen.png"))
        bkg_img = bkg_img.resize((1280, 720), Image.Resampling.LANCZOS)
        self.bkg = ImageTk.PhotoImage(bkg_img)
        self.canvas.create_image(0, 0, image=self.bkg, anchor="nw")

        # Frame principal por cima do canvas
        self.frame_principal = ttk.Frame(self.canvas, padding=20)
        self.frame_principal.place(relx=0.5, rely=0.5, anchor="center")  # Centraliza no canvas

        # Label de boas-vindas
        #self.label = ttk.Label(self.frame_principal, text="DocsGen - Ferramenta de Geração de Documentos", font=("Arial", 14, "bold"))
        #self.label.grid(row=0, column=0, pady=50)

        # Botão de saída
        #self.btn_sair = tb.Button(self.frame_principal, text="Sair", bootstyle="danger", command=self.root.quit)
        #self.btn_sair.grid(row=1, column=0, pady=10)

        # Menu
        menubar = tk.Menu(self.root)

        menu_anuencias = tk.Menu(menubar, tearoff=0)
        menu_anuencias.add_command(label="Gerar Nova", command=self.abrir_anuencias)
        #menu_anuencias.add_command(label="Tecnico de O&M", command=self.anuencias_tec_om)
        #menu_anuencias.add_command(label="Supervisor de O&M", command=self.anuencias_sup_om)
        #menu_anuencias.add_command(label="Tecnico de Pá", command=self.anuencias_tec_pa)
        #menu_anuencias.add_command(label="Tecnico de Pá Especialista", command=self.anuencias_tec_pa_esp)
        #menu_anuencias.add_command(label="Tecnico de Seg. Trabalho", command=self.anuencias_tec_seg)
        #menu_anuencias.add_command(label="Tecnico de Serviços Especiais Operacional", command=self.anuencias_tec_serv_esp_op)
        #menu_anuencias.add_command(label="Consultor Administrativo", command=self.anuencias_consultor_adm)
        #menu_anuencias.add_command(label="Almoxarife", command=self.anuencias_almoxarife)

        menu_os = tk.Menu(menubar, tearoff=0)
        #menu_os.add_command(label="Gerar Nova", command=self.abrir_os)
        menu_os.add_command(label="Gerar OS", command=self.gerador_os)        

        menu_sit = tk.Menu(menubar, tearoff=0)
        menu_sit.add_command(label="Gerar Novo", command=self.abrir_sit)

        menu_lift = tk.Menu(menubar, tearoff=0)
        menu_lift.add_command(label="Gerar Novo", command=self.abrir_lift)

        menu_assinaturas = tk.Menu(menubar, tearoff=0)
        menu_assinaturas.add_command(label="Remover Assinaturas", command=self.abrir_assinaturas)

        menubar.add_cascade(label="Anuências", menu=menu_anuencias)
        menubar.add_cascade(label="OS", menu=menu_os)
        menubar.add_cascade(label="SIT", menu=menu_sit)
        menubar.add_cascade(label="Lift User", menu=menu_lift)
        menubar.add_cascade(label="Assinaturas", menu=menu_assinaturas)
        menubar.add_command(label="Sair", command=self.root.quit)

        self.root.config(menu=menubar)

    def _executar_script_com_progress(self, caminho_script, nome_script):
        """Executa script com barra de progresso em thread separada"""
        janela_progress = tk.Toplevel(self.root)
        janela_progress.title(f"Carregando {nome_script}...")
        janela_progress.geometry("400x150")
        janela_progress.resizable(False, False)
        janela_progress.transient(self.root)
        janela_progress.grab_set()
        
        # Centralizar janela de progresso
        janela_progress.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (janela_progress.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (janela_progress.winfo_height() // 2)
        janela_progress.geometry(f"+{x}+{y}")
        
        # Label informativo
        label = ttk.Label(janela_progress, text=f"Iniciando {nome_script}...\nPor favor, aguarde...", justify=tk.CENTER)
        label.pack(pady=20)
        
        # Barra de progresso
        progress = ttk.Progressbar(janela_progress, mode='indeterminate', length=350)
        progress.pack(pady=10, padx=20)
        progress.start()
        
        # Label de status
        status_label = ttk.Label(janela_progress, text="Carregando...", foreground="gray")
        status_label.pack(pady=5)
        
        def executar():
            try:
                # Verificar se arquivo existe
                if not os.path.exists(caminho_script):
                    self.root.after(0, lambda: messagebox.showerror(
                        "Erro",
                        f"Arquivo não encontrado:\n{caminho_script}"
                    ))
                    janela_progress.destroy()
                    return
                
                # Executar script com stdout e stderr capturados
                processo = subprocess.Popen(
                    [PYTHONW, caminho_script],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True
                )
                
                try:
                    # Aguardar 3 segundos para verificar se há erro na inicialização
                    stdout, stderr = processo.communicate(timeout=3)
                    
                    # Se chegou aqui, o processo terminou (erro ou execução rápida)
                    if processo.returncode == 0:
                        status_label.config(text="Concluído.", foreground="green")
                        self.root.after(500, janela_progress.destroy)
                    else:
                        erro_msg = stderr if stderr else f"Código de erro: {processo.returncode}"
                        self.root.after(0, lambda: messagebox.showerror(
                            "Erro ao executar",
                            f"Erro ao executar {nome_script}:\n\n{erro_msg[:500]}"
                        ))
                        janela_progress.destroy()
                        
                except subprocess.TimeoutExpired:
                    # Timeout significa que o programa continua rodando (sucesso para GUI)
                    status_label.config(text="Programa iniciado!", foreground="green")
                    self.root.after(500, janela_progress.destroy)
                    # O processo continua rodando em segundo plano
                    
            except subprocess.TimeoutExpired:
                processo.kill()
                self.root.after(0, lambda: messagebox.showerror(
                    "Timeout",
                    f"O script {nome_script} demorou muito tempo para responder."
                ))
                janela_progress.destroy()
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "Erro",
                    f"Erro ao executar {nome_script}:\n\n{str(e)}"
                ))
                janela_progress.destroy()
        
        # Executar em thread separada para não travar a interface
        thread = threading.Thread(target=executar, daemon=True)
        thread.start()

    def abrir_anuencias(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DocsGen_Anuencias.py"),
            "Anuências"
        )

    def anuencias_tec_om(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DG_AnuenciasTecOM.py"),
            "Anuências - Técnico O&M"
        )

    def anuencias_sup_om(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DG_AnuenciasSupOM.py"),
            "Anuências - Supervisor O&M"
        )

    def anuencias_tec_pa(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DG_AnuenciasTecPA.py"),
            "Anuências - Técnico Pá"
        )

    def anuencias_tec_pa_esp(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DG_AnuenciasTecPAESP.py"),
            "Anuências - Técnico Pá Especialista"
        )

    def anuencias_tec_seg(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DG_AnuenciasTecSEG.py"),
            "Anuências - Técnico Segurança"
        )

    def anuencias_tec_serv_esp_op(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DG_AnuenciasTecSERVESPCIOP.py"),
            "Anuências - Técnico Serviços Especiais"
        )

    def anuencias_consultor_adm(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DG_AnuenciasTecCONSULTORADM.py"),
            "Anuências - Consultor Administrativo"
        )

    def anuencias_almoxarife(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DG_AnuenciasAlmox.py"),
            "Anuências - Almoxarife"
        )

    def gerador_os(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DocsGen_OS.py"),
            "Gerador de Ordens de Serviço"
        )

    def abrir_sit(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DocsGen_SIT.py"),
            "Gerador SIT"
        )

    def abrir_lift(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DG_LiftUser.py"),
            "Lift User"
        )

    def abrir_assinaturas(self):
        self._executar_script_com_progress(
            os.path.join(BASE_DIR, "DG_RemoveSign.py"),
            "Remover Assinaturas"
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()