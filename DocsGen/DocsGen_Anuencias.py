import os
import tkinter as tk
import comtypes.client
import re
import traceback
import uuid
import time
import glob
from plyer import notification
from tkinter import ttk, filedialog, messagebox
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from ttkbootstrap import Style
from PIL import Image, ImageTk
from datetime import datetime
import threading
import pythoncom

# Obter a data de hoje
hoje = datetime.today()
# Formatar a data como DDMMAAAA
data_hoje = hoje.strftime("%d%m%Y")

# Obtém o diretório onde o script está localizado
caminho_base = os.path.dirname(os.path.abspath(__file__))

# Lista de meses em português
meses = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

# Configuração dos Cargos
CONFIG_CARGOS = {
    "Técnico O&M": {
        "pasta": "tec_oem",
        "sufixo": "_TEC_OM",
        "docs": ["NR10", "NR10 SEP", "NR12", "NR33", "NR35"]
    },
    "Supervisor O&M": {
        "pasta": "sup_oem",
        "sufixo": "_sup_OM",
        "docs": ["NR10", "NR10 SEP", "NR12", "NR33", "NR35"]
    },
    "Técnico PA": {
        "pasta": "tec_pa",
        "sufixo": "_TEC_PA",
        "docs": ["NR10", "NR10 SEP", "NR12", "NR33", "NR35"]
    },
    "Almoxarife": {
        "pasta": "almox",
        "sufixo": "_almox",
        "docs": ["NR10", "NR10 SEP", "NR12", "NR33"]
    },
    "Técnico PA Especializado": {
        "pasta": "tec_pa_esp",
        "sufixo": "_TEC_PA_ESP",
        "docs": ["NR10", "NR10 SEP", "NR12", "NR33", "NR35"]
    },
    "Técnico Segurança": {
        "pasta": "tec_seg",
        "sufixo": "_TEC_SEG",
        "docs": ["NR10", "NR10 SEP", "NR12", "NR33", "NR35"]
    },
    "Consultor Administrativo": {
        "pasta": "consultor_adm",
        "sufixo": "_CONSULTOR_ADM",
        "docs": ["NR10", "NR10 SEP", "NR12", "NR33", "NR35"]
    },
    "Técnico Serv. Esp. Op.": {
        "pasta": "tec_serv_esp_op",
        "sufixo": "_TEC_SERV_ESP_OP",
        "docs": ["NR10", "NR10 SEP", "NR12", "NR33", "NR35"]
    }
}

class genAnuencias:
    def __init__(self, root):
        # Crie a interface gráfica
        style = Style(theme='darkly')
        self.root = root
        root = style.master
        root.title("Gerador de Anuências - Unificado")
        frame = ttk.Frame(root)
        frame['padding'] = (10, 10, 10, 10)
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        self.dir_pasta_selecionada = ""

        # Carregar a imagem
        try:
            self.bg_image = Image.open(r"C:\PyProjects\DocsGen\bkg_anuencias.png")
            self.bg_photo = ImageTk.PhotoImage(self.bg_image)
            ttk.Label(frame, image=self.bg_photo).grid(row=0, column=0, columnspan=6, pady=(0, 20))
        except Exception:
            pass

        #LabelFrame - Dados Documento
        lblframe_dados = ttk.LabelFrame(frame, text="Dados Documento:", padding=10)
        lblframe_dados.grid(row=1, column=0, columnspan=6, sticky="ew", padx=10, pady=5)   

        ttk.Label(lblframe_dados, text="Funcionário: ").grid(row=0, column=0, sticky=tk.W)
        self.entr_funcionario = ttk.Entry(lblframe_dados, width=60)
        self.entr_funcionario.grid(row=0, column=1, columnspan=4, sticky=tk.W, pady=5)

        ttk.Label(lblframe_dados, text="CPF: ").grid(row=1, column=0, sticky=tk.W)
        self.entr_CPF = ttk.Entry(lblframe_dados, width=30)
        self.entr_CPF.grid(row=1, column=1, sticky=tk.W, pady=5)

        ttk.Label(lblframe_dados, text="Iniciais: ").grid(row=1, column=2, sticky=tk.W)
        self.entr_apelido = ttk.Entry(lblframe_dados, width=15)
        self.entr_apelido.grid(row=1, column=3, sticky=tk.W, pady=5)

        # Seleção de Cargo
        lblframe_cargo = ttk.LabelFrame(frame, text="Seleção de Cargo:", padding=10)
        lblframe_cargo.grid(row=2, column=0, columnspan=6, sticky="ew", padx=10, pady=5)
        
        ttk.Label(lblframe_cargo, text="Cargo: ").grid(row=0, column=0, sticky=tk.W)
        self.combo_cargo = ttk.Combobox(lblframe_cargo, values=list(CONFIG_CARGOS.keys()), state="readonly", width=40)
        self.combo_cargo.grid(row=0, column=1, sticky=tk.W, padx=10)
        self.combo_cargo.bind("<<ComboboxSelected>>", self.atualizar_opcoes_cargo)

        # Frame para Checkboxes (Dinâmico)
        self.lblframe_anuencias = ttk.LabelFrame(frame, text="Anuências Disponíveis:", padding=10)
        self.lblframe_anuencias.grid(row=3, column=0, columnspan=6, sticky="ew", padx=10, pady=5)
        
        self.vars_docs = {} # Dicionário para guardar as variáveis dos checkboxes

        # Opções
        lblframe_opcoes = ttk.LabelFrame(frame, text="Opções: ", padding=10)
        lblframe_opcoes.grid(row=4, column=0, columnspan=6, sticky="ew", padx=10, pady=5)                

        ttk.Label(lblframe_opcoes, text="Salvar em: ").grid(row=0, column=0, sticky=tk.W)      
        ttk.Button(lblframe_opcoes, text="...", command=self.selecionar_pasta).grid(row=0, column=1, sticky=tk.W, padx=(0, 10))   
        self.lbl_pastaselecionada = tk.Label(lblframe_opcoes, text="Selecione a pasta", wraplength=350)
        self.lbl_pastaselecionada.grid(row=0, column=2, columnspan=4, sticky=tk.W, padx=(0,10))

        # Botão Gerar
        ttk.Button(frame, text="Gerar Anuências", command=self.verificar_checkbuttons).grid(row=5, column=0, sticky=tk.W, padx=10, pady=20)
        self.lbl_avisoGeracao = ttk.Label(frame, text="")
        self.lbl_avisoGeracao.grid(row=5, column=1, columnspan=4, sticky=tk.W)

        self.dados = {}

    def atualizar_opcoes_cargo(self, event):
        # Limpar checkboxes antigos
        for widget in self.lblframe_anuencias.winfo_children():
            widget.destroy()
        self.vars_docs.clear()

        cargo = self.combo_cargo.get()
        if cargo in CONFIG_CARGOS:
            docs = CONFIG_CARGOS[cargo]["docs"]
            for i, doc in enumerate(docs):
                var = tk.BooleanVar()
                chk = ttk.Checkbutton(self.lblframe_anuencias, text=doc, variable=var)
                chk.grid(row=0, column=i, padx=10, sticky=tk.W)
                self.vars_docs[doc] = var

    def selecionar_pasta(self):
        pasta_selecionada = filedialog.askdirectory()
        if pasta_selecionada:
            self.lbl_pastaselecionada.config(text=pasta_selecionada)   
            self.dir_pasta_selecionada = pasta_selecionada  

    def substituir_texto(self, paragraph, substituicoes):
        # Substituir o texto nos runs do parágrafo
        for palavra_antiga, palavra_nova in substituicoes.items():
            if palavra_nova is None:
                palavra_nova = ""
            if palavra_antiga in paragraph.text:
                if palavra_antiga == 'NOMEFUNCIONARIO':
                    for run in paragraph.runs:
                        if palavra_antiga in run.text:
                            run.text = run.text.replace(palavra_antiga, palavra_nova)
                            run.bold = True  # Mantém o negrito
                        run.font.name = 'Arial'
                else:
                    for run in paragraph.runs:
                        if palavra_antiga in run.text:
                            run.text = run.text.replace(palavra_antiga, palavra_nova)
                        run.font.name = 'Arial'
        return paragraph

    def substituir_texto_tabela(self, cell, substituicoes):
        # Substituir o texto nas células da tabela
        for paragraph in cell.paragraphs:
            paragraph = self.substituir_texto(paragraph, substituicoes)  # Chamando a função que substitui no parágrafo
        # Centralizar verticalmente o conteúdo da célula
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        return cell

    def validar_cpf(self, cpf: str) -> bool:
            """Valida um CPF informado com pontos e traço."""

            # Remover caracteres não numéricos (pontos e traço)
            cpf = re.sub(r"[\.\-]", "", cpf)

            # Verificar se tem exatamente 11 dígitos
            if not cpf.isdigit() or len(cpf) != 11:
                return False

            # Verificar se todos os dígitos são iguais (ex: 111.111.111-11 é inválido)
            if cpf == cpf[0] * 11:
                return False

            # Cálculo do primeiro dígito verificador
            soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
            digito1 = (soma * 10 % 11) % 10

            # Cálculo do segundo dígito verificador
            soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
            digito2 = (soma * 10 % 11) % 10

            # Verifica se os dígitos calculados são iguais aos do CPF informado
            return digito1 == int(cpf[9]) and digito2 == int(cpf[10])

    def mostrar_progresso(self):
        """Exibe barra de progresso e inicia a geração em thread"""
        self.janela_progress = tk.Toplevel(self.root)
        self.janela_progress.title("Gerando Documentos...")
        self.janela_progress.geometry("400x150")
        self.janela_progress.resizable(False, False)
        self.janela_progress.transient(self.root)
        self.janela_progress.grab_set()
        
        # Centralizar
        self.janela_progress.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (self.janela_progress.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (self.janela_progress.winfo_height() // 2)
        self.janela_progress.geometry(f"+{x}+{y}")
        
        ttk.Label(self.janela_progress, text="Gerando documentos...\nPor favor, aguarde...", justify=tk.CENTER).pack(pady=20)
        
        self.progress = ttk.Progressbar(self.janela_progress, mode='indeterminate', length=350)
        self.progress.pack(pady=10, padx=20)
        self.progress.start()
        
        # Iniciar thread
        threading.Thread(target=self._executar_geracao, daemon=True).start()

    def limpar_temporarios(self, pasta):
        """Remove arquivos temporários da pasta especificada."""
        try:
            padrao = os.path.join(caminho_base, pasta, "temp_*.docx")
            arquivos = glob.glob(padrao)
            for arquivo in arquivos:
                # Tentar deletar com retentativa
                for _ in range(3):
                    try:
                        os.remove(arquivo)
                        break
                    except Exception:
                        time.sleep(0.5)
        except Exception:
            pass

    def _executar_geracao(self):
        pasta = ""
        try:
            pythoncom.CoInitialize()
            cargo = self.combo_cargo.get()
            config = CONFIG_CARGOS.get(cargo)
            
            if not config:
                return

            pasta = config["pasta"]
            sufixo = config["sufixo"]
            
            # Limpar temporários antigos antes de começar
            self.limpar_temporarios(pasta)
            
            # Preparar dados para substituição
            nome_func = self.entr_funcionario.get()
            cpf_func = self.entr_CPF.get()
            
            # Garantir que não sejam None
            if nome_func is None: nome_func = ""
            if cpf_func is None: cpf_func = ""

            self.dados = {
                'NOMEFUNCIONARIO': str(nome_func),
                'DIAANUENCIA': datetime.today().strftime("%d"),
                'MESANUENCIA': str(meses.get(datetime.today().month, datetime.today().strftime("%B"))),
                'ANOANUENCIA': datetime.today().strftime("%Y"),
                'CPFFUNCIONARIO': str(cpf_func)
            }

            for doc_nome, var in self.vars_docs.items():
                if var.get():
                    # Determinar nome do template
                    # Ex: NR10 -> nr10_pasta.docx
                    # Ex: NR10 SEP -> nr10_sep_pasta.docx
                    prefixo = doc_nome.lower().replace(" ", "_") # nr10, nr10_sep
                    template_name = f"{prefixo}_{pasta}.docx"
                    
                    self.gerar_documento(template_name, pasta, sufixo, doc_nome)
                    
                    self.root.after(0, lambda d=doc_nome: notification.notify(
                        title="Aviso",
                        message=f"{d} Gerada com Sucesso!",
                        timeout=5
                    ))

            # Finalização
            self.root.after(0, lambda: [
                self.janela_progress.destroy(),
                self.lbl_avisoGeracao.config(text=""),
                messagebox.showinfo("Concluído", "Documento(s) Salvo(s) com Sucesso!")
            ])

        except Exception as e:
            erro_msg = traceback.format_exc()
            self.root.after(0, lambda: [
                self.janela_progress.destroy(),
                messagebox.showerror("Erro", f"Ocorreu um erro:\n{erro_msg}")
            ])
        finally:
            if pasta:
                self.limpar_temporarios(pasta)
            pythoncom.CoUninitialize()

    def gerar_documento(self, template_name, pasta, sufixo, doc_nome):
        template_path = os.path.join(caminho_base, pasta, template_name)
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template não encontrado: {template_path}")

        doc = Document(template_path)

        for paragraph in doc.paragraphs:
            self.substituir_texto(paragraph, self.dados)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    self.substituir_texto_tabela(cell, self.dados)

        # Nome do arquivo de saída
        # Ex: NR10_Apelido_Data_Sufixo.pdf
        prefixo_saida = doc_nome.replace(" ", "_") # NR10, NR10_SEP
        
        # Validar caracteres inválidos no apelido
        apelido = self.entr_apelido.get()
        if apelido is None: apelido = ""
        apelido = re.sub(r'[<>:"/\\|?*]', '', apelido) # Remove caracteres inválidos
        
        nome_arquivo = f"{prefixo_saida}_{apelido}_{data_hoje}{sufixo}.pdf"
        
        # Gerar nome temporário único para evitar conflitos de arquivo em uso
        unique_id = uuid.uuid4().hex
        temp_docx = os.path.join(caminho_base, pasta, f"temp_{prefixo_saida}_{unique_id}.docx")
        pdf_path = os.path.join(self.dir_pasta_selecionada, nome_arquivo)

        # Verificar se o PDF de destino já está aberto (tentativa simples de abrir para escrita)
        if os.path.exists(pdf_path):
            try:
                with open(pdf_path, 'ab'):
                    pass
            except PermissionError:
                raise PermissionError(f"O arquivo PDF '{nome_arquivo}' parece estar aberto em outro programa. Feche-o e tente novamente.")

        doc.save(temp_docx)

        word = None
        doc_word = None
        try:
            word = comtypes.client.CreateObject('Word.Application')
            # word.Visible = False # Opcional, mas bom para evitar janelas piscando
            doc_word = word.Documents.Open(os.path.abspath(temp_docx))
            
            # Tentar salvar, capturando erros específicos do Word/COM
            try:
                doc_word.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
            except Exception as e:
                if "permission" in str(e).lower() or "access is denied" in str(e).lower():
                     raise PermissionError(f"Erro de permissão ao salvar '{nome_arquivo}'. Verifique se o arquivo está aberto.")
                raise e
                
        except Exception as e:
            raise e
        finally:
            if doc_word:
                try:
                    doc_word.Close()
                except:
                    pass
            if word:
                try:
                    word.Quit()
                except:
                    pass
            
            # Garantir remoção do arquivo temporário
            if os.path.exists(temp_docx):
                try:
                    time.sleep(0.5) # Pequena pausa para liberar o arquivo
                    os.remove(temp_docx)
                except:
                    pass
            
    def verificar_checkbuttons(self):      

        cpf_digitado = self.entr_CPF.get()
        texto_pastaselecionada = self.lbl_pastaselecionada.cget("text")
        cargo = self.combo_cargo.get()

        if not cargo:
            messagebox.showinfo("Alerta!", "Selecione um Cargo!")
            return

        if texto_pastaselecionada == "Selecione a pasta":
            messagebox.showinfo("Alerta!", "Selecione a pasta para salvar o(s) documento(s)!") 
        else:
            if self.validar_cpf(cpf_digitado):    
                self.lbl_avisoGeracao.config(text="Processando...")
                self.mostrar_progresso()
            else:
                messagebox.showinfo("Alerta!", "CPF Inválido!") 
        
# Criar janela do Tkinter
if __name__ == "__main__":
    root = tk.Tk()
    app = genAnuencias(root)
    root.mainloop()
