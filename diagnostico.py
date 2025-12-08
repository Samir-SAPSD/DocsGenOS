#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Ferramenta de Diagnóstico - DocsGenOS
Verifica se o ambiente está configurado corretamente
"""

import os
import sys
import subprocess
from datetime import datetime

def criar_relatorio():
    """Cria um relatório de diagnóstico do sistema"""
    
    nome_arquivo = os.path.join(os.path.dirname(__file__), f"diagnostico_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    
    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write("RELATÓRIO DE DIAGNÓSTICO - DocsGenOS\n")
        f.write(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write("=" * 80 + "\n\n")
        
        # Informações do Python
        f.write("--- INFORMAÇÕES DO PYTHON ---\n")
        f.write(f"Versão: {sys.version}\n")
        f.write(f"Executável: {sys.executable}\n")
        f.write(f"Arquitetura: {sys.platform}\n")
        f.write(f"Codificação padrão: {sys.getdefaultencoding()}\n\n")
        
        # Diretórios importantes
        f.write("--- DIRETÓRIOS ---\n")
        base_dir = os.path.dirname(os.path.abspath(__file__))
        f.write(f"Diretório base: {base_dir}\n")
        f.write(f"Diretório DocsGen: {os.path.join(base_dir, 'DocsGen')}\n\n")
        
        # Verificar arquivos
        f.write("--- ARQUIVOS NECESSÁRIOS ---\n")
        arquivos_necessarios = [
            ("DocsGen/DocsGen.py", "Script principal"),
            ("DocsGen/DocsGen_OS.py", "Gerador de Ordens de Serviço"),
            ("DocsGen/DocsGen_SIT.py", "Gerador SIT"),
            ("DocsGen/DG_RemoveSign.py", "Removedor de Assinaturas"),
            ("DocsGen/listaDescFuncoes.txt", "Lista de Funções"),
            ("requirements.txt", "Dependências"),
        ]
        
        for arquivo, descricao in arquivos_necessarios:
            caminho_completo = os.path.join(base_dir, arquivo)
            existe = "✓ OK" if os.path.exists(caminho_completo) else "✗ FALTANDO"
            f.write(f"  {existe} - {descricao}: {caminho_completo}\n")
        f.write("\n")
        
        # Verificar bibliotecas
        f.write("--- BIBLIOTECAS PYTHON ---\n")
        bibliotecas = [
            "tkinter", "ttkbootstrap", "PIL", "docx", "comtypes", "plyer",
            "openpyxl", "fitz", "win32com"
        ]
        
        for lib in bibliotecas:
            try:
                __import__(lib)
                f.write(f"  ✓ OK - {lib}\n")
            except ImportError as e:
                f.write(f"  ✗ ERRO - {lib}: {str(e)}\n")
        f.write("\n")
        
        # Testar execução de scripts
        f.write("--- TESTE DE EXECUÇÃO ---\n")
        docsgen_py = os.path.join(base_dir, "DocsGen", "DocsGen.py")
        if os.path.exists(docsgen_py):
            try:
                resultado = subprocess.run(
                    [sys.executable, "-m", "py_compile", docsgen_py],
                    capture_output=True,
                    text=True,
                    timeout=10
                )
                if resultado.returncode == 0:
                    f.write(f"  ✓ OK - DocsGen.py pode ser compilado\n")
                else:
                    f.write(f"  ✗ ERRO - DocsGen.py não compila:\n{resultado.stderr}\n")
            except Exception as e:
                f.write(f"  ✗ ERRO ao testar DocsGen.py: {str(e)}\n")
        f.write("\n")
        
        # Informações de permissões
        f.write("--- PERMISSÕES ---\n")
        try:
            docsgen_dir = os.path.join(base_dir, "DocsGen")
            if os.path.exists(docsgen_dir):
                pode_ler = os.access(docsgen_dir, os.R_OK)
                pode_escrever = os.access(docsgen_dir, os.W_OK)
                pode_executar = os.access(docsgen_dir, os.X_OK)
                f.write(f"  Leitura DocsGen/: {'✓' if pode_ler else '✗'}\n")
                f.write(f"  Escrita DocsGen/: {'✓' if pode_escrever else '✗'}\n")
                f.write(f"  Execução DocsGen/: {'✓' if pode_executar else '✗'}\n")
        except Exception as e:
            f.write(f"  ✗ Erro ao verificar permissões: {str(e)}\n")
        f.write("\n")
        
        f.write("=" * 80 + "\n")
        f.write("FIM DO RELATÓRIO\n")
        f.write("=" * 80 + "\n")
    
    return nome_arquivo

if __name__ == "__main__":
    print("Gerando relatório de diagnóstico...")
    arquivo = criar_relatorio()
    print(f"\nRelatório salvo em: {arquivo}")
    print("\nAbrindo arquivo...")
    os.startfile(arquivo)
