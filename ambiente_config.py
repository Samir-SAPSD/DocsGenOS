import subprocess
import sys
import os

def obter_python_executavel():
    """Obtém o caminho correto do executável Python independente da versão."""
    # sys.executable retorna o caminho completo do Python que está executando o script
    python_exe = sys.executable
    python_dir = os.path.dirname(python_exe)
    pythonw_exe = os.path.join(python_dir, "pythonw.exe")
    
    # Informações da versão
    versao = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    
    print(f"\n--- Informações do Python ---")
    print(f"Versão: {versao}")
    print(f"Executável (python.exe): {python_exe}")
    print(f"Executável (pythonw.exe): {pythonw_exe if os.path.exists(pythonw_exe) else 'Não encontrado'}")
    print(f"Diretório: {python_dir}")
    
    return {
        'python': python_exe,
        'pythonw': pythonw_exe if os.path.exists(pythonw_exe) else python_exe,
        'dir': python_dir,
        'versao': versao
    }

def instalar_bibliotecas(requirements_file):
    try:
        with open(requirements_file, 'r') as file:
            bibliotecas = file.readlines()
            bibliotecas = [biblioteca.strip() for biblioteca in bibliotecas if biblioteca.strip()]

            for biblioteca in bibliotecas:
                print(f"Instalando/atualizando '{biblioteca}'...")
                result = subprocess.run(
                    [sys.executable, '-m', 'pip', 'install', biblioteca],
                    capture_output=True,
                    text=True
                )
                if result.returncode == 0:
                    print(f"A biblioteca '{biblioteca}' foi instalada/atualizada com sucesso.")
                else:
                    print(f"Erro ao instalar '{biblioteca}': {result.stderr}")
                    
    except FileNotFoundError:
        print(f"Erro: O arquivo '{requirements_file}' não foi encontrado.")

def criar_atalho(nome_atalho, caminho_script, pasta_destino, python_info, icone=None, usar_console=False):
    """Cria um atalho .lnk para um script Python no Windows."""
    try:
        # Caminho completo do atalho
        atalho_path = os.path.join(pasta_destino, f"{nome_atalho}.lnk")
        
        # Usar o executável correto baseado na versão detectada
        if usar_console:
            python_path = python_info['python']  # python.exe (com console)
        else:
            python_path = python_info['pythonw']  # pythonw.exe (sem console)
        
        # PowerShell/WScript.Shell evita depender do pywin32, que falha ao ser importado logo após a instalação
        ps_script = (
            "$s = (New-Object -ComObject WScript.Shell).CreateShortcut($env:DG_LNK);"
            "$s.TargetPath = $env:DG_TARGET;"
            "$s.Arguments = '\"' + $env:DG_ARGS + '\"';"
            "$s.WorkingDirectory = $env:DG_WORKDIR;"
            "$s.Description = $env:DG_DESC;"
            "if ($env:DG_ICON) { $s.IconLocation = $env:DG_ICON };"
            "$s.Save()"
        )
        env = os.environ.copy()
        env.update({
            'DG_LNK': atalho_path,
            'DG_TARGET': python_path,
            'DG_ARGS': caminho_script,
            'DG_WORKDIR': os.path.dirname(caminho_script),
            'DG_DESC': f"{nome_atalho} (Python {python_info['versao']})",
            'DG_ICON': icone if icone and os.path.exists(icone) else '',
        })
        result = subprocess.run(
            ['powershell', '-NoProfile', '-NonInteractive', '-Command', ps_script],
            capture_output=True, text=True, env=env
        )
        if result.returncode != 0:
            print(f"Erro ao criar atalho '{nome_atalho}': {result.stderr.strip()}")
            return False
        print(f"Atalho '{nome_atalho}' criado em: {atalho_path}")
        print(f"  -> Usando: {python_path}")
        return True
        
    except Exception as e:
        print(f"Erro ao criar atalho '{nome_atalho}': {e}")
        return False

def criar_atalhos(python_info):
    """Cria os atalhos para os scripts principais."""
    # Diretório base do projeto
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Pasta onde os atalhos serão salvos
    pasta_atalhos = os.path.join(base_dir, "shortCuts")
    os.makedirs(pasta_atalhos, exist_ok=True)
    
    # Lista de atalhos a serem criados: (nome, caminho_relativo, usar_console)
    atalhos = [
        ("DocsGen - Gerador de Documentos", os.path.join(base_dir, "DocsGen", "DocsGen.py"), False),
        ("RemoveSign", os.path.join(base_dir, "RemoveSign", "RemoveSign.py"), False),
        ("Configurar Ambiente", os.path.join(base_dir, "ambiente_config.py"), True),
    ]
    
    print("\n--- Criando Atalhos ---")
    for nome, caminho, usar_console in atalhos:
        if os.path.exists(caminho):
            criar_atalho(nome, caminho, pasta_atalhos, python_info, usar_console=usar_console)
        else:
            print(f"Aviso: Arquivo não encontrado: {caminho}")

# Detectar versão e caminhos do Python
python_info = obter_python_executavel()

# Defina o arquivo de requisitos
base_dir = os.path.dirname(os.path.abspath(__file__))
requirements_file = os.path.join(base_dir, 'requirements.txt')

# Chame a função para instalar as bibliotecas
print("\n--- Instalando Bibliotecas ---")
instalar_bibliotecas(requirements_file)

# Criar os atalhos com os caminhos corretos do Python
criar_atalhos(python_info)
