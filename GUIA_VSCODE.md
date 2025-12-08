# Guia: Executar DocsGenOS no Visual Studio Code

## 1. Abrir o Projeto

1. Abra VS Code
2. Clique em **File** → **Open Folder**
3. Navegue até `DocsGenOS` e clique em **Selecionar Pasta**

## 2. Configurar o Python

1. Pressione `Ctrl + Shift + P` para abrir a Paleta de Comandos
2. Digite `Python: Select Interpreter`
3. Escolha a versão Python 3.14 (ou a versão que você tem instalada)
   - Procure por `Python 3.14` em `AppData\Local\Programs\Python\Python314`

## 3. Instalar Extensões Recomendadas

VS Code pedirá para instalar as extensões Python. Clique em **Install All** para instalar:
- Python
- Pylance
- Debugpy

## 4. Executar o Código

### Opção A: Usar Tasks (Mais Fácil)

Pressione `Ctrl + Shift + B` ou vá em **Terminal** → **Run Task**

Escolha uma das opções:
- **Configurar Ambiente** - Instala bibliotecas necessárias
- **Executar DocsGen** - Inicia o programa principal
- **Executar DocsGen OS** - Executa o gerador de Ordens de Serviço
- **Executar RemoveSign** - Executa o removedor de assinaturas
- **Executar Diagnóstico** - Executa o diagnóstico do sistema

### Opção B: Usar Debugger

1. Clique no ícone **Run and Debug** na barra lateral esquerda
2. Na lista de configurações, selecione uma opção (ex: "Python: DocsGen")
3. Clique no botão **Run** (ícone de play verde)

### Opção C: Executar Arquivo Atual

1. Abra o arquivo Python desejado
2. Pressione `F5` ou clique em **Run** → **Start Debugging**
3. Escolha **debugpy** como debugger

## 5. Ver Saída

A saída do programa aparecerá no painel **Terminal** (parte inferior do VS Code)

## 6. Configurar Python Path (se necessário)

Se o VS Code não encontrar as bibliotecas:

1. Pressione `Ctrl + ,` para abrir Settings
2. Procure por `Python: Default Interpreter Path`
3. Insira o caminho completo:
   ```
   C:\Users\SAPSD\AppData\Local\Programs\Python\Python314\python.exe
   ```

## 7. Atalhos Úteis

| Atalho | Ação |
|--------|------|
| `Ctrl + Shift + B` | Executar task padrão (Configurar Ambiente) |
| `F5` | Iniciar debugging |
| `Ctrl + K, Ctrl + 0` | Recolher todas as pastas |
| `Ctrl + /` | Comentar/descomentar linha |
| `Ctrl + Shift + P` | Paleta de Comandos |

## 8. Troubleshooting

### Python não reconhecido
```bash
Ctrl + Shift + P → Python: Select Interpreter → Escolha a correta
```

### Módulos não encontrados
```bash
Ctrl + Shift + B → Configurar Ambiente
```

### Terminal não abrir
```bash
Terminal → New Terminal
```

## 9. Debug Avançado

Para adicionar breakpoints:
1. Clique na área à esquerda do número da linha
2. Um ponto vermelho aparecerá
3. Pressione `F5` para iniciar o debug
4. Use `F10` (step over) ou `F11` (step into) para navegar

## 10. Integração com Git

No VS Code você pode:
- Ver mudanças nos arquivos (aba **Source Control**)
- Fazer commits
- Ver histórico

Basta clicar na aba **Source Control** na barra lateral esquerda.
