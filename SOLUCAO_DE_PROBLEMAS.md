# DocsGenOS - Guia de Solução de Problemas

## Problema: Script não executa completamente

Se o programa inicia mas não termina de carregar, siga os passos abaixo:

### 1. Executar o Diagnóstico

Abra um terminal e execute:

```bash
python diagnostico.py
```

Isso vai gerar um arquivo `diagnostico_YYYYMMDD_HHMMSS.txt` com informações sobre:
- Versão do Python
- Bibliotecas instaladas
- Arquivos necessários
- Permissões de acesso

### 2. Verificar os Requisitos

Certifique-se de que todas as bibliotecas estão instaladas:

```bash
python ambiente_config.py
```

### 3. Causas Comuns

#### A. Bibliotecas Faltando
- **Solução**: Execute `python ambiente_config.py` para instalar todas as dependências

#### B. Permissões de Arquivo
- **Solução**: Verifique se você tem permissão de leitura/escrita na pasta DocsGenOS
- Execute como Administrador se necessário

#### C. Caminho de Arquivo Quebrado
- **Solução**: Verifique se os caminhos no código estão corretos
- Todos os caminhos devem ser relativos ao diretório do script

#### D. Antivírus/Firewall Bloqueando
- **Solução**: Adicione a pasta DocsGenOS à lista de exceções do antivírus

#### E. Falta de Espaço em Disco
- **Solução**: Libere espaço em disco (as imagens dos PDFs podem ocupar bastante espaço temporário)

#### F. Microsoft Word Não Instalado
- **Solução**: Alguns scripts usam COM para converter para PDF e precisam do Word instalado

### 4. Visualizar Logs de Erro

Quando um script falha, uma janela de erro aparece com a mensagem. Se quiser mais detalhes:

1. Execute o script diretamente em um terminal:
```bash
python DocsGen/DocsGen_OS.py
```

2. Anote a mensagem de erro completa

### 5. Contatar Suporte

Se nenhuma solução funcionar, forneça:
- O arquivo `diagnostico_*.txt`
- A mensagem de erro completa
- A sua versão do Windows
- Se você está usando como Administrador
