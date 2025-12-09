# DocsGenOS - Gerador de Documentos de Segurança

Sistema automatizado para geração de Ordens de Serviço (OS) e Anuências (Autorizações) para a Vestas, focado em segurança do trabalho e conformidade com Normas Regulamentadoras (NRs).

## 📋 Funcionalidades

*   **Gerador de Ordens de Serviço (OS)**: Cria documentos de OS baseados no GHE (Grupo Homogêneo de Exposição) do funcionário, selecionando automaticamente os riscos e EPIs associados.
*   **Gerador de Anuências**: Cria cartas de anuência para NR-10, NR-10 SEP, NR-12, NR-33 e NR-35.
*   **Configuração Centralizada**: Cargos e GHEs são gerenciados em um único arquivo JSON, compartilhado entre os módulos.
*   **Gestão de HSE**: Lista de profissionais de segurança configurável via arquivo de texto.
*   **Templates Word**: Modelos `.docx` padronizados e editáveis.

## 🚀 Como Usar

### Pré-requisitos
*   Python 3.x instalado.
*   Dependências listadas em `requirements.txt`.

### Instalação
1.  Clone o repositório ou baixe os arquivos.
2.  Instale as dependências necessárias:
    ```bash
    pip install -r requirements.txt
    ```

### Executando os Módulos

#### 1. Gerar Anuências
Execute o script principal de anuências:
```bash
python DocsGen/DocsGen_Anuencias.py
```
1.  Preencha os dados do funcionário (Nome, CPF, Iniciais).
2.  Selecione o **HSE Responsável**.
3.  Selecione o **GHE** e o **Cargo** (a lista de cargos atualiza automaticamente).
4.  Marque as caixas das NRs desejadas (NR10, NR35, etc.).
5.  Escolha a pasta de destino e clique em "Gerar Anuências".

#### 2. Gerar Ordens de Serviço (OS)
Execute o script de OS:
```bash
python DocsGen/DocsGen_OS.py
```
1.  Siga o fluxo similar de preenchimento de dados e seleção de GHE/Cargo.
2.  Selecione os riscos específicos (Físicos, Químicos, Ergonômicos, Mecânicos).
3.  Gere o documento.

## ⚙️ Configuração

### Cargos e GHEs (`DocsGen/ghe_config.json`)
Este arquivo JSON controla a relação entre GHEs e Cargos para ambos os sistemas. Para adicionar um novo cargo, edite este arquivo e inclua o cargo na lista do GHE correspondente.

```json
{
  "01": [
    "ANALISTA DE CUSTO",
    "NOVO CARGO AQUI"
  ],
  "02": [...]
}
```

### Profissionais HSE (`DocsGen/listaHSE.txt`)
Arquivo de texto que alimenta a lista de responsáveis nos formulários. O formato deve ser respeitado:
`CARGO;NOME COMPLETO;REGISTRO`

Exemplo:
```text
HSE;MANOEL JEFETE DA SILVA TENONIO;MTE/RN: 1805
HSE;BRUNA PETRONI CEZARIO;CREA-RN: 2122993685
```
*Nota: Se o cargo for "HSE", o sistema o converterá automaticamente para "Técnico(a) de Segurança do Trabalho" nos documentos.*

### Templates (`DocsGen/templates/`)
Os modelos Word (`.docx`) utilizam "placeholders" para substituição automática. Não altere o nome dos arquivos, apenas o conteúdo se necessário.

**Variáveis disponíveis para uso no Word:**
*   `NOMEFUNCIONARIO`: Nome do colaborador.
*   `FUNCAOFUNCIONARIO`: Cargo do colaborador.
*   `CPFFUNCIONARIO`: CPF do colaborador.
*   `NOMEHSE`: Nome do responsável HSE.
*   `FUNCAOHSE`: Função do responsável HSE.
*   `REGISTROHSE`: Registro profissional do HSE.
*   `DIAANUENCIA`, `MESANUENCIA`, `ANOANUENCIA`: Data atual.

## 📂 Estrutura do Projeto

*   `DocsGen/`: Pasta principal contendo os scripts Python.
    *   `DocsGen_Anuencias.py`: Script de Anuências.
    *   `DocsGen_OS.py`: Script de Ordens de Serviço.
    *   `ghe_config.json`: Configuração de Cargos/GHE.
    *   `listaHSE.txt`: Lista de profissionais HSE.
    *   `templates/`: Modelos de documentos Word.
*   `RemoveSign/`: Utilitário para remoção de assinaturas.
*   `requirements.txt`: Lista de bibliotecas Python necessárias.

---
**DocsGenOS** - Simplificando a conformidade em SSMA.
