# Diário Oficial de Vila Velha — Coleta Automática de Movimentações

Automação que monitora o **Diário Oficial do Município de Vila Velha (ES)**, extrai nomeações, exonerações, vacâncias e transferências de cargos comissionados/efetivos publicados diariamente, e consolida tudo diretamente em uma planilha no **Google Sheets**.

---

## O que o script faz

1. **Baixa** a última edição do Diário Oficial publicada no portal oficial (`diariooficial.vilavelha.es.gov.br`), usando Playwright para simular o clique no botão de "Última Edição".
2. **Extrai o texto** do PDF (com `pdfplumber`), detectando automaticamente se a página está em uma ou duas colunas para não embaralhar o conteúdo.
3. **Localiza atos** de Portaria/Decreto e identifica por meio de expressões regulares:
   - `Exonerar` → **Exonerado**
   - `Nomear` → **Nomeado**
   - `Declarar vacância` → **Vacância**
   - `Transferir a lotação` → **Transferido**
4. **Atualiza a planilha mestre no Google Sheets**, conectando-se via API oficial do Google. Ele soma os novos registros aos já existentes na sua primeira aba, garantindo que tudo permaneça em ordem cronológica.

---

## Estrutura do projeto

```text
diarioof/
├── downloads/                   # PDFs baixados (pasta criada automaticamente)
├── logs/                        # Logs de execução (pasta criada automaticamente)
├── diario_exonerados.py         # Script principal
├── run_diario_exonerados.bat    # Atalho para rodar localmente no Windows
├── requirements.txt             # Dependências Python
├── credenciais-gcp.json         # Suas credenciais de API do Google Cloud (Oculto no Git)
├── token.json                   # Token de acesso salvo após o 1º login (Oculto no Git)
├── .gitignore                   # Configuração de arquivos ignorados pelo Git
└── README.md
```

---

## Rodando localmente

### Pré-requisitos
- Python 3.12+
- Windows

### Configuração Inicial (Uma única vez)

```bash
# 1. Criar e ativar o ambiente virtual
python -m venv .venv
.\.venv\Scripts\activate

# 2. Instalar dependências
pip install -r requirements.txt

# 3. Instalar o navegador usado pelo Playwright
playwright install chromium
```

### Execução no dia a dia

Basta dar um duplo clique no arquivo **`run_diario_exonerados.bat`**.

Na **primeira execução**, o script abrirá o seu navegador solicitando que você faça login na sua conta do Google e autorize o acesso à sua planilha. Depois que você aceitar, ele gerará um arquivo `token.json` e as próximas execuções ocorrerão de forma totalmente silenciosa e automática nos bastidores.

---

## Observações importantes

- O script depende da estrutura atual do portal do Diário Oficial de Vila Velha. Se o site mudar de layout, o seletor `#btn1` e a lógica de captura do PDF podem precisar de ajuste em `baixar_ultima_edicao()`.
- As expressões regulares foram desenhadas para o padrão de redação usado nas Portarias/Decretos do município. Mudanças na formatação dos atos publicados podem exigir ajustes nos padrões.
