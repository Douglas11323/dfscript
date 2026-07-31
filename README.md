# Diário Oficial de Vila Velha — Coleta Automática de Movimentações

Automação que monitora o **Diário Oficial do Município de Vila Velha (ES)**, extrai
nomeações, exonerações, vacâncias e transferências de cargos comissionados/efetivos
publicados diariamente, e consolida tudo em uma planilha Excel histórica.

Roda automaticamente de segunda a sexta às **8h (horário de Brasília)** via GitHub
Actions, sem precisar de servidor ou máquina ligada.

---

## O que o script faz

1. **Baixa** a última edição do Diário Oficial publicada no portal oficial
   (`diariooficial.vilavelha.es.gov.br`), usando Playwright para simular o clique
   no botão de "Última Edição".
2. **Extrai o texto** do PDF (com `pdfplumber`), detectando automaticamente se a
   página está em uma ou duas colunas para não embaralhar o conteúdo.
3. **Localiza atos** de Portaria/Decreto e, dentro deles, identifica por meio de
   expressões regulares:
   - `Exonerar` → **Exonerado**
   - `Nomear` → **Nomeado**
   - `Declarar vacância` → **Vacância**
   - `Transferir a lotação` → **Transferido**
4. **Atualiza a planilha mestre** (`saida/movimentacoes.xlsx`), somando os novos
   registros aos já existentes, removendo duplicatas e mantendo a ordem lógica
   (ex.: exoneração de um servidor sempre aparece antes de uma nova nomeação na
   mesma data).
5. **Commita a planilha** de volta no repositório automaticamente, criando um
   histórico versionado de todas as movimentações — dia após dia.

A planilha gerada tem 5 abas: `Movimentações` (visão geral) e uma aba separada
para cada situação (`Nomeações`, `Exonerações`, `Vacâncias`, `Transferências`).

---

## Estrutura do projeto

```
diarioof/
├── .github/
│   └── workflows/
│       └── coletar.yml          # agenda a execução diária (GitHub Actions)
├── saida/
│   └── movimentacoes.xlsx       # planilha histórica (versionada no git)
├── downloads/                   # PDFs baixados (não versionado)
├── diario_exonerados.py         # script principal
├── run_diario_exonerados.bat    # atalho para rodar localmente no Windows
├── requirements.txt             # dependências Python
├── .gitignore
└── README.md
```

---

## Rodando localmente

### Pré-requisitos
- Python 3.12+
- Windows, macOS ou Linux

### Passo a passo

```bash
# 1. Criar e ativar o ambiente virtual
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate

# 2. Instalar dependências
pip install -r requirements.txt

# 3. Instalar o navegador usado pelo Playwright
playwright install --with-deps chromium

# 4. Rodar a coleta
python diario_exonerados.py
```

No Windows, depois de configurado o ambiente virtual, basta dar duplo clique em
`run_diario_exonerados.bat` para rodar sem abrir o terminal manualmente.

A planilha atualizada fica em `saida/movimentacoes.xlsx`. Os logs de cada
execução ficam em `saida/logs/`.

---

## Automação (GitHub Actions)

O workflow em `.github/workflows/coletar.yml` roda automaticamente:

- **Quando:** segunda a sexta, às 11h UTC (8h em Brasília)
- **O que faz:** baixa a edição do dia, extrai as movimentações, atualiza
  `saida/movimentacoes.xlsx` e faz commit + push da planilha atualizada
- **Execução manual:** também é possível disparar a qualquer momento pela aba
  **Actions** do repositório, clicando em *"Run workflow"*

Não é necessário nenhum servidor ligado — o GitHub Actions cuida de tudo.

---

## Observações importantes

- O script depende da estrutura atual do portal do Diário Oficial de Vila Velha.
  Se o site mudar de layout, o seletor `#btn1` (botão "Última Edição") e a lógica
  de captura do PDF podem precisar de ajuste em `baixar_ultima_edicao()`.
- As expressões regulares foram desenhadas para o padrão de redação usado nas
  Portarias/Decretos do município. Mudanças na formatação dos atos publicados
  podem exigir ajustes nos padrões (`PADRAO_EXONERAR`, `PADRAO_NOMEAR`, etc.).
- A deduplicação usa como chave `Data + Servidor + Situação`, então o mesmo
  servidor pode aparecer mais de uma vez na planilha se tiver mais de uma
  movimentação em datas diferentes (o que é esperado).
