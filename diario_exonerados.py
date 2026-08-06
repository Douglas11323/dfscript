"""
Automação: Diário Oficial de Vila Velha -> Planilha de Movimentações (Execução Local)
==================================================================================
Baixa a última edição do Diário Oficial de Vila Velha, extrai o texto do PDF,
localiza nomeações, exonerações, vacâncias e transferências de cargos
comissionados/efetivos e consolida tudo em uma planilha Excel local.

Changelog desta revisão (correção do bug de extração de 04/08/2026):
- Reescrita completa da extração de texto de páginas em duas colunas. O
  heurístico antigo (`eh_duas_colunas`) media a fração de PALAVRAS próximas
  do centro da página, o que é impreciso: em qualquer página de 2 colunas,
  dezenas de palavras de AMBAS as colunas caem perto do centro só porque as
  colunas são estreitas — isso fazia páginas genuinamente de 2 colunas serem
  classificadas como 1 coluna, usando pdfplumber puro, que intercala linha a
  linha o texto da coluna esquerda com a da direita. O resultado eram
  parágrafos de Portarias diferentes grudados um no outro, quebrando todos os
  regex (nome truncado, cargo/secretaria "vazando" para dentro do texto de
  outra Portaria, etc.) — exatamente o que aconteceu com a Portaria 413/2026.
- A nova extração (`extrair_texto`) trabalha linha a linha: agrupa palavras em
  linhas visuais, detecta a "calha" (gutter) real entre as colunas medindo o
  menor espaço em branco compartilhado por várias linhas — não um limiar fixo
  de 50% da largura — e só then divide cada linha exatamente nesse ponto.
  Páginas de coluna única (ex.: a Resolução do IPVV, que ocupa a largura
  inteira) continuam sendo extraídas normalmente, sem qualquer corte, porque
  o algoritmo simplesmente não encontra uma calha estável e cai no
  `page.extract_text()` padrão. Isso deve funcionar em qualquer edição futura,
  não só na de hoje.
- Regex de "Exonerar" corrigido: o conector antes de "cargo" aceitava tanto
  "do" quanto "de". Como nomes brasileiros frequentemente contêm "de"
  ("Wanilda DE Andrade..."), o regex não-guloso escolhia a interpretação mais
  curta e cortava o nome no primeiro "de" que encontrasse — mesmo com texto
  perfeitamente limpo. Agora o conector exige literalmente "do (seu) cargo",
  eliminando essa ambiguidade independente da extração de texto.
- Aceita tanto "Exonerar" quanto o verbo já conjugado "Exonera" (variação já
  observada em edições anteriores, ver Portaria 405/2026).
- Captura de cargo/secretaria agora tem limite de tamanho (evita que, em um
  texto malformado sem pontuação, o regex "vaze" por centenas de caracteres
  para dentro do próximo artigo/portaria).
- [NOVO] "Padrão/Símbolo/CC" agora é OPCIONAL em Exonerar e Nomear: cargos
  EFETIVOS (ex.: Bibliotecário, Professor, etc.) não têm código de padrão/CC
  — o texto vai direto de "do cargo efetivo de X," para "da Secretaria Y."
  A regex antiga exigia sempre um token em maiúsculas ([A-Z0-9-]+) entre a
  vírgula do cargo e a secretaria, o que fazia esses casos não darem match
  nenhum (nem entravam na planilha, silenciosamente). Ver Portaria 418/2026
  (exoneração de Aline Larangeira Chahoud, cargo efetivo de Bibliotecário,
  Secretaria Municipal de Educação — sem padrão CC). Casos com CC (ex.:
  Portaria 422/2026, "padrão CC-1") continuam funcionando normalmente, pois
  o grupo opcional é tentado primeiro pelo motor de regex antes de ser
  pulado.
"""

import logging
import re
import statistics
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber
from playwright.sync_api import sync_playwright

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DE DIRETÓRIOS E URLS (ESTRUTURA 100% LOCAL)
# ---------------------------------------------------------------------------

URL_DIARIO = "https://diariooficial.vilavelha.es.gov.br"
SELETOR_ULTIMA_EDICAO = "#btn1"

DIRETORIO_BASE = Path(__file__).resolve().parent

PASTA_DOWNLOADS = DIRETORIO_BASE / "downloads"
PASTA_SAIDA = DIRETORIO_BASE / "saida"
PASTA_LOGS = PASTA_SAIDA / "logs"

PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
PASTA_LOGS.mkdir(parents=True, exist_ok=True)

HEADLESS = True

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------

_data_log = datetime.now().strftime("%Y-%m-%d")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(PASTA_LOGS / f"execucao_{_data_log}.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

# ---------------------------------------------------------------------------
# PATTERNS
# ---------------------------------------------------------------------------

PADRAO_EXONERAR = re.compile(
    r"\b(?:Art\.\s*\d+[º°]?|DECRETA:?|RESOLVE:?)\s+"
    r"Exonera(?:r)?\b\s*,?\s*(?:a\s+pedido\s*,?\s*)?"
    r"(?P<nome>[^,]{3,70}?)\s*,?\s*"
    r"(?:matr[íi]cula\s+n[ºo°]?\.?\s*[\d/]+\s*,?\s*)?"
    # Conector fixo: só "do (seu) cargo ..." — nunca "de", que é ambíguo com
    # nomes contendo "de" no meio (ex.: "Wanilda de Andrade Oliveira Pedro").
    r"\bdo\s+(?:seu\s+)?cargo\s+"
    r"(?:efetivo\s+de|comissionado\s+de|em\s+comiss[ãa]o\s+de|de)\s+"
    r"(?P<cargo>[^,]{2,90}?),\s*"
    # Padrão/Símbolo/CC é OPCIONAL: cargos efetivos não têm esse código, o
    # texto vai direto para "da/do/na Secretaria...". Quando existe (cargos
    # comissionados), continua sendo capturado normalmente, pois o motor de
    # regex tenta esse grupo antes de descartá-lo.
    r"(?:(?:padr[ãa]o|s[íi]mbolo|n[íi]vel)?\s*(?P<padrao_cc>[A-Z0-9-]+),\s*)?"
    r"(?:da|do|no|na)\s+(?P<secretaria>[^.]{2,150}?)\.",
    re.IGNORECASE | re.DOTALL,
)

PADRAO_NOMEAR = re.compile(
    r"\b(?:Art\.\s*\d+[º°]?|DECRETA:?|RESOLVE:?)\s+"
    r"Nomear\b\s+(?P<nome>[^,]{3,70}?)\s+"
    r"para\s+exercer\s+(?:o\s+)?cargo\s+"
    r"(?:comissionado\s+de|em\s+comiss[ãa]o\s+de|de)\s*"
    r"(?P<cargo>[^,]{2,90}?),\s*"
    # Mesma correção: Padrão/Símbolo/CC opcional (nomeação para cargo
    # efetivo — ex.: posse de concursado aprovado — também não tem CC).
    r"(?:(?:padr[ãa]o|s[íi]mbolo|n[íi]vel)?\s*(?P<padrao_cc>[A-Z0-9-]+),\s*)?"
    r"(?:da|do|no|na)\s+(?P<secretaria>[^.]{2,150}?)\.",
    re.IGNORECASE | re.DOTALL,
)

PADRAO_VACANCIA = re.compile(
    r"\b(?:Art\.\s*\d+[º°]?|DECRETA:?|RESOLVE:?)\s+"
    r"Declarar\b\s+vac[âa]ncia\s+do\s+cargo\s+efetivo\s+de\s+"
    r"(?P<cargo>[^,]{2,90}?),\s*"
    r"(?:da|do|no|na)\s+(?P<secretaria>[^,.]{2,150}?),\s*"
    r"ocupado\s+pel[oa]\s+[Ss]ervidor[a]?\s+"
    r"(?P<nome>[^,]{3,70}?),\s*"
    r"(?:matr[íi]cula\s+n[ºo°]?\.?\s*[\d/]+)?",
    re.IGNORECASE | re.DOTALL,
)

PADRAO_TRANSFERENCIA = re.compile(
    r"\b(?:Art\.\s*\d+[º°]?|DECRETA:?|RESOLVE:?)\s+"
    r"Transferir\b\s+a\s+lota[çc][aã]o\s+de\s+"
    r"(?P<nome>[^,]{3,70}?),\s*"
    r"ocupante\s+do\s+cargo\s+comissionado\s+de\s+"
    r"(?P<cargo>[^,]{2,90}?),\s*"
    r"(?:padr[ãa]o|s[íi]mbolo|n[íi]vel)?\s*(?P<padrao_cc>[A-Z0-9-]+),\s*"
    r"(?:da|do|no|na)\s+(?P<secretaria_origem>[^.]{2,150}?)\s+para\s+(?:a|o)\s+"
    r"(?P<secretaria_destino>[^.]{2,150}?)\.",
    re.IGNORECASE | re.DOTALL,
)

# Cabeçalho de Portaria/Decreto (ex.: "PORTARIA Nº 409/2026", "PORTARIA SEMAS Nº 076/2026",
# "DECRETO Nº 300/2026"). Os lookaheads negativos evitam confundir com referências feitas
# dentro do corpo do texto, como "Decreto nº 038/2017 que dispõe..." ou "Decreto nº 072, de...".
PADRAO_ATO = re.compile(
    r"\b(?P<tipo>PORTARIA|DECRETO)\s+(?:[A-ZÇÃÕÁÉÍÓÚ]+\s+)?N[ºO°]\.?\s*(?P<numero>\d{1,5}/\d{4})"
    r"(?!\s*,)(?!\s*que\b)",
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# ETAPA 1: BAIXAR O PDF DEDICADO
# ---------------------------------------------------------------------------

def baixar_ultima_edicao() -> Path:
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=HEADLESS)
        contexto = navegador.new_context(accept_downloads=True)
        pagina = contexto.new_page()

        logging.info("Acessando o portal do Diário Oficial...")
        pagina.goto(URL_DIARIO, wait_until="domcontentloaded")

        data_hoje = datetime.now().strftime("%Y-%m-%d")
        caminho_pdf = PASTA_DOWNLOADS / f"diario_{data_hoje}.pdf"

        buffer_pdf = []

        def processar_resposta(resposta):
            try:
                content_type = resposta.headers.get("content-type", "").lower()
                if "application/pdf" in content_type:
                    body = resposta.body()
                    if body.startswith(b"%PDF"):
                        buffer_pdf.append(body)
            except Exception as e:
                logging.debug(f"Response ignorada ({resposta.url}): {e}")

        contexto.on("response", processar_resposta)

        logging.info("Clicando em 'Última Edição'...")

        download_capturado = None
        nova_aba = None
        try:
            with contexto.expect_event("page", timeout=15000) as nova_aba_info:
                try:
                    with pagina.expect_download(timeout=8000) as download_info:
                        pagina.click(SELETOR_ULTIMA_EDICAO)
                    download_capturado = download_info.value
                except Exception:
                    pass
            nova_aba = nova_aba_info.value
        except Exception:
            logging.info("Nenhuma nova aba detectada; seguindo com a aba atual.")

        if nova_aba is not None and download_capturado is None:
            try:
                download_capturado = nova_aba.wait_for_event("download", timeout=8000)
            except Exception:
                logging.info("Nenhum evento de download na nova aba (timeout de 8s).")

        if download_capturado is not None:
            download_capturado.save_as(str(caminho_pdf))
            navegador.close()
            logging.info(f"PDF baixado via evento de download em: {caminho_pdf}")
            return caminho_pdf

        if nova_aba is not None:
            nova_aba.wait_for_timeout(6000)

        if buffer_pdf:
            caminho_pdf.write_bytes(buffer_pdf[0])
            navegador.close()
            logging.info(f"PDF baixado via response HTTP em: {caminho_pdf}")
            return caminho_pdf

        if nova_aba is not None:
            for elem in ["embed", "iframe", "object"]:
                loc = nova_aba.locator(elem)
                if loc.count() > 0:
                    url_src = loc.first.get_attribute("src") or loc.first.get_attribute("data")
                    if url_src:
                        resp = contexto.request.get(url_src)
                        if resp.ok and resp.body().startswith(b"%PDF"):
                            caminho_pdf.write_bytes(resp.body())
                            navegador.close()
                            logging.info(f"PDF baixado via elemento <{elem}> em: {caminho_pdf}")
                            return caminho_pdf

            debug_html = PASTA_LOGS / f"debug_pagina_{data_hoje}.html"
            debug_html.write_text(nova_aba.content(), encoding="utf-8")
            logging.error(f"URL da nova aba no momento da falha: {nova_aba.url}")
            logging.error(f"HTML da nova aba salvo para inspeção em: {debug_html}")

        navegador.close()
        raise RuntimeError("Não foi possível capturar o fluxo do PDF.")

# ---------------------------------------------------------------------------
# ETAPA 2: EXTRAÇÃO DE TEXTO (reescrita — reconstrução por linha/coluna)
# ---------------------------------------------------------------------------

def _agrupar_linhas(palavras, tolerancia=2.5):
    linhas = []
    atual = []
    topo_atual = None
    for w in sorted(palavras, key=lambda w: (w["top"], w["x0"])):
        if topo_atual is None or abs(w["top"] - topo_atual) <= tolerancia:
            atual.append(w)
            topo_atual = w["top"] if topo_atual is None else topo_atual
        else:
            linhas.append(atual)
            atual = [w]
            topo_atual = w["top"]
    if atual:
        linhas.append(atual)
    return linhas


def _gaps_da_linha(palavras_ordenadas):
    """Lista de (tamanho_do_gap, indice_de_corte, x1_antes, x0_depois)."""
    gaps = []
    for i, (a, b) in enumerate(zip(palavras_ordenadas, palavras_ordenadas[1:])):
        gaps.append((b["x0"] - a["x1"], i + 1, a["x1"], b["x0"]))
    return gaps


def extrair_texto_pagina(pagina) -> str:
    palavras = pagina.extract_words()
    if not palavras:
        return pagina.extract_text() or ""

    largura = pagina.width
    linhas_brutas = _agrupar_linhas(palavras)
    linhas = []
    for ln in linhas_brutas:
        ln_ordenada = sorted(ln, key=lambda w: w["x0"])
        linhas.append(dict(
            top=min(w["top"] for w in ln_ordenada),
            xmin=min(w["x0"] for w in ln_ordenada),
            xmax=max(w["x1"] for w in ln_ordenada),
            palavras=ln_ordenada,
            gaps=_gaps_da_linha(ln_ordenada),
        ))

    # Passo 1: candidatas óbvias a "linha dividida entre colunas"
    GAP_GRANDE = max(50, largura * 0.08)
    candidatas = []
    for l in linhas:
        if not l["gaps"]:
            continue
        g, _, xa, xb = max(l["gaps"], key=lambda t: t[0])
        meio_do_gap = (xa + xb) / 2
        if g > GAP_GRANDE and largura * 0.2 < meio_do_gap < largura * 0.8:
            candidatas.append((xa, xb))

    MIN_CANDIDATAS = 4
    if len(candidatas) < MIN_CANDIDATAS:
        # Página de coluna única: sem corte nenhum.
        return pagina.extract_text() or ""

    calha_esq = max(c[0] for c in candidatas)
    calha_dir = min(c[1] for c in candidatas)
    if calha_esq >= calha_dir:
        calha_esq = statistics.median(c[0] for c in candidatas)
        calha_dir = statistics.median(c[1] for c in candidatas)
    calha_meio = (calha_esq + calha_dir) / 2

    TOLERANCIA_BORDA = 2.0
    saida = []
    buffer_esq, buffer_dir = [], []
    coluna_iniciada = False

    def flush():
        saida.extend(buffer_esq)
        saida.extend(buffer_dir)
        buffer_esq.clear()
        buffer_dir.clear()

    for l in sorted(linhas, key=lambda l: l["top"]):
        indice_corte = None
        for g, idx, xa, xb in l["gaps"]:
            if xa <= calha_dir and xb >= calha_esq and g >= 6:
                indice_corte = idx
                break

        if indice_corte is not None:
            esq = l["palavras"][:indice_corte]
            dir_ = l["palavras"][indice_corte:]
            buffer_esq.append(" ".join(w["text"] for w in esq))
            buffer_dir.append(" ".join(w["text"] for w in dir_))
            coluna_iniciada = True
            continue

        texto_linha = " ".join(w["text"] for w in l["palavras"])
        if l["xmax"] <= calha_dir + TOLERANCIA_BORDA:
            buffer_esq.append(texto_linha)
            coluna_iniciada = True
        elif l["xmin"] >= calha_esq - TOLERANCIA_BORDA:
            buffer_dir.append(texto_linha)
            coluna_iniciada = True
        else:
            if not coluna_iniciada:
                saida.append(texto_linha)
            else:
                flush()
                saida.append(texto_linha)

    flush()
    return "\n".join(saida)


def extrair_texto(caminho_pdf: Path) -> str:
    texto_completo = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto_completo.append(extrair_texto_pagina(pagina))

    texto_final = "\n".join(texto_completo)

    if not texto_final.strip():
        raise RuntimeError("O PDF não retornou texto selecionável.")

    return texto_final

# ---------------------------------------------------------------------------
# ETAPA 3: PROCESSAMENTO E LIMPEZA DOS DADOS
# ---------------------------------------------------------------------------

def limpar_secretaria(texto: str) -> str:
    if not texto:
        return ""
    match = re.search(r"(Secretaria\s+Municipal\s+de\s+[^.─\n]+|Secretaria\s+[^.─\n]+)", texto, re.IGNORECASE)
    if match:
        sec = match.group(1).strip()
        sec = re.split(r"\.|Art\.|Portaria|Decreto|\,", sec, flags=re.IGNORECASE)[0].strip()
        return sec
    return texto.strip()


_PALAVRAS_CHAVE_SEPARAR = [
    "para exercer",
    "cargo comissionado",
    "cargo em comissão",
    "padrão",
    "Secretaria",
    "Nomear",
    "Exonerar",
]


def corrigir_espacos_faltantes(texto: str) -> str:
    for palavra in _PALAVRAS_CHAVE_SEPARAR:
        texto = re.sub(rf"(?<=\S)(?={re.escape(palavra)})", " ", texto)
    return texto


def localizar_atos(texto: str) -> list[tuple[int, str]]:
    atos = []
    for m in PADRAO_ATO.finditer(texto):
        tipo = m.group("tipo").capitalize()
        numero = m.group("numero")
        atos.append((m.start(), f"{tipo} nº {numero}"))
    return atos


def ato_vigente(posicao: int, atos: list[tuple[int, str]]) -> str:
    rotulo = ""
    for pos_ato, label in atos:
        if pos_ato <= posicao:
            rotulo = label
        else:
            break
    return rotulo


def extrair_movimentacoes(texto: str) -> list[dict]:
    texto_limpo = re.sub(r"\s+", " ", texto)
    texto_limpo = corrigir_espacos_faltantes(texto_limpo)

    atos = localizar_atos(texto_limpo)

    data_hoje = datetime.now().strftime("%d/%m/%Y")
    encontrados = []

    for m in PADRAO_EXONERAR.finditer(texto_limpo):
        dados = m.groupdict()
        encontrados.append((m.start(), {
            "Data": data_hoje,
            "Portaria Nº": ato_vigente(m.start(), atos),
            "Servidor": re.sub(r"\s+", " ", dados["nome"] or "").strip(),
            "Situação": "Exonerado",
            "Cargo": re.sub(r"\s+", " ", dados["cargo"] or "").strip(),
            "Padrão CC": re.sub(r"\s+", " ", dados.get("padrao_cc") or "").strip(),
            "Secretaria": limpar_secretaria(dados.get("secretaria", "")),
            "Secretaria Destino": "",
        }))

    for m in PADRAO_NOMEAR.finditer(texto_limpo):
        dados = m.groupdict()
        encontrados.append((m.start(), {
            "Data": data_hoje,
            "Portaria Nº": ato_vigente(m.start(), atos),
            "Servidor": re.sub(r"\s+", " ", dados["nome"] or "").strip(),
            "Situação": "Nomeado",
            "Cargo": re.sub(r"\s+", " ", dados["cargo"] or "").strip(),
            "Padrão CC": re.sub(r"\s+", " ", dados.get("padrao_cc") or "").strip(),
            "Secretaria": limpar_secretaria(dados.get("secretaria", "")),
            "Secretaria Destino": "",
        }))

    for m in PADRAO_VACANCIA.finditer(texto_limpo):
        dados = m.groupdict()
        encontrados.append((m.start(), {
            "Data": data_hoje,
            "Portaria Nº": ato_vigente(m.start(), atos),
            "Servidor": re.sub(r"\s+", " ", dados["nome"] or "").strip(),
            "Situação": "Vacância",
            "Cargo": re.sub(r"\s+", " ", dados["cargo"] or "").strip(),
            "Padrão CC": "",
            "Secretaria": limpar_secretaria(dados.get("secretaria", "")),
            "Secretaria Destino": "",
        }))

    for m in PADRAO_TRANSFERENCIA.finditer(texto_limpo):
        dados = m.groupdict()
        encontrados.append((m.start(), {
            "Data": data_hoje,
            "Portaria Nº": ato_vigente(m.start(), atos),
            "Servidor": re.sub(r"\s+", " ", dados["nome"] or "").strip(),
            "Situação": "Transferido",
            "Cargo": re.sub(r"\s+", " ", dados["cargo"] or "").strip(),
            "Padrão CC": re.sub(r"\s+", " ", dados["padrao_cc"] or "").strip(),
            "Secretaria": limpar_secretaria(dados.get("secretaria_origem", "")),
            "Secretaria Destino": limpar_secretaria(dados.get("secretaria_destino", "")),
        }))

    encontrados.sort(key=lambda item: item[0])
    return [movimentacao for _, movimentacao in encontrados]

# ---------------------------------------------------------------------------
# ETAPA 4: GERAÇÃO/ATUALIZAÇÃO DA PLANILHA EXCEL LOCAL (COM LÓGICA DE ORDEM)
# ---------------------------------------------------------------------------

CAMINHO_PLANILHA_MESTRE = PASTA_SAIDA / "movimentacoes.xlsx"

COLUNAS = [
    "Data",
    "Portaria Nº",
    "Servidor",
    "Situação",
    "Cargo",
    "Padrão CC",
    "Secretaria",
    "Secretaria Destino",
]

ORDEM_SITUACAO = {
    "Exonerado": 1,
    "Nomeado": 2,
    "Vacância": 3,
    "Transferido": 4
}


def carregar_planilha_existente(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        return pd.DataFrame(columns=COLUNAS)
    try:
        return pd.read_excel(caminho, sheet_name="Movimentações", dtype=str)
    except Exception:
        logging.warning("Não foi possível ler a planilha existente; criando do zero.")
        return pd.DataFrame(columns=COLUNAS)


def gerar_planilha(movimentacoes: list[dict], caminho_pdf: Path) -> Path:
    df_novo = pd.DataFrame(movimentacoes)
    if df_novo.empty:
        df_novo = pd.DataFrame(columns=COLUNAS)
    else:
        df_novo = df_novo[COLUNAS]

    df_existente = carregar_planilha_existente(CAMINHO_PLANILHA_MESTRE)

    df_total = pd.concat([df_existente, df_novo], ignore_index=True)

    df_total = df_total.drop_duplicates(
        subset=["Data", "Servidor", "Situação"],
        keep="last",
    )

    df_total["_Ordem_Situacao"] = df_total["Situação"].map(ORDEM_SITUACAO).fillna(99)

    df_total = df_total.sort_values(
        by=["Data", "Servidor", "_Ordem_Situacao"],
        ascending=[True, True, True]
    ).drop(columns=["_Ordem_Situacao"])

    try:
        writer = pd.ExcelWriter(CAMINHO_PLANILHA_MESTRE, engine="openpyxl")
        caminho_saida = CAMINHO_PLANILHA_MESTRE
    except PermissionError:
        hora_atual = datetime.now().strftime("%H%M%S")
        caminho_saida = PASTA_SAIDA / f"movimentacoes_{hora_atual}.xlsx"
        logging.warning(f"Arquivo mestre aberto no Excel! Salvando cópia em: {caminho_saida.name}")
        writer = pd.ExcelWriter(caminho_saida, engine="openpyxl")

    with writer:
        df_total.to_excel(writer, sheet_name="Movimentações", index=False)

        if not df_total.empty:
            df_total[df_total["Situação"] == "Nomeado"].to_excel(writer, sheet_name="Nomeações", index=False)
            df_total[df_total["Situação"] == "Exonerado"].to_excel(writer, sheet_name="Exonerações", index=False)
            df_total[df_total["Situação"] == "Vacância"].to_excel(writer, sheet_name="Vacâncias", index=False)
            df_total[df_total["Situação"] == "Transferido"].to_excel(writer, sheet_name="Transferências", index=False)

    novos_adicionados = len(df_total) - len(df_existente)
    logging.info(f"Planilha atualizada em: {caminho_saida} "
                 f"({novos_adicionados} nova(s) linha(s), {len(df_total)} no total)")
    return caminho_saida

# ---------------------------------------------------------------------------
# EXECUÇÃO PRINCIPAL
# ---------------------------------------------------------------------------

def main():
    logging.info("1/4 - Baixando última edição do Diário Oficial...")
    caminho_pdf = baixar_ultima_edicao()

    logging.info("2/4 - Extraindo texto do PDF...")
    texto = extrair_texto(caminho_pdf)

    logging.info("3/4 - Localizando exonerações, vacâncias e nomeações...")
    movimentacoes = extrair_movimentacoes(texto)

    nomeados_cnt = sum(1 for m in movimentacoes if m["Situação"] == "Nomeado")
    exonerados_cnt = sum(1 for m in movimentacoes if m["Situação"] == "Exonerado")
    vacancias_cnt = sum(1 for m in movimentacoes if m["Situação"] == "Vacância")
    transferencias_cnt = sum(1 for m in movimentacoes if m["Situação"] == "Transferido")

    logging.info(f"{nomeados_cnt} nomeação(ões) encontrada(s)")
    logging.info(f"{exonerados_cnt} exoneração(ões) encontrada(s)")
    logging.info(f"{vacancias_cnt} vacância(s) encontrada(s)")
    logging.info(f"{transferencias_cnt} transferência(s) de lotação encontrada(s)")

    logging.info("4/4 - Atualizando planilha mestre local...")
    gerar_planilha(movimentacoes, caminho_pdf)


if __name__ == "__main__":
    try:
        main()
    except Exception as erro:
        logging.exception(f"Falha na execução: {erro}")
        sys.exit(1)