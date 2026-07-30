"""
Automação: Diário Oficial de Vila Velha -> Planilha de Movimentações (Execução Local)
==================================================================================
CORREÇÃO: a captura do PDF na Etapa 1 agora trata explicitamente o caso em que
o Chromium dispara um evento de DOWNLOAD nativo (Content-Disposition: attachment)
em vez de renderizar o PDF na página/embed. Esse era o motivo mais provável do
erro "Não foi possível capturar o fluxo do PDF." — o listener de response
engolia a exceção silenciosamente (try/except Exception: pass) e o PDF nunca
era mesmo interceptado por response nem encontrado como <embed>/<iframe>.
"""

import logging
import re
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

# Precisa ser True no GitHub Actions (não há interface gráfica no runner).
# Só use False localmente, se precisar depurar visualmente.
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
# PATTERNS (inalterados)
# ---------------------------------------------------------------------------

PADRAO_EXONERAR = re.compile(
    r"(?:Art\.\s*\d+º?|DECRETA:?|RESOLVE:?)\s*"
    r"\bExonerar?\b\s*,?\s*(?:a\s+pedido\s*,?\s*)?"
    r"(?P<nome>[^,]{3,70}?),\s*"
    r"(?:matr[íi]cula\s+n[ºo°]?\.?\s*[\d/]+\s*,\s*)?"
    r"(?:do|de)\s+(?:[^.]*?\s+)?cargo\s+(?:comissionado\s+de|em\s+comissão\s+de)\s+"
    r"(?P<cargo>.+?),\s*"
    r"(?:padrão|símbolo|nível)?\s*(?P<padrao_cc>[A-Z0-9-]+),\s*"
    r"(?:da|do|no|na)\s+(?P<secretaria>[^.]+?)\.",
    re.IGNORECASE | re.DOTALL,
)

PADRAO_NOMEAR = re.compile(
    r"(?:Art\.\s*\d+º?|DECRETA:?|RESOLVE:?)\s*"
    r"\bNomear\b\s+(?P<nome>[^,]{3,70}?)\s+"
    r"para\s+exercer\s+(?:[^.]*?\s+)?cargo\s+(?:comissionado\s+de|em\s+comissão\s+de)?\s*"
    r"(?P<cargo>.+?),\s*"
    r"(?:padrão|símbolo|nível)?\s*(?P<padrao_cc>[A-Z0-9-]+),\s*"
    r"(?:da|do|no|na)\s+(?P<secretaria>[^.]+?)\.",
    re.IGNORECASE | re.DOTALL,
)

PADRAO_VACANCIA = re.compile(
    r"(?:Art\.\s*\d+º?|DECRETA:?|RESOLVE:?)\s*"
    r"\bDeclarar\b\s+vac[âa]ncia\s+do\s+cargo\s+efetivo\s+de\s+"
    r"(?P<cargo>[^,]+?),\s*"
    r"(?:da|do|no|na)\s+(?P<secretaria>[^,]+?),\s*"
    r"ocupado\s+pel[oa]\s+[Ss]ervidor[a]?\s+"
    r"(?P<nome>[^,]{3,70}?),\s*"
    r"(?:matr[íi]cula\s+n[ºo°]?\.?\s*[\d/]+\s*,?\s*)?",
    re.IGNORECASE | re.DOTALL,
)

PADRAO_TRANSFERENCIA = re.compile(
    r"(?:Art\.\s*\d+º?|DECRETA:?|RESOLVE:?)\s*"
    r"\bTransferir\b\s+a\s+lota[çc][aã]o\s+de\s+"
    r"(?P<nome>[^,]{3,70}?),\s*"
    r"ocupante\s+do\s+cargo\s+comissionado\s+de\s+"
    r"(?P<cargo>[^,]+?),\s*"
    r"(?:padrão|símbolo|nível)?\s*(?P<padrao_cc>[A-Z0-9-]+),\s*"
    r"(?:da|do|no|na)\s+(?P<secretaria_origem>[^.]+?)\s+para\s+(?:a|o)\s+"
    r"(?P<secretaria_destino>[^.]+?)\.",
    re.IGNORECASE | re.DOTALL,
)

# ---------------------------------------------------------------------------
# ETAPA 1: BAIXAR O PDF DEDICADO (CORRIGIDA)
# ---------------------------------------------------------------------------

def baixar_ultima_edicao() -> Path:
    """Abre o portal, clica em 'Última Edição' e intercepta o PDF.

    Estratégia em 3 camadas, na ordem de prioridade:
      1. Evento de DOWNLOAD nativo do Chromium (o mais comum quando o servidor
         manda Content-Disposition: attachment).
      2. Response HTTP com content-type application/pdf (site que renderiza
         inline).
      3. Elemento <embed>/<iframe>/<object> apontando para o arquivo.
    Cada camada agora loga o motivo se falhar, em vez de engolir a exceção.
    """
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
                # Antes: "except Exception: pass" — escondia o motivo real.
                logging.debug(f"Response ignorada ({resposta.url}): {e}")

        contexto.on("response", processar_resposta)

        logging.info("Clicando em 'Última Edição'...")

        # Camada 1: tenta capturar um DOWNLOAD nativo e, em paralelo, uma nova aba.
        download_capturado = None
        nova_aba = None
        try:
            with contexto.expect_event("page", timeout=15000) as nova_aba_info:
                try:
                    with pagina.expect_download(timeout=8000) as download_info:
                        pagina.click(SELETOR_ULTIMA_EDICAO)
                    download_capturado = download_info.value
                except Exception:
                    # Não veio download na página atual; talvez tenha vindo na aba nova.
                    pass
            nova_aba = nova_aba_info.value
        except Exception:
            logging.info("Nenhuma nova aba detectada; seguindo com a aba atual.")

        # Se a nova aba abriu, também pode ter disparado um download nela.
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

        # Camada 2: já deve ter sido preenchido pelo listener de response.
        if nova_aba is not None:
            nova_aba.wait_for_timeout(6000)

        if buffer_pdf:
            caminho_pdf.write_bytes(buffer_pdf[0])
            navegador.close()
            logging.info(f"PDF baixado via response HTTP em: {caminho_pdf}")
            return caminho_pdf

        # Camada 3: procura embed/iframe/object na nova aba.
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

            # DEBUG: se chegou até aqui, salva um snapshot da aba para inspeção manual.
            debug_html = PASTA_LOGS / f"debug_pagina_{data_hoje}.html"
            debug_html.write_text(nova_aba.content(), encoding="utf-8")
            logging.error(f"URL da nova aba no momento da falha: {nova_aba.url}")
            logging.error(f"HTML da nova aba salvo para inspeção em: {debug_html}")

        navegador.close()
        raise RuntimeError("Não foi possível capturar o fluxo do PDF.")

# ---------------------------------------------------------------------------
# ETAPA 2: EXTRAÇÃO DE TEXTO
# ---------------------------------------------------------------------------

def extrair_texto(caminho_pdf: Path) -> str:
    """Extrai o texto do PDF respeitando o layout em DUAS COLUNAS do Diário Oficial."""
    texto_completo = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            largura = pagina.width
            meio = largura / 2

            coluna_esquerda = pagina.crop((0, 0, meio, pagina.height))
            coluna_direita = pagina.crop((meio, 0, largura, pagina.height))

            texto_esquerda = coluna_esquerda.extract_text() or ""
            texto_direita = coluna_direita.extract_text() or ""

            texto_completo.append(texto_esquerda)
            texto_completo.append(texto_direita)

    texto_final = "\n".join(texto_completo)

    if not texto_final.strip():
        raise RuntimeError("O PDF não retornou texto selecionável.")

    return texto_final

# ---------------------------------------------------------------------------
# ETAPA 3: PROCESSAMENTO E LIMPEZA DOS DADOS
# ---------------------------------------------------------------------------

def limpar_secretaria(texto: str) -> str:
    """Extrai apenas o nome oficial da Secretaria, descartando ementas/decretos vizinhos."""
    if not texto:
        return ""
    match = re.search(r"(Secretaria\s+Municipal\s+de\s+[^.─\n]+|Secretaria\s+[^.─\n]+)", texto, re.IGNORECASE)
    if match:
        sec = match.group(1).strip()
        sec = re.split(r"\.|Art\.|Portaria|Decreto", sec, flags=re.IGNORECASE)[0].strip()
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
    """Insere espaço antes de palavras-chave quando coladas ao texto anterior."""
    for palavra in _PALAVRAS_CHAVE_SEPARAR:
        texto = re.sub(rf"(?<=\S)(?={re.escape(palavra)})", " ", texto)
    return texto


def extrair_movimentacoes(texto: str) -> list[dict]:
    texto_limpo = re.sub(r"\s+", " ", texto)
    texto_limpo = corrigir_espacos_faltantes(texto_limpo)

    data_hoje = datetime.now().strftime("%d/%m/%Y")
    lista_movimentacoes = []

    for m in PADRAO_NOMEAR.finditer(texto_limpo):
        dados = m.groupdict()
        sec_limpa = limpar_secretaria(dados.get("secretaria", ""))
        lista_movimentacoes.append({
            "Data": data_hoje,
            "Servidor": re.sub(r"\s+", " ", dados["nome"] or "").strip(),
            "Situação": "Nomeado",
            "Cargo": re.sub(r"\s+", " ", dados["cargo"] or "").strip(),
            "Padrão CC": re.sub(r"\s+", " ", dados["padrao_cc"] or "").strip(),
            "Secretaria": sec_limpa,
            "Secretaria Destino": "",
        })

    for m in PADRAO_EXONERAR.finditer(texto_limpo):
        dados = m.groupdict()
        sec_limpa = limpar_secretaria(dados.get("secretaria", ""))
        lista_movimentacoes.append({
            "Data": data_hoje,
            "Servidor": re.sub(r"\s+", " ", dados["nome"] or "").strip(),
            "Situação": "Exonerado",
            "Cargo": re.sub(r"\s+", " ", dados["cargo"] or "").strip(),
            "Padrão CC": re.sub(r"\s+", " ", dados["padrao_cc"] or "").strip(),
            "Secretaria": sec_limpa,
            "Secretaria Destino": "",
        })

    for m in PADRAO_VACANCIA.finditer(texto_limpo):
        dados = m.groupdict()
        sec_limpa = limpar_secretaria(dados.get("secretaria", ""))
        lista_movimentacoes.append({
            "Data": data_hoje,
            "Servidor": re.sub(r"\s+", " ", dados["nome"] or "").strip(),
            "Situação": "Vacância",
            "Cargo": re.sub(r"\s+", " ", dados["cargo"] or "").strip(),
            "Padrão CC": "",
            "Secretaria": sec_limpa,
            "Secretaria Destino": "",
        })

    for m in PADRAO_TRANSFERENCIA.finditer(texto_limpo):
        dados = m.groupdict()
        sec_origem_limpa = limpar_secretaria(dados.get("secretaria_origem", ""))
        sec_destino_limpa = limpar_secretaria(dados.get("secretaria_destino", ""))
        lista_movimentacoes.append({
            "Data": data_hoje,
            "Servidor": re.sub(r"\s+", " ", dados["nome"] or "").strip(),
            "Situação": "Transferido",
            "Cargo": re.sub(r"\s+", " ", dados["cargo"] or "").strip(),
            "Padrão CC": re.sub(r"\s+", " ", dados["padrao_cc"] or "").strip(),
            "Secretaria": sec_origem_limpa,
            "Secretaria Destino": sec_destino_limpa,
        })

    return lista_movimentacoes

# ---------------------------------------------------------------------------
# ETAPA 4: GERAÇÃO/ATUALIZAÇÃO DA PLANILHA EXCEL LOCAL
# ---------------------------------------------------------------------------

CAMINHO_PLANILHA_MESTRE = PASTA_SAIDA / "movimentacoes.xlsx"

COLUNAS = [
    "Data",
    "Servidor",
    "Situação",
    "Cargo",
    "Padrão CC",
    "Secretaria",
    "Secretaria Destino",
]


def carregar_planilha_existente(caminho: Path) -> pd.DataFrame:
    """Lê a aba 'Movimentações' da planilha mestre local, se ela já existir."""
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
        subset=["Data", "Servidor", "Situação", "Cargo", "Padrão CC", "Secretaria"],
        keep="first",
    )

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