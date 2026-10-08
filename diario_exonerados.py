import logging
import re
import statistics
import sys
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber
from playwright.sync_api import sync_playwright

import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DE DIRETÓRIOS E PLANILHA
# ---------------------------------------------------------------------------

URL_DIARIO = "https://diariooficial.vilavelha.es.gov.br"
SELETOR_ULTIMA_EDICAO = "#btn1"

DIRETORIO_BASE = Path(__file__).resolve().parent
PASTA_DOWNLOADS = DIRETORIO_BASE / "downloads"
PASTA_LOGS = DIRETORIO_BASE / "logs"

PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
PASTA_LOGS.mkdir(parents=True, exist_ok=True)

HEADLESS = True

ID_PLANILHA = "10luQWC_abnlC09R1btbR1Bn9S8n11e-fJeir9ov2UaM"

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
# AUTENTICAÇÃO GOOGLE SHEETS (OAUTH 2.0)
# ---------------------------------------------------------------------------

ESCOPOS = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def autenticar_google_sheets():
    caminho_credencial = DIRETORIO_BASE / "credenciais-gcp.json"
    caminho_token = DIRETORIO_BASE / "token.json"
    creds = None
    
    if caminho_token.exists():
        creds = Credentials.from_authorized_user_file(str(caminho_token), ESCOPOS)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(caminho_credencial), ESCOPOS)
            creds = flow.run_local_server(port=0)
            
        with open(caminho_token, 'w') as token:
            token.write(creds.to_json())
            
    return gspread.authorize(creds)

# ---------------------------------------------------------------------------
# PATTERNS
# ---------------------------------------------------------------------------

_SECRETARIA_TAIL = (
    r"(?P<secretaria>[^,.]{2,150}?)"
    r"(?:,\s*(?:com\s+efeitos\b|de\s+acordo\s+com\b|conforme\b)[^.]*)?\."
)

PADRAO_EXONERAR = re.compile(
    r"\b(?:Art\.\s*\d+[º°]?|DECRETA:?|RESOLVE:?)\s+"
    r"Exonera(?:r)?\b\s*,?\s*(?:a\s+pedido\s*,?\s*)?"
    r"(?P<nome>[^,]{3,70}?)\s*,?\s*"
    r"(?:matr[íi]cula\s+n[ºo°]?\.?\s*[\d/]+\s*,?\s*)?"
    r"\b(?:do\s+(?:seu\s+)?|para\s+exercer\s+(?:o\s+)?)cargo\s+"
    r"(?:efetivo\s+de|comissionado\s+de|em\s+comiss[ãa]o\s+de|de)\s+"
    r"(?P<cargo>.+?),\s*"
    r"(?:(?:padr[ãa]o|s[íi]mbolo|n[íi]vel)?\s*(?P<padrao_cc>[A-Z0-9-]+),\s*)?"
    r"(?:da|do|no|na)\s+" + _SECRETARIA_TAIL,
    re.IGNORECASE | re.DOTALL,
)

PADRAO_NOMEAR = re.compile(
    r"\b(?:Art\.\s*\d+[º°]?|DECRETA:?|RESOLVE:?)\s+"
    r"Nomear\b\s+(?P<nome>[^,]{3,70}?)\s+"
    r"para\s+exercer\s+(?:o\s+)?cargo\s+"
    r"(?:comissionado\s+do\s+cargo\s+)?"
    r"(?:comissionado\s+de|em\s+comiss[ãa]o\s+de|de)\s*"
    r"(?P<cargo>.+?),\s*"
    r"(?:(?:padr[ãa]o|s[íi]mbolo|n[íi]vel)?\s*(?P<padrao_cc>[A-Z0-9-]+),\s*)?"
    r"(?:da|do|no|na)\s+" + _SECRETARIA_TAIL,
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

PADRAO_TORNAR_SEM_EFEITO = re.compile(
    r"\b(?:Art\.\s*\d+[º°]?|DECRETA:?|RESOLVE:?)\s+"
    r"Tornar\s+sem\s+efeito\s+o\s+art\.?\s*\d+[º°]?\s+da\s+Portaria\s+"
    r"n[ºo°]?\.?\s*(?P<portaria_original>\d{1,5}/\d{4})\s+que\s+"
    r"(?P<verbo>exonerou|nomeou)\s+"
    r"(?P<nome>[^,]{3,70}?)\s*,?\s*"
    r"(?:para\s+exercer\s+(?:o\s+)?)?"
    r"(?:do\s+)?cargo\s+"
    r"(?:comissionado\s+de|em\s+comiss[ãa]o\s+de|efetivo\s+de|de)\s*"
    r"(?P<cargo>.+?),\s*"
    r"(?:(?:padr[ãa]o|s[íi]mbolo|n[íi]vel)?\s*(?P<padrao_cc>[A-Z0-9-]+),\s*)?"
    r"(?:da|do|no|na)\s+" + _SECRETARIA_TAIL,
    re.IGNORECASE | re.DOTALL,
)

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
            except Exception:
                pass

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
            logging.info(f"PDF baixado em: {caminho_pdf}")
            return caminho_pdf

        if nova_aba is not None:
            nova_aba.wait_for_timeout(6000)

        if buffer_pdf:
            caminho_pdf.write_bytes(buffer_pdf[0])
            navegador.close()
            logging.info(f"PDF baixado em: {caminho_pdf}")
            return caminho_pdf

        navegador.close()
        raise RuntimeError("Não foi possível capturar o fluxo do PDF.")

# ---------------------------------------------------------------------------
# ETAPA 2: EXTRAÇÃO DE TEXTO
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
        return pagina.extract_text() or ""

    calha_esq = max(c[0] for c in candidatas)
    calha_dir = min(c[1] for c in candidatas)
    if calha_esq >= calha_dir:
        calha_esq = statistics.median(c[0] for c in candidatas)
        calha_dir = statistics.median(c[1] for c in candidatas)

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
    "Tornar sem efeito",
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
            "Ato Original Anulado": "",
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
            "Ato Original Anulado": "",
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
            "Ato Original Anulado": "",
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
            "Ato Original Anulado": "",
        }))

    # [NOVO] Tornar sem efeito
    for m in PADRAO_TORNAR_SEM_EFEITO.finditer(texto_limpo):
        dados = m.groupdict()
        verbo = (dados.get("verbo") or "").lower()
        situacao = "Sem Efeito - Exoneração" if verbo == "exonerou" else "Sem Efeito - Nomeação"
        portaria_original = dados.get("portaria_original") or ""
        encontrados.append((m.start(), {
            "Data": data_hoje,
            "Portaria Nº": ato_vigente(m.start(), atos),
            "Servidor": re.sub(r"\s+", " ", dados["nome"] or "").strip(),
            "Situação": situacao,
            "Cargo": re.sub(r"\s+", " ", dados["cargo"] or "").strip(),
            "Padrão CC": re.sub(r"\s+", " ", dados.get("padrao_cc") or "").strip(),
            "Secretaria": limpar_secretaria(dados.get("secretaria", "")),
            "Secretaria Destino": "",
            "Ato Original Anulado": f"Portaria nº {portaria_original}" if portaria_original else "",
        }))

    encontrados.sort(key=lambda item: item[0])
    return [movimentacao for _, movimentacao in encontrados]

# ---------------------------------------------------------------------------
# ETAPA 4: GERAÇÃO/ATUALIZAÇÃO DA PLANILHA NO GOOGLE SHEETS
# ---------------------------------------------------------------------------

COLUNAS = [
    "Data",
    "Portaria Nº",
    "Servidor",
    "Situação",
    "Cargo",
    "Padrão CC",
    "Secretaria",
    "Secretaria Destino",
    "Ato Original Anulado",
]

ORDEM_SITUACAO = {
    "Exonerado": 1,
    "Nomeado": 2,
    "Vacância": 3,
    "Transferido": 4,
    "Sem Efeito - Exoneração": 5,
    "Sem Efeito - Nomeação": 6,
}

def atualizar_aba(planilha, nome_aba, df_subset):
    try:
        aba = planilha.worksheet(nome_aba)
    except gspread.exceptions.WorksheetNotFound:
        aba = planilha.add_worksheet(title=nome_aba, rows=100, cols=len(COLUNAS))
        aba.append_row(COLUNAS)
    
    if df_subset.empty:
        return
        
    registros_existentes = aba.get_all_records()
    df_existente = pd.DataFrame(registros_existentes)
    
    if df_existente.empty:
        df_existente = pd.DataFrame(columns=COLUNAS)
    else:
        # Garantir que as colunas existam
        for col in COLUNAS:
            if col not in df_existente.columns:
                df_existente[col] = ""
        df_existente = df_existente[COLUNAS]
        
    df_total = pd.concat([df_existente, df_subset], ignore_index=True)
    
    # Preencher NaN com string vazia
    df_total = df_total.fillna("")
    
    # Remover duplicadas
    df_total = df_total.drop_duplicates(
        subset=["Data", "Servidor", "Situação"],
        keep="last",
    )
    
    # Ordenar
    df_total["_Ordem_Situacao"] = df_total["Situação"].map(ORDEM_SITUACAO).fillna(99)
    df_total = df_total.sort_values(
        by=["Data", "Servidor", "_Ordem_Situacao"],
        ascending=[True, True, True]
    ).drop(columns=["_Ordem_Situacao"])
    
    # Limpar a aba inteira e reescrever (mais seguro para garantir ordem correta)
    aba.clear()
    
    # Prepara os dados para o gspread (lista de listas)
    dados_insercao = [df_total.columns.values.tolist()] + df_total.values.tolist()
    aba.update(range_name='A1', values=dados_insercao)


def gerar_planilha(movimentacoes: list[dict], caminho_pdf: Path):
    df_novo = pd.DataFrame(movimentacoes)
    if df_novo.empty:
        df_novo = pd.DataFrame(columns=COLUNAS)
    else:
        df_novo = df_novo[COLUNAS]

    # Preencher NaN com string vazia
    df_novo = df_novo.fillna("")

    logging.info("Conectando ao Google Sheets...")
    gc = autenticar_google_sheets()
    planilha = gc.open_by_key(ID_PLANILHA)
    
    nome_aba_principal = planilha.worksheets()[0].title
    # Atualizar aba principal
    atualizar_aba(planilha, nome_aba_principal, df_novo)
    
    if not df_novo.empty:
        atualizar_aba(planilha, "Nomeações", df_novo[df_novo["Situação"] == "Nomeado"])
        atualizar_aba(planilha, "Exonerações", df_novo[df_novo["Situação"] == "Exonerado"])
        atualizar_aba(planilha, "Vacâncias", df_novo[df_novo["Situação"] == "Vacância"])
        atualizar_aba(planilha, "Transferências", df_novo[df_novo["Situação"] == "Transferido"])
        atualizar_aba(planilha, "Sem Efeito", df_novo[df_novo["Situação"].isin(["Sem Efeito - Exoneração", "Sem Efeito - Nomeação"])])

    logging.info(f"Planilha Google atualizada com sucesso. ID: {ID_PLANILHA}")

# ---------------------------------------------------------------------------
# EXECUÇÃO PRINCIPAL
# ---------------------------------------------------------------------------

def main():
    logging.info("1/4 - Baixando última edição do Diário Oficial...")
    caminho_pdf = baixar_ultima_edicao()

    logging.info("2/4 - Extraindo texto do PDF...")
    texto = extrair_texto(caminho_pdf)

    logging.info("3/4 - Localizando exonerações, vacâncias, nomeações e atos sem efeito...")
    movimentacoes = extrair_movimentacoes(texto)

    nomeados_cnt = sum(1 for m in movimentacoes if m["Situação"] == "Nomeado")
    exonerados_cnt = sum(1 for m in movimentacoes if m["Situação"] == "Exonerado")
    vacancias_cnt = sum(1 for m in movimentacoes if m["Situação"] == "Vacância")
    transferencias_cnt = sum(1 for m in movimentacoes if m["Situação"] == "Transferido")
    sem_efeito_cnt = sum(1 for m in movimentacoes if m["Situação"].startswith("Sem Efeito"))

    logging.info(f"{nomeados_cnt} nomeação(ões) encontrada(s)")
    logging.info(f"{exonerados_cnt} exoneração(ões) encontrada(s)")
    logging.info(f"{vacancias_cnt} vacância(s) encontrada(s)")
    logging.info(f"{transferencias_cnt} transferência(s) de lotação encontrada(s)")
    logging.info(f"{sem_efeito_cnt} ato(s) 'tornado(s) sem efeito' encontrado(s)")

    logging.info("4/4 - Atualizando planilha no Google Sheets...")
    gerar_planilha(movimentacoes, caminho_pdf)


if __name__ == "__main__":
    try:
        main()
    except Exception as erro:
        logging.exception(f"Falha na execução: {erro}")
        sys.exit(1)