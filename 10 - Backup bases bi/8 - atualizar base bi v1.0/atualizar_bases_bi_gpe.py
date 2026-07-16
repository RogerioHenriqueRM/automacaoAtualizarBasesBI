from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError 
from pathlib import Path
import subprocess
import time
import sys

# =========================================================
# AUTOMACAO GPE - SAP -> PASTAS LOCAIS -> SHAREPOINT
# =========================================================
# Antes de executar:
# 1) Abra o SAP e deixe logado.
# 2) Ajuste CAMINHO_VBS e CAMINHO_BAT_EDGE se necessario.
# 3) Execute: python atualizar_bases_bi_gpe.py
# =========================================================

CAMINHO_VBS = r"C:\Users\rogerio.matos\Documents\28 - extração SAP\SAP_GPE_TODAS_TRANSACOES_v1.0.vbs"
CAMINHO_BAT_EDGE = r"C:\Users\rogerio.matos\Documents\9 - automacao_.zip\abrir_edge.bat"
EDGE_CDP = "http://localhost:9222"

# Tempo base para o SharePoint carregar paginas e processar uploads.
TEMPO_CARREGAR_PAGINA_MS = 15000
TEMPO_APOS_UPLOAD_MS = 15000

TRANSACOES = {
    "MB51": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\1 - MB51",
        "nome_base": "MB51 ATUALIZADA",
        "sharepoint": "https://fsagrbr.sharepoint.com/sites/gpe/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2Fgpe%2FDocumentos%20Compartilhados%2FBase%20de%20Dados%20%2D%20Gest%C3%A3o%20e%20Planejamento%20de%20Estoques%2FBases%20do%20BI%2FMB51&sortField=Created&isAscending=false&viewid=a61170b8%2D2935%2D411f%2Da644%2D4e565573140d",
    },
    "MB52": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\2 - MB52",
        "nome_base": "MB52 ATUALIZADA",
        "sharepoint": "https://fsagrbr.sharepoint.com/sites/gpe/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2Fgpe%2FDocumentos%20Compartilhados%2FBase%20de%20Dados%20%2D%20Gest%C3%A3o%20e%20Planejamento%20de%20Estoques%2FBases%20do%20BI%2FMB52&viewid=a61170b8%2D2935%2D411f%2Da644%2D4e565573140d&ga=1",
    },
    "ME2N": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\4 - ME2N",
        "nome_base": "ME2N ATUALIZADA",
        "sharepoint": "https://fsagrbr.sharepoint.com/sites/gpe/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2Fgpe%2FDocumentos%20Compartilhados%2FBase%20de%20Dados%20%2D%20Gest%C3%A3o%20e%20Planejamento%20de%20Estoques%2FBases%20do%20BI%2FME2N&sortField=Created&isAscending=false&viewid=a61170b8%2D2935%2D411f%2Da644%2D4e565573140d",
    },
    "MM60": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\5 - MM60",
        "nome_base": "MM60 ATUALIZADA",
        "sharepoint": "https://fsagrbr.sharepoint.com/sites/gpe/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2Fgpe%2FDocumentos%20Compartilhados%2FBase%20de%20Dados%20%2D%20Gest%C3%A3o%20e%20Planejamento%20de%20Estoques%2FBases%20do%20BI%2FMM60&sortField=Created&isAscending=false&viewid=a61170b8%2D2935%2D411f%2Da644%2D4e565573140d",
    },
    "ZMM039": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\6 - ZMM039",
        "nome_base": "ZMM039_GPE_ATUALIZADA",
        "sharepoint": "https://fsagrbr.sharepoint.com/sites/gpe/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2Fgpe%2FDocumentos%20Compartilhados%2FBase%20de%20Dados%20%2D%20Gest%C3%A3o%20e%20Planejamento%20de%20Estoques%2FBases%20do%20BI%2FZMM039&sortField=Created&isAscending=false&viewid=a61170b8%2D2935%2D411f%2Da644%2D4e565573140d",
    },
    "ZMM028": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\7 - ZMM028",
        "nome_base": "ZMM028 ATUALIZADA",
        "sharepoint": "https://fsagrbr.sharepoint.com/sites/gpe/Documentos%20Compartilhados/Forms/AllItems.aspx?id=%2Fsites%2Fgpe%2FDocumentos%20Compartilhados%2FBase%20de%20Dados%20%2D%20Gest%C3%A3o%20e%20Planejamento%20de%20Estoques%2FBases%20do%20BI%2FZMM028&viewid=a61170b8%2D2935%2D411f%2Da644%2D4e565573140d&ga=1",
    },
}

EXTENSOES_VALIDAS = (".xlsx", ".xls", ".xlsm", ".csv", ".txt")


def log(texto: str) -> None:
    print(texto, flush=True)


def garantir_pastas_locais() -> None:
    for cfg in TRANSACOES.values():
        Path(cfg["pasta_local"]).mkdir(parents=True, exist_ok=True)


def executar_vbs_sap() -> None:
    vbs = Path(CAMINHO_VBS)
    if not vbs.exists():
        raise FileNotFoundError(f"VBS nao encontrado: {vbs}")

    log("Executando VBS do SAP...")
    resultado = subprocess.run(
        ["cscript.exe", "//nologo", str(vbs)],
        shell=False,
        capture_output=True,
        text=True,
    )

    if resultado.stdout:
        log(resultado.stdout)
    if resultado.stderr:
        log(resultado.stderr)

    if resultado.returncode != 0:
        raise RuntimeError(f"VBS finalizou com erro. Codigo: {resultado.returncode}")


def arquivo_gerado(transacao: str, cfg: dict) -> Path:
    pasta = Path(cfg["pasta_local"])
    nome_base = cfg["nome_base"].lower()

    candidatos = []
    for item in pasta.iterdir():
        if item.is_file() and item.suffix.lower() in EXTENSOES_VALIDAS:
            if item.stem.lower() == nome_base or item.stem.lower().startswith(nome_base):
                candidatos.append(item)

    if not candidatos:
        arquivos = [p.name for p in pasta.iterdir() if p.is_file()]
        raise FileNotFoundError(
            f"Nao encontrei arquivo da {transacao} em {pasta}. Arquivos encontrados: {arquivos}"
        )

    return max(candidatos, key=lambda p: p.stat().st_mtime)


def abrir_edge_debug() -> None:
    bat = Path(CAMINHO_BAT_EDGE)
    if not bat.exists():
        raise FileNotFoundError(f"BAT do Edge nao encontrado: {bat}")

    log("Abrindo Edge pela porta debug...")
    subprocess.Popen(str(bat), shell=True)
    time.sleep(7)


def clicar_primeiro_que_existir(page, tentativas: list, timeout: int = 5000) -> bool:
    for tentativa in tentativas:
        try:
            tipo = tentativa[0]
            if tipo == "role":
                _, role, name = tentativa
                page.get_by_role(role, name=name).click(timeout=timeout)
            elif tipo == "text":
                _, text, exact = tentativa
                page.get_by_text(text, exact=exact).click(timeout=timeout)
            else:
                continue
            return True
        except Exception:
            pass
    return False


def enviar_arquivo_sharepoint(page, arquivo: Path, url_destino: str, transacao: str) -> None:
    log(f"\nAbrindo pasta SharePoint da {transacao}...")
    page.goto(url_destino)
    page.wait_for_timeout(TEMPO_CARREGAR_PAGINA_MS)

    log(f"Enviando: {arquivo}")

    abriu_menu = clicar_primeiro_que_existir(
        page,
        [
            ("role", "button", "Criar ou carregar"),
            ("role", "button", "Carregar"),
            ("role", "button", "Upload"),
            ("text", "Criar ou carregar", True),
            ("text", "Carregar", True),
            ("text", "Upload", True),
        ],
        timeout=10000,
    )

    if not abriu_menu:
        raise RuntimeError("Nao consegui abrir o menu de upload/carregar no SharePoint.")

    page.wait_for_timeout(3000)

    with page.expect_file_chooser(timeout=15000) as fc_info:
        clicou_arquivo = clicar_primeiro_que_existir(
            page,
            [
                ("role", "menuitem", "Carregamento de arquivos"),
                ("role", "menuitem", "Arquivos"),
                ("role", "menuitem", "Files"),
                ("text", "Carregamento de arquivos", True),
                ("text", "Arquivos", True),
                ("text", "Files upload", True),
                ("text", "Files", True),
            ],
            timeout=8000,
        )
        if not clicou_arquivo:
            raise RuntimeError("Nao encontrei a opcao de carregar arquivo no menu.")

    file_chooser = fc_info.value
    file_chooser.set_files(str(arquivo))

    log("Aguardando processamento do upload...")
    page.wait_for_timeout(TEMPO_APOS_UPLOAD_MS)

    substituiu = clicar_primeiro_que_existir(
        page,
        [
            ("role", "button", "Substituir"),
            ("role", "button", "Replace"),
            ("text", "Substituir", True),
            ("text", "Replace", True),
        ],
        timeout=8000,
    )

    if substituiu:
        log("Arquivo existente substituido.")
        page.wait_for_timeout(8000)
    else:
        log("Arquivo novo enviado ou substituicao nao foi solicitada.")

    log(f"Upload concluido: {transacao}")


def enviar_todos_sharepoint(arquivos_por_transacao: dict) -> None:
    abrir_edge_debug()

    with sync_playwright() as p:
        log("Conectando ao Edge ja logado...")
        browser = p.chromium.connect_over_cdp(EDGE_CDP)

        if not browser.contexts:
            raise RuntimeError("Nao encontrei contexto do Edge. Verifique se o .bat abriu o Edge na porta 9222.")

        context = browser.contexts[0]
        page = context.new_page()

        for transacao, arquivo in arquivos_por_transacao.items():
            cfg = TRANSACOES[transacao]
            enviar_arquivo_sharepoint(page, arquivo, cfg["sharepoint"], transacao)

        page.close()


def main() -> int:
    try:
        garantir_pastas_locais()
        executar_vbs_sap()

        log("\nConferindo arquivos gerados pelo SAP...")
        arquivos_por_transacao = {}
        for transacao, cfg in TRANSACOES.items():
            arquivo = arquivo_gerado(transacao, cfg)
            arquivos_por_transacao[transacao] = arquivo
            log(f"{transacao}: {arquivo}")

        enviar_todos_sharepoint(arquivos_por_transacao)
        log("\nAUTOMACAO FINALIZADA COM SUCESSO.")
        return 0

    except Exception as erro:
        log("\nERRO NA AUTOMACAO")
        log(str(erro))
        return 1


if __name__ == "__main__":
    sys.exit(main())
    
    