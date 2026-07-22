from pathlib import Path
import subprocess
import shutil
import sys
import time

# =========================================================
# AUTOMACAO GPE - SAP -> PASTAS LOCAIS -> SHAREPOINT SYNC
# Versao v2: sem FileChooser.set_files, portanto sem limite de 50 MB
# =========================================================
# Como funciona:
# 1) Executa o VBS do SAP.
# 2) Localiza os arquivos extraidos nas pastas locais.
# 3) Copia/substitui os arquivos na pasta local sincronizada do SharePoint/OneDrive.
# 4) O proprio OneDrive envia os arquivos para o SharePoint.
# =========================================================
 
CAMINHO_VBS = False
 
CAMINHO_VBS = r"C:\Users\rogerio.matos\Documents\28 - extração SAP\SAP_GPE_TODAS_TRANSACOES_v1.4.vbs"

# Se a descoberta automatica nao encontrar a pasta sincronizada, preencha aqui.
# Exemplo:
# SHAREPOINT_SYNC_ROOT = r"C:\Users\rogerio.matos\FS AGRISOLUTIONS\gpe - Documentos\Base de Dados - Gestão e Planejamento de Estoques\Bases do BI"
SHAREPOINT_SYNC_ROOT = r"C:\Users\rogerio.matos\FS Agrisolutions Industria de Biocombustiveis Ltda\Gestão e Planejamento de Estoques - Bases do BI"

TRANSACOES = {
    "MB51": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\1 - MB51",
        "nome_base": "MB51 ATUALIZADA",
        "pasta_sharepoint": "MB51",
    },
    "MB52": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\2 - MB52",
        "nome_base": "MB52 ATUALIZADA",
        "pasta_sharepoint": "MB52",
    },
    "ME2N": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\4 - ME2N",
        "nome_base": "ME2N ATUALIZADA",
        "pasta_sharepoint": "ME2N",
    },
    "MM60": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\5 - MM60",
        "nome_base": "MM60 ATUALIZADA",
        "pasta_sharepoint": "MM60",
    },
    "ZMM039": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\6 - ZMM039",
        "nome_base": "ZMM039_GPE_ATUALIZADA",
        "pasta_sharepoint": "ZMM039",
    },
    "ZMM028": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\7 - ZMM028",
        "nome_base": "ZMM028 ATUALIZADA",
        "pasta_sharepoint": "ZMM028",
    },
    "MB25": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\9 -  MB25",
        "nome_base": "MB25 ATUALIZADA",
        "pasta_sharepoint": "MB25"
    },
    "MI24": {
        "pasta_local": r"C:\Users\rogerio.matos\Documents\21 - atualizar bases bi\3 - MI24",
        "nome_base": "MI24 ATUALIZADA",
        "pasta_sharepoint": "MI24"
    }
}

EXTENSOES_VALIDAS = (".xlsx", ".xls", ".xlsm", ".csv", ".txt")


def log(msg: str) -> None:
    print(msg, flush=True)


def executar_vbs_sap() -> None:
    vbs = Path(CAMINHO_VBS)
    if not vbs.exists():
        raise FileNotFoundError(f"VBS nao encontrado: {vbs}")

    log("Executando extracao SAP pelo VBS...")
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
        raise RuntimeError(f"O VBS terminou com erro. Codigo: {resultado.returncode}")


def localizar_arquivo_gerado(transacao: str, cfg: dict) -> Path:
    pasta = Path(cfg["pasta_local"])
    if not pasta.exists():
        raise FileNotFoundError(f"Pasta local nao encontrada para {transacao}: {pasta}")

    nome_base = cfg["nome_base"].lower()
    candidatos = []

    for item in pasta.iterdir():
        if item.is_file() and item.suffix.lower() in EXTENSOES_VALIDAS:
            stem = item.stem.lower()
            if stem == nome_base or stem.startswith(nome_base):
                candidatos.append(item)

    if not candidatos:
        arquivos = [p.name for p in pasta.iterdir() if p.is_file()]
        raise FileNotFoundError(
            f"Arquivo da {transacao} nao encontrado em {pasta}. Arquivos encontrados: {arquivos}"
        )

    return max(candidatos, key=lambda p: p.stat().st_mtime)


def descobrir_raiz_sharepoint_sync() -> Path:
    if SHAREPOINT_SYNC_ROOT.strip():
        raiz = Path(SHAREPOINT_SYNC_ROOT)
        if raiz.exists():
            return raiz
        raise FileNotFoundError(f"SHAREPOINT_SYNC_ROOT preenchido, mas nao existe: {raiz}")

    usuario = Path.home()
    possiveis_bases = [
        usuario,
        usuario / "OneDrive - FS",
        usuario / "OneDrive - FS AGRISOLUTIONS",
        usuario / "FS AGRISOLUTIONS",
        usuario / "FS Agrisolutions",
    ]

    nomes_alvo = [
        Path("Base de Dados - Gestão e Planejamento de Estoques") / "Bases do BI",
        Path("Base de Dados - Gestao e Planejamento de Estoques") / "Bases do BI",
    ]

    for base in possiveis_bases:
        if not base.exists():
            continue
        for alvo in nomes_alvo:
            encontrados = list(base.rglob(str(alvo)))
            for encontrado in encontrados:
                if all((encontrado / cfg["pasta_sharepoint"]).exists() for cfg in TRANSACOES.values()):
                    return encontrado

    raise FileNotFoundError(
        "Nao encontrei automaticamente a pasta sincronizada do SharePoint.\n"
        "Abra a pasta do SharePoint no navegador, clique em 'Sincronizar' ou 'Adicionar atalho ao OneDrive',\n"
        "depois copie o caminho local da pasta 'Bases do BI' e preencha a variavel SHAREPOINT_SYNC_ROOT no codigo."
    )


def copiar_para_sharepoint_sync(arquivos_por_transacao: dict, raiz_sharepoint: Path) -> None:
    log(f"\nPasta sincronizada encontrada: {raiz_sharepoint}")

    for transacao, arquivo_origem in arquivos_por_transacao.items():
        cfg = TRANSACOES[transacao]
        pasta_destino = raiz_sharepoint / cfg["pasta_sharepoint"]

        if not pasta_destino.exists():
            raise FileNotFoundError(f"Pasta de destino da {transacao} nao encontrada: {pasta_destino}")

        arquivo_destino = pasta_destino / arquivo_origem.name

        log(f"\n{transacao}")
        log(f"Origem : {arquivo_origem}")
        log(f"Destino: {arquivo_destino}")

        if arquivo_destino.exists():
            arquivo_destino.unlink()
            log("Arquivo antigo removido para substituicao.")

        shutil.copy2(arquivo_origem, arquivo_destino)
        log("Arquivo copiado para a pasta sincronizada do SharePoint.")

    log("\nAgora o OneDrive/SharePoint vai sincronizar os arquivos automaticamente.")


def main() -> int:
    try:
        executar_vbs_sap()

        log("\nLocalizando arquivos gerados...")
        arquivos_por_transacao = {}
        for transacao, cfg in TRANSACOES.items():
            arquivo = localizar_arquivo_gerado(transacao, cfg)
            arquivos_por_transacao[transacao] = arquivo
            tamanho_mb = arquivo.stat().st_size / 1024 / 1024
            log(f"{transacao}: {arquivo.name} - {tamanho_mb:.2f} MB")

        raiz_sharepoint = descobrir_raiz_sharepoint_sync()
        copiar_para_sharepoint_sync(arquivos_por_transacao, raiz_sharepoint)

        log("\nAUTOMACAO FINALIZADA COM SUCESSO.")
        return 0

    except Exception as erro:
        log("\nERRO NA AUTOMACAO")
        log(str(erro))
        return 1


if __name__ == "__main__":
    codigo = main()
    input("\nPressione ENTER para fechar...")
    sys.exit(codigo)
