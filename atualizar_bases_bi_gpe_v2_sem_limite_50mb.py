from pathlib import Path
import subprocess
import shutil
import sys
import time
import os

# =========================================================
# AUTOMACAO GPE - SAP -> PASTAS LOCAIS -> SHAREPOINT SYNC
# Versao v2.1
#
# Como funciona:
# 1) Executa o VBS do SAP.
# 2) Localiza os arquivos extraidos nas pastas locais.
# 3) Copia/substitui os arquivos na pasta local sincronizada.
# 4) Atualiza a data do arquivo para ajudar o OneDrive a detectar a alteracao.
# 5) O proprio OneDrive envia os arquivos para o SharePoint.
# =========================================================

CAMINHO_VBS = (
    r"C:\Users\rogerio.matos\Documents\28 - extração SAP"
    r"\SAP_GPE_TODAS_TRANSACOES_v1.4.vbs"
)

# Pasta local sincronizada com o SharePoint/OneDrive.
SHAREPOINT_SYNC_ROOT = (
    r"C:\Users\rogerio.matos"
    r"\FS Agrisolutions Industria de Biocombustiveis Ltda"
    r"\Gestão e Planejamento de Estoques - Bases do BI"
)

TRANSACOES = {
    "MB51": {
        "pasta_local": (
            r"C:\Users\rogerio.matos\Documents"
            r"\21 - atualizar bases bi\1 - MB51"
        ),
        "nome_base": "MB51 ATUALIZADA",
        "pasta_sharepoint": "MB51",
    },

    "MB52": {
        "pasta_local": (
            r"C:\Users\rogerio.matos\Documents"
            r"\21 - atualizar bases bi\2 - MB52"
        ),
        "nome_base": "MB52 ATUALIZADA",
        "pasta_sharepoint": "MB52",
    },

    "ME2N": {
        "pasta_local": (
            r"C:\Users\rogerio.matos\Documents"
            r"\21 - atualizar bases bi\4 - ME2N"
        ),
        "nome_base": "ME2N ATUALIZADA",
        "pasta_sharepoint": "ME2N",
    },

    "MM60": {
        "pasta_local": (
            r"C:\Users\rogerio.matos\Documents"
            r"\21 - atualizar bases bi\5 - MM60"
        ),
        "nome_base": "MM60 ATUALIZADA",
        "pasta_sharepoint": "MM60",
    },

    "ZMM039": {
        "pasta_local": (
            r"C:\Users\rogerio.matos\Documents"
            r"\21 - atualizar bases bi\6 - ZMM039"
        ),
        "nome_base": "ZMM039_GPE_ATUALIZADA",
        "pasta_sharepoint": "ZMM039",
    },

    "ZMM028": {
        "pasta_local": (
            r"C:\Users\rogerio.matos\Documents"
            r"\21 - atualizar bases bi\7 - ZMM028"
        ),
        "nome_base": "ZMM028 ATUALIZADA",
        "pasta_sharepoint": "ZMM028",
    },

    "MB25": {
        "pasta_local": (
            r"C:\Users\rogerio.matos\Documents"
            r"\21 - atualizar bases bi\9 -  MB25"
        ),
        "nome_base": "MB25 ATUALIZADA",
        "pasta_sharepoint": "MB25",
    },

    "MI24": {
        "pasta_local": (
            r"C:\Users\rogerio.matos\Documents"
            r"\21 - atualizar bases bi\3 - MI24"
        ),
        "nome_base": "MI24 ATUALIZADA",

        # Nomes que o VBS pode gerar na pasta local.
        "nomes_origem": [
            "MI24 ATUALIZADA",
            "MI24_GPE",
            "MI24",
        ],

        "pasta_sharepoint": "MI24",
    },
}

EXTENSOES_VALIDAS = (
    ".xlsx",
    ".xls",
    ".xlsm",
    ".csv",
    ".txt",
)


def log(msg: str) -> None:
    print(msg, flush=True)


def executar_vbs_sap() -> None:
    vbs = Path(CAMINHO_VBS)

    if not vbs.exists():
        raise FileNotFoundError(
            f"VBS nao encontrado: {vbs}"
        )

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
        raise RuntimeError(
            f"O VBS terminou com erro. "
            f"Codigo: {resultado.returncode}"
        )


def localizar_arquivo_gerado(
    transacao: str,
    cfg: dict
) -> Path:

    pasta = Path(cfg["pasta_local"])

    if not pasta.exists():
        raise FileNotFoundError(
            f"Pasta local nao encontrada para "
            f"{transacao}: {pasta}"
        )

    nomes_origem = cfg.get(
        "nomes_origem",
        [cfg["nome_base"]],
    )

    nomes_origem = [
        nome.lower().strip()
        for nome in nomes_origem
    ]

    candidatos = []

    for item in pasta.iterdir():

        if not item.is_file():
            continue

        if item.suffix.lower() not in EXTENSOES_VALIDAS:
            continue

        # Ignora arquivos temporarios do Excel.
        if item.name.startswith("~$"):
            continue

        stem = item.stem.lower().strip()

        nome_valido = any(
            stem == nome
            or stem.startswith(nome)
            for nome in nomes_origem
        )

        if nome_valido:
            candidatos.append(item)

    if not candidatos:
        arquivos_encontrados = [
            p.name
            for p in pasta.iterdir()
            if p.is_file()
        ]

        raise FileNotFoundError(
            f"Arquivo da {transacao} nao encontrado.\n"
            f"Pasta: {pasta}\n"
            f"Nomes procurados: {nomes_origem}\n"
            f"Arquivos encontrados: "
            f"{arquivos_encontrados}"
        )

    arquivo_mais_recente = max(
        candidatos,
        key=lambda p: p.stat().st_mtime,
    )

    return arquivo_mais_recente


def esperar_arquivo_estabilizar(
    arquivo: Path,
    tentativas: int = 15,
    intervalo: int = 2
) -> None:
    """
    Aguarda o SAP/Excel terminar de gravar o arquivo.

    O tamanho precisa permanecer igual em duas verificacoes
    consecutivas.
    """

    tamanho_anterior = -1
    tamanho_estavel = 0

    for tentativa in range(1, tentativas + 1):

        if not arquivo.exists():
            time.sleep(intervalo)
            continue

        try:
            tamanho_atual = arquivo.stat().st_size

            if tamanho_atual == tamanho_anterior:
                tamanho_estavel += 1
            else:
                tamanho_estavel = 0

            if tamanho_estavel >= 2:
                log(
                    f"Arquivo estabilizado: "
                    f"{arquivo.name}"
                )
                return

            tamanho_anterior = tamanho_atual

        except OSError:
            pass

        log(
            f"Aguardando o arquivo ficar disponivel "
            f"({tentativa}/{tentativas})..."
        )

        time.sleep(intervalo)

    raise RuntimeError(
        f"O arquivo ainda parece estar em uso "
        f"ou sendo gravado: {arquivo}"
    )


def descobrir_raiz_sharepoint_sync() -> Path:

    if SHAREPOINT_SYNC_ROOT.strip():

        raiz = Path(SHAREPOINT_SYNC_ROOT)

        if raiz.exists():
            return raiz

        raise FileNotFoundError(
            "SHAREPOINT_SYNC_ROOT preenchido, "
            f"mas nao existe: {raiz}"
        )

    usuario = Path.home()

    possiveis_bases = [
        usuario,
        usuario / "OneDrive - FS",
        usuario / "OneDrive - FS AGRISOLUTIONS",
        usuario / "FS AGRISOLUTIONS",
        usuario / "FS Agrisolutions",
    ]

    nomes_alvo = [
        (
            Path(
                "Base de Dados - Gestão e "
                "Planejamento de Estoques"
            )
            / "Bases do BI"
        ),
        (
            Path(
                "Base de Dados - Gestao e "
                "Planejamento de Estoques"
            )
            / "Bases do BI"
        ),
    ]

    for base in possiveis_bases:

        if not base.exists():
            continue

        for alvo in nomes_alvo:

            encontrados = list(
                base.rglob(str(alvo))
            )

            for encontrado in encontrados:

                todas_pastas_existem = all(
                    (
                        encontrado
                        / cfg["pasta_sharepoint"]
                    ).exists()
                    for cfg in TRANSACOES.values()
                )

                if todas_pastas_existem:
                    return encontrado

    raise FileNotFoundError(
        "Nao encontrei automaticamente a pasta "
        "sincronizada do SharePoint.\n"
        "Abra a pasta do SharePoint no navegador, "
        "clique em 'Sincronizar' ou "
        "'Adicionar atalho ao OneDrive'.\n"
        "Depois copie o caminho local da pasta "
        "'Bases do BI' e preencha a variavel "
        "SHAREPOINT_SYNC_ROOT."
    )


def remover_arquivo_destino(
    arquivo_destino: Path,
    tentativas: int = 10
) -> None:

    if not arquivo_destino.exists():
        return

    for tentativa in range(1, tentativas + 1):

        try:
            os.remove(arquivo_destino)

            log(
                "Arquivo antigo removido para "
                "substituicao."
            )

            return

        except PermissionError:

            if tentativa == tentativas:
                raise PermissionError(
                    "Nao foi possivel remover o arquivo "
                    "antigo porque ele esta aberto ou "
                    "bloqueado:\n"
                    f"{arquivo_destino}"
                )

            log(
                f"Arquivo de destino em uso. "
                f"Nova tentativa "
                f"({tentativa}/{tentativas})..."
            )

            time.sleep(3)


def copiar_arquivo_com_tentativas(
    arquivo_origem: Path,
    arquivo_destino: Path,
    tentativas: int = 10
) -> None:

    for tentativa in range(1, tentativas + 1):

        try:
            # shutil.copy nao preserva todos os metadados
            # do arquivo original. Isso ajuda o OneDrive
            # a identificar o arquivo como novo/alterado.
            shutil.copy(
                arquivo_origem,
                arquivo_destino,
            )

            # Atualiza a data e hora de modificacao.
            arquivo_destino.touch()

            return

        except PermissionError:

            if tentativa == tentativas:
                raise PermissionError(
                    "Nao foi possivel copiar o arquivo "
                    "porque a origem ou o destino esta "
                    "aberto/bloqueado.\n"
                    f"Origem: {arquivo_origem}\n"
                    f"Destino: {arquivo_destino}"
                )

            log(
                f"Arquivo em uso. Nova tentativa "
                f"({tentativa}/{tentativas})..."
            )

            time.sleep(3)


def verificar_copia(
    arquivo_origem: Path,
    arquivo_destino: Path
) -> None:

    if not arquivo_destino.exists():
        raise RuntimeError(
            "O arquivo nao foi criado no destino:\n"
            f"{arquivo_destino}"
        )

    tamanho_origem = arquivo_origem.stat().st_size
    tamanho_destino = arquivo_destino.stat().st_size

    if tamanho_origem != tamanho_destino:
        raise RuntimeError(
            "Os tamanhos dos arquivos sao diferentes.\n"
            f"Origem: {tamanho_origem} bytes\n"
            f"Destino: {tamanho_destino} bytes"
        )


def copiar_para_sharepoint_sync(
    arquivos_por_transacao: dict,
    raiz_sharepoint: Path
) -> None:

    log(
        f"\nPasta sincronizada encontrada: "
        f"{raiz_sharepoint}"
    )

    for transacao, arquivo_origem in (
        arquivos_por_transacao.items()
    ):

        cfg = TRANSACOES[transacao]

        pasta_destino = (
            raiz_sharepoint
            / cfg["pasta_sharepoint"]
        )

        if not pasta_destino.exists():
            raise FileNotFoundError(
                f"Pasta de destino da {transacao} "
                f"nao encontrada:\n"
                f"{pasta_destino}"
            )

        # Usa sempre o nome padrao definido na configuracao,
        # preservando somente a extensao do arquivo de origem.
        arquivo_destino = (
            pasta_destino
            / (
                f'{cfg["nome_base"]}'
                f'{arquivo_origem.suffix.lower()}'
            )
        )

        log(f"\n{transacao}")
        log(f"Origem : {arquivo_origem}")
        log(f"Destino: {arquivo_destino}")

        esperar_arquivo_estabilizar(
            arquivo_origem
        )

        remover_arquivo_destino(
            arquivo_destino
        )

        # Pequena pausa entre a exclusao e a nova copia.
        # Isso ajuda o OneDrive a detectar a alteracao.
        time.sleep(1)

        copiar_arquivo_com_tentativas(
            arquivo_origem,
            arquivo_destino,
        )

        verificar_copia(
            arquivo_origem,
            arquivo_destino,
        )

        tamanho_mb = (
            arquivo_destino.stat().st_size
            / 1024
            / 1024
        )

        log(
            f"Arquivo copiado e verificado com sucesso. "
            f"Tamanho: {tamanho_mb:.2f} MB"
        )

        # Aguarda o OneDrive detectar a alteracao antes
        # de passar para o proximo arquivo.
        time.sleep(5)

    log(
        "\nTodos os arquivos foram copiados para "
        "a pasta sincronizada."
    )

    log(
        "O OneDrive agora deve enviar os arquivos "
        "para o SharePoint."
    )


def main() -> int:

    try:
        executar_vbs_sap()

        log("\nLocalizando arquivos gerados...")

        arquivos_por_transacao = {}

        for transacao, cfg in TRANSACOES.items():

            arquivo = localizar_arquivo_gerado(
                transacao,
                cfg,
            )

            arquivos_por_transacao[
                transacao
            ] = arquivo

            tamanho_mb = (
                arquivo.stat().st_size
                / 1024
                / 1024
            )

            log(
                f"{transacao}: "
                f"{arquivo.name} - "
                f"{tamanho_mb:.2f} MB"
            )

        raiz_sharepoint = (
            descobrir_raiz_sharepoint_sync()
        )

        copiar_para_sharepoint_sync(
            arquivos_por_transacao,
            raiz_sharepoint,
        )

        log(
            "\nAUTOMACAO FINALIZADA COM SUCESSO."
        )

        return 0

    except Exception as erro:

        log("\nERRO NA AUTOMACAO")
        log(str(erro))

        return 1


if __name__ == "__main__":

    codigo = main()

    input(
        "\nPressione ENTER para fechar..."
    )

    sys.exit(codigo)