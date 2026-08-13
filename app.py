import os
import re
import time
import json
import logging
import threading
from pathlib import Path
from datetime import datetime, timedelta
from io import BytesIO

import requests
from bs4 import BeautifulSoup
from flask import Flask, jsonify, send_file
from openpyxl import Workbook


app = Flask(__name__)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)


# ==========================================================
# CONFIGURAÇÕES
# ==========================================================

LOTERIAS = {
    "PT-SP": {
        "slug": "resultados-pt-sp-do-dia-",
        "url_base": "https://www.resultadofacil.com.br/",
    },

    "LOTEP": {
        "slug": "resultados-lotep-do-dia-",
        "url_base": "https://www.resultadofacil.com.br/",
    },

    "PT-BA": {
        "slug": "resultados-paratodos-bahia-do-dia-",
        "url_base": "https://www.resultadofacil.com.br/",
    },

    "LOTECE": {
        "slug": "resultados-lotece---loteria-dos-sonhos-do-dia-",
        "url_base": "https://www.resultadofacil.com.br/",
    },
}


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/142.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
}

# ==========================================================
# COLETA HISTÓRICA / DISCO PERSISTENTE
# ==========================================================

DATA_INICIAL_COLETA = datetime(
    2025,
    1,
    1
)

DATA_FINAL_COLETA = datetime(
    2026,
    8,
    12
)

DIRETORIO_DADOS = Path(
    os.environ.get(
        "DATA_DIR",
        "/var/data"
    )
)

DIRETORIO_DADOS.mkdir(
    parents=True,
    exist_ok=True
)

ARQUIVO_ESTADO = (
    DIRETORIO_DADOS
    / "estado_coleta.json"
)

ARQUIVO_RESULTADOS = (
    DIRETORIO_DADOS
    / "resultados.jsonl"
)

ARQUIVO_AUDITORIA = (
    DIRETORIO_DADOS
    / "auditoria.jsonl"
)

ARQUIVO_EXCEL = (
    DIRETORIO_DADOS
    / "resultadofacil_2025_2026.xlsx"
)

COLETA_LOCK = threading.Lock()

def estado_inicial_coleta():
    return {
        "status": "nao_iniciada",
        "data_inicial": DATA_INICIAL_COLETA.strftime(
            "%d/%m/%Y"
        ),
        "data_final": DATA_FINAL_COLETA.strftime(
            "%d/%m/%Y"
        ),
        "data_atual": None,
        "loteria_atual": None,
        "paginas_processadas": 0,
        "resultados_coletados": 0,
        "falhas": 0,
        "progresso": 0.0,
        "mensagem": (
            "A coleta ainda não foi iniciada."
        ),
    }


def salvar_estado_coleta(estado):
    temporario = ARQUIVO_ESTADO.with_suffix(
        ".tmp"
    )

    with open(
        temporario,
        "w",
        encoding="utf-8"
    ) as arquivo:
        json.dump(
            estado,
            arquivo,
            ensure_ascii=False,
            indent=2,
        )

    os.replace(
        temporario,
        ARQUIVO_ESTADO
    )


def carregar_estado_coleta():
    if not ARQUIVO_ESTADO.exists():
        estado = estado_inicial_coleta()
        salvar_estado_coleta(
            estado
        )
        return estado

    try:
        with open(
            ARQUIVO_ESTADO,
            "r",
            encoding="utf-8"
        ) as arquivo:
            return json.load(
                arquivo
            )

    except Exception:
        logging.exception(
            "Erro ao carregar estado da coleta."
        )

        return estado_inicial_coleta()


def adicionar_jsonl(
    caminho,
    registro
):
    with open(
        caminho,
        "a",
        encoding="utf-8"
    ) as arquivo:
        arquivo.write(
            json.dumps(
                registro,
                ensure_ascii=False,
            )
        )
        arquivo.write("\n")


# ==========================================================
# UTILITÁRIOS
# ==========================================================

def somente_digitos(valor):
    return re.sub(r"\D", "", str(valor or ""))


def formatar_milhar(valor):
    digitos = somente_digitos(valor)

    if not digitos:
        return ""

    return digitos[-4:].zfill(4)


def normalizar_texto(texto):
    return re.sub(
        r"\s+",
        " ",
        str(texto or "")
    ).strip()


def calcular_premios_6_7(premios):
    valores = [int(p) for p in premios[:5]]

    soma = sum(valores)

    m6 = str(soma)[-4:].zfill(4)

    produto = valores[0] * valores[1]

    m7 = str(
        produto // 1000
    )[-3:].zfill(3)

    return m6, m7


def montar_url(loteria, data_obj):
    cfg = LOTERIAS[loteria]

    data_iso = data_obj.strftime(
        "%Y-%m-%d"
    )

    return (
        cfg["url_base"]
        + cfg["slug"]
        + data_iso
    )


# ==========================================================
# EXTRAÇÃO
# ==========================================================

def extrair_horario(texto):
    texto = normalizar_texto(
        texto
    ).lower()

    match = re.search(
        r"\b(\d{1,2})\s*(?:h|horas?)",
        texto
    )

    if match:
        return match.group(1).zfill(2)

    match = re.search(
        r"\b(\d{1,2}):\d{2}\b",
        texto
    )

    if match:
        return match.group(1).zfill(2)

    return ""


def parece_federal(texto):
    texto = normalizar_texto(
        texto
    ).lower()

    return "federal" in texto

def obter_campo_item(item, nomes):
    """
    O Resultado Fácil apresenta pequenas variações no JSON-LD.
    Alguns registros usam name/value e outros podem aparecer
    como nome/valor.

    Esta função aceita essas variações.
    """

    for nome in nomes:
        valor = item.get(nome)

        if valor not in [None, ""]:
            return str(valor).strip()

    return ""


def extrair_dataset_resultadofacil(html):
    """
    Localiza o Dataset de resultados dentro dos blocos JSON-LD.
    """

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    scripts = soup.find_all(
        "script",
        attrs={"type": "application/ld+json"}
    )

    for script in scripts:
        conteudo = script.string

        if not conteudo:
            conteudo = script.get_text(
                strip=True
            )

        if not conteudo:
            continue

        try:
            dados = json.loads(
                conteudo
            )
        except Exception:
            continue

        grafo = dados.get(
            "@graph",
            []
        )

        if not isinstance(
            grafo,
            list
        ):
            continue

        for item in grafo:

            if not isinstance(
                item,
                dict
            ):
                continue

            tipo = str(
                item.get(
                    "@type",
                    ""
                )
            ).lower()

            if tipo != "dataset":
                continue

            variaveis = item.get(
                "variableMeasured",
                []
            )

            if isinstance(
                variaveis,
                list
            ):
                return {
                    "dataset": item,
                    "variaveis": variaveis,
                }

    return {
        "dataset": None,
        "variaveis": [],
    }


def normalizar_nome_resultado(texto):
    texto = normalizar_texto(
        texto
    )

    texto = texto.replace(
        "º",
        "o"
    )

    texto = texto.replace(
        "ª",
        "a"
    )

    return texto


def extrair_posicao_premio(texto):
    """
    Identifica somente posições de 1º a 5º prêmio.
    """

    texto = normalizar_nome_resultado(
        texto
    ).lower()

    match = re.search(
        r"\b([1-5])\s*[oa]?\s*pr[eê]mio\b",
        texto
    )

    if not match:
        return None

    return int(
        match.group(1)
    )


def extrair_horario_variavel(texto):
    """
    Obtém o horário do nome da variável do Dataset.

    Exemplos:
    PTSP 08:20
    BA 10:20
    LOTEP 10:45
    CE 15:45
    """

    texto = normalizar_texto(
        texto
    )

    horarios = re.findall(
        r"\b([01]?\d|2[0-3]):([0-5]\d)\b",
        texto
    )

    if horarios:
        hora = horarios[-1][0]

        return hora.zfill(2)

    match = re.search(
        r"\b(\d{1,2})\s*h\b",
        texto.lower()
    )

    if match:
        return match.group(1).zfill(2)

    return ""


def pertence_loteria(
    loteria,
    nome_variavel
):
    """
    Impede que resultados da Federal ou de outra banca
    sejam incorporados à loteria analisada.
    """

    texto = normalizar_texto(
        nome_variavel
    ).upper()

    # Federal misturada nas páginas
    if "FEDERAL" in texto:
        return False

    if loteria == "PT-SP":
        return (
            "PTSP" in texto
            or "PT SP" in texto
            or "PTN SP" in texto
        )

    if loteria == "LOTEP":
        return "LOTEP" in texto

    if loteria == "PT-BA":
        return (
            " BA " in f" {texto} "
            or texto.startswith("BA ")
            or " - BA" in texto
        )

    if loteria == "LOTECE":
        return (
            "LOTECE" in texto
            or " CE," in texto
            or " CE " in texto
        )

    return False


def extrair_milhar_valor(valor):
    """
    O campo value normalmente vem assim:

    5388 · Grupo 22 · Tigre

    Captura apenas a primeira milhar.
    """

    valor = normalizar_texto(
        valor
    )

    match = re.search(
        r"(?<!\d)(\d{1,4})(?!\d)",
        valor
    )

    if not match:
        return ""

    return match.group(1).zfill(4)


def extrair_resultados_dataset(
    html,
    loteria,
    data_obj,
    url
):
    estrutura = extrair_dataset_resultadofacil(
        html
    )

    variaveis = estrutura[
        "variaveis"
    ]

    sorteios = {}

    for item in variaveis:

        if not isinstance(
            item,
            dict
        ):
            continue

        nome = obter_campo_item(
            item,
            [
                "name",
                "nome",
            ]
        )

        valor = obter_campo_item(
            item,
            [
                "value",
                "valor",
            ]
        )

        if not nome or not valor:
            continue

        if not pertence_loteria(
            loteria,
            nome
        ):
            continue

        posicao = extrair_posicao_premio(
            nome
        )

        if posicao is None:
            continue

        horario = extrair_horario_variavel(
            nome
        )

        if not horario:
            continue

        milhar = extrair_milhar_valor(
            valor
        )

        if not milhar:
            continue

        if horario not in sorteios:
            sorteios[horario] = {}

        sorteios[horario][
            posicao
        ] = milhar

    resultados = []

    for horario in sorted(
        sorteios.keys()
    ):

        premios_dict = sorteios[
            horario
        ]

        # Só aceita sorteio completo do 1º ao 5º
        if not all(
            p in premios_dict
            for p in range(1, 6)
        ):
            continue

        premios = [
            premios_dict[1],
            premios_dict[2],
            premios_dict[3],
            premios_dict[4],
            premios_dict[5],
        ]

        m6, m7 = calcular_premios_6_7(
            premios
        )

        resultados.append({
            "data": data_obj.strftime(
                "%d/%m/%Y"
            ),
            "loteria": loteria,
            "horario": horario,
            "m1": premios[0],
            "m2": premios[1],
            "m3": premios[2],
            "m4": premios[3],
            "m5": premios[4],
            "m6": m6,
            "m7": m7,
            "url": url,
        })

    return resultados

def buscar_resultados_data(
    loteria,
    data_obj
):
    url = montar_url(
        loteria,
        data_obj
    )

    try:
        resp = requests.get(
            url,
            headers=HEADERS,
            timeout=30,
        )

    except Exception as e:
        return {
            "ok": False,
            "status": "erro_rede",
            "erro": str(e),
            "url": url,
            "resultados": [],
        }

    if resp.status_code == 404:
        return {
            "ok": True,
            "status": "nao_encontrado",
            "url": url,
            "resultados": [],
        }

    if resp.status_code == 403:
        return {
            "ok": False,
            "status": "bloqueado_403",
            "url": url,
            "resultados": [],
        }

    if resp.status_code != 200:
        return {
            "ok": False,
            "status": f"http_{resp.status_code}",
            "url": url,
            "resultados": [],
        }

    resultados = extrair_resultados_dataset(
        resp.text,
        loteria,
        data_obj,
        url,
    )

    if not resultados:
        return {
            "ok": True,
            "status": "sem_resultados_validos",
            "url": url,
            "resultados": [],
        }

    return {
        "ok": True,
        "status": "ok",
        "url": url,
        "resultados": resultados,
    }

def diagnosticar_pagina_resultadofacil(loteria, data_obj):
    url = montar_url(
        loteria,
        data_obj
    )

    resp = requests.get(
        url,
        headers=HEADERS,
        timeout=30,
    )

    return {
        "status_http": resp.status_code,
        "url": url,
        "html": resp.text[:50000],
    }
    



# ==========================================================
# ROTAS DE TESTE
# ==========================================================

@app.route("/coleta/status")
def status_coleta():
    estado = carregar_estado_coleta()

    return jsonify({
        "ok": True,
        "disco": str(
            DIRETORIO_DADOS
        ),
        "estado": estado,
        "arquivos": {
            "estado": ARQUIVO_ESTADO.exists(),
            "resultados": ARQUIVO_RESULTADOS.exists(),
            "auditoria": ARQUIVO_AUDITORIA.exists(),
            "excel": ARQUIVO_EXCEL.exists(),
        },
    })

@app.route("/")
def home():
    return jsonify({
        "ok": True,
        "servico": "Coletor Histórico Resultado Fácil",
        "rotas": {
            "/teste/<loteria>/<data>": (
                "Testa uma loteria em uma data"
            ),
            "/teste-periodo": (
                "Testa período curto"
            ),
        },
    })


@app.route(
    "/teste/<loteria>/<data_teste>"
)
def teste_loteria(
    loteria,
    data_teste
):
    loteria = loteria.upper()

    if loteria not in LOTERIAS:
        return jsonify({
            "ok": False,
            "erro": "Loteria inválida",
            "loterias": list(
                LOTERIAS.keys()
            ),
        }), 400

    try:
        data_obj = datetime.strptime(
            data_teste,
            "%Y-%m-%d"
        )

    except ValueError:
        return jsonify({
            "ok": False,
            "erro": (
                "Data inválida. "
                "Use YYYY-MM-DD."
            ),
        }), 400

    retorno = buscar_resultados_data(
        loteria,
        data_obj
    )

    return jsonify(retorno)

@app.route(
    "/debug-html/<loteria>/<data_teste>"
)
def debug_html(
    loteria,
    data_teste
):
    loteria = loteria.upper()

    if loteria not in LOTERIAS:
        return jsonify({
            "ok": False,
            "erro": "Loteria inválida",
        }), 400

    try:
        data_obj = datetime.strptime(
            data_teste,
            "%Y-%m-%d"
        )

    except ValueError:
        return jsonify({
            "ok": False,
            "erro": "Data inválida. Use YYYY-MM-DD.",
        }), 400

    try:
        retorno = diagnosticar_pagina_resultadofacil(
            loteria,
            data_obj
        )

        return jsonify({
            "ok": True,
            **retorno,
        })

    except Exception as e:
        return jsonify({
            "ok": False,
            "erro": str(e),
        }), 500


@app.route("/teste-periodo")
def teste_periodo():
    data_ini = datetime(
        2026,
        8,
        10
    )

    data_fim = datetime(
        2026,
        8,
        12
    )

    saida = []

    data_atual = data_ini

    while data_atual <= data_fim:

        for loteria in LOTERIAS:

            retorno = buscar_resultados_data(
                loteria,
                data_atual
            )

            saida.append({
                "data": data_atual.strftime(
                    "%Y-%m-%d"
                ),
                "loteria": loteria,
                "status": retorno["status"],
                "quantidade": len(
                    retorno["resultados"]
                ),
                "resultados": retorno[
                    "resultados"
                ],
            })

            time.sleep(0.8)

        data_atual += timedelta(
            days=1
        )

    return jsonify({
        "ok": True,
        "total_consultas": len(saida),
        "consultas": saida,
    })


if __name__ == "__main__":
    porta = int(
        os.environ.get(
            "PORT",
            5000
        )
    )

    app.run(
        host="0.0.0.0",
        port=porta
    )
