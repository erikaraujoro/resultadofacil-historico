import os
import re
import time
import logging
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

LOTTERIAS = {
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
    cfg = LOTerias[loteria]

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


def extrair_premios_bloco(bloco):
    texto = normalizar_texto(
        bloco.get_text(
            " ",
            strip=True
        )
    )

    numeros = re.findall(
        r"\b\d{1,4}\b",
        texto
    )

    milhares = []

    for numero in numeros:
        milhar = formatar_milhar(
            numero
        )

        if len(milhar) == 4:
            milhares.append(milhar)

    # Remove repetições consecutivas simples
    filtradas = []

    for milhar in milhares:
        if (
            not filtradas
            or filtradas[-1] != milhar
        ):
            filtradas.append(milhar)

    if len(filtradas) < 5:
        return []

    return filtradas[:5]


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

    soup = BeautifulSoup(
        resp.text,
        "html.parser"
    )

    resultados = []

    # Procura blocos que possuam textos de horário
    elementos = soup.find_all(
        ["div", "section", "article", "table"]
    )

    chaves = set()

    for elemento in elementos:
        texto = normalizar_texto(
            elemento.get_text(
                " ",
                strip=True
            )
        )

        if not texto:
            continue

        horario = extrair_horario(
            texto
        )

        if not horario:
            continue

        if parece_federal(texto):
            continue

        premios = extrair_premios_bloco(
            elemento
        )

        if len(premios) != 5:
            continue

        m6, m7 = calcular_premios_6_7(
            premios
        )

        data_br = data_obj.strftime(
            "%d/%m/%Y"
        )

        chave = (
            data_br,
            loteria,
            horario,
            *premios,
        )

        if chave in chaves:
            continue

        chaves.add(chave)

        resultados.append({
            "data": data_br,
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

    resultados.sort(
        key=lambda x: x["horario"]
    )

    return {
        "ok": True,
        "status": "ok",
        "url": url,
        "resultados": resultados,
    }


# ==========================================================
# ROTAS DE TESTE
# ==========================================================

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

    if loteria not in LOTerias:
        return jsonify({
            "ok": False,
            "erro": "Loteria inválida",
            "loterias": list(
                LOTerias.keys()
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

        for loteria in LOTerias:

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
