"""
Modulo de pagamento PIX via OpenPix.
Documentacao: https://developers.openpix.com.br/
"""
import requests
import uuid
import time

# ========== CONFIGURACOES ==========
OPENPIX_TOKEN = "COLOQUE_SEU_TOKEN_AQUI"  # <-- SUBSTITUA PELO SEU TOKEN DA OPENPIX
OPENPIX_BASE = "https://api.openpix.com.br/api/openpix/v1"

_HEADERS = {
    "Authorization": OPENPIX_TOKEN,
    "Content-Type": "application/json",
}


def criar_cobranca(valor_reais: float, descricao: str = "Acesso Planilha de Ganhos", identificador: str = None):
    """
    Cria uma cobranca PIX dinamica na OpenPix.
    Retorna dict com: id, brCode, qrCodeImage (base64), correlationID, status
    """
    if OPENPIX_TOKEN == "COLOQUE_SEU_TOKEN_AQUI":
        return {"erro": "Token OpenPix nao configurado. Edite pagamento.py e insira seu token."}

    identificador = identificador or str(uuid.uuid4())
    valor_centavos = int(round(valor_reais * 100))

    payload = {
        "correlationID": identificador,
        "value": valor_centavos,
        "comment": descricao,
    }

    try:
        resp = requests.post(
            f"{OPENPIX_BASE}/charge",
            json=payload,
            headers=_HEADERS,
            timeout=20,
        )
        resp.raise_for_status()
        data = resp.json().get("charge", {})
        return {
            "id": data.get("correlationID"),
            "brCode": data.get("brCode"),
            "qrCodeImage": data.get("qrCodeImage"),  # base64 PNG
            "status": data.get("status"),
            "valor": valor_reais,
        }
    except requests.exceptions.HTTPError as e:
        return {"erro": f"Erro na API OpenPix: {e}", "detalhes": resp.text if 'resp' in dir() else str(e)}
    except Exception as e:
        return {"erro": f"Falha ao criar cobranca: {e}"}


def consultar_cobranca(identificador: str):
    """
    Consulta o status de uma cobranca pelo correlationID.
    Retorna: COMPLETED | ACTIVE | EXPIRED | etc
    """
    if OPENPIX_TOKEN == "COLOQUE_SEU_TOKEN_AQUI":
        return None

    try:
        resp = requests.get(
            f"{OPENPIX_BASE}/charge/{identificador}",
            headers=_HEADERS,
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json().get("charge", {})
        return data.get("status")
    except Exception:
        return None


def aguardar_pagamento(identificador: str, timeout_segundos: int = 300, intervalo: int = 5, callback=None):
    """
    Faz polling ate o pagamento ser confirmado ou timeout.
    callback(status) eh chamado a cada verificacao.
    Retorna True se pago, False se timeout.
    """
    inicio = time.time()
    while time.time() - inicio < timeout_segundos:
        status = consultar_cobranca(identificador)
        if callback:
            callback(status)
        if status == "COMPLETED":
            return True
        time.sleep(intervalo)
    return False
