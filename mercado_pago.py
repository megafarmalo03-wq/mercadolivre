"""
Modulo de pagamento via Mercado Pago - PIX QR Code.
Suporta tanto PIX dinamico (API) quanto estatico (fallback).
"""
import requests
import uuid
import time

# ========== CONFIGURACOES ==========
ACCESS_TOKEN = "APP_USR-1241304769992786-062223-e46e28076bcb308db533c794ab9ae1ed-3493278878"
TOKEN_CONFIGURADO = True
BASE_URL = "https://api.mercadopago.com"


def _headers():
    return {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json",
        "X-Idempotency-Key": str(uuid.uuid4()),
    }


def criar_pix(valor: float, descricao: str = "Acesso Planilha de Ganhos", referencia: str = ""):
    """Cria um pagamento PIX dinamico via Mercado Pago."""
    if not TOKEN_CONFIGURADO:
        return {"erro": "Token nao configurado"}

    url = f"{BASE_URL}/v1/payments"
    payload = {
        "transaction_amount": float(valor),
        "description": descricao,
        "payment_method_id": "pix",
        "external_reference": referencia,
        "payer": {
            "email": "cliente@email.com",
            "first_name": "Cliente",
            "last_name": "App",
            "identification": {"type": "CPF", "number": "12345678909"}
        }
    }

    try:
        resp = requests.post(url, headers=_headers(), json=payload, timeout=20)
        resp.raise_for_status()
        data = resp.json()

        point = data.get("point_of_interaction", {})
        trx = point.get("transaction_data", {})

        return {
            "id": data.get("id"),
            "status": data.get("status"),
            "qr_code": trx.get("qr_code"),
            "qr_code_base64": trx.get("qr_code_base64"),
            "ticket_url": trx.get("ticket_url"),
            "external_reference": referencia,
        }
    except requests.exceptions.HTTPError as e:
        return {"erro": f"API Error {resp.status_code}: {resp.text[:200]}", "status_code": resp.status_code}
    except Exception as e:
        return {"erro": str(e)}


def consultar_pagamento(payment_id: str):
    """Consulta status de um pagamento pelo ID."""
    if not TOKEN_CONFIGURADO or not payment_id:
        return None

    try:
        resp = requests.get(
            f"{BASE_URL}/v1/payments/{payment_id}",
            headers={"Authorization": f"Bearer {ACCESS_TOKEN}"},
            timeout=15
        )
        resp.raise_for_status()
        return resp.json().get("status")
    except Exception:
        return None


def buscar_por_referencia(external_reference: str):
    """Busca pagamentos pela external_reference."""
    if not TOKEN_CONFIGURADO:
        return []

    try:
        resp = requests.get(
            f"{BASE_URL}/v1/payments/search",
            headers={"Authorization": f"Bearer {ACCESS_TOKEN}"},
            params={
                "external_reference": external_reference,
                "sort": "date_created",
                "criteria": "desc",
                "limit": 5,
            },
            timeout=15
        )
        resp.raise_for_status()
        return resp.json().get("results", [])
    except Exception:
        return []


def ultimo_status(external_reference: str):
    """Retorna o status mais recente do pagamento."""
    pagamentos = buscar_por_referencia(external_reference)
    if not pagamentos:
        return None
    return pagamentos[0].get("status")


def verificar_token():
    """Verifica se o token esta valido."""
    if not TOKEN_CONFIGURADO:
        return False
    try:
        resp = requests.get(
            f"{BASE_URL}/users/me",
            headers={"Authorization": f"Bearer {ACCESS_TOKEN}"},
            timeout=10
        )
        return resp.status_code == 200
    except Exception:
        return False

