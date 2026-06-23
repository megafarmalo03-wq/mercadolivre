"""
Modulo de pagamento via Mercado Pago - PIX QR Code dinamico.
O dinheiro cai direto na conta Mercado Pago do usuario.
Documentacao: https://www.mercadopago.com.br/developers/pt/reference
"""
import requests
import uuid

# ========== CONFIGURACOES ==========
ACCESS_TOKEN = "APP_USR-1241304769992786-062223-e46e28076bcb308db533c794ab9ae1ed-3493278878"
TOKEN_CONFIGURADO = True
BASE_URL = "https://api.mercadopago.com"


def criar_pix(valor: float, descricao: str = "Acesso Planilha de Ganhos", referencia: str = ""):
    """
    Cria um pagamento PIX dinamico via Mercado Pago.
    Retorna dict com:
        - id (id do pagamento MP)
        - qr_code (string base64 da imagem QR)
        - qr_code_base64 (imagem QR em base64 para exibir)
        - ticket_url (link para pagamento no celular)
        - status
    """
    if not TOKEN_CONFIGURADO:
        return {"erro": "Token nao configurado"}

    idempotency = str(uuid.uuid4())
    url = f"{BASE_URL}/v1/payments"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json",
        "X-Idempotency-Key": idempotency,
    }
    payload = {
        "transaction_amount": float(valor),
        "description": descricao,
        "payment_method_id": "pix",
        "external_reference": referencia,
        "notification_url": "https://www.google.com",  # webhook (opcional)
        "payer": {
            "email": "cliente@email.com",
            "first_name": "Cliente",
            "last_name": "App",
            "identification": {
                "type": "CPF",
                "number": "12345678909"
            }
        }
    }

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=20)
        resp.raise_for_status()
        data = resp.json()

        point = data.get("point_of_interaction", {})
        trx = point.get("transaction_data", {})

        return {
            "id": data.get("id"),
            "status": data.get("status"),
            "qr_code": trx.get("qr_code"),                 # copia e cola
            "qr_code_base64": trx.get("qr_code_base64"),   # imagem base64
            "ticket_url": trx.get("ticket_url"),           # link para pagar no celular
            "external_reference": referencia,
        }
    except requests.exceptions.HTTPError as e:
        return {"erro": f"Erro API Mercado Pago: {e}", "detalhes": resp.text if 'resp' in dir() else str(e)}
    except Exception as e:
        return {"erro": str(e)}


def consultar_pagamento(payment_id: str):
    """
    Consulta o status de um pagamento pelo ID.
    Retorna: approved | pending | in_process | rejected | cancelled | None
    """
    if not TOKEN_CONFIGURADO:
        return None

    url = f"{BASE_URL}/v1/payments/{payment_id}"
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}

    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        return resp.json().get("status")
    except Exception:
        return None


def buscar_por_referencia(external_reference: str):
    """
    Busca pagamentos pela external_reference.
    Retorna lista de dicts.
    """
    if not TOKEN_CONFIGURADO:
        return []

    url = f"{BASE_URL}/v1/payments/search"
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}
    params = {
        "external_reference": external_reference,
        "sort": "date_created",
        "criteria": "desc",
        "limit": 5,
    }

    try:
        resp = requests.get(url, headers=headers, params=params, timeout=15)
        resp.raise_for_status()
        return resp.json().get("results", [])
    except Exception:
        return []


def ultimo_status(external_reference: str):
    """
    Retorna o status mais recente do pagamento para a referencia.
    """
    pagamentos = buscar_por_referencia(external_reference)
    if not pagamentos:
        return None
    return pagamentos[0].get("status")
