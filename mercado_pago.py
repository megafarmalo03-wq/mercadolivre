"""
Modulo de pagamento via Mercado Pago (guest mode - sem conta PJ).
Para PIX estatico, gera um link de checkout para pagamento.
"""
import requests

# ========== CONFIGURACOES ==========
# Para usar, crie uma conta em https://www.mercadopago.com.br/developers
# Va em Users and Permissions > Credentials > Access Token
# Troque o texto abaixo pelo seu token de teste ou producao.
# Sem token, o app mostra apenas QR Code estatico + codigo de libercao manual.
ACCESS_TOKEN = "APP_USR-1241304769992786-062223-e46e28076bcb308db533c794ab9ae1ed-3493278878"
BASE_URL = "https://api.mercadopago.com"


def criar_preferencia(valor: float, descricao: str = "Acesso Planilha de Ganhos", referencia: str = ""):
    """
    Cria uma preferencia de pagamento (link de checkout Mercado Pago).
    Retorna dict com: init_point (URL de pagamento), id
    """
    if ACCESS_TOKEN == "APP_USR-1241304769992786-062223-e46e28076bcb308db533c794ab9ae1ed-3493278878":
        return {"erro": "Token nao configurado"}

    url = f"{BASE_URL}/checkout/preferences"
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}", "Content-Type": "application/json"}
    payload = {
        "items": [{
            "title": descricao,
            "quantity": 1,
            "currency_id": "BRL",
            "unit_price": float(valor),
        }],
        "external_reference": referencia,
        "back_urls": {
            "success": "https://www.google.com",
            "failure": "https://www.google.com",
            "pending": "https://www.google.com",
        },
    }

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        return {
            "id": data.get("id"),
            "init_point": data.get("init_point"),     # link de pagamento web
            "sandbox_init_point": data.get("sandbox_init_point"),
        }
    except Exception as e:
        return {"erro": str(e)}


def buscar_pagamentos(external_reference: str):
    """
    Busca pagamentos pela referencia externa.
    Retorna list de dicts com status.
    """
    if ACCESS_TOKEN == "APP_USR-1241304769992786-062223-e46e28076bcb308db533c794ab9ae1ed-3493278878":
        return []

    # Mercado Pago nao tem endpoint direto por external_reference na v1 publica.
    # Alternativa: usar Webhook (requer servidor publico) ou polling via API de busca.
    # Aqui fazemos polling pela API de payments com filtro por external_reference.
    url = f"{BASE_URL}/v1/payments/search"
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}
    params = {"external_reference": external_reference, "sort": "date_created", "criteria": "desc", "limit": 5}

    try:
        resp = requests.get(url, headers=headers, params=params, timeout=15)
        resp.raise_for_status()
        return resp.json().get("results", [])
    except Exception:
        return []


def ultimo_status(external_reference: str):
    """
    Retorna o status mais recente do pagamento para a referencia.
    Estados comuns: approved, pending, in_process, rejected, cancelled
    """
    pagamentos = buscar_pagamentos(external_reference)
    if not pagamentos:
        return None
    return pagamentos[0].get("status")
