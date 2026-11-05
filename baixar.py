"""
Baixa as planilhas de origem do Google Sheets ou do OneDrive / SharePoint.

Dois modos, escolhidos automaticamente:

1. LINK PUBLICO — mais simples. Serve para Google Sheets (compartilhado como
   "Qualquer pessoa com o link") e para OneDrive/SharePoint com link anonimo.
   Nao exige cadastrar aplicativo nenhum. No caso da Microsoft, muitas empresas
   bloqueiam compartilhamento anonimo por politica; ai use o modo 2.

2. MICROSOFT GRAPH — usado quando TENANT_ID / CLIENT_ID / CLIENT_SECRET
   estao definidos. Funciona com link interno ("Pessoas da sua organizacao"),
   que e o caso da maioria dos tenants corporativos.

Saida: arquivos gravados em entrada/.
"""

import base64
import os
import re
import pathlib
import sys

import requests

ENTRADA = pathlib.Path(__file__).resolve().parents[1] / "entrada"
ENTRADA.mkdir(exist_ok=True)

ARQUIVOS = [
    ("LINK_PRODUTIVIDADE", "produtividade.xlsx", True),   # obrigatorio
    ("LINK_CUSTOS", "custos.xlsx", False),                # opcional
]


def token_graph():
    """Token app-only. Retorna None se as credenciais nao estiverem configuradas."""
    tenant = os.environ.get("TENANT_ID")
    cid = os.environ.get("CLIENT_ID")
    seg = os.environ.get("CLIENT_SECRET")
    if not (tenant and cid and seg):
        return None
    r = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "client_id": cid,
            "client_secret": seg,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=60,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def id_compartilhamento(url):
    """Converte a URL de compartilhamento no share id que o Graph espera."""
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
    return "u!" + b64.rstrip("=")


def baixa_via_graph(url, destino, token):
    api = (
        "https://graph.microsoft.com/v1.0/shares/"
        f"{id_compartilhamento(url)}/driveItem/content"
    )
    r = requests.get(api, headers={"Authorization": f"Bearer {token}"},
                     allow_redirects=True, timeout=300)
    r.raise_for_status()
    destino.write_bytes(r.content)


def url_google(url):
    """Se for um link do Google Sheets, devolve a URL que exporta a planilha em xlsx."""
    if "docs.google.com/spreadsheets" not in url:
        return None
    # link "publicar na web" usa /d/e/<token>, que tem endpoint proprio
    m = re.search(r"/spreadsheets/d/e/([a-zA-Z0-9-_]+)", url)
    if m:
        return f"https://docs.google.com/spreadsheets/d/e/{m.group(1)}/pub?output=xlsx"
    # link normal de compartilhamento
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if m:
        return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx"
    return None


def baixa_via_link(url, destino):
    """Link publico. Serve para Google Sheets e para links anonimos do OneDrive."""
    google = url_google(url)
    if google:
        alvo = google
        dica = ("o link do Google Sheets nao esta publico. Em Compartilhar, "
                "marque 'Qualquer pessoa com o link' como Leitor.")
    else:
        sep = "&" if "?" in url else "?"
        alvo = url + sep + "download=1"
        dica = ("o link provavelmente exige login. Deixe o arquivo como "
                "'Qualquer pessoa com o link' ou use o modo Microsoft Graph.")

    r = requests.get(alvo, allow_redirects=True, timeout=300)
    r.raise_for_status()
    if r.content[:2] != b"PK":  # xlsx e um zip; comeca com PK
        raise RuntimeError("a resposta nao é um arquivo xlsx — " + dica)
    destino.write_bytes(r.content)


def main():
    token = token_graph()
    modo = "Microsoft Graph" if token else "link anonimo"
    print(f"Modo de download: {modo}")

    baixados = 0
    for env, nome, obrigatorio in ARQUIVOS:
        url = os.environ.get(env, "").strip()
        if not url:
            if obrigatorio:
                sys.exit(f"ERRO: o secret {env} nao foi definido.")
            print(f"- {env} nao definido, seguindo sem esse arquivo.")
            continue

        destino = ENTRADA / nome
        try:
            if token and not url_google(url):
                baixa_via_graph(url, destino, token)
            else:
                baixa_via_link(url, destino)
            print(f"- {nome}: {destino.stat().st_size:,} bytes")
            baixados += 1
        except Exception as erro:
            if obrigatorio:
                sys.exit(f"ERRO ao baixar {nome}: {erro}")
            print(f"- {nome}: falhou ({erro}). Seguindo sem esse arquivo.")

    if not baixados:
        sys.exit("ERRO: nenhum arquivo baixado.")


if __name__ == "__main__":
    main()
