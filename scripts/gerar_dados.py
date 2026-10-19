"""
Le as planilhas e grava dados/dados.json, que e o arquivo consumido pelo painel.

Entrada esperada (pasta entrada/):
  produtividade.xlsx  -> aba "Produt_Montagem-Formas_R02", "BD-Obras", "BD-Custos"
  custos.xlsx         -> aba "Área Forma"   (opcional)

Se algum arquivo faltar, a parte correspondente do JSON fica vazia e o painel
continua funcionando com o que existe.
"""

import datetime as dt
import json
import unicodedata
import math
import pathlib
import re
import sys

import pandas as pd
from openpyxl import load_workbook

RAIZ = pathlib.Path(__file__).resolve().parents[1]
ENTRADA = RAIZ / "entrada"
SAIDA = RAIZ / "dados" / "dados.json"

PREFIXO_ABA = "Prod_"         # cada obra tem sua aba: Prod_Serena, Prod_Arte-Penha, ...
ABA_RESUMO = "GERAL"          # quadro consolidado por obra
ABA_FICHA = "Ficha-Técnica-Obras"   # caracteristicas de cada obra
ABA_ANTIGA = "Produt_Montagem-Formas_R02"   # formato anterior, aba unica
BUSCA_CABECALHO = 30          # ate que linha procurar o cabecalho

# Data de atualizacao mostrada no cabecalho e no rodape do painel. Quando preenchida
# ("AAAA-MM-DD"), vale esta; com None, vale a da planilha (aba GERAL, DATA ATUALIZACAO).
DATA_ATUALIZACAO_FIXA = "2026-10-19"

# Como reconhecer cada coluna do quadro GERAL. Cada entrada lista os termos que
# precisam aparecer no titulo e os que nao podem aparecer. Assim a planilha pode
# ganhar, perder ou reordenar colunas sem quebrar a leitura.
REGRAS_RESUMO = {
    "efetivo":   (["EFETIVO", "MEDIO"], ["AJUST"]),
    "efetivoAj": (["EFETIVO", "AJUST"], []),
    "m2med":     (["PRODUTIV", "MEDIA", "M2"], ["AJUST", "MAXIM"]),
    "m2aj":      (["PRODUTIV", "MEDIA", "AJUST", "M2"], ["MAXIM"]),
    "m2max":     (["PRODUTIV", "MAXIM", "M2"], []),
    "m3med":     (["PRODUTIV", "MEDIA", "M3"], ["AJUST", "MAXIM"]),
    "m3aj":      (["PRODUTIV", "MEDIA", "AJUST", "M3"], ["MAXIM"]),
    "m3max":     (["PRODUTIV", "MAXIM", "M3"], []),
    "esp":       (["ESPESSURA", "MEDIA"], ["AJUST", "MAXIM"]),
    "espaj":     (["ESPESSURA", "AJUST"], ["MAXIM"]),
    "espmax":    (["ESPESSURA", "MAXIM"], []),
    "concrMed":  (["CONCRETAGENS"], []),
}


def num(v):
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f)) else round(f, 4)
    except (TypeError, ValueError):
        return None


def data(v):
    if v is None:
        return None
    try:
        return pd.to_datetime(v).strftime("%Y-%m-%d")
    except Exception:
        return None


def limpa(txt):
    """Titulo sem acento, sem quebra de linha e em maiusculas."""
    t = unicodedata.normalize("NFD", str(txt or ""))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return " ".join(t.upper().split())


def acha_cabecalho(ws, obrigatoria="OBRA"):
    """Descobre em que linha esta o cabecalho, procurando uma coluna conhecida."""
    for i, linha in enumerate(ws.iter_rows(max_row=BUSCA_CABECALHO, values_only=True), 1):
        if any(limpa(c) == obrigatoria for c in linha):
            return i
    return None


def numera_repetidas(cab):
    """As colunas dos dias repetem o mesmo titulo seis vezes. Numera as
    repeticoes como .1, .2 ... para que cada dia possa ser lido pelo nome."""
    vistos, saida = {}, []
    for nome in cab:
        n = vistos.get(nome, 0)
        saida.append(nome if n == 0 else f"{nome}.{n}")
        vistos[nome] = n + 1
    return saida


def aba_obsoleta(nome):
    """Abas marcadas como OBSOLET(A/O) na planilha ficam de fora: guardam a
    versao antiga do mesmo dado e contariam a obra duas vezes."""
    return "OBSOLET" in limpa(nome)


def abas_de_obra(wb):
    """Abas Prod_<obra>. Se nao houver nenhuma, cai no formato antigo de aba unica."""
    abas = [s for s in wb.sheetnames
            if s.startswith(PREFIXO_ABA) and not aba_obsoleta(s)]
    if abas:
        return abas
    if ABA_ANTIGA in wb.sheetnames:
        print(f"Nenhuma aba '{PREFIXO_ABA}*'; usando '{ABA_ANTIGA}'.")
        return [ABA_ANTIGA]
    sys.exit(f"ERRO: nao encontrei abas '{PREFIXO_ABA}*' nem '{ABA_ANTIGA}'.")


def le_semanal(wb):
    """Junta as abas de obra. As colunas sao lidas pelo nome, entao a ordem
    pode variar de uma aba para outra sem quebrar nada."""
    semanas = []
    for aba in abas_de_obra(wb):
        semanas += le_aba(wb[aba], aba)
    if not semanas:
        sys.exit("ERRO: nenhuma linha de obra encontrada.")
    return semanas


def le_aba(ws, nome_aba):
    inicio = acha_cabecalho(ws)
    if not inicio:
        print(f"- {nome_aba}: cabecalho nao encontrado, ignorada.")
        return []
    linhas = list(ws.iter_rows(min_row=inicio, values_only=True))
    if not linhas:
        return []
    cab = numera_repetidas([str(c) for c in linhas[0]])
    df = pd.DataFrame(linhas[1:], columns=cab)
    if "OBRA" not in df.columns:
        print(f"- {nome_aba}: sem coluna OBRA, ignorada.")
        return []
    df = df[df["OBRA"].notna()]

    def col(base, i):
        """Nome da coluna do dia i. O pandas numera as repetidas com .1, .2...
        Quando a coluna aparece uma vez so (caso de TORRE em algumas abas),
        o mesmo valor vale para todos os dias."""
        nome = base if i == 0 else f"{base}.{i}"
        return nome if nome in df.columns else (base if base in df.columns else None)

    semanas = []
    for _, r in df.iterrows():
        detalhe = []
        for i in range(6):                       # colunas DIA 01 a DIA 06
            c_area = col("ÁREA MONTADA (m2)", i)
            area = num(r.get(c_area)) if c_area else None
            if area and area > 0:
                def v(base):
                    c = col(base, i)
                    return r.get(c) if c else None
                detalhe.append({
                    "data": data(v("DATA CONCRET.")),
                    "torre": str(v("TORRE")),
                    "pavto": str(v("PAVTO.")),
                    "trecho": str(v("TRECHO MONTADO")),
                    "area": area,
                    "vol": num(v("VOLUME CONCRETADO (m3)")),
                })
        _area = num(r.get("SOMA DA ÁREA DE FÔRMA MONTADA NA SEMANA (m2)")) or 0
        _mont = num(r.get("QTDE. MÉDIA DE MONTADORES NA SEMANA")) or 0
        # semana com área produzida mas sem efetivo lançado: falha de
        # apontamento na planilha, não produtividade zero de verdade.
        # Sem "sem" a produtividade fica sem numero, em vez de um 0 enganoso.
        _sem_efetivo = _area > 0 and _mont == 0
        # torre da semana: a aba Prod_ e de uma torre so; o detalhe diario diz qual
        _torre = detalhe[0]["torre"] if detalhe else str(r.get("TORRE") or "").strip()
        semanas.append({
            "obra": r["OBRA"],
            "torre": "" if _torre in ("-", "None", "nan") else _torre,
            "ano": int(num(r["ANO"]) or 0),
            "sem": int(num(r["SEMANA DO ANO"]) or 0),
            "periodo": r["PERÍODO"],
            "ini": data(r.get("DATA CONCRET.")),
            "mes": data(r.get("MÊS")),
            "concr": num(r.get("QTDE. CONCRETAGENS DA TORRE NA SEMANA")),
            "sabado": (str(r.get("EQUIPE TRABALHOU NO SÁBADO?")) or "").strip(),
            "feriados": num(r.get("QTDE. FERIADOS NA SEMANA")),
            "area": _area,
            "vol": num(r.get("SOMA DE VOLUME CONCRETADO NA SEMANA (m3)")) or 0,
            "mont": _mont,
            "dias": num(r.get("DIAS TRABALHADOS NA SEMANA")) or 0,
            "horas": num(r.get("HORAS TRABALHADAS NA SEMANA")) or 0,
            "pm2sem": None if _sem_efetivo else num(r.get("PRODUTIVIDADE MÉDIA SEMANAL POR MONTADOR  (m2/semana)")),
            "pm2dia": None if _sem_efetivo else num(r.get("PRODUTIVIDADE MÉDIA DIÁRIA POR MONTADOR (m2/dia)")),
            "pm2h": None if _sem_efetivo else num(r.get("PRODUTIVIDADE MÉDIA POR HORA POR MONTADOR (m2/h)")),
            "pm3dia": None if _sem_efetivo else num(r.get("PRODUTIVIDADE MÉDIA DIÁRIA POR MONTADOR (m3/dia)")),
            "pm3h": None if _sem_efetivo else num(r.get("PRODUTIVIDADE MÉDIA POR HORA POR MONTADOR (m3/h)")),
            "esp": num(r.get("ESPESSURA MÉDIA PRODUZIDA (m2/m3/dia)")),
            "rup": num(r.get("RUP - RAZÃO UNITÁRIA DE PRODUÇÃO (H.h/m²)")),
            "obs": str(r.get("OBSERVAÇÕES")) if r.get("OBSERVAÇÕES") else "",
            "detalhe": detalhe,
        })
    return semanas


def numero_torre(txt):
    """'Torre 07' -> 7 ; 'Torre 01 (unica)' -> 1 ; sem numero -> None."""
    m = re.search(r"\d+", str(txt or ""))
    return int(m.group()) if m else None


def junta_textos(valores):
    """Une textos de empreiteiros sem repetir nenhum trecho, mantendo a ordem
    em que aparecem (a ordem de entrada na obra: 'A > B' = A saiu, B entrou)."""
    saida = []
    for v in valores:
        for parte in re.split(r"\s*>\s*", str(v or "")):
            parte = parte.strip()
            if parte and parte not in saida:
                saida.append(parte)
    return " > ".join(saida)


def unifica_obra(obra, linhas):
    """Uma obra com varias linhas no GERAL (uma por torre) vira uma linha so:
    MEDIA SIMPLES das torres em cada coluna de media (cada torre pesa igual,
    como no consolidado da planilha). Colunas de maximo: o maior entre as torres."""
    if len(linhas) == 1:
        item = dict(linhas[0])
        item["torres"] = [{"torre": item.pop("torre", ""), "status": item.pop("status", ""),
                           "empreiteiros": item.pop("empreiteiros", "")}]
        item["status"] = item["torres"][0]["status"]
        item["empreiteiros"] = item["torres"][0]["empreiteiros"]
        return item

    item = {"obra": obra}
    for chave in REGRAS_RESUMO:
        vals = [l.get(chave) for l in linhas if l.get(chave) is not None]
        if not vals:
            item[chave] = None
        elif chave.endswith("max"):
            item[chave] = max(vals)
        else:
            item[chave] = round(sum(vals) / len(vals), 4)
    status = [str(l.get("status") or "").strip() for l in linhas]
    # a obra so esta concluida quando todas as torres estao
    item["status"] = ("Concluído" if all(limpa(x).startswith("CONCLU") for x in status if x)
                      and any(status) else "Em andamento")
    item["empreiteiros"] = junta_textos(l.get("empreiteiros") for l in linhas)
    item["torres"] = [{"torre": str(l.get("torre") or ""), "status": str(l.get("status") or ""),
                       "empreiteiros": str(l.get("empreiteiros") or "")} for l in linhas]
    return item


def le_resumo(wb):
    """Quadro consolidado por obra, na aba GERAL. As colunas sao reconhecidas
    pelo titulo, entao mudar a ordem ou inserir colunas nao quebra a leitura.
    Se a planilha traz uma linha por torre, elas sao unificadas por obra."""
    if ABA_RESUMO not in wb.sheetnames:
        return []
    ws = wb[ABA_RESUMO]
    inicio = acha_cabecalho(ws)
    if not inicio:
        print(f"- aba {ABA_RESUMO}: cabecalho nao encontrado.")
        return []

    linhas = list(ws.iter_rows(min_row=inicio, values_only=True))
    cab = [limpa(c) for c in linhas[0]]
    col_obra = cab.index("OBRA")
    col_torre = next((i for i, t in enumerate(cab) if t == "TORRE"), None)
    col_status = next((i for i, t in enumerate(cab) if t.startswith("STATUS")), None)
    col_emp = next((i for i, t in enumerate(cab) if "EMPREITEIRO" in t), None)

    mapa = {}
    for chave, (precisa, proibe) in REGRAS_RESUMO.items():
        for i, titulo in enumerate(cab):
            if not titulo or i == col_obra:
                continue
            if all(t in titulo for t in precisa) and not any(t in titulo for t in proibe):
                mapa[chave] = i
                break
        else:
            print(f"- aba {ABA_RESUMO}: coluna '{chave}' nao encontrada.")

    def cel(r, i):
        return r[i] if i is not None and i < len(r) else None

    por_obra = {}
    for r in linhas[1:]:
        obra = r[col_obra] if col_obra < len(r) else None
        if not obra or not isinstance(obra, str):
            continue
        item = {"obra": obra.strip(),
                "torre": cel(r, col_torre), "status": cel(r, col_status),
                "empreiteiros": cel(r, col_emp)}
        for chave, i in mapa.items():
            item[chave] = num(r[i]) if i < len(r) else None
        por_obra.setdefault(item["obra"], []).append(item)

    saida = []
    for obra, ls in por_obra.items():
        if len(ls) > 1:
            print(f"GERAL: {obra} tem {len(ls)} linhas (torres) -> unificada em uma.")
        saida.append(unifica_obra(obra, ls))
    return saida


def acha_atualizacao(wb):
    """Procura o rotulo DATA ATUALIZACAO e devolve o valor da celula ao lado."""
    for aba in (ABA_RESUMO, ABA_ANTIGA):
        if aba not in wb.sheetnames:
            continue
        for linha in wb[aba].iter_rows(max_row=BUSCA_CABECALHO, values_only=True):
            for i, c in enumerate(linha):
                if "DATA ATUALIZ" in limpa(c):
                    for v in linha[i + 1:i + 4]:
                        d = data(v)
                        if d:
                            return d
    return None


def le_ficha(wb):
    """Ficha tecnica das obras. O cabecalho e localizado pela coluna OBRA."""
    if ABA_FICHA not in wb.sheetnames:
        return {}
    ws = wb[ABA_FICHA]
    inicio = acha_cabecalho(ws)
    if not inicio:
        print(f"- aba {ABA_FICHA}: cabecalho nao encontrado.")
        return {}

    linhas = list(ws.iter_rows(min_row=inicio, values_only=True))
    cab = linhas[0]
    # colunas validas: as que tem titulo
    idx = [i for i, t in enumerate(cab) if t not in (None, "")]
    i_obra = next(i for i in idx if limpa(cab[i]) == "OBRA")
    campos = [" ".join(str(cab[i]).split()) for i in idx if i != i_obra]

    def texto(v):
        if v is None:
            return ""
        if isinstance(v, dt.datetime):
            return v.strftime("%d/%m/%Y")
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return "\n".join(l.strip() for l in str(v).split("\n") if l.strip())

    obras = []
    for r in linhas[1:]:
        nome = r[i_obra] if i_obra < len(r) else None
        if not nome or not isinstance(nome, str):
            continue
        obras.append({"obra": nome.strip(),
                      "valores": [texto(r[i]) if i < len(r) else "" for i in idx if i != i_obra]})
    return {"campos": campos, "obras": obras}


def pavto_valido(pav):
    """So conta pavimento terreo, tipo (numerado), platibanda ou atico —
    os unicos estagios que compoem o ciclo estrutural da torre."""
    t = limpa(pav)
    if not t:
        return False
    if t in ("TERREO", "PLATIBANDA", "ATICO"):
        return True
    return bool(re.match(r"^\d+[°º]?\s*PAVTO\.?$", t))


def le_ciclo(wb):
    """Ciclo de concretagem: dias corridos entre uma concretagem e a anterior
    da MESMA torre. Calculado aqui, e não lido da coluna RITMO da planilha,
    porque essa coluna compara com a linha anterior da aba — que pode ser de
    outra torre, quando duas torres têm concretagens perto uma da outra no
    calendário e suas linhas ficam intercaladas. A primeira concretagem de
    cada torre não tem uma anterior para comparar e fica de fora."""
    brutos = []
    for aba in [s for s in wb.sheetnames if s.startswith("C.Geo") and not aba_obsoleta(s)]:
        ws = wb[aba]
        obra = ws.cell(2, 1).value
        if not obra or not isinstance(obra, str):
            print(f"- {aba}: sem nome de obra na célula A2, ignorada.")
            continue
        obra = obra.strip()

        linha_cab = acha_cabecalho(ws, obrigatoria="TORRE")
        if not linha_cab:
            print(f"- {aba}: cabeçalho com 'Torre' não encontrado, ignorada.")
            continue

        col = {}
        for c in range(1, ws.max_column + 1):
            titulo = str(ws.cell(linha_cab, c).value or "").strip()
            if titulo in ("Data", "Trecho", "Torre", "Pavto."):
                col[titulo] = c
        faltando = {"Data", "Torre"} - col.keys()
        if faltando:
            print(f"- {aba}: colunas {faltando} não encontradas, ignorada.")
            continue

        # ws.cell(r,c) é lentíssimo em modo read_only para milhares de linhas;
        # iter_rows varre a planilha sequencialmente, na velocidade certa.
        maxc = max(col.values())
        n = 0
        for linha in ws.iter_rows(min_row=linha_cab + 1, max_col=maxc, values_only=True):
            torre = linha[col["Torre"] - 1]
            data_c = linha[col["Data"] - 1]
            if not torre or not isinstance(data_c, dt.datetime):
                continue
            pav = linha[col["Pavto."] - 1] if "Pavto." in col else None
            if not pavto_valido(pav):
                continue
            tre = linha[col["Trecho"] - 1] if "Trecho" in col else None
            brutos.append({
                "obra": obra,
                "data": data_c.date(),
                "torre": str(torre).strip(),
                "pavto": str(pav or "").strip(),
                "trecho": str(tre or "").strip(),
            })
            n += 1
        print(f"{aba:<28} {n} linhas lidas ({obra})")

    brutos.sort(key=lambda e: (e["obra"], e["torre"], e["data"]))
    eventos = []
    anterior = {}
    for e in brutos:
        chave = (e["obra"], e["torre"])
        prev = anterior.get(chave)
        anterior[chave] = e["data"]
        if prev is None:
            continue
        eventos.append({
            "obra": e["obra"],
            "data": e["data"].strftime("%Y-%m-%d"),
            "ritmo": (e["data"] - prev).days,
            "torre": e["torre"],
            "pavto": e["pavto"],
            "trecho": e["trecho"],
        })
    eventos.sort(key=lambda e: (e["obra"], e["data"]))
    return eventos


def le_bd_obras(wb):
    if "BD-Obras" not in wb.sheetnames:
        return []
    return [
        {"obra": r[0], "trecho": str(r[1]), "desc": r[2],
         "area": num(r[3]), "vol": num(r[4])}
        for r in wb["BD-Obras"].iter_rows(min_row=4, values_only=True) if r[0]
    ]


def le_bd_custos(wb):
    if "BD-Custos" not in wb.sheetnames:
        return [], []
    ws = wb["BD-Custos"]
    cargos = [
        {"cargo": r[0], "salario": num(r[1]), "comEncargos": num(r[2])}
        for r in ws.iter_rows(min_row=8, max_row=10, values_only=True) if r[0]
    ]
    equip = [
        {"tipo": r[0], "mensal": num(r[1])}
        for r in ws.iter_rows(min_row=18, max_row=19, values_only=True) if r[0]
    ]
    return cargos, equip


def le_area_forma(caminho):
    """Aba 'Área Forma' da planilha de mao de obra. Opcional."""
    if not caminho.exists():
        return []
    wb = load_workbook(caminho, read_only=True, data_only=True)
    aba = next((s for s in wb.sheetnames if s.strip().lower().startswith("área forma")
                or s.strip().lower().startswith("area forma")), None)
    if not aba:
        return []
    campos = ["obra", "quartos", "tipologia", "aptos", "areaFormaSLaje", "areaLaje",
              "indice", "custoParedeMO", "areaMontagem", "rsPorM2", "rsPorApto",
              "custoInst", "instConc60", "instPorM2", "totalPorM2"]
    saida = []
    for r in wb[aba].iter_rows(min_row=11, max_row=40, values_only=True):
        if not r[1] or not isinstance(r[1], str):
            continue
        item = {"obra": r[1], "quartos": r[2], "tipologia": r[3]}
        for i, campo in enumerate(campos[3:], start=4):
            item[campo] = num(r[i])
        if item["areaMontagem"] and item["rsPorM2"]:   # descarta linhas de nota
            saida.append(item)
    return saida


def main():
    prod = ENTRADA / "produtividade.xlsx"
    if not prod.exists():
        sys.exit(f"ERRO: {prod} nao encontrado. Rode baixar.py antes.")

    wb = load_workbook(prod, read_only=True, data_only=True)

    atualizacao = DATA_ATUALIZACAO_FIXA or acha_atualizacao(wb)

    cargos, equip = le_bd_custos(wb)
    semanas = le_semanal(wb)
    saida = {
        "geradoEm": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "atualizacaoPlanilha": atualizacao,
        "semanas": semanas,
        "resumo": le_resumo(wb),
        "trechos": le_bd_obras(wb),
        "cargos": cargos,
        "equip": equip,
        "areaForma": le_area_forma(ENTRADA / "custos.xlsx"),
        "ficha": le_ficha(wb),
        "ciclo": le_ciclo(wb),
    }

    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    SAIDA.write_text(json.dumps(saida, ensure_ascii=False, separators=(",", ":")),
                     encoding="utf-8")

    com_producao = sum(1 for s in saida["semanas"] if s["area"] > 0)
    print(f"dados/dados.json gerado: {SAIDA.stat().st_size:,} bytes")
    print(f"  {len(saida['semanas'])} semanas ({com_producao} com producao)")
    for o in sorted(set(s["obra"] for s in saida["semanas"])):
        n = sum(1 for s in saida["semanas"] if s["obra"] == o and s["area"] > 0)
        print(f"    {o}: {n} semanas com producao")
    print(f"  {len(saida['areaForma'])} linhas de orcamento")
    f = saida['ficha']
    if f:
        print(f"  ficha tecnica: {len(f['obras'])} obras x {len(f['campos'])} campos")
    print(f"  ciclo de concretagem: {len(saida['ciclo'])} eventos")


if __name__ == "__main__":
    main()
