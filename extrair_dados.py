r"""
Extrai os dados do painel "Departamentos" a partir de RESULTADO.xlsx
(aba "MES ATUAL") e salva um dados.json pronto pro gerador de HTML
consumir.

Layout da aba: blocos por supervisor (linha com nome do supervisor na
col C, ex: "EDMAR"), seguidos de uma linha de cabeçalho ("RCA" / "VENDEDOR")
e 1 linha por RCA até a linha "TOTAL" que fecha o bloco.

Por RCA (só as 10 categorias de produto — Mix Médio e SKU ficam de fora):
- C = código, D = nome completo (formato "NOME - ROTA...", às vezes só "NOME")
- N = BACON realizado             | meta fixa na linha 5, col N
- P = BOVINO realizado            | meta fixa na linha 5, col P
- R = BATATA realizado            | meta fixa na linha 5, col R
- T = SUINO realizado             | meta fixa na linha 5, col T
- V = CALABRESA realizado         | meta fixa na linha 5, col V
- X = PÃES realizado              | meta fixa na linha 5, col X
- Z = FRESCAIS realizado          | meta fixa na linha 5, col Z
- AB = SABORIZADAS realizado      | meta não tem célula numérica própria —
- AD = LACTEOS realizado          |   vem do texto "MINIMO N" no cabeçalho
- AM = THERMO realizado           |   de cada bloco (col AB/AD/AM da linha do supervisor)

A linha 5 (mínimos "oficiais") é global — vale pra todos os blocos/supervisores,
mesmo que o texto do cabeçalho de algum bloco mostre um número desatualizado
(ex: bloco do RICHARD mostra "MINIMO 2,15" no texto pra Mix Médio, que nem
entra mais aqui — mas confirma que o texto do cabeçalho pode ficar desatualizado).

"Bateu a meta" (status do card) = realizado >= meta em TODAS as 10 categorias
de produto.
"""

import json
import os
import re

import openpyxl

CAMINHO_RESULTADO = r"C:\Users\edmar\Desktop\ACOMPANHA RESULTADO\RESULTADO.xlsx"
PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
# Caminhos relativos: a pasta do projeto mudou de C:\AutomacaoMaxGestao para
# C:\Drivers\AutomacaoMaxGestao e o caminho fixo antigo deixou de existir.
CAMINHO_MELHORIA_SALARIAL = os.path.join(PASTA_BASE, "..", "melhoria_salarial", "dados.json")
CAMINHO_PAINEL_PILARES = os.path.join(PASTA_BASE, "..", "painel_pilares", "dados.json")

CAMINHO_SAIDA = os.path.join(PASTA_BASE, "dados.json")

# Nome do supervisor na planilha -> nome de exibição/agrupamento no painel.
NORMALIZAR_SUPERVISOR = {
    "EDMAR": "EDMAR",
    "LEANDRO FREITAS": "LEANDRO",
    "FLAVIANE": "FLAVIANE",
    "IDEGLAN": "IDEGLAN",
    "RICARDO": "RICARDO",
    "SUP RICHARD": "RICHARD",
    "RODRIGO": "RODRIGO",
}

# (chave, rótulo de exibição, coluna padrão do realizado, início do texto do cabeçalho)
# Mix Médio e SKU ficam de fora — o painel mostra só as categorias de produto.
# AJUSTE (08/10): a planilha perdeu uma coluna e tudo andou 1 pra esquerda;
# agora a coluna de cada categoria é achada pelo texto da linha de cabeçalho
# ("RCA" / "VENDEDOR" / ... "BACON" ...). A coluna padrão só vale se o
# cabeçalho não for encontrado.
CATEGORIAS = [
    ("bacon", "Bacon", 13, "BACON"),            # M
    ("bovino", "Bovino", 15, "BOVINO"),         # O
    ("batata", "Batata", 17, "BATATA"),         # Q
    ("suino", "Suíno", 19, "SUINO"),            # S
    ("calabresa", "Calabresa", 21, "CALABRESA"),  # U
    ("paes", "Pães", 23, "PÃES"),               # W
    ("frescais", "Frescais", 25, "FRESCAIS"),   # Y
    ("lacteos", "Lácteos", 29, "LACTEOS"),      # AC
    ("thermo", "Thermo", 38, "THERMO"),         # AL
]
COL_MEDIA_PEDIDOS = 7  # G — "MEDIA"/"REAL": média de pedidos/dia do RCA no mês
CABECALHO_MEDIA = ("MEDIA", "REAL")

# "saborizadas" (col AB) retirada do painel a pedido do Edmar (25/08) —
# fica de fora da contagem "bateu a meta" também.

# Até 19/08 só saborizadas/AB, lacteos/AD e thermo/AM tinham o texto
# "MINIMO N" no cabeçalho de cada bloco — as outras 7 categorias usavam um
# valor numérico fixo da linha 5 (mesmo pra todo mundo). Confirmado em 21/08
# que agora TODO bloco de supervisor tem seu próprio "MINIMO N" pras 10
# categorias (e alguns minimos mudaram, ex: Bovino é 15 e não mais 10) —
# a linha 5 ficou obsoleta, a meta de cada categoria sempre vem do texto do
# cabeçalho do bloco correspondente.


def _num(v):
    return v if isinstance(v, (int, float)) else 0


def _parse_minimo(texto):
    """'MINIMO 15' -> 15.0 · 'MINIMO 2,50' -> 2.5 · None/outro -> 0.0"""
    if not isinstance(texto, str):
        return 0.0
    m = re.search(r"[\d.,]+", texto)
    if not m:
        return 0.0
    return float(m.group(0).replace(",", "."))


def _nome_e_rota(nome_completo):
    """'FABIO L. - GYN RT 21 - SEG 01' -> ('FABIO L.', 'GYN RT 21 - SEG 01')
    Nomes sem ' - ' (com espaços dos dois lados) ficam só com o nome, sem rota."""
    partes = [p.strip() for p in nome_completo.split(" - ")]
    nome = partes[0].strip().rstrip("-").strip()  # "LUIZ GUSTAVO -" -> "LUIZ GUSTAVO"
    rota = " - ".join(p for p in partes[1:] if p).strip()
    return nome, rota


def _ler_meta_pedidos_dia():
    """Cross-referencia com o painel Performance (melhoria_salarial) pra
    pegar a Meta de Pedidos/Dia — é um valor ÚNICO (não por RCA), o mesmo
    "desafio" pra equipe toda; quando o Edmar edita esse campo no
    Performance, o valor vivo (do navegador) sobrepõe este aqui via
    localStorage compartilhado (ver JS de painel_departamentos)."""
    if not os.path.exists(CAMINHO_MELHORIA_SALARIAL):
        return 0
    with open(CAMINHO_MELHORIA_SALARIAL, "r", encoding="utf-8") as f:
        dados = json.load(f)
    return round(dados["constantes"]["meta_pedidos_dia"])


def _ler_posit_atual():
    """Cross-referencia com o Painel 4 Pilares pra pegar a 'Média de
    pedidos' (coluna BD da SOMA NAO SALVA ENCIMA.xlsx) de cada RCA — é uma
    métrica diferente da 'média pedidos mês atual' (essa vem do total de
    pedidos ÷ dias úteis calculado no Performance)."""
    if not os.path.exists(CAMINHO_PAINEL_PILARES):
        return {}
    with open(CAMINHO_PAINEL_PILARES, "r", encoding="utf-8") as f:
        dados = json.load(f)
    return {str(r["codigo"]): r["media_pedidos"] for r in dados}


def extrair():
    wb = openpyxl.load_workbook(CAMINHO_RESULTADO, data_only=True)
    ws = wb["MES ATUAL"]

    def val(row, col):
        return _num(ws.cell(row=row, column=col).value)

    meta_pedidos_dia = _ler_meta_pedidos_dia()
    posit_atual = _ler_posit_atual()

    rcas = []
    supervisor_atual = None
    linha_supervisor = None
    metas_bloco = {}
    cols = {chave: col for chave, _, col, _ in CATEGORIAS}
    col_media = COL_MEDIA_PEDIDOS

    for r in range(1, ws.max_row + 1):
        c3 = ws.cell(row=r, column=3).value
        c4 = ws.cell(row=r, column=4).value

        if isinstance(c3, str) and c3.strip() == "RCA":
            # Linha "RCA / VENDEDOR / ... BACON ..." — localiza as colunas pelo texto
            # e lê os "MINIMO N" da linha do supervisor logo acima, nas mesmas colunas.
            textos = {c: str(ws.cell(row=r, column=c).value or "").strip().upper()
                      for c in range(4, ws.max_column + 1)}
            for chave, _, _, cab in CATEGORIAS:
                achou = [c for c, t in textos.items() if t.startswith(cab)]
                if achou:
                    cols[chave] = achou[0]
            achou = [c for c, t in textos.items() if t in CABECALHO_MEDIA]
            if achou:
                col_media = achou[0]
            if linha_supervisor:
                metas_bloco = {chave: _parse_minimo(ws.cell(row=linha_supervisor, column=cols[chave]).value)
                               for chave, _, _, _ in CATEGORIAS}
            continue

        if isinstance(c3, str) and c4 is None and c3 not in ("TOTAL",):
            # Linha de cabeçalho de bloco (nome do supervisor)
            supervisor_atual = NORMALIZAR_SUPERVISOR.get(c3.strip(), c3.strip())
            linha_supervisor = r
            metas_bloco = {}
            continue

        if not isinstance(c3, int) or not c4 or supervisor_atual is None:
            continue

        # AJUSTE (07/10): linhas novas vêm com o código na frente do nome
        # ("722 - LUIZ GUSTAVO -"). Tira o código do nome e, se ele divergir da
        # coluna C (ex.: C=678 com "679 - RAFAEL COSTA"), vale o do nome.
        m = re.match(r"\s*(\d+)\s*-\s*(.*)$", str(c4))
        if m:
            if int(m.group(1)) != c3:
                print(f"  Aviso: código {c3} na coluna C, mas o nome diz {m.group(1)} ({m.group(2).strip()}); usando {m.group(1)}.")
                c3 = int(m.group(1))
            c4 = m.group(2)
        nome_rca, rota = _nome_e_rota(str(c4))

        categorias_dados = {}
        atingidas = 0
        for chave, label, _, _ in CATEGORIAS:
            meta = metas_bloco.get(chave, 0)
            real = val(r, cols[chave])
            bateu_categoria = real >= meta if meta else False
            if bateu_categoria:
                atingidas += 1
            categorias_dados[chave] = {"label": label, "meta": meta, "real": real, "bateu": bateu_categoria}

        rcas.append({
            "codigo": str(c3),
            "nome": nome_rca,
            "rota": rota,
            "supervisor": supervisor_atual,
            "categorias": categorias_dados,
            "categorias_atingidas": atingidas,
            "total_categorias": len(CATEGORIAS),
            "bateu": atingidas == len(CATEGORIAS),
            "media_pedidos_atual": meta_pedidos_dia,
            # AJUSTE (08/10): "Atual" passa a vir da coluna G da própria RESULTADO
            # (MEDIA/REAL = média de pedidos/dia do RCA); o 4 Pilares só se a célula estiver vazia.
            "posit_atual": val(r, col_media) or posit_atual.get(str(c3), 0),
        })

    return rcas


URL_PLANILHA_METAS = "https://script.google.com/macros/s/AKfycbwg-btEbvpUKtnkijNgUJ4gXQHr_bAyNxfmkPCFN1Zk-FE7IOi-2RgkjMeeMtLljmV01A/exec"
MAPA_CATEGORIA_PERFORMANCE = {
    "bacon": "bacon", "calabresa": "calabresa", "frescais": "frescais", "paes": "paes",
    "lactios": "lacteos", "batata": "batata", "bovino": "bovino", "suino": "suino", "thermo_cat": "thermo",
}


def aplicar_metas_enviadas(rcas):
    """Metas que cada vendedor enviou pelo botão "Feito" do Performance
    (planilha Google). Gravadas nos dados para valer também nos links do
    claude.ai, que não conseguem buscar a planilha ao abrir; o painel no
    GitHub Pages ainda busca de novo ao abrir, pra pegar envios mais novos."""
    import urllib.request
    try:
        with urllib.request.urlopen(URL_PLANILHA_METAS, timeout=60) as resp:
            respostas = json.load(resp)
    except Exception as e:
        print(f"  Aviso: não consegui ler as metas enviadas pelos vendedores ({e}).")
        return
    por_rca = {str(x.get("rca")): x for x in respostas}
    aplicadas = 0
    for r in rcas:
        x = por_rca.get(str(r["codigo"]))
        if not x:
            continue
        est = x.get("estado") or {}
        metas = est.get("metasCategoria") or {}
        for chave_perf, chave_dep in MAPA_CATEGORIA_PERFORMANCE.items():
            v = metas.get(chave_perf)
            if isinstance(v, (int, float)) and chave_dep in r["categorias"]:
                r["categorias"][chave_dep]["meta"] = v
        atingidas = 0
        for cat in r["categorias"].values():
            cat["bateu"] = cat["real"] >= cat["meta"] if cat["meta"] else False
            atingidas += cat["bateu"]
        r["categorias_atingidas"] = atingidas
        r["bateu"] = atingidas == r["total_categorias"]
        desafio = est.get("metaPedidosDia", x.get("metaPedidosDia"))
        if isinstance(desafio, (int, float)) and desafio > 0:
            r["media_pedidos_atual"] = round(desafio)
        aplicadas += 1
    print(f"  Metas enviadas pelos vendedores aplicadas: {aplicadas} RCAs.")


def _atualizar_mestres():
    """Antes de ler: puxa os .xls para as planilhas mestre e acerta os dias
    úteis/trabalhados (ver ..\\atualizar_mestres.py). Falha aqui não impede o painel."""
    import sys
    sys.path.insert(0, os.path.join(PASTA_BASE, ".."))
    try:
        from atualizar_mestres import atualizar_mestres
        atualizar_mestres()
    except Exception as e:
        print(f"  Aviso: não consegui atualizar as planilhas mestre ({e}); usando os dados já salvos.")


if __name__ == "__main__":
    _atualizar_mestres()
    dados = extrair()
    aplicar_metas_enviadas(dados)
    with open(CAMINHO_SAIDA, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    print(f"{len(dados)} RCAs extraídos. Salvo em: {CAMINHO_SAIDA}")
