"""Motor offline da Ferramenta PDPAS: catálogo, enquadramento e alertas.

Não acessa a rede nem o SEI. Recebe um pedido (dict) e devolve o enquadramento do
código, a lista de alertas e as peças que o processo exige.

Severidades:
  BLOQUEIO  a norma veda ou o processo não pode seguir assim.
  ATENCAO   pode seguir, mas falta comprovar ou conferir algo.
  INFO      orientação do fluxo.

Fontes das regras: Decreto 44.322/2023; Portaria SES 473/2023; Manual DIORG/GMPROC de 03/09/2024;
Memorandos Circulares 6/2023 e 1/2025 da GEOF Norte, 2/2024 da SRSNO e 3/2024 da GEOF; Lei 14.133/2021
(art. 75 e § 1º); pactuação em colegiado (um processo por código); pareceres e TRs de 2026 lidos no SEI.
"""
from __future__ import annotations

import json
import math
import re
import unicodedata
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path

import historico
import textos

RAIZ = Path(__file__).resolve().parent
ARQ_CATALOGO = RAIZ / "dados" / "catalogo.json"
ARQ_REGRAS = RAIZ / "regras.json"
ARQ_GASTOS = RAIZ / "dados" / "gastos_ano.json"
ARQ_HISTORICO = RAIZ / "dados" / "historico_pdpas.json"
HIST: dict | None = None
CAT_ITEM = None     # app e testes apontam para Catalogo.achar


# Fontes do teto que dependem de onde a ferramenta roda (24/11/2026). O app do computador e o site trocam estas funções;
# sem troca (testes, uso direto) não há nenhuma.
def _nenhum_pedido(codigo: str, ano: int, excluir: str = "") -> list[dict]:
    return []


def _nenhuma_compra(codigo: str, ano: int) -> list[dict]:
    return []


# pedidos gerados na ferramenta, do mesmo código e ano, ainda não recebidos nem cancelados: [{"id", "valor"}]
PEDIDOS_EM_ANDAMENTO = _nenhum_pedido
# compras da planilha (computador: linhas "Lançado à mão"; site: planilha carregada): [{"valor", "processo", "nf"}]
COMPRAS_DA_PLANILHA = _nenhuma_compra
# site: diz se falta a planilha de compras anteriores (então a compra anterior precisa ser confirmada à mão)
FALTA_PLANILHA = None


def conta_no_teto(pedido: dict, fluxo: dict | None, codigo: str, ano: int, criado_ano: int | None) -> float | None:
    """Valor estimado de um pedido gerado que ainda pesa no teto do código; None se não conta.
    Sai da conta: outro código ou outro ano, cancelado, parado por ata vigente, ou já atestado/pago (aí a nota fiscal
    entra pelo histórico ou pela planilha, e somar o pedido de novo contaria duas vezes)."""
    f = fluxo or {}
    d, ans = f.get("d") or {}, f.get("ans") or {}
    if str(pedido.get("codigo_ses") or "").strip().upper() != str(codigo or "").strip().upper() or not codigo:
        return None
    if criado_ano is not None and criado_ano != ano:
        return None
    if f.get("cancelado") or ans.get("g") == "s" or d.get("t16") == "d" or d.get("t18") == "d":
        return None
    return valor_estimado(pedido)

BLOQUEIO, ATENCAO, INFO = "BLOQUEIO", "ATENCAO", "INFO"
TIPOS = ("consumo", "permanente", "servico")


@dataclass
class Alerta:
    severidade: str
    regra: str
    mensagem: str


def carregar_regras(caminho: Path | None = None) -> dict:
    """regras.json é editado à mão (prazos, teto): BOM do Bloco de Notas e vírgula sobrando viram mensagem clara, não traceback."""
    arq = Path(caminho or ARQ_REGRAS)
    try:
        return json.loads(arq.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as e:
        raise ValueError(f"{arq.name} ilegível ({str(e)[:80]}). Confira vírgulas, aspas e chaves.") from e


def _sem_acento(t: str) -> str:
    n = unicodedata.normalize("NFKD", t or "")
    return "".join(c for c in n if not unicodedata.combining(c)).upper()


class Catalogo:
    def __init__(self, caminho: Path = ARQ_CATALOGO):
        dados = json.loads(Path(caminho).read_text(encoding="utf-8"))
        self.atualizado_em = dados.get("atualizado_em", "")
        self.itens = dados["itens"]
        self.por_codigo: dict[str, list[dict]] = {}
        vistos: set[tuple] = set()
        unicos = []
        for it in self.itens:
            cod = str(it.get("codigo_ses") or "").strip().upper()
            # o Power BI repete a última linha de cada página na seguinte: linha idêntica entra uma vez só
            assinatura = (cod, it.get("codigo_br"), it.get("classif_categoria"), it.get("desc_resumida"), it.get("desc_completa"))
            if assinatura in vistos:
                continue
            vistos.add(assinatura)
            unicos.append(it)
            it["_desc"] = _sem_acento(it.get("desc_resumida") or "")
            it["_busca"] = _sem_acento(f"{cod} {it.get('codigo_br') or ''} ") + it["_desc"]
            self.por_codigo.setdefault(cod, []).append(it)
        self.itens = unicos

    def achar(self, codigo: str) -> list[dict]:
        return self.por_codigo.get(str(codigo or "").strip().upper(), [])

    def buscar(self, termo: str, limite: int = 40) -> list[dict]:
        """Todas as palavras precisam aparecer; descrição que começa com o termo vem antes."""
        partes = _sem_acento(termo).split()
        if not partes:
            return []
        achados = [i for i in self.itens if all(p in i["_busca"] for p in partes)]
        inicio = partes[0]
        achados.sort(key=lambda i: (not i["_desc"].startswith(inicio), i["_desc"]))
        return achados[:limite]


def categoria_num(item: dict) -> str:
    return str(item.get("classif_categoria") or "").split(" - ")[0].strip()


def classificar_codigo(codigo: str) -> dict:
    """Memorando Circular 2/2024: só números = almoxarifado (NUAL); inicia com P = patrimônio (NPDOC).

    Os códigos do catálogo que a nota do memorando não cobre (M…, pontuados, GRAMPE) voltam
    como 'indefinido': a ferramenta não deduz o destino deles.
    """
    c = str(codigo or "").strip().upper()
    if re.fullmatch(r"\d+", c):
        return {"grupo": "consumo", "destino": "NUAL", "despesa": "corrente (custeio)"}
    if c.startswith("P"):
        return {"grupo": "permanente", "destino": "NPDOC", "despesa": "capital"}
    return {"grupo": "indefinido", "destino": None, "despesa": None}


def moeda(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _num(v) -> float | None:
    """Número vindo da janela ou de um rascunho: aceita 12, '12', '1.200', '1.234,56'; texto ilegível vira None."""
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return float(v) if math.isfinite(v) else None     # NaN/inf vindos de um JSON adulterado não passam pelo teto
    return historico.numero(v)


def ler_unidades(pedido: dict) -> tuple[dict, list[str]]:
    """Uma unidade por linha: 'UBS 6 Sobradinho; 1'. Aceita ';', ':', '=' ou tabulação (colar do Excel)."""
    unidades: dict[str, float] = {}
    invalidas = []
    for linha in str(pedido.get("unidades") or "").splitlines():
        if not linha.strip():
            continue
        m = re.fullmatch(r"\s*(.+?)\s*[;:=\t]\s*([\d.,]+)\s*", linha)
        q = historico.numero(m[2]) if m else None      # "1.500" é mil e quinhentos, não 1,5
        if q and q > 0:
            unidades[m[1]] = unidades.get(m[1], 0) + q
        else:
            invalidas.append(linha.strip())
    return unidades, invalidas


def ler_patrimonios(pedido: dict) -> list[dict]:
    """Uma linha por bem: 'patrimônio; descrição, marca, modelo, série; local; SEI da ficha SisGepat'."""
    bens = []
    for linha in str(pedido.get("patrimonios") or "").splitlines():
        if not linha.strip():
            continue
        partes = [x.strip() for x in re.split(r"[;\t]", linha)] + ["", "", ""]
        bens.append({"patrimonio": partes[0], "descricao": partes[1], "local": partes[2], "ficha": partes[3]})
    return bens


def somar_quantidades(pedido: dict) -> tuple[float, dict]:
    """Usa a grade por GSAP; sem ela, a lista de unidades; no reparo, um por patrimônio."""
    q = {k: _num(v) for k, v in (pedido.get("quantidades") or {}).items() if v not in (None, "", 0, "0")}
    q = {k: v for k, v in q.items() if v and v > 0}      # negativo (rascunho/JSON editado) reduzia o total e escapava do teto
    q = q or ler_unidades(pedido)[0]
    # um por bem: linha sem nº de patrimônio ganha rótulo próprio (antes todas viravam a mesma chave e contavam 1)
    if not q and pedido.get("tipo") in (None, "", "servico"):     # lista de bens escondida de um rascunho de serviço não vira quantidade
        q = {(f"Patrimônio {b['patrimonio']}" if b["patrimonio"] else f"Bem {i} (sem nº de patrimônio)"): 1.0
             for i, b in enumerate(ler_patrimonios(pedido), 1)}
    return sum(q.values()), q


def area_parecer(item: dict | None, codigo: str, regras: dict) -> dict | None:
    """Sugere a área central do parecer pelo que os pareceres de 2026 mostraram. Ver regras.json."""
    desc = _sem_acento((item or {}).get("desc_resumida") or "")
    for r in regras.get("area_parecer") or []:
        q = r["quando"]
        # palavra no início de palavra (como textos._grupo): "CUBA" não pode casar "INCUBADORA"
        if (any(re.search(rf"\b{re.escape(p)}", desc) for p in q.get("palavras", [])) or any(codigo.startswith(p) for p in q.get("prefixo", []))
                or (item and categoria_num(item) in q.get("categoria", []))):
            return r
    return None


def valor_estimado(pedido: dict) -> float | None:
    total, _ = somar_quantidades(pedido)
    preco = _num(pedido.get("preco_estimado"))
    if preco is None or preco < 0:
        return None
    from decimal import ROUND_HALF_UP, Decimal
    # em float, 0,125 × 1 arredondava para 0,12 (meio-par sobre o binário): no limiar do teto cada centavo conta
    return float((Decimal(str(total)) * Decimal(str(preco))).quantize(Decimal("0.01"), ROUND_HALF_UP))


def gastos_do_ano(ano: int, caminho: Path | None = None) -> list[dict]:
    """Lançamentos: {"ano", "natureza", "valor", "origem"}. Vêm da lista da GEOF ou de processos anteriores.
    Arquivo editado à mão: ilegível vira aviso (e lista vazia), não derruba a análise; "valor" com vírgula é aceito."""
    caminho = Path(caminho or ARQ_GASTOS)
    if not caminho.exists():
        return []
    try:
        dados = json.loads(caminho.read_text(encoding="utf-8-sig"))
        if not isinstance(dados, list):
            raise ValueError("esperava uma lista de lançamentos")
    except (OSError, ValueError) as e:
        aviso = f"{caminho.name} ilegível ({str(e)[:60]}): os lançamentos do ano não entraram na conta do teto."
        avisos = getattr(historico, "AVISOS", None)
        if isinstance(avisos, list) and aviso not in avisos:
            avisos.append(aviso)
        return []
    return [g for g in dados if isinstance(g, dict) and str(g.get("ano") or "").strip() == str(ano)]


def hist() -> dict:
    global HIST
    if HIST is None:
        HIST = historico.carregar(ARQ_HISTORICO)
    return HIST


def _br(iso: str) -> str:
    return "/".join(reversed(iso.split("-"))) if iso else ""


def gasto_do_item(codigo: str, ano: int) -> dict | None:
    """Já comprado do código no ano: notas fiscais do histórico + lançamentos de gastos_ano.json com esse código
    (processos que ainda não aparecem nas Relações). None se não houver registro."""
    g = historico.gasto_ano(hist(), codigo, ano)
    total, processos = g["total"], set(g["processos"])
    nas_notas = set(g["processos"])         # processo já contado pelas NFs: a ordem da GEOF é a mesma compra
    cod = str(codigo or "").strip().upper()
    vistos = set()
    for lanc in gastos_do_ano(ano):
        if str(lanc.get("codigo") or "").strip().upper() != cod or lanc.get("processo") in nas_notas:
            continue
        # só junta como repetido quando há nº SEI igual; sem SEI, dois lançamentos iguais são somados (decisão de 05/11/2026)
        marca = (lanc.get("processo"), lanc.get("sei"), _num(lanc.get("valor")) or 0.0)
        if lanc.get("sei") and marca in vistos:
            continue
        vistos.add(marca)
        total = round(total + marca[2], 2)
        if lanc.get("processo"):
            processos.add(lanc["processo"])
    return {"total": total, "processos": sorted(processos), "ate": _br(g["ate"])} if total else None


def gasto_considerado(pedido: dict, regras: dict) -> dict:
    """O 'já comprado' que entra na conta do teto, igual na faixa ao vivo e no alerta.

    Soma por item (Decreto 44.322, art. 7º, III): o maior entre o histórico do código e o valor digitado em
    'Já comprado no ano'. O digitado nunca abaixa o histórico. Se houver natureza escolhida, vale o maior entre
    isso e o total lançado da natureza (lista da GEOF), que é a soma por ramo do Decreto 44.330, art. 234.
    """
    ano = regras["ano"]
    natureza = pedido.get("natureza")
    ja_natureza = sum((_num(g.get("valor")) or 0.0) for g in gastos_do_ano(ano) if natureza and g.get("natureza") == natureza)
    codigo = str(pedido.get("codigo_ses") or "").strip().upper()
    registro = gasto_do_item(codigo, ano) if codigo else None
    ja_lista = registro["total"] if registro else 0.0
    # compras da planilha que o histórico ainda não tem (o mesmo processo não entra duas vezes)
    contados = set(registro["processos"]) if registro else set()
    planilha = [c for c in (COMPRAS_DA_PLANILHA(codigo, ano) if codigo else [])
                if (_num(c.get("valor")) or 0) > 0 and not (c.get("processo") and c["processo"] in contados)]
    ja_planilha = round(sum(_num(c["valor"]) for c in planilha), 2)
    # pedidos em andamento do mesmo código no ano: o que já foi pedido conta no teto (Colegiado/decisão de 24/11/2026)
    andamento = [p for p in (PEDIDOS_EM_ANDAMENTO(codigo, ano, str(pedido.get("_pedido_id") or "")) if codigo else [])
                 if (_num(p.get("valor")) or 0) > 0]
    ja_andamento = round(sum(_num(p["valor"]) for p in andamento), 2)
    comprado = round(ja_lista + ja_planilha, 2)
    digitado = _num(pedido.get("gasto_item"))
    ja_item = (max(comprado, digitado) if digitado is not None else comprado) + ja_andamento
    return {"ja": round(max(ja_natureza, ja_item), 2), "ja_item": round(ja_item, 2), "ja_lista": round(ja_lista, 2),
            "ja_natureza": round(ja_natureza, 2), "digitado": digitado, "registro": registro,
            "ja_planilha": ja_planilha, "n_planilha": len(planilha), "ja_comprado": comprado,
            "ja_andamento": ja_andamento, "andamento": [str(p.get("id") or "") for p in andamento]}


def origem_do_gasto(g: dict) -> list[str]:
    """Partes do 'já considerado no ano', em texto curto para a tela."""
    partes = []
    if g["ja_lista"]:
        partes.append(f"{moeda(g['ja_lista'])} em notas do histórico")
    if g["ja_planilha"]:
        partes.append(f"{moeda(g['ja_planilha'])} da planilha de compras")
    if g["digitado"] is not None and g["digitado"] > g["ja_comprado"]:
        partes.append(f"{moeda(g['digitado'])} informados à mão")
    if g["ja_andamento"]:
        n = len(g["andamento"])
        partes.append(f"{moeda(g['ja_andamento'])} em {n} pedido{'s' if n > 1 else ''} já gerado{'s' if n > 1 else ''} e em "
                      f"andamento ({', '.join(g['andamento'])})")
    if g["ja_natureza"] > g["ja_item"]:
        partes.append(f"{moeda(g['ja_natureza'])} na natureza da despesa escolhida")
    return partes


def config_intencao(regras: dict) -> dict:
    """Parâmetros da abertura do permanente (áudios da Diretoria, 17/11/2026). regras.json antigo, sem o bloco, usa estes padrões."""
    c = regras.get("intencao") or {}
    try:
        dias = int(c.get("prazo_dias_uteis") or 7)
    except (TypeError, ValueError):
        dias = 7
    if dias < 1:
        dias = 7
    extras = c.get("destinatarios_extras") or []
    if isinstance(extras, str):
        extras = extras.splitlines()
    feriados = c["feriados"] if isinstance(c.get("feriados"), list) else textos.FERIADOS_PADRAO     # [] = sem feriado, de propósito
    return {"dias": dias, "feriados": feriados,
            "extras": [str(e).strip() for e in extras if str(e).strip()]}


def prazo_intencao(regras: dict, hoje: date) -> str:
    c = config_intencao(regras)
    return textos.prazo_gsaps(hoje, c["dias"], c["feriados"])


def sugestoes(pedido: dict, regras: dict, hoje: date | None = None) -> dict:
    """O que o histórico já responde sobre o código, para ninguém precisar digitar."""
    hoje = hoje or date.today()
    codigo = str(pedido.get("codigo_ses") or "").strip().upper()
    if pedido.get("tipo") == "servico":
        campos = textos.para_servico(pedido, ler_patrimonios(pedido), hoje)
        campos["prazo_manifestacao"] = prazo_intencao(regras, hoje)
        return {"campos": campos}
    if not codigo:
        return {}
    h = hist()
    compras = historico.compras_do_item(h, codigo)
    visao = h["visao"].get(codigo)
    total, _ = somar_quantidades(pedido)
    s: dict = {"compras": [{"data": _br(c["nf_data"]), "qtd": c["qtd"], "unit": c["unit"], "total": c["total"],
                            "processo": c["processo"], "fornecedor": c["fornecedor"]} for c in compras[:3]]}
    if compras and compras[0]["unit"]:
        s["preco"], s["preco_fonte"] = compras[0]["unit"], f"última NF, {_br(compras[0]['nf_data'])}"
    elif visao and visao.get("atas"):
        s["preco"], s["preco_fonte"] = visao["atas"][0]["preco"], "preço da ata"
    elif visao and visao.get("precos"):
        s["preco"], s["preco_fonte"] = visao["precos"][0]["preco"], f"última compra da SES, {_br(visao['precos'][0]['data'])}"
    teto = _num((regras.get("teto_dispensa") or {}).get("valor"))
    gasto = gasto_considerado(pedido, regras)            # a mesma conta do alerta de teto
    s["gasto_ano"] = gasto["ja"]
    s["gasto_origem"] = origem_do_gasto(gasto)      # de onde vem o "já considerado": a tela diz, para o teto não parecer mágico
    if teto is not None:        # "valor": null = teto ainda não confirmado (o alerta de teto avisa); sem saldo nem qtd_max
        s["saldo_teto"] = round(teto - gasto["ja"], 2)      # negativo = já passou
        s["teto"], s["alerta_pct"] = teto, float(regras.get("alerta_teto_percentual") or 0.8)
        preco_ref = _num(pedido.get("preco_estimado")) or s.get("preco")
        if preco_ref and s["saldo_teto"] > 0:
            from decimal import ROUND_FLOOR, Decimal
            s["qtd_max"] = int((Decimal(str(s["saldo_teto"])) / Decimal(str(preco_ref))).to_integral_value(ROUND_FLOOR))
    c = historico.cmm(h, codigo, hoje)
    if c and c["valor"]:
        qtd_meses = "1 mês" if c["meses"] == 1 else f"{c['meses']} meses"
        periodo = _br(c["de"]) if c["de"] == c["ate"] else f"{_br(c['de'])} a {_br(c['ate'])}"
        s["cmm"], s["cmm_fonte"] = c["valor"], f"média de {qtd_meses} com saída, {periodo}"
        s["_just"] = (f"Consumo médio mensal de {c['valor']:,}".replace(",", ".") + f" na APS Norte, média de {qtd_meses} "
                                f"com saída, {periodo}, conforme Relação de Estoque Físico da GEOF.")
        s["consumo_mensal"] = f"{c['valor']:,}".replace(",", ".")
        if total:
            s["tempo_duracao"] = f"{total / c['valor']:.1f} meses".replace(".", ",")
    desc = (compras[0]["descricao"] if compras else "") or (visao or {}).get("descricao", "")
    categoria = ""
    if CAT_ITEM:
        achado = CAT_ITEM(codigo)
        if achado:
            desc, categoria = achado[0].get("desc_resumida") or desc, categoria_num(achado[0])
    s["campos"] = textos.para_item(desc, _sem_acento(desc), categoria, pedido.get("tipo") or "consumo", c["valor"] if c else None, visao, hoje, codigo)
    aprendidos = (h.get("textos") or {}).get(codigo) or {}
    frase_cmm = re.compile(r"\s*O consumo médio da APS Norte é de [\d.]+ por mês\.")
    for campo, v in aprendidos.items():
        texto = v.get("texto") if isinstance(v, dict) else None
        if not isinstance(texto, str):
            continue
        if campo == "justificativa" and frase_cmm.search(texto):
            # o texto validado antes traz o CMM daquele dia; troca pelo atual para o TR não citar dois valores
            texto = frase_cmm.sub("", texto).rstrip()
            if c and c["valor"]:
                texto += f" O consumo médio da APS Norte é de {c['valor']:,} por mês.".replace(",", ".")
        s["campos"][campo] = texto
    if not c:
        s["campos"]["consumo_mensal"] = f"{visao['cmm_rede']:g} na rede SES" if visao and visao.get("cmm_rede") else "Sem registro na APS Norte"
        s["campos"]["tempo_duracao"] = "Não estimado por falta de consumo registrado"
    s["campos"]["prazo_manifestacao"] = prazo_intencao(regras, hoje)    # 7 dias úteis, com feriados do regras.json
    s["campos"].setdefault("antecedentes", "Não há compra deste item via PDPAS no histórico da APS Norte.")
    if compras:
        u = compras[0]
        qtd = f"{u['qtd']:g}" if u["qtd"] else "?"
        s["antecedentes"] = (f"Última compra via PDPAS: NF {u['nf']}, de {_br(u['nf_data'])}, {qtd} unidades, processo {u['processo']}, "
                             f"fornecedor {u['fornecedor']}.")
    if visao:
        s["visao_data"] = _br(visao.get("data", ""))
        s["art21"] = art21_da_visao(visao)
    for chave, campo in (("_just", "just_quantidade"), ("consumo_mensal", "consumo_mensal"), ("tempo_duracao", "tempo_duracao"),
                         ("antecedentes", "antecedentes")):
        if s.get(chave):
            s["campos"][campo] = s.pop(chave)
    return s


def aprender(pedido: dict, hoje: date | None = None) -> None:
    """Guarda o texto validado para sugerir igual da próxima vez. Sugestão não validada não entra."""
    codigo = str(pedido.get("codigo_ses") or "").strip().upper()
    if not codigo:
        return
    pendentes = set(pedido.get("nao_validados") or [])
    data = (hoje or date.today()).isoformat()
    for campo in ("justificativa", "beneficios"):
        texto = (pedido.get(campo) or "").strip()
        if texto and campo not in pendentes and "[A PREENCHER" not in texto and "[VALIDAR" not in texto:
            hist().setdefault("textos", {}).setdefault(codigo, {})[campo] = {"texto": texto, "data": data}


def art21_da_visao(v: dict) -> dict:
    """Converte a Visão Gerencial nos campos da janela. Quem decide o que fazer com isso é ela."""
    campos = {"data_consulta": _br(v.get("data", "")), "sei_visao": v.get("sei", "")}
    # Só afirma "sem" quando o documento foi lido de fato; leitura incerta deixa o campo para ela confirmar.
    if v.get("central") is not None:
        campos["estoque_central"] = "tem" if v["central"] else "sem"
    if v.get("atas"):
        campos["ata"] = "vigente"
    elif not v.get("atas_incerto"):
        campos["ata"] = "sem"
    outras = [o for o in v.get("outras_unidades") or [] if o.get("disponivel")]
    campos["outra_unidade"] = "reservado" if outras else "nao"
    if outras:
        campos["texto_outra_unidade"] = "Há estoque em " + ", ".join(f"{o['local']} ({o['disponivel']:g})" for o in outras) + \
            ": [A PREENCHER: por que não atende]"
    if v.get("atas"):
        a = v["atas"][0]
        campos.update(ata_numero=a["ata"], ata_fornecedor=a["fornecedor"], ata_validade=a["validade"], ata_saldo=a["saldo"],
                      ata_preco=a["preco"])
    if v.get("abertos"):
        p = v["abertos"][0]["processo"]
        campos.update(processo_andamento="contratacao", processo_numero=p, ata_empenho=f"pedido em aberto no processo {p}")
    else:
        campos["processo_andamento"] = "nao"
    return campos


def naturezas_do_ano(ano: int) -> list[dict]:
    soma: dict[str, float] = {}
    for g in gastos_do_ano(ano):
        nat, valor = str(g.get("natureza") or "").strip(), _num(g.get("valor"))
        if nat and valor is not None:       # lançamento incompleto ou "valor" com vírgula não pode impedir a janela de abrir
            soma[nat] = soma.get(nat, 0) + valor
    return [{"natureza": k, "total": round(v, 2)} for k, v in sorted(soma.items())]


def _data(txt):
    if not txt:
        return None
    try:
        return datetime.strptime(txt, "%d/%m/%Y").date()
    except ValueError:
        return date.fromisoformat(txt)      # AAAA-MM-DD, comum em arquivo editado à mão


def _ler_data(txt):
    try:
        return _data(str(txt or "").strip())
    except ValueError:
        return None


NOME_TIPO = {"consumo": "consumo", "permanente": "permanente", "servico": "serviço"}


def _prazos(tipo: str, regras: dict, hoje: date) -> list[Alerta]:
    p = regras["prazos_geof"]
    if p.get("ano") != regras["ano"] or not p.get(tipo):
        ref = p["referencia"]
        return [Alerta(ATENCAO, "Prazo GEOF",
                       f"Prazo de {regras['ano']} ainda não publicado pela GEOF. Em {ref['ano']} o limite para "
                       f"{NOME_TIPO.get(tipo, tipo)} foi {ref[tipo]} ({ref['fonte']}).")]
    limite = _ler_data(p[tipo])
    if limite is None:
        return [Alerta(ATENCAO, "Prazo GEOF", f"Não consegui ler a data do prazo de {NOME_TIPO.get(tipo, tipo)} em regras.json "
                                              f"(prazos_geof: “{p[tipo]}”). Escreva como dd/mm/aaaa.")]
    faltam = (limite - hoje).days
    if faltam < 0:
        return [Alerta(BLOQUEIO, "Prazo GEOF", f"O prazo para {NOME_TIPO.get(tipo, tipo)} ({p[tipo]}) já passou.")]
    if faltam <= 10:
        return [Alerta(ATENCAO, "Prazo GEOF", f"Faltam {faltam} dias para o prazo de {NOME_TIPO.get(tipo, tipo)} ({p[tipo]}).")]
    return [Alerta(INFO, "Prazo GEOF", f"Prazo para {NOME_TIPO.get(tipo, tipo)}: {p[tipo]} (faltam {faltam} dias).")]


def _teto(pedido: dict, regras: dict, grupo: str) -> list[Alerta]:
    t = regras.get("teto_dispensa") or {}
    if _num(t.get("valor")) is None:        # "valor": null = ainda não confirmado: antes era TypeError e nenhum alerta saía
        return [Alerta(ATENCAO, "Teto da dispensa",
                       "Valor do teto da dispensa não cadastrado em regras.json (teto_dispensa.valor): não dá para conferir "
                       "o teto. Confirme o valor do ano com a GEOF e preencha.")]
    teto = float(_num(t.get("valor")))
    valor = valor_estimado(pedido)
    # a flag só vale em serviço: um rascunho de serviço reaberto como consumo aplicava o teto de engenharia a uma compra
    engenharia = grupo == "servico" and bool(pedido.get("engenharia"))
    if engenharia:
        te = regras.get("teto_engenharia") or {}
        if _num(te.get("valor")) is None:
            return [Alerta(ATENCAO, "Teto da dispensa",
                           "Obra ou serviço de engenharia segue o art. 75, I, não o inciso II, e o valor desse teto "
                           "ainda não está em regras.json. Confirme com a GEOF e corrija a fundamentação do TR.")]
        teto = float(_num(te["valor"]))
    regra_inciso = "art. 75, I" if engenharia else "art. 75, II"
    if valor is None:
        return [Alerta(ATENCAO, "Teto da dispensa", "Sem preço estimado: não dá para conferir o teto da dispensa.")]
    fmt = moeda
    g = gasto_considerado(pedido, regras)
    ja, registro, gasto_item = g["ja"], g["registro"], g["digitado"]
    soma = round(valor + ja, 2)
    composicao = f"{fmt(valor)} deste pedido" + (f" + {fmt(ja)} já lançados = {fmt(soma)}" if ja else "")
    origens = []
    if registro and (gasto_item is None or gasto_item <= g["ja_comprado"]):
        partes = ([f"notas de {regras['ano']} até {registro['ate']}"] if registro["ate"] else ["lançamentos do ano"]) + \
                 ([f"processos {', '.join(registro['processos'])}"] if registro["processos"] else [])
        origens.append(", ".join(partes))
    if g["ja_planilha"] and (gasto_item is None or gasto_item <= g["ja_comprado"]):
        origens.append(f"{fmt(g['ja_planilha'])} da planilha de compras ({g['n_planilha']} lançamento"
                       + ("s" if g["n_planilha"] > 1 else "") + ")")
    if gasto_item is not None and gasto_item > g["ja_comprado"]:
        origens.append(f"{fmt(gasto_item)} informados à mão")
    if g["ja_andamento"]:
        n = len(g["andamento"])
        origens.append(f"{fmt(g['ja_andamento'])} em {n} pedido{'s' if n > 1 else ''} em andamento: {', '.join(g['andamento'])}")
    if origens:
        composicao += f" ({'; '.join(origens)})"
    saida = []
    if gasto_item is not None and gasto_item < g["ja_comprado"]:
        saida.append(Alerta(ATENCAO, "Teto da dispensa",
                            f"Você informou {fmt(gasto_item)} já comprados, mas o histórico e a planilha registram "
                            f"{fmt(g['ja_comprado'])} do código em {regras['ano']}. A conta usa o maior valor."))
    if g["ja_andamento"]:
        saida.append(Alerta(ATENCAO, "Teto da dispensa",
                            f"Já há pedido deste código em andamento em {regras['ano']} ({', '.join(g['andamento'])}): o valor "
                            "estimado dele entra na conta. Pedido repetido ou desistido: cancele-o na tela Minutas/documentos "
                            "gerados, para sair da conta."))
    aviso = ("O Decreto 44.322, art. 7º, III, soma por item; o Decreto 44.330, art. 234, § 1º, soma por unidade gestora e "
             "por ramo de atividade (classe do item no catálogo), então só a lista da GEOF dá o total real.")
    if soma > teto:
        saida.append(Alerta(BLOQUEIO, "Teto da dispensa",
                            f"{composicao}, acima do teto de {fmt(teto)}. Não cabe dispensa pelo {regra_inciso}: a licitação "
                            f"é do nível central (Decreto 44.322, art. 7º, IV). {aviso}"))
    elif soma >= teto * float(regras.get("alerta_teto_percentual") or 0.8):
        saida.append(Alerta(ATENCAO, "Teto da dispensa",
                            f"{composicao} ({soma / teto:.1%} do teto de {fmt(teto)}). {aviso} Abrir código parecido para "
                            "seguir comprando é fracionamento: a ASSITEC indeferiu pedido assim (SEI 213219196)."))
    else:
        saida.append(Alerta(INFO, "Teto da dispensa",
                            f"{composicao} ({soma / teto:.0%} do teto de {fmt(teto)}). {aviso}"))
    if FALTA_PLANILHA and FALTA_PLANILHA() and gasto_item is None:
        saida.append(Alerta(ATENCAO, "Compra anterior",
                            f"Sem planilha de compras neste aparelho: confirme se já houve compra deste item em {regras['ano']}. "
                            "Se houve, carregue a planilha de compras ou preencha o valor em “Já comprado deste item no ano” "
                            "(etapa Quantidade, “Ajustar a soma do teto”). Se não houve, preencha 0."))
    elif not pedido.get("natureza") and gasto_item is None and not g["ja_comprado"]:
        saida.append(Alerta(ATENCAO, "Teto da dispensa",
                            "Informe o já comprado deste item no ano ou escolha a natureza da despesa."))
    if engenharia:
        saida.append(Alerta(ATENCAO, "Teto da dispensa", f"Obra ou engenharia usa o art. 75, I ({fmt(teto)} no decreto "
                                                         "federal). Confirme o valor com a GEOF e cite o inciso I no TR."))
    elif not t.get("confirmado_gdf"):
        saida.append(Alerta(ATENCAO, "Teto da dispensa",
                            f"Valor do teto em {regras['ano']} baseado no decreto federal; confirme com a GEOF "
                            "se o GDF adota o mesmo valor."))
    return saida


def _quantidades(pedido: dict, regras: dict) -> list[Alerta]:
    total, q = somar_quantidades(pedido)
    unidades, invalidas = ler_unidades(pedido)
    # grade de zeros/vazios com a quantidade real nas unidades não é "por GSAP" (dava "área fora da lista" falso)
    por_gsap = any((_num(v) or 0) > 0 for v in (pedido.get("quantidades") or {}).values())
    saida = []
    if invalidas:
        saida.append(Alerta(ATENCAO, "Quantidade", "Unidade sem quantidade legível, fora do TR: "
                                                   f"{'; '.join(invalidas)}. Use o formato 'UBS 6 Sobradinho; 1'."))
    if not q:
        return [*saida, Alerta(BLOQUEIO, "Quantidade", "Nenhuma quantidade informada por GSAP, unidade ou patrimônio.")]
    if list(unidades) == ["APS Norte"]:
        saida.append(Alerta(ATENCAO, "Quantidade", "Só o total: o TR sai sem a divisão por GSAP, e o Mem. Circ. 2/2024 pede o "
                                                   "consolidado das GSAPs. Serve para rascunho ou item da região inteira."))
    if pedido.get("tipo") != "servico" and not (pedido.get("just_quantidade") or "").strip():
        saida.append(Alerta(ATENCAO, "Quantidade", "Falta a memória do quantitativo: de onde saiu o número. Os NLF usam a "
                                                   "movimentação de 12 meses, só dos meses com abastecimento. A área "
                                                   "técnica central pede esclarecimento quando não há (SEI 218202353)."))
    if por_gsap and unidades and round(sum(unidades.values()), 6) != round(total, 6):     # 0,1+0,2 ≠ 0,3 em float
        saida.append(Alerta(ATENCAO, "Quantidade", f"Soma das unidades ({sum(unidades.values()):g}) difere da "
                                                   f"soma das GSAPs ({total:g})."))
    fora = [k for k in q if por_gsap and k not in regras["gsaps"] and k not in (pedido.get("outras_areas") or [])]
    if fora:
        saida.append(Alerta(ATENCAO, "Quantidade", f"Área fora da lista de GSAPs: {', '.join(fora)}."))
    consolidado = _num(pedido.get("consolidado"))
    if consolidado is not None and round(consolidado, 6) != round(total, 6):
        saida.append(Alerta(ATENCAO, "Quantidade",
                            f"Soma das GSAPs ({total:g}) difere do consolidado informado ({consolidado:g})."))
    limite = regras.get("limite_qtd_por_ubs")
    if limite:
        acima = [k for k, v in q.items() if v > float(limite)]
        if acima:
            saida.append(Alerta(ATENCAO, "Quantidade", f"Acima do limite configurado ({limite}): {', '.join(acima)}."))
    sem = [g for g in regras["gsaps"] if g not in q]
    if sem and por_gsap:
        saida.append(Alerta(INFO, "Quantidade",
                            f"{len(sem)} GSAPs sem quantidade. Confira se todas foram consultadas "
                            "(a Diretoria avalia se o TR contemplou todas as áreas)."))
    return saida


def analisar(pedido: dict, catalogo: Catalogo, regras: dict, hoje: date | None = None) -> dict:
    hoje = hoje or date.today()
    tipo = pedido.get("tipo")
    alertas: list[Alerta] = []
    add = alertas.append
    codigo = str(pedido.get("codigo_ses") or "").strip().upper()
    itens = catalogo.achar(codigo) if codigo else []
    item = itens[0] if itens else None
    enq = {"grupo": "servico", "destino": None, "despesa": "corrente (custeio)"}

    if tipo not in TIPOS:
        add(Alerta(BLOQUEIO, "Tipo", "Escolha consumo, permanente ou serviço."))
        tipo = None

    if tipo != "servico":
        enq = classificar_codigo(codigo) if codigo else {"grupo": "indefinido", "destino": None, "despesa": None}
        if not codigo:
            add(Alerta(BLOQUEIO, "Art. 22, III", "Sem código SES. Só itens padronizados podem ser comprados."))
        elif not item:
            add(Alerta(BLOQUEIO, "Art. 22, III",
                       f"Código {codigo} não está no catálogo (versão {catalogo.atualizado_em[:10]}). Produto para "
                       "saúde não padronizado é vedado. Pelo Mem. Circ. 2/2024, peça a padronização à "
                       "SES/SAIS/ASSISTEC por TR com justificativa técnica."))
        else:
            if len({(_sem_acento(i.get("desc_resumida")), categoria_num(i)) for i in itens}) > 1:
                add(Alerta(ATENCAO, "Catálogo", f"O código {codigo} aparece com descrições ou categorias diferentes no "
                                                "catálogo. Confira qual vale."))
            if enq["grupo"] == "indefinido":
                add(Alerta(ATENCAO, "Código", f"Código {codigo} fora da regra P/número do Mem. Circ. 2/2024. "
                                              "Confirme natureza e destino com o NUAL ou o NPDOC."))
            if categoria_num(item) == "0":
                add(Alerta(ATENCAO, "Art. 22, I",
                           "Medicamento: confira se é do componente especializado (vedado). O catálogo não traz "
                           "essa informação."))
        if tipo and enq["grupo"] in ("consumo", "permanente") and tipo != enq["grupo"]:
            add(Alerta(BLOQUEIO, "Natureza do item",
                       f"Pedido aberto como {NOME_TIPO.get(tipo, tipo)}, mas o código {codigo} é de {enq['grupo']} "
                       f"(destino {enq['destino']}). Foi o que encerrou o processo do colchonete 14183 em 2025."))

    outros = [str(c).strip().upper() for c in (pedido.get("outros_codigos") or []) if str(c).strip()]
    if outros:
        add(Alerta(ATENCAO, "Um processo por código",
                   f"Pactuado em colegiado: um processo por código de material. Separe: {', '.join(outros)}."))

    vedacao = pedido.get("vedacao")
    if vedacao:
        add(Alerta(BLOQUEIO, "Decreto 44.322, art. 4º", "Vedado com PDPAS, inciso "
                   + (regras.get("vedacoes_art4") or {}).get(vedacao, vedacao) + "."))
    if pedido.get("ordem_judicial"):
        add(Alerta(BLOQUEIO, "Art. 22, II", "Vedado usar PDPAS para atender ordem judicial."))
    if pedido.get("encargos"):
        add(Alerta(BLOQUEIO, "Art. 22, IV", "Vedado pagar juros, multas e correção monetária com PDPAS."))
    if pedido.get("despesa_anterior"):
        add(Alerta(BLOQUEIO, "Art. 22, V", "Vedado pagar despesa anterior ao crédito do recurso."))

    # Art. 21: as três condições são cumulativas; § 1º e § 2º são as saídas previstas.
    if tipo in ("consumo", "permanente"):
        estoque = pedido.get("estoque_central")
        if estoque == "tem":
            add(Alerta(BLOQUEIO, "Art. 21, I", "Item disponível nas Centrais de Abastecimento: não cabe PDPAS."))
        elif estoque != "sem":
            add(Alerta(ATENCAO, "Art. 21, I", "Falta o Nada Consta do sistema de materiais, atual."))
        ata = pedido.get("ata")
        if ata == "vigente":
            add(Alerta(BLOQUEIO, "Art. 21, II", "Há ata vigente: o processo para, porque a SES já adquire o item pela ata "
                                                "(Colegiado 23/11/2026)."))
        elif ata == "atrasada":
            if pedido.get("comprovante_atraso"):
                add(Alerta(INFO, "Art. 21, § 1º", "Atraso ou inexecução de ata comprovado: exceção aplicável."))
            else:
                add(Alerta(BLOQUEIO, "Art. 21, § 1º", "Ata com atraso: anexe a comprovação motivada."))
        elif ata != "sem":
            add(Alerta(ATENCAO, "Art. 21, II", "Falta a Declaração de Inexistência de ARP válida."))
        andamento = pedido.get("processo_andamento")
        if andamento == "homologada":
            add(Alerta(BLOQUEIO, "Manual, atividade 3", "Depois da homologação da licitação regular não dá para seguir "
                                                        "por PDPAS."))
        elif andamento in ("etp", "tr", "licitacao", "contratacao"):
            add(Alerta(ATENCAO, "Processo em andamento",
                       "Cite o processo no TR e avise a área técnica central para ela ajustar a grade de distribuição e "
                       "evitar compra em duplicidade (SEI 218244958)."))
        elif andamento != "nao" and tipo == "permanente":
            add(Alerta(ATENCAO, "Processo em andamento", "A tabela do parecer técnico pergunta se há processo de aquisição "
                                                         "em andamento e em qual fase. Confira antes."))
        if tipo == "consumo":
            outra = pedido.get("outra_unidade")
            if outra == "sim":
                add(Alerta(BLOQUEIO, "Decreto 44.322, art. 7º, V", "O item está disponível em outra unidade da rede: "
                                                                   "peça remanejamento antes de comprar."))
            elif outra == "reservado":
                add(Alerta(ATENCAO, "Decreto 44.322, art. 7º, V", "Explique no TR por que o estoque das outras unidades não "
                                                                  "atende, como faz o TR do NLF Oeste (SEI 217882310)."))
            elif outra != "nao":
                add(Alerta(ATENCAO, "Decreto 44.322, art. 7º, V", "O decreto exige indisponibilidade em qualquer outra "
                                                                  "unidade da rede, não só nas centrais. Confira na Visão "
                                                                  "Gerencial."))
        limite = regras.get("idade_max_consulta_dias")
        consulta = _ler_data(pedido.get("data_consulta"))
        if not consulta:
            add(Alerta(ATENCAO, "Art. 16, II", "Informe a data do Nada Consta, da Visão Gerencial e da Declaração: os "
                                               "pareceres registram a data de emissão de cada uma."))
        elif consulta > hoje:
            add(Alerta(ATENCAO, "Art. 16, II", f"A data das consultas ({consulta:%d/%m/%Y}) está no futuro: confira o campo."))
        elif limite and (hoje - consulta).days > int(limite):
            add(Alerta(ATENCAO, "Art. 16, II", f"Consultas de {consulta:%d/%m/%Y}, há {(hoje - consulta).days} dias. "
                                               "Refaça para o processo seguir com declaração atual."))
        if item and categoria_num(item) == "0":
            cmm = _num(pedido.get("cmm"))       # como os demais campos numéricos: "1,5" não derruba a análise
            total = somar_quantidades(pedido)[0]
            if cmm is None:
                add(Alerta(ATENCAO, "Medicamento", "Informe o consumo médio mensal (CMM): sai no TR de medicamento."))
            elif total > cmm and not (pedido.get("just_acima_cmm") or "").strip():
                add(Alerta(ATENCAO, "Medicamento", f"Pedido de {total:g} acima do CMM de {cmm:g}: justifique. "
                                                   "Os NLF fazem isso em bloco próprio do TR (SEI 216270304)."))
    if not (pedido.get("justificativa") or "").strip():
        add(Alerta(ATENCAO, "Art. 21, III", "Falta a justificativa de relevância para a população."))

    if tipo == "servico":
        if pedido.get("altera_predio"):
            add(Alerta(ATENCAO, "Art. 20", "Reparo que altera o prédio exige anuência prévia da SINFRA."))
        if pedido.get("engenharia"):
            add(Alerta(INFO, "Art. 20", "Obra ou engenharia: o TR da GAOAPS-NO usa composições com código e quantidade por "
                                        "etapa, após visita técnica in loco (SEI 204620555)."))
        if pedido.get("reparo_equipamento"):
            bens = ler_patrimonios(pedido)
            if not bens:
                add(Alerta(BLOQUEIO, "Mem. Circ. 6/2023", "Reparo de equipamento: liste cada patrimônio com marca, modelo "
                                                          "e local."))
            elif any(not b["ficha"] for b in bens):
                add(Alerta(ATENCAO, "Mem. Circ. 6/2023", "Falta o SEI da ficha SisGepat de algum patrimônio. O NPDOC "
                                                         "fornece uma por equipamento."))
            if not pedido.get("sem_contrato_manutencao"):
                add(Alerta(ATENCAO, "Mem. Circ. 6/2023", "Falta a declaração do NECFM ou do executor de que não há contrato "
                                                         "de manutenção do equipamento na SES."))
            add(Alerta(INFO, "Parecer técnico", "Serviço em equipamento médico: a GEMED costuma analisar as propostas depois "
                                                "do mapa (SEI 218275490)."))

    if tipo:
        alertas.extend(_quantidades(pedido, regras))
        alertas.extend(_teto(pedido, regras, tipo))
        alertas.extend(_prazos(tipo, regras, hoje))
        if int(regras.get("ano") or hoje.year) != hoje.year:
            add(Alerta(ATENCAO, "Ano de referência", f"regras.json ainda está em {regras['ano']}: o teto, os prazos e o "
                                                     f"gasto do ano são de {regras['ano']}, não de {hoje.year}. Atualize "
                                                     "o arquivo (pasta Dados) quando sair o memorando da GEOF."))
        fer, ilegiveis = textos.ler_feriados(config_intencao(regras)["feriados"])
        if ilegiveis:
            add(Alerta(ATENCAO, "Prazo das gerências", "Feriado que não consegui ler em regras.json (intencao.feriados): "
                                                      f"{'; '.join(ilegiveis)}. Escreva como dd/mm/aaaa; até lá ele não é pulado "
                                                      "na contagem dos dias úteis."))
        if fer and max(fer) < hoje:
            add(Alerta(ATENCAO, "Prazo das gerências", f"A lista de feriados de regras.json termina em {max(fer):%d/%m/%Y}: o prazo "
                                                      "em dias úteis está sendo contado sem feriado. Atualize intencao.feriados."))
        dec = regras.get("decreto_encerramento")
        if dec:
            add(Alerta(ATENCAO, dec["numero"], dec["aviso"]))
    if tipo == "permanente" or (tipo and enq["grupo"] == "permanente"):
        sug = area_parecer(item, codigo, regras)
        onde = f" Provável área: {sug['area']} ({sug['motivo']}; {sug['fonte']}).{sug.get('nota', '')}" if sug else \
            " Área não identificada pelo código: confirme com a GEOF."
        add(Alerta(INFO, "Parecer técnico", "Permanente, mobiliário e equipamento exigem parecer favorável da Administração "
                                            f"Central (Decreto 44.322, art. 3º, parágrafo único), antes da cotação.{onde}"))
        if regras.get("aviso_capital"):
            add(Alerta(ATENCAO, "Art. 16, IV", regras["aviso_capital"]))
    if tipo == "permanente":
        if not str(pedido.get("gsap_origem") or "").strip():
            add(Alerta(ATENCAO, "Abertura", "Falta o setor de origem (etapa Texto): os memorandos saem com o campo a preencher."))
        if not str(pedido.get("processo_intencao") or "").strip():
            add(Alerta(INFO, "Abertura", "Sem o nº do processo do memorando de intenção, despachos, TR e memorando ao NPDOC e à "
                                         "GEOF saem com o campo a preencher. Informe na etapa Texto quando o processo existir."))
    if tipo == "consumo":
        add(Alerta(INFO, "Custeio", "Bem de consumo é custeio, fora do fluxo de bens permanentes da DIRAPS: em geral é executado "
                                    "pelo Núcleo de Logística Farmacêutica e pelo Núcleo de Almoxarifado (medicamentos e itens do dia a dia "
                                    "da UBS). Orientação da Diretoria, 17/11/2026."))
    elif tipo == "servico":
        add(Alerta(INFO, "Custeio", "Serviço é custeio, fora do fluxo de bens permanentes da DIRAPS: cobre pequenos reparos e "
                                    "manutenção de equipamentos SEM contrato de manutenção vigente, e em geral é executado pela Diretoria "
                                    "Administrativa. Orientação da Diretoria, 17/11/2026."))
    if tipo == "consumo" and regras["entrega"].get("consumo_obs"):
        add(Alerta(INFO, "Local de entrega", regras["entrega"]["consumo_obs"]))
    if tipo == "consumo" and not regras["entrega"].get("consumo"):
        add(Alerta(ATENCAO, "Local de entrega", "Endereço de entrega de consumo não cadastrado em regras.json. "
                                                "O TR sairá com campo a preencher."))

    pendentes = pedido.get("nao_validados") or []
    if pendentes:
        add(Alerta(ATENCAO, "Sugestões", f"{len(pendentes)} sugestões ainda sem validar. Elas saem em amarelo nas minutas."))
    ordem = {BLOQUEIO: 0, ATENCAO: 1, INFO: 2}
    alertas.sort(key=lambda a: ordem[a.severidade])
    return {
        "tipo": tipo,
        "codigo": codigo,
        "item": {k: v for k, v in item.items() if not k.startswith("_")} if item else None,
        "enquadramento": enq,
        "quantidade_total": somar_quantidades(pedido)[0],
        "valor_estimado": valor_estimado(pedido),
        "alertas": [asdict(a) for a in alertas],
        "pode_gerar": tipo is not None,
        "tem_bloqueio": any(a.severidade == BLOQUEIO for a in alertas),
    }


DUVIDA_ORDEM = ("Colegiado 23/11/2026: avaliar se a cotação e o mapa de preços (GEOF) devem vir antes do parecer central e da aprovação e autorização do Diretor Administrativo, logo depois do Nada Consta, porque a declaração de que há recurso depende do preço. Por enquanto, vale a ordem atual.")

FASES = ("1. Demanda", "2. Instrução", "3. Parecer central", "4. Autorização", "5. Cotação", "6. Escolha do fornecedor",
         "7. Contratação", "8. Recebimento e atesto", "9. Pagamento e contas")


def checklist(tipo: str, pedido: dict | None = None) -> list[dict]:
    """Sequência do Manual de 03/09/2024 conferida com árvores de processos de 2026 (SEI, 02/11/2026)."""
    pedido = pedido or {}
    destino = "NPDOC" if tipo == "permanente" else "NUAL"
    f = FASES
    C = "Colegiado 23/11/2026"
    L = [
        (f[0], "Setor de origem", ("Memorando de intenção à sua diretoria, em processo PÚBLICO: serviço, necessidade e bens ou "
                                  "locais atendidos" if tipo == "servico" else
                                  "Memorando de intenção à sua diretoria, em processo PÚBLICO: item localizado no catálogo SES "
                                  "(link da SES), necessidade e quantidade"), C + "; Rotina DIRAPS Norte (17/11/2026)"),
        (f[0], "Diretoria do setor de origem", "Despacho às gerências de TODA a região (incluindo as que historicamente ficam de "
                                               "fora), com prazo de 7 dias úteis para dizerem se têm interesse",
         "Rotina DIRAPS Norte (Diretoria, 17/11/2026)"),
        (f[0], "Assessoria da diretoria", "Registrar na planilha de acompanhamento o nº do processo do memorando de intenção e o item",
         "Rotina DIRAPS Norte (Diretoria, 17/11/2026)"),
        (f[0], "Setor de origem", "Ao fim dos 7 dias úteis, consolidar o interesse das gerências e registrar o consolidado no "
                                  "processo do memorando", C + "; Mem. Circ. 2/2024, item 1"),
        (f[0], "Setor de origem", "Termo de Referência em processo SEI NOVO, PÚBLICO, sem documento restrito e relacionado ao "
                                  "processo do memorando de intenção", C),
        (f[0], "Setor de origem", "Memorando com justificativa (ou DFD)", "Manual, atividade 1"),
        (f[1], "Setor de origem", "Conferir no TR a fonte pagadora: razão social e CNPJ da APS Norte (00.394.700/0050-96)",
         "Rotina DIRAPS Norte (Diretoria, 18/11/2026)"),
    ]
    if tipo == "permanente":
        R = "Rotina DIRAPS Norte (Diretoria, 18/11/2026)"
        L += [
            (f[1], "Gestores interessados", "TR assinado pelos gestores de TODAS as gerências que manifestaram interesse no "
                                            "memorando de intenção", R),
            (f[1], "Diretoria do setor de origem", "Bloco de assinatura da diretoria: revisão de pendências e assinatura do TR",
             "Art. 16, I; " + R),
        ]
    else:
        L.append((f[1], "Diretoria do setor de origem", "Bloco de assinatura da diretoria: assinatura do TR", "Art. 16, I"))
    L.append((f[1], "SRSNO", "Bloco de assinatura da Superintendência: revisão e assinatura do TR; de lá, o processo segue "
                             + ("à GEOF" if tipo == "servico" else f"ao {destino} e à GEOF"), "Art. 16, I; " + C))
    if tipo == "servico":
        if pedido.get("reparo_equipamento"):
            L += [(f[1], "NPDOC", "Ficha SisGepat de cada equipamento (Carga Geral - Bens Móveis)", "Mem. Circ. 6/2023"),
                  (f[1], "NECFM ou executor", "Declaração de inexistência de contrato de manutenção", "Mem. Circ. 6/2023")]
        if pedido.get("visita_tecnica"):
            L.append((f[1], "Setor de origem", "Modelo do Comprovante de Visita Técnica para os orçamentos", "TR 95/2026 SRSSO"))
        if pedido.get("altera_predio"):
            L.append((f[1], "SINFRA", "Anuência prévia da SINFRA", "Art. 20"))
    else:
        L += [
            (f[1], destino, "Nada Consta do sistema de materiais, com data", "Art. 16, II"),
            (f[1], destino, "Declaração de existência ou inexistência de ata de registro de preços, ou justificativa", "Art. 16, III"),
            (f[1], destino, "Relatório completo da Visão Gerencial", "Manual, atividade 1"),
        ]
        if pedido.get("ata") == "vigente":
            L.append((f[1], "Setor de origem", "Ata vigente: o processo para aqui, porque a SES já adquire o item pela ata",
                      "Art. 21, II; Colegiado 23/11/2026"))
    if tipo == "permanente":
        L.append((f[2], "Área central da SES", "Parecer técnico favorável, antes da cotação (ordem em discussão: ver a "
                                               "dúvida em aberto)",
                  "Decreto 44.322, art. 3º, p. único; Manual, atividade 2"))
    L += [
        (f[3], "Diretor Administrativo", "Aprovação do TR por despacho motivado", "Decreto 44.330, art. 71, § 3º"),
        (f[3], "Diretor Administrativo", "Declaração do ordenador: disponibilidade financeira", "Art. 16, IV"),
        (f[3], "Diretor Administrativo", "Autorização da aquisição", "Art. 16, V"),
        (f[4], "GEOF", "Pedido de Cotação no sistema e convocação por e-mail do SEI", "Manual, atividade 6"),
        (f[4], "GEOF", "Propostas em 3 dias úteis; registrar os consultados que não responderam", "Decreto 44.330, art. 93, § 6º"),
        (f[4], "GEOF", "Pesquisa de Preços: ao menos 3 preços válidos, mediana e memória de cálculo", "Decreto 44.330, arts. 90 e 93"),
        (f[4], "GEOF", "Mapa Comparativo de Preços", "Manual, atividade 7"),
        (f[5], "Setor de origem", "Parecer com a empresa vencedora; justificar se não for o menor preço", "Art. 16, VI"),
    ]
    if tipo == "permanente":
        L += [(f[5], "Setor de origem", "Avaliar no mínimo 3 empresas e escolher a que atende aos requisitos do item com o melhor "
                                   "preço; devolver o processo à GEOF, que acerta a entrega com a empresa",
               "Rotina DIRAPS Norte (Diretoria, 18/11/2026)")]
    if tipo == "servico" and pedido.get("reparo_equipamento"):
        L.append((f[5], "SINFRA/GEMED", "Análise das propostas, em equipamento médico", "Parecer 490/2026"))
    L += [
        (f[5], "GEOF", "Certidões: Receita e PGFN, GDF, Trabalhista, FGTS, CGU", "Manual, atividade 7"),
        (f[5], "GEOF", "Consulta às empresas suspensas ou impedidas", "Art. 16, IX"),
        (f[6], "Diretor Administrativo", "Termo de Adjudicação e Homologação", "Manual, atividade 9"),
        (f[6], "Diretor Administrativo", "Ordem de Serviço - PDPAS", "Manual, atividade 9"),
        (f[6], "UGE", "Publicação no PNCP e no DODF em até 10 dias úteis", "Art. 16, X"),
        (f[6], "GEOF", "Envio da Ordem ao fornecedor por e-mail do SEI", "Manual, atividade 10"),
    ]
    if tipo == "permanente":
        L += [
            (f[7], "NPDOC", "Recebimento no sistema, Nota Fiscal e Termo de Recebimento Provisório", "Manual, atividade 13"),
            (f[7], "Setor de origem", "Ir ao NPDOC, avaliar o bem contra o TR e fazer o Atesto de Recebimento Definitivo com fotos "
                                 "dos bens e servidor responsável", "Art. 36, XIII; Mem. Circ. 1/2025"),
            (f[7], "NPDOC", "Distribuir às gerências conforme o consolidado do processo do memorando de intenção",
             "Rotina DIRAPS Norte (Diretoria, 18/11/2026)"),
        ]
    elif tipo == "consumo":
        L.append((f[7], destino, "Nota Fiscal no processo logo após o recebimento", "Mem. Circ. 1/2025"))
        L.append((f[7], "Setor de origem", "Atesto da Nota Fiscal com a declaração do PDPAS", "Mem. Circ. 1/2025"))
    else:
        L.append((f[7], "Responsável pelo aceite", "Nota Fiscal atestada após a execução total", "Mem. Circ. 1/2025"))
    L += [(f[8], "GEOF", "Pagamento e comprovante", "Art. 16, VII e VIII"),
          (f[8], "GEOF", "Prestação de contas bimestral até o dia 10 e anual até fevereiro; sem ela o repasse é suspenso",
           "Art. 34, § 1º; Decreto 44.322, art. 11")]
    if tipo == "permanente":
        L.append((f[8], "NPDOC", "Tombamento, Termo de Movimentação e relatório de incorporação no SisGepat",
                  "Art. 28, § 1º; Art. 36, XIII"))
    return [{"fase": a, "responsavel": b, "peca": c, "base": d} for a, b, c, d in L]
