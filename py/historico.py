"""Histórico de compras e consumo da APS e leitura de documentos do SEI e do sistema de materiais.

O histórico mora em um único arquivo JSON, sem dado de paciente nem nome de servidor. Ele nasce com a carga
embutida no programa, acompanha a pasta do exe e pode ser levado para outra máquina: importar o arquivo de
outra máquina soma o que falta e não duplica nada.

Lê, em PDF ou texto: Relação de Estoque Físico da GEOF, Visão Gerencial, Nada Consta e o PDF do processo inteiro.
"""
from __future__ import annotations

import io
import json
import os
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path

VERSAO = 1
MESES = ("JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO",
         "NOVEMBRO", "DEZEMBRO")
PROCESSO = r"\d{5}-\d{8}/\d{4}-\d{2}"
CENTRAIS = ("ALMOXARIFADO CENTRAL", "FARMACIA CENTRAL", "NUMEBE", "NUDIME", "FARMACIA ODONTOLOGIA", "FARMACIA AMBULATORIAL",
            "PATRIMÔNIO CENTRAL", "BÁSICOS E ESTRATÉGICOS")


def vazio() -> dict:
    return {"versao": VERSAO, "atualizado_em": "", "compras": {}, "consumo": {}, "visao": {}, "importacoes": [], "textos": {}}


def numero(txt) -> float | None:
    """Aceita 1.200,50 / 1200.50 / 199.800 / 0,0797, formatos que convivem nas relações da GEOF."""
    s = re.sub(r"[R$\s]", "", str(txt or ""))
    if not re.fullmatch(r"-?[\d.,]+", s) or not re.search(r"\d", s):
        return None
    if "," in s and "." in s and s.rfind(".") > s.rfind(","):
        s = s.replace(",", "")                  # formato americano (1,234.56): antes virava 1.23456
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s) and not re.fullmatch(r"-?0\.\d+", s):
        s = s.replace(".", "")    # ponto de milhar ("199.800"); "0.797" continua decimal
    try:
        return float(s)
    except ValueError:
        return None


def _data_iso(txt: str) -> str:
    try:
        return datetime.strptime(txt.strip(), "%d/%m/%Y").date().isoformat()
    except ValueError:
        return ""


# ---------- Relação de Estoque Físico ----------

def _meses_da_relacao(celulas: list[str]) -> tuple[list[str], list[str]]:
    """Meses de entradas e de saídas. A GEOF usa 2 meses por coluna, mas já usou 4 e 5 no mesmo quadro."""
    texto = " ".join(celulas).upper()
    fim = re.search(r"SALDO\s+AT[ÉE]\s+([A-ZÇ]+)[\s/.-]*(?:DE\s+)?(\d{4})", texto)      # "ATÉ AGOSTO 2026", "ATÉ AGOSTO/2026", "ATE"
    if not fim or fim[1] not in MESES:
        return [], []
    mes_fim, ano_fim = MESES.index(fim[1]) + 1, int(fim[2])
    nomes = []
    for c in celulas:
        m = re.fullmatch(r"([A-Za-zçÇ]+)\s*(?:/\s*(\d{4}))?", c.strip())
        if m and m[1].upper() in MESES:
            nomes.append(MESES.index(m[1].upper()) + 1)
        elif nomes:
            break
    # o bloco de saídas começa onde o mês deixa de ser o sucessor do anterior: "NOV DEZ JAN FEV | NOV DEZ JAN FEV"
    # (a regra antiga, "mês <= anterior", cortava em JAN e atribuía o consumo aos meses errados sem aviso)
    corte = next((i for i in range(1, len(nomes)) if nomes[i] != nomes[i - 1] % 12 + 1), len(nomes))
    rotulo = lambda m: f"{ano_fim - 1 if m > mes_fim else ano_fim}-{m:02d}"
    return [rotulo(m) for m in nomes[:corte]], [rotulo(m) for m in nomes[corte:]]


def ler_relacao(texto: str, origem: str = "") -> dict:
    """Cada compra vira uma linha; as saídas de cada mês alimentam o consumo por código.

    Itens da mesma nota vêm em linhas mescladas: a segunda linha começa direto no código e herda a NF.
    """
    celulas = [c.strip() for c in re.split(r"[\t\r\n]+", texto) if c.strip()]
    entradas, saidas = _meses_da_relacao(celulas)
    tamanho = 2 + len(entradas) + len(saidas)
    data = r"\d{2}/\d{2}/\d{4}"
    n = len(celulas)
    inicio_linha = lambda x: x + 3 < n and re.fullmatch(data, celulas[x]) and re.fullmatch(data, celulas[x + 3]) \
        and len(celulas[x + 1]) <= 12
    compras, consumo, descartadas = [], [], 0

    def item(pos: int, nf: str, nf_data: str) -> int | None:
        """Lê do código até os saldos. Devolve a posição seguinte, ou None se não for uma linha de compra."""
        nonlocal descartadas
        k = pos + 1
        while k < n and not re.fullmatch(PROCESSO, celulas[k]) and not celulas[k].startswith("16.") and not inicio_linha(k):
            k += 1
        if k >= n or not re.fullmatch(PROCESSO, celulas[k]) or k - 3 < pos + 1:
            return None
        qtd, unit, total = numero(celulas[k - 3]), numero(celulas[k - 2]), numero(celulas[k - 1])
        if total is None:
            descartadas += 1
            return k + 1
        descricao = " ".join(celulas[pos + 1:k - 3])
        j = k + 1
        while j < min(k + 8, n) and not re.search(r"\d{2}\.?\d{3}\.?\d{3}.*\d{4}-?\d{2}$", celulas[j]):
            j += 1
        fornecedor = " ".join(celulas[k + 1:j])
        fim_nums = j + 1
        while fim_nums < min(j + 1 + tamanho, n) and numero(celulas[fim_nums]) is not None and not inicio_linha(fim_nums + 1):
            fim_nums += 1
        nums = [numero(c) for c in celulas[j + 1:fim_nums]]
        codigo = celulas[pos].upper()
        alias = re.match(r"(\d{3,6})\s*-", descricao)
        registro = {"codigo": codigo, "descricao": re.sub(r"^\d{3,6}\s*-\s*", "", descricao)[:120], "nf": nf,
                    "nf_data": _data_iso(nf_data), "qtd": qtd, "unit": unit, "total": total, "processo": celulas[k],
                    "fornecedor": fornecedor[:70], "origem": origem}
        if alias and alias[1] != codigo:
            registro["alias"] = alias[1]
        compras.append(registro)
        if saidas and len(nums) == tamanho:
            for mes, valor in zip(saidas, nums[1 + len(entradas):-1], strict=True):
                consumo.append((codigo, mes, valor))
        return fim_nums

    i = 0
    while i + 5 < n:
        if not inicio_linha(i):
            i += 1
            continue
        nf, nf_data = celulas[i + 2], celulas[i + 3]
        pos = item(i + 4, nf, nf_data)
        if pos is None:
            descartadas += 1
            i += 1
            continue
        while pos < n and re.fullmatch(r"[A-Za-z]{0,3}\d{2,12}", celulas[pos]) and not inicio_linha(pos) \
                and not inicio_linha(pos + 1):
            seguinte = item(pos, nf, nf_data)
            if seguinte is None:
                break
            pos = seguinte
        i = pos
    return {"compras": compras, "consumo": consumo, "meses": saidas, "descartadas": descartadas}


# ---------- Visão Gerencial e Nada Consta ----------

def _trecho(texto: str, inicio: str, fim: str) -> str:
    m = re.search(inicio, texto)
    if not m:
        return ""
    n = re.search(fim, texto[m.end():])
    return texto[m.end():m.end() + n.start()] if n else texto[m.end():]


def _bloco_visao(t: str, sei: str = "") -> dict | None:
    p = re.search(r"PRODUTO:\s*(\S+)\s*-\s*(.+)", t)
    if not p:
        return None
    pegar = lambda rx: (re.search(rx, t) or [None, ""])[1].strip()
    locais = _trecho(t, r"Locais com Maior Disponibilidade", r"INDICADORES DE REPOSI")
    outros, central_itens = [], []
    for m in re.finditer(r"(?:^|\s{2,})(\d[\d.]*)\s+([A-ZÀ-Ú][^\n]*?)\s{2,}([\d.,]+)\s+([\d.,]+)(?=\s|$)", locais, re.M):
        nome = m[2].strip()
        par = {"local": nome, "disponivel": numero(m[3]) or 0}
        (central_itens if any(c in nome.upper() for c in CENTRAIS) else outros).append(par)
    total = re.search(r"Total nos Locais Centrais:\s*([\d.,]+)\s+([\d.,]+)", t)
    atas_txt = _trecho(t, r"Atas de Registro de Preço vigentes", r"Histórico de Compras Regular")
    atas = []
    sem_ata = "Nenhuma ata em vigência" in atas_txt
    if not sem_ata:
        for lin in atas_txt.splitlines():
            m = re.search(r"([\d.,]+)\s+([\d.,]+)\s+R\$\s*([\d.,]+)\s+(\d{2}/\d{2}/\d{4})", lin)
            if m:
                cab = re.match(r"\s*\d+\s+(\S+)\s+(\S+)", lin)
                atas.append({"ata": cab[1] if cab else "", "pregao": cab[2] if cab else "", "quantidade": numero(m[1]),
                             "saldo": numero(m[2]), "preco": numero(m[3]), "validade": m[4],
                             "fornecedor": re.sub(r"\s+", " ", lin[cab.end():m.start()]).strip(" 0123456789") if cab else ""})
    abertos_txt = _trecho(t, r"Pedidos e Compras em Aberto", r"INDICADORES DE CONSUMO")
    abertos = []
    for m in re.finditer(rf"({PROCESSO})([^\n]*)", abertos_txt):
        resto = m[2]
        afm = re.search(r"\S*AFM\S*", resto)
        emp = re.search(r"\d{4}NE\d+", resto)
        abertos.append({"processo": m[1], "afm": afm[0] if afm else "", "empenho": emp[0] if emp else ""})
    precos = []
    for m in re.finditer(r"^\s*\d+\s+(\d{2}/\d{2}/\d{4})\s+([\d.,]+)\s+\S+\s+R\$\s*([\d.,]+)", _trecho(t, "Histórico de Preços", "Atas de Registro"), re.M):
        precos.append({"data": _data_iso(m[1]), "qtd": numero(m[2]), "preco": numero(m[3])})
    consumo_locais = []
    for m in re.finditer(r"^\s*(\d[\d.]*)\s+([A-ZÀ-Ú][^\n]*?)\s{2,}([\d.,]+)\s*$", _trecho(t, "Locais de Maior Consumo", "INDICADORES DE MOVIMENTA"), re.M):
        consumo_locais.append({"local": m[2].strip(), "consumo": numero(m[3])})
    meses = {f"{m[1]}-{int(m[2]):02d}": {"mov": numero(m[3]), "cmm": numero(m[4])}
             for m in re.finditer(r"(\d{4})/(\d{1,2})\s+([\d.,]+)\s+([\d.,]+)", t)}
    impresso = _data_iso(pegar(r"Impresso em:\s*(\d{2}/\d{2}/\d{4})"))
    return {"codigo": p[1].upper(), "descricao": p[2].strip()[:120], "data": impresso, "sei": sei,
            "codigo_br": pegar(r"Código BR:\s*(\d+)"), "elenco": pegar(r"Pertence ao Elenco\?\s*(Sim|Não)"),
            "custo_medio": numero(pegar(r"Custo Médio:\s*R\$\s*([\d.,]+)")), "estoque_rede": numero(pegar(r"Estoque na rede:\s*([\d.,]+)")),
            "cmm_rede": numero(pegar(r"Consumo Médio Mensal:\s*([\d.,]+)")),
            # None = o total não foi encontrado no texto (não é o mesmo que zero): a janela pede confirmação
            "central": numero(total[1]) if total else None,
            "cobertura_dias": numero(pegar(r"Cobertura de Estoque:\s*(\d+) dias")), "outras_unidades": outros,
            # atas_incerto = o quadro de atas existe, não diz "Nenhuma ata" e nenhuma linha foi lida (linha quebrada, PDF girado)
            "atas": atas, "atas_incerto": bool(atas_txt.strip()) and not sem_ata and not atas,
            "abertos": abertos, "precos": precos, "consumo_locais": consumo_locais, "meses": meses}


def ler_visoes(texto: str) -> list[dict]:
    """Um PDF de processo pode trazer várias Visões; devolve uma por produto, a mais recente."""
    achados: dict[str, dict] = {}
    partes = re.split(r"(?=VISÃO GERENCIAL DO PRODUTO)", texto)
    for parte in partes[1:]:
        sei = (re.search(r"Visão Gerencial[^\n(]*\((\d{8,9})\)", parte) or [None, ""])[1]
        v = _bloco_visao(parte, sei)
        if v and (v["codigo"] not in achados or v["data"] >= achados[v["codigo"]]["data"]):
            achados[v["codigo"]] = v
    return list(achados.values())


def ler_nada_consta(texto: str) -> list[dict]:
    saida = []
    for parte in re.split(r"(?=DECLARAÇÃO DE NADA CONSTA)", texto)[1:]:
        cod = re.search(r"Código do Produto:\s*(\S+)", parte)
        data = re.search(r"na data de\s*(\d{2}/\d{2}/\d{4})", parte)
        sei = re.search(r"Nada Consta[^\n(]*\((\d{8,9})\)", parte)
        if cod:
            saida.append({"codigo": cod[1].upper(), "data": _data_iso(data[1]) if data else "", "sei": sei[1] if sei else ""})
    return saida


def extrair_texto(nome: str, dados: bytes) -> str:
    if dados[:5] == b"%PDF-":
        from pypdf import PdfReader
        leitor = PdfReader(io.BytesIO(dados))

        def _pagina(p):
            try:
                return p.extract_text(extraction_mode="layout", layout_mode_strip_rotated=False) or ""
            except Exception:  # noqa: BLE001 — uma página girada/quebrada não deve perder as demais
                return ""
        return "\n".join(_pagina(p) for p in leitor.pages)
    texto = None
    for cod in ("utf-8-sig", "cp1252"):
        try:
            texto = dados.decode(cod)
            break
        except UnicodeDecodeError:
            continue
    texto = dados.decode("latin-1") if texto is None else texto
    if re.search(r"<(html|table|body)", texto[:5000], re.I):
        texto = re.sub(r"(?is)<(script|style).*?</\1>", " ", texto)
        texto = re.sub(r"(?i)</(td|tr|p|div|th)>|<br\s*/?>", "\n", texto)
        texto = re.sub(r"<[^>]+>", " ", texto).replace("&nbsp;", " ")
    return texto


# ---------- Histórico ----------

def _chave_compra(c: dict) -> str:
    return f"{c['codigo']}|{c['nf']}|{c['processo']}|{c['total']:.2f}"


_OBRIGATORIAS = {"codigo", "nf", "nf_data", "total", "processo"}


def _chave_norm(c: dict) -> str:
    """Mesma compra com a NF em grafias diferentes ("5.190" / "5190" / "005190") não pode entrar duas vezes.
    _chave_compra fica como está (as chaves já gravadas precisam continuar casando); aqui só se compara o conteúdo."""
    nf = re.sub(r"\D", "", str(c.get("nf") or "")).lstrip("0") or str(c.get("nf") or "").strip().upper()
    return f"{str(c.get('codigo') or '').strip().upper()}|{nf}|{str(c.get('processo') or '').strip()}|{float(c.get('total') or 0):.2f}"


def _compra_valida(c) -> bool:
    """Tem os campos obrigatórios, total numérico e data da nota em texto (as consultas ordenam e fatiam essa data)."""
    if not isinstance(c, dict) or not c.keys() >= _OBRIGATORIAS:
        return False
    total = c.get("total")
    if isinstance(total, bool) or not isinstance(total, (int, float)) or total != total:
        return False
    return isinstance(c.get("nf_data"), str)


def fundir(base: dict, outro: dict) -> tuple[dict, int]:
    """Soma o que falta em base. Rodar de novo com o mesmo arquivo não muda nada."""
    novos = 0
    vistos_c = {_chave_norm(c) for c in base["compras"].values() if isinstance(c, dict)}
    for k, c in (outro.get("compras") or {}).items():
        # entrada incompleta ou de tipo errado (JSON de outra máquina) quebraria as consultas em toda abertura: fica de fora
        if not _compra_valida(c):
            continue
        if k not in base["compras"] and _chave_norm(c) not in vistos_c:
            base["compras"][k] = c
            vistos_c.add(_chave_norm(c))
            novos += 1
    for cod, meses in (outro.get("consumo") or {}).items():
        if not isinstance(meses, dict):
            continue
        destino = base["consumo"].setdefault(cod, {})
        for mes, valor in meses.items():
            if not isinstance(valor, (int, float)):
                continue
            if mes not in destino:
                novos += 1
            destino[mes] = max(destino.get(mes, 0), valor)
    for cod, v in (outro.get("visao") or {}).items():
        if not isinstance(v, dict):
            continue
        if cod not in base["visao"] or v.get("data", "") > base["visao"][cod].get("data", ""):
            base["visao"][cod] = v
            novos += 1
    for cod, campos in (outro.get("textos") or {}).items():
        if not isinstance(campos, dict):
            continue
        destino = base.setdefault("textos", {}).setdefault(cod, {})
        for campo, v in campos.items():
            if not isinstance(v, dict):
                continue
            if campo not in destino or v.get("data", "") > destino[campo].get("data", ""):
                destino[campo] = v
                novos += 1
    vistos = {json.dumps(i, sort_keys=True) for i in base["importacoes"]}
    base["importacoes"] += [i for i in outro.get("importacoes") or [] if json.dumps(i, sort_keys=True) not in vistos]
    return base, novos


def carregar(arquivo: Path) -> dict:
    """Nunca descarta um arquivo em silêncio: ilegível ou de outra versão é posto de lado com outro nome."""
    arquivo = Path(arquivo)
    if not arquivo.exists():
        return vazio()
    # OSError (arquivo em uso, OneDrive "somente online") SOBE para quem chamou: não é defeito do conteúdo, e pôr de lado
    # o histórico bom da pasta compartilhada fazia a gravação seguinte sobrescrevê-lo só com o conteúdo local
    texto = arquivo.read_text(encoding="utf-8-sig")
    try:
        dados = json.loads(texto)
        if not isinstance(dados, dict) or not isinstance(dados.get("compras"), dict):
            raise ValueError("não é um histórico")
    except ValueError as e:
        _por_de_lado(arquivo, "ilegivel")
        AVISOS.append(f"Histórico ilegível em {arquivo.name} ({str(e)[:80]}): guardado ao lado com outro nome e ignorado.")
        return vazio()
    if dados.get("versao") != VERSAO:
        _por_de_lado(arquivo, f"versao{dados.get('versao')}")
        AVISOS.append(f"Histórico de outra versão em {arquivo.name}: guardado ao lado com outro nome e ignorado.")
        return vazio()
    for chave, padrao in vazio().items():       # arquivo de versão certa, mas sem alguma chave, derrubava as sugestões
        dados.setdefault(chave, padrao)
    return dados


AVISOS: list[str] = []      # o app mostra na janela o que foi posto de lado


def _por_de_lado(arquivo: Path, motivo: str) -> None:
    try:
        arquivo.replace(arquivo.with_name(f"{arquivo.stem}.{motivo}-{datetime.now():%Y%m%d-%H%M%S}.json"))
    except OSError:
        pass


def salvar(h: dict, arquivo: Path) -> None:
    import time
    arquivo = Path(arquivo)
    arquivo.parent.mkdir(parents=True, exist_ok=True)
    h["atualizado_em"] = datetime.now().isoformat(timespec="seconds")
    temporario = arquivo.with_name(f"{arquivo.stem}.{os.getpid()}.tmp")    # nome próprio: duas instâncias não disputam
    try:
        with open(temporario, "w", encoding="utf-8") as f:
            f.write(json.dumps(h, ensure_ascii=False, indent=0))
            f.flush()
            os.fsync(f.fileno())        # suspensão/queda logo depois não deixa o histórico com 0 bytes
    except BaseException:
        temporario.unlink(missing_ok=True)      # sem .tmp órfão (no compartilhado ele seria sincronizado)
        raise
    for tentativa in range(3):      # o OneDrive pode segurar o arquivo por um instante
        try:
            temporario.replace(arquivo)
            return
        except PermissionError:
            if tentativa == 2:
                temporario.unlink(missing_ok=True)
                raise
            time.sleep(0.5)


def importar_texto(h: dict, texto: str, origem: str) -> dict:
    """Reconhece o tipo pelo conteúdo. Devolve o resumo e o que foi lido para preencher a janela."""
    resumo = {"origem": origem, "tipo": [], "novos": 0, "visoes": [], "nada_consta": [], "descartadas": 0}
    texto = unicodedata.normalize("NFC", texto)     # PDF extraído com acento decomposto não casava "FÍSICO" nem "ATÉ"
    if "ESTOQUE FÍSICO" in texto.upper():
        r = ler_relacao(texto, origem)
        lote = vazio()
        for c in r["compras"]:
            lote["compras"][_chave_compra(c)] = c
        for cod, mes, valor in r["consumo"]:
            lote["consumo"].setdefault(cod, {})
            lote["consumo"][cod][mes] = max(lote["consumo"][cod].get(mes, 0), valor)
        _, n = fundir(h, lote)
        resumo["tipo"].append(f"Relação de Estoque Físico ({len(r['compras'])} compras)"
                              + (", sem meses de saída reconhecidos" if r["compras"] and not r["consumo"] else ""))
        resumo["novos"] += n
        resumo["descartadas"] += r["descartadas"]
    visoes = ler_visoes(texto)
    if visoes:
        lote = vazio()
        lote["visao"] = {v["codigo"]: v for v in visoes}
        _, n = fundir(h, lote)
        resumo["tipo"].append(f"Visão Gerencial ({len(visoes)})")
        resumo["novos"] += n
        resumo["visoes"] = visoes
    nadas = ler_nada_consta(texto)
    if nadas:
        resumo["tipo"].append(f"Nada Consta ({len(nadas)})")
        resumo["nada_consta"] = nadas
    if resumo["tipo"]:
        h["importacoes"].append({"data": date.today().isoformat(), "origem": origem, "tipo": ", ".join(resumo["tipo"])})
    return resumo


def importar_arquivo(h: dict, nome: str, dados: bytes) -> dict:
    if nome.lower().endswith(".json"):
        outro = json.loads(dados.decode("utf-8-sig"))       # histórico salvo com BOM por outro programa
        if not isinstance(outro, dict) or outro.get("versao") != VERSAO or not isinstance(outro.get("compras"), dict):
            raise ValueError(f"{nome} não é um histórico da Ferramenta PDPAS")
        _, n = fundir(h, outro)
        return {"origem": nome, "tipo": ["Histórico de outra máquina"], "novos": n, "visoes": [], "nada_consta": [], "descartadas": 0}
    return importar_texto(h, extrair_texto(nome, dados), nome)


# ---------- Consultas ----------

def compras_do_item(h: dict, codigo: str) -> list[dict]:
    cod = str(codigo or "").strip().upper()
    achados = [c for c in h["compras"].values() if cod and (c["codigo"] == cod or c.get("alias") == cod)]
    return sorted(achados, key=lambda c: c["nf_data"], reverse=True)


def gasto_ano(h: dict, codigo: str, ano: int) -> dict:
    lista = [c for c in compras_do_item(h, codigo) if c["nf_data"][:4] == str(ano)]
    ate = max((c["nf_data"] for c in h["compras"].values()), default="")
    return {"total": round(sum(c["total"] for c in lista), 2), "processos": sorted({c["processo"] for c in lista}), "ate": ate}


def cmm(h: dict, codigo: str, hoje: date | None = None) -> dict | None:
    """Média das saídas nos últimos 12 meses, só dos meses com saída, como fazem os TRs dos NLF."""
    hoje = hoje or date.today()
    meses = h["consumo"].get(str(codigo or "").strip().upper()) or {}
    corte = f"{hoje.year - 1}-{hoje.month:02d}"      # exclusivo: 12 meses, do mês seguinte ao de um ano atrás até hoje
    validos = {m: v for m, v in meses.items() if m > corte and v and v > 0}
    if not validos:
        return None
    media = sum(validos.values()) / len(validos)
    valor = int(media + 0.5)        # meio para cima, como a planilha da GEOF (round() do Python faz 2,5 -> 2)
    if valor == 0:
        return None                 # média abaixo de 0,5 não serve de CMM e evitaria divisão por zero
    return {"valor": valor, "meses": len(validos), "de": min(validos), "ate": max(validos)}
