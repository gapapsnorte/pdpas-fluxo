"""Planilha única de controle das compras do PDPAS: uma linha por pedido feito na ferramenta e uma por compra já
registrada no histórico (Relações de Estoque Físico do SEI). Uma aba só, sem células mescladas."""
from __future__ import annotations

import csv
import io
import math
from pathlib import Path

import motor

NOME = "Controle de compras PDPAS.xlsx"
COLUNAS = ["Situação", "Origem", "Data", "Código SES", "Item", "Tipo", "Quantidade", "Valor (R$)", "Fornecedor", "Nota fiscal",
           "Processo SEI", "Números SEI anotados", "Fase atual", "Passos feitos", "Unidades", "Minutas geradas", "Atualizado em",
           "Identificação"]
LARGURAS = [16, 30, 12, 12, 48, 13, 12, 14, 34, 12, 24, 28, 24, 10, 40, 10, 18, 26]
TIPOS = {"permanente": "Permanente", "consumo": "Consumo", "servico": "Serviço"}
_INICIO_DE_FORMULA = ("=", "+", "-", "@", "\t", "\r")
ORIGEM_MANUAL = "Lançado à mão"     # na coluna Origem: compra feita fora da ferramenta, digitada na planilha (24/11/2026)


def _sem_acento(t) -> str:
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", str(t or "")) if unicodedata.category(c) != "Mn").lower().strip()


def eh_manual(origem) -> bool:
    return _sem_acento(origem).startswith("lancado a mao")


def _numero(v) -> float | None:
    """1.234,56 | 1234,56 | 1234.56 | R$ 1.234,56 | número do Excel."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v) if math.isfinite(v) else None
    s = str(v or "").replace("R$", "").replace("\xa0", "").replace(" ", "").strip()
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        x = float(s)
    except ValueError:
        return None
    return x if math.isfinite(x) else None


def _ano(v) -> int | None:
    """Ano de dd/mm/aaaa, aaaa-mm-dd ou da data serial do Excel."""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 20000 < v < 80000:
        from datetime import date, timedelta
        return (date(1899, 12, 30) + timedelta(days=int(v))).year
    s = str(v or "").strip()
    for a, b in ((6, 10), (0, 4)):
        if len(s) >= b and s[a:b].isdigit() and ((a == 0 and s[4:5] == "-") or (a == 6 and s[2:3] == "/")):
            return int(s[a:b])
    return None


def ler_tabela(nome: str, dados: bytes) -> list[list]:
    """Linhas da primeira aba de um .xlsx, ou de um .csv (ponto e vírgula, vírgula ou tabulação). Sem openpyxl:
    roda igual no computador e no site."""
    if str(nome).lower().endswith(".xlsx"):
        import re
        import zipfile
        from xml.etree import ElementTree as ET
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        try:
            z = zipfile.ZipFile(io.BytesIO(dados))
            nomes = z.namelist()
            comp = []
            if "xl/sharedStrings.xml" in nomes:
                for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(ns + "si"):
                    comp.append("".join(x.text or "" for x in si.iter(ns + "t")))
            folhas = sorted(n for n in nomes if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n))
            raiz = ET.fromstring(z.read("xl/worksheets/sheet1.xml" if "xl/worksheets/sheet1.xml" in nomes else folhas[0]))
        except (zipfile.BadZipFile, KeyError, IndexError, ET.ParseError) as e:
            raise ValueError("Não consegui ler esta planilha .xlsx. Salve de novo no Excel ou exporte como .csv.") from e
        linhas = []
        for row in raiz.iter(ns + "row"):
            ln = {}
            for c in row.iter(ns + "c"):
                ref, tipo = c.get("r") or "", c.get("t")
                col = 0
                for ch in re.match(r"[A-Z]*", ref).group(0):
                    col = col * 26 + ord(ch) - 64
                v = c.find(ns + "v")
                if tipo == "s" and v is not None:
                    val = comp[int(v.text)] if int(v.text) < len(comp) else ""
                elif tipo == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter(ns + "t"))
                elif v is None:
                    val = ""
                elif tipo in ("str", "e"):
                    val = v.text or ""
                else:
                    try:
                        val = float(v.text)
                    except (TypeError, ValueError):
                        val = v.text or ""
                ln[max(col, 1) - 1] = val
            linhas.append([ln.get(i, "") for i in range(max(ln) + 1)] if ln else [])
        return linhas
    for cod in ("utf-8-sig", "cp1252"):
        try:
            texto = dados.decode(cod)
            break
        except UnicodeDecodeError:
            continue
    primeira = texto.split("\n", 1)[0]
    sep = max((";", ",", "\t"), key=primeira.count)
    return [ln for ln in csv.reader(io.StringIO(texto), delimiter=sep)]


_CAMPOS = {"codigo": ("codigo ses", "codigo", "cod. ses", "codigo do item"), "valor": ("valor (r$)", "valor total", "valor"),
           "data": ("data", "data da nf", "data da nota"), "situacao": ("situacao",), "origem": ("origem",),
           "processo": ("processo sei", "processo"), "nf": ("nota fiscal", "nf")}


def compras_da_tabela(linhas: list[list], so_manuais: bool = False) -> list[dict]:
    """Compras de uma planilha (a desta ferramenta ou outra com colunas de código e valor). Linhas "Em andamento",
    canceladas ou apagadas não contam: pedido em andamento entra pela própria ferramenta."""
    def posicoes(ln):
        cab = [_sem_acento(x) for x in ln]
        return {k: next((cab.index(n) for n in nomes if n in cab), None) for k, nomes in _CAMPOS.items()}

    i = next((j for j, ln in enumerate(linhas[:20]) if None not in (posicoes(ln)["codigo"], posicoes(ln)["valor"])), None)
    if i is not None:
        pos = posicoes(linhas[i])
    else:
        raise ValueError("Não achei as colunas de código e de valor na planilha. Use a planilha de controle da ferramenta "
                         "ou uma com as colunas “Código SES”, “Valor (R$)” e “Data”.")
    pega = lambda ln, k: ln[pos[k]] if pos[k] is not None and pos[k] < len(ln) else ""
    saida = []
    for ln in linhas[i + 1:]:
        situacao, origem = _sem_acento(pega(ln, "situacao")), pega(ln, "origem")
        if situacao.startswith(("em andamento", "cancelad", "apagad")) or (so_manuais and not eh_manual(origem)):
            continue
        codigo, valor = str(pega(ln, "codigo") or "").strip().upper(), _numero(pega(ln, "valor"))
        if isinstance(pega(ln, "codigo"), float) and pega(ln, "codigo") == int(pega(ln, "codigo")):
            codigo = str(int(pega(ln, "codigo")))
        if not codigo or not valor or valor <= 0:
            continue
        saida.append({"codigo": codigo, "valor": round(valor, 2), "ano": _ano(pega(ln, "data")),
                      "processo": str(pega(ln, "processo") or "").strip(), "nf": str(pega(ln, "nf") or "").strip(),
                      "origem": str(origem or "").strip()})
    return saida


def linhas_manuais(arq: Path) -> list[list]:
    """Linhas "Lançado à mão" de uma planilha de controle já existente, para não sumirem quando ela é refeita."""
    if not arq.exists():
        return []
    linhas = ler_tabela(arq.name, arq.read_bytes())
    if not linhas:
        return []
    cab = [_sem_acento(x) for x in linhas[0]]
    if "origem" not in cab:
        return []
    o = cab.index("origem")
    saida = []
    for ln in linhas[1:]:
        if o < len(ln) and eh_manual(ln[o]):
            nova = (list(ln) + [""] * len(COLUNAS))[:len(COLUNAS)]
            nova[7] = _numero(nova[7]) if _numero(nova[7]) is not None else nova[7]
            saida.append(nova)
    return saida


def _texto_seguro(v):
    """Texto que começa como fórmula (digitado, vindo do SEI ou de um backup alheio) entra como texto, com apóstrofo:
    sem isso o Excel executaria =HYPERLINK(...) ou =cmd|... ao abrir a planilha."""
    return "'" + v if isinstance(v, str) and v.startswith(_INICIO_DE_FORMULA) else v


def _celula_csv(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if not math.isfinite(v):
            return ""
        return str(int(v)) if v == int(v) else str(v).replace(".", ",")
    return _texto_seguro(v)


def _br(iso: str) -> str:
    """aaaa-mm-dd -> dd/mm/aaaa; o que não for data ISO volta como veio."""
    s = str(iso or "")
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}" + (" " + s[11:16] if len(s) >= 16 else "")
    return s


def linha_pedido(pid: str, pedido: dict, fluxo: dict, item: str, criado: str, n_minutas: int, cat, regras) -> list:
    try:
        a = motor.analisar(pedido, cat, regras)
        qtd, val = a.get("quantidade_total"), a.get("valor_estimado")
    except Exception:  # noqa: BLE001  (pedido antigo ou incompleto: a linha sai sem quantidade e valor)
        qtd = val = None
    f = fluxo or {}
    fase = f.get("fase")
    feitos = sum(1 for v in (f.get("d") or {}).values() if v == "d")
    return ["Em andamento", "Pedido feito na ferramenta", criado, pedido.get("codigo_ses") or "", item,
            TIPOS.get(pedido.get("tipo") or "", ""), qtd, val, "", "", pedido.get("processo_intencao") or "",
            ", ".join(str(v) for v in (f.get("sei") or {}).values() if v),
            motor.FASES[fase] if isinstance(fase, int) and 0 <= fase < len(motor.FASES) else "", feitos,
            " | ".join(str(pedido.get("unidades") or "").splitlines()), n_minutas, _br(f.get("atualizado") or ""), pid]


def linhas_compras(hist: dict) -> list[list]:
    """Compras já feitas, a mais recente primeiro. A data é a da nota fiscal."""
    cs = sorted((hist or {}).get("compras", {}).values(), key=lambda c: str(c.get("nf_data") or ""), reverse=True)
    return [["Comprado", "Relação de Estoque Físico (" + str(c.get("origem") or "sem nº SEI") + ")", _br(c.get("nf_data")),
             c.get("codigo") or "", c.get("descricao") or "", TIPOS.get(motor.classificar_codigo(c.get("codigo"))["grupo"], ""), c.get("qtd"), c.get("total"), c.get("fornecedor") or "",
             c.get("nf") or "", c.get("processo") or "", "", "", "", "", "", "", ""] for c in cs]


def gravar_xlsx(destino: Path, linhas: list[list]) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Controle"
    ws.append(COLUNAS)
    for ln in linhas:
        ws.append([_texto_seguro(v) for v in ln])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F5FA8")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, w in enumerate(LARGURAS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r in ws.iter_rows(min_row=2, min_col=8, max_col=8):
        r[0].number_format = "#,##0.00"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    tmp = destino.with_suffix(".tmp.xlsx")
    wb.save(tmp)
    tmp.replace(destino)          # se a planilha estiver aberta no Excel, falha aqui (PermissionError) e a antiga fica intacta
    return destino


def csv_texto(linhas: list[list]) -> str:
    """Mesma planilha em CSV que o Excel abre (ponto e vírgula, vírgula decimal)."""
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";", lineterminator="\r\n")
    w.writerow(COLUNAS)
    for ln in linhas:
        w.writerow([_celula_csv(v) for v in ln])
    return "﻿" + buf.getvalue()
