"""Planilha única de controle das compras do PDPAS: uma linha por pedido feito na ferramenta e uma por compra já
registrada no histórico (Relações de Estoque Físico do SEI). Uma aba só, sem células mescladas."""
from __future__ import annotations

import csv
import io
from pathlib import Path

import motor

NOME = "Controle de compras PDPAS.xlsx"
COLUNAS = ["Situação", "Origem", "Data", "Código SES", "Item", "Tipo", "Quantidade", "Valor (R$)", "Fornecedor", "Nota fiscal",
           "Processo SEI", "Números SEI anotados", "Fase atual", "Passos feitos", "Unidades", "Minutas geradas", "Atualizado em",
           "Identificação"]
LARGURAS = [16, 30, 12, 12, 48, 13, 12, 14, 34, 12, 24, 28, 24, 10, 40, 10, 18, 26]
TIPOS = {"permanente": "Permanente", "consumo": "Consumo", "servico": "Serviço"}


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
             c.get("codigo") or "", c.get("descricao") or "", "Consumo", c.get("qtd"), c.get("total"), c.get("fornecedor") or "",
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
        ws.append(ln)
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
        w.writerow([(str(int(v)) if v == int(v) else str(v).replace(".", ",")) if isinstance(v, float) else ("" if v is None else v) for v in ln])
    return "﻿" + buf.getvalue()
