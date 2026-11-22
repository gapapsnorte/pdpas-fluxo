"""Versão web da Ferramenta PDPAS: mesmas regras (motor.py, pecas.py, textos.py), rodando no navegador via Pyodide.

Não há servidor: os pedidos e rascunhos ficam no localStorage do aparelho (e no arquivo de backup que a pessoa baixa).
As minutas em Word são geradas na hora, a partir do pedido guardado, e baixadas.
"""
from __future__ import annotations

import base64
import io
import json
import re
import zipfile
from datetime import date, datetime
from pathlib import Path

import historico
import motor
import pecas

CHAVE = "pdpas-web-v1"
CHAVE_HIST = "pdpas-web-historico-v1"
TMP = Path("/tmp/processos")
CAT = motor.Catalogo()
motor.CAT_ITEM = CAT.achar
PID = re.compile(r"[\w-]{1,80}")
RID = re.compile(r"[\w-]{1,40}")
HIST_ARQ = motor.ARQ_HISTORICO


def _ls():
    from js import localStorage
    return localStorage


def _ler() -> dict:
    try:
        d = json.loads(_ls().getItem(CHAVE) or "{}")
    except (ValueError, TypeError):
        d = {}
    if not isinstance(d, dict):
        d = {}
    d.setdefault("pedidos", {})
    d.setdefault("rascunhos", {})
    return d


def _gravar(d: dict) -> None:
    try:
        _ls().setItem(CHAVE, json.dumps(d, ensure_ascii=False))
    except Exception as e:  # noqa: BLE001
        raise ValueError("Este navegador não deixou guardar o pedido no aparelho (espaço cheio ou modo privado). Baixe o backup.") from e


def carregar_historico_guardado() -> None:
    """Histórico importado antes (Visão Gerencial, Nada Consta, Relações) volta para o motor."""
    txt = _ls().getItem(CHAVE_HIST)
    if txt:
        HIST_ARQ.parent.mkdir(parents=True, exist_ok=True)
        HIST_ARQ.write_text(str(txt), encoding="utf-8")


def _gravar_historico() -> None:
    h = motor.hist()
    historico.salvar(h, HIST_ARQ)
    try:
        _ls().setItem(CHAVE_HIST, HIST_ARQ.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        pass


def _item(i: dict) -> dict:
    enq = motor.classificar_codigo(i.get("codigo_ses"))
    return {"codigo": i.get("codigo_ses"), "br": i.get("codigo_br"), "categoria": i.get("classif_categoria"),
            "resumo": i.get("desc_resumida"), "completa": i.get("desc_completa"), "grupo": enq["grupo"], "destino": enq["destino"]}


def _busca(termo: str) -> list[dict]:
    h = motor.hist()
    comprados = {c["codigo"] for c in h["compras"].values()} | {c["alias"] for c in h["compras"].values() if c.get("alias")}
    alvo = termo.strip().upper()
    itens, vistos = [], set()
    for i in CAT.buscar(termo, 10 ** 6):
        cod = str(i.get("codigo_ses") or "").upper()
        if cod in vistos:
            continue
        vistos.add(cod)
        itens.append({**_item(i), "comprado": cod in comprados})
    itens.sort(key=lambda i: (str(i["codigo"]).upper() != alvo, not i["comprado"]))
    return itens[:40]


def _resumo() -> dict:
    h = motor.hist()
    datas = [c["nf_data"] for c in h["compras"].values() if c["nf_data"]]
    meses = [m for v in h["consumo"].values() for m in v]
    return {"compras": len(h["compras"]), "itens": len(h["consumo"]), "de": min(datas, default=""), "ate": max(datas, default=""),
            "consumo_ate": max(meses, default=""), "pasta": "", "compartilhada": ""}


def _sei() -> dict:
    return {"disponivel": False, "devida": False, "rodando": False, "fase": "", "historico": _resumo()}


def _pasta_nome(analise: dict) -> str:
    nome = re.sub(r"[^\w-]+", "_", analise.get("codigo") or "servico")
    return f"{date.today().isoformat()}_{nome}"


def _id_pedido(pid: str) -> str:
    pid = (pid or "").replace("\\", "/").rstrip("/").split("/")[-1]
    if not PID.fullmatch(pid):
        raise ValueError("pedido sem identificação")
    return pid


def _pedido_guardado(pid: str, estado: dict) -> dict:
    p = estado["pedidos"].get(pid)
    if not p:
        raise ValueError("Não encontrei esse pedido neste aparelho. Se ele foi feito em outro computador, carregue o backup.")
    return p


def _info(pid: str, reg: dict) -> dict:
    p = reg.get("pedido") or {}
    achados = CAT.achar(p.get("codigo_ses")) if p.get("codigo_ses") else []
    item = (achados[0].get("desc_resumida") if achados else "") or p.get("objeto") or ""
    return {"id": pid, "pasta": f"/Processos/{pid}", "tipo": p.get("tipo") or "", "codigo": p.get("codigo_ses") or "", "item": item,
            "ata": p.get("ata") or "", "reparo": bool(p.get("reparo_equipamento")), "predio": bool(p.get("altera_predio")),
            "gsap_origem": p.get("gsap_origem") or "", "criado": reg.get("criado_br") or "",
            "arquivos": reg.get("arquivos") or [], "fluxo": reg.get("fluxo") or {}}


def _quando() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M")


def _gerar_pasta(pid: str, pedido: dict, so: str | None = None) -> tuple[list[Path], dict]:
    regras = motor.carregar_regras()
    analise = motor.analisar(pedido, CAT, regras)
    pasta = TMP / pid
    pasta.mkdir(parents=True, exist_ok=True)
    if so:
        return [pecas.gerar_peca(so, pedido, analise, regras, pasta)], analise
    return pecas.gerar_tudo(pedido, analise, regras, pasta), analise


def _arquivos_de(pid: str, estado: dict) -> list[dict]:
    return estado["pedidos"][pid].get("arquivos") or []


def _baixar_arquivo(pid: str, nome: str, estado: dict) -> dict:
    reg = _pedido_guardado(pid, estado)
    pedido = reg["pedido"]
    arqs, _ = _gerar_pasta(pid, pedido)            # sempre uma cópia nova, com a data de hoje
    por_nome = {a.name: a for a in arqs}
    if nome:
        a = por_nome.get(nome)
        if a is None:
            raise ValueError("Não achei essa minuta no pedido.")
        return {"__download": {"nome": a.name, "b64": base64.b64encode(a.read_bytes()).decode(), "mime":
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for a in arqs:
            z.write(a, a.name)
    return {"__download": {"nome": f"minutas_{pid}.zip", "b64": base64.b64encode(buf.getvalue()).decode(), "mime": "application/zip"}}


def exportar() -> str:
    return json.dumps({"formato": "pdpas-web", "versao": 1, "exportado": datetime.now().isoformat(timespec="seconds"),
                       "estado": _ler(), "historico": _ls().getItem(CHAVE_HIST) or ""}, ensure_ascii=False)


def importar_backup(txt: str) -> str:
    d = json.loads(txt)
    if not isinstance(d, dict) or d.get("formato") != "pdpas-web" or not isinstance(d.get("estado"), dict):
        raise ValueError("Este arquivo não é um backup da Ferramenta PDPAS (versão web).")
    novo, atual = d["estado"], _ler()
    novo.setdefault("pedidos", {})
    novo.setdefault("rascunhos", {})
    for k in ("pedidos", "rascunhos"):                  # junta: o que é mais recente vence
        for i, v in novo[k].items():
            a = atual[k].get(i)
            if not a or str(v.get("atualizado") or v.get("criado") or "") >= str(a.get("atualizado") or a.get("criado") or ""):
                atual[k][i] = v
    _gravar(atual)
    return json.dumps({"pedidos": len(atual["pedidos"]), "rascunhos": len(atual["rascunhos"])})


def tratar(metodo: str, caminho: str, consulta: str, corpo: str) -> str:
    """Devolve JSON {"code": n, "obj": {...}}. Mesmas rotas do app.py, sem arquivos nem janelas."""
    try:
        q = dict(p.split("=", 1) for p in consulta.split("&") if "=" in p) if consulta else {}
        from urllib.parse import unquote
        q = {k: unquote(v.replace("+", " ")) for k, v in q.items()}
        d = json.loads(corpo) if corpo else {}
        return json.dumps({"code": 200, "obj": _rota(metodo, caminho, q, d)}, ensure_ascii=False)
    except Exception as e:  # noqa: BLE001
        return json.dumps({"code": 400, "obj": {"erro": str(e)}}, ensure_ascii=False)


def _rota(metodo: str, caminho: str, q: dict, d: dict):
    if caminho == "/api/ping":
        return {"ok": True}
    estado = _ler()
    if caminho == "/api/inicio":
        r = motor.carregar_regras()
        ras = sorted([{**v, "id": i} for i, v in estado["rascunhos"].items() if (v.get("pedido") or {}).get("tipo")],
                     key=lambda x: x.get("atualizado", ""), reverse=True)
        return {"gsaps": r["gsaps"], "ano": r["ano"], "catalogo_em": CAT.atualizado_em[:10], "historico": _resumo(), "rascunhos": ras,
                "n_itens": len(CAT.itens), "saida": "", "naturezas": motor.naturezas_do_ano(r["ano"]), "avisos": historico.AVISOS, "sei": _sei()}
    if caminho == "/api/sei_status":
        return _sei()
    if caminho in ("/api/sei_cancelar", "/api/sair"):
        return {"ok": True}
    if caminho in ("/api/sei_sync", "/api/sei_adiar", "/api/abrir_dados"):
        raise ValueError("Isso só existe na ferramenta instalada no computador.")
    if caminho == "/api/pedidos":
        lista = sorted(estado["pedidos"].items(), key=lambda kv: kv[1].get("criado", ""), reverse=True)
        return {"pedidos": [_info(i, r) for i, r in lista]}
    if caminho == "/api/pedido":
        pid = _id_pedido(q.get("id", ""))
        return _info(pid, _pedido_guardado(pid, estado))
    if caminho == "/api/buscar":
        return {"itens": _busca(q.get("q", ""))}
    if caminho == "/api/rascunho":
        rid = str(d.get("id") or "")
        if not RID.fullmatch(rid):
            raise ValueError("rascunho sem identificação")
        if d.get("apagar"):
            estado["rascunhos"].pop(rid, None)
        else:
            estado["rascunhos"][rid] = {**d, "atualizado": datetime.now().isoformat(timespec="seconds")}
        _gravar(estado)
        return {"ok": True}
    if caminho == "/api/sugestoes":
        return motor.sugestoes(d, motor.carregar_regras())
    if caminho == "/api/analisar":
        return motor.analisar(d, CAT, motor.carregar_regras())
    if caminho == "/api/importar":
        if d.get("texto"):
            texto, nome, dados = d["texto"], d.get("nome") or "texto colado", None
        else:
            nome, dados = d["nome"], base64.b64decode(d["b64"])
            if nome.lower().endswith(".json"):
                texto = None
            else:
                try:
                    texto = historico.extrair_texto(nome, dados)
                except Exception as e:  # noqa: BLE001
                    raise ValueError("Na versão web, leia só texto colado ou histórico .json. Para PDF, copie o texto do SEI e cole.") from e
        h = motor.hist()
        r = historico.importar_arquivo(h, nome, dados) if texto is None else historico.importar_texto(h, texto, nome)
        if not r["tipo"]:
            raise ValueError(f"Não reconheci {r['origem']}: esperava Relação de Estoque Físico, Visão Gerencial, Nada Consta ou histórico.")
        _gravar_historico()
        r["campos"] = [{"codigo": v["codigo"], **motor.art21_da_visao(v)} for v in r.pop("visoes")]
        r["historico"] = _resumo()
        return r
    if caminho == "/api/gerar":
        regras = motor.carregar_regras()
        analise = motor.analisar(d, CAT, regras)
        base, pid, n = _pasta_nome(analise), None, 2
        pid = base
        while pid in estado["pedidos"]:
            pid, n = f"{base}_{n}", n + 1
        arqs, _ = _gerar_pasta(pid, d)
        motor.aprender(d)
        _gravar_historico()
        q_ = _quando()
        estado["pedidos"][pid] = {"pedido": d, "fluxo": {}, "criado": datetime.now().isoformat(timespec="seconds"),
                                  "criado_br": date.today().strftime("%d/%m/%Y"),
                                  "arquivos": [{"nome": a.name, "quando": q_} for a in arqs]}
        _gravar(estado)
        return {"pasta": f"/Processos/{pid}", "arquivos": [a.name for a in arqs], "tem_bloqueio": analise["tem_bloqueio"], "aviso": ""}
    if caminho == "/api/fluxo":
        pid = _id_pedido(str(d.get("id") or ""))
        reg = _pedido_guardado(pid, estado)
        est = d.get("estado")
        if not isinstance(est, dict):
            raise ValueError("andamento inválido")
        limpo = {k: est.get(k) for k in ("d", "ans", "sei", "h", "fase") if k in est}
        limpo["atualizado"] = datetime.now().isoformat(timespec="seconds")
        reg["fluxo"] = limpo
        _gravar(estado)
        return {"ok": True}
    if caminho == "/api/gerar_peca":
        pid = _id_pedido(str(d.get("id") or ""))
        reg = _pedido_guardado(pid, estado)
        arqs, _ = _gerar_pasta(pid, reg["pedido"], so=str(d.get("peca") or ""))
        lista = [a for a in reg.get("arquivos", []) if a["nome"] != arqs[0].name]
        lista.append({"nome": arqs[0].name, "quando": _quando()})
        reg["arquivos"] = sorted(lista, key=lambda a: a["nome"])
        _gravar(estado)
        return {"arquivo": arqs[0].name, "arquivos": reg["arquivos"]}
    if caminho == "/api/abrir":
        alvo = str(d.get("alvo") or "").replace("\\", "/").rstrip("/")
        partes = [p for p in alvo.split("/") if p]
        if len(partes) >= 2 and partes[0] == "Processos":
            pid = _id_pedido(partes[1])
            return _baixar_arquivo(pid, partes[2] if len(partes) > 2 else "", estado)
        raise ValueError("não encontrei essa pasta")
    raise ValueError("rota desconhecida")
