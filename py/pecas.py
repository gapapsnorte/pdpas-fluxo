"""Gera as minutas em .docx: memo de intenção do setor de origem, Termo de Referência e checklist.

O texto do TR segue os TRs 25 e 31/2024 e 6/2025 da GAPAPS, com a base legal atualizada
(a Portaria 200/2023 citada neles foi revogada pela Portaria 473/2023).
Campos que a ferramenta não sabe saem como [A PREENCHER: ...], destacados em amarelo.
O TR de serviço segue o TR 28/2026 - SES/SRSNO/HRPL/GENF/NME (SEI 217574930).
O TR de permanente segue o TR 1/2026 - SES/SRSNO/DIRAPS/GAPAPS (SEI 218150428), com a mesma numeração.
As colunas de consumo seguem o TR 8/2026 do NUAL (SEI 210165374).
"""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.shared import Cm, Pt

import motor

MESES = ("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
         "setembro", "outubro", "novembro", "dezembro")
_CAMPO = re.compile(r"(\[A PREENCHER[^\]]*\]|\[VALIDAR: [^\]]*\])")
_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")     # texto colado de PDF/Word: o lxml recusa caracteres de controle
# "cmm" entrou em 05/11/2026: era sugerido em amarelo na janela, mas saía no TR sem [VALIDAR]
TEXTUAIS = ("objeto", "descricao", "local", "justificativa", "antecedentes", "contexto", "beneficios", "just_quantidade", "consumo_mensal",
            "tempo_duracao", "justificativa_ata", "just_acima_cmm", "texto_outra_unidade", "responsavel_aceite", "prazo_manifestacao", "cmm")


def data_extenso(d: date) -> str:
    return f"{d.day:02d} de {MESES[d.month - 1]} de {d.year}"


def _doc(regras: dict) -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.5)
    sec.top_margin = sec.bottom_margin = Cm(2)
    est = doc.styles["Normal"]
    est.font.name = "Calibri"
    est.font.size = Pt(11)
    for linha in regras["unidade"]["cabecalho"]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.add_run(linha).bold = True
    doc.add_paragraph()
    return doc


def _par(doc, texto: str, negrito=False, justificar=True):
    p = doc.add_paragraph()
    if justificar:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    for pedaco in _CAMPO.split(_CTRL.sub(" ", texto)):
        if not pedaco:
            continue
        r = p.add_run(pedaco)
        r.bold = negrito
        if _CAMPO.fullmatch(pedaco):
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return p


def _titulo(doc, texto: str):
    p = _par(doc, texto, negrito=True, justificar=False)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.keep_with_next = True


def _tabela(doc, cab: list[str], linhas: list[list[str]], larguras: list[float] | None = None):
    t = doc.add_table(rows=1, cols=len(cab))
    t.style = "Table Grid"
    for i, c in enumerate(cab):
        cel = t.rows[0].cells[i]
        cel.text = ""
        cel.paragraphs[0].add_run(_CTRL.sub(" ", str(c))).bold = True
    for ln in linhas:
        cels = t.add_row().cells
        for i, v in enumerate(ln):
            cels[i].text = ""
            for pedaco in _CAMPO.split(_CTRL.sub(" ", str(v))):
                if pedaco:
                    r = cels[i].paragraphs[0].add_run(pedaco)
                    if _CAMPO.fullmatch(pedaco):
                        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    if larguras:
        t.autofit = False
        for i, w in enumerate(larguras):
            t.columns[i].width = Cm(w)
        for row in t.rows:
            for i, w in enumerate(larguras):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return t


def _nome_arquivo(pedido: dict, analise: dict) -> str:
    base = analise.get("codigo") or "servico"
    return re.sub(r"[^\w-]+", "_", str(base))


def _desc(pedido: dict, analise: dict) -> tuple[str, str]:
    item = analise.get("item") or {}
    resumo = item.get("desc_resumida") or pedido.get("objeto") or "[A PREENCHER: objeto]"
    completa = item.get("desc_completa") or pedido.get("descricao") or "[A PREENCHER: descrição completa]"
    # O catálogo cola frases ("EXAME.Construído"); a minúscula evita mexer em abreviações e decimais.
    completa = re.sub(r"\s*\|\s*", " ", completa)   # o catálogo separa as linhas com "|"
    return resumo, re.sub(r"([a-zà-úA-ZÀ-Ú]{3})\.(?=[A-ZÀ-Ú][a-zà-ú])", r"\1. ", completa)


def _qtd(v: float) -> str:
    """1234567 -> '1.234.567'; 2,5 -> '2,5'. (f'{v:g}' virava '1.23457e+06' acima de 6 algarismos.)"""
    from decimal import Decimal
    d = Decimal(str(v)).normalize()
    texto = format(d, "f")
    inteiro, _, frac = texto.partition(".")
    inteiro = f"{int(inteiro):,}".replace(",", ".")
    return f"{inteiro},{frac}" if frac else inteiro


def _medicamento(analise: dict) -> bool:
    return motor.categoria_num(analise.get("item") or {}) == "0"


def _item_saude(analise: dict) -> bool:
    """Medicamento, material médico, odonto, laboratório, OPME e material médico de almoxarifado."""
    return motor.categoria_num(analise.get("item") or {}) in ("0", "1", "3", "4", "6", "11")


def _abastecimento(pedido: dict, analise: dict) -> str:
    """Segue os TRs dos NLF: cita o número SEI de cada peça em vez de só afirmar que falta."""
    if not (pedido.get("sei_nada_consta") or pedido.get("sei_visao") or pedido.get("sei_declaracao") or pedido.get("data_consulta")):
        return "[A PREENCHER: Nada Consta, Visão Gerencial e Declaração de ata, com número SEI e data da consulta]"
    campo = lambda c, f: _ou_campo(pedido, c, f)
    data = (pedido.get("data_consulta") or "").strip()
    ata = {"sem": "não há Ata de Registro de Preços vigente", "vigente": "há Ata de Registro de Preços vigente",
           "atrasada": "há ata com atraso de entrega"}.get(pedido.get("ata"), "[A PREENCHER: situação da ata]")
    return (f"Conforme Declaração ({campo('sei_declaracao', 'nº SEI')}), {ata}. Considerando o Relatório de Nada Consta "
            f"({campo('sei_nada_consta', 'nº SEI')}) e o Relatório - Visão Gerencial ({campo('sei_visao', 'nº SEI')})"
            + (f", emitidos em {data}," if data else "") + " na presente data não há estoque disponível para reabastecimento.")


def _verificacao_fornecedor(n: str) -> list[str]:
    return [f"{n}.3 Previamente à emissão da Ordem, a Administração verificará a existência de sanção que impeça a "
            "contratação, mediante consulta a cadastros oficiais: SICAF; Certidão Negativa Correcional da CGU (ePAD, "
            "CGU-PJ, CEIS, CNEP e CEPIM).",
            f"{n}.4 Serão exigidos a Certidão Negativa de Débitos relativa a Créditos Tributários Federais e à Dívida Ativa "
            "da União, a Certidão Negativa de Débitos junto ao GDF, o Certificado de Regularidade do FGTS e a Certidão "
            "Negativa de Débitos Trabalhistas."]


APROVACAO = ("Considerando os termos da Lei nº 14.133/2021, do Decreto nº 44.330/2023 e da Portaria nº 473/2023, aprovo o "
             "presente Termo de Referência de Dispensa de Licitação - PDPAS e ratifico a veracidade das informações "
             "prestadas, assim como afirmo a ausência de direcionamento de marca e/ou modelo.")


def _item_cod(analise: dict, resumo: str) -> str:
    """'descrição, código SES X' para item de catálogo; serviço não tem código e sai só com o objeto."""
    if analise.get("tipo") == "servico":
        return resumo
    return f"{resumo}, código SES {analise.get('codigo') or '[A PREENCHER: código SES]'}"


def _rotulo(analise: dict, resumo: str) -> str:
    cod = analise.get("codigo")
    return f"{cod} - {resumo}" if cod and analise.get("tipo") != "servico" else resumo


def _proc_intencao(pedido: dict) -> str:
    """Nº do processo do memorando de intenção. Não confundir com processo_numero, que é a compra regular em andamento."""
    return str(pedido.get("processo_intencao") or "").strip() or "[A PREENCHER: nº do processo do memorando de intenção]"


def _gerencias_interessadas(pedido: dict) -> list[str]:
    _, q = motor.somar_quantidades(pedido)
    return [g for g, v in q.items() if v and g != "APS Norte"]


def _assinaturas_permanente(pedido: dict, regras: dict) -> list[str]:
    """Ordem pactuada (Diretoria, áudio de 18/11/2026): gestores que sinalizaram no memorando de intenção → bloco da
    Diretoria (revisão de pendências e assinatura) → bloco da Superintendência (revisão e assinatura)."""
    gs = _gerencias_interessadas(pedido)
    linhas = [f"Gestores das gerências que manifestaram interesse no memorando de intenção ({_proc_intencao(pedido)}):"]
    linhas += [f"{g}: [A PREENCHER: nome e matrícula do gestor]" for g in gs] or \
              ["[A PREENCHER: gerência, nome e matrícula de cada gestor que manifestou interesse]"]
    nome, sigla = _diretoria(regras)
    linhas += [f"{nome} ({sigla}), em bloco de assinatura, "
               "após revisão: [A PREENCHER: nome e matrícula]",
               "Superintendência da Região de Saúde Norte (SES/SRSNO), em bloco de assinatura, após revisão: "
               "[A PREENCHER: nome e matrícula]"]
    return linhas


def _origem(pedido: dict) -> str:
    return str(pedido.get("gsap_origem") or "").strip() or "[A PREENCHER: setor de origem]"


def _diretoria(regras: dict) -> tuple[str, str]:
    """Diretoria do setor de origem (Colegiado 23/11/2026): na APS Norte é a DIRAPS; outra instância troca em regras.json."""
    o = regras.get("origem") or {}     # regras.json antigo em Dados (editado à mão) não tem a chave: vale a DIRAPS Norte
    return (o.get("diretoria") or "Diretoria Regional de Atenção Primária à Saúde da Região Norte",
            o.get("diretoria_sigla") or "SES/SRSNO/DIRAPS")


def _qtd_origem(pedido: dict) -> str:
    """Quantidade do setor de origem, só se ela já estiver na grade; a ferramenta nunca sugere quantidade."""
    _, q = motor.somar_quantidades(pedido)
    orig = str(pedido.get("gsap_origem") or "").strip()
    if orig and orig in q:
        return _qtd(q[orig])
    return "[A PREENCHER: quantidade necessária no setor de origem]"


def memo_intencao(pedido: dict, analise: dict, regras: dict, destino: Path, hoje: date | None = None) -> Path:
    """Passo 1 (Colegiado 23/11/2026): o setor que percebeu a necessidade abre processo PÚBLICO com o memorando de intenção à
    sua diretoria. O memorando fala só da necessidade do setor: não descreve a responsabilidade dos demais."""
    hoje = hoje or date.today()
    resumo, _ = _desc(pedido, analise)
    cod = analise.get("codigo") or "[A PREENCHER: código SES]"
    nome, sigla = _diretoria(regras)
    doc = _doc(regras)
    _par(doc, f"Memorando Nº [A PREENCHER: número]/{hoje.year} - {sigla}/{_origem(pedido)}                    "
              f"Brasília, {data_extenso(hoje)}.", justificar=False)
    _par(doc, f"À {nome} ({sigla}),", justificar=False)
    _par(doc, f"Assunto: PDPAS - intenção de aquisição de {resumo}", negrito=True, justificar=False)
    for n, t in enumerate([
        (f"Este setor identificou a necessidade de contratação do serviço: {resumo}." if analise.get("tipo") == "servico" else
         f"Este setor identificou a necessidade de {resumo}, código SES {cod}, localizado no catálogo de bens e insumos "
         "padronizados da SES-DF."),
        f"A necessidade decorre de: {_ou_campo(pedido, 'justificativa', 'justificativa da necessidade')}",
        f"Quantidade necessária neste setor: {_qtd_origem(pedido)}.",
        "Registra-se esta intenção para o levantamento da necessidade das demais gerências da região, com vistas a um único "
        "processo de aquisição, conforme o Memorando Circular Nº 2/2024 - SES/SRSNO (00060-00326247/2024-15).",
    ], 1):
        _par(doc, f"{n}.  {t}")
    arq = destino / f"1_Memo_intencao_{_nome_arquivo(pedido, analise)}.docx"
    doc.save(arq)
    return arq


def despacho_intencao(pedido: dict, analise: dict, regras: dict, destino: Path, hoje: date | None = None) -> Path:
    """Passo 2: a diretoria do setor de origem despacha a todas as gerências da região e fixa o prazo de 7 dias úteis.
    Minuta oferecida a quem prepara o despacho na diretoria; não sai do setor de origem."""
    hoje = hoje or date.today()
    resumo, _ = _desc(pedido, analise)
    cfg = motor.config_intencao(regras)
    limite = pedido.get("prazo_manifestacao") or "[A PREENCHER: data-limite]"
    _, sigla = _diretoria(regras)
    doc = _doc(regras)
    _par(doc, f"Despacho - {sigla}                    Brasília, {data_extenso(hoje)}.", justificar=False)
    _par(doc, "Às Gerências de Serviços de Atenção Primária à Saúde da Região Norte (SES/SRSNO/DIRAPS/GSAPs),", justificar=False)
    if cfg["extras"]:
        _par(doc, "Com cópia para: " + "; ".join(cfg["extras"]) + ".", justificar=False)
    else:
        _par(doc, "Com cópia para: [A PREENCHER: áreas que sempre devem receber o despacho e que historicamente ficam de "
                  "fora].", justificar=False)
    _par(doc, f"Assunto: PDPAS - intenção de aquisição de {resumo}", negrito=True, justificar=False)
    for n, t in enumerate([
        f"Encaminha-se o memorando de intenção da {_origem(pedido)} ({_proc_intencao(pedido)}), "
        f"para {'contratação' if analise.get('tipo') == 'servico' else 'aquisição'} via PDPAS de {_item_cod(analise, resumo)}.",
        "Considerando o Memorando Circular Nº 2/2024 - SES/SRSNO (00060-00326247/2024-15), que orienta levantar e consolidar a "
        "necessidade, unificando itens repetidos, a fim de zelar pelo princípio da economicidade e pelo não fracionamento das "
        "despesas.",
        ("Considerando a necessidade de atender, em um só processo, todas as GSAPs que se manifestarem."
         if analise.get("tipo") == "servico" else
         "Considerando o pactuado em colegiado com as GSAPs, de um processo por código de material, de forma a atender todas as "
         "GSAPs que se manifestarem."),
        f"Solicita-se que cada gerência se posicione, no prazo de {cfg['dias']} dias úteis (até {limite}), informando se tem ou "
        "não interesse no item e, havendo, a quantidade necessária para suprir a demanda do local de trabalho.",
        f"Findo o prazo, o processo retorna à {_origem(pedido)}, que consolidará as manifestações e dará seguimento ao "
        "processo de solicitação.",
    ], 1):
        _par(doc, f"{n}.  {t}")
    arq = destino / f"1b_Despacho_DIRAPS_gerencias_{_nome_arquivo(pedido, analise)}.docx"
    doc.save(arq)
    return arq


def despacho_consolidado(pedido: dict, analise: dict, regras: dict, destino: Path, hoje: date | None = None) -> Path:
    """Passo 3 (Colegiado 23/11/2026): terminado o prazo das gerências, o próprio setor de origem consolida o interesse e o
    registra no processo do memorando de intenção."""
    hoje = hoje or date.today()
    resumo, _ = _desc(pedido, analise)
    total, q = motor.somar_quantidades(pedido)
    _, sigla = _diretoria(regras)
    doc = _doc(regras)
    _par(doc, f"Despacho - {sigla}/{_origem(pedido)}                    Brasília, {data_extenso(hoje)}.", justificar=False)
    _par(doc, f"Para registro no processo do memorando de intenção ({_proc_intencao(pedido)}).", justificar=False)
    _par(doc, f"Assunto: PDPAS - consolidado de {resumo}", negrito=True, justificar=False)
    _par(doc, f"1.  Encerrado o prazo de manifestação das gerências em {_ou_campo(pedido, 'prazo_manifestacao', 'data-limite')}, "
              f"segue o consolidado de {_item_cod(analise, resumo)}, para toda a Atenção Primária da Região Norte.")
    por_bem = bool(q) and not any((motor._num(v) or 0) > 0 for v in (pedido.get("quantidades") or {}).values()) \
        and not motor.ler_unidades(pedido)[0]
    linhas = [[g, _qtd(v)] for g, v in q.items() if v] or [["[A PREENCHER: gerências que manifestaram interesse]", "[A PREENCHER]"]]
    linhas.append(["Total", _qtd(total) if total else "[A PREENCHER]"])
    _tabela(doc, ["Bem" if por_bem else "Gerência", "Quantidade"], linhas, [11, 5])
    _par(doc, "2.  Gerências que não se manifestaram no prazo: [A PREENCHER: lista, conforme a planilha de acompanhamento].")
    _par(doc, "3.  Com o consolidado acima, este setor construirá o Termo de Referência"
              + ("" if analise.get("tipo") == "servico" else ", em processo por código de material, conforme pactuado em colegiado")
              + f", em processo SEI novo, público, sem documento restrito e relacionado a este processo ({_proc_intencao(pedido)}).")
    arq = destino / f"1c_Despacho_consolidado_{_nome_arquivo(pedido, analise)}.docx"
    doc.save(arq)
    return arq



def termo_referencia(pedido: dict, analise: dict, regras: dict, destino: Path, hoje: date | None = None) -> Path:
    tipo = analise["tipo"]
    resumo, completa = _desc(pedido, analise)
    total, qs = motor.somar_quantidades(pedido)
    un = regras["unidade"]
    servico = tipo == "servico"
    inciso = "I" if (servico and pedido.get("engenharia")) else "II"     # obra/engenharia: art. 75, I
    cod = analise.get("codigo") or "-"
    cod_br = (analise.get("item") or {}).get("codigo_br") or "-"
    local_perm = regras["entrega"].get("permanente_sigla") or "NPDOC"
    doc = _doc(regras)
    _par(doc, f"Termo de Referência - {un['sigla']}", negrito=True, justificar=False)

    objeto = "Contratação de serviço" if servico else ("Aquisição de material permanente" if tipo == "permanente"
                                                       else "Aquisição de material de consumo")
    _titulo(doc, "1. OBJETO")
    if tipo == "permanente":
        _par(doc, f"1.1 O presente Termo de Referência tem por objeto a aquisição do bem permanente {resumo}, conforme "
                  "especificações técnicas detalhadas, visando atender às demandas assistenciais das unidades de "
                  "Atenção Primária da Região Norte.")
        _par(doc, "1.2 O objeto está caracterizado como bem permanente, não se enquadrando na categoria de bem de luxo.")
        _par(doc, "1.3 A contratação será realizada por dispensa de licitação com recursos do Programa de Descentralização "
                  "Progressiva de Ações de Saúde - PDPAS, em consonância com o art. 75, inciso II, da Lei nº 14.133/2021, "
                  "observando o parcelamento sempre que técnica e economicamente viável (art. 40, V, “b”).")
        _par(doc, "1.4 O prazo de vigência contratual será de até 30 (trinta) dias contados da publicação da homologação "
                  "da compra, sendo improrrogável.")
        _par(doc, "1.5 Nos termos dos §§ 4º e 5º do art. 105 do Decreto nº 44.330/2023, a estimativa de preços será "
                  "realizada concomitantemente à seleção da proposta mais vantajosa, por meio de solicitação formal de "
                  "cotações a fornecedores.")
    else:
        _par(doc, f"{objeto}: {resumo}, por DISPENSA DE LICITAÇÃO, nos termos do inciso {inciso} do artigo 75 da Lei "
                  "n° 14.133/2021 e suas alterações, mediante recursos financeiros do Programa de Descentralização "
                  "Progressiva de Ações de Saúde - PDPAS, para atender às necessidades da Atenção Primária da Região "
                  "de Saúde Norte, conforme especificações e quantitativos constantes neste Termo de Referência.")
        _par(doc, "A contratação se dará na categoria de entrega imediata, com prazo de até 30 (trinta) dias da ordem de "
                  "fornecimento, nos termos do art. 254 do Decreto nº 44.330/2023. A estimativa de preços será realizada "
                  "concomitantemente à seleção da proposta, conforme os §§ 4º e 5º do art. 105 do mesmo Decreto.")

    _titulo(doc, "2. JUSTIFICATIVA")
    for sub, campo, falta in (("2.1 MOTIVAÇÃO PARA A CONTRATAÇÃO", "justificativa", "motivação e relevância para a população"),
                              ("2.2 ANTECEDENTES", "antecedentes", "antecedentes"),
                              ("2.3 CONTEXTO ATUAL", "contexto", "contexto atual"),
                              ("2.4 BENEFÍCIOS/OBJETIVOS", "beneficios", "benefícios esperados")):
        _par(doc, sub, negrito=True, justificar=False)
        _par(doc, (pedido.get(campo) or "").strip() or f"[A PREENCHER: {falta}]")
    if not servico:
        _par(doc, "2.5 SITUAÇÃO ATUAL DO ABASTECIMENTO", negrito=True, justificar=False)
        _par(doc, _abastecimento(pedido, analise))
        if tipo == "consumo" and pedido.get("outra_unidade") == "reservado" and (pedido.get("texto_outra_unidade") or "").strip():
            _par(doc, pedido["texto_outra_unidade"].strip())
        andamento = {"etp": "em elaboração do ETP", "tr": "em elaboração do TR", "licitacao": "em licitação",
                     "contratacao": "em contratação"}.get(pedido.get("processo_andamento"))
        if andamento:
            _par(doc, f"Há processo de aquisição {andamento}: "
                      f"{_ou_campo(pedido, 'processo_numero', 'número do processo')}.")

    _titulo(doc, "3. DESCRIÇÃO DO OBJETO")
    _par(doc, "3.1 ESPECIFICAÇÃO TÉCNICA", negrito=True, justificar=False)
    bens = motor.ler_patrimonios(pedido) if servico else []
    if bens:
        _tabela(doc, ["Item", "Patrimônio", "Descrição, marca, modelo e série", "Local", "SEI ficha SisGepat"],
                [[str(i), b["patrimonio"], b["descricao"], b["local"], b["ficha"] or "[A PREENCHER]"]
                 for i, b in enumerate(bens, 1)], [1, 2.4, 6.6, 3.4, 2.6])
        _par(doc, completa)
    elif tipo == "consumo":
        mensal = (pedido.get("consumo_mensal") or "").strip() or "[A PREENCHER: consumo mensal]"
        duracao = (pedido.get("tempo_duracao") or "").strip() or "[A PREENCHER: tempo de duração]"
        _tabela(doc, ["Item", "Descrição completa", "Quantidade", "Código SES", "Código BR",
                      "Consumo mensal (estimativa)", "Tempo de duração (estimativa)"],
                [["1", completa, _qtd(total) if total else "[A PREENCHER: quantidade]", cod, cod_br, mensal, duracao]], [1, 5.2, 1.9, 1.7, 1.7, 2.3, 2.2])
    else:
        _tabela(doc, ["Item", "Descrição completa", "Quantidade", "Código SES", "Código BR"],
                [["1", completa, _qtd(total) if total else "[A PREENCHER: quantidade]", cod, cod_br]], [1.2, 7.6, 2.6, 2.3, 2.3])
    if tipo == "permanente":
        _par(doc, "3.1.1 Os equipamentos deverão ser novos, estar em perfeito estado e condições de uso, entregues "
                  "devidamente montados e prontos para uso, acompanhados de nota fiscal, com garantia mínima não inferior "
                  f"a 12 (doze) meses, a partir do aceite do produto pelo {local_perm}.")
    _tabela(doc, ["Programa de trabalho", "Fonte de recursos"],
            [["Programa de Descentralização Progressiva de Ações de Saúde - PDPAS",
              "PDPAS - Atenção Primária"]], [8, 8])
    _par(doc, "3.2 JUSTIFICATIVA DO QUANTITATIVO", negrito=True, justificar=False)
    unidades = motor.ler_unidades(pedido)[0]
    origem = "dos quantitativos por unidade" if unidades else "das manifestações das GSAPs"
    if bens:
        _par(doc, (pedido.get("just_quantidade") or "").strip()
             or f"Um serviço por bem relacionado no item 3.1, no total de {len(bens)} bem(ns).")
    else:
        grade = unidades or qs
        _par(doc, (pedido.get("just_quantidade") or "").strip()
             or f"Quantitativo definido a partir {origem}, conforme memória de cálculo abaixo.")
        _tabela(doc, ["Unidade" if unidades else "GSAP", _rotulo(analise, resumo)],
                ([[g, _qtd(v)] for g, v in grade.items()] + [["Total", _qtd(sum(grade.values()))]]) if grade
                else [["[A PREENCHER: GSAP ou unidade]", "[A PREENCHER: quantidade]"]], [8, 8])
    if _medicamento(analise):
        cmm = pedido.get("cmm")
        _par(doc, f"Consumo médio mensal (CMM): {cmm if cmm not in (None, '') else '[A PREENCHER: CMM]'}.")
        if (pedido.get("just_acima_cmm") or "").strip():
            _par(doc, f"Justificativa da solicitação acima do CMM: {pedido['just_acima_cmm'].strip()}")
        _par(doc, "3.3 SITUAÇÃO CMED", negrito=True, justificar=False)
        _par(doc, "[A PREENCHER: se o item está sujeito ao CAP (Resolução CMED nº 3/2011) e se consta do Convênio ICMS nº 87/2002]")

    _titulo(doc, "4. DA FUNDAMENTAÇÃO LEGAL")
    teto = ((regras.get("teto_dispensa") or {}).get("base") or "[A PREENCHER: ato que atualiza o teto da dispensa]").split("atualizado pelo ")[-1]
    base = [
        "Portaria SES nº 473, de 04 de dezembro de 2023, alterada pela Portaria nº 115/2024, que regulamenta o "
        "Programa de Descentralização Progressiva de Ações de Saúde - PDPAS, instituído pelo Decreto nº 44.322/2023;",
        f"Inciso {inciso} do art. 75 da Lei nº 14.133/2021 e suas alterações, regulamentada no GDF pelo Decreto nº 44.330/2023;",
        f"{teto}, que atualiza os valores estabelecidos na Lei nº 14.133/2021;",
    ]
    if tipo == "permanente":
        base.append("Portaria SEF nº 135, de 26 de julho de 2016, e Instrução Normativa nº 03, de 15 de maio de 2018 - "
                    "SUCON/SEF, sobre o registro no SisGepatWEB e a incorporação ao patrimônio do DF;")
        base.append("Memorando Circular Nº 2/2024 - SES/SRSNO, que estabelece o fluxo de aquisição de bem permanente "
                    "com recurso do PDPAS.")
    for i, t in enumerate(base, 1):
        _par(doc, f"4.{i} {t}")

    _titulo(doc, "5. DO CRITÉRIO DE JULGAMENTO")
    _par(doc, ("5.1 O critério de julgamento adotado será o de MENOR PREÇO por patrimônio." if bens else
               "5.1 O critério de julgamento adotado será o de MENOR PREÇO por item, em atenção ao artigo 33 da Lei "
               "nº 14.133/2021."))
    _par(doc, "5.2 Serão desclassificadas as propostas que não atenderem à descrição completa do item, bem como aquelas "
              "que apresentarem preços excessivos, comparados aos preços de mercado, em consonância com o disposto no "
              "artigo 59 da Lei nº 14.133/2021 e no art. 93, § 8º, do Decreto nº 44.330/2023.")
    # numeração sequencial: cada tipo recebe só as cláusulas que lhe cabem, sem lacuna nem repetição de número
    k = 2
    def c5(texto: str) -> None:
        nonlocal k
        k += 1
        _par(doc, f"5.{k} {texto}")
    if tipo in ("consumo", "permanente"):
        c5("A proposta não poderá ter validade inferior a 30 (trinta) dias.")
    if _medicamento(analise):
        c5("Junto à proposta, a empresa enviará à Gerência de Orçamento e Finanças: Autorização de Funcionamento "
           "de Empresa (AFE) da ANVISA, nos termos da RDC nº 16/2014, e Autorização Especial (AE) para medicamento ou "
           "substância sujeita a controle especial; e Alvará ou Licença Sanitária vigente.")
    if servico and pedido.get("visita_tecnica"):
        c5("A proposta deverá vir acompanhada do Comprovante de Visita Técnica para Orçamentos do PDPAS, "
           "preenchido na unidade onde o serviço será executado.")
    if tipo == "permanente":
        c5("Junto à proposta comercial, a licitante deverá apresentar catálogo e/ou prospectos, em língua "
           "portuguesa, que contenham a descrição detalhada do produto ofertado, que será analisado conforme o "
           "descritivo deste Termo de Referência pela área demandante.")
    c5("A contratação deverá ser feita preferencialmente com microempresa, empresa de pequeno porte ou "
       "microempreendedor individual, conforme o art. 234 do Decreto nº 44.330/2023.")

    if tipo == "permanente":
        _clausulas_permanente(doc, regras, un, pedido)
        arq = destino / f"2_Termo_de_Referencia_{_nome_arquivo(pedido, analise)}.docx"
        doc.save(arq)
        return arq

    n = 6
    if not servico:
        _titulo(doc, f"{n}. DA AMOSTRA")
        _par(doc, f"{n}.1 Solicitamos o envio de uma fotografia do produto ofertado, a título de amostra, para conferência "
                  "e validação quanto à conformidade com as especificações apresentadas.")
        if pedido.get("amostra_fisica"):
            _par(doc, f"{n}.2 {AMOSTRA_FISICA}")
        n += 1

    _titulo(doc, f"{n}. DO FORNECIMENTO E DO INSTRUMENTO DE AJUSTE")
    _par(doc, f"{n}.1 Para o fornecimento do objeto, deverão ser observadas as especificações estabelecidas no item 3.")
    if servico:
        _par(doc, f"{n}.2 O serviço deverá ser realizado no local especificado e a nota fiscal atestada pelo responsável "
                  "técnico após a total realização do serviço.")
        _par(doc, f"{n}.3 A Autorização de Serviço substituirá a nota de empenho, conforme art. 233 do Decreto nº 44.330/2023.")
    else:
        _par(doc, f"{n}.2 Os objetos deverão ser entregues com a respectiva nota fiscal, para ateste e liquidação da despesa.")
        _par(doc, f"{n}.3 A Autorização de Compra substituirá a nota de empenho, conforme art. 233 do Decreto nº 44.330/2023.")
    n += 1

    _titulo(doc, f"{n}. DO LOCAL DE ENTREGA")
    if servico:
        local = pedido.get("local") or "[A PREENCHER: local de execução do serviço]"
        _par(doc, f"{n}.1 O serviço será executado em: {local}.")
        _par(doc, f"{n}.2 Responsável pelo aceite: {_ou_campo(pedido, 'responsavel_aceite', 'nome, matrícula, cargo e contato')}.")
    else:
        local = regras["entrega"].get(tipo) or "[A PREENCHER: endereço de entrega de consumo]"
        # o endereço de regras.json pode trazer o horário do setor: nesse caso o TR não repete outro por cima
        tem_hora = bool(re.search(r"[0-9]{1,2} ?h", local)) and "A PREENCHER" not in local
        horario = "" if tem_hora else ", no horário de 8h às 12h e de 13h às 17h"
        _par(doc, f"{n}.1 Os materiais deverão ser entregues no prazo de até 30 dias, contados do recebimento da "
                  f"Autorização de Compra{horario}, em: {local}.")
    n += 1

    _titulo(doc, f"{n}. DEVERES DA UNIDADE SOLICITANTE")
    instrumento = "Autorização de Serviço" if servico else "Autorização de Compra"
    deveres = [f"Receber o objeto, no prazo e nas condições estabelecidas na {instrumento};",
               "Verificar a conformidade do objeto com as especificações, para fins de aceitação e recebimento;",
               "Comunicar à empresa, por escrito, imperfeições, falhas ou irregularidades verificadas;",
               "Efetuar o pagamento da Nota Fiscal atestada, no valor correspondente ao fornecimento."]
    if tipo == "permanente":
        deveres.append("Providenciar registro no Sistema Geral de Patrimônio (SisGepat) e incorporação do bem ao "
                       "patrimônio da Unidade.")
    for i, t in enumerate(deveres, 1):
        _par(doc, f"{n}.{i} {t}")
    n += 1

    _titulo(doc, f"{n}. DEVERES DA EMPRESA VENCEDORA")
    for i, t in enumerate([
        "Cumprir todas as obrigações constantes do Termo de Referência, do Pedido de Cotação e da proposta, assumindo "
        "como exclusivamente seus os riscos e as despesas decorrentes da boa e perfeita execução do objeto;",
        "Entregar o objeto em perfeitas condições, acompanhado da nota fiscal, com marca, fabricante, modelo, "
        "procedência e prazo de garantia ou validade;",
        "Responsabilizar-se pelos vícios e danos decorrentes do objeto, conforme os artigos 12, 13 e 17 a 27 do Código "
        "de Defesa do Consumidor (Lei nº 8.078, de 1990);",
        "Substituir, reparar ou corrigir, às suas expensas, no prazo de 30 dias da notificação, o objeto com defeitos;",
        "Manter, durante todo o processo, as condições de habilitação e qualificação exigidas na dispensa de licitação.",
    ], 1):
        _par(doc, f"{n}.{i} {t}")
    n += 1

    _titulo(doc, f"{n}. DA DISPONIBILIDADE DE RECURSOS")
    _par(doc, f"{n}.1 Recurso financeiro do PDPAS, disponibilizado em conta bancária da {un['razao']}, CNPJ {un['cnpj']}.")
    n += 1

    _titulo(doc, f"{n}. DOS PRAZOS E DAS CONDIÇÕES")
    if servico:
        _par(doc, f"{n}.1 O serviço poderá ser rejeitado, no todo ou em parte, quando em desacordo com as especificações "
                  "deste Termo de Referência e da proposta, devendo ser adequado no prazo de 10 (dez) dias da notificação.")
    else:
        _par(doc, f"{n}.1 O prazo de entrega é de até 30 (trinta) dias corridos, contados do recebimento da Autorização "
                  "de Compras, em remessa única.")
    _par(doc, f"{n}.2 Os preços devem incluir fretes, seguros, impostos, taxas e demais custos diretos ou indiretos.")
    if _item_saude(analise):
        _par(doc, f"{n}.3 Na entrega, a validade do produto deve ser de no mínimo 75% do prazo total de validade.")
        _par(doc, f"{n}.4 As embalagens primárias e/ou secundárias devem conter a expressão “PROIBIDA A VENDA NO COMÉRCIO”. "
                  "Em embalagens estéreis não serão aceitos carimbos alcoólicos; preconiza-se etiqueta inviolável.")
        _par(doc, f"{n}.5 A Nota Fiscal deve trazer lote e validade de cada item entregue.")
    n += 1

    if not servico and not _medicamento(analise):
        _titulo(doc, f"{n}. DA GARANTIA DO OBJETO")
        _par(doc, f"{n}.1 Garantia contra defeitos de fabricação por, no mínimo, 12 (doze) meses a partir do aceite.")
        n += 1
    if servico:
        _titulo(doc, f"{n}. DA GARANTIA DO SERVIÇO")
        _par(doc, f"{n}.1 O serviço e as peças repostas terão garantia mínima de 3 (três) meses a partir do aceite, ou o "
                  "prazo do fabricante, se maior.")
        n += 1

    _titulo(doc, f"{n}. DO PAGAMENTO")
    _par(doc, f"{n}.1 Mediante apresentação de Nota Fiscal devidamente atestada, em favor da {un['razao']}, CNPJ "
              f"{un['cnpj']}, com a menção “Aquisição e/ou contratação com recursos do PDPAS”"
              + (" e os números de patrimônio atendidos." if bens else "."))
    _par(doc, f"{n}.2 Os pagamentos iguais ou superiores a R$ 5.000,00 serão feitos, exclusivamente, por crédito em "
              "conta corrente no Banco de Brasília S/A - BRB, ressalvadas as exceções do art. 13, § 4º, da Portaria 473/2023.")
    for t in _verificacao_fornecedor(str(n)):
        _par(doc, t)
    n += 1

    _titulo(doc, f"{n}. DO REAJUSTE")
    _par(doc, f"{n}.1 O valor do objeto desta dispensa de licitação é fixo e irreajustável.")
    n += 1

    _titulo(doc, f"{n}. DAS SANÇÕES APLICÁVEIS")
    _par(doc, f"{n}.1 A empresa vencedora estará sujeita às sanções administrativas previstas nos artigos 155 a 163 da "
              "Lei nº 14.133/2021.")
    n += 1
    _titulo(doc, f"{n}. DAS ASSINATURAS E APROVAÇÃO")
    _par(doc, APROVACAO)
    _par(doc, "[A PREENCHER: nome e matrícula dos responsáveis]")
    arq = destino / f"2_Termo_de_Referencia_{_nome_arquivo(pedido, analise)}.docx"
    doc.save(arq)
    return arq


AMOSTRA_FISICA = ("Além da fotografia, antes da entrega a empresa deverá apresentar amostra física do produto ofertado, para avaliação da área "
                  "demandante quanto à conformidade com as especificações deste Termo de Referência; a entrega somente será "
                  "autorizada após a aprovação da amostra.")


def _clausulas_permanente(doc, regras: dict, un: dict, pedido: dict | None = None) -> None:
    """Itens 6 a 19 do TR 1/2026, corrigidos: um só marco dos 30 dias, garantia em vez de 'validade',
    10 dias para material rejeitado na entrega e 30 dias para defeito depois do aceite.
    A sanção remete à Lei 14.133, cujo art. 156, § 4º, limita o impedimento a 3 anos (o TR falava em 5)."""
    sigla = (regras.get("entrega") or {}).get("permanente_sigla") or "NPDOC"
    endereco = (regras.get("entrega") or {}).get("permanente") or "[A PREENCHER: endereço de entrega de bem permanente]"
    razao, cnpj = un["razao"], un["cnpj"]
    pedido = pedido or {}
    amostra = ["6.1 Solicitamos o envio de uma fotografia do produto ofertado, a título de amostra, para conferência e "
               "validação quanto à conformidade com as especificações apresentadas. O encaminhamento da imagem é necessário "
               "para subsidiar a análise e dar prosseguimento ao processo de aquisição."]
    if pedido.get("amostra_fisica"):
        amostra.append(f"6.2 {AMOSTRA_FISICA}")
    secoes = [
        ("6. DA AMOSTRA", amostra),
        ("7. FORNECIMENTO DOS BENS E INSTRUMENTO DE AJUSTE", [
            "7.1 DO FORNECIMENTO DOS PRODUTOS",
            "7.1.1 Para fornecimento do objeto do presente Instrumento, deverão ser observadas as especificações "
            "técnicas estabelecidas na descrição do objeto;",
            "7.1.2 Os objetos deverão ser entregues com a respectiva nota fiscal, para ateste e liquidação da despesa;",
            "7.1.3 Os produtos serão fornecidos de forma integral, com o certificado de calibração devidamente aprovado "
            "pelo INMETRO, quando for o caso;",
            "7.1.4 Os equipamentos deverão ser fornecidos novos, de primeiro uso, em perfeito estado e plenas condições "
            "de funcionamento e conservação, devidamente embalados, montados e prontos para uso; os itens que exijam "
            "montagem serão entregues montados e disponíveis para uso, sem ônus adicional para a Administração;",
            "7.1.5 A contratada deverá fornecer manual de uso e instruções de conservação, quando aplicável."]),
        ("8. DO INSTRUMENTO DE AJUSTE", [
            "8.1 A Autorização de Compra substituirá a nota de empenho, conforme art. 233 do Decreto nº 44.330/2023."]),
        ("9. PRAZO DE ENTREGA E RECEBIMENTO", [
            "9.1 Entrega em até 30 (trinta) dias, contados do recebimento da Autorização de Compra;",
            f"9.2 Os materiais permanentes, objeto deste Termo de Referência, deverão ser entregues no {sigla}, no horário "
            f"compreendido de 8h às 12h e de 13h às 17h. Endereço: {endereco}.",
            "9.3 Os bens poderão ser rejeitados, no todo ou em parte, quando em desacordo com as especificações constantes "
            "neste Termo de Referência e na proposta.",
            "9.4 Somente após a aprovação pela área demandante, os bens entregues serão recebidos definitivamente.",
            "9.5 Caso necessite de instalação, a contratada deverá instalar o bem em até 3 (três) dias após o acionamento "
            "pela Administração."]),
        ("10. DEVERES E RESPONSABILIDADES DA UNIDADE SOLICITANTE", [
            "10.1 Receber o objeto, no prazo e condições estabelecidas na Autorização de Compra e seus anexos, se houver;",
            "10.2 Verificar minuciosamente a conformidade dos bens recebidos com as especificações constantes da "
            "Autorização de Compra, da proposta e seus anexos, para fins de aceitação e recebimento;",
            f"10.3 Comunicar à empresa vencedora da dispensa, por escrito, no prazo de 5 dias úteis do aceite do produto "
            f"pelo {sigla}, sobre imperfeições, falhas ou irregularidades verificadas no objeto fornecido, para que seja "
            "substituído, reparado ou corrigido;",
            "10.4 Efetuar o pagamento da Nota Fiscal atestada, no valor correspondente ao fornecimento do objeto, no prazo "
            "e forma estabelecidos;",
            "10.5 Providenciar registro no Sistema Geral de Patrimônio (SisGepat) e incorporação do bem ao patrimônio da "
            "Unidade."]),
        ("11. DEVERES E RESPONSABILIDADES DA EMPRESA VENCEDORA DA DISPENSA DE LICITAÇÃO", [
            "11.1 Cumprir todas as obrigações constantes do Termo de Referência, do Pedido de Cotação, de seus anexos e da "
            "proposta apresentada, assumindo como exclusivamente seus os riscos e as despesas decorrentes da boa e perfeita "
            "execução do objeto;",
            "11.2 Efetuar a entrega do objeto em perfeitas condições, conforme especificações, prazo e local constantes no "
            "Termo de Referência, no Pedido de Cotação e em seus anexos, acompanhado da respectiva nota fiscal, na qual "
            "constarão as indicações referentes à marca, ao fabricante, ao modelo, à procedência e ao prazo de garantia;",
            "11.3 O objeto deve estar acompanhado do manual do usuário, com uma versão em português, e da relação da rede de "
            "assistência técnica autorizada;",
            "11.4 Responsabilizar-se pelos vícios e danos decorrentes do objeto, de acordo com os artigos 12, 13 e 17 a 27 "
            "do Código de Defesa do Consumidor (Lei nº 8.078, de 1990);",
            "11.5 Substituir, reparar ou corrigir, às suas expensas, no prazo de 30 dias a partir da notificação, o objeto "
            "com avarias ou defeitos constatados após o aceite;",
            "11.6 Comunicar à Unidade solicitante, em prazo que antecede a data da entrega, os motivos que impossibilitem o "
            "cumprimento do prazo previsto, com a devida justificativa e comprovação;",
            "11.7 Manter, durante todo o processo, em compatibilidade com as obrigações assumidas, todas as condições de "
            "habilitação e qualificação exigidas na dispensa de licitação."]),
        ("12. DA DISPONIBILIDADE DE RECURSOS", [
            f"12.1 Recurso financeiro do Programa de Descentralização Progressiva de Ações de Saúde - PDPAS, disponibilizado "
            f"em conta bancária da {razao}, CNPJ {cnpj}."]),
        ("13. DA VIGÊNCIA, DOS PRAZOS E DAS CONDIÇÕES", [
            "13.1 O prazo para entrega do produto é de até 30 (trinta) dias corridos, contados do recebimento da Autorização "
            "de Compras, em remessa única;",
            "13.2 Os bens rejeitados na entrega, quando em desacordo com as especificações constantes neste Termo de "
            "Referência e na proposta, deverão ser substituídos no prazo de 10 (dez) dias, a contar da notificação;",
            "13.3 Dentro dos preços constantes da Dispensa de Licitação já devem estar inclusos os valores de fretes, "
            "calibração/aferição dos equipamentos, seguros, impostos, taxas e outros que possam influenciar direta ou "
            "indiretamente nos preços dos itens solicitados."]),
        ("14. DA GARANTIA DO OBJETO", [
            f"14.1 Os materiais fornecidos devem contar com garantia contra defeitos de fabricação, por um período mínimo de "
            f"12 (doze) meses contado a partir da data de aceite do produto pelo {sigla}, e dispor de assistência técnica "
            "autorizada pelo fabricante no Distrito Federal.",
            "14.2 Dentro do prazo de garantia, a unidade solicitante poderá notificar a empresa fornecedora para providências "
            "de reparo ou substituição. Nesse caso, sendo confirmada a necessidade, a empresa terá 30 dias para busca e "
            "devolução ou entrega de novo produto.",
            "14.3 No ato da entrega, a empresa fornecedora deverá apresentar o certificado de garantia, que poderá ser "
            "disponibilizado por meio de documento próprio ou mediante anotação impressa ou carimbada na respectiva Nota "
            "Fiscal, devendo constar, ainda, a relação da rede de assistência técnica autorizada.",
            "14.4 O fabricante deverá disponibilizar canal de atendimento, por meio de número telefônico sediado em Brasília "
            "e/ou serviço 0800, bem como endereço eletrônico (e-mail), destinado à prestação de suporte técnico.",
            "14.5 As peças que apresentarem vício ou defeito durante o período de vigência da garantia deverão ser substituídas "
            "por componentes novos, originais e de primeiro uso, que atendam aos mesmos padrões de qualidade e desempenho, "
            "iguais ou superiores aos das peças originais utilizadas na fabricação do equipamento, sem custos adicionais."]),
        ("15. DO PAGAMENTO", [
            f"15.1 Mediante apresentação de Nota Fiscal da empresa vencedora da dispensa de licitação - PDPAS, devidamente "
            f"atestada e em favor da {razao}, CNPJ {cnpj}, com a menção “Aquisição e/ou contratação com recursos do PDPAS”.",
            "15.2 Os pagamentos de valores iguais ou superiores a R$ 5.000,00 (cinco mil reais) serão feitos, exclusivamente, "
            "mediante crédito em conta corrente, em nome do beneficiário, junto ao Banco de Brasília S/A - BRB, excluindo-se "
            "destas disposições os seguintes casos: a) os pagamentos a empresas vinculadas ou supervisionadas pela "
            "Administração Pública Federal; b) os pagamentos a empresas de outros Estados da Federação que não mantenham "
            "filiais e/ou representações no DF.", *_verificacao_fornecedor("15")]),
        ("16. DO REAJUSTE", ["16.1 O valor do objeto desta dispensa de licitação é fixo e irreajustável."]),
        ("17. MECANISMOS FORMAIS DE COMUNICAÇÃO", ["17.1 Será utilizado o envio de mensagens eletrônicas (e-mail) ou carta."]),
        ("18. DAS SANÇÕES APLICÁVEIS", [
            "18.1 Aquele que, convocado dentro do prazo de validade de sua proposta, deixar de entregar documentação exigida "
            "no pedido de cotação e seus anexos, apresentar documentação falsa, ensejar o retardamento da execução de seu "
            "objeto, não mantiver a proposta, falhar ou fraudar, comportar-se de modo inidôneo, fizer declaração falsa ou "
            "cometer fraude fiscal, garantido o direito à ampla defesa, ficará sujeito às sanções da Lei nº 14.133/2021, sem "
            "prejuízo da aplicação de multa;",
            "18.2 No caso de não cumprimento do apresentado na proposta, a empresa vencedora da dispensa de licitação - PDPAS "
            "estará sujeita às sanções administrativas previstas nos artigos 155 a 163 da Lei nº 14.133/2021."]),
        ("19. DAS ASSINATURAS", [APROVACAO, "Responsável pela Unidade Solicitante (elaborou este Termo de Referência e "
                                 "acompanha o processo até a entrega final): [A PREENCHER: nome e matrícula]",
                                 *_assinaturas_permanente(pedido, regras)]),
    ]
    for titulo, paragrafos in secoes:
        _titulo(doc, titulo)
        for t in paragrafos:
            _par(doc, t)


def _ponto(t: str) -> str:
    t = t.strip()
    return t if t.endswith((".", "!", "?", "]")) else t + "."


def _ou_campo(pedido: dict, chave: str, falta: str) -> str:
    return str(pedido.get(chave) or "").strip() or f"[A PREENCHER: {falta}]"



def parecer_demandante(pedido: dict, analise: dict, regras: dict, destino: Path, hoje: date | None = None) -> Path:
    """Fase 6, depois do mapa, no formato do Parecer Técnico 56/2026 do NLF Norte (SEI 218235374)."""
    resumo, _ = _desc(pedido, analise)
    cod = analise.get("codigo") or "-"
    total = motor.somar_quantidades(pedido)[0]
    doc = _doc(regras)
    _par(doc, f"Parecer Técnico - {regras['unidade']['sigla']}", negrito=True, justificar=False)
    _par(doc, "À SES/SRSNO/DA/GEOF", justificar=False)
    _par(doc, "Assunto: Parecer Técnico referente ao PDPAS nº [A PREENCHER: número do PDPAS]", negrito=True, justificar=False)
    _par(doc, "Após a análise das propostas apresentadas, do Mapa Comparativo ([A PREENCHER: nº SEI]) e das pesquisas de "
              "preços ([A PREENCHER: nº SEI]), a empresa [A PREENCHER: vencedora] apresentou o menor preço para "
              f"{'o serviço' if analise.get('tipo') == 'servico' else 'o item'} {_rotulo(analise, resumo)}: valor unitário de [A PREENCHER: valor], frente à mediana de [A PREENCHER: mediana].")
    _par(doc, "A proposta ([A PREENCHER: nº SEI]) atende às especificações do Termo de Referência ([A PREENCHER: nº SEI]), "
              "conforme catálogo ou fotografia enviados pela empresa.")
    _par(doc, "[A PREENCHER: se a vencedora não for a de menor preço, justificar a escolha (art. 16, VI)]")
    _par(doc, "[A PREENCHER: se a proposta estiver acima da mediana da pesquisa de preços, justificar o prosseguimento e o "
              "impacto do fracasso da compra]")
    _tabela(doc, ["Item", "Descrição", "Quantidade", "Código", "Vencedora", "Marca", "Entrega"],
            [["1", resumo, _qtd(total) if total else "[A PREENCHER: quantidade]", cod, "[A PREENCHER]", "[A PREENCHER]", "[A PREENCHER]"]],
            [1, 5.4, 1.9, 1.9, 2.3, 1.9, 1.6])
    _par(doc, "Diante do exposto, esta área manifesta-se favorável à aquisição.")
    _par(doc, "Atenciosamente,", justificar=False)
    arq = destino / f"5_Parecer_apos_mapa_{_nome_arquivo(pedido, analise)}.docx"
    doc.save(arq)
    return arq


def atesto(pedido: dict, analise: dict, regras: dict, destino: Path, hoje: date | None = None) -> Path:
    """Fase 8: declaração do Memorando Circular 1/2025 da GEOF Norte (SEI 174878263)."""
    tipo = analise["tipo"]
    frase = {"consumo": "o material nela descrito foi recebido em", "permanente": "o bem patrimonial nela descrito foi recebido em",
             "servico": "o serviço nela descrito foi totalmente executado em"}[tipo]
    doc = _doc(regras)
    _par(doc, f"Atesto - {regras['unidade']['sigla']}", negrito=True, justificar=False)
    _par(doc, f"Atesto, para os devidos fins, que a Nota Fiscal nº [A PREENCHER: número], anexa ao presente processo no "
              f"Documento SEI nº [A PREENCHER: nº SEI], foi conferida e {frase} [A PREENCHER: data]. Declaro, ainda, que se "
              "trata de aquisição e/ou contratação com recursos do PDPAS.")
    if tipo == "permanente":
        sigla = (regras.get("entrega") or {}).get("permanente_sigla") or "NPDOC"
        _par(doc, f"O bem foi avaliado presencialmente no {sigla} pela área demandante e está de acordo com as especificações "
                  "do Termo de Referência.")
        _par(doc, "Seguem as fotografias do bem ([A PREENCHER: nº SEI]) e a identificação do servidor responsável pelo "
                  "patrimônio: [A PREENCHER: nome e matrícula]. Localização: [A PREENCHER: unidade].")
    arq = destino / f"6_Atesto_{_nome_arquivo(pedido, analise)}.docx"
    doc.save(arq)
    return arq


def relatorio_checklist(pedido: dict, analise: dict, regras: dict, destino: Path, hoje: date | None = None) -> Path:
    hoje = hoje or date.today()
    resumo, _ = _desc(pedido, analise)
    doc = _doc(regras)
    _par(doc, f"Conferência do processo PDPAS: {resumo}", negrito=True, justificar=False)
    enq = analise["enquadramento"]
    _par(doc, f"Código: {analise.get('codigo') or '-'}  |  Tipo: {motor.NOME_TIPO.get(analise['tipo'], analise['tipo'])}  |  Destino: "
              f"{enq.get('destino') or '-'}  |  Despesa: {enq.get('despesa') or '-'}  |  Gerado em {hoje:%d/%m/%Y}",
         justificar=False)
    _titulo(doc, "Linha para a planilha de acompanhamento (assessoria da diretoria)")
    _tabela(doc, ["Nº do processo do memorando de intenção", "Item", "Setor de origem", "Data do despacho", "Fim do prazo"],
            [[_proc_intencao(pedido), _rotulo(analise, resumo),
              _origem(pedido), f"{hoje:%d/%m/%Y}", _ou_campo(pedido, "prazo_manifestacao", "data-limite")]],
            [3.6, 5.2, 3, 2.1, 2.1])
    _titulo(doc, "Alertas")
    _tabela(doc, ["Nível", "Regra", "O que fazer"],
            [[{"BLOQUEIO": "Bloqueio", "ATENCAO": "Atenção", "INFO": "Orientação"}[a["severidade"]], a["regra"], a["mensagem"]]
             for a in analise["alertas"]] or [["-", "-", "Sem alertas"]], [2.2, 3.3, 10.5])
    _titulo(doc, "Dúvida em aberto")
    duv = doc.add_paragraph().add_run(motor.DUVIDA_ORDEM)
    duv.font.highlight_color = WD_COLOR_INDEX.YELLOW
    _titulo(doc, "Peças do processo")
    _par(doc, "Quem abre o processo responde só pelas linhas do setor de origem. As demais aparecem para mostrar o caminho: "
              "o processo vai andando de um responsável para o outro.", justificar=False)
    lista = motor.checklist(analise["tipo"], pedido)
    ordem = 0
    for fase in motor.FASES:
        linhas = [c for c in lista if c["fase"] == fase]
        if linhas:
            ordem += 1      # consumo e serviço não têm a fase do parecer central: renumera para não saltar do 2 para o 4
            _par(doc, f"{ordem}. {fase.split('. ', 1)[-1]}", negrito=True, justificar=False)
            _tabela(doc, ["OK", "Quem faz essa parte?", "Peça", "Base"],
                    [["☐", c["responsavel"], c["peca"], c["base"]] for c in linhas], [1, 3.2, 8.3, 3.5])
    arq = destino / f"3_Checklist_{_nome_arquivo(pedido, analise)}.docx"
    doc.save(arq)
    return arq


def _marcar_nao_validados(pedido: dict) -> dict:
    """Texto sugerido que ela não validou sai marcado na minuta; devolve uma cópia do pedido."""
    pedido = dict(pedido)
    for campo in pedido.get("nao_validados") or []:
        texto = str(pedido.get(campo) or "").strip()
        if campo in TEXTUAIS and texto:
            pedido[campo] = "[VALIDAR: " + texto.replace("[", "").replace("]", "") + "]"
    return pedido


def gerar_peca(chave: str, pedido: dict, analise: dict, regras: dict, destino: Path, hoje: date | None = None) -> Path:
    """Gera (ou refaz) uma única peça na pasta do pedido: é o que o fluxo usa em cada passo."""
    pecas_por_chave = {"memo_intencao": memo_intencao, "despacho_intencao": despacho_intencao,
                       "despacho_consolidado": despacho_consolidado, "termo_referencia": termo_referencia,
                       "parecer_demandante": parecer_demandante, "atesto": atesto, "checklist": relatorio_checklist}
    if chave not in pecas_por_chave:
        raise ValueError("Peça desconhecida.")
    if not analise.get("pode_gerar"):
        raise ValueError("Escolha o tipo do pedido antes de gerar as peças.")
    destino = Path(destino)
    destino.mkdir(parents=True, exist_ok=True)
    return pecas_por_chave[chave](_marcar_nao_validados(pedido), analise, regras, destino, hoje)


def gerar_tudo(pedido: dict, analise: dict, regras: dict, destino: Path, hoje: date | None = None) -> list[Path]:
    if not analise.get("pode_gerar"):
        raise ValueError("Escolha o tipo do pedido antes de gerar as peças.")
    destino = Path(destino)
    destino.mkdir(parents=True, exist_ok=True)
    pedido = _marcar_nao_validados(pedido)
    arquivos: list[Path] = []
    try:
        arquivos += [memo_intencao(pedido, analise, regras, destino, hoje),
                     termo_referencia(pedido, analise, regras, destino, hoje),
                     relatorio_checklist(pedido, analise, regras, destino, hoje)]
        arquivos += [despacho_intencao(pedido, analise, regras, destino, hoje),
                     despacho_consolidado(pedido, analise, regras, destino, hoje)]
        arquivos += [parecer_demandante(pedido, analise, regras, destino, hoje), atesto(pedido, analise, regras, destino, hoje)]
    except Exception:
        for a in arquivos:      # falha na 3ª peça não deixa as duas primeiras soltas numa pasta "pronta"
            Path(a).unlink(missing_ok=True)
        raise
    return arquivos
