"""Sugestões de texto para os campos livres, por regra, sem IA. Quem preenche valida cada uma.

Os grupos dizem para que o item serve na APS. A frase é genérica de propósito: sem dado do histórico ou da
Visão Gerencial, a sugestão não afirma falta de estoque nem quantidade. O texto validado em processo anterior
do mesmo código tem prioridade sobre estes modelos.
"""
from __future__ import annotations

import json
import re
from datetime import date, datetime, timedelta
from pathlib import Path

ARQ_ITENS = Path(__file__).resolve().parent / "dados" / "textos_itens.json"
_ITENS: dict | None = None


def texto_do_item(codigo: str) -> dict | None:
    """Texto próprio de cada código, gerado uma vez por gerar_textos_itens.py e guardado com o programa."""
    global _ITENS
    if _ITENS is None:
        try:
            _ITENS = json.loads(ARQ_ITENS.read_text(encoding="utf-8-sig")) if ARQ_ITENS.exists() else {}
        except (OSError, ValueError):       # arquivo regenerado com defeito: cai nos grupos genéricos, sem derrubar a sugestão
            _ITENS = {}
    p = _ITENS.get(str(codigo or "").strip().upper())
    # só vale com os três textos (u = uso, f = falta, b = benefício) presentes e em texto
    return p if isinstance(p, dict) and all(isinstance(p.get(k), str) and p[k].strip() for k in ("u", "f", "b")) else None

# (padrão na descrição sem acento, uso, o que a falta compromete, benefício)
GRUPOS = [
    (r"\bCURATIVO|COBERTURA|HIDROCOLOIDE|ALGINATO|ATADURA|COMPRESSA|GAZE|MICROPORE|ESPARADRAPO|FITA CIRURGICA",
     "no cuidado de feridas e em procedimentos de enfermagem", "o tratamento de feridas e a prevenção de infecções",
     "o cuidado adequado das feridas e a redução de complicações"),
    (r"TESTE RAPIDO|IMUNOCROMATOGRAF", "na testagem rápida e no diagnóstico oportuno",
     "o diagnóstico oportuno e o início do tratamento", "a testagem rápida na própria unidade e o tratamento oportuno"),
    (r"COLETA|TUBO PARA|COLETOR|LANCETA|FITA COM AREA REAGENTE|TIRA REAGENTE|LAMINA|PORTA-LAMINAS|ESCOVA CERVICAL|ESPATULA",
     "na coleta de exames e no monitoramento de usuários", "a coleta de exames e o acompanhamento dos usuários",
     "a coleta de exames na própria unidade e o acompanhamento dos usuários"),
    (r"ESPECULO|HISTEROMETRO|PRESERVATIVO|DIU\b|IMPLANTE",
     "na saúde sexual e reprodutiva e na prevenção do câncer do colo do útero",
     "o acesso à saúde sexual e reprodutiva e o rastreamento do câncer do colo do útero",
     "o acesso das usuárias e dos usuários à saúde sexual e reprodutiva e ao rastreamento"),
    (r"AGULHA|SERINGA|ESCALPE|CATETER|EQUIPO|SORO|CLORETO DE SODIO",
     "na administração de medicamentos e em procedimentos de enfermagem",
     "a administração de medicamentos e a resolutividade das unidades",
     "a realização segura dos procedimentos nas unidades"),
    (r"LUVA|MASCARA CIRURGICA|AVENTAL|TOUCA|OCULOS|PROTETOR|CAPOTE|LENCOL",
     "na proteção de profissionais e usuários", "a segurança de profissionais e usuários durante a assistência",
     "a proteção de profissionais e usuários e a prevenção de infecções"),
    (r"GRAU CIRURGICO|DETERGENTE ENZIMATICO|AUTOCLAVE|SELADORA|INDICADOR BIOLOGICO|INTEGRADOR",
     "no processamento e na esterilização de materiais", "a esterilização de materiais e a segurança dos procedimentos",
     "o processamento seguro dos materiais usados na assistência"),
    (r"ALCOOL|HIPOCLORITO|CLOREXIDINA|DESINFETANTE|SABONETE", "na higienização e na desinfecção",
     "a higienização das mãos, das superfícies e dos materiais", "a prevenção de infecções relacionadas à assistência"),
    (r"ESPACADOR|NEBULIZ|CANULA NASAL|OXIGENIO|REINALANTE|UMIDIFICADOR|AMBU|REANIMADOR",
     "no atendimento de usuários com condições respiratórias", "o atendimento de crises respiratórias nas unidades",
     "o atendimento respiratório oportuno e a redução de encaminhamentos evitáveis"),
    (r"LARINGOSCOPIO|DESFIBRILADOR|CARDIOVERSOR|CARRO DE EMERGENCIA|ASPIRADOR|MONITOR MULTI",
     "no atendimento de urgências e intercorrências", "a resposta da equipe a urgências e intercorrências clínicas",
     "a resposta rápida e segura às urgências nas unidades"),
    (r"ESFIGMO|TENSIOMETRO|ESTETOSCOPIO|OTOSCOPIO|OFTALMOSCOPIO|DOPPLER|BALANCA|GLICOSIMETRO|TERMOMETRO|OXIMETRO|ANTROPOMETRO|ELETROCARDIOGRAFO",
     "na avaliação clínica e no acompanhamento de usuários", "a avaliação clínica e o acompanhamento das condições crônicas",
     "a avaliação clínica adequada e o acompanhamento longitudinal"),
    (r"PINCA|TESOURA|CABO DE BISTURI|PORTA AGULHA|CUBA|BANDEJA|AFASTADOR|CURETA", "em pequenos procedimentos e curativos",
     "a realização de pequenos procedimentos na unidade", "a resolutividade da APS em pequenos procedimentos"),
    (r"HALTER|CANELEIRA|TAPETE|COLCHONETE|FAIXA ELASTICA|BOLA", "em práticas corporais e reabilitação",
     "as atividades de práticas corporais e reabilitação", "a oferta regular de práticas corporais e reabilitação"),
    (r"AMOXICILINA|AZITROMICINA|CEFALEXINA|CIPROFLOXACINO|CLAVULANATO|METRONIDAZOL|NITROFURANTOINA|BENZILPENICILINA",
     "no tratamento de infecções", "o tratamento oportuno de infecções", "o tratamento oportuno na própria APS"),
    (r"ANLODIPINO|LOSARTANA|ENALAPRIL|CAPTOPRIL|HIDROCLOROTIAZIDA|CARVEDILOL|ATENOLOL|FUROSEMIDA|METFORMINA|GLIBENCLAMIDA|INSULINA|SINVASTATINA|HIDRALAZINA",
     "no cuidado de condições crônicas", "o controle de hipertensão, diabetes e outras condições crônicas",
     "a continuidade do tratamento das condições crônicas"),
    (r"SALBUTAMOL|BECLOMETASONA|BUDESONIDA|IPRATROPIO", "no tratamento de asma e outras doenças respiratórias",
     "o controle de asma e de doenças respiratórias", "o controle das doenças respiratórias e a redução de crises"),
    (r"DIPIRONA|PARACETAMOL|IBUPROFENO|TENOXICAM|BROMOPRIDA|ONDANSETRONA|BUTILBROMETO|PREDNISONA",
     "no manejo de sintomas e de condições agudas", "o atendimento de demandas espontâneas nas unidades",
     "o alívio de sintomas e a resolutividade do atendimento"),
]
PADRAO = {
    "0": ("no tratamento de usuários atendidos na APS", "a continuidade do tratamento dos usuários", "a continuidade do tratamento na APS"),
    "1": ("na assistência às pessoas atendidas", "a realização dos procedimentos assistenciais", "a continuidade e a segurança da assistência"),
    "11": ("na assistência às pessoas atendidas", "a realização dos procedimentos assistenciais", "a continuidade e a segurança da assistência"),
    "2": ("na rotina das unidades", "o funcionamento regular das unidades", "o funcionamento regular das unidades"),
    "3": ("no atendimento odontológico", "o acesso à saúde bucal e os procedimentos odontológicos", "a continuidade dos procedimentos odontológicos"),
    "4": ("nos exames laboratoriais", "a realização de exames e o diagnóstico", "a realização de exames e o diagnóstico oportuno"),
    "5": ("no cuidado nutricional", "o cuidado nutricional dos usuários", "a continuidade do cuidado nutricional"),
    "6": ("em procedimentos especializados", "a realização dos procedimentos", "a continuidade dos procedimentos"),
    "9": ("na assistência prestada nas unidades", "a capacidade de atendimento das unidades", "a capacidade e a qualidade do atendimento"),
    "10": ("na estrutura das unidades", "as condições de trabalho e de acolhimento", "as condições de trabalho e de acolhimento dos usuários"),
}


_PREFIXOS = {"NEBULIZ", "IMUNOCROMATOGRAF", "ESFIGMO", "MONITOR MULTI", "TUBO PARA", "GRAU CIRURGICO", "FITA COM AREA REAGENTE"}


def _termo_re(t: str) -> str:
    t = t.replace(r"\b", "")
    return rf"{t}\w*" if t in _PREFIXOS else rf"{t}(?:S|ES|AS|OS)?\b"


_GRUPOS_RE = [(re.compile(r"\b(?:" + "|".join(_termo_re(t) for t in p.split("|")) + ")"), u, r, b) for p, u, r, b in GRUPOS]


def _grupo(desc_sem_acento: str, categoria: str) -> tuple[str, str, str]:
    # cada termo é palavra inteira (com plural): "LAMINA" não casa em "ESCLAMINA", "BOLA" não casa em "BOLACHA";
    # só os prefixos declarados (NEBULIZ -> NEBULIZADOR) casam o começo da palavra
    for padrao, uso, risco, beneficio in _GRUPOS_RE:
        if padrao.search(desc_sem_acento):
            return uso, risco, beneficio
    return PADRAO.get(categoria, PADRAO["2"])


def _nome(desc: str) -> str:
    d = (desc or "o item").split(",")[0].strip()
    return d[:1].upper() + d[1:].lower() if d.isupper() else d


# Feriados do DF que caem nos prazos de fim de ano; o regras.json ("intencao.feriados") sobrescreve.
FERIADOS_PADRAO = ("2026-11-20", "2026-11-30", "2026-12-25", "2027-01-01")


def ler_feriados(feriados) -> tuple[list[date], list[str]]:
    """Aceita AAAA-MM-DD e dd/mm/aaaa (o regras.json é editado à mão). Devolve as datas e o que não deu para ler:
    antes, feriado em formato brasileiro era ignorado em silêncio e o prazo saía em dia de feriado."""
    datas, ilegiveis = [], []
    for f in feriados or ():
        txt = str(f).strip()
        for ler in (date.fromisoformat, lambda t: datetime.strptime(t, "%d/%m/%Y").date()):
            try:
                datas.append(ler(txt))
                break
            except ValueError:
                continue
        else:
            ilegiveis.append(txt)
    return datas, ilegiveis


def prazo_gsaps(hoje: date, dias: int = 7, feriados=FERIADOS_PADRAO) -> str:
    """Dias ÚTEIS (sem sábado, domingo e feriado) a partir de `hoje`: a Diretoria dá 7 para as gerências se posicionarem."""
    fer = set(ler_feriados(feriados)[0])
    d = hoje
    while dias > 0:
        d += timedelta(days=1)
        if d.weekday() < 5 and d not in fer:
            dias -= 1
    return d.strftime("%d/%m/%Y")


def para_item(desc: str, desc_sem_acento: str, categoria: str, tipo: str, cmm: int | None, visao: dict | None, hoje: date,
              codigo: str = "") -> dict:
    uso, risco, beneficio = _grupo(desc_sem_acento, categoria)
    nome = _nome(desc)
    proprio = texto_do_item(codigo)
    if proprio:
        risco, beneficio = proprio["f"].rstrip(". "), re.sub(r"^[Aa]ssegurar\s+", "", proprio["b"].rstrip(". "))
        u = proprio["u"].rstrip(". ")
        just = (f"{nome}: {u[:1].lower()}{u[1:]}. A falta deste item nas unidades de Atenção Primária da Região Norte compromete {risco}, "
                "o que caracteriza a relevância da aquisição para a saúde da população atendida.")
    else:
        just = (f"O item {nome.lower()} é usado {uso} nas unidades de Atenção Primária da Região Norte. A indisponibilidade compromete {risco}, "
                "o que caracteriza a relevância da aquisição para a saúde da população atendida.")
    if cmm:
        just += f" O consumo médio da APS Norte é de {cmm:,} por mês.".replace(",", ".")
    campos = {"justificativa": just,
              "beneficios": f"Assegurar {beneficio}, com continuidade da assistência nas unidades da Região Norte.",
              "prazo_manifestacao": prazo_gsaps(hoje),
              # frases de fato sem fonte viram campo a preencher (decisão de 05/11/2026)
              "just_acima_cmm": f"A quantidade busca evitar nova interrupção de {nome.lower()}, considerando que "
                                "[A PREENCHER: situação da reposição pelo abastecimento central].",
              "justificativa_ata": f"Embora exista ata vigente, [A PREENCHER: por que a entrega de {nome.lower()} pela compra central "
                                   f"não atende à necessidade imediata das unidades], o que compromete {risco}."}
    if visao:
        # None = o número não foi lido na Visão (PDF girado, linha quebrada); não é zero. motor.art21_da_visao já respeita
        # isso desde 03/11; o texto sugerido afirmava "estoque 0" e "não há ata" sem fonte.
        data = "/".join(reversed((visao.get("data") or "").split("-")))
        rede, central = visao.get("estoque_rede"), visao.get("central")
        if rede is None or central is None:
            partes = [f"Na Visão Gerencial de {data}: [A PREENCHER: estoque na rede e nos locais centrais]."]
        else:
            partes = [f"Na Visão Gerencial de {data}, o estoque na rede era de {rede:g} e o dos locais centrais, de {central:g}."]
        if visao.get("atas"):
            a = visao["atas"][0]
            partes.append(f"Há a ata {a['ata']}, válida até {a['validade']}.")
        elif visao.get("atas_incerto"):
            partes.append("[A PREENCHER: situação das atas de registro de preços, não lida na Visão]")
        else:
            partes.append("Não há ata de registro de preços vigente.")
        if visao.get("abertos"):
            partes.append(f"Há pedido em aberto no processo {visao['abertos'][0]['processo']}.")
        campos["contexto"] = " ".join(partes)
    else:
        campos["contexto"] = f"O item {nome.lower()} é padronizado pela SES-DF. A situação do estoque e das atas será comprovada pelos relatórios do sistema de materiais."
    if tipo == "permanente":
        campos["just_quantidade"] = "Quantitativo definido a partir das manifestações das unidades, com substituição de bens inservíveis ou atendimento de nova necessidade."
    else:
        campos["just_quantidade"] = "Quantitativo definido a partir das manifestações das unidades e da demanda assistencial."
    return campos


def para_servico(pedido: dict, bens: list[dict], hoje: date) -> dict:
    predio = pedido.get("altera_predio") or pedido.get("engenharia")
    if pedido.get("reparo_equipamento") or bens:
        nome = bens[0]["descricao"].split(",")[0] if bens and bens[0]["descricao"] else "equipamento"
        qtd = f"{len(bens)} bens" if len(bens) > 1 else nome
        objeto = f"Manutenção corretiva de {qtd}"
        campos = {"objeto": objeto, "descricao": f"{objeto}, com avaliação técnica e reposição de peças, se necessário, conforme os patrimônios listados.",
                  "justificativa": f"O equipamento{f' ({nome})' if bens else ''} está sem funcionamento adequado e [A PREENCHER: situação do contrato de manutenção], o que compromete a assistência prestada na unidade.",
                  "contexto": "A unidade depende do equipamento para a rotina assistencial e [A PREENCHER: por que o reparo é preferível à reposição].",
                  "beneficios": "Restabelecer o funcionamento do equipamento, preservar o patrimônio público e manter a assistência sem interrupção."}
        if bens and bens[0]["local"]:
            campos["local"] = bens[0]["local"]
    elif predio:
        objeto = (pedido.get("objeto") or "Reparo nas instalações da unidade").strip()
        campos = {"objeto": objeto, "descricao": f"{objeto}, conforme levantamento em visita técnica.",
                  "justificativa": "As condições atuais do espaço comprometem a segurança, a ambiência e o acolhimento das pessoas atendidas.",
                  "contexto": "O problema foi identificado pela unidade e o reparo não está coberto por contrato de manutenção predial vigente.",
                  "beneficios": "Adequar o espaço às condições de segurança, ambiência e acolhimento, sem alterar a capacidade de atendimento."}
    else:
        objeto = (pedido.get("objeto") or "Contratação de serviço").strip()
        campos = {"objeto": objeto, "descricao": f"{objeto}, conforme especificações deste Termo de Referência.",
                  "justificativa": "O serviço é necessário ao funcionamento regular das unidades e [A PREENCHER: situação de contrato vigente na SES-DF].",
                  "contexto": "A demanda foi levantada pelas unidades de Atenção Primária da Região Norte.",
                  "beneficios": "Manter o funcionamento regular das unidades e a continuidade da assistência."}
    campos.update({"antecedentes": "Não há registro de contratação anterior deste serviço no histórico da ferramenta.",
                   "responsavel_aceite": "Gerência de Serviços de Atenção Primária da unidade atendida",
                   "prazo_manifestacao": prazo_gsaps(hoje)})
    return campos
