"""
Gera a planilha de demonstracao Controle_Devolucao_Pecas_Demo.xlsx com dados
100% ficticios, reproduzindo a estrutura do projeto real (abas, tabela
estruturada, formulas SUMIFS/INDEX-MATCH, validacao de dados e graficos).

Uso:  python gerar_demo.py [qtd_linhas]
"""
import random
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

random.seed(42)
N = int(sys.argv[1]) if len(sys.argv) > 1 else 2000
SAIDA = "Controle_Devolucao_Pecas_Demo.xlsx"

# ---------------------------------------------------------------- listas
STATUS_GRUPO = [
    ("Ok. Devolvida", "Concluído"),
    ("Em transporte", "Em andamento"),
    ("Em andamento", "Em andamento"),
    ("Aguardando Transporte", "Em andamento"),
    ("Aguardando devolução", "Pendente"),
    ("NF Vencida", "Pendente"),
    ("Debitar do STA", "Perda / Débito"),
    ("Não Devolve-Norte", "Não aplicável"),
    ("NPD", "Não aplicável"),
]
STATUS = [s for s, _ in STATUS_GRUPO]
PESOS_STATUS = [55, 8, 7, 5, 12, 4, 5, 2, 2]
GRUPOS = ["Concluído", "Em andamento", "Pendente", "Perda / Débito", "Não aplicável"]
TRANSPORTADORAS = ["LogExpress", "TransNorte", "ViaCargo", "RotaSul", "CargoMax", "EntregaJá"]
COORDENADORES = ["CENTRO OESTE", "NORDESTE", "NORTE", "SUDESTE MG", "SUDESTE RJ e ES", "SUDESTE SP", "SUL"]
TIPOS = ["Refrigeradores", "Chopeiras", "Misturadores"]
CONTRATOS = ["Cliente A - Refrigerador", "Cliente B - Chopeira", "Cliente C - Misturador"]
UF_REGIAO = {
    "CENTRO OESTE": ["DF", "GO", "MT", "MS"],
    "NORDESTE": ["BA", "PE", "CE", "MA", "RN", "PB", "AL", "SE", "PI"],
    "NORTE": ["AM", "PA", "RO", "AC", "AP", "RR", "TO"],
    "SUDESTE MG": ["MG"],
    "SUDESTE RJ e ES": ["RJ", "ES"],
    "SUDESTE SP": ["SP"],
    "SUL": ["PR", "SC", "RS"],
}
REGIAO_MACRO = {"CENTRO OESTE": "Centro-Oeste", "NORDESTE": "Nordeste", "NORTE": "Norte",
                "SUDESTE MG": "Sudeste", "SUDESTE RJ e ES": "Sudeste", "SUDESTE SP": "Sudeste", "SUL": "Sul"}
NOMES_UF = {"DF": "Distrito Federal", "GO": "Goiás", "MT": "Mato Grosso", "MS": "Mato Grosso do Sul",
            "BA": "Bahia", "PE": "Pernambuco", "CE": "Ceará", "MA": "Maranhão", "RN": "Rio Grande do Norte",
            "PB": "Paraíba", "AL": "Alagoas", "SE": "Sergipe", "PI": "Piauí", "AM": "Amazonas", "PA": "Pará",
            "RO": "Rondônia", "AC": "Acre", "AP": "Amapá", "RR": "Roraima", "TO": "Tocantins",
            "MG": "Minas Gerais", "RJ": "Rio de Janeiro", "ES": "Espírito Santo", "SP": "São Paulo",
            "PR": "Paraná", "SC": "Santa Catarina", "RS": "Rio Grande do Sul"}

PREFIXOS = ["Refrigeração", "Assistência Técnica", "Clima", "Frio", "Gelo", "Polar", "Ártico", "Norte Frio",
            "Refrigel", "Termotec", "Cool Service", "Ice Tech", "Frigo", "Zero Grau", "Brisa"]
SUFIXOS = ["Ltda", "ME", "Serviços", "Manutenção", "& Cia", "Refrigeração", "Técnica", "Express"]

FAMILIAS = [("Compressor", 320, "Refrigeradores"), ("Termostato", 45, "Refrigeradores"),
            ("Ventilador", 60, "Refrigeradores"), ("Placa eletrônica", 180, "Refrigeradores"),
            ("Filtro secador", 25, "Refrigeradores"), ("Gaxeta da porta", 35, "Refrigeradores"),
            ("Controlador digital", 150, "Refrigeradores"), ("Relé de partida", 30, "Refrigeradores"),
            ("Torneira de chopp", 90, "Chopeiras"), ("Serpentina", 210, "Chopeiras"),
            ("Regulador de pressão", 120, "Chopeiras"), ("Válvula extratora", 75, "Chopeiras"),
            ("Bomba de recirculação", 260, "Misturadores"), ("Motor agitador", 190, "Misturadores"),
            ("Bico dosador", 40, "Misturadores"), ("Sensor de nível", 55, "Misturadores")]

# ---------------------------------------------------------------- estilos
AZUL = "1F3A5F"
F_TIT = Font(bold=True, size=14, color=AZUL)
F_SUB = Font(italic=True, color="666666")
F_CAB = Font(bold=True, color="FFFFFF")
FILL_CAB = PatternFill("solid", fgColor=AZUL)
FILL_KPI = PatternFill("solid", fgColor="EAF1F9")
FMT_BRL = '"R$" #,##0.00'
FMT_PCT = "0.0%"
FMT_DATA = "dd/mm/yyyy"


def cabecalho(ws, linha, textos, col_ini=1):
    for i, t in enumerate(textos):
        c = ws.cell(row=linha, column=col_ini + i, value=t)
        c.font = F_CAB
        c.fill = FILL_CAB
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def larguras(ws, mapa):
    for col, w in mapa.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- dados ficticios
def gerar_stas(qtd=60):
    stas, usados = [], set()
    cod = 10000
    while len(stas) < qtd:
        nome = f"{random.choice(PREFIXOS)} {random.choice(SUFIXOS)}"
        if nome in usados:
            continue
        usados.add(nome)
        cod += random.randint(7, 90)
        coord = random.choice(COORDENADORES)
        uf = random.choice(UF_REGIAO[coord])
        slug = nome.lower().replace(" & cia", "").replace(" ", ".")
        slug = (slug.replace("ç", "c").replace("ã", "a").replace("á", "a").replace("é", "e")
                .replace("ê", "e").replace("í", "i").replace("ó", "o").replace("ú", "u"))
        email = f"{slug}@exemplo.com.br" if random.random() > 0.08 else ""
        stas.append(dict(cod=cod, nome=f"{nome} ({cod})", uf=uf, coord=coord, email=email,
                         mei="Sim" if random.random() < 0.15 else "Não",
                         cadastro=date(2024, 1, 1) + timedelta(days=random.randint(0, 600))))
    return stas


def gerar_pecas():
    pecas, seq = [], 1
    for fam, valor, tipo in FAMILIAS:
        for var in range(1, random.randint(2, 4) + 1):
            cod = f"PC-{seq:04d}"
            seq += 1
            pecas.append(dict(cod=cod, desc=f"{fam} modelo {var:02d}", valor=valor + random.choice([0, 5, 10, -5]),
                              familia=fam.upper(), tipo=tipo, generico=f"PG-{fam[:3].upper()}",
                              dev=round(valor * 0.1, 2)))
    return pecas


def gerar_base(stas, pecas, n):
    linhas = []
    ini, fim = date(2025, 9, 1), date(2026, 8, 31)
    chamado = 20000000
    nf_por_transp = {t: 100 for t in TRANSPORTADORAS}
    for _ in range(n):
        sta = random.choice(stas)
        peca = random.choice(pecas)
        chamado += random.randint(1, 40)
        enc = ini + timedelta(days=random.randint(0, (fim - ini).days))
        status = random.choices(STATUS, PESOS_STATUS)[0]
        if sta["coord"] == "NORTE" and random.random() < 0.5:
            status = "Não Devolve-Norte"
        grupo = dict(STATUS_GRUPO)[status]
        validacao = enc + timedelta(days=random.randint(5, 60)) if grupo == "Concluído" else None
        transp = random.choice(TRANSPORTADORAS) if grupo in ("Concluído", "Em andamento") else ""
        nf = ""
        if transp:
            if random.random() < 0.3:
                nf_por_transp[transp] += 1
            nf = nf_por_transp[transp]
        linhas.append(dict(chamado=chamado, mei=sta["mei"], uf=sta["uf"], coord=sta["coord"],
                           contrato=random.choice(CONTRATOS), tipo=peca["tipo"], valor=peca["valor"],
                           enc=enc, validacao=validacao, status=status, cod_sta=sta["cod"], sta=sta["nome"],
                           cod_peca=peca["cod"], desc=peca["desc"], qtd=random.choices([1, 2, 3], [80, 15, 5])[0],
                           nf=nf, transp=transp))
    linhas.sort(key=lambda r: r["enc"])
    return linhas


# ---------------------------------------------------------------- planilha
def montar():
    stas = gerar_stas()
    pecas = gerar_pecas()
    base = gerar_base(stas, pecas, N)
    n_sta = len(stas)
    ult = N + 1

    wb = Workbook()
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_base = wb.create_sheet("Base_RPTG")
    ws_par = wb.create_sheet("Parametros")
    ws_rsm = wb.create_sheet("Resumo_Status_Mes")
    ws_rtr = wb.create_sheet("Resumo_Transportadora")
    ws_rco = wb.create_sheet("Resumo_Coordenador")
    ws_ruf = wb.create_sheet("Resumo_UF")
    ws_top = wb.create_sheet("Top_Pecas")
    ws_age = wb.create_sheet("Aging_Pendentes")
    ws_val = wb.create_sheet("Base_Valores")
    ws_sta = wb.create_sheet("STAs")
    ws_ras = wb.create_sheet("Rastreio")
    ws_qd = wb.create_sheet("Qualidade_Dados")
    ws_mail = wb.create_sheet("SendBulkEmails")

    # ---------------- Parametros
    ws_par["A1"] = "Listas oficiais (usadas na validação de dados da Base_RPTG) e configurações das macros"
    ws_par["A1"].font = F_TIT
    ws_par["A2"], ws_par["B2"] = "Envio Automático Ativo:", "NAO"
    ws_par["A3"], ws_par["B3"] = "Dia da Semana (2=Seg..6=Sex, 8=Diário):", 3
    ws_par["A4"], ws_par["B4"] = "Horário de Disparo:", "09:30"
    cabecalho(ws_par, 3, ["Grupo_Status"], 3)
    cabecalho(ws_par, 3, ["Transportadora"], 5)
    cabecalho(ws_par, 3, ["Coordenador"], 7)
    cabecalho(ws_par, 3, ["Tipo"], 9)
    cabecalho(ws_par, 3, ["MEI"], 11)
    cabecalho(ws_par, 3, ["Status", "Grupo_Status"], 13)
    for i, g in enumerate(GRUPOS):
        ws_par.cell(row=4 + i, column=3, value=g)
    for i, t in enumerate(TRANSPORTADORAS):
        ws_par.cell(row=4 + i, column=5, value=t)
    for i, c in enumerate(COORDENADORES):
        ws_par.cell(row=4 + i, column=7, value=c)
    for i, t in enumerate(TIPOS):
        ws_par.cell(row=4 + i, column=9, value=t)
    ws_par["K4"], ws_par["K5"] = "Não", "Sim"
    for i, (s, g) in enumerate(STATUS_GRUPO):
        ws_par.cell(row=4 + i, column=13, value=s)
        ws_par.cell(row=4 + i, column=14, value=g)
    # coluna A5:A13 = lista de status (usada na validacao e na checagem)
    ws_par["A5"] = "Status oficiais:"
    for i, s in enumerate(sorted(STATUS)):
        ws_par.cell(row=6 + i, column=1, value=s)
    # configuracoes das macros (mesmas celulas que o codigo VBA le)
    ws_par["O3"] = "Configurações usadas pelas macros VBA"
    ws_par["O3"].font = Font(bold=True)
    cfg = {
        4: ("CNPJ do tomador do frete (rastreio Braspress) — só números", "00000000000000"),
        6: ("Caminho da imagem do banner da assinatura (opcional)", ""),
        8: ("Mês de referência do comunicado", date(2026, 6, 1)),
        9: ("Link da tabela de códigos", "https://exemplo.com.br/tabela-codigos"),
        10: ("Texto do bloco Atenção!", "Caso algum STA esteja com devolução de {MES_REF} pendente, encaminhe a nota até {DATA_NOTA}; peças de {MES_ANT} devem ser enviadas até {DATA_PECAS}."),
        11: ("CONSULTORES POR REGIÃO (região na coluna O, e-mails em Cc na coluna P)", None),
        12: ("Sul", "consultor.sul1@exemplo.com.br; consultor.sul2@exemplo.com.br"),
        13: ("Sudeste", "consultor.sudeste1@exemplo.com.br; consultor.sudeste2@exemplo.com.br"),
        14: ("Nordeste", "consultor.nordeste1@exemplo.com.br"),
        15: ("Centro oeste", "consultor.co1@exemplo.com.br"),
        16: ("Pasta dos PDFs anexados", r"C:\Comunicados_STAs"),
        18: ("Pasta onde procurar o export do sistema de chamados (importação)", r"C:\Exports_Chamados"),
        19: ("Último arquivo importado (preenchido pela macro)", None),
        20: ("E-mail do Financeiro (planilha de débito)", "financeiro@exemplo.com.br"),
        21: ("Pasta onde salvar a planilha de débito (vazio = pasta desta planilha)", ""),
        23: ("Cobrança semanal - último envio automático", None),
        24: ("Cobrança automática - dia da semana (1=dom ... 7=sáb)", 3),
        25: ("Cobrança automática - horário", "09:00"),
        26: ("Cobrança automática - intervalo em dias", 7),
    }
    for lin, (rot, val) in cfg.items():
        ws_par.cell(row=lin, column=15, value=rot)
        if val is not None:
            ws_par.cell(row=lin, column=16, value=val)
    ws_par["P8"].number_format = FMT_DATA
    larguras(ws_par, {"A": 38, "C": 16, "E": 16, "G": 18, "I": 16, "K": 8, "M": 22, "N": 16, "O": 58, "P": 60})
    wb.defined_names["mapGrupoStatus"] = DefinedName("mapGrupoStatus", attr_text="Parametros!$M$4:$N$12")

    # ---------------- Base_Valores
    cabecalho(ws_val, 1, ["Cód. Peça", "Peça", "Valor Unitário (Débito)", "Família", "Descrição da família",
                          "Cód Genérico", "Valor unitário de devolução", "Tipo"])
    for i, p in enumerate(pecas, start=2):
        ws_val.append([p["cod"], p["desc"], p["valor"], p["familia"], f"{p['familia']} - RETORNO EM GARANTIA",
                       p["generico"], p["dev"], p["tipo"]])
        ws_val.cell(row=i, column=3).number_format = FMT_BRL
        ws_val.cell(row=i, column=7).number_format = FMT_BRL
    tv = Table(displayName="tblValores", ref=f"A1:H{len(pecas) + 1}")
    tv.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_val.add_table(tv)
    ws_val.freeze_panes = "A2"
    larguras(ws_val, {"A": 12, "B": 30, "C": 20, "D": 24, "E": 38, "F": 14, "G": 24, "H": 16})

    # ---------------- STAs
    cabecalho(ws_sta, 1, ["UF", "ESTADO COMPLETO", "STA", "Região", "Coordenador", "Cód STA", "Email",
                          "Data_Cadastro", "STA ativa?", "MEI?", "E-mail válido?"])
    for i, s in enumerate(stas, start=2):
        ws_sta.append([s["uf"], NOMES_UF[s["uf"]], s["nome"], REGIAO_MACRO[s["coord"]], s["coord"], s["cod"],
                       s["email"], s["cadastro"], "Sim", s["mei"],
                       f'=IF(AND(G{i}<>"",ISNUMBER(FIND("@",G{i})),ISNUMBER(FIND(".",G{i}))),"Sim","Não")'])
        ws_sta.cell(row=i, column=8).number_format = FMT_DATA
    ts = Table(displayName="tblSTAs", ref=f"A1:K{n_sta + 1}")
    ts.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_sta.add_table(ts)
    ws_sta.freeze_panes = "A2"
    larguras(ws_sta, {"A": 6, "B": 22, "C": 40, "D": 14, "E": 18, "F": 10, "G": 40, "H": 14, "I": 11, "J": 8, "K": 14})

    # ---------------- Base_RPTG
    cols = ["Email_STA", "Chamado", "MEI", "UF", "Coordenador", "Contrato", "Tipo", "Valor_Debito_Unit",
            "Valor_Total", "Dta_Encerramento", "Data_Validacao", "Status", "Grupo_Status", "Faixa_Aging",
            "Dias_Em_Aberto", "Codigo_STA", "STA", "Codigo_Peca", "Descricao", "QTD", "NF", "Mes", "Ano",
            "Transportadora", "Status_Rastreio"]
    cabecalho(ws_base, 1, cols)
    sta_rng_cod = f"STAs!$F$2:$F${n_sta + 1}"
    sta_rng_email = f"STAs!$G$2:$G${n_sta + 1}"
    for i, r in enumerate(base, start=2):
        ws_base.append([
            f'=IFERROR(INDEX({sta_rng_email},MATCH(P{i},{sta_rng_cod},0))&"","")',
            r["chamado"], r["mei"], r["uf"], r["coord"], r["contrato"], r["tipo"], r["valor"],
            f"=IFERROR(T{i}*H{i},0)",
            r["enc"], r["validacao"], r["status"],
            f'=IF(L{i}="","",IFERROR(VLOOKUP(L{i},mapGrupoStatus,2,0),"Outros"))',
            f'=IF(O{i}="","",IF(O{i}<=30,"0-30 dias",IF(O{i}<=60,"31-60 dias",IF(O{i}<=90,"61-90 dias","Mais de 90 dias"))))',
            f'=IF(OR(M{i}="Pendente",M{i}="Em andamento"),IF(J{i}="","",TODAY()-J{i}),"")',
            r["cod_sta"], r["sta"], r["cod_peca"], r["desc"], r["qtd"], r["nf"],
            f"=IF(J{i}=\"\",\"\",MONTH(J{i}))", f"=IF(J{i}=\"\",\"\",YEAR(J{i}))",
            r["transp"],
            f'=IFERROR(INDEX(Rastreio!$H$2:$H$200,MATCH(X{i}&"|"&U{i},Rastreio!$A$2:$A$200,0)),"")',
        ])
        ws_base.cell(row=i, column=8).number_format = FMT_BRL
        ws_base.cell(row=i, column=9).number_format = FMT_BRL
        ws_base.cell(row=i, column=10).number_format = FMT_DATA
        ws_base.cell(row=i, column=11).number_format = FMT_DATA
    tb = Table(displayName="tblBase", ref=f"A1:Y{ult}")
    tb.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_base.add_table(tb)
    ws_base.freeze_panes = "C2"
    for col, lista in {"L": "Parametros!$M$4:$M$12", "X": "Parametros!$E$4:$E$9",
                       "E": "Parametros!$G$4:$G$10", "G": "Parametros!$I$4:$I$6", "C": "Parametros!$K$4:$K$5"}.items():
        dv = DataValidation(type="list", formula1=f"={lista}", allow_blank=True)
        dv.error, dv.errorTitle = "Escolha um valor da lista oficial (aba Parametros).", "Valor fora da lista"
        ws_base.add_data_validation(dv)
        dv.add(f"{col}2:{col}{ult}")
    larguras(ws_base, {"A": 36, "B": 11, "C": 6, "D": 5, "E": 17, "F": 24, "G": 15, "H": 12, "I": 12, "J": 13,
                       "K": 13, "L": 22, "M": 16, "N": 15, "O": 13, "P": 11, "Q": 40, "R": 11, "S": 30, "T": 6,
                       "U": 7, "V": 6, "W": 7, "X": 16, "Y": 16})

    # ---------------- Rastreio
    cabecalho(ws_ras, 1, ["Chave", "Transportadora", "NF", "Codigo_STA", "STA", "Itens em transporte", "Valor (R$)",
                          "Status_Rastreio", "Ultima_Ocorrencia", "Data_Ocorrencia", "Previsao_Entrega",
                          "Data_Entrega", "Conhecimento_CTe", "Consultado_em", "Link_Consulta"])
    vistos, lin = set(), 2
    for r in base:
        if r["transp"] and r["nf"] and (r["transp"], r["nf"]) not in vistos and dict(STATUS_GRUPO)[r["status"]] == "Em andamento":
            vistos.add((r["transp"], r["nf"]))
            ws_ras.append([f"=B{lin}&\"|\"&C{lin}", r["transp"], r["nf"], r["cod_sta"], r["sta"],
                           f'=SUMIFS(tblBase[QTD],tblBase[Transportadora],B{lin},tblBase[NF],C{lin},tblBase[Grupo_Status],"Em andamento")',
                           f'=SUMIFS(tblBase[Valor_Total],tblBase[Transportadora],B{lin},tblBase[NF],C{lin},tblBase[Grupo_Status],"Em andamento")',
                           random.choice(["Em trânsito", "Saiu para entrega", "Aguardando coleta", ""]),
                           None, None, None, None, None, None,
                           f'=IF(B{lin}="Braspress",HYPERLINK("https://www.braspress.com.br/site/w/tracking/search?cnpj="&Parametros!$P$4&"&documentType=NOTAFISCAL&numero="&C{lin},"Abrir rastreio"),"")'])
            ws_ras.cell(row=lin, column=7).number_format = FMT_BRL
            lin += 1
    ws_ras.freeze_panes = "A2"
    larguras(ws_ras, {"A": 16, "B": 16, "C": 7, "D": 11, "E": 40, "F": 18, "G": 14, "H": 18, "I": 18, "J": 14,
                      "K": 14, "L": 14, "M": 16, "N": 16, "O": 16})

    # ---------------- Resumo_Status_Mes
    ws_rsm["A1"] = "Itens por status e mês (fórmulas — atualiza sozinho ao editar a Base_RPTG)"
    ws_rsm["A1"].font = F_TIT
    meses = [(2025, m) for m in range(9, 13)] + [(2026, m) for m in range(1, 13)]
    cabecalho(ws_rsm, 3, ["Status"] + [f"{a}-{m:02d}" for a, m in meses] + ["Total"])
    for i, s in enumerate(sorted(STATUS), start=4):
        ws_rsm.cell(row=i, column=1, value=s)
        for j, (a, m) in enumerate(meses, start=2):
            ws_rsm.cell(row=i, column=j, value=f"=SUMIFS(tblBase[QTD],tblBase[Status],$A{i},tblBase[Ano],{a},tblBase[Mes],{m})")
        ws_rsm.cell(row=i, column=len(meses) + 2, value=f"=SUM(B{i}:{get_column_letter(len(meses) + 1)}{i})")
    ws_rsm.column_dimensions["A"].width = 24
    ch = LineChart()
    ch.title, ch.height, ch.width = "Itens por status ao longo dos meses", 9, 24
    ch.add_data(Reference(ws_rsm, min_col=1, max_col=len(meses) + 1, min_row=4, max_row=3 + len(STATUS)),
                from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(ws_rsm, min_col=2, max_col=len(meses) + 1, min_row=3))
    ws_rsm.add_chart(ch, "A15")

    # ---------------- Resumos por categoria
    def resumo(ws, titulo, campo, categorias, extra=22):
        ws["A1"], ws["A2"] = titulo, "Digite uma nova categoria nas linhas em branco no fim da lista: as fórmulas já estão prontas."
        ws["A1"].font, ws["A2"].font = F_TIT, F_SUB
        cabecalho(ws, 3, ["Categoria", "Itens", "Valor total (R$)", "Devolvidos", "% devolvido", "Em aberto"])
        fim = 3 + extra
        for i in range(4, fim + 1):
            if i - 4 < len(categorias):
                ws.cell(row=i, column=1, value=categorias[i - 4])
            ws.cell(row=i, column=2, value=f'=IF($A{i}="","",SUMIFS(tblBase[QTD],tblBase[{campo}],$A{i}))')
            ws.cell(row=i, column=3, value=f'=IF($A{i}="","",SUMIFS(tblBase[Valor_Total],tblBase[{campo}],$A{i}))')
            ws.cell(row=i, column=4, value=f'=IF($A{i}="","",SUMIFS(tblBase[QTD],tblBase[{campo}],$A{i},tblBase[Grupo_Status],"Concluído"))')
            ws.cell(row=i, column=5, value=f'=IFERROR(D{i}/B{i},"")')
            ws.cell(row=i, column=6, value=f'=IF($A{i}="","",SUMIFS(tblBase[QTD],tblBase[{campo}],$A{i},tblBase[Grupo_Status],"Pendente")+SUMIFS(tblBase[QTD],tblBase[{campo}],$A{i},tblBase[Grupo_Status],"Em andamento"))')
            ws.cell(row=i, column=3).number_format = FMT_BRL
            ws.cell(row=i, column=5).number_format = FMT_PCT
        t = fim + 1
        ws.cell(row=t, column=1, value="Total da base").font = Font(bold=True)
        ws.cell(row=t, column=2, value="=SUM(tblBase[QTD])")
        ws.cell(row=t, column=3, value="=SUM(tblBase[Valor_Total])").number_format = FMT_BRL
        ws.cell(row=t, column=4, value='=SUMIFS(tblBase[QTD],tblBase[Grupo_Status],"Concluído")')
        ws.cell(row=t, column=5, value=f'=IFERROR(D{t}/B{t},"")').number_format = FMT_PCT
        ws.cell(row=t, column=6, value='=SUMIFS(tblBase[QTD],tblBase[Grupo_Status],"Pendente")+SUMIFS(tblBase[QTD],tblBase[Grupo_Status],"Em andamento")')
        ws.cell(row=t + 1, column=1, value="Não listado acima (categoria nova)")
        for c in "BCDF":
            ws[f"{c}{t + 1}"] = f"={c}{t}-SUM({c}4:{c}{fim})"
        larguras(ws, {"A": 26, "B": 10, "C": 18, "D": 12, "E": 12, "F": 11})
        g = BarChart()
        g.type, g.title, g.height, g.width = "bar", titulo, 8, 16
        g.add_data(Reference(ws, min_col=2, min_row=3, max_row=3 + len(categorias)), titles_from_data=True)
        g.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + len(categorias)))
        g.legend = None
        ws.add_chart(g, "H3")

    resumo(ws_rtr, "Desempenho por transportadora", "Transportadora", TRANSPORTADORAS)
    resumo(ws_rco, "Desempenho por regional / coordenador", "Coordenador", COORDENADORES)
    ufs = sorted({u for lst in UF_REGIAO.values() for u in lst})
    resumo(ws_ruf, "Volume por UF", "UF", ufs, extra=len(ufs) + 5)

    # ---------------- Top_Pecas
    ws_top["A1"] = "Peças com maior volume (edite os códigos da coluna A; o restante é fórmula)"
    ws_top["A1"].font = F_TIT
    cabecalho(ws_top, 3, ["Código", "Descrição", "Itens", "Valor total (R$)", "Em aberto"])
    for i, p in enumerate(pecas[:25], start=4):
        ws_top.cell(row=i, column=1, value=p["cod"])
        ws_top.cell(row=i, column=2, value=f'=IFERROR(INDEX(tblValores[Peça],MATCH($A{i},tblValores[Cód. Peça],0)),"")')
        ws_top.cell(row=i, column=3, value=f'=IF($A{i}="","",SUMIFS(tblBase[QTD],tblBase[Codigo_Peca],$A{i}))')
        ws_top.cell(row=i, column=4, value=f'=IF($A{i}="","",SUMIFS(tblBase[Valor_Total],tblBase[Codigo_Peca],$A{i}))').number_format = FMT_BRL
        ws_top.cell(row=i, column=5, value=f'=IF($A{i}="","",SUMIFS(tblBase[QTD],tblBase[Codigo_Peca],$A{i},tblBase[Grupo_Status],"Pendente")+SUMIFS(tblBase[QTD],tblBase[Codigo_Peca],$A{i},tblBase[Grupo_Status],"Em andamento"))')
    ws_top.freeze_panes = "A4"
    larguras(ws_top, {"A": 12, "B": 32, "C": 10, "D": 18, "E": 11})

    # ---------------- Aging
    ws_age["A1"] = "Aging dos itens em aberto — pendentes e em andamento (dias desde o encerramento do chamado)"
    ws_age["A1"].font = F_TIT
    cabecalho(ws_age, 3, ["Faixa", "Itens", "Valor em risco (R$)"])
    for i, f in enumerate(["0-30 dias", "31-60 dias", "61-90 dias", "Mais de 90 dias"], start=4):
        ws_age.cell(row=i, column=1, value=f)
        ws_age.cell(row=i, column=2, value=f"=SUMIFS(tblBase[QTD],tblBase[Faixa_Aging],$A{i})")
        ws_age.cell(row=i, column=3, value=f"=SUMIFS(tblBase[Valor_Total],tblBase[Faixa_Aging],$A{i})").number_format = FMT_BRL
    larguras(ws_age, {"A": 18, "B": 10, "C": 20})
    pie = PieChart()
    pie.title, pie.height, pie.width = "Itens em aberto por faixa de aging", 8, 12
    pie.add_data(Reference(ws_age, min_col=2, min_row=3, max_row=7), titles_from_data=True)
    pie.set_categories(Reference(ws_age, min_col=1, min_row=4, max_row=7))
    ws_age.add_chart(pie, "E3")

    # ---------------- Qualidade_Dados (numeros do diagnostico do projeto real, sem dados)
    ws_qd["A1"] = "Diagnóstico de qualidade dos dados (correções aplicadas na base original do projeto)"
    ws_qd["A1"].font = F_TIT
    cabecalho(ws_qd, 3, ["Categoria", "Coluna", "Ocorrências", "Detalhe", "Ação aplicada"])
    diag = [
        ("Padronização", "Status", 5025, "Mesmo status escrito de formas diferentes (ex.: 'Ok. Devolvida' / 'ok. Devolvida' / 'OK. Devolvida')", "Textos unificados em 9 status oficiais"),
        ("Padronização", "Transportadora", 410, "Nome com caixa diferente ('Begur'/'begur')", "Nomes de transportadora unificados"),
        ("Padronização", "Coordenador", 949, "'NORTE'/'Norte', 'SUDESTE  MG' com espaço duplo", "Regionais unificadas"),
        ("Padronização", "Tipo", 2, "'Coolers'/'coolers'", "Tipos unificados"),
        ("Padronização", "MEI", 0, "'Não' com espaço no fim, 'Sim'/'sim'", "Padronizado como Sim/Não"),
        ("Erro de fórmula", "Diversas", 7, "Células com erro (#N/D, #VALOR!) vindas de PROCV", "Convertidas em célula vazia"),
        ("Texto no lugar de vazio", "NF / Data_Validacao / MEI", 6197, "Textos 'NULL' e 'Não Encontrado' em campos numéricos e de data", "Convertidos em célula vazia"),
        ("Tipo de dado", "Data_Validacao", 21187, "Datas gravadas como texto misturadas com datas reais", "Todas convertidas em data real"),
        ("Tipo de dado", "Dta_Encerramento / Data_Validacao", 95705, "Datas armazenadas como número de série", "Convertidas em data real dd/mm/aaaa"),
        ("Nomenclatura", "Chmado / mei", 2, "Nomes de coluna com erro de digitação e caixa inconsistente", "Renomeadas para 'Chamado' e 'MEI'"),
        ("Integridade", "Email_STA", 21345, "Registros cujo código de STA não tinha e-mail cadastrado", "Coluna Email_STA por INDEX/MATCH; vazios visíveis para cadastro"),
        ("Integridade", "STAs!Email", 85, "STAs sem e-mail ou com formato inválido", "Coluna 'E-mail válido?' criada para conferência"),
        ("Duplicidade", "Chamado + Peça + NF + Status + Período", 0, "Linhas com a mesma combinação", "Sinalizadas (não removidas)"),
    ]
    for i, linha in enumerate(diag, start=4):
        for j, v in enumerate(linha, start=1):
            ws_qd.cell(row=i, column=j, value=v)
    l0 = 4 + len(diag) + 2
    ws_qd.cell(row=l0, column=1, value="Checagem ao vivo da Base_RPTG (recalcula sozinha)").font = Font(bold=True)
    checks = [
        ("Linhas na base", "=ROWS(tblBase[Chamado])"),
        ("Status em branco", "=COUNTBLANK(tblBase[Status])"),
        ("Status fora da lista oficial", '=SUMPRODUCT((tblBase[Status]<>"")*(COUNTIF(Parametros!$M$4:$M$12,tblBase[Status])=0))'),
        ("Linhas sem código de STA", "=COUNTBLANK(tblBase[Codigo_STA])"),
        ("Linhas sem e-mail de STA", '=COUNTIF(tblBase[Email_STA],"")'),
        ("Linhas sem valor de débito", "=COUNTBLANK(tblBase[Valor_Debito_Unit])"),
        ("Linhas sem data de encerramento", "=COUNTBLANK(tblBase[Dta_Encerramento])"),
        ("Itens em aberto com mais de 90 dias", '=COUNTIF(tblBase[Faixa_Aging],"Mais de 90 dias")'),
        ("STAs cadastradas sem e-mail válido", '=COUNTIF(tblSTAs[E-mail válido?],"Não")'),
    ]
    for i, (rot, f) in enumerate(checks, start=l0 + 1):
        ws_qd.cell(row=i, column=1, value=rot)
        ws_qd.cell(row=i, column=2, value=f)
    larguras(ws_qd, {"A": 34, "B": 30, "C": 12, "D": 70, "E": 52})

    # ---------------- SendBulkEmails (1 linha por STA, tudo por formula)
    cabecalho(ws_mail, 1, ["EmailAddress", "Subject", "Body", "Cód STA", "Nome do STA", "Itens em aberto",
                           "Valor em aberto", "Dias em aberto (máx.)", "Último envio", "", "Meses em aberto",
                           "Status", "Valor Total (R$)", "Região", "Consultores em Cópia (CC)"])
    for i in range(2, n_sta + 2):
        ws_mail.cell(row=i, column=4, value=f"=STAs!F{i}")
        ws_mail.cell(row=i, column=5, value=f"=STAs!C{i}")
        ws_mail.cell(row=i, column=6, value=f'=SUMIFS(tblBase[QTD],tblBase[Codigo_STA],D{i},tblBase[Grupo_Status],"Pendente")+SUMIFS(tblBase[QTD],tblBase[Codigo_STA],D{i},tblBase[Grupo_Status],"Em andamento")')
        ws_mail.cell(row=i, column=7, value=f'=SUMIFS(tblBase[Valor_Total],tblBase[Codigo_STA],D{i},tblBase[Grupo_Status],"Pendente")+SUMIFS(tblBase[Valor_Total],tblBase[Codigo_STA],D{i},tblBase[Grupo_Status],"Em andamento")').number_format = FMT_BRL
        ws_mail.cell(row=i, column=8, value=f'=IFERROR(_xlfn.MAXIFS(tblBase[Dias_Em_Aberto],tblBase[Codigo_STA],D{i}),0)')
        ws_mail.cell(row=i, column=1, value=f'=IF(F{i}>0,STAs!G{i},"")')
        ws_mail.cell(row=i, column=14, value=f"=STAs!D{i}")
        ws_mail.cell(row=i, column=15, value=f'=IFERROR(INDEX(Parametros!$P$12:$P$15,MATCH(N{i},Parametros!$O$12:$O$15,0)),"")')
        ws_mail.cell(row=i, column=2, value=f'=IF(F{i}>0,"Devolução RPTG - "&N{i}&" - "&F{i}&" peça(s) pendente(s) - STA "&D{i},"")')
        ws_mail.cell(row=i, column=3, value=(
            f'=IF(F{i}>0,"Prezado(a) parceiro(a) "&E{i}&","&CHAR(10)&CHAR(10)&'
            f'"Consta em nosso controle de devolução RPTG "&F{i}&" peça(s) pendente(s) de devolução, no valor de "&"R$ "&FIXED(G{i},2)&'
            f'", a mais antiga há "&H{i}&" dias."&CHAR(10)&CHAR(10)&"Solicitamos o envio das peças ou a justificativa em até 5 dias úteis."&CHAR(10)&CHAR(10)&'
            f'"Atenciosamente,"&CHAR(10)&"Qualidade - RPTG","")'))
    ws_mail.freeze_panes = "A2"
    larguras(ws_mail, {"A": 36, "B": 50, "C": 60, "D": 10, "E": 36, "F": 10, "G": 14, "H": 12, "I": 14, "K": 14,
                       "L": 12, "M": 14, "N": 14, "O": 50})

    # ---------------- Dashboard
    ws_dash["B2"] = "Controle de Devolução de Peças (RPTG) — Painel de Controle"
    ws_dash["B2"].font = Font(bold=True, size=18, color=AZUL)
    ws_dash["B3"] = '="Base de demonstração com "&ROWS(tblBase[Chamado])&" registros fictícios · atualizado em "&DAY(TODAY())&"/"&MONTH(TODAY())&"/"&YEAR(TODAY())'
    ws_dash["B3"].font = F_SUB
    kpis = [
        ("Itens totais", "=SUM(tblBase[QTD])", "#,##0"),
        ("Itens devolvidos", '=SUMIFS(tblBase[QTD],tblBase[Grupo_Status],"Concluído")', "#,##0"),
        ("% devolvido", "=IFERROR(C6/C5,0)", FMT_PCT),
        ("Itens em aberto", '=SUMIFS(tblBase[QTD],tblBase[Grupo_Status],"Pendente")+SUMIFS(tblBase[QTD],tblBase[Grupo_Status],"Em andamento")', "#,##0"),
        ("Valor total debitável", "=SUM(tblBase[Valor_Total])", FMT_BRL),
        ("Valor em risco (em aberto)", '=SUMIFS(tblBase[Valor_Total],tblBase[Grupo_Status],"Pendente")+SUMIFS(tblBase[Valor_Total],tblBase[Grupo_Status],"Em andamento")', FMT_BRL),
        ("Valor já debitado do STA", '=SUMIFS(tblBase[Valor_Total],tblBase[Status],"Debitar do STA")', FMT_BRL),
        ("Aging médio dos pendentes (dias)", "=IFERROR(AVERAGE(tblBase[Dias_Em_Aberto]),0)", "0"),
        ("STAs com itens em aberto", f'=COUNTIF(SendBulkEmails!F2:F{n_sta + 1},">0")', "#,##0"),
        ("Itens sem e-mail de STA", '=COUNTIF(tblBase[Email_STA],"")', "#,##0"),
    ]
    for i, (rot, f, fmt) in enumerate(kpis, start=5):
        a, b = ws_dash.cell(row=i, column=2, value=rot), ws_dash.cell(row=i, column=3, value=f)
        a.fill, b.fill = FILL_KPI, FILL_KPI
        a.font, b.font = Font(bold=True, color=AZUL), Font(bold=True, size=12)
        b.number_format = fmt
        b.alignment = Alignment(horizontal="right")
    cabecalho(ws_dash, 4, ["Grupo", "Itens"], 6)
    for i, g in enumerate(GRUPOS, start=5):
        ws_dash.cell(row=i, column=6, value=g)
        ws_dash.cell(row=i, column=7, value=f"=SUMIFS(tblBase[QTD],tblBase[Grupo_Status],$F{i})")
    bc = BarChart()
    bc.title, bc.height, bc.width, bc.legend = "Itens por grupo de status", 7.5, 14, None
    bc.add_data(Reference(ws_dash, min_col=7, min_row=4, max_row=9), titles_from_data=True)
    bc.set_categories(Reference(ws_dash, min_col=6, min_row=5, max_row=9))
    ws_dash.add_chart(bc, "I2")
    ws_dash["B17"] = "Como usar"
    ws_dash["B17"].font = Font(bold=True, size=12, color=AZUL)
    uso = [
        "1. Lance ou cole novos dados apenas na aba Base_RPTG (tabela tblBase). Todo o resto se atualiza sozinho.",
        "1a. Preencha só as colunas de entrada: Email_STA, Valor_Total, Grupo_Status, Faixa_Aging, Dias_Em_Aberto, Mes, Ano e Status_Rastreio são fórmulas.",
        "1b. Cadastre a STA na aba STAs (Cód STA + Email) — é de lá que vêm o e-mail e a lista da aba SendBulkEmails.",
        "2. Status, Transportadora, Coordenador, Tipo e MEI têm lista suspensa (aba Parametros) — evita a mesma informação escrita de formas diferentes.",
        "3. Grupo_Status agrupa os 9 status em Concluído / Em andamento / Pendente / Perda-Débito / Não aplicável (intervalo nomeado mapGrupoStatus).",
        "4. Dias_Em_Aberto e Faixa_Aging medem o tempo desde o encerramento do chamado para o que ainda não voltou.",
        "5. Abas Resumo_* e Aging_Pendentes usam SUMIFS: não precisam de 'Atualizar' como as tabelas dinâmicas.",
        "6. Qualidade_Dados lista o que foi corrigido na base original e faz checagens ao vivo.",
        "7. SendBulkEmails tem 1 linha por STA cadastrada, com assunto e corpo por fórmula; as macros VBA (pasta src/) fazem o envio pelo Outlook.",
        "8. Todos os nomes, e-mails, códigos e valores desta planilha são fictícios.",
    ]
    for i, t in enumerate(uso, start=18):
        ws_dash.cell(row=i, column=2, value=t)
    larguras(ws_dash, {"A": 2, "B": 34, "C": 18, "D": 2, "E": 2, "F": 16, "G": 10})
    ws_dash.sheet_view.showGridLines = False
    ws_dash.page_setup.orientation = "landscape"
    ws_dash.page_setup.fitToWidth = 1
    ws_dash.page_setup.fitToHeight = 1
    ws_dash.sheet_properties.pageSetUpPr.fitToPage = True
    ws_dash.print_area = "A1:R28"

    wb.save(SAIDA)
    print(f"gerado {SAIDA}: {N} linhas, {n_sta} STAs, {len(pecas)} peças")


if __name__ == "__main__":
    montar()
