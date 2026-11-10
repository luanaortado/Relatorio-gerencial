from __future__ import annotations

import json
import re
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
HTML = ROOT / "index.html"
DRE = Path("/Users/luanaortado/Downloads/DRE Financeira Projetada 2026 - OUT.xlsx")

MONTHS = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "aug", "sep", "oct"]
COLS = dict(zip(MONTHS, [1, 2, 3, 4, 6, 8, 10, 12, 14, 16]))

GROUPS = {
    "administrativo": (12, range(13, 33)),
    "comercial": (33, range(34, 48)),
    "operacional": (48, range(49, 66)),
    "impostos": (66, range(67, 74)),
    "financeiro": (74, range(75, 80)),
}

LABELS = {
    "ATS": "ATs (Assistências Técnicas)",
    "TRANSPORTADORAS": "Transportadoras + Correios",
    "DESP COM VEICULO PRÓPRIO": "Despesas com Veículo Próprio",
    "AQUISIÇÃO DE VEICULOS": "Aquisição de Veículos",
    "MATERIAIS UTILIZADOS EM ASSIST": "Materiais Utilizados em Assistência",
    "ROUPAS PROF E MAT DE PROTEÇÃO": "Roupas Profissionais e Materiais de Proteção",
    "MANUTENÇÃO MÁQUINAS E EQUIPAME": "Manutenção de Máquinas e Equipamentos",
    "COMISSÕES REPRESENTANTES E TER": "Comissões de Representantes e Terceiros",
    "AQUISIÇÃO MÁQUINAS E EQUIPAMEN": "Aquisição de Máquinas e Equipamentos",
    "ASSISTêNCIA MÉDICA/ODONTOLÓGIC": "Assistência Médica/Odontológica",
    "BENEFICIOS - VT / VR / VA": "Benefícios - VT / VR / VA",
    "CONSULTORIA E ASSESSORIA": "Consultoria e Assessoria",
    "PAGAMENTO DE EMPRÉSTIMOS (PRIN": "Pagamento de Empréstimos (Principal)",
    "TARIFAS BANCARIAS": "Tarifas Bancárias",
    "IMPOSTOS NA IMPORTACAO": "Impostos na Importação",
}


def money(value):
    return round(float(value or 0), 2)


def clean_label(value):
    raw = str(value or "").strip()
    return LABELS.get(raw, raw.title())


def js(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


wb = load_workbook(DRE, data_only=True, read_only=True)
ws = wb["FC 2026"]

official = {}
details = {}
for month, col in COLS.items():
    entrada = money(ws.cell(3, col + 1).value)
    despesa = money(ws.cell(7, col + 1).value)
    resultado = money(ws.cell(9, col + 1).value)
    caixa = money(ws.cell(10, col + 1).value)
    fin_total = money(ws.cell(74, col + 1).value)
    expense_groups = {
        "administrativo": money(ws.cell(12, col + 1).value),
        "comercial": money(ws.cell(33, col + 1).value),
        "operacional": money(ws.cell(48, col + 1).value),
        "impostos": money(ws.cell(66, col + 1).value),
    }
    expense_sum = sum(expense_groups.values())
    official[month] = {
        "receita": money(ws.cell(4, col + 1).value),
        "outras": money(ws.cell(5, col + 1).value),
        "emprestimos": money(ws.cell(6, col + 1).value),
        "entradaTotal": entrada,
        "despesaOp": despesa,
        "resultadoOp": resultado,
        "margemOp": round(resultado / entrada * 100, 1) if entrada else 0,
        "blocoFin": fin_total,
        "geracaoCaixa": caixa,
        "margemCaixa": round(caixa / entrada * 100, 1) if entrada else 0,
        "comprometimento": round(100 - caixa / entrada * 100, 1) if entrada else 0,
        "bloco": {
            "pagamentos": money(ws.cell(75, col + 1).value),
            "juros": money(ws.cell(76, col + 1).value),
            "tarifas": money(ws.cell(77, col + 1).value),
            "adiantamento": money(ws.cell(78, col + 1).value),
            "estorno": money(ws.cell(79, col + 1).value),
        },
        "despesas": {
            k: {"total": v, "pct": round(v / expense_sum * 100, 1) if expense_sum else 0}
            for k, v in expense_groups.items()
        },
        "nota": "Fonte oficial: DRE Financeira Projetada 2026 - OUT.xlsx, conciliada em novembro de 2026.",
    }
    if month == "jun":
        official[month]["nota"] += " Em junho, R$ 18.249,94 foram reclassificados de impostos para pagamento de empréstimos, sem alterar a geração de caixa."

    month_detail = {}
    for key, (total_row, item_rows) in GROUPS.items():
        total = money(ws.cell(total_row, col + 1).value)
        items = []
        for row in item_rows:
            val = money(ws.cell(row, col + 1).value)
            if abs(val) >= 0.005:
                items.append({"nome": clean_label(ws.cell(row, 1).value), "val": val})
        item_sum = round(sum(item["val"] for item in items), 2)
        difference = round(total - item_sum, 2)
        if abs(difference) >= 0.01:
            items.append({"nome": "Ajustes e outras rubricas", "val": difference})
        items.sort(key=lambda item: item["val"], reverse=True)
        for item in items:
            item["pct"] = round(item["val"] / total * 100) if total else 0
        month_detail[key] = {
            "total": total,
            "pct": round(total / (despesa + fin_total) * 100, 1) if despesa + fin_total else 0,
            "itens": items,
        }
    details[month] = month_detail

oct_data = {
    "label": "Outubro",
    "com": {
        "novosQtd": 5, "novosVal": 8150, "distratosQtd": 2, "distratosVal": 3855.42,
        "reajustesContratos": 4, "reajustesAprovados": 4, "reajustesVal": 298.06,
        "reajustesPct": 3.85, "frotaAtiva": 286, "contratosAtivos": 129,
        "ticketMedio": 1893.4822727, "faturamentoAtivo": 541535.93, "impactoFat": 4294.58,
        "distribuicaoUF": [
            {"uf": "AC", "qtd": 1, "pct": .3}, {"uf": "AM", "qtd": 4, "pct": 1.4},
            {"uf": "BA", "qtd": 23, "pct": 8.0}, {"uf": "CE", "qtd": 12, "pct": 4.2},
            {"uf": "DF", "qtd": 8, "pct": 2.8}, {"uf": "GO", "qtd": 2, "pct": .7},
            {"uf": "MA", "qtd": 6, "pct": 2.1}, {"uf": "MG", "qtd": 17, "pct": 5.9},
            {"uf": "MS", "qtd": 2, "pct": .7}, {"uf": "MT", "qtd": 3, "pct": 1.0},
            {"uf": "PA", "qtd": 23, "pct": 8.0}, {"uf": "PE", "qtd": 18, "pct": 6.3},
            {"uf": "PR", "qtd": 8, "pct": 2.8}, {"uf": "RJ", "qtd": 41, "pct": 14.3},
            {"uf": "RS", "qtd": 4, "pct": 1.4}, {"uf": "SE", "qtd": 7, "pct": 2.4},
            {"uf": "SP", "qtd": 106, "pct": 37.1}, {"uf": "TO", "qtd": 1, "pct": .3},
        ],
    },
    "ops": {
        "notaSLA": "O SLA fechou outubro em 81,4%, melhora de 3,2 p.p. sobre setembro, ainda 3,6 p.p. abaixo da referência de 85%. O principal desvio foi o Amazonas, com TMA de 19,5 dias.",
        "chamados": 129, "dentroDeSLA": 105, "encerrados": 102, "trocas": 17, "desligamentos": 5,
        "encerradosMes": 102, "encerradosMesPct": 79.07, "encerradosSLAqtd": 105,
        "encerradosSLApct": 81.4, "diciclos": 17, "diciclosPct": 13.18,
        "desligamentosEsp": 5, "desligamentosEspPct": 3.88,
        "chamadosPorUF": [
            {"uf":"AC","qtd":2,"sla":1},{"uf":"AM","qtd":2,"sla":19.5},{"uf":"BA","qtd":17,"sla":5.3},
            {"uf":"CE","qtd":4,"sla":4.25},{"uf":"DF","qtd":3,"sla":6.3},{"uf":"GO","qtd":1,"sla":2},
            {"uf":"MA","qtd":0,"sla":0},{"uf":"MG","qtd":10,"sla":4.9},{"uf":"MS","qtd":0,"sla":0},
            {"uf":"MT","qtd":5,"sla":6},{"uf":"PA","qtd":3,"sla":5},{"uf":"PE","qtd":5,"sla":2},
            {"uf":"PR","qtd":1,"sla":7},{"uf":"RJ","qtd":13,"sla":5.7},{"uf":"RS","qtd":3,"sla":6},
            {"uf":"SE","qtd":1,"sla":0},{"uf":"SP","qtd":57,"sla":3.35},{"uf":"TO","qtd":2,"sla":5},
        ],
        "motivos": [
            {"m":"Guidão/Postura","q":27,"pct":20.93},{"m":"Pneu Furado/Murcho","q":18,"pct":13.95},
            {"m":"Não Carrega/Bateria","q":17,"pct":13.18},{"m":"Pane/Desligamento","q":15,"pct":11.63},
            {"m":"Carregador","q":13,"pct":10.08},
        ],
        "modelos": [
            {"nome":"S9","qtd":44,"pct":34.11},{"nome":"I2","qtd":17,"pct":13.18},
            {"nome":"S10","qtd":15,"pct":11.63},{"nome":"H9","qtd":14,"pct":10.85},
            {"nome":"X2","qtd":8,"pct":6.20},
        ],
        "sintomas": [
            {"nome":"Guidão/Postura","qtd":27,"pct":20.93},{"nome":"Pneu Furado/Murcho","qtd":18,"pct":13.95},
            {"nome":"Não Carrega/Bateria","qtd":17,"pct":13.18},{"nome":"Pane/Desligamento","qtd":15,"pct":11.63},
            {"nome":"Carregador","qtd":13,"pct":10.08},
        ],
        "descontos": 0, "gastosTransportadora": 21586.88, "gastosAT": 16121.11,
        "faturamentoServ": 541535.93,
        "despesasOp": [
            {"nome":"Transportadoras + Correios","val":21586.88,"pct_fat":3.99},
            {"nome":"ATs","val":16121.11,"pct_fat":2.98},
        ],
        "baterias": {"prod":[72,53,83,99,70,67,83,95,47,46],"dist":[
            {"m":"H9/H9X","q":11},{"m":"I2/X2","q":8},{"m":"I3/X3","q":0},
            {"m":"S10/S10X","q":11},{"m":"S8/R8/BR8","q":4},{"m":"S9/S9X","q":12}]},
        "carregadores": {"prod":[54,44,56,83,46,32,43,98,28,55],"dist":[
            {"m":"BR8/R8/S9/S9X/H9/H9X","q":30},{"m":"I3/X3","q":2},
            {"m":"S10/S10X","q":16},{"m":"S6","q":2},{"m":"S8","q":5}]},
        "manutencao": {
            "tecnicos":[
                {"nome":"Aran","meses":[15,15,12,12,13,15,20,28,15,8]},
                {"nome":"Brandon","meses":[10,1,11,12,15,15,17,29,4,4]},
                {"nome":"Maxwell","meses":[19,14,11,14,23,16,10,7,24,13]},
            ],
            "iva":[
                {"nome":"Aran Monteiro","iva":776.0,"pct":96.69,"patrimônios":22},
                {"nome":"Maxwell Gatti","iva":719.2,"pct":85.97,"patrimônios":34},
                {"nome":"Brandon David","iva":451.3,"pct":40.98,"patrimônios":13},
            ],
            "porModelo":[
                {"modelo":"S9","qtd":5},{"modelo":"H9X","qtd":3},{"modelo":"X2","qtd":3},
                {"modelo":"BR8","qtd":2},{"modelo":"H9","qtd":2},{"modelo":"I3","qtd":2},
                {"modelo":"S10","qtd":2},{"modelo":"S10X","qtd":2},{"modelo":"S8","qtd":2},
                {"modelo":"R8","qtd":1},{"modelo":"X3","qtd":1},
            ],
        },
        "horasTecnicos": 527.93, "patrimoniosAtendidos": 54, "mediaHorasProdutivas": 10.02,
    },
}

logistics = {
    "jan": [31600.39, 29878.56, 1721.83, 0, 16974.67, 1815.95, 661557.44],
    "fev": [19058.77, 13262.98, 5795.79, 0, 12995.80, 0, 658323.81],
    "mar": [9593.94, 6362.51, 3231.43, 0, 14879.14, 8628.41, 665968.97],
    "abr": [26570.55, 22284.64, 4285.91, 2397.64, 16300.92, 12658.60, 665202.69],
    "mai": [22988.80, 17832.99, 5155.81, 0, 17934.87, 5453.38, 672777.00],
    "jun": [26280.12, 20846.65, 5433.47, 2515.09, 11978.69, 1809.72, 678100.00],
    "jul": [23409.81, 17827.95, 5581.86, 7694.81, 17791.56, 0, 630295.31],
    "aug": [31072.66, 25930.66, 5142.00, 1383.43, 13299.13, 0, 531393.39],
    "sep": [28382.54, 23774.53, 4608.01, 4841.96, 9537.99, 339.28, 535053.29],
    "oct": [21586.88, 18715.52, 2871.36, 2083.36, 16121.11, 0, 541535.93],
}

oct_uf = [
    ["AC",2230.22],["AM",6114],["BA",40842.70],["CE",25240.21],["DF",15349.51],
    ["GO",3057],["MA",11940],["MG",30181.35],["MS",3418.50],["MT",4950],["PA",62787.90],
    ["PE",33264.12],["PR",14579.19],["RJ",70217],["RS",6675.50],["SE",14013.90],
    ["SP",194648.67],["TO",2026.16],
]

block = ["// ─── ATUALIZAÇÃO OFICIAL — OUTUBRO/2026 ───────────────────────────────"]
block.append(f"DATA.oct={js(oct_data)};")
block.append("const FIN_OFICIAL_OUT=" + js(official) + ";")
block.append("Object.entries(FIN_OFICIAL_OUT).forEach(([m,v])=>{FIN_DATA[m]=v;DATA[m].fin=v;});")
block.append("const DRE_OFICIAL_OUT=" + js(details) + ";")
block.append("Object.entries(DRE_OFICIAL_OUT).forEach(([m,v])=>DRE_DETALHE[m]=v);")
log_obj = {
    month: {
        "logisticaDre": vals[0], "transp": vals[1], "correios": vals[2],
        "restituido": vals[3], "transpLiquido": round(vals[1] - vals[3], 2),
        "ats": vals[4], "descontos": vals[5], "fat": vals[6],
    }
    for month, vals in logistics.items()
}
block.append("const LOGISTICA_DRE_OFICIAL=" + js(log_obj) + ";")
block.append("Object.entries(LOGISTICA_DRE_OFICIAL).forEach(([m,v])=>Object.assign(DESP_OP[m]||(DESP_OP[m]={}),v));")
block.append("DESP_OP_ITENS.splice(0,DESP_OP_ITENS.length,['logisticaDre','Transportadoras + Correios (DRE)'],['ats','Gastos com ATs'],['descontos','Descontos Concedidos']);")
block.append("EXEC.oct={novosQtd:5,novosVal:8150,distratosQtd:2,distratosVal:3855.42,reajustesAprovados:4,reajustesVal:298.06,receitaPontualTotal:19066.92,avanco:'SLA melhorou 3,2 p.p. e a frota voltou a crescer, com cinco novos veículos no mês.',risco:'SLA permanece abaixo de 85% e o Amazonas registrou TMA de 19,5 dias.'};")
block.append("// ─── FIM ATUALIZAÇÃO OFICIAL — OUTUBRO/2026 ───────────────────────────")
block_text = "\n".join(block) + "\n\n"

text = HTML.read_text(encoding="utf-8")
start = "// ─── ATUALIZAÇÃO OFICIAL — OUTUBRO/2026"
end = "// ─── FIM ATUALIZAÇÃO OFICIAL — OUTUBRO/2026 ───────────────────────────\n\n"
if start in text:
    text = re.sub(re.escape(start) + r".*?" + re.escape(end), block_text, text, flags=re.S)
else:
    text = text.replace("// ─── AGREGAÇÃO CONSOLIDADA", block_text + "// ─── AGREGAÇÃO CONSOLIDADA", 1)

replacements = {
    "Jan–Set": "Jan–Out", "Jan a Set": "Jan a Out", "Mai a Set": "Mai a Out",
    "Soma dos 9 meses": "Soma dos 10 meses", "Operacional e Comercial: 9 meses": "Operacional e Comercial: 10 meses",
    "Produção somada (8 meses)": "Produção somada (10 meses)", "Jan–Out (8 meses)": "Jan–Out (10 meses)",
}
for old, new in replacements.items():
    text = text.replace(old, new)
text = text.replace("const LUCROS_MESES_PAGOS=8;", "const LUCROS_MESES_PAGOS=10;")
text = text.replace("let currentMonth='sep'", "let currentMonth='oct'")
text = text.replace("const months=['jan','fev','mar','abr','mai','jun','jul','aug','sep'];", "const months=['jan','fev','mar','abr','mai','jun','jul','aug','sep','oct'];")
text = text.replace("const monthLabels=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set'];", "const monthLabels=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out'];")
text = text.replace("const opsMonths=['jan','fev','mar','abr','mai','jun','jul','aug','sep'];", "const opsMonths=['jan','fev','mar','abr','mai','jun','jul','aug','sep','oct'];")
text = text.replace("const opsMonthLabels=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set'];", "const opsMonthLabels=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out'];")
text = text.replace("const financeMonths=['jan','fev','mar','abr','mai','jun','jul','aug','sep'];", "const financeMonths=['jan','fev','mar','abr','mai','jun','jul','aug','sep','oct'];")
text = text.replace("const financeMonthLabels=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set'];", "const financeMonthLabels=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out'];")
text = text.replace("const mmFin=['jan','fev','mar','abr','mai','jun','jul','aug','sep'];", "const mmFin=['jan','fev','mar','abr','mai','jun','jul','aug','sep','oct'];")
text = text.replace("const mmOps=['jan','fev','mar','abr','mai','jun','jul','aug','sep'];", "const mmOps=['jan','fev','mar','abr','mai','jun','jul','aug','sep','oct'];")

HTML.write_text(text, encoding="utf-8")
print("Atualização de outubro aplicada com DRE oficial e logística conciliada.")
