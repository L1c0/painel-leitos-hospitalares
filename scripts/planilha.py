import csv
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

from estilo import Estilo

PASTA = Path(__file__).resolve().parent.parent
# verde da folha e ocre da cagaita madura, fruta do cerrado que dá nome ao hospital
e = Estilo("3D4A2A", "B87505", "E9D9A6", "FAF7F0")
INICIO, FIM = date(2026, 5, 1), date(2026, 10, 31)
MESES = ["Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro"]
DATA, PCT = "dd/mm/yyyy", "0.0%"


def ler(nome):
    with (PASTA / "dados" / nome).open(encoding="utf-8-sig") as arquivo:
        return list(csv.DictReader(arquivo, delimiter=";"))


def data(texto):
    return date.fromisoformat(texto) if texto else None


livro = Workbook()
painel = livro.active
painel.title = "Painel"
censo, calc = livro.create_sheet("Censo"), livro.create_sheet("Cálculos")
base, leitos, notas = livro.create_sheet("Internações"), livro.create_sheet("Leitos"), livro.create_sheet("Notas")

linhas = [[int(r["Internação"]), r["Unidade"], data(r["Entrada"]), data(r["Saída"]),
           r["Motivo da saída"] or None, r["Faixa etária"], r["Fonte pagadora"]] for r in ler("internacoes.csv")]
e.tabela(base, "tbInternacoes", ["Internação", "Unidade", "Entrada", "Saída", "Motivo da saída", "Faixa etária", "Fonte pagadora"],
         linhas, {3: DATA, 4: DATA}, [13, 20, 13, 13, 23, 15, 17])
capacidade = [[r["Unidade"], data(r["De"]), data(r["Até"]), int(r["Leitos"])] for r in ler("leitos.csv")]
e.tabela(leitos, "tbLeitos", ["Unidade", "De", "Até", "Leitos"], capacidade, {2: DATA, 3: DATA}, [22, 13, 13, 10])
unidades = list(dict.fromkeys(r[0] for r in capacidade))

U, ENT, SAI = "tbInternacoes[Unidade]", "tbInternacoes[Entrada]", "tbInternacoes[Saída]"
MOTIVO = "tbInternacoes[Motivo da saída]"
diario = []
dia = INICIO
while dia <= FIM:
    for unidade in unidades:
        n = len(diario) + 2
        diario.append([
            dia, unidade,
            f'=SUMIFS(tbLeitos[Leitos],tbLeitos[Unidade],B{n},tbLeitos[De],"<="&A{n},tbLeitos[Até],">="&A{n})',
            # quem já entrou, menos quem já saiu
            f'=COUNTIFS({U},B{n},{ENT},"<="&A{n})-COUNTIFS({U},B{n},{ENT},"<="&A{n},{SAI},"<="&A{n})',
            f"=COUNTIFS({U},B{n},{ENT},A{n})",
            f"=COUNTIFS({U},B{n},{SAI},A{n})",
            f'=COUNTIFS({U},B{n},{SAI},A{n},{MOTIVO},"Óbito")',
            f"=IF(D{n}>=C{n},1,0)",
        ])
    dia += timedelta(1)
e.tabela(censo, "tbCenso", ["Data", "Unidade", "Leitos", "Pacientes", "Entradas", "Saídas", "Óbitos", "Lotado"],
         diario, {1: DATA}, [13, 20, 10, 11, 11, 10, 10, 10])

calc["A1"] = "Contas do painel. Os filtros e a meta ficam nas células amarelas do Painel."
for n, (rotulo, formula) in enumerate([
    ("Unidade escolhida", '=IF(Painel!$G$4="Todas","*",Painel!$G$4)'),
    ("Início do período", '=IF(Painel!$C$4="Todos",$B$12,INDEX($B$12:$B$17,MATCH(Painel!$C$4,$A$12:$A$17,0)))'),
    ("Fim do período", '=IF(Painel!$C$4="Todos",$C$17,INDEX($C$12:$C$17,MATCH(Painel!$C$4,$A$12:$A$17,0)))'),
    ("Meta de ocupação", "=Painel!$K$4"),
], 3):
    calc.cell(n, 1, rotulo)
    calc.cell(n, 2, formula).number_format = PCT if n == 6 else DATA

NOMES = ["Pacientes-dia", "Leitos-dia", "Ocupação", "Entradas", "Saídas", "Óbitos", "Permanência média", "Mortalidade", "Dias lotados"]


def indicadores(n, unidade, inicio, fim):
    def soma(coluna):
        return f'=SUMIFS(tbCenso[{coluna}],tbCenso[Unidade],{unidade},tbCenso[Data],">="&{inicio},tbCenso[Data],"<="&{fim})'

    formulas = [soma("Pacientes"), soma("Leitos"), f"=IFERROR(D{n}/E{n},0)", soma("Entradas"), soma("Saídas"),
                soma("Óbitos"), f"=IFERROR(D{n}/H{n},0)", f"=IFERROR(I{n}/H{n},0)", soma("Lotado")]
    for c, (formula, formato) in enumerate(zip(formulas, ["#,##0", "#,##0", PCT, "#,##0", "#,##0", "#,##0", "0.0", PCT, "#,##0"], strict=True), 4):
        calc.cell(n, c, formula).number_format = formato


e.cabecalho(calc, 8, ["Com os filtros", "", ""] + NOMES)
indicadores(9, "$B$3", "$B$4", "$B$5")
e.cabecalho(calc, 11, ["Mês", "Início", "Fim"] + NOMES + ["Meta"])
for i, mes in enumerate(MESES):
    n = 12 + i
    calc.cell(n, 1, mes)
    calc.cell(n, 2, date(2026, 5 + i, 1)).number_format = DATA
    calc.cell(n, 3, f"=EOMONTH(B{n},0)").number_format = DATA
    indicadores(n, "$B$3", f"B{n}", f"C{n}")
    calc.cell(n, 13, "=$B$6").number_format = PCT
e.cabecalho(calc, 20, ["Unidade", "", ""] + NOMES)
for n, unidade in enumerate(unidades, 21):
    calc.cell(n, 1, unidade)
    indicadores(n, f"$A{n}", "$B$4", "$B$5")
calc["O11"], calc["P11"] = "Lista de meses", "Lista de unidades"
for n, mes in enumerate(["Todos"] + MESES, 12):
    calc.cell(n, 15, mes)
for n, unidade in enumerate(["Todas"] + unidades, 12):
    calc.cell(n, 16, unidade)
for linha in calc.iter_rows():
    for celula in linha:
        if celula.row not in (8, 11, 20) or celula.column > 13:
            celula.font = e.fonte()
calc.column_dimensions["A"].width = 24
for c in "BCDEFGHIJKLMOP":
    calc.column_dimensions[c].width = 16

e.painel(painel, "Hospital Lagoa da Cagaita: ocupação de leitos",
         "Hospital fictício de Montes Claros. Dados sintéticos de maio a outubro de 2026, sem nenhum paciente real.", "N")
for c, largura in zip("ABCDEFGHIJKLMN", [2, 17, 13, 14, 14, 13, 17, 13, 12, 12, 12, 12, 12, 2], strict=True):
    painel.column_dimensions[c].width = largura
e.filtro(painel, "B4", "C4", "Mês", "Todos")
e.filtro(painel, "F4", "G4", "Unidade", "Todas")
painel.merge_cells("I4:J4")
e.filtro(painel, "I4", "K4", "Meta de ocupação", 0.85, "0%")
for celula, regra in (
    ("C4", DataValidation(type="list", formula1="'Cálculos'!$O$12:$O$18")),
    ("G4", DataValidation(type="list", formula1="'Cálculos'!$P$12:$P$17")),
    ("K4", DataValidation(type="decimal", operator="between", formula1="0.5", formula2="1")),
):
    regra.showErrorMessage = True
    painel.add_data_validation(regra)
    regra.add(celula)
for inicio, fim, rotulo, coluna, formato in (
    ("B", "C", "Taxa de ocupação", "F", PCT), ("D", "E", "Permanência média (dias)", "J", "0.0"),
    ("F", "G", "Internações no período", "G", "#,##0"), ("H", "I", "Saídas no período", "H", "#,##0"),
    ("J", "K", "Mortalidade (óbitos / saídas)", "K", PCT), ("L", "M", "Dias lotados (soma das unidades)", "L", "#,##0"),
):
    e.cartao(painel, inicio, fim, rotulo, f"=Cálculos!{coluna}9", formato)
e.alerta(painel, "B7:C7", "$B$7>$K$4")

e.cabecalho(painel, 9, ["Unidade", "Ocupação", "Permanência", "Internações", "Saídas", "Óbitos", "Dias lotados"], 2)
for i, unidade in enumerate(unidades):
    n = 10 + i
    painel.cell(n, 2, unidade)
    for c, (coluna, formato) in enumerate([("F", PCT), ("J", "0.0"), ("G", "#,##0"), ("H", "#,##0"), ("I", "#,##0"), ("L", "#,##0")], 3):
        painel.cell(n, c, f"=Cálculos!{coluna}{21 + i}").number_format = formato
    for c in range(2, 9):
        painel.cell(n, c).fill = e.cor("FFFFFF")
e.alerta(painel, "C10:C14", "$C10>$K$4")
painel["B15"] = "A tabela segue o mês escolhido e sempre mostra as cinco unidades. Vermelho: acima da meta."
painel["B15"].font = e.fonte(9, "5C554D", italic=True)


def arrumar(grafico, titulo):
    grafico.title = titulo
    grafico.title.overlay = False
    grafico.roundedCorners = False
    grafico.y_axis.scaling.min, grafico.y_axis.scaling.max = 0, 1
    grafico.y_axis.number_format, grafico.y_axis.majorUnit = "0%", 0.25
    grafico.y_axis.majorGridlines = None
    grafico.x_axis.delete = grafico.y_axis.delete = False
    grafico.height, grafico.width = 8, 13.3


por_mes = BarChart()
por_mes.add_data(Reference(calc, min_col=6, min_row=11, max_row=17), titles_from_data=True)
por_mes.set_categories(Reference(calc, min_col=1, min_row=12, max_row=17))
por_mes.series[0].graphicalProperties.solidFill = "177E52"
por_mes.gapWidth = 70
arrumar(por_mes, "Ocupação por mês (segue a unidade)")
linha_meta = LineChart()
linha_meta.add_data(Reference(calc, min_col=13, min_row=11, max_row=17), titles_from_data=True)
linha_meta.series[0].graphicalProperties.line.solidFill = "B87505"
linha_meta.series[0].graphicalProperties.line.width = 22000
linha_meta.series[0].marker.symbol = "none"
por_mes += linha_meta
por_mes.legend.position = "b"
por_mes.legend.overlay = False
painel.add_chart(por_mes, "B17")

por_unidade = BarChart()
por_unidade.type = "bar"
por_unidade.add_data(Reference(calc, min_col=6, min_row=20, max_row=25), titles_from_data=True)
por_unidade.set_categories(Reference(calc, min_col=1, min_row=21, max_row=25))
por_unidade.series[0].graphicalProperties.solidFill = "177E52"
por_unidade.gapWidth = 60
por_unidade.legend = None
arrumar(por_unidade, "Ocupação por unidade (segue o mês)")
por_unidade.x_axis.scaling.orientation = "maxMin"
painel.add_chart(por_unidade, "H17")
painel.freeze_panes = "A5"

e.notas(notas, [
    "#Sobre esta planilha",
    "Protótipo criado com IA generativa. Hospital fictício e dados sintéticos.",
    "#Decisões",
    "O paciente é contado no leito à meia-noite. Quem entra hoje conta hoje. Quem sai hoje não conta.",
    "Taxa de ocupação é pacientes-dia dividido por leitos-dia. Os dias são somados antes de dividir, sem média de porcentagens.",
    "Permanência média é pacientes-dia dividido por saídas, com os óbitos dentro das saídas.",
    "Mortalidade aqui é a bruta: todos os óbitos divididos por todas as saídas. Não serve para comparar a UTI com a maternidade.",
    "Os leitos têm data de início e fim. A Clínica Médica ganhou 4 leitos em 15/07 e a Cirúrgica perdeu 5 entre 10/08 e 21/08, por reforma.",
    "A meta de 85% é um parâmetro do exercício. Cada hospital define a sua.",
    "#O que esta planilha não faz",
    "Não acompanha transferência entre unidades do mesmo hospital: cada internação fica numa unidade só.",
    "Não tem nome, documento, diagnóstico nem data de nascimento. Com dado real, precisaria de autorização e de cuidado com a LGPD.",
])
livro.save(PASTA / "painel_leitos.xlsx")
