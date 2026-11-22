from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


class Estilo:
    def __init__(self, escuro, destaque, claro, fundo):
        self.escuro, self.destaque, self.claro, self.fundo = escuro, destaque, claro, fundo

    def fonte(self, tamanho=10, cor="2B2723", **extras):
        return Font(name="Arial", size=tamanho, color=cor, **extras)

    def cor(self, cor):
        return PatternFill("solid", fgColor=cor)

    def cabecalho(self, aba, linha, nomes, coluna=1):
        for i, nome in enumerate(nomes):
            celula = aba.cell(linha, coluna + i, nome)
            celula.font = self.fonte(cor="FFFFFF", bold=True)
            celula.fill = self.cor(self.escuro)
            celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def tabela(self, aba, nome, nomes, linhas, formatos, larguras):
        self.cabecalho(aba, 1, nomes)
        for n, linha in enumerate(linhas, 2):
            for c, valor in enumerate(linha, 1):
                celula = aba.cell(n, c, valor)
                celula.font = self.fonte()
                if c in formatos:
                    celula.number_format = formatos[c]
        for c, largura in enumerate(larguras):
            aba.column_dimensions[chr(65 + c)].width = largura
        tabela = Table(displayName=nome, ref=f"A1:{chr(64 + len(nomes))}{len(linhas) + 1}")
        tabela.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
        aba.add_table(tabela)
        aba.freeze_panes = "A2"

    def painel(self, aba, titulo, aviso, ultima):
        aba.sheet_view.showGridLines = False
        aba.sheet_view.zoomScale = 90
        aba.sheet_properties.tabColor = self.escuro
        for linha in aba[f"A1:{ultima}44"]:
            for celula in linha:
                celula.fill = self.cor(self.fundo)
                celula.font = self.fonte()
        for linha in aba[f"A1:{ultima}2"]:
            for celula in linha:
                celula.fill = self.cor(self.escuro)
        aba["B1"] = titulo
        aba["B1"].font = self.fonte(17, "FFFFFF", bold=True)
        aba["B2"] = aviso
        aba["B2"].font = self.fonte(9, self.claro)
        aba.row_dimensions[1].height = 34
        aba.page_setup.orientation = "landscape"
        aba.page_setup.paperSize = 9
        aba.sheet_properties.pageSetUpPr.fitToPage = True
        aba.page_setup.fitToHeight = 1
        aba.print_area = f"A1:{ultima}44"

    def filtro(self, aba, rotulo, valor, texto, inicial, formato=None):
        aba[rotulo] = texto
        aba[rotulo].font = self.fonte(bold=True)
        aba[rotulo].alignment = Alignment(horizontal="right")
        aba[valor] = inicial
        aba[valor].font = self.fonte(11, "0000FF", bold=True)
        aba[valor].fill = self.cor("FFF2B3")
        aba[valor].alignment = Alignment(horizontal="center")
        if formato:
            aba[valor].number_format = formato

    def cartao(self, aba, inicio, fim, rotulo, formula, formato, linha=6):
        traco = Side(style="medium", color=self.destaque)
        aba.merge_cells(f"{inicio}{linha}:{fim}{linha}")
        aba.merge_cells(f"{inicio}{linha + 1}:{fim}{linha + 1}")
        topo, valor = aba[f"{inicio}{linha}"], aba[f"{inicio}{linha + 1}"]
        topo.value, valor.value = rotulo, formula
        topo.font = self.fonte(9, "5C554D")
        valor.font = self.fonte(19, self.escuro, bold=True)
        valor.number_format = formato
        for celulas in aba[f"{inicio}{linha}:{fim}{linha + 1}"]:
            for celula in celulas:
                celula.fill = self.cor("FFFFFF")
                celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if celula.row == linha:
                    celula.border = Border(top=traco)
        aba.row_dimensions[linha].height = 27
        aba.row_dimensions[linha + 1].height = 38

    def alerta(self, aba, faixa, formula):
        aba.conditional_formatting.add(
            faixa, FormulaRule(formula=[formula], fill=self.cor("F6D5CC"), font=Font(color="8A2A16", bold=True))
        )

    def notas(self, aba, linhas):
        aba.sheet_view.showGridLines = False
        aba.column_dimensions["A"].width = 115
        for n, texto in enumerate(linhas, 1):
            celula = aba.cell(n, 1, texto.lstrip("#"))
            celula.alignment = Alignment(wrap_text=True, vertical="top")
            celula.font = self.fonte(12, self.escuro, bold=True) if texto.startswith("#") else self.fonte()
