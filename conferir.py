"""Refaz o censo em Python puro e compara com o que o Excel calculou."""
import csv
import sys
import warnings
from datetime import date, timedelta
from pathlib import Path

from openpyxl import load_workbook

PASTA = Path(__file__).resolve().parent
warnings.filterwarnings("ignore", message="Data Validation extension")


def ler(nome):
    with (PASTA / "dados" / nome).open(encoding="utf-8-sig", newline="") as arquivo:
        return list(csv.DictReader(arquivo, delimiter=";"))


def main():
    base, leitos = ler("internacoes.csv"), ler("leitos.csv")
    inicio, fim = date(2026, 5, 1), date(2026, 10, 31)
    dias = [(inicio + timedelta(i)).isoformat() for i in range((fim - inicio).days + 1)]
    pacientes = sum(r["Entrada"] <= d and (not r["Saída"] or r["Saída"] > d) for d in dias for r in base)
    capacidade = sum(int(r["Leitos"]) for d in dias for r in leitos if r["De"] <= d <= r["Até"])
    painel = load_workbook(PASTA / "painel_leitos.xlsx", data_only=True)["Painel"]
    contas = {
        "ocupação": (painel["B7"].value, pacientes / capacidade),
        "internações": (painel["F7"].value, sum(r["Entrada"] in dias for r in base)),
        "saídas": (painel["H7"].value, sum(r["Saída"] in dias for r in base)),
    }
    erros = [nome for nome, (excel, python) in contas.items() if excel is None or abs(excel - python) > 1e-9]
    print("diverge em: " + ", ".join(erros) if erros else "Excel e Python batem")
    return 1 if erros else 0


if __name__ == "__main__":
    sys.exit(main())
