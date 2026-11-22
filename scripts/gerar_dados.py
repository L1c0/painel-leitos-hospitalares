# Dados sintéticos: nenhum paciente existe e o hospital é fictício.
import csv
import random
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

sorteio = random.Random(3817)
DADOS = Path(__file__).resolve().parent.parent / "dados"
INICIO, FIM = date(2026, 5, 1), date(2026, 10, 31)

LEITOS = [
    ("Clínica Médica", date(2026, 5, 1), date(2026, 7, 14), 32),
    ("Clínica Médica", date(2026, 7, 15), FIM, 36),
    ("Cirúrgica", date(2026, 5, 1), date(2026, 8, 9), 24),
    ("Cirúrgica", date(2026, 8, 10), date(2026, 8, 21), 19),
    ("Cirúrgica", date(2026, 8, 22), FIM, 24),
    ("Pediatria", INICIO, FIM, 14),
    ("UTI Adulto", INICIO, FIM, 10),
    ("Maternidade", INICIO, FIM, 17),
]
# chegadas por dia, permanência média e chance de óbito
PERFIL = {
    "Clínica Médica": (4.7, 6.6, 0.051),
    "Cirúrgica": (5.6, 3.3, 0.012),
    "Pediatria": (2.5, 3.4, 0.004),
    "UTI Adulto": (1.15, 7.4, 0.168),
    "Maternidade": (4.6, 2.6, 0.0),
}


def leitos_no_dia(unidade, dia):
    dia = min(max(dia, INICIO), FIM)
    return next(q for u, de, ate, q in LEITOS if u == unidade and de <= dia <= ate)


def chegadas(unidade, dia):
    media = PERFIL[unidade][0]
    if unidade == "Cirúrgica":
        media *= 0.38 if dia.weekday() >= 5 else 1.22
    # bronquiolite e gripe lotam a pediatria no outono e no inverno
    if unidade == "Pediatria" and 5 <= dia.month <= 7:
        media *= 1.75
    if unidade == "Clínica Médica" and dia.month in (6, 7):
        media *= 1.22
    return sum(sorteio.random() < media / 40 for _ in range(40))


def faixa_etaria(unidade):
    if unidade == "Pediatria":
        return sorteio.choices(["0 a 4", "5 a 14"], [68, 32])[0]
    if unidade == "Maternidade":
        return sorteio.choices(["15 a 29", "30 a 44"], [63, 37])[0]
    return sorteio.choices(["15 a 29", "30 a 44", "45 a 59", "60 a 74", "75 ou mais"], [9, 15, 22, 31, 23])[0]


def gerar():
    ocupados = Counter()
    internacoes = []
    numero = 60418
    dia = date(2026, 3, 22)
    while dia <= FIM:
        for unidade, (_, permanencia, risco) in PERFIL.items():
            for _ in range(chegadas(unidade, dia)):
                noites = max(1, round(sorteio.gammavariate(2.3, permanencia / 2.3)))
                if sorteio.random() < 0.018:
                    noites = 0
                saida = dia + timedelta(noites)
                # quase não tem alta no domingo
                if saida.weekday() == 6 and noites and sorteio.random() < 0.55:
                    saida += timedelta(1)
                estadia = [dia + timedelta(i) for i in range((saida - dia).days)]
                if any(ocupados[unidade, d] >= leitos_no_dia(unidade, d) for d in estadia):
                    continue
                for d in estadia:
                    ocupados[unidade, d] += 1
                numero += sorteio.choices([1, 2, 5], [93, 5, 2])[0]
                if saida < INICIO:
                    continue
                motivo = "Óbito" if sorteio.random() < risco else sorteio.choices(["Alta", "Transferência externa"], [96, 4])[0]
                fonte = sorteio.choices(["SUS", "Convênio", "Particular"], [74, 21, 5])[0]
                ainda_internado = saida > FIM
                internacoes.append([numero, unidade, dia, None if ainda_internado else saida,
                                    "" if ainda_internado else motivo, faixa_etaria(unidade), fonte])
        dia += timedelta(1)
    return internacoes


def gravar(nome, cabecalho, linhas):
    with (DADOS / nome).open("w", newline="", encoding="utf-8-sig") as arquivo:
        saida = csv.writer(arquivo, delimiter=";")
        saida.writerow(cabecalho)
        saida.writerows(linhas)


if __name__ == "__main__":
    gravar("internacoes.csv", ["Internação", "Unidade", "Entrada", "Saída", "Motivo da saída", "Faixa etária", "Fonte pagadora"], gerar())
    gravar("leitos.csv", ["Unidade", "De", "Até", "Leitos"], LEITOS)
