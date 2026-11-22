# Painel de ocupação de leitos (Excel 2016)

> Protótipo criado com IA generativa, a partir dos meus prompts. O hospital é fictício e os dados são sintéticos: não existe paciente.

Qual unidade está lotada, e em que mês? O hospital fecha seis meses com 77,2% de ocupação, e a média esconde a Pediatria, que chega a 90,1% em julho.

[Baixar a planilha](painel_leitos.xlsx)

![Painel com indicadores, tabela por unidade e dois gráficos](prints/painel.png)

## O que tem na planilha

- Censo de cada dia à meia-noite, com CONT.SES
- Leitos com data de início e fim, porque a quantidade muda
- Filtros de mês e unidade, e meta numa célula
- Só funções do Excel 2016

## Para conferir

| Filtro | Ocupação | Permanência | Internações |
|---|---|---|---|
| Todos, Todas | 77,2% | 4,2 | 3.355 |
| Julho, Pediatria | 90,1% | 3,7 | 107 |
| Agosto, Cirúrgica | 78,8% | 3,3 | 161 |

```
python -m pip install -r requirements.txt
python conferir.py
```

## Gerar do zero (Windows com Excel)

```
python scripts/gerar_dados.py
python scripts/planilha.py
python scripts/recalcular.py painel_leitos.xlsx
```

## Limites

Não acompanha transferência entre unidades. Mortalidade bruta não serve para comparar unidades diferentes.
