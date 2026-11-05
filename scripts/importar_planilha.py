"""
Importação inicial (rodar uma vez): traz o histórico da planilha de inadimplência para a pasta de dados do painel.

  python scripts/importar_planilha.py --planilha "Inadimplência Carozzo 05112026.xlsx" --esperado "Consolidado_Titulos_Boletagem_Mar-Set2026_5.xlsx"

O que é importado:
- aba " 📋 Tratados"        -> dados/tratados.json  (títulos tratados e o Status de Saída)
- aba "֍ Variação Semanal"  -> dados/semanal.json   (saldo semanal por empreendimento, R$ mil -> R$)
- aba " 📋 Base Anterior"    -> dados/retratos/<data>.json (semana anterior, para o comparativo e a recuperação)
- planilha de esperado      -> dados/esperado_manual.json (recebimento esperado mar-out/26; Feira sem JDL)
Parcelas que o Excel transformou em data ("3/5" -> 05/03) voltam a ser texto.
"""
import sys, datetime
from collections import defaultdict
import openpyxl
from comum import pasta_dados, gravar_json, parc_texto, chave, cfg as ler_cfg

arg = lambda n: sys.argv[sys.argv.index(n) + 1] if n in sys.argv else None
DADOS = pasta_dados()
cfg = ler_cfg()
REGRA = int(cfg.get("regra_dias", 10))
MES = {"Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
       "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12}


def d_iso(v):
    if isinstance(v, datetime.datetime): return v.date().isoformat()
    if isinstance(v, datetime.date): return v.isoformat()
    return str(v)[:10] if v else None


def f(v):
    try: return float(str(v).replace(",", ".")) if v not in (None, "") else 0.0
    except ValueError: return 0.0


def titulo(r):
    """Linha das abas Base/Tratados (colunas A..Y) -> título do painel."""
    return {"chave": chave(r[3], r[5]), "venc": d_iso(r[0]), "cliente": str(r[1] or "").strip(), "cd_cliente": str(r[2]).strip(),
            "documento": str(r[3]).strip(), "titulo": str(r[4]).strip(), "parc": parc_texto(r[5]), "tc": str(r[6] or "").strip(),
            "unid": str(r[7] or "").strip(), "valor_original": f(r[8]), "total": f(r[22]), "dias": int(f(r[14])),
            "empresa": str(r[24] or "").strip(), "cd_empresa": str(r[23]).strip()}


planilha = arg("--planilha")
if planilha and (DADOS / "tratados.json").exists() and "--forcar" not in sys.argv:
    sys.exit("A pasta de dados já tem Tratados (dados/tratados.json). Importar de novo apagaria o que foi classificado pelo painel.\n"
             "Se for isso mesmo, rode de novo com --forcar.")
if planilha:
    wb = openpyxl.load_workbook(planilha, data_only=True)
    nomes = {ws.title.strip(): ws for ws in wb.worksheets}
    aba = lambda trecho: next(ws for t, ws in nomes.items() if trecho in t)

    # --- Tratados ---
    trat = []
    for r in aba("Tratados").iter_rows(min_row=2, values_only=True):
        if not r[0] or not r[3]: continue
        t = titulo(r)
        t.update(status=(str(r[27]).strip() if r[27] else ""), obs=(str(r[28]).strip() if r[28] else ""),
                 tratado_em=None, origem="planilha")
        trat.append(t)
    gravar_json(DADOS / "tratados.json", trat)
    print(f"Tratados: {len(trat)} títulos")

    # --- Variação semanal (R$ mil) ---
    ws = aba("Variação Semanal")
    rot = {"Ícone": "TERRAZZA FEIRA SPE LTDA", "Balena": "TERRAZZA BALENA SPE LTDA", "Carozzo Wellness": "CAROZZO WELLNESS JARDIM ARMAÇÃO LTDA"}
    sem = defaultdict(dict)
    datas_icone = []
    for r in range(1, 30):
        nome = ws.cell(r, 1).value
        if nome in rot:
            # linha de datas e linha de valores logo abaixo do nome
            rd = next(x for x in range(r + 1, r + 4)
                      if any(isinstance(ws.cell(x, c).value, datetime.datetime) for c in range(1, ws.max_column + 1)))
            for c in range(1, ws.max_column + 1):
                d, v = ws.cell(rd, c).value, ws.cell(rd + 1, c).value
                if isinstance(d, datetime.datetime) and v is not None:
                    sem[d.date().isoformat()][rot[nome]] = round(f(v) * 1000, 2)
                    if nome == "Ícone": datas_icone.append(d.date().isoformat())
    # cada empreendimento foi anotado em dias um pouco diferentes (ex.: 04/09 e 05/09): junta datas a até 3 dias
    grupos = []
    for d in sorted(sem):
        dd = datetime.date.fromisoformat(d)
        if grupos and (dd - grupos[-1][0]).days <= 3: grupos[-1][1].append(d)
        else: grupos.append([dd, [d]])
    semanal = []
    for _, ds in grupos:
        v = {}
        for d in ds: v.update(sem[d])
        semanal.append({"data": max(ds), "valores": v})
    gravar_json(DADOS / "semanal.json", semanal)
    print(f"Variação semanal: {len(semanal)} semanas ({semanal[0]['data']} a {semanal[-1]['data']})")

    # --- Base Anterior -> retrato da semana anterior ---
    data_ant = arg("--data-anterior") or sorted(datas_icone)[-2]
    base = [titulo(r) for r in aba("Base Anterior").iter_rows(min_row=2, values_only=True) if r[0] and r[3]]
    subj = {}
    for r in aba("Base Anterior").iter_rows(min_row=2, values_only=True):
        if r[0] and r[3]: subj[chave(r[3], r[5])] = str(r[26] or "").strip().lower() == "sim"
    maior = defaultdict(int)
    for t in base: maior[t["cd_cliente"]] = max(maior[t["cd_cliente"]], t["dias"])
    for t in base: t.update(dias_aj=maior[t["cd_cliente"]], subjudice=subj.get(t["chave"], False))
    gravar_json(DADOS / "retratos" / f"{data_ant}.json", {"data": data_ant, "titulos": base, "origem": "planilha (Base Anterior)"})
    print(f"Base anterior: {len(base)} títulos, R$ {sum(t['total'] for t in base):,.2f} -> retratos/{data_ant}.json")

esperado = arg("--esperado")
if esperado:
    wb = openpyxl.load_workbook(esperado, data_only=False)
    out = defaultdict(dict)
    cons = wb["Consolidado"]
    for r in range(6, 14):
        m = MES.get(str(cons.cell(r, 1).value).strip())
        if not m: continue
        mes = f"2026-{m:02d}"
        for col, emp in ((2, "TERRAZZA BALENA SPE LTDA"), (4, "CAROZZO WELLNESS JARDIM ARMAÇÃO LTDA")):
            v = cons.cell(r, col).value
            if isinstance(v, (int, float)): out[emp][mes] = round(v, 2)
    fs = wb["Feira sem JDL"]
    for r in range(6, 14):
        m = MES.get(str(fs.cell(r, 1).value).strip())
        b, c = fs.cell(r, 2).value, fs.cell(r, 3).value
        if m and isinstance(b, (int, float)):
            out["TERRAZZA FEIRA SPE LTDA"][f"2026-{m:02d}"] = round(b - (c or 0), 2)
    gravar_json(DADOS / "esperado_manual.json", out)
    for e, ms in out.items():
        print(f"Esperado {e}: " + ", ".join(f"{k} {v:,.2f}" for k, v in sorted(ms.items())))
