"""
Antecipações de clientes (aba Antecipações do painel).

Fonte: Sienge > Relatório de Antecipações/Inadimplências, exportado em XLSX, um arquivo por empresa, salvo em
pasta_antecipacoes (padrão entrada/antecipacoes). Parâmetros: Empresa; Período de recebimento; Plano financeiro
1.01.01.02; Tipo de relatório Antecipação; Ordenação Data Pagamento; Incluir títulos quitados.

Regra: entra a parcela paga com vencimento 2 ou mais meses à frente do mês do pagamento
(mês do vencimento - mês da baixa >= antecipacao_meses). O valor vai para o mês da baixa (Valor Baixa).
Parcelas de TC em excluir_tc (AT = aporte) não entram.

- Vários arquivos da mesma empresa: o mais novo substitui o período que ele cobre (pega estornos e baixas novas).
- dados/antecipacoes_historico.json {empresa: {AAAA-MM: valor}}: meses fechados vindos da planilha manual;
  quando o mês está ali, o "Pago" do mês é esse valor (o Sienge fica só como detalhe).
- dados/antecipacoes_manual.json {empresa: {AAAA-MM: valor}}: linha "A pagar", alimentada pelo painel.
"""
import re, datetime
from collections import defaultdict
from comum import caminho, ler_json

DATA = re.compile(r"^\d{2}/\d{2}/\d{4}$")


def _num(v):
    if v is None or v == "": return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = re.sub(r"[^\d,.-]", "", str(v))
    return float(s.replace(".", "").replace(",", ".")) if s else 0.0


def _d(s):
    return datetime.datetime.strptime(str(s).strip(), "%d/%m/%Y").date()


def _mes(d):
    return f"{d.year}-{d.month:02d}"


def ler_relatorio(f):
    """-> (empresa, (ini, fim), [linhas])"""
    import openpyxl
    ws = openpyxl.load_workbook(f, data_only=True, read_only=True).active
    emp = per = col = None
    linhas = []
    cur_dt = cur_cli = None
    for r in ws.iter_rows(values_only=True):
        r = list(r)
        if not any(v not in (None, "") for v in r): continue
        cab = str(r[0] or "").strip()
        if cab == "Empresa":
            emp = next((str(v).strip() for v in r[1:] if v not in (None, "")), None); continue
        if cab.startswith("Período de recebimento"):
            txt = next((str(v) for v in r[1:] if v not in (None, "")), "")
            ds = re.findall(r"\d{2}/\d{2}/\d{4}", txt)
            if len(ds) == 2: per = (_d(ds[0]), _d(ds[1]))
            continue
        if cab == "Dt. Baixa":
            col = {str(v).strip(): i for i, v in enumerate(r) if v not in (None, "")}; continue
        if not col or cab.startswith("Total"): continue
        g = lambda k: r[col[k]] if k in col and col[k] < len(r) else None
        doc, venc = g("Documento"), g("Dt. Vencto")
        if not doc or not DATA.match(str(venc or "").strip()): continue
        if DATA.match(cab): cur_dt = cab
        if g("Cliente"): cur_cli = str(g("Cliente")).strip()
        if not cur_dt: continue
        linhas.append({"baixa": _d(cur_dt), "cliente": cur_cli, "documento": str(doc).strip(), "parc": str(g("Tit./Parc") or "").strip(),
                       "tc": str(g("TC") or "").strip(), "venc": _d(venc), "valor": round(_num(g("Valor Baixa")), 2),
                       "liquido": round(_num(g("Líquido")), 2)})
    return emp, per, linhas


def montar_antecipacoes(cfg, dados_dir, empreendimentos, avisos):
    pasta = caminho(cfg.get("pasta_antecipacoes", "entrada/antecipacoes"))
    n_meses = int(cfg.get("antecipacao_meses", 2))
    excl_tc = set(cfg.get("excluir_tc", []))
    hist = ler_json(dados_dir / "antecipacoes_historico.json", {})
    manual = ler_json(dados_dir / "antecipacoes_manual.json", {})
    arqs = sorted([f for f in pasta.glob("*.xlsx") if f.is_file() and not f.name.startswith("~$")], key=lambda f: f.stat().st_mtime) if pasta.is_dir() else []

    por_emp, cobertura, fontes = defaultdict(list), {}, defaultdict(list)
    for f in arqs:
        try:
            emp, per, linhas = ler_relatorio(f)
        except Exception as e:
            avisos.append(f"Antecipações: não consegui ler {f.name} ({e})."); continue
        if not emp or not per:
            avisos.append(f"Antecipações: {f.name} não parece o Relatório de Antecipações do Sienge."); continue
        ini, fim = per
        por_emp[emp] = [x for x in por_emp[emp] if not (ini <= x["baixa"] <= fim)] + linhas
        c = cobertura.get(emp)
        cobertura[emp] = (min(c[0], ini), max(c[1], fim)) if c else (ini, fim)
        fontes[emp].append(f.name)

    hoje = max([c[1] for c in cobertura.values()], default=datetime.date.today())
    mes_fim = _mes(hoje)
    inicio_padrao = cfg.get("antecipacao_inicio", "2025-08")
    saida, detalhe = [], []
    for e in empreendimentos:
        emp = e["empresa"]
        ant = []
        for x in por_emp.get(emp, []):
            dm = (x["venc"].year * 12 + x["venc"].month) - (x["baixa"].year * 12 + x["baixa"].month)
            if dm >= n_meses and x["tc"] not in excl_tc:
                ant.append({**x, "meses": dm})
        sienge = defaultdict(lambda: [0.0, 0])
        for x in ant:
            s = sienge[_mes(x["baixa"])]; s[0] += x["valor"]; s[1] += 1
        h, m = hist.get(emp, {}), manual.get(emp, {})
        if h: ini = min(h)
        elif sienge: ini = min(sienge)
        elif emp in cobertura: ini = max(inicio_padrao, _mes(cobertura[emp][0]))
        else: continue
        meses, mm = {}, ini
        while mm <= mes_fim:
            sv, sn = sienge.get(mm, (0.0, 0))
            if mm in h:
                meses[mm] = {"pago": float(h[mm] or 0), "fonte": "planilha", "sienge": round(sv, 2), "titulos": sn}
            else:
                cob = cobertura.get(emp)
                coberto = bool(cob) and _mes(cob[0]) <= mm <= _mes(cob[1])
                meses[mm] = {"pago": round(sv, 2) if coberto else None, "fonte": "sienge" if coberto else "sem dados", "sienge": round(sv, 2), "titulos": sn}
            meses[mm]["a_pagar"] = float(m[mm]) if mm in m and m[mm] not in (None, "") else None
            y, mo = map(int, mm.split("-")); mo += 1
            if mo > 12: mo, y = 1, y + 1
            mm = f"{y}-{mo:02d}"
        cob = cobertura.get(emp)
        saida.append({"empresa": emp, "nome": e["nome"], "inicio": ini, "fim": mes_fim, "meses": meses,
                      "ate_planilha": max(h) if h else None,
                      "sienge_de": cob[0].isoformat() if cob else None, "sienge_ate": cob[1].isoformat() if cob else None,
                      "arquivos": fontes.get(emp, [])})
        for x in ant:
            if _mes(x["baixa"]) >= ini:
                detalhe.append({"empresa": emp, "baixa": x["baixa"].isoformat(), "cliente": x["cliente"], "documento": x["documento"],
                                "parc": x["parc"], "tc": x["tc"], "venc": x["venc"].isoformat(), "meses": x["meses"], "valor": x["valor"]})
    detalhe.sort(key=lambda x: (x["baixa"], x["cliente"] or ""), reverse=True)
    if not arqs:
        avisos.append("Antecipações: nenhum relatório em " + str(pasta) + ".")
    return {"regra_meses": n_meses, "atualizado_ate": hoje.isoformat(), "empresas": saida, "titulos": detalhe}
