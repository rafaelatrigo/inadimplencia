"""
Monta a planilha de inadimplência da semana (Inadimplência_Carozzo_DDMMAAAA.xlsx) a partir da planilha anterior,
do CSV do Contas a Receber e (opcional) do PDF de Títulos em Cobrança do Sienge.

  python scripts/montar_planilha.py --anterior "...\\Inadimplência_Carozzo_05112026.xlsx" --csv "relatorio - 2026-11-12....csv"
         [--cobranca relatorio_40.pdf] --data 12/11/2026 --saida "...\\Inadimplência_Carozzo_12112026.xlsx"

O que faz (o mesmo passo a passo da planilha, sem copiar e colar):
- Base Anterior  <- a Base Atual da planilha anterior (valores), com as fórmulas AB/AC refeitas
- Base Atual     <- CSV: vencidos até a data, sem os Tratados, sem TC excluídos (config: excluir_tc, ex.: AT) e sem
                    ignorar_titulos; regra dos 10 dias (cliente entra se tiver algum título com mais de 10 dias)
- Base Posição Clientes <- último lançamento de cada cliente no PDF de cobrança (sem PDF: mantém o status anterior)
- Variação Semanal <- nova coluna da semana (R$ mil) e gráficos estendidos até ela
- Fórmulas das outras abas: intervalos das bases ajustados ao novo número de linhas; o Excel recalcula ao abrir
Os gráficos, imagens e formatação da planilha são preservados (edição direta do arquivo, sem regravar pelo openpyxl).
A aba Tratados é mantida como está.
"""
import sys, re, csv, zipfile, datetime, shutil, pathlib, warnings
from collections import defaultdict
from xml.sax.saxutils import escape
import openpyxl
from comum import chave, parc_texto, cfg as ler_cfg

arg = lambda n, d=None: sys.argv[sys.argv.index(n) + 1] if n in sys.argv else d
cfg = ler_cfg()
REGRA = int(cfg.get("regra_dias", 10))
EXCL_TC = set(cfg.get("excluir_tc", []))
IGN = {chave(x["documento"], p) for x in cfg.get("ignorar_titulos", []) for p in x.get("parcelas", [])}
TEXTO = {"B", "D", "F", "G", "Y"}          # colunas de texto nas bases (Parc sempre texto!)
DATAS = {"A", "K"}
COLS = [openpyxl.utils.get_column_letter(i) for i in range(1, 26)]   # A..Y
NOMES_EMP = {"TERRAZZA FEIRA SPE LTDA": "Ícone", "TERRAZZA BALENA SPE LTDA": "Balena",
             "CAROZZO WELLNESS JARDIM ARMAÇÃO LTDA": "Carozzo Wellness"}


def serial(d):
    if isinstance(d, str):
        d = datetime.date.fromisoformat(d[:10])
    if isinstance(d, datetime.datetime):
        d = d.date()
    return (d - datetime.date(1899, 12, 30)).days


def num(v):
    try: return float(str(v).replace(",", ".")) if v not in (None, "") else None
    except ValueError: return None


def cel(ref, v, s, forcar_texto=False):
    st = f' s="{s}"' if s else ""
    if v is None or v == "":
        return f'<c r="{ref}"{st}/>'
    if not forcar_texto and isinstance(v, (int, float)) and not isinstance(v, bool):
        return f'<c r="{ref}"{st}><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'
    return f'<c r="{ref}"{st} t="inlineStr"><is><t xml:space="preserve">{escape(str(v))}</t></is></c>'


def formula(ref, f, s, texto=False):
    st = f' s="{s}"' if s else ""
    tp = ' t="str"' if texto else ""
    return f'<c r="{ref}"{st}{tp}><f>{escape(f)}</f></c>'


def estilos_linha(xml, linha):
    m = re.search(r'<row [^>]*r="%d"[^>]*>(.*?)</row>' % linha, xml, re.S)
    return dict(re.findall(r'<c r="([A-Z]+)\d+"[^>]*?s="(\d+)"', m.group(1))) if m else {}


def atributos_linha(xml, linha):
    m = re.search(r'<row ([^>]*r="%d"[^>]*)>' % linha, xml)
    a = m.group(1) if m else 'r="%d"' % linha
    return re.sub(r'\s*r="\d+"', "", re.sub(r'\s*spans="[^"]*"', "", a))


def trocar_linhas(xml, novas_linhas, ultima_col):
    """Mantém a linha 1 (cabeçalho) e substitui todas as outras linhas da aba."""
    ini = xml.index("<sheetData>") + len("<sheetData>"); fim = xml.index("</sheetData>")
    dados = xml[ini:fim]
    m1 = re.search(r'<row [^>]*r="1"[^>]*>.*?</row>|<row [^>]*r="1"[^>]*/>', dados, re.S)
    cab = m1.group(0) if m1 else ""
    xml = xml[:ini] + cab + "".join(novas_linhas) + xml[fim:]
    n = len(novas_linhas) + 1
    xml = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:{ultima_col}{n}"/>', xml)
    xml = re.sub(r'(<autoFilter ref=")A1:([A-Z]+)\d+', lambda m: f'{m.group(1)}A1:{m.group(2)}{n}', xml)
    return xml


def main():
    f_ant, f_csv, f_cob = pathlib.Path(arg("--anterior")), pathlib.Path(arg("--csv")), arg("--cobranca")
    data = datetime.datetime.strptime(arg("--data"), "%d/%m/%Y").date()
    saida = pathlib.Path(arg("--saida"))

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wbv = openpyxl.load_workbook(f_ant, data_only=True)
    abas = {ws.title.strip(): ws.title for ws in wbv.worksheets}
    nome = lambda trecho: next(t for k, t in abas.items() if trecho in k)
    A_ATUAL, A_ANT, A_TRAT, A_POS, A_VAR = (nome("Base Atual"), nome("Base Anterior"), nome("Tratados"),
                                             nome("Posição Clientes"), nome("Variação Semanal"))
    linhas = lambda ws: [r for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0] and len(r) > 24 and r[3]]

    base_old = linhas(wbv[A_ATUAL])
    k_trat = {chave(r[3], r[5]) for r in linhas(wbv[A_TRAT])}
    pos_old = {}
    for r in wbv[A_POS].iter_rows(min_row=2, values_only=True):
        if r and len(r) > 5 and r[2] and r[5]: pos_old[str(r[2]).strip().upper()] = str(r[5])

    # ---- nova base ----
    rows = list(csv.DictReader(open(f_csv, encoding="utf-8-sig")))
    vis = [r for r in rows if r["Data vecto"][:10] <= data.isoformat() and r["TC"] not in EXCL_TC
           and chave(r["Documento"], r["Parc"]) not in IGN and chave(r["Documento"], r["Parc"]) not in k_trat]
    maior = defaultdict(int)
    for r in vis: maior[r["Cd Cliente"]] = max(maior[r["Cd Cliente"]], int(float(r["Dias"] or 0)))
    base = sorted([r for r in vis if maior[r["Cd Cliente"]] > REGRA],
                  key=lambda r: (r["Nome Empresa"], r["Cliente"], r["Data vecto"]))
    cab_csv = list(rows[0].keys())

    # ---- status da cobrança (PDF) ----
    status = {}
    if f_cob:
        import gerar_dados
        cob = gerar_dados.ler_cobranca(pathlib.Path(f_cob))
        nomes = {}
        for r in rows: nomes[r["Cd Cliente"].strip()] = r["Cliente"].strip().upper()
        for cd, st in cob["status"].items():
            nm = nomes.get(cd) or ""
            if nm: status[nm] = f'{datetime.date.fromisoformat(st["data"]):%d/%m/%Y} - {st["texto"]}'

    tmp = saida.with_suffix(".tmp.d")
    if tmp.exists(): shutil.rmtree(tmp)
    with zipfile.ZipFile(f_ant) as z: z.extractall(tmp)
    wbx = (tmp / "xl/workbook.xml").read_text(encoding="utf-8")
    rels = (tmp / "xl/_rels/workbook.xml.rels").read_text(encoding="utf-8")
    alvo = {}
    for n_, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        m = re.search(r'<Relationship [^>]*Id="%s"[^>]*>' % rid, rels)
        alvo[n_.replace("&amp;", "&")] = tmp / "xl" / re.search(r'Target="([^"]+)"', m.group(0)).group(1)
    ler = lambda a: alvo[a].read_text(encoding="utf-8")
    gravar = lambda a, x: alvo[a].write_text(x, encoding="utf-8")

    n_new, n_old = len(base), len(base_old)
    q = lambda a: "'" + a.replace("'", "''") + "'!"

    # ---- Base Atual ----
    x = ler(A_ATUAL); es = estilos_linha(x, 2); attr = atributos_linha(x, 2); out = []
    for i, r in enumerate(base, start=2):
        cs = []
        for j, col in enumerate(COLS):
            v = r[cab_csv[j]]
            if col in DATAS and v: v = serial(v)
            elif col == "F": v = parc_texto(v)
            elif col not in TEXTO and num(v) is not None: v = num(v); v = int(v) if v == int(v) and col not in ("I", "L", "M", "N", "P", "Q", "R", "S", "T", "U", "V", "W") else v
            cs.append(cel(f"{col}{i}", v, es.get(col), forcar_texto=(col in TEXTO)))
        cs.append(formula(f"Z{i}", f"_xlfn.MAXIFS(O$2:O${n_new + 1},B$2:B${n_new + 1},B{i})", es.get("Z")))
        cs.append(cel(f"AA{i}", "Não", es.get("AA"), True))
        out.append(f'<row r="{i}" {attr}>' + "".join(cs) + "</row>")
    gravar(A_ATUAL, trocar_linhas(x, out, "AA"))

    # ---- Base Anterior (= Base Atual da semana passada) ----
    x = ler(A_ANT); es = estilos_linha(x, 2); attr = atributos_linha(x, 2); out = []
    for i, r in enumerate(base_old, start=2):
        cs = []
        for j, col in enumerate(COLS):
            v = r[j]
            if col in DATAS and isinstance(v, (datetime.date, datetime.datetime)): v = serial(v)
            elif col == "F": v = parc_texto(v)
            elif col not in TEXTO and num(v) is not None: v = num(v)
            cs.append(cel(f"{col}{i}", v, es.get(col), forcar_texto=(col in TEXTO)))
        cs.append(cel(f"Z{i}", r[25] if len(r) > 25 else None, es.get("Z")))
        cs.append(cel(f"AA{i}", (r[26] if len(r) > 26 else None) or "Não", es.get("AA"), True))
        cs.append(formula(f"AB{i}", f'IF(COUNTIFS({q(A_TRAT)}D:D,D{i},{q(A_TRAT)}F:F,F{i})=0,"NÃO","SIM")', es.get("AB"), True))
        cs.append(formula(f"AC{i}", f'IF(COUNTIFS({q(A_ATUAL)}D$2:D${n_new + 1},D{i},{q(A_ATUAL)}F$2:F${n_new + 1},F{i})=0,"SIM","NÃO")', es.get("AC"), True))
        out.append(f'<row r="{i}" {attr}>' + "".join(cs) + "</row>")
    gravar(A_ANT, trocar_linhas(x, out, "AC"))

    # ---- Base Posição Clientes (um por cliente da base nova) ----
    x = ler(A_POS); es = estilos_linha(x, 2); attr = atributos_linha(x, 2); out = []; vistos = set()
    for r in base:
        nm = r["Cliente"].strip()
        if nm.upper() in vistos: continue
        vistos.add(nm.upper()); i = len(out) + 2
        st = status.get(nm.upper()) or pos_old.get(nm.upper(), "")
        out.append(f'<row r="{i}" {attr}>' + cel(f"A{i}", None, es.get("A")) + cel(f"B{i}", r["Nome Empresa"], es.get("B"), True)
                   + cel(f"C{i}", nm, es.get("C"), True) + cel(f"D{i}", None, es.get("D")) + cel(f"E{i}", None, es.get("E"))
                   + cel(f"F{i}", st, es.get("F"), True) + cel(f"G{i}", None, es.get("G")) + "</row>")
    gravar(A_POS, trocar_linhas(x, out, "G"))

    # ---- Variação Semanal: nova coluna ----
    x = ler(A_VAR)
    tot = defaultdict(float)
    for r in base: tot[r["Nome Empresa"]] += float(r["Total"] or 0)
    ultima = max(re.findall(r'<c r="([A-Z]+)2"', x), key=lambda c: openpyxl.utils.column_index_from_string(c))
    nova = openpyxl.utils.get_column_letter(openpyxl.utils.column_index_from_string(ultima) + 1)
    for linha_data, emp in ((2, "TERRAZZA FEIRA SPE LTDA"), (7, "TERRAZZA BALENA SPE LTDA"), (12, "CAROZZO WELLNESS JARDIM ARMAÇÃO LTDA")):
        for ln, val, s in ((linha_data, serial(data), "123"), (linha_data + 1, round(tot.get(emp, 0) / 1000), "124")):
            m = re.search(r'(<row [^>]*r="%d"[^>]*>)(.*?)(</row>)' % ln, x, re.S)
            stl = (re.findall(r'<c r="%s%d"[^>]*?s="(\d+)"' % (ultima, ln), m.group(2)) or [s])[0]
            corpo = re.sub(r'<c r="%s%d"[^>]*?(/>|>.*?</c>)' % (nova, ln), "", m.group(2))
            x = x[:m.start()] + m.group(1) + corpo + cel(f"{nova}{ln}", val, stl) + m.group(3) + x[m.end():]
    x = re.sub(r'(<dimension ref="A1:)[A-Z]+(\d+)', lambda m: m.group(1) + nova + m.group(2), x)
    x = re.sub(r'(<row [^>]*spans="1:)(\d+)"', lambda m: m.group(1) + str(max(int(m.group(2)), openpyxl.utils.column_index_from_string(nova))) + '"', x)
    gravar(A_VAR, x)
    for ch in (tmp / "xl/charts").glob("chart*.xml"):
        c = ch.read_text(encoding="utf-8")
        c2 = re.sub(r"(Variação Semanal'!\$A\$\d+:\$)%s(\$\d+)" % ultima, lambda m: m.group(1) + nova + m.group(2), c)
        if c2 != c:
            c2 = re.sub(r"<c:(numCache|strCache)>.*?</c:\1>", "", c2, flags=re.S)   # cache velho: o Excel refaz
            ch.write_text(c2, encoding="utf-8")

    # ---- intervalos das bases nas fórmulas das outras abas ----
    def ajusta(xml, aba, n):
        nome_xml = re.escape("'" + escape(aba).replace("'", "''") + "'!")
        return re.sub("(" + nome_xml + r"\$?[A-Z]+\$?2:\$?[A-Z]+\$?)(\d+)",
                      lambda m: m.group(0) if int(m.group(2)) > 5000 else m.group(1) + str(n + 1), xml)
    for a, p in alvo.items():
        if a in (A_ATUAL, A_ANT): continue
        x = p.read_text(encoding="utf-8"); x2 = ajusta(ajusta(x, A_ATUAL, n_new), A_ANT, n_old)
        if x2 != x: p.write_text(x2, encoding="utf-8")
    wbx = re.sub(r"(_xlnm\._FilterDatabase[^>]*>'?" + re.escape(escape(A_ATUAL)) + r"'?!\$A\$1:\$[A-Z]+\$)\d+", lambda m: m.group(1) + str(n_new + 1), wbx)
    wbx = re.sub(r"(_xlnm\._FilterDatabase[^>]*>'?" + re.escape(escape(A_ANT)) + r"'?!\$A\$1:\$[A-Z]+\$)\d+", lambda m: m.group(1) + str(n_old + 1), wbx)

    # ---- recálculo completo ao abrir; sem calcChain ----
    wbx = re.sub(r"<calcPr([^>]*?)/>", lambda m: "<calcPr" + re.sub(r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)) + ' fullCalcOnLoad="1"/>', wbx)
    (tmp / "xl/workbook.xml").write_text(wbx, encoding="utf-8")
    cc = tmp / "xl/calcChain.xml"
    if cc.exists():
        cc.unlink()
        ct = (tmp / "[Content_Types].xml").read_text(encoding="utf-8")
        (tmp / "[Content_Types].xml").write_text(re.sub(r'<Override [^>]*calcChain[^>]*/>', "", ct), encoding="utf-8")
        r = (tmp / "xl/_rels/workbook.xml.rels").read_text(encoding="utf-8")
        (tmp / "xl/_rels/workbook.xml.rels").write_text(re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "", r), encoding="utf-8")

    with zipfile.ZipFile(f_ant) as zorig:
        ordem = [i.filename for i in zorig.infolist()]
    destino = saida.with_suffix(".novo.xlsx")
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for nome_arq in ordem:
            p = tmp / nome_arq
            if p.exists(): z.write(p, nome_arq)
    shutil.rmtree(tmp)
    destino.replace(saida)
    print(f"OK: {saida.name} | Base Atual {n_new} títulos R$ {sum(float(r['Total'] or 0) for r in base):,.2f} | "
          f"Base Anterior {n_old} | Posição {len(vistos)} clientes ({'com' if status else 'sem'} PDF de cobrança) | "
          f"Variação Semanal coluna {nova}: " + ", ".join(f"{NOMES_EMP.get(k, k)} {v/1000:,.0f}" for k, v in tot.items()))


if __name__ == "__main__":
    main()
