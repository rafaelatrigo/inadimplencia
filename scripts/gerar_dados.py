"""
Gera os dados do painel de inadimplência (site/data/painel.enc.json, criptografado).

De onde vem cada número (pastas em scripts/config.json):
- BASE: o CSV do Sienge "Contas a receber" mais recente em pasta_entrada
  (parâmetros: dez/25 até D-0, flags Parcelas SPC/Serasa, Inadimplentes, Sub judice e Parcelas de resíduo,
  plano financeiro 1.01.01.02). Só títulos já vencidos (vencimento até a data do relatório).
  - Tira os títulos classificados nos Tratados (dados/tratados.json).
  - Regra dos 10 dias (regra_dias): o cliente entra se tiver ALGUM título com mais de 10 dias de atraso;
    aí entram todos os títulos vencidos dele. Cliente só com títulos de até 10 dias fica de fora.
  - Dias ajustado = maior atraso do cliente (é o que define a faixa do aging).
- STATUS DA COBRANÇA e SUB JUDICE: o PDF "Títulos em Cobrança" mais recente em pasta_entrada
  (flags Em cobrança, Inadimplentes, Sub judice; situação Todas). Último lançamento de cada cliente.
- RECEBIMENTO ESPERADO: os PDFs "Contas a Receber (por Data de Vencimento)" enviados ao BPO, em pasta_boletagem
  (um por empresa e mês, coluna Total, "Total da empresa"). Clientes em esperado_excluir_clientes saem do total.
  Meses antigos sem PDF vêm de dados/esperado_manual.json.
- SEMANA ANTERIOR: o retrato salvo na semana anterior (dados/retratos/AAAA-MM-DD.json).

Opções:  --se-mudou   só gera se algum arquivo de entrada, os Tratados ou os fechamentos mudaram
"""
import sys, json, csv, re, datetime, pathlib, os, base64, hashlib, getpass, unicodedata
from collections import defaultdict
from comum import AQUI, RAIZ, cfg as ler_cfg, caminho, pasta_dados, ler_json, gravar_json, parc_texto, chave, num_br

try:
    import pdfplumber
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
except ImportError:
    sys.exit("Faltam bibliotecas. Rode:  python -m pip install pdfplumber cryptography")

cfg = ler_cfg()
DADOS = pasta_dados(cfg)
ENTRADA = caminho(cfg.get("pasta_entrada", "entrada"))
BOLETAGEM = caminho(cfg.get("pasta_boletagem", "entrada/boletagem"))
REGRA = int(cfg.get("regra_dias", 10))
avisos = []


def sem_espaco(s):
    return re.sub(r"\s+", "", s or "")


def mais_recente(arqs):
    return max(arqs, key=lambda f: f.stat().st_mtime) if arqs else None


def listar(pasta, padrao, recursivo=False):
    if not pasta.is_dir():
        return []
    it = pasta.rglob(padrao) if recursivo else pasta.glob(padrao)
    return [f for f in it if f.is_file() and not f.name.startswith("~$")]


def texto_pdf(arq, so_primeira=False):
    with pdfplumber.open(arq) as pdf:
        pags = pdf.pages[:1] if so_primeira else pdf.pages
        return "\n".join((p.extract_text(layout=True) or "") for p in pags)


def data_br(s):
    return datetime.datetime.strptime(s, "%d/%m/%Y").date()


# ============ 1. Contas a receber (CSV do Sienge) ============
def achar_csv():
    cand = []
    for f in listar(ENTRADA, cfg.get("padrao_contas_receber", "*.csv")):
        try:
            cab = f.read_text(encoding="utf-8-sig", errors="replace").split("\n", 1)[0]
        except OSError:
            continue
        if "Data vecto" in cab and "Nome Empresa" in cab:
            cand.append(f)
    f = mais_recente(cand)
    if not f:
        sys.exit(f"Não achei o CSV do Contas a Receber em {ENTRADA}\n(cabeçalho com 'Data vecto' e 'Nome Empresa').")
    return f


def data_do_arquivo(f):
    m = re.search(r"(20\d\d)-(\d\d)-(\d\d)", f.name)
    if m:
        return datetime.date(int(m[1]), int(m[2]), int(m[3]))
    return datetime.date.fromtimestamp(f.stat().st_mtime)


def ler_csv(f):
    linhas = []
    with open(f, encoding="utf-8-sig", newline="") as fh:
        for r in csv.DictReader(fh):
            if not r.get("Data vecto"):
                continue
            n = lambda k: float(str(r.get(k) or "0").replace(",", ".") or 0)
            linhas.append({
                "chave": chave(r["Documento"], r["Parc"]),
                "venc": r["Data vecto"][:10], "cliente": (r["Cliente"] or "").strip(),
                "cd_cliente": str(r["Cd Cliente"]).strip(), "documento": r["Documento"].strip(),
                "titulo": str(r.get("Título") or r.get("Titulo") or "").strip(), "parc": parc_texto(r["Parc"]),
                "tc": (r["TC"] or "").strip(), "unid": str(r["Unid. princ"] or "").strip(),
                "valor_original": n("Valor original"), "saldo": n("Saldo atual"), "dias": int(float(r["Dias"] or 0)),
                "acrescimo": n("Acréscimo"), "total": n("Total"),
                "empresa": (r["Nome Empresa"] or "").strip(), "cd_empresa": str(r["Cd Empresa"]).strip()})
    return linhas


# ============ 2. Títulos em cobrança (PDF) ============
RE_TIT = re.compile(r"Título\s*(\d+)\s+Documento\s*(.+?)\s*$")
RE_CLI = re.compile(r"Cliente\s*(\d+)\s*-\s*(.+?)\s+Saldo\s*Inadimplente")
RE_SIT = re.compile(r"Situação\s*(\S.*?)\s*$")
RE_LANC = re.compile(r"^\s*(\d\d/\d\d/\d{4})\s+(\S{1,3})\s+(?:(OK)\s+)?(.*\S)\s*$")
RE_RODAPE = re.compile(r"^\s*(\d\d/\d\d/\d{4})\s*-\s*\d\d:\d\d")
CABECALHOS = ("Títulos em", "Empresa", "Empreendimento", "Unidades", "Processo", "Foro", "Data ")


def achar_cobranca():
    cand = []
    for f in listar(ENTRADA, cfg.get("padrao_cobranca", "*.pdf")):
        try:
            if "TítulosemCobrança" in sem_espaco(texto_pdf(f, True)):
                cand.append(f)
        except Exception:
            continue
    return mais_recente(cand)


def ler_cobranca(f):
    txt = texto_pdf(f)
    m = RE_RODAPE.search(txt) or re.search(r"(\d\d/\d\d/\d{4})\s*-\s*\d\d:\d\d", txt)
    data_rel = data_br(m.group(1)) if m else datetime.date.fromtimestamp(f.stat().st_mtime)
    doc = cli = None
    situacao, lanc = {}, defaultdict(list)
    ultimo = None
    for ln in txt.split("\n"):
        if not ln.strip():
            continue
        if RE_RODAPE.match(ln):
            ultimo = None; continue
        mt = RE_TIT.search(ln)
        if mt:
            doc = mt.group(2); ultimo = None; continue
        mc = RE_CLI.search(ln)
        if mc:
            cli = (mc.group(1), mc.group(2).strip()); ultimo = None; continue
        ms = RE_SIT.search(ln)
        if ms and doc:
            situacao[chave(doc, "").split("|")[0]] = ms.group(1).strip()
        ml = RE_LANC.match(ln)
        if ml and cli:
            ultimo = {"data": data_br(ml.group(1)), "texto": ml.group(4).strip(), "doc": doc, "nome": cli[1]}
            lanc[cli[0]].append(ultimo); continue
        if any(ln.strip().startswith(c) or sem_espaco(ln).startswith(sem_espaco(c)) for c in CABECALHOS):
            ultimo = None; continue
        if ultimo and not ms:                       # continuação da descrição na linha de baixo
            ultimo["texto"] += " " + ln.strip()
    status, datas_estranhas = {}, []
    limite_antigo = datetime.date(data_rel.year - 1, 1, 1)
    for cd, ls in lanc.items():
        vistos = set()
        for l in ls:
            k = (l["data"], l["texto"])
            if (l["data"] > data_rel or l["data"] < limite_antigo) and k not in vistos:
                vistos.add(k)
                datas_estranhas.append({"cd_cliente": cd, "cliente": l["nome"], "documento": l["doc"],
                                        "data": l["data"].isoformat(), "texto": l["texto"][:160]})
        validos = [l for l in ls if l["data"] <= data_rel]
        if validos:
            u = max(enumerate(validos), key=lambda x: (x[1]["data"], x[0]))[1]   # data mais recente; empate = o último lançado
            status[cd] = {"data": u["data"].isoformat(), "texto": u["texto"], "n": len(ls)}
    subjudice = sorted(d for d, s in situacao.items() if sem_espaco(s).lower().startswith("subjudice"))
    return {"arquivo": f.name, "data": data_rel.isoformat(), "status": status, "subjudice_docs": subjudice,
            "datas_estranhas": datas_estranhas}


# ============ 3. Recebimento esperado (PDFs da boletagem) ============
RE_EMP = re.compile(r"Empresa\s*(\d+)\s*-\s*(.+?)\s{2,}Período\s*(\d\d/\d\d/\d{4})\s*a\s*(\d\d/\d\d/\d{4})")
RE_NUM = re.compile(r"\d{1,3}(?:\.\d{3})*,\d\d")
RE_LINHA = re.compile(r"^\s*(\d\d/\d\d/\d{4})\s+\S")


def ler_boletagem(f, excluir):
    txt = texto_pdf(f)
    m = RE_EMP.search(txt)
    if not m:
        return None
    empresa = m.group(2).strip()
    ini, fim = data_br(m.group(3)), data_br(m.group(4))
    total = None
    for chave_tot in ("Total da empresa", "Total geral"):
        for ln in txt.split("\n"):
            if sem_espaco(ln).startswith(sem_espaco(chave_tot)) or chave_tot in ln:
                nums = RE_NUM.findall(ln)
                if nums:
                    total = num_br(nums[-1])
        if total is not None:
            break
    if total is None:
        return None
    # clientes excluídos (ex.: JDL no Feira): soma o Total das linhas deles e tira do total
    codigos = {str(x["codigo_cliente"]) for x in excluir if x["empresa"].upper() == empresa.upper()}
    excluido, linhas = 0.0, txt.split("\n")
    if codigos:
        for i, ln in enumerate(linhas):
            if not RE_LINHA.match(ln):
                continue
            bloco = ln + " " + (linhas[i + 1] if i + 1 < len(linhas) and not RE_LINHA.match(linhas[i + 1]) else "")
            mc = re.search(r"\(\s*(\d+)\s*\)", re.sub(r"\(\s*([^)]*?)\s*\)", lambda x: "(" + sem_espaco(x.group(1)) + ")", bloco))
            if mc and mc.group(1) in codigos:
                nums = RE_NUM.findall(ln)
                if nums:
                    excluido += num_br(nums[-1])
    return {"empresa": empresa, "mes": ini.strftime("%Y-%m"), "periodo": [ini.isoformat(), fim.isoformat()],
            "total_pdf": round(total, 2), "excluido": round(excluido, 2), "valor": round(total - excluido, 2), "arquivo": f.name}


def esperado_por_mes():
    manual = ler_json(DADOS / "esperado_manual.json", {})     # {empresa: {AAAA-MM: valor}}
    cache = ler_json(DADOS / "cache_boletagem.json", {})
    excluir = cfg.get("esperado_excluir_clientes", [])
    sinal_excl = json.dumps(excluir, sort_keys=True)
    lidos, novo_cache = [], {}
    for f in listar(BOLETAGEM, cfg.get("padrao_boletagem", "*.pdf"), recursivo=True):
        k = f"{f}|{f.stat().st_mtime}|{sinal_excl}"
        if k in cache:
            r = cache[k]
        else:
            try:
                r = ler_boletagem(f, excluir) if "ContasaReceber(porDatadeVencimento)" in sem_espaco(texto_pdf(f, True)) else None
            except Exception as e:
                avisos.append(f"Não consegui ler {f.name}: {e}"); r = None
        novo_cache[k] = r
        if r:
            lidos.append((f.stat().st_mtime, r))
    gravar_json(DADOS / "cache_boletagem.json", novo_cache)
    esp = {e: {m: {"valor": v, "fonte": "histórico (planilha)"} for m, v in ms.items()} for e, ms in manual.items()}
    for _, r in sorted(lidos, key=lambda x: x[0]):            # o PDF mais recente de cada empresa/mês vale
        obs = f"{r['arquivo']}" + (f" (menos R$ {r['excluido']:,.2f} de clientes excluídos)".replace(",", "X").replace(".", ",").replace("X", ".") if r["excluido"] else "")
        esp.setdefault(r["empresa"], {})[r["mes"]] = {"valor": r["valor"], "fonte": obs}
    return esp


# ============ 4. Monta a base ============
# ============ 4. Planilha de inadimplência (fonte principal, quando configurada) ============
def achar_planilha():
    pasta = caminho(cfg["pasta_planilha"]) if cfg.get("pasta_planilha") else None
    if not pasta or not pasta.is_dir():
        if pasta: avisos.append(f"Pasta da planilha não encontrada: {pasta}. Usei os relatórios do Sienge.")
        return None
    cand = listar(pasta, cfg.get("padrao_planilha", "Inadimpl*Carozzo*.xlsx"))
    def data_nome(f):
        m = re.search(r"(\d{2})(\d{2})(20\d{2})", f.stem)
        try: return datetime.date(int(m[3]), int(m[2]), int(m[1])) if m else None
        except ValueError: return None
    datados = [(data_nome(f), f.stat().st_mtime, f) for f in cand if data_nome(f)]
    if not datados:
        return None
    datados.sort(key=lambda x: (x[0], x[1]))
    d, _, f = datados[-1]                       # a data do nome do arquivo manda; empate = a salva por último
    anteriores = sorted({x[0] for x in datados if x[0] < d})
    return f, d, (anteriores[-1] if anteriores else None)


def _f(v):
    try: return float(str(v).replace(",", ".")) if v not in (None, "") else 0.0
    except ValueError: return 0.0


def _dia(v):
    if isinstance(v, datetime.datetime): return v.date().isoformat()
    if isinstance(v, datetime.date): return v.isoformat()
    try: return data_br(str(v)[:10]).isoformat()
    except ValueError: return str(v)[:10] if v else None


def _titulo(r):
    """Linha das abas Base Atual / Base Anterior / Tratados (colunas A..Y) -> título."""
    return {"chave": chave(r[3], r[5]), "venc": _dia(r[0]), "cliente": str(r[1] or "").strip(), "cd_cliente": str(r[2]).strip(),
            "documento": str(r[3]).strip(), "titulo": str(r[4]).strip(), "parc": parc_texto(r[5]), "tc": str(r[6] or "").strip(),
            "unid": str(r[7] or "").strip(), "valor_original": _f(r[8]), "total": _f(r[22]), "dias": int(_f(r[14])),
            "empresa": str(r[24] or "").strip(), "cd_empresa": str(r[23]).strip()}


def copia_planilha(f):
    """Copia a planilha para a pasta de dados antes de ler. Se ela estiver aberta no Excel ou sincronizando no OneDrive
    (PermissionError), tenta de novo algumas vezes; se não der, usa a última cópia boa e avisa no painel."""
    import shutil, time
    destino = DADOS / "planilha_ultima_lida.xlsx"
    for tentativa in range(4):
        try:
            shutil.copyfile(f, destino)
            gravar_json(DADOS / "planilha_ultima_lida.json", {"arquivo": f.name, "lida_em": datetime.datetime.now().isoformat(timespec="seconds")})
            (DADOS / ".planilha_bloqueada").unlink(missing_ok=True)
            return destino
        except PermissionError:
            time.sleep(3)
    (DADOS / ".planilha_bloqueada").write_text(f.name)     # o serviço tenta de novo a cada 3 minutos até conseguir ler
    info = ler_json(DADOS / "planilha_ultima_lida.json", None)
    if destino.exists() and info:
        avisos.append(f"A planilha {f.name} estava bloqueada (aberta no Excel ou sincronizando no OneDrive). "
                      f"Usei a última cópia lida ({info['arquivo']}, {info['lida_em'][:16].replace('T', ' ')}). Feche a planilha e clique em Publicar agora.")
        return destino
    avisos.append(f"Não consegui abrir {f.name} (aberta no Excel ou sincronizando no OneDrive): só a visão automática (Sienge). Feche a planilha e publique de novo.")
    return None


def ler_planilha(f):
    import openpyxl, warnings
    f = copia_planilha(f)
    if not f:
        return None
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
    abas = {ws.title.strip(): ws for ws in wb.worksheets}
    aba = lambda trecho: next((ws for t, ws in abas.items() if trecho.lower() in t.lower()), None)
    def linhas(ws):
        return [r for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0] and len(r) > 24 and r[3]] if ws else []
    sim = lambda v: str(v or "").strip().lower() == "sim"
    atual = [dict(_titulo(r), subj_planilha=sim(r[26] if len(r) > 26 else None)) for r in linhas(aba("Base Atual"))]
    anterior = [dict(_titulo(r), subj_planilha=sim(r[26] if len(r) > 26 else None)) for r in linhas(aba("Base Anterior"))]
    trat = []
    for r in linhas(aba("Tratados")):
        t = _titulo(r)
        t.update(status=str(r[27]).strip() if len(r) > 27 and r[27] else "", obs=str(r[28]).strip() if len(r) > 28 and r[28] else "",
                 tratado_em=None, origem="planilha")
        trat.append(t)
    # Base Posição Clientes: última descrição de cada cliente ("27/10/2026 - TEXTO")
    posicao = {}
    ws = aba("Posição Clientes") or aba("Posicao Clientes")
    if ws:
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or len(r) < 6 or not r[2] or not r[5]: continue
            m = re.match(r"\s*(\d\d/\d\d/\d{4})\s*-\s*(.*)", str(r[5]))
            if m:
                posicao[str(r[2]).strip().upper()] = {"data": data_br(m.group(1)).isoformat(), "texto": m.group(2).strip()}
    # Variação Semanal (R$ mil)
    sem = defaultdict(dict)
    ws = aba("Variação Semanal")
    if ws:
        grade = [list(r) for r in ws.iter_rows(values_only=True)]
        rot = {"Ícone": "TERRAZZA FEIRA SPE LTDA", "Balena": "TERRAZZA BALENA SPE LTDA", "Carozzo Wellness": "CAROZZO WELLNESS JARDIM ARMAÇÃO LTDA"}
        for i, r in enumerate(grade[:40]):
            nome = r[0] if r else None
            if nome not in rot: continue
            rd = next((x for x in range(i + 1, min(i + 4, len(grade))) if any(isinstance(v, datetime.datetime) for v in grade[x])), None)
            if rd is None or rd + 1 >= len(grade): continue
            for c, d in enumerate(grade[rd]):
                v = grade[rd + 1][c] if c < len(grade[rd + 1]) else None
                if isinstance(d, datetime.datetime) and v not in (None, ""):
                    sem[d.date().isoformat()][rot[nome]] = round(_f(v) * 1000, 2)
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
    wb.close()
    return {"atual": atual, "anterior": anterior, "tratados": trat, "posicao": posicao, "semanal": semanal}


# ============ 5. Monta as duas visões: automática (Sienge) e planilha ============
def dias_ajustado(lst):
    maior = defaultdict(int)
    for r in lst: maior[r["cd_cliente"]] = max(maior[r["cd_cliente"]], r["dias"])
    for r in lst: r["dias_aj"] = maior[r["cd_cliente"]]


def enxuto(r):
    return {k: r[k] for k in ("chave", "venc", "cliente", "cd_cliente", "documento", "titulo", "parc", "tc", "unid",
                              "valor_original", "total", "dias", "dias_aj", "empresa", "subjudice") if k in r}


def com_semana(hist, data, base):
    tot = defaultdict(float)
    for r in base: tot[r["empresa"]] += r["total"]
    h = [x for x in hist if x["data"] != data] + [{"data": data, "valores": {k: round(v, 2) for k, v in tot.items()}}]
    return sorted(h, key=lambda x: x["data"])


def desambiguar(lst):
    """O Sienge às vezes repete Documento + Parcela em títulos diferentes (ex.: duas parcelas "1/1" do mesmo contrato com
    vencimentos diferentes). Nesses casos a chave ganha o vencimento, para um não sumir no lugar do outro."""
    for t in lst: t["chave"] = t["chave"].split("#")[0]
    n = defaultdict(int)
    for t in lst: n[t["chave"]] += 1
    for t in lst:
        if n[t["chave"]] > 1: t["chave"] += "#" + str(t.get("venc"))
    return lst


def sem_ignorados(lst):
    """Títulos em config.json > ignorar_titulos saem de tudo (base, semana anterior, Tratados), como se nunca tivessem existido."""
    ign = {chave(x["documento"], p) for x in cfg.get("ignorar_titulos", []) for p in x.get("parcelas", [])}
    return [t for t in lst if t["chave"].split("#")[0] not in ign] if ign else lst


def montar():
    doc = lambda t: chave(t["documento"], "").split("|")[0]
    usa_plan = bool(cfg.get("pasta_planilha"))
    trat_painel = sem_ignorados([t for t in ler_json(DADOS / "tratados.json", []) if not (usa_plan and t.get("origem") == "planilha")])
    cob, f_cob = None, achar_cobranca()
    if f_cob: cob = ler_cobranca(f_cob)
    else: avisos.append("Não achei o PDF de Títulos em Cobrança na pasta de entrada.")
    if not cob: cob = {"arquivo": None, "data": None, "status": {}, "subjudice_docs": [], "datas_estranhas": []}
    try: f_csv = achar_csv()
    except SystemExit: f_csv = None
    sienge = desambiguar(sem_ignorados(ler_csv(f_csv))) if f_csv else []
    plan = achar_planilha() if usa_plan else None
    P = ler_planilha(plan[0]) if plan else None
    if plan and not P: plan = None
    if P:
        for k in ("atual", "anterior", "tratados"): P[k] = desambiguar(sem_ignorados(P[k]))
    if usa_plan and not plan: avisos.append("Não achei a planilha de inadimplência na pasta configurada: só a visão automática (Sienge).")
    if not f_csv and not P: sys.exit("Não achei o CSV do Contas a Receber nem a planilha de inadimplência.")

    # ---- Tratados: aba Tratados da planilha + classificados pelo painel (o painel vale por cima) ----
    por_chave = {t["chave"]: t for t in (P["tratados"] if P else [])}
    for t in trat_painel: por_chave[t["chave"]] = t
    tratados = list(por_chave.values())
    k_trat, k_painel = set(por_chave), {t["chave"] for t in trat_painel}

    # ---- Sub judice: PDF de cobrança OU marcado "Sim" na planilha (vale para as duas visões e para a semana anterior) ----
    subj = set(cob["subjudice_docs"])
    if P: subj |= {doc(t) for t in P["atual"] + P["anterior"] if t.get("subj_planilha")}
    if not cfg.get("considerar_subjudice", False):
        subj = set()                 # decisão: sub judice não é evidenciado; entra normalmente nas faixas de atraso

    # ---- Status da cobrança: PDF e Base Posição Clientes; vale o lançamento mais recente ----
    status = dict(cob["status"])
    if P:
        cd_por_nome = {t["cliente"].upper(): t["cd_cliente"] for t in P["atual"] + P["anterior"] + tratados + sienge}
        for nome, st in P["posicao"].items():
            cd = cd_por_nome.get(nome)
            if cd and (cd not in status or st["data"] >= status[cd]["data"]): status[cd] = st
    cob["status"] = status

    visoes, fora_regra = {}, []
    hist_plan = P["semanal"] if P else []

    # ---- visão automática: CSV do Sienge + regra dos 10 dias, sem os Tratados ----
    if f_csv:
        data_base = data_do_arquivo(f_csv)
        cand = [r for r in sienge if r["venc"] <= data_base.isoformat() and r["chave"] not in k_trat]
        maior = defaultdict(int)
        for r in cand: maior[r["cd_cliente"]] = max(maior[r["cd_cliente"]], r["dias"])
        base = [dict(r, subjudice=doc(r) in subj) for r in cand if maior[r["cd_cliente"]] > REGRA]
        fora_regra = sorted({r["cliente"] for r in cand if maior[r["cd_cliente"]] <= REGRA})
        anteriores = sorted(p for p in (DADOS / "retratos").glob("*.json") if p.stem < data_base.isoformat())
        ant = ler_json(anteriores[-1], {"data": None, "titulos": []}) if anteriores else {"data": None, "titulos": []}
        ant["titulos"] = desambiguar(sem_ignorados(ant["titulos"]))
        for t in ant["titulos"]: t["subjudice"] = bool(t.get("subjudice")) or doc(t) in subj
        dias_ajustado(base); dias_ajustado(ant["titulos"])
        base.sort(key=lambda r: (r["empresa"], r["cliente"], r["venc"]))
        gravar_json(DADOS / "retratos" / f"{data_base.isoformat()}.json", {"data": data_base.isoformat(), "titulos": base})
        gravar_json(DADOS / "sienge_atual.json", {"data": data_base.isoformat(), "titulos": {r["chave"]: r for r in sienge}})
        hist = hist_plan or ler_json(DADOS / "semanal.json", [])
        semanal = com_semana(hist, data_base.isoformat(), base)
        gravar_json(DADOS / "semanal.json", semanal)
        visoes["sienge"] = {"rotulo": "Automática (Sienge)", "fonte": f"{f_csv.name} · Contas a Receber do Sienge, regra dos {REGRA} dias",
                            "data_base": data_base.isoformat(), "data_anterior": ant["data"],
                            "titulos": [enxuto(r) for r in base], "anterior": [enxuto(r) for r in ant["titulos"]], "semanal": semanal}

    # ---- visão planilha: abas Base Atual e Base Anterior da planilha mais recente ----
    if P:
        f_plan, d_plan, d_ant = plan
        base = [dict(t, subjudice=doc(t) in subj) for t in P["atual"] if t["chave"] not in k_painel]
        ant_t = [dict(t, subjudice=doc(t) in subj) for t in P["anterior"]]
        dias_ajustado(base); dias_ajustado(ant_t)
        base.sort(key=lambda r: (r["empresa"], r["cliente"], r["venc"]))
        visoes["planilha"] = {"rotulo": "Planilha", "fonte": f"{f_plan.name} · abas Base Atual e Base Anterior",
                              "data_base": d_plan.isoformat(), "data_anterior": d_ant.isoformat() if d_ant else None,
                              "titulos": [enxuto(r) for r in base], "anterior": [enxuto(r) for r in ant_t],
                              "semanal": com_semana(hist_plan, d_plan.isoformat(), base)}

    # Tratados: ainda no Contas a Receber = aguardando compensação; saiu = compensado (sem o CSV, fica aguardando)
    no_sienge = {r["chave"] for r in sienge}
    caixa = {x["status"]: x["caixa"] for x in cfg["status_tratados"]}
    trat_out = [dict(t, caixa=caixa.get(t.get("status") or "", "Sem status"), no_sienge=(t["chave"] in no_sienge) if sienge else True)
                for t in tratados]

    todos = [t for v in visoes.values() for t in v["titulos"] + v["anterior"]]
    emps_cfg = cfg["empreendimentos"]
    extras = sorted({r["empresa"] for r in todos} - {e["empresa"] for e in emps_cfg})
    empreendimentos = [{"nome": e["nome"], "empresa": e["empresa"]} for e in emps_cfg] + [{"nome": x.title(), "empresa": x} for x in extras]
    padrao = "sienge" if "sienge" in visoes else "planilha"
    v0 = visoes[padrao]
    return {
        "gerado_em": datetime.datetime.now().isoformat(timespec="seconds"),
        "visoes": visoes, "visao_padrao": padrao,
        # campos da visão padrão (compatibilidade e resumo no console)
        "data_base": v0["data_base"], "data_anterior": v0["data_anterior"], "titulos": v0["titulos"], "anterior": v0["anterior"],
        "semanal": v0["semanal"], "regra_dias": REGRA,
        "empreendimentos": empreendimentos,
        "tratados": trat_out,
        "status_tratados": cfg["status_tratados"],
        "cobranca": {"arquivo": cob["arquivo"], "data": cob["data"], "status": cob["status"],
                     "subjudice_docs": sorted(subj), "datas_estranhas": cob["datas_estranhas"]},
        "esperado": esperado_por_mes(),
        "fechamentos": ler_json(DADOS / "fechamentos.json", []),
        "fora_regra": fora_regra,
        "fontes": {"contas_receber": v0["fonte"], "cobranca": f_cob.name if f_cob else None},
        "avisos": avisos,
    }


# ============ 6. Criptografia ============
def ler_senha():
    a = AQUI / "senha.txt"
    s = a.read_text(encoding="utf-8-sig").strip() if a.exists() else ""
    if not s:
        s = getpass.getpass("Senha de acesso ao painel (não aparece ao digitar): ").strip()
    if len(s) < 12:
        sys.exit("A senha (scripts/senha.txt) precisa ter pelo menos 12 caracteres.")
    return s


def assinatura():
    pl = caminho(cfg["pasta_planilha"]) if cfg.get("pasta_planilha") else None
    arqs = listar(ENTRADA, "*.*") + listar(BOLETAGEM, "*.pdf", True) + (listar(pl, cfg.get("padrao_planilha", "Inadimpl*Carozzo*.xlsx")) if pl else []) + \
        [DADOS / n for n in ("tratados.json", "fechamentos.json", "esperado_manual.json") if (DADOS / n).exists()] + [AQUI / "config.json"]
    return hashlib.sha256("|".join(sorted(f"{f}:{f.stat().st_mtime}" for f in arqs)).encode()).hexdigest()


if __name__ == "__main__":
    arq_assin = DADOS / ".assinatura"
    assin = assinatura()
    if "--se-mudou" in sys.argv and arq_assin.exists() and arq_assin.read_text() == assin:
        print("Nada a fazer: as entradas não mudaram desde a última geração."); sys.exit(0)
    dados = montar()
    senha = ler_senha()
    salt, iv = os.urandom(16), os.urandom(12)
    ITER = 600_000
    k = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(senha.encode("utf-8"))
    ct = AESGCM(k).encrypt(iv, json.dumps(dados, ensure_ascii=False).encode("utf-8"), None)
    b64 = lambda b: base64.b64encode(b).decode("ascii")
    destino = RAIZ / "site" / "data"; destino.mkdir(parents=True, exist_ok=True)
    (destino / "painel.enc.json").write_text(json.dumps({"v": 1, "kdf": "PBKDF2-SHA256", "iter": ITER, "salt": b64(salt),
                                                        "iv": b64(iv), "dados": b64(ct)}), encoding="utf-8")
    json.dump({"gerado_em": dados["gerado_em"]}, open(destino / "versao.json", "w"))
    arq_assin.write_text(assin)

    # resumo no console
    print(f"Base: {dados['fontes']['contas_receber']} - data {dados['data_base']}; semana anterior: {dados['data_anterior'] or '-'}")
    for k, v in dados["visoes"].items():
        print(f"  visão {v['rotulo']:<22} data {v['data_base']}  {len(v['titulos']):4d} títulos  R$ {sum(r['total'] for r in v['titulos']):14,.2f}  (anterior {v['data_anterior']}: R$ {sum(r['total'] for r in v['anterior']):,.2f})")
    tot = defaultdict(lambda: [0, 0.0])
    for r in dados["titulos"]:
        tot[r["empresa"]][0] += 1; tot[r["empresa"]][1] += r["total"]
    for e, (n, v) in tot.items():
        print(f"  {e:<40} {n:4d} títulos  R$ {v:14,.2f}")
    print(f"  {'TOTAL':<40} {len(dados['titulos']):4d} títulos  R$ {sum(r['total'] for r in dados['titulos']):14,.2f}")
    print(f"Tratados: {len(dados['tratados'])} · Sub judice: {len(dados['cobranca']['subjudice_docs']) if dados['cobranca'] else '-'}")
    for e, ms in dados["esperado"].items():
        print(f"  esperado {e[:30]:<30} " + "  ".join(f"{m}: {v['valor']:,.2f}" for m, v in sorted(ms.items())))
    for a in avisos:
        print("AVISO:", a)
    print("OK: site/data/painel.enc.json gerado.")
