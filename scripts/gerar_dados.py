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
def montar():
    f_csv = achar_csv()
    data_base = data_do_arquivo(f_csv)
    sienge = ler_csv(f_csv)
    tratados = ler_json(DADOS / "tratados.json", [])
    chaves_trat = {t["chave"] for t in tratados}
    cob = None
    f_cob = achar_cobranca()
    if f_cob:
        cob = ler_cobranca(f_cob)
    else:
        avisos.append("Não achei o PDF de Títulos em Cobrança: status da cobrança e Sub judice ficaram sem atualizar.")
    subj = set(cob["subjudice_docs"]) if cob else set()

    vencidos = [r for r in sienge if r["venc"] <= data_base.isoformat()]
    candidatos = [r for r in vencidos if r["chave"] not in chaves_trat]
    maior = defaultdict(int)
    for r in candidatos:
        maior[r["cd_cliente"]] = max(maior[r["cd_cliente"]], r["dias"])
    base = []
    for r in candidatos:
        if maior[r["cd_cliente"]] > REGRA:
            r = dict(r, dias_aj=maior[r["cd_cliente"]], subjudice=chave(r["documento"], "").split("|")[0] in subj)
            base.append(r)
    base.sort(key=lambda r: (r["empresa"], r["cliente"], r["venc"]))
    fora_regra = sorted({r["cliente"] for r in candidatos if maior[r["cd_cliente"]] <= REGRA})

    # retratos: o desta semana (sobrescreve se rodar de novo no mesmo dia) e o anterior
    gravar_json(DADOS / "retratos" / f"{data_base.isoformat()}.json", {"data": data_base.isoformat(), "titulos": base})
    anteriores = sorted(p for p in (DADOS / "retratos").glob("*.json") if p.stem < data_base.isoformat())
    ant = ler_json(anteriores[-1], {"data": None, "titulos": []}) if anteriores else {"data": None, "titulos": []}

    # Sienge completo (o serviço usa para registrar um título nos Tratados com os dados do dia)
    gravar_json(DADOS / "sienge_atual.json", {"data": data_base.isoformat(), "titulos": {r["chave"]: r for r in sienge}})

    # Tratados: ainda no Sienge = aguardando compensação; saiu do Sienge = compensado (ou distrato concluído)
    no_sienge = {r["chave"] for r in sienge}
    caixa = {s["status"]: s["caixa"] for s in cfg["status_tratados"]}
    trat_out = []
    for t in tratados:
        trat_out.append(dict(t, caixa=caixa.get(t.get("status") or "", "Sem status"), no_sienge=t["chave"] in no_sienge))

    # histórico semanal (R$ por empresa)
    semanal = ler_json(DADOS / "semanal.json", [])
    tot = defaultdict(float)
    for r in base:
        tot[r["empresa"]] += r["total"]
    semanal = [s for s in semanal if s["data"] != data_base.isoformat()] + [{"data": data_base.isoformat(), "valores": {k: round(v, 2) for k, v in tot.items()}}]
    semanal.sort(key=lambda s: s["data"])
    gravar_json(DADOS / "semanal.json", semanal)

    emps_cfg = cfg["empreendimentos"]
    conhecidas = {e["empresa"] for e in emps_cfg}
    extras = sorted({r["empresa"] for r in base + ant["titulos"]} - conhecidas)
    empreendimentos = [{"nome": e["nome"], "empresa": e["empresa"]} for e in emps_cfg] + [{"nome": x.title(), "empresa": x} for x in extras]

    enxuto = lambda r: {k: r[k] for k in ("chave", "venc", "cliente", "cd_cliente", "documento", "titulo", "parc", "tc", "unid",
                                         "valor_original", "total", "dias", "dias_aj", "empresa", "subjudice") if k in r}
    return {
        "gerado_em": datetime.datetime.now().isoformat(timespec="seconds"),
        "data_base": data_base.isoformat(), "data_anterior": ant["data"], "regra_dias": REGRA,
        "empreendimentos": empreendimentos,
        "titulos": [enxuto(r) for r in base],
        "anterior": [enxuto(r) for r in ant["titulos"]],
        "tratados": trat_out,
        "status_tratados": cfg["status_tratados"],
        "cobranca": ({"arquivo": cob["arquivo"], "data": cob["data"], "status": cob["status"],
                      "subjudice_docs": cob["subjudice_docs"], "datas_estranhas": cob["datas_estranhas"]} if cob else None),
        "esperado": esperado_por_mes(),
        "fechamentos": ler_json(DADOS / "fechamentos.json", []),
        "semanal": semanal,
        "fora_regra": fora_regra,
        "fontes": {"contas_receber": f_csv.name, "cobranca": f_cob.name if f_cob else None},
        "avisos": avisos,
    }


# ============ 5. Criptografia ============
def ler_senha():
    a = AQUI / "senha.txt"
    s = a.read_text(encoding="utf-8-sig").strip() if a.exists() else ""
    if not s:
        s = getpass.getpass("Senha de acesso ao painel (não aparece ao digitar): ").strip()
    if len(s) < 12:
        sys.exit("A senha (scripts/senha.txt) precisa ter pelo menos 12 caracteres.")
    return s


def assinatura():
    arqs = listar(ENTRADA, "*.*") + listar(BOLETAGEM, "*.pdf", True) + \
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
    print(f"Base: {dados['fontes']['contas_receber']} (data {dados['data_base']}); semana anterior: {dados['data_anterior'] or '-'}")
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
