"""Funções comuns do painel de inadimplência (usadas pelo gerar_dados, importar_planilha e serviço local)."""
import json, sys, pathlib, re, datetime, os, tempfile

AQUI = pathlib.Path(__file__).resolve().parent
RAIZ = AQUI.parent


def arq_config():
    return pathlib.Path(sys.argv[sys.argv.index("--config") + 1]) if "--config" in sys.argv else AQUI / "config.json"


def cfg():
    return json.loads(arq_config().read_text(encoding="utf-8-sig"))


def caminho(p):
    """Caminho do config.json: absoluto ou relativo à pasta do painel."""
    p = pathlib.Path(os.path.expandvars(str(p)))
    return p if p.is_absolute() else RAIZ / p


def pasta_dados(c=None):
    p = caminho((c or cfg()).get("pasta_dados", "dados"))
    (p / "retratos").mkdir(parents=True, exist_ok=True)
    return p


def ler_json(arq, padrao):
    try:
        return json.loads(pathlib.Path(arq).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return padrao


def gravar_json(arq, dados):
    """Grava de forma atômica (não corrompe o arquivo se o PC desligar no meio)."""
    arq = pathlib.Path(arq)
    arq.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=arq.parent, suffix=".tmp")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=1)
    os.replace(tmp, arq)


def parc_texto(v):
    """Parcela sempre como texto. Corrige parcelas que o Excel transformou em data
    ("3/5" virou 05/03/2026 = 46145; "11/35" virou nov/1935 = 13089)."""
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        d = v
    elif isinstance(v, (int, float)) and not isinstance(v, bool):
        d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
    else:
        return str(v).strip()
    if d.year < 2000:                 # "11/35" -> 01/11/1935: mês/ano
        return f"{d.month}/{d.year % 100}"
    return f"{d.day}/{d.month}"       # "3/5" -> 03/05/2026: dia/mês


def chave(documento, parc):
    """Chave do título = Documento + Parcela, sem as marcas (C) cobrança escritural, (S) SPC/Serasa e (*) do Sienge.
    Assim a mesma parcela é reconhecida mesmo se ganhar ou perder a marca de uma semana para outra."""
    doc = re.sub(r"\s+", "", str(documento or "")).upper().replace("/", ".")
    p = re.sub(r"[^0-9/]", "", parc_texto(parc))
    return f"{doc}|{p}"


def num_br(s):
    """'1.234,56' -> 1234.56"""
    return float(str(s).replace(".", "").replace(",", "."))


def dia_iso(d):
    return d.isoformat() if d else None
