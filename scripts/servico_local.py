"""
Serviço local do painel de inadimplência.

Fica rodando em segundo plano no PC da Rafaela (inicia junto com o Windows) e faz três coisas:
1. Publicação automática: a cada minuto olha a pasta de entrada (CSV do Contas a Receber, PDF de Títulos em
   Cobrança) e a pasta da boletagem. Quando um arquivo novo é salvo (e fica 1 minuto sem mudar), gera e publica.
2. Botão "Publicar agora" do painel.
3. Edição pelo painel (só neste computador): classificar títulos nos Tratados, mudar o status, desfazer,
   e registrar o fechamento mensal do % de inadimplência (o que foi enviado à Rita).

Os dados editados ficam na pasta de dados (config.json > pasta_dados), NUNCA no GitHub sem criptografia.
Só aceita chamadas deste computador (127.0.0.1) e do endereço do painel definido em config.json.
Registro em: atualizacao.log (pasta do painel).
"""
import json, sys, os, time, threading, subprocess, datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from comum import AQUI, RAIZ, cfg, caminho, pasta_dados, ler_json, gravar_json

args_cfg = ["--config", sys.argv[sys.argv.index("--config") + 1]] if "--config" in sys.argv else []
LOG = RAIZ / "atualizacao.log"
NO_WIN = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def log(msg):
    linha = f"[{datetime.datetime.now():%d/%m/%Y %H:%M:%S}] {msg}"
    try:
        with open(LOG, "a", encoding="utf-8") as f: f.write(linha + "\n")
    except OSError: pass
    print(linha, flush=True)


estado = {"executando": False, "etapa": "parado", "inicio": None, "fim": None, "resultado": None, "mensagem": "", "origem": None}
trava, trava_dados = threading.Lock(), threading.Lock()
pendente = {"sim": False}


def rodar(cmd):
    r = subprocess.run(cmd, cwd=RAIZ, capture_output=True, text=True, encoding="utf-8", errors="replace", creationflags=NO_WIN,
                       env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    saida = (r.stdout or "") + (r.stderr or "")
    if saida.strip(): log(saida.strip()[-3000:])
    return r.returncode, saida


def publicar(origem, so_se_mudou):
    with trava:
        if estado["executando"]:
            pendente["sim"] = True          # uma edição chegou durante a publicação: publica de novo no fim
            return False
        estado.update(executando=True, etapa="gerando", inicio=datetime.datetime.now().isoformat(timespec="seconds"),
                      fim=None, resultado=None, mensagem="", origem=origem)
    threading.Thread(target=_publicar, args=(origem, so_se_mudou), daemon=True).start()
    return True


def _publicar(origem, so_se_mudou):
    log(f"==== publicação iniciada ({origem}) ====")
    try:
        rc, saida = rodar([sys.executable, str(AQUI / "gerar_dados.py"), *args_cfg] + (["--se-mudou"] if so_se_mudou else []))
        if rc != 0: raise RuntimeError("erro ao gerar os dados (veja atualizacao.log)")
        if "Nada a fazer" in saida:
            estado.update(resultado="sem_mudanca", mensagem="As bases não mudaram desde a última publicação."); return
        estado["etapa"] = "enviando"
        rodar(["git", "pull", "--quiet", "--rebase", "origin", "main"])
        rodar(["git", "add", "-A"])
        if subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=RAIZ, creationflags=NO_WIN).returncode == 0:
            estado.update(resultado="sem_mudanca", mensagem="Nada novo para enviar."); return
        rc, _ = rodar(["git", "commit", "-m", f"Atualiza dados do painel ({origem}) {datetime.datetime.now():%d/%m/%Y %H:%M}"])
        if rc != 0: raise RuntimeError("erro no commit (veja atualizacao.log)")
        rc, _ = rodar(["git", "push", "-u", "origin", "HEAD:main"])
        if rc != 0: raise RuntimeError("erro ao enviar para o GitHub (veja atualizacao.log)")
        estado.update(resultado="ok", mensagem="Enviado. O GitHub Pages publica em 1-2 minutos.")
    except Exception as e:
        estado.update(resultado="erro", mensagem=str(e)); log(f"ERRO: {e}")
    finally:
        estado.update(executando=False, etapa="parado", fim=datetime.datetime.now().isoformat(timespec="seconds"))
        log(f"==== fim: {estado['resultado']} - {estado['mensagem']} ====")
        if pendente["sim"]:
            pendente["sim"] = False
            publicar("edição no painel", so_se_mudou=False)


# ---------- edições feitas pelo painel ----------
def tratar(corpo):
    """corpo = {"itens": [{"chave", "status", "obs"}]}  ·  status "" + remover=true desfaz (o título volta para a base)."""
    c = cfg(); dados = pasta_dados(c)
    validos = {s["status"] for s in c["status_tratados"]}
    with trava_dados:
        trat = ler_json(dados / "tratados.json", [])
        por_chave = {t["chave"]: t for t in trat}
        sienge = ler_json(dados / "sienge_atual.json", {"titulos": {}})["titulos"]
        retrato = {}
        for p in sorted((dados / "retratos").glob("*.json"))[-1:]:
            retrato = {t["chave"]: t for t in ler_json(p, {"titulos": []})["titulos"]}
        feitos = 0
        for it in corpo.get("itens", []):
            k = it.get("chave")
            if not k: continue
            if it.get("remover"):
                if k in por_chave: del por_chave[k]; feitos += 1
                continue
            st = it.get("status", "")
            if st and st not in validos: raise ValueError(f"Status inválido: {st}")
            if k in por_chave:
                por_chave[k].update(status=st, obs=it.get("obs", por_chave[k].get("obs", "")))
            else:
                base = retrato.get(k) or sienge.get(k)
                if not base: raise ValueError(f"Título {k} não está na base atual.")
                novo = {x: base.get(x) for x in ("chave", "venc", "cliente", "cd_cliente", "documento", "titulo", "parc", "tc",
                                                  "unid", "valor_original", "total", "dias", "empresa", "cd_empresa")}
                novo.update(status=st, obs=it.get("obs", ""), tratado_em=datetime.date.today().isoformat(), origem="painel")
                por_chave[k] = novo
            feitos += 1
        gravar_json(dados / "tratados.json", list(por_chave.values()))
    log(f"Tratados: {feitos} alteração(ões) pelo painel")
    return feitos


def fechamento(corpo):
    """corpo = {"mes": "AAAA-MM", "linhas": [...], "total": {...}}: o % enviado à Rita, congelado."""
    if not corpo.get("mes") or not corpo.get("linhas"): raise ValueError("Fechamento sem mês ou sem linhas.")
    dados = pasta_dados()
    with trava_dados:
        fx = [f for f in ler_json(dados / "fechamentos.json", []) if f["mes"] != corpo["mes"]]
        fx.append({"mes": corpo["mes"], "registrado_em": datetime.datetime.now().isoformat(timespec="minutes"),
                   "data_base": corpo.get("data_base"), "linhas": corpo["linhas"], "total": corpo.get("total")})
        fx.sort(key=lambda f: f["mes"])
        gravar_json(dados / "fechamentos.json", fx)
    log(f"Fechamento {corpo['mes']} registrado")


# ---------- observador das pastas ----------
def assinatura_pastas(c):
    out = []
    fontes = [(caminho(c.get("pasta_entrada", "entrada")), "*.*", False),
              (caminho(c.get("pasta_boletagem", "entrada/boletagem")), "*.pdf", True)]
    if c.get("pasta_planilha"):            # planilha de inadimplência no OneDrive: quando alguém salva, o painel atualiza
        fontes.append((caminho(c["pasta_planilha"]), c.get("padrao_planilha", "Inadimpl*Carozzo*.xlsx"), False))
    for pasta, padrao, rec in fontes:
        if not pasta.is_dir(): out.append(None); continue
        fs = (pasta.rglob(padrao) if rec else pasta.glob(padrao))
        out.append(sorted((f.name, f.stat().st_mtime) for f in fs if f.is_file() and not f.name.startswith("~$")))
    return out


def observar():
    c = cfg(); sl = c.get("servico_local", {})
    intervalo = int(sl.get("intervalo_segundos", 60)); espera = int(sl.get("estabilizar_segundos", 60))
    ultima = assinatura_pastas(c); mudou_em = None
    log(f"Observando as pastas a cada {intervalo}s. Publicação automática {espera}s depois de salvar.")
    publicar("início do serviço", so_se_mudou=True)
    while True:
        time.sleep(intervalo)
        try:
            atual = assinatura_pastas(cfg())
            if atual != ultima:
                ultima, mudou_em = atual, time.time()
                log(f"Mudança detectada nas pastas - aguardando {espera}s sem novas alterações")
            elif mudou_em and time.time() - mudou_em >= espera:
                if publicar("relatório salvo", so_se_mudou=True): mudou_em = None
        except Exception as e:
            log(f"Observador: {e}")


# ---------- servidor HTTP local ----------
class Handler(BaseHTTPRequestHandler):
    def origem_ok(self):
        o = self.headers.get("Origin")
        return o is None or o in cfg().get("servico_local", {}).get("origens", [])

    def cors(self):
        o = self.headers.get("Origin")
        if o and self.origem_ok():
            self.send_header("Access-Control-Allow-Origin", o); self.send_header("Vary", "Origin")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type")
            self.send_header("Access-Control-Allow-Private-Network", "true")

    def responder(self, cod, dados):
        corpo = json.dumps(dados, ensure_ascii=False).encode("utf-8")
        self.send_response(cod); self.cors()
        self.send_header("Content-Type", "application/json; charset=utf-8"); self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(corpo))); self.end_headers(); self.wfile.write(corpo)

    def corpo(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n).decode("utf-8") or "{}") if n else {}

    def do_OPTIONS(self):
        self.send_response(204 if self.origem_ok() else 403); self.cors(); self.end_headers()

    def do_GET(self):
        if not self.origem_ok(): return self.responder(403, {"erro": "origem não autorizada"})
        if self.path.startswith("/status"): return self.responder(200, {"servico": "painel-inadimplencia", **estado})
        self.responder(404, {"erro": "não encontrado"})

    def do_POST(self):
        if not self.origem_ok(): return self.responder(403, {"erro": "origem não autorizada"})
        try:
            if self.path.startswith("/publicar"):
                ini = publicar("botão Publicar agora", so_se_mudou=False)
                return self.responder(202 if ini else 409, {"iniciado": ini, **estado})
            if self.path.startswith("/tratar"):
                n = tratar(self.corpo()); publicar("Tratados pelo painel", so_se_mudou=False)
                return self.responder(200, {"ok": True, "alterados": n, **estado})
            if self.path.startswith("/fechamento"):
                fechamento(self.corpo()); publicar("fechamento mensal", so_se_mudou=False)
                return self.responder(200, {"ok": True, **estado})
        except Exception as e:
            log(f"ERRO na edição: {e}")
            return self.responder(400, {"erro": str(e)})
        self.responder(404, {"erro": "não encontrado"})

    def log_message(self, *a): pass


if __name__ == "__main__":
    porta = int(cfg().get("servico_local", {}).get("porta", 8733))
    try:
        srv = ThreadingHTTPServer(("127.0.0.1", porta), Handler)
    except OSError:
        log(f"A porta {porta} já está em uso - o serviço provavelmente já está rodando."); sys.exit(0)
    log(f"Serviço local iniciado em http://127.0.0.1:{porta}")
    if "--sem-observador" not in sys.argv:
        threading.Thread(target=observar, daemon=True).start()
    srv.serve_forever()
