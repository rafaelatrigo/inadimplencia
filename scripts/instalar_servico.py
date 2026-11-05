"""Instala o serviço local do painel de inadimplência: cria o atalho de inicialização do Windows e liga o serviço agora.
O atalho é gravado em Unicode, para funcionar mesmo com acento no caminho da pasta (ex.: "Inadimplência")."""
import os, sys, pathlib, subprocess, time, urllib.request

AQUI = pathlib.Path(__file__).resolve().parent
servico = AQUI / "servico_local.py"
pyw = pathlib.Path(sys.executable).with_name("pythonw.exe")
if not pyw.exists():
    sys.exit(f"Não encontrei o pythonw.exe ao lado de {sys.executable}. Reinstale o Python marcando 'Add python.exe to PATH'.")
inicializar = pathlib.Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
vbs = inicializar / "PainelInadimplencia.vbs"
vbs.write_text(f'CreateObject("WScript.Shell").Run """{pyw}"" ""{servico}""", 0, False\r\n', encoding="utf-16")
print(f"Atalho de inicialização: {vbs}")

subprocess.Popen([str(pyw), str(servico)], cwd=str(AQUI.parent),
                 creationflags=getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NO_WINDOW", 0))
for _ in range(15):
    time.sleep(1)
    try:
        with urllib.request.urlopen("http://127.0.0.1:8733/status", timeout=2) as r:
            if b"painel-inadimplencia" in r.read():
                print("Serviço rodando em http://127.0.0.1:8733")
                break
    except Exception:
        pass
else:
    log = AQUI.parent / "atualizacao.log"
    print("O serviço não respondeu. Veja o arquivo atualizacao.log" if log.exists() else
          "O serviço não respondeu e não criou o atualizacao.log. Rode no terminal:  python scripts\\servico_local.py  e me mande o erro.")
    sys.exit(1)
