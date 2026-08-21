#!/usr/bin/env python3
"""
Orquestrador da atualizacao automatica do Dashboard Regulatorios - Pronutrition.

Passos:
1. Le a planilha "LISTA MESTRE DE PRODUTOS - PRONUTRITION.xlsx".
2. Recalcula todos os KPIs (extract_kpis.build).
3. Injeta o novo JSON no dashboard HTML local (OneDrive) via inject_data.inject.
4. Clona/atualiza o repositorio GitHub, copia o dashboard + dados atualizados,
   comita e faz push (o que dispara o deploy automatico no Netlify, pois o
   site esta conectado ao repo via Continuous Deployment).

Uso:
    python3 update_dashboard.py

IMPORTANTE: os caminhos abaixo apontam para dentro do mount da sandbox
("/sessions/<id-da-sessao>/mnt/..."), e o <id-da-sessao> muda a cada nova
sessao do Cowork. Por isso os caminhos SEMPRE devem ser passados via
variaveis de ambiente (DASH_XLSX_PATH e DASH_HTML_PATH) resolvidas no
inicio de cada execucao (ver instrucoes da skill). Os valores default
abaixo servem apenas de fallback/documentacao e provavelmente NAO existirao
em uma sessao nova.
"""
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from extract_kpis import build  # noqa: E402
from inject_data import inject  # noqa: E402

# ---- Caminhos: SEMPRE fornecidos via variavel de ambiente pela skill ----
# Nomes de pasta usados para localizar o mount, caso as env vars nao sejam passadas:
XLSX_REL = "1A. Lista Mestre de Produtos - PRONUTRITION/LISTA MESTRE DE PRODUTOS - PRONUTRITION.xlsx"
HTML_REL = "20. Time Regulatórios/8. INDICADORES/D. Dashboards/Dashboard Regulatórios - RASCUNHO Novos Indicadores.html"
TOKEN_REL = "20. Time Regulatórios/8. INDICADORES/D. Dashboards/.config/github_token.txt"


def _resolve(env_var, rel_path, fallback):
    val = os.environ.get(env_var)
    if val:
        return val
    if os.path.exists(fallback):
        return fallback
    # tenta descobrir o mount atual pesquisando /sessions/*/mnt/<rel_path>
    import glob
    matches = glob.glob(f"/sessions/*/mnt/{rel_path}")
    if matches:
        return matches[0]
    raise RuntimeError(
        f"Nao foi possivel localizar '{rel_path}'. Defina a variavel de ambiente {env_var} "
        f"com o caminho correto para esta sessao (veja a secao 'Shell access' do system prompt)."
    )


XLSX_PATH = _resolve("DASH_XLSX_PATH", XLSX_REL, "/sessions/exciting-pensive-cerf/mnt/" + XLSX_REL)
DASHBOARD_HTML_PATH = _resolve("DASH_HTML_PATH", HTML_REL, "/sessions/exciting-pensive-cerf/mnt/" + HTML_REL)
TOKEN_FILE = _resolve("DASH_TOKEN_PATH", TOKEN_REL, "/sessions/exciting-pensive-cerf/mnt/" + TOKEN_REL)
KPIS_JSON_PATH = os.path.join(SCRIPT_DIR, "kpis_all.json")

GITHUB_REPO = "github.com/luanalante-byte/Dashboard-Regulat-rio.git"
CLONE_DIR = "/tmp/dashboard_repo_autoupdate_v2"


def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def read_token():
    token = os.environ.get("GITHUB_TOKEN")
    if token:
        return token.strip()
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    raise RuntimeError(
        f"Token do GitHub nao encontrado. Defina GITHUB_TOKEN ou crie o arquivo {TOKEN_FILE}"
    )


def step_extract_and_inject():
    log(f"Lendo planilha: {XLSX_PATH}")
    if not os.path.exists(XLSX_PATH):
        raise FileNotFoundError(f"Planilha nao encontrada em {XLSX_PATH}")

    data = build(XLSX_PATH)
    log(
        "KPIs recalculados: "
        f"cotacoes={sum(data['cotacoes']['status'].values())}, "
        f"docs_tipos={len(data['docs']['tipos'])}, "
        f"notif_vig={data['notif_vig']['total']}, "
        f"gastos_clientes={data['gastos_clientes']['n_registros']}, "
        f"rev_arte={data['rev_arte']['total_revisoes']}"
    )

    with open(KPIS_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

    if not os.path.exists(DASHBOARD_HTML_PATH):
        raise FileNotFoundError(f"Dashboard HTML nao encontrado em {DASHBOARD_HTML_PATH}")

    inject(DASHBOARD_HTML_PATH, KPIS_JSON_PATH)
    log(f"Dashboard local atualizado: {DASHBOARD_HTML_PATH}")
    log(f"Atualizado em: {data['meta']['atualizado_em']}")
    return data


def run(cmd, cwd=None, check=True):
    log("$ " + " ".join(cmd))
    result = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if result.stdout.strip():
        print(result.stdout.strip())
    if result.stderr.strip():
        print(result.stderr.strip())
    if check and result.returncode != 0:
        raise RuntimeError(f"Comando falhou ({result.returncode}): {' '.join(cmd)}")
    return result


def step_push_github(data):
    token = read_token()
    remote_url = f"https://{token}@{GITHUB_REPO}"

    if os.path.exists(CLONE_DIR):
        shutil.rmtree(CLONE_DIR)

    log("Clonando repositorio...")
    run(["git", "clone", remote_url, CLONE_DIR])

    run(["git", "config", "user.email", "labdesenvolvimento@pronutrition.com.br"], cwd=CLONE_DIR)
    run(["git", "config", "user.name", "Pronutrition Regulatorio (auto-update)"], cwd=CLONE_DIR)

    os.makedirs(os.path.join(CLONE_DIR, "dashboard"), exist_ok=True)
    os.makedirs(os.path.join(CLONE_DIR, "data"), exist_ok=True)
    os.makedirs(os.path.join(CLONE_DIR, "scripts"), exist_ok=True)

    shutil.copy(DASHBOARD_HTML_PATH, os.path.join(CLONE_DIR, "dashboard", "index.html"))
    shutil.copy(KPIS_JSON_PATH, os.path.join(CLONE_DIR, "data", "kpis_all.json"))
    shutil.copy(os.path.join(SCRIPT_DIR, "extract_kpis.py"), os.path.join(CLONE_DIR, "scripts", "extract_kpis.py"))
    shutil.copy(os.path.join(SCRIPT_DIR, "inject_data.py"), os.path.join(CLONE_DIR, "scripts", "inject_data.py"))
    shutil.copy(__file__, os.path.join(CLONE_DIR, "scripts", "update_dashboard.py"))

    run(["git", "add", "-A"], cwd=CLONE_DIR)
    status = run(["git", "status", "--porcelain"], cwd=CLONE_DIR)
    if not status.stdout.strip():
        log("Nenhuma alteracao detectada em relacao ao ultimo commit. Nada para enviar.")
        return False

    commit_msg = f"Atualizacao automatica dos indicadores ({data['meta']['atualizado_em']})"
    run(["git", "commit", "-m", commit_msg], cwd=CLONE_DIR)
    run(["git", "push", "origin", "main"], cwd=CLONE_DIR)
    log("Push concluido. O deploy no Netlify deve iniciar automaticamente (Continuous Deployment).")
    return True


def main():
    data = step_extract_and_inject()
    pushed = step_push_github(data)
    if pushed:
        log("Atualizacao completa: dashboard local + GitHub + (Netlify via auto-deploy).")
    else:
        log("Dashboard local atualizado. Sem mudancas para publicar no GitHub.")


if __name__ == "__main__":
    main()
