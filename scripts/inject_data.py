"""
inject_data.py

Le o dashboard HTML, localiza `const D = {...};` embutido no <script> e
substitui pelo conteudo de um JSON novo (gerado por extract_kpis.build()).

Uso:
    from inject_data import inject
    inject("Dashboard Regulatórios - RASCUNHO Novos Indicadores.html", "kpis_all_new.json")
"""

import json
import re


_PATTERN = re.compile(r"const D = (\{.*?\});\n", re.DOTALL)


def inject(html_path, json_path):
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    new_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))

    match = _PATTERN.search(html)
    if not match:
        raise ValueError("Não foi possível localizar 'const D = {...};' no HTML.")

    new_html = html[: match.start()] + f"const D = {new_json};\n" + html[match.end():]

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(new_html)

    return new_html


if __name__ == "__main__":
    import sys

    html_path = sys.argv[1]
    json_path = sys.argv[2]
    inject(html_path, json_path)
    print(f"Injetado {json_path} em {html_path}")
