"""Captura as telas do Portal Prontuário (Evidências) da Sambaíba.

Uso:
    pip install playwright requests
    playwright install chromium        # na sua máquina
    export PORTAL_USER='seu_usuario'
    export PORTAL_PASS='sua_senha'
    python3 capturar.py                # gera shots/*.png

Faz login no formulário do portal e tira screenshot de página inteira de
cada módulo, em resolução 2x. As imagens ficam em shots/ (fora do git).
"""
import os
from playwright.sync_api import sync_playwright

BASE = "https://nimer-prontuario.sambaibasp.cloud"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shots")

TELAS = [
    ("/",              "01_inicio"),
    ("/evidencias/nova", "02_nova_evidencia"),
    ("/evidencias",    "03_evidencias"),
    ("/validacao",     "04_validacao"),
    ("/fila-nimer",    "05_fila_nimer"),
    ("/retorno-nimer", "06_retorno_nimer"),
    ("/usuarios",      "07_usuarios"),
]


def main():
    user = os.environ["PORTAL_USER"]
    pwd = os.environ["PORTAL_PASS"]
    os.makedirs(OUT, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport={"width": 1440, "height": 900},
                                  device_scale_factor=2)
        page = ctx.new_page()
        # login
        page.goto(f"{BASE}/login", wait_until="networkidle")
        page.fill("input[name=login]", user)
        page.fill("input[name=senha]", pwd)
        page.click("button[type=submit]")
        page.wait_for_load_state("networkidle")
        if "/login" in page.url:
            raise SystemExit("Login falhou — confira PORTAL_USER / PORTAL_PASS.")
        # capturas
        for path, nome in TELAS:
            page.goto(BASE + path, wait_until="networkidle")
            page.wait_for_timeout(1500)  # tabelas montam via JS
            page.screenshot(path=os.path.join(OUT, f"{nome}.png"), full_page=True)
            print("ok", nome)
        browser.close()


if __name__ == "__main__":
    main()
