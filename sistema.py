from bs4 import BeautifulSoup
import re
import os
from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time

# ── Configuração ──────────────────────────────────────────────────────────────
OUTPUT       = r"Z:\Ester-Dev\advogados.xlsx"
OABS_IGNORAR = {"MG-176171", "MG-124826"}
DATA_INICIO  = "2024-01-01"
DATA_FIM     = date.today().isoformat()
URL_BASE     = (
    "https://comunica.pje.jus.br/consulta"
    "?dataDisponibilizacaoInicio={inicio}"
    "&dataDisponibilizacaoFim={fim}"
    "&numeroProcesso={numero}"
)
# ─────────────────────────────────────────────────────────────────────────────


def get_html_via_selenium(url):
    options = Options()
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1400,900")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })

    try:
        print(f"  Abrindo: {url}")
        driver.get(url)

        print("  Aguardando resultados carregarem...")
        for i in range(20):
            time.sleep(2)
            html = driver.page_source
            if any(k in html for k in ["col-md-10", "Advogado", "dvogado"]):
                print(f"  Conteúdo detectado após {(i+1)*2}s.")
                break
            print(f"  Aguardando... {(i+1)*2}s")
        else:
            print("  AVISO: conteúdo não detectado no tempo limite.")

        html = driver.page_source
        with open(r"Z:\Ester-Dev\diagnostico.html", "w", encoding="utf-8") as f:
            f.write(html)
        return html

    finally:
        driver.quit()


def parse_advogados(html, numero_processo):
    soup = BeautifulSoup(html, "html.parser")
    advogados = []

    # Busca todas as divs que contenham "col-md-10" entre suas classes
    for div in soup.find_all("div"):
        classes = div.get("class", [])
        if "col-md-10" in classes:
            raw = div.get_text(strip=True)
            match = re.match(r"^(.+?)\s*-\s*OAB\s+([A-Z]{2})-(\d+)$", raw)
            if match:
                oab_key = f"{match.group(2)}-{match.group(3)}"
                if oab_key in OABS_IGNORAR:
                    print(f"  [ignorado] {match.group(1).strip()} | OAB {oab_key}")
                    continue
                # Evita duplicatas
                if not any(a["Número OAB"] == match.group(3) for a in advogados):
                    advogados.append({
                        "Processo": numero_processo,
                        "Nome": match.group(1).strip(),
                        "UF OAB": match.group(2).strip(),
                        "Número OAB": match.group(3).strip(),
                    })

    return advogados


def save_xlsx(advogados, path):
    headers = ["Processo", "Nome", "UF OAB", "Número OAB"]
    header_fill = PatternFill("solid", start_color="003366")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    row_font    = Font(name="Arial", size=11)

    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
        next_row = ws.max_row + 1
        print(f"  Arquivo existente — adicionando a partir da linha {next_row}.")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Advogados"
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 45
        ws.column_dimensions["C"].width = 10
        ws.column_dimensions["D"].width = 15
        ws.row_dimensions[1].height = 20
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        next_row = 2

    for adv in advogados:
        for col_idx, key in enumerate(headers, 1):
            cell = ws.cell(row=next_row, column=col_idx, value=adv[key])
            cell.font = row_font
            cell.alignment = Alignment(horizontal="left", vertical="center")
        next_row += 1

    wb.save(path)
    print(f"  Planilha salva em: {path}")


# ── Main ──────────────────────────────────────────────────────────────────────
print("=" * 50)
print("  Extrator de Advogados - PJe")
print("=" * 50)

numero_raw = input("\nDigite o número do processo: ").strip()
numero = re.sub(r"[\s.\-/]", "", numero_raw)

if not numero:
    print("Nenhum número informado. Encerrando.")
else:
    numero_fmt = (
        f"{numero[:7]}-{numero[7:9]}.{numero[9:13]}.{numero[13]}.{numero[14:16]}.{numero[16:]}"
        if len(numero) == 20 else numero_raw
    )
    print(f"\nNúmero formatado : {numero_fmt}")
    print(f"Período de busca : {DATA_INICIO} até {DATA_FIM}")

    url = URL_BASE.format(inicio=DATA_INICIO, fim=DATA_FIM, numero=numero)

    try:
        html = get_html_via_selenium(url)
        advogados = parse_advogados(html, numero_fmt)

        if advogados:
            save_xlsx(advogados, OUTPUT)
            print(f"\nAdicionados {len(advogados)} advogado(s):")
            for a in advogados:
                print(f"  {a['Nome']} | OAB {a['UF OAB']}-{a['Número OAB']}")
        else:
            print("\nNenhum advogado encontrado.")
            print("Abra Z:\\Ester-Dev\\diagnostico.html no navegador para verificar.")

    except Exception as e:
        print(f"\nErro: {e}")

input("\nPressione Enter para fechar...")
