"""
Monitor de Normas Legales - Diario Oficial El Peruano
=====================================================

Revisa https://diariooficial.elperuano.pe/Normas en un rango de fechas y
selecciona las normas que involucran a:

  1. ATU (Autoridad de Transporte Urbano para Lima y Callao) y transporte urbano
  2. Normatividad presupuestal (transferencias, créditos suplementarios, etc.)
  3. Inversiones (Invierte.pe, proyectos de inversión, APP, OxI, etc.)

Genera un resumen en Word (.docx) y, si está configurado en config.ini,
lo envía por Telegram (listado de normas + el Word adjunto).

Uso:
    python monitor_normas.py                    # normas de hoy
    python monitor_normas.py --dias 7           # últimos 7 días
    python monitor_normas.py --desde 01/10/2026 --hasta 22/10/2026
    python monitor_normas.py --no-enviar        # solo genera el Word
"""

import argparse
import configparser
import os
import re
import sys
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from html import escape
from pathlib import Path

try:
    # Usa los certificados de Windows (necesario si la red inspecciona HTTPS)
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm

BASE_DIR = Path(__file__).resolve().parent
HORA_LIMA = timezone(timedelta(hours=-5))  # Perú no tiene horario de verano
URL_PORTAL = "https://diariooficial.elperuano.pe/Normas"
URL_FILTRO = "https://diariooficial.elperuano.pe/Normas/Filtro"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/130.0 Safari/537.36",
    "Accept-Language": "es-PE,es;q=0.9",
}

# --------------------------------------------------------------------------
# Criterios de búsqueda (sobre texto normalizado: minúsculas y sin tildes)
# --------------------------------------------------------------------------
CATEGORIAS = {
    "ATU y transporte urbano": [
        r"autoridad de transporte urbano",
        r"transporte urbano (de|para) lima",
        r"lima y callao.{0,40}transporte|transporte.{0,40}lima y callao",
        r"linea \d+ (de la red basica )?del metro",
        r"red basica del metro",
        r"\bel metropolitano\b|sistema (de transporte )?metropolitano de transporte|\bcosac\b",
        r"corredores? complementarios?",
        r"sistema integrado de transporte",
    ],
    "Normatividad presupuestal": [
        r"presupuest",
        r"transferencias? de partidas",
        r"transferencias? financieras?",
        r"creditos? suplementarios?",
        r"modificaci(on|ones) presupuestari",
        r"demandas? adicionales",
        r"reserva de contingencia",
        r"recursos ordinarios",
        r"equilibrio financiero",
        r"endeudamiento",
        r"operaci(on|ones) de endeudamiento",
        r"programacion multianual presupuestaria",
        r"gasto publico",
        r"hacienda publica",
        r"tesoro publico",
    ],
    "Inversiones": [
        r"inversi(on|ones) public",
        r"invierte\.?pe",
        r"programacion multianual y gestion de inversiones",
        r"proyectos? de inversion",
        r"\bioarr\b",
        r"programa multianual de inversiones",
        r"cartera de inversiones",
        r"asociaci(on|ones) publico.?privadas?",
        r"proinversion",
        r"obras? por impuestos",
        r"contratos? de concesion|concesiones? (vial|ferroviari|de transporte|de infraestructura)",
        r"expedientes? tecnicos?",
        r"autoridad nacional de infraestructura",
        r"(obras?|proyectos?|brecha) de infraestructura",
        r"contrataciones (publicas|del estado)",
        r"inversiones",
    ],
}

# Actos administrativos de personal que no interesan, salvo que mencionen a la ATU
REGEX_EXCLUIR = re.compile(
    r"^(designan|dan por concluid|aceptan (la )?renuncia|encargan|autorizan (el )?viaje|"
    r"otorgan concesion unica|otorgan concesion electrica)"
)

# Menciones a la ATU dentro del texto completo (sobre texto original)
REGEX_ATU_TEXTO = re.compile(
    r"\bATU\b|Autoridad de Transporte Urbano( para Lima y Callao)?", re.IGNORECASE
)
REGEX_MONTO = re.compile(r"S/\s?[\d.,]+(?:\s?(?:millones|mil))?")


@dataclass
class Norma:
    fecha: str
    entidad: str
    titulo: str
    sumilla: str
    url: str
    url_pdf: str
    extraordinaria: bool
    texto: str = ""
    categorias: list = field(default_factory=list)
    menciones_atu: list = field(default_factory=list)


def normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFD", texto or "")
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto.lower())


# --------------------------------------------------------------------------
# Descarga y lectura del portal
# --------------------------------------------------------------------------
def nueva_sesion() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    s.get(URL_PORTAL, timeout=30)  # obtiene cookies del portal
    return s


def listar_normas(s: requests.Session, desde: date, hasta: date, solo_extra=False):
    params = {"dateparam": datetime.now().strftime("%m/%d/%Y 00:00:00")}
    data = {
        "cddesde": desde.strftime("%d/%m/%Y"),
        "cdhasta": hasta.strftime("%d/%m/%Y"),
        "btnBuscar": "Buscar",
    }
    if solo_extra:
        data["tipo"] = "on"
    r = s.post(URL_FILTRO, params=params, data=data, timeout=120,
               headers={"X-Requested-With": "XMLHttpRequest"})
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")

    normas = []
    for art in soup.select("article.edicionesoficiales_articulos"):
        cont = art.select_one(".ediciones_texto")
        if not cont:
            continue
        enlace = cont.select_one("h5 a")
        fecha_tag = cont.select_one("p b")
        parrafos = cont.find_all("p")
        pdf = art.select_one(".ediciones_botones a[href$='/pdf']")
        normas.append(Norma(
            fecha=(fecha_tag.get_text(strip=True).replace("Fecha:", "").strip()
                   if fecha_tag else ""),
            entidad=cont.h4.get_text(" ", strip=True) if cont.h4 else "",
            titulo=enlace.get_text(" ", strip=True) if enlace else "",
            sumilla=parrafos[-1].get_text(" ", strip=True) if len(parrafos) > 1 else "",
            url=enlace["href"] if enlace else "",
            url_pdf=pdf["href"] if pdf else "",
            extraordinaria=bool(cont.select_one(".extraordinaria")),
        ))
    return normas


def codigo_norma(norma: Norma) -> str:
    """Código de publicación, p. ej. 2567343-1, tomado del enlace del portal."""
    m = re.search(r"/(\d+-\d+)(?:/|$)", norma.url)
    return m.group(1) if m else ""


def url_texto(norma: Norma) -> str:
    # Los enlaces /dispositivo/NL/... del portal suelen dar 404; este visor sí responde
    op = codigo_norma(norma)
    return f"https://busquedas.elperuano.pe/api/visor_html/{op}" if op else norma.url


def descargar_texto(s: requests.Session, norma: Norma, reintentos=2) -> str:
    op = codigo_norma(norma)
    if not op:
        return ""
    for intento in range(reintentos + 1):
        try:
            r = s.get(url_texto(norma), timeout=30)
            if r.status_code == 404:
                return ""
            r.raise_for_status()
            r.encoding = "utf-8"
            soup = BeautifulSoup(r.text, "lxml")
            nodo = soup.find(id=f"x{op}") or soup.select_one(".story") or soup.body
            return nodo.get_text("\n", strip=True) if nodo else ""
        except requests.RequestException:
            time.sleep(2 * (intento + 1))
    return ""


# --------------------------------------------------------------------------
# Clasificación y extractos
# --------------------------------------------------------------------------
PATRONES = {cat: [re.compile(p) for p in pats] for cat, pats in CATEGORIAS.items()}


def clasificar(norma: Norma):
    cabecera = normalizar(f"{norma.entidad} {norma.titulo} {norma.sumilla}")
    cats = [cat for cat, pats in PATRONES.items() if any(p.search(cabecera) for p in pats)]

    # La ATU puede aparecer solo en el texto o anexos (p. ej. transferencias del MEF);
    # en las normas que emite la propia ATU el extracto no aporta nada
    emitida_por_atu = "autoridad de transporte urbano" in normalizar(norma.entidad)
    if norma.texto and not emitida_por_atu:
        for m in REGEX_ATU_TEXTO.finditer(norma.texto):
            ini, fin = max(0, m.start() - 220), min(len(norma.texto), m.end() + 220)
            frag = re.sub(r"\s+", " ", norma.texto[ini:fin]).strip()
            if not any(frag[:80] in x for x in norma.menciones_atu):
                norma.menciones_atu.append(f"…{frag}…")
            if len(norma.menciones_atu) >= 3:
                break
    menciona_atu = bool(emitida_por_atu or norma.menciones_atu
                        or REGEX_ATU_TEXTO.search(norma.sumilla))
    if REGEX_EXCLUIR.search(normalizar(norma.sumilla)) and not menciona_atu:
        cats = []
    if menciona_atu and "ATU y transporte urbano" not in cats:
        cats.insert(0, "ATU y transporte urbano")
    norma.categorias = cats
    return bool(cats)


def extracto_articulos(texto: str, max_chars=700) -> str:
    """Primer artículo de la parte resolutiva (lo que la norma dispone)."""
    if not texto:
        return ""
    m = re.search(r"(Art[íi]culo\s+1\.?[º°o]?\.?\s*[-–.]?.*)", texto, re.S)
    frag = m.group(1) if m else ""
    frag = re.sub(r"\s+", " ", frag).strip()
    if len(frag) > max_chars:
        frag = frag[:max_chars].rsplit(" ", 1)[0] + "…"
    return frag


def montos(texto: str, n=4):
    vistos = []
    for m in REGEX_MONTO.findall(texto or ""):
        m = m.strip().rstrip(".,")
        if m not in vistos and re.search(r"\d", m):
            vistos.append(m)
        if len(vistos) >= n:
            break
    return vistos


# --------------------------------------------------------------------------
# Documento Word
# --------------------------------------------------------------------------
def agregar_hipervinculo(parrafo, url, texto):
    rel = parrafo.part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), rel)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color"); color.set(qn("w:val"), "0563C1")
    subr = OxmlElement("w:u"); subr.set(qn("w:val"), "single")
    rpr.append(color); rpr.append(subr); run.append(rpr)
    t = OxmlElement("w:t"); t.text = texto; run.append(t)
    hl.append(run)
    parrafo._p.append(hl)


def sombrear(celda, color_hex):
    tcpr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), color_hex)
    tcpr.append(shd)


def generar_word(seleccion, total, desde, hasta, ruta: Path):
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2)
    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(10.5)

    t = doc.add_heading("Resumen de Normas Legales – Diario Oficial El Peruano", 0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    periodo = (desde.strftime("%d/%m/%Y") if desde == hasta
               else f"{desde:%d/%m/%Y} al {hasta:%d/%m/%Y}")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Periodo revisado: {periodo}\n").bold = True
    p.add_run(f"Generado: {datetime.now():%d/%m/%Y %H:%M}  ·  Fuente: diariooficial.elperuano.pe")

    doc.add_heading("1. Resumen ejecutivo", 1)
    doc.add_paragraph(
        f"Se revisaron {total} normas publicadas en el periodo. "
        f"{len(seleccion)} resultan de interés para la ATU por tratar sobre la entidad, "
        f"la normatividad presupuestal o las inversiones."
    )
    tabla = doc.add_table(rows=1, cols=2)
    tabla.style = "Light Grid Accent 1"
    tabla.rows[0].cells[0].text = "Tema"
    tabla.rows[0].cells[1].text = "N.° de normas"
    for cat in CATEGORIAS:
        fila = tabla.add_row().cells
        fila[0].text = cat
        fila[1].text = str(sum(cat in n.categorias for n in seleccion))

    atu = [n for n in seleccion if n.menciones_atu or "ATU y transporte urbano" in n.categorias]
    if atu:
        doc.add_paragraph()
        p = doc.add_paragraph()
        r = p.add_run("Atención – normas que mencionan a la ATU o al transporte urbano de Lima y Callao:")
        r.bold = True
        r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        for n in atu:
            doc.add_paragraph(f"{n.titulo} ({n.entidad}): {n.sumilla}", style="List Bullet")

    if not seleccion:
        doc.add_paragraph("No se encontraron normas relacionadas en el periodo.")
        doc.save(ruta)
        return

    num = 2
    for cat in CATEGORIAS:
        normas_cat = [n for n in seleccion if n.categorias and n.categorias[0] == cat]
        if not normas_cat:
            continue
        doc.add_heading(f"{num}. {cat} ({len(normas_cat)})", 1)
        num += 1

        tabla = doc.add_table(rows=1, cols=4)
        tabla.style = "Table Grid"
        for i, enc in enumerate(["Fecha", "Entidad", "Norma", "Sumilla"]):
            c = tabla.rows[0].cells[i]
            c.text = enc
            c.paragraphs[0].runs[0].bold = True
            sombrear(c, "1F4E79")
            c.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        for n in normas_cat:
            fila = tabla.add_row().cells
            fila[0].text = n.fecha + (" (Ext.)" if n.extraordinaria else "")
            fila[1].text = n.entidad
            fila[2].text = ""
            agregar_hipervinculo(fila[2].paragraphs[0], url_texto(n), n.titulo)
            fila[3].text = n.sumilla
        for fila in tabla.rows:
            for i, ancho in enumerate([Cm(2.2), Cm(3.5), Cm(4), Cm(7.3)]):
                fila.cells[i].width = ancho
                for par in fila.cells[i].paragraphs:
                    for run in par.runs:
                        run.font.size = Pt(9)

        doc.add_paragraph()
        for n in normas_cat:
            doc.add_heading(n.titulo, 3)
            p = doc.add_paragraph()
            p.add_run("Entidad: ").bold = True
            p.add_run(f"{n.entidad}   ")
            p.add_run("Fecha: ").bold = True
            p.add_run(n.fecha + ("  (Edición Extraordinaria)" if n.extraordinaria else ""))
            p = doc.add_paragraph()
            p.add_run("Temas: ").bold = True
            p.add_run(", ".join(n.categorias))
            p = doc.add_paragraph()
            p.add_run("Sumilla: ").bold = True
            p.add_run(n.sumilla)
            disp = extracto_articulos(n.texto)
            if disp:
                p = doc.add_paragraph()
                p.add_run("Dispone: ").bold = True
                p.add_run(disp)
            mts = montos(n.texto)
            if mts:
                p = doc.add_paragraph()
                p.add_run("Montos citados: ").bold = True
                p.add_run(", ".join(mts))
            for frag in n.menciones_atu:
                p = doc.add_paragraph()
                r = p.add_run("Mención a la ATU: ")
                r.bold = True
                r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
                p.add_run(frag).italic = True
            p = doc.add_paragraph()
            agregar_hipervinculo(p, url_texto(n), "Ver texto completo")
            p.add_run("   ")
            agregar_hipervinculo(p, n.url, "Ver en El Peruano")
            if n.url_pdf:
                p.add_run("   ")
                agregar_hipervinculo(p, n.url_pdf, "PDF")

    doc.save(ruta)


# --------------------------------------------------------------------------
# Envío
# --------------------------------------------------------------------------
def texto_breve(seleccion, total, desde, hasta):
    """Listado en HTML de Telegram, partido en mensajes de menos de 4096 caracteres."""
    periodo = (f"{desde:%d/%m/%Y}" if desde == hasta else f"{desde:%d/%m/%Y} al {hasta:%d/%m/%Y}")
    lineas = [f"<b>Normas El Peruano – {periodo}</b>",
              f"Revisadas: {total} | De interés: {len(seleccion)}"]
    for cat in CATEGORIAS:
        normas_cat = [n for n in seleccion if n.categorias and n.categorias[0] == cat]
        if normas_cat:
            lineas.append(f"\n<b>{escape(cat)} ({len(normas_cat)})</b>")
            for n in normas_cat:
                lineas.append(f'• <a href="{escape(url_texto(n))}">{escape(n.titulo)}</a> '
                              f"({escape(n.entidad)}): {escape(n.sumilla[:200])}")
    mensajes, actual = [], ""
    for linea in lineas:
        if len(actual) + len(linea) > 3800:
            mensajes.append(actual)
            actual = ""
        actual += linea + "\n"
    mensajes.append(actual)
    return mensajes


def enviar_telegram(cfg, ruta_docx: Path, mensajes, titulo: str):
    t = cfg["telegram"]
    token = os.environ.get("MONITOR_TELEGRAM_TOKEN") or t.get("bot_token", "")
    chat_id = os.environ.get("MONITOR_TELEGRAM_CHAT_ID") or t.get("chat_id", "")
    if not (token and chat_id):
        print("[telegram] Falta bot_token o chat_id en config.ini; no se envía.")
        return
    api = f"https://api.telegram.org/bot{token}"
    for texto in mensajes:
        r = requests.post(f"{api}/sendMessage", timeout=60, data={
            "chat_id": chat_id, "text": texto, "parse_mode": "HTML",
            "disable_web_page_preview": "true"})
        if not r.ok:
            print(f"[telegram] Error {r.status_code}: {r.text[:300]}")
            return
    with open(ruta_docx, "rb") as fh:
        r = requests.post(f"{api}/sendDocument", timeout=120,
                          data={"chat_id": chat_id, "caption": titulo},
                          files={"document": (ruta_docx.name, fh)})
    if not r.ok:
        print(f"[telegram] Error al enviar el Word {r.status_code}: {r.text[:300]}")
        return
    print(f"[telegram] Enviado: {len(mensajes)} mensaje/s + Word")


# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Monitor de normas El Peruano para la ATU")
    ap.add_argument("--desde", help="dd/mm/aaaa (por defecto: hoy)")
    ap.add_argument("--hasta", help="dd/mm/aaaa (por defecto: hoy)")
    ap.add_argument("--dias", type=int, help="revisar los últimos N días (incluye hoy)")
    ap.add_argument("--sin-texto", action="store_true",
                    help="no descargar el texto completo (más rápido, pero no detecta "
                         "menciones a la ATU dentro de anexos)")
    ap.add_argument("--no-enviar", action="store_true", help="solo genera el Word")
    ap.add_argument("--chat-id", action="store_true",
                    help="muestra el chat_id de quien escribió al bot y termina")
    ap.add_argument("--diario", action="store_true",
                    help="normas de hoy más las ediciones extraordinarias de ayer "
                         "(publicadas de noche, después de la revisión del día anterior)")
    args = ap.parse_args()

    hoy = datetime.now(HORA_LIMA).date()
    hasta = datetime.strptime(args.hasta, "%d/%m/%Y").date() if args.hasta else hoy
    if args.dias:
        desde = hasta - timedelta(days=args.dias - 1)
    else:
        desde = datetime.strptime(args.desde, "%d/%m/%Y").date() if args.desde else hasta

    cfg = configparser.ConfigParser()
    cfg.read(BASE_DIR / "config.ini", encoding="utf-8")
    for sec in ("telegram", "general"):
        if not cfg.has_section(sec):
            cfg.add_section(sec)

    if args.chat_id:
        token = os.environ.get("MONITOR_TELEGRAM_TOKEN") or cfg["telegram"].get("bot_token", "")
        if not token:
            print("Primero coloque bot_token en config.ini.")
            return 1
        api = f"https://api.telegram.org/bot{token}"
        yo = requests.get(f"{api}/getMe", timeout=30).json()
        print(f"Bot: @{yo.get('result', {}).get('username')}  {yo.get('description', '')}")
        r = requests.get(f"{api}/getUpdates", timeout=30).json()
        if not r.get("ok"):
            print(f"Telegram respondió: {r.get('description')}")
        chats = {}
        for u in r.get("result", []):
            chat = (u.get("message") or u.get("channel_post") or {}).get("chat")
            if chat:
                chats[chat["id"]] = chat.get("title") or chat.get("first_name", "")
        if not chats:
            print("No hay mensajes: envíe un mensaje a su bot y vuelva a ejecutar.")
        for cid, nombre in chats.items():
            print(f"chat_id = {cid}   ({nombre})")
        return 0

    print(f"Consultando normas del {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}…")
    s = nueva_sesion()
    normas = listar_normas(s, desde, hasta)
    if args.diario:
        urls = {n.url for n in normas}
        ayer = desde - timedelta(days=1)
        extra = [n for n in listar_normas(s, ayer, ayer, solo_extra=True)
                 if n.extraordinaria and n.url not in urls]
        print(f"  + {len(extra)} de ediciones extraordinarias del {ayer:%d/%m/%Y}")
        normas = extra + normas
    print(f"  {len(normas)} normas publicadas.")

    if not args.sin_texto and normas:
        print("  Descargando texto completo de cada norma…")
        with ThreadPoolExecutor(max_workers=4) as ex:
            for i, (n, txt) in enumerate(zip(normas, ex.map(lambda n: descargar_texto(s, n), normas)), 1):
                n.texto = txt
                if i % max(25, len(normas) // 10) == 0 or i == len(normas):
                    print(f"    {i}/{len(normas)}", flush=True)

    seleccion = [n for n in normas if clasificar(n)]
    orden = list(CATEGORIAS)
    seleccion.sort(key=lambda n: (orden.index(n.categorias[0]), n.fecha, n.entidad))
    print(f"  {len(seleccion)} normas de interés.")

    carpeta = Path(cfg["general"].get("carpeta_salida", BASE_DIR / "resumenes"))
    carpeta.mkdir(parents=True, exist_ok=True)
    nombre = (f"Normas_ElPeruano_{desde:%Y%m%d}.docx" if desde == hasta
              else f"Normas_ElPeruano_{desde:%Y%m%d}_{hasta:%Y%m%d}.docx")
    ruta = carpeta / nombre
    generar_word(seleccion, len(normas), desde, hasta, ruta)
    print(f"  Word generado: {ruta}")

    if args.no_enviar:
        return
    titulo = f"Normas El Peruano – {len(seleccion)} de interés ({desde:%d/%m/%Y}" + \
             ("" if desde == hasta else f" al {hasta:%d/%m/%Y}") + ")"
    try:
        enviar_telegram(cfg, ruta, texto_breve(seleccion, len(normas), desde, hasta), titulo)
    except Exception as e:  # noqa: BLE001
        print(f"[telegram] Error: {e}")


if __name__ == "__main__":
    sys.exit(main())
