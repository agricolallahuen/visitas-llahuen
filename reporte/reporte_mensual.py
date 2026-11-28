"""Reporte mensual de visitas (PowerPoint) - Agrícola Llahuen.

Arma una presentación con el formato de la "Base presentación mensual":
portada, fotos de predios, planta y fruto, resumen de visitas (2 por
diapositiva), variedades nuevas, problemas detectados, oportunidades,
agenda próxima y cierre. La envía por correo.

Usa los mismos secretos que el resumen semanal (SUPABASE_URL, SUPABASE_KEY,
TEAM_CODE, MAIL_USER, MAIL_PASS). Destinatario: MAIL_TO_MENSUAL.
Opcionales: MES (AAAA-MM, por defecto el mes en curso), HOY, AUTOR, SOLO_PPT=1,
DEMO_JSON (datos de prueba).
"""
import base64
import io
import os
import smtplib
import sys
from collections import OrderedDict
from datetime import date, datetime, timedelta
from email.message import EmailMessage

from PIL import Image as PILImage
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree

import requests
from zoneinfo import ZoneInfo

# Este script es independiente del resumen semanal: no importa nada de él.
TZ = ZoneInfo("America/Santiago")
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
ESTADOS = ["Excelente", "Bueno", "Regular", "Crítico"]
VARIEDADES_BASE = ["Camarosa", "Camino Real", "Fronteras", "Adelanto", "Merced", "Belvedere", "Castaic",
                   "San Andreas", "Albion", "Monterey", "Cabrillo", "Alturas", "Portola", "Artesia",
                   "Carpinteria", "Sweet Carolina"]


# ----------------------------------------------------------------- datos
def _headers():
    key = os.environ["SUPABASE_KEY"]
    h = {"apikey": key, "x-llahuen-key": os.environ["TEAM_CODE"], "Accept": "application/json"}
    if key.startswith("eyJ"):
        h["Authorization"] = "Bearer " + key
    return h


def _base():
    return os.environ["SUPABASE_URL"].rstrip("/") + "/rest/v1/"


def cargar_docs():
    if os.environ.get("DEMO_JSON"):
        import json
        with open(os.environ["DEMO_JSON"], encoding="utf-8") as f:
            return json.load(f)["visitas"]
    r = requests.get(_base() + "visitas", headers=_headers(), timeout=60,
                     params={"select": "id,data", "order": "actualizado.desc", "limit": "5000"})
    if r.status_code != 200:
        sys.exit(f"Error leyendo visitas de Supabase: HTTP {r.status_code}. "
                 "Revisa SUPABASE_URL, SUPABASE_KEY y TEAM_CODE.")
    return [dict(d["data"] or {}, id=d["id"]) for d in r.json()]


def cargar_fotos(ids):
    out = {}
    if not ids:
        return out
    if os.environ.get("DEMO_JSON"):
        import json
        with open(os.environ["DEMO_JSON"], encoding="utf-8") as f:
            fotos = json.load(f).get("fotos", {})
        return {i: fotos[i] for i in ids if i in fotos}
    ids = list(ids)
    for i in range(0, len(ids), 6):
        parte = ids[i:i + 6]
        try:
            r = requests.get(_base() + "fotos", headers=_headers(), timeout=60,
                             params={"select": "id,data", "id": "in.(" + ",".join(parte) + ")"})
            if r.status_code == 200:
                for row in r.json():
                    out[row["id"]] = row["data"]
        except requests.RequestException:
            pass
    return out


def migrar(v):
    v = dict(v)
    dest = []
    for x in v.get("destacados") or []:
        x = dict(x)
        if x.get("tipo") == "Reclamo":
            x.setdefault("categoria", "Otro")
            x.setdefault("estado", v.get("reclamoEstado") or "Abierto")
        elif x.get("tipo") == "Oportunidad":
            x.setdefault("sub", "Otra")
            x.setdefault("estado", "Detectada")
        dest.append(x)
    v["destacados"] = dest
    v["variedades"] = v.get("variedades") or []
    v["evalVar"] = v.get("evalVar") or {}
    v["fotos"] = v.get("fotos") or []
    return v


def estado_visita(v):
    if v.get("estado") in ESTADOS:
        return v["estado"]
    ests = [e.get("estado") for e in v["evalVar"].values() if e.get("estado") in ESTADOS]
    return min(ests, key=lambda s: -ESTADOS.index(s)) if ests else ""


def variedad_oport(v, d):
    """Variedad de la oportunidad: la marcada en la app o la(s) mencionada(s) en el detalle."""
    if d.get("variedad"):
        return d["variedad"]
    import unicodedata

    def n(t):
        return "".join(c for c in unicodedata.normalize("NFD", str(t or "").lower()) if unicodedata.category(c) != "Mn")
    texto = n(d.get("texto"))
    nombres = list(dict.fromkeys(VARIEDADES_BASE + list(v["variedades"])))
    hall = [x for x in nombres if n(x) and n(x) in texto]
    return ", ".join(hall) if hall else "Sin especificar"


# ----------------------------------------------------------------- estilo
# Identidad de llahuen.com: verde corporativo #096242, fondo crema #FBF3E8,
# rojo frutilla del logo, etiquetas en mayúsculas espaciadas, fotos del vivero.
# Composición inspirada en el deck de Stripe (blanco, títulos grandes, bloques
# en diagonal, íconos en círculos) y gráficos de datos modernos.
from PIL import ImageDraw  # noqa: E402
from pptx.chart.data import CategoryChartData  # noqa: E402
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION  # noqa: E402

VERDE = RGBColor(0x09, 0x62, 0x42)
VERDE_OSC = RGBColor(0x06, 0x45, 0x2E)
VERDE_CLARO = RGBColor(0xE4, 0xF0, 0xE8)
VERDE_MED = RGBColor(0x9C, 0xC9, 0xAE)
ROJO = RGBColor(0xD7, 0x26, 0x2E)
ROSA = RGBColor(0xFB, 0xE3, 0xE4)
CREMA = RGBColor(0xFB, 0xF3, 0xE8)
CREMA_OSC = RGBColor(0xF1, 0xE6, 0xD4)
TEXTO = RGBColor(0x1E, 0x2A, 0x23)
GRIS = RGBColor(0x6B, 0x70, 0x6C)
GRIS_CLARO = RGBColor(0xF5, 0xF6, 0xF3)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)
AMARILLO = RGBColor(0xFB, 0xEF, 0xC9)
CAFE = RGBColor(0x8F, 0x3A, 0x1E)
EST_COL = {"Excelente": RGBColor(0x09, 0x62, 0x42), "Bueno": RGBColor(0x6F, 0xA8, 0x3E),
           "Regular": RGBColor(0xE0, 0x9A, 0x2B), "Crítico": ROJO}
FONT = "Calibri"
W, H = 13.333, 7.5

# Logo y fotos de www.llahuen.com guardados en reporte/assets/ (fijos en el repositorio;
# no se descargan del sitio cada mes).
ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
REGION_ORDEN = ["Arica y Parinacota", "Tarapacá", "Antofagasta", "Atacama", "Coquimbo", "Valparaíso",
                "Metropolitana", "O'Higgins", "Maule", "Ñuble", "Biobío", "La Araucanía", "Los Ríos",
                "Los Lagos", "Aysén", "Magallanes"]
REGION_ZONA = {"Arica y Parinacota": "Norte", "Tarapacá": "Norte", "Antofagasta": "Norte", "Atacama": "Norte",
               "Coquimbo": "Norte", "Valparaíso": "Zona Central", "Metropolitana": "Zona Central",
               "O'Higgins": "Zona Central", "Maule": "Maule", "Ñuble": "Ñuble", "Biobío": "Araucanía",
               "La Araucanía": "Araucanía", "Los Ríos": "Sur", "Los Lagos": "Sur", "Aysén": "Sur", "Magallanes": "Sur"}
ZONAS = ["Norte", "Zona Central", "Maule", "Ñuble", "Araucanía", "Sur"]


def sitio_img(clave):
    """Imagen de marca guardada en reporte/assets (o en SITIO_DIR para pruebas)."""
    carpeta = os.environ.get("SITIO_DIR") or ASSETS
    ruta = os.path.join(carpeta, clave + (".png" if clave == "logo" else ".jpg"))
    try:
        return PILImage.open(ruta)
    except OSError:
        return None


def logo_sitio():
    """Logo de Agrícola Llahuen del sitio (solo la marca, sin el sello de 50 años)."""
    im = sitio_img("logo")
    if im is None:
        return None
    im = im.convert("RGBA")
    bg = PILImage.new("RGBA", im.size, (255, 255, 255, 0))
    w, h = im.size
    # el archivo trae la marca a la izquierda y el sello "50 años" a la derecha
    marca = im.crop((0, 0, int(w * 0.5), h))
    fondo = PILImage.new("RGBA", marca.size, (255, 255, 255, 255))
    plano = PILImage.alpha_composite(fondo, marca).convert("L")
    caja = plano.point(lambda p: 255 if p < 235 else 0).getbbox()
    if caja:
        pad = 12
        marca = marca.crop((max(0, caja[0] - pad), max(0, caja[1] - pad),
                            min(marca.width, caja[2] + pad), min(marca.height, caja[3] + pad)))
    _ = bg
    return png(marca)


def hexrgb(c):
    return (c[0], c[1], c[2]) if isinstance(c, tuple) else tuple(int(str(c)[i:i + 2], 16) for i in (0, 2, 4))


# ----------------------------------------------------------------- utilidades
def fd(iso):
    try:
        d = date.fromisoformat(str(iso)[:10])
        return f"{d.day:02d}/{d.month:02d}/{d.year}"
    except ValueError:
        return str(iso or "")


def fcorta(iso):
    try:
        d = date.fromisoformat(str(iso)[:10])
        return f"{d.day} {MESES[d.month - 1][:3]}"
    except ValueError:
        return str(iso or "")


def corta(t, n):
    t = " ".join(str(t or "").split())
    return t if len(t) <= n else t[: n - 1].rsplit(" ", 1)[0] + "…"


def logo_bytes():
    import re
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "index.html")
    try:
        html = open(ruta, encoding="utf-8").read()
        tag = re.search(r'<img[^>]*\bid="logo"[^>]*>', html).group(0)
        src = re.search(r'src="data:image/[^;]+;base64,([^"]+)"', tag).group(1)
        return base64.b64decode(src)
    except Exception:
        return None


def abrir_foto(data_url):
    try:
        im = PILImage.open(io.BytesIO(base64.b64decode(data_url.split(",", 1)[1])))
        try:
            from PIL import ImageOps
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        return im.convert("RGB")
    except Exception:
        return None


def recorte(im, w, h, lado=900):
    """Recorta la foto al centro con la proporción del marco (sin bordes vacíos)."""
    ar = w / h
    iw, ih = im.size
    if iw / ih > ar:
        nw = int(ih * ar)
        im = im.crop(((iw - nw) // 2, 0, (iw - nw) // 2 + nw, ih))
    else:
        nh = int(iw / ar)
        im = im.crop((0, (ih - nh) // 2, iw, (ih - nh) // 2 + nh))
    im.thumbnail((lado, lado))
    return im


def png(im):
    b = io.BytesIO()
    im.save(b, "PNG")
    return b.getvalue()


def jpg(im):
    b = io.BytesIO()
    im.save(b, "JPEG", quality=78)
    return b.getvalue()


# ----------------------------------------------------------------- ilustraciones (dibujadas en código)
def _frutilla(d, cx, cy, r):
    rojo, verde, semilla = (214, 40, 57), (46, 125, 79), (255, 214, 102)
    pts = []
    import math
    for k in range(80):
        t = 2 * math.pi * k / 80
        x = math.sin(t)
        y = -math.cos(t)
        ancho = 0.82 * (1 - 0.35 * max(0, y)) if y > 0 else 0.92
        pts.append((cx + x * r * ancho * (1 + 0.15 * (y < 0)), cy + y * r * (1.08 if y > 0 else 0.62)))
    d.polygon(pts, fill=rojo)
    for (sx, sy) in [(-.42, -.18), (0, -.22), (.42, -.18), (-.55, .12), (-.18, .1), (.2, .1), (.56, .12),
                     (-.35, .4), (0, .42), (.35, .4), (-.15, .68), (.18, .68)]:
        x, y = cx + sx * r, cy + sy * r
        d.ellipse((x - r * .045, y - r * .07, x + r * .045, y + r * .07), fill=semilla)
    hoja = [(cx, cy - r * .75)]
    for k in range(7):
        ang = math.pi + k * math.pi / 6
        rr = r * (.55 if k % 2 == 0 else .28)
        hoja.append((cx + math.cos(ang) * rr, cy - r * .6 + math.sin(ang) * rr * .5))
    d.polygon(hoja, fill=verde)
    d.rectangle((cx - r * .05, cy - r * 1.05, cx + r * .05, cy - r * .6), fill=verde)


def _flor(d, cx, cy, r, petalo=(255, 255, 255)):
    import math
    for k in range(5):
        ang = -math.pi / 2 + k * 2 * math.pi / 5
        px, py = cx + math.cos(ang) * r * .5, cy + math.sin(ang) * r * .5
        d.ellipse((px - r * .42, py - r * .42, px + r * .42, py + r * .42), fill=petalo)
    d.ellipse((cx - r * .3, cy - r * .3, cx + r * .3, cy + r * .3), fill=(250, 196, 54))
    for k in range(10):
        ang = k * 2 * math.pi / 10
        px, py = cx + math.cos(ang) * r * .2, cy + math.sin(ang) * r * .2
        d.ellipse((px - r * .04, py - r * .04, px + r * .04, py + r * .04), fill=(214, 150, 30))


def _hoja(img, cx, cy, r, color=(46, 125, 79), vena=(30, 90, 55)):
    import math
    S = img.size[0]
    for ang in (-50, 0, 50):
        capa = PILImage.new("RGBA", (S, S), (0, 0, 0, 0))
        dc = ImageDraw.Draw(capa)
        dc.ellipse((cx - r * .3, cy - r * .95, cx + r * .3, cy - r * .05), fill=color)
        dc.line((cx, cy - r * .9, cx, cy - r * .1), fill=vena, width=max(2, int(r * .04)))
        for k in range(3):
            yy = cy - r * (.75 - k * .2)
            dc.line((cx, yy, cx - r * .18, yy - r * .1), fill=vena, width=max(1, int(r * .025)))
            dc.line((cx, yy, cx + r * .18, yy - r * .1), fill=vena, width=max(1, int(r * .025)))
        capa = capa.rotate(-ang, center=(cx, cy), resample=PILImage.BICUBIC)
        img.alpha_composite(capa)
    ImageDraw.Draw(img).line((cx, cy, cx, cy + r * .5), fill=color, width=max(3, int(r * .07)))
    _ = math


def icono(tipo, fondo=(255, 255, 255), lado=400):
    """Ícono redondo con una frutilla, flor u hoja (PNG con transparencia)."""
    S = lado * 3
    img = PILImage.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    if fondo is not None:
        d.ellipse((0, 0, S - 1, S - 1), fill=fondo)
    c, r = S / 2, S * .3
    if tipo == "frutilla":
        _frutilla(d, c, c + r * .12, r)
    elif tipo == "flor":
        _flor(d, c, c, r * 1.05, (255, 255, 255) if fondo != (255, 255, 255) else (255, 235, 238))
        if fondo == (255, 255, 255):
            pass
    else:
        _hoja(img, c, c + r * .35, r * 1.15)
    return png(img.resize((lado, lado), PILImage.LANCZOS))


def foto_diagonal(im, w_in, h_in, corte=0.28):
    """Foto con el borde izquierdo en diagonal (como los bloques del deck de Stripe)."""
    im = recorte(im, w_in, h_in, 1400).convert("RGBA")
    mask = PILImage.new("L", im.size, 0)
    iw, ih = im.size
    ImageDraw.Draw(mask).polygon([(iw * corte, 0), (iw, 0), (iw, ih), (0, ih)], fill=255)
    im.putalpha(mask)
    return png(im)


# ----------------------------------------------------------------- dibujo
class Deck:
    def __init__(self, logo):
        self.p = Presentation()
        self.p.slide_width, self.p.slide_height = Inches(W), Inches(H)
        self.blank = self.p.slide_layouts[6]
        self.n = 0
        self.fotos = {}
        self.logo = logo
        self._cache = {}

    def foto(self, fid):
        if fid not in self._cache:
            self._cache[fid] = abrir_foto(self.fotos[fid]) if fid in self.fotos else None
        return self._cache[fid]

    def slide(self, fondo=BLANCO):
        s = self.p.slides.add_slide(self.blank)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = fondo
        self.n += 1
        return s

    def texto(self, s, x, y, w, h, partes, size, color=TEXTO, bold=False, align=PP_ALIGN.LEFT,
              anchor=MSO_ANCHOR.TOP, espacio=0):
        """partes: str, lista de párrafos, o párrafos como listas de (texto, color, bold)."""
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, 0)
        parrafos = partes if isinstance(partes, list) else [partes]
        for i, par in enumerate(parrafos):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            if espacio and i:
                p.space_before = Pt(espacio)
            runs = par if isinstance(par, list) else [(par, color, bold)]
            for txt, col, b in runs:
                r = p.add_run()
                r.text = txt
                r.font.size, r.font.bold, r.font.name = Pt(size), b, FONT
                r.font.color.rgb = col
        return tb

    def rect(self, s, x, y, w, h, color, forma=MSO_SHAPE.RECTANGLE, radio=None):
        r = s.shapes.add_shape(forma, Inches(x), Inches(y), Inches(w), Inches(h))
        r.fill.solid()
        r.fill.fore_color.rgb = color
        r.line.fill.background()
        r.shadow.inherit = False
        if radio is not None and forma == MSO_SHAPE.ROUNDED_RECTANGLE:
            r.adjustments[0] = radio
        return r

    def poligono(self, s, puntos, color):
        fb = s.shapes.build_freeform(Inches(puntos[0][0]), Inches(puntos[0][1]), scale=1.0)
        fb.add_line_segments([(Inches(x), Inches(y)) for x, y in puntos[1:]], close=True)
        sh = fb.convert_to_shape()
        sh.fill.solid()
        sh.fill.fore_color.rgb = color
        sh.line.fill.background()
        sh.shadow.inherit = False
        return sh

    def imagen(self, s, data, x, y, w=None, h=None):
        kw = {}
        if w:
            kw["width"] = Inches(w)
        if h:
            kw["height"] = Inches(h)
        return s.shapes.add_picture(io.BytesIO(data), Inches(x), Inches(y), **kw)

    def icono(self, s, tipo, x, y, d, fondo=(255, 255, 255)):
        return self.imagen(s, icono(tipo, fondo), x, y, d, d)

    def marco(self, s, titulo, sub="", titulo_extra=None):
        """Diapositiva de contenido: título verde arriba a la izquierda, logo abajo."""
        largo = len(titulo) + len(titulo_extra or "")
        size = 30 if largo <= 44 else (25 if largo <= 56 else 21)
        partes = [[(titulo, VERDE, True)] + ([(titulo_extra, ROJO, True)] if titulo_extra else [])]
        self.texto(s, 0.6, 0.42, 11.6, 0.7, partes, size)
        if sub:
            self.texto(s, 0.6, 1.08, 10.5, 0.35, sub, 13, GRIS)
        if self.logo:
            self.imagen(s, self.logo, 0.6, 6.82, h=0.46)
        self.texto(s, 11.9, 6.95, 0.83, 0.3, str(self.n), 10, GRIS, align=PP_ALIGN.RIGHT)

    def pastilla(self, s, x, y, estado, w=1.05):
        if estado not in EST_COL:
            return
        r = self.rect(s, x, y, w, 0.3, EST_COL[estado], MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
        tf = r.text_frame
        tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = estado
        run.font.size, run.font.bold, run.font.name = Pt(10.5), True, FONT
        run.font.color.rgb = BLANCO

    def chip(self, s, x, y, texto, fondo=ROSA, color=ROJO, size=10):
        w = min(3.4, 0.25 + 0.075 * len(texto) * size / 10)
        r = self.rect(s, x, y, w, 0.28, fondo, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
        tf = r.text_frame
        tf.margin_top = tf.margin_bottom = 0
        tf.margin_left = tf.margin_right = Inches(0.06)
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = texto
        run.font.size, run.font.bold, run.font.name = Pt(size), True, FONT
        run.font.color.rgb = color
        return w

    def foto_redonda(self, s, fid, x, y, w, h, radio=6000):
        im = self.foto(fid)
        if im is None:
            self.rect(s, x, y, w, h, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE, 0.08)
            return False
        pic = self.imagen(s, jpg(recorte(im, w, h)), x, y, w, h)
        pic.auto_shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
        geom = pic._element.spPr.find(qn("a:prstGeom"))
        av = geom.find(qn("a:avLst"))
        if av is None:
            av = etree.SubElement(geom, qn("a:avLst"))
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", "adj")
        gd.set("fmla", f"val {radio}")
        return True

    def grilla(self, s, items, x0, y0, cols, w, h, gx, gy, cap_h, sub=False):
        for i, it in enumerate(items):
            fila, col = divmod(i, cols)
            x = x0 + col * (w + gx)
            y = y0 + fila * (h + cap_h + gy)
            self.foto_redonda(s, it["id"], x, y, w, h)
            pars = [[(it["cap"], TEXTO, True)]]
            if sub and it.get("sub"):
                pars.append([(it["sub"], GRIS, False)])
            self.texto(s, x + 0.02, y + h + 0.07, w - 0.04, cap_h - 0.07, pars, 9.5)

    def tabla(self, s, x, y, w, filas, ancho_et=1.05, size=9):
        t = s.shapes.add_table(len(filas), 2, Inches(x), Inches(y), Inches(w), Inches(0.24 * len(filas))).table
        tbl = t._tbl
        tblPr = tbl.tblPr
        style = tblPr.find(qn("a:tableStyleId"))
        if style is not None:
            style.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # sin estilo, sin bordes
        t.columns[0].width = Inches(ancho_et)
        t.columns[1].width = Inches(w - ancho_et)
        t.first_row = False
        t.horz_banding = False
        colores = {"Oportunidad": (AMARILLO, RGBColor(0x8A, 0x5A, 0x00)), "Reclamos": (ROSA, ROJO)}
        for i, (et, val) in enumerate(filas):
            fondo, col = colores.get(et, (VERDE_CLARO, VERDE))
            for j, (txt, f, c, b) in enumerate([(et, fondo, col, True), (val, BLANCO, TEXTO, False)]):
                cel = t.cell(i, j)
                cel.fill.solid()
                cel.fill.fore_color.rgb = f
                cel.margin_left = cel.margin_right = Inches(0.07)
                cel.margin_top = cel.margin_bottom = Inches(0.035)
                cel.vertical_anchor = MSO_ANCHOR.TOP
                tf = cel.text_frame
                tf.word_wrap = True
                r = tf.paragraphs[0].add_run()
                r.text = txt
                r.font.size, r.font.bold, r.font.name = Pt(size), b, FONT
                r.font.color.rgb = c
        return t

    def racimo(self, s, x0, y0, escala=1.0, fondo_claro=True):
        """Grupo de íconos flotantes (frutillas, flores y hojas), como el deck de Stripe."""
        blanco = (255, 255, 255)
        verde, rojo, rosa = (230, 240, 232), (251, 227, 230), (42, 104, 70)
        piezas = [("frutilla", 1.55, 1.2, 0.9, blanco), ("flor", 0.75, 0.25, 0.8, (214, 40, 57)),
                  ("hoja", 0.85, 2.9, 0.75, blanco if not fondo_claro else verde), ("flor", 3.0, 0.6, 0.65, rosa),
                  ("frutilla", 3.1, 2.25, 0.62, rojo if fondo_claro else blanco), ("hoja", 2.35, 0.0, 0.5, verde),
                  ("flor", 2.45, 3.35, 0.45, (214, 40, 57)), ("frutilla", 0.1, 1.75, 0.45, blanco)]
        for tipo, dx, dy, d, fondo in piezas:
            self.icono(s, tipo, x0 + dx * escala, y0 + dy * escala, d * escala * 1.3, fondo)


    # ---- detalles de marca y gráficos
    def etiqueta(self, s, x, y, w, texto, color=VERDE, size=10.5, align=PP_ALIGN.LEFT):
        """Etiqueta en mayúsculas espaciadas, como en llahuen.com."""
        tb = self.texto(s, x, y, w, 0.3, texto.upper(), size, color, True, align)
        for r in tb.text_frame.paragraphs[0].runs:
            r.font._rPr.set("spc", "250")
        return tb

    def transparencia(self, shape, alpha=0.75):
        clr = shape._element.spPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        etree.SubElement(clr, qn("a:alpha"), val=str(int(alpha * 100000)))

    def foto_fondo(self, s, im, oscuro=0.62, color=VERDE_OSC):
        """Foto a pantalla completa con un velo verde para que el texto se lea."""
        if im is not None:
            self.imagen(s, jpg(recorte(im.convert("RGB"), W, H, 1600)), 0, 0, W, H)
        velo = self.rect(s, 0, 0, W, H, color)
        if im is not None:
            self.transparencia(velo, oscuro)

    def _ejes_limpios(self, chart, size=10):
        chart.has_title = False
        chart.font.size = Pt(size)
        chart.font.name = FONT
        chart.font.color.rgb = TEXTO

    def dona(self, s, x, y, d, cats, vals, colores, centro, sub):
        cd = CategoryChartData()
        cd.categories = cats
        cd.add_series("Visitas", vals)
        gf = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(x), Inches(y), Inches(d), Inches(d), cd)
        ch = gf.chart
        self._ejes_limpios(ch, 11)
        ch.has_legend = False
        plot = ch.plots[0]
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.show_value, dl.show_category_name, dl.show_percentage = True, False, False
        dl.font.size, dl.font.bold = Pt(12), True
        dl.font.color.rgb = BLANCO
        for i, pt in enumerate(plot.series[0].points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = colores[i]
            pt.format.line.color.rgb = BLANCO
            pt.format.line.width = Pt(2)
        hole = ch._chartSpace.find(".//" + qn("c:holeSize"))
        if hole is not None:
            hole.set("val", "62")
        self.texto(s, x, y + d / 2 - 0.45, d, 0.6, str(centro), 34, VERDE, True, PP_ALIGN.CENTER)
        self.texto(s, x, y + d / 2 + 0.12, d, 0.3, sub, 11, GRIS, False, PP_ALIGN.CENTER)
        return ch

    def leyenda(self, s, x, y, items, size=11, paso=0.36):
        for i, (txt, col) in enumerate(items):
            self.rect(s, x, y + i * paso + 0.05, 0.18, 0.18, col, MSO_SHAPE.OVAL)
            self.texto(s, x + 0.3, y + i * paso, 3.2, 0.3, txt, size, TEXTO)

    def barras_apiladas(self, s, x, y, w, h, cats, series, colores, leyenda=True):
        cd = CategoryChartData()
        cd.categories = cats
        for nombre, vals in series:
            cd.add_series(nombre, vals)
        gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = gf.chart
        self._ejes_limpios(ch, 10.5)
        plot = ch.plots[0]
        plot.gap_width = 55
        plot.overlap = 100
        for i, ser in enumerate(plot.series):
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = colores[i]
            ser.format.line.color.rgb = BLANCO
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(10)
        plot.data_labels.font.bold = True
        plot.data_labels.font.color.rgb = BLANCO
        plot.data_labels.position = XL_LABEL_POSITION.CENTER
        va = ch.value_axis
        va.visible = False
        va.has_major_gridlines = False
        ca = ch.category_axis
        ca.format.line.fill.background()
        ca.tick_labels.font.size = Pt(10.5)
        ca.reverse_order = True
        ch.has_legend = leyenda
        if leyenda:
            ch.legend.position = XL_LEGEND_POSITION.TOP
            ch.legend.include_in_layout = False
            ch.legend.font.size = Pt(10)
        return ch

    def barras(self, s, x, y, w, h, cats, vals, color):
        cd = CategoryChartData()
        cd.categories = cats
        cd.add_series("Cantidad", vals)
        gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = gf.chart
        self._ejes_limpios(ch, 10.5)
        ch.has_legend = False
        plot = ch.plots[0]
        plot.gap_width = 70
        ser = plot.series[0]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = color
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(11)
        plot.data_labels.font.bold = True
        plot.data_labels.font.color.rgb = color
        plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        ch.value_axis.visible = False
        ch.value_axis.has_major_gridlines = False
        ch.category_axis.format.line.fill.background()
        ch.category_axis.reverse_order = True
        return ch


# ----------------------------------------------------------------- contenido
def caption(f):
    return " · ".join(x for x in [f.get("cat") or "Foto", f.get("variedad")] if x)


def es_nueva(v, n):
    return bool((v["evalVar"].get(n) or {}).get("nueva"))


def ubic(v):
    partes = [fd(v.get("fecha")), ", ".join(x for x in [v.get("comuna"), v.get("region")] if x), v.get("tipo")]
    if v.get("superficie"):
        partes.append(f"{str(v['superficie']).replace('.', ',')} ha")
    return " · ".join(p for p in partes if p)


def fplant(v, n):
    e = v["evalVar"].get(n) or {}
    return sorted({l.get("fplant") for l in (e.get("lotes") or {}).values() if l.get("fplant")})


def filas_visita(v, largo=1.0):
    n = lambda k: int(k * largo)  # noqa: E731
    dat = " · ".join([v.get("predio") or "Predio no indicado", v.get("etapa") or "etapa no indicada",
                      "Var.: " + (", ".join(v["variedades"]) or "no indicada")])
    fito = (v["fito"] + (": " + v["fitoDetalle"] if v.get("fitoDetalle") else "")) if v.get("fito") \
        else "Sin información sanitaria registrada."
    notas = " ".join(x for x in [v.get("notas"), ("Manejo: " + v["manejo"]) if v.get("manejo") else ""] if x)
    hall = []
    for nom in v["variedades"]:
        e = v["evalVar"].get(nom) or {}
        fp = fplant(v, nom)
        meta = ", ".join(x for x in [e.get("estado"), ("plant. " + ", ".join(fd(x) for x in fp)) if fp else ""] if x)
        txt = nom + (f" ({meta})" if meta else "")
        if e.get("resultado"):
            txt += ": " + e["resultado"]
        hall.append(txt)
    op = [(d.get("texto") or "") + f" [Var.: {variedad_oport(v, d)}]" for d in v["destacados"] if d.get("tipo") == "Oportunidad"]
    rec = [(d.get("texto") or "") + f" ({d.get('estado', 'Abierto')})" for d in v["destacados"] if d.get("tipo") == "Reclamo"]
    return [("Datos", corta(dat, n(110))), ("Fitosanidad", corta(fito, n(110))),
            ("Notas", corta(notas, n(210)) or "Sin notas registradas."),
            ("Hallazgos", corta(" · ".join(hall), n(210)) or "Sin hallazgos registrados."),
            ("Oportunidad", corta(" · ".join(op), n(120)) or "Sin oportunidad registrada."),
            ("Reclamos", corta(" · ".join(rec), n(120)) or "Sin reclamos.")]


def problema(v):
    return (estado_visita(v) in ("Regular", "Crítico") or v.get("fito") == "Problema"
            or any(d.get("tipo") == "Reclamo" and d.get("estado") != "Resuelto" for d in v["destacados"]))


def construir(mes, hoy, autor):
    ini = mes
    fin = (mes.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    corte = min(fin, hoy)
    docs = [migrar(d) for d in cargar_docs() if not d.get("ejemplo")]
    todos = sorted([v for v in docs if not v.get("prog") and ini.isoformat() <= str(v.get("fecha", ""))[:10] <= corte.isoformat()],
                   key=lambda v: (v.get("fecha", ""), v.get("creado") or 0))
    # Reclamos, riesgos u oportunidades agregados en la app sin visita presencial ("caso"):
    # no cuentan como visitas, pero sus registros sí se informan.
    vis = [v for v in todos if not v.get("caso")]
    casos = [v for v in todos if v.get("caso")]
    agenda = sorted([a for a in docs if a.get("prog") and a.get("estado") == "pendiente"
                     and hoy.isoformat() < str(a.get("fecha", ""))[:10] <= (hoy + timedelta(days=28)).isoformat()],
                    key=lambda a: (a.get("fecha", ""), a.get("hora", "")))
    ops = [(v, d) for v in todos for d in v["destacados"] if d.get("tipo") == "Oportunidad"]
    recs = [(v, d) for v in todos for d in v["destacados"] if d.get("tipo") == "Reclamo"]
    nuevas_tot = sorted({n for v in vis for n in v["variedades"] if es_nueva(v, n)})
    mes_txt = MESES[ini.month - 1]
    rango = f"{ini.day} al {corte.day} de {mes_txt} de {ini.year}"

    def fotos_cat(palabras, n):
        out, vistos = [], set()
        for v in vis:
            if v.get("cliente") in vistos:
                continue
            for f in v["fotos"]:
                if f.get("id") and any(p in (f.get("cat") or "").lower() for p in palabras):
                    out.append({"id": f["id"], "cap": caption(f), "sub": v.get("cliente", "")})
                    vistos.add(v.get("cliente"))
                    break
        return out[:n]
    predios = fotos_cat(["general predio"], 8)
    plantas = fotos_cat(["apariencia planta", "planta sana"], 4)
    frutas = fotos_cat(["fruta"], 4)
    ids = {x["id"] for x in predios + plantas + frutas}
    for v in vis:
        ids |= {f["id"] for f in v["fotos"] if f.get("id")}

    logo = logo_sitio() or logo_bytes()
    img_campo, img_plantas, img_atardecer = sitio_img("campo"), sitio_img("plantas"), sitio_img("atardecer")
    D = Deck(logo)
    D.fotos = cargar_fotos(ids)
    ROSA_TXT = RGBColor(0xFF, 0xC9, 0xCC)

    # 1. portada: foto del vivero (llahuen.com) con velo verde
    s = D.slide(VERDE)
    D.foto_fondo(s, img_campo, 0.70)
    if logo:
        D.rect(s, 0.6, 0.6, 2.6, 1.75, CREMA, MSO_SHAPE.ROUNDED_RECTANGLE, 0.08)
        D.imagen(s, logo, 0.78, 0.72, h=1.5)
    D.racimo(s, 9.4, 0.55, 0.78, fondo_claro=False)
    D.etiqueta(s, 0.6, 3.0, 8, f"Reporte mensual · {mes_txt} {ini.year}", ROSA_TXT, 12)
    D.texto(s, 0.6, 3.4, 8.6, 1.9, "Postventa y atención a clientes estratégicos", 42, BLANCO, True)
    D.texto(s, 0.6, 5.35, 8.6, 0.5, "Lo que vimos en terreno este mes, desde la app Visitas Llahuen", 17, VERDE_MED)
    D.texto(s, 0.6, 6.4, 8.6, 0.7, [f"{rango}", f"{autor} · {fd(hoy.isoformat())}"], 12, VERDE_MED)

    # 2. el mes en cifras: números grandes con etiqueta espaciada (estilo llahuen.com) y foto en diagonal
    s = D.slide(CREMA)
    D.marco(s, "El mes en cifras", rango)
    stats = [(len(vis), "visitas", VERDE), (len({v.get('cliente') for v in vis}), "productores", VERDE),
             (len({v.get('region') for v in vis if v.get('region')}), "regiones", VERDE),
             (len(recs), "reclamos", ROJO), (len(ops), "oportunidades", ROJO), (len(nuevas_tot), "var. nuevas", ROJO)]
    for i, (num, et, col) in enumerate(stats):
        f, c = divmod(i, 3)
        x, y = 0.6 + c * 2.35, 1.85 + f * 1.95
        D.texto(s, x, y, 2.2, 1.0, str(num), 60, col, True)
        D.etiqueta(s, x + 0.03, y + 1.08, 2.25, et, GRIS, 10.5)
    ha = sum(float(v.get("superficie") or 0) for v in vis)
    if ha:
        D.texto(s, 0.6, 5.85, 6.6, 0.4, [[("Superficie visitada: ", VERDE, True), (f"{ha:,.1f} ha".replace(",", "X").replace(".", ",").replace("X", "."), TEXTO, False)]], 12)
    foto_ancla = img_plantas or next((D.foto(p["id"]) for p in predios if D.foto(p["id"]) is not None), None)
    if foto_ancla is not None:
        D.imagen(s, foto_diagonal(foto_ancla.convert("RGB"), 5.6, H), W - 5.6, 0, 5.6, H)
    else:
        D.poligono(s, [(W - 5.6 + 1.6, 0), (W, 0), (W, H), (W - 5.6, H)], VERDE)
        D.racimo(s, W - 4.1, 1.6, 0.85, fondo_claro=False)

    if not vis:
        s = D.slide(VERDE_OSC)
        D.texto(s, 1.0, 2.8, 11.3, 1.5, "No hay visitas registradas en la app para este mes.", 36, BLANCO, True, PP_ALIGN.CENTER)

    if vis:
        # 3. dónde estuvimos: Chile de norte a sur con burbujas por región + barras por zona
        s = D.slide()
        D.marco(s, "Dónde estuvimos", "Visitas por región, de norte a sur")
        por_reg = {}
        for v in vis:
            por_reg[v.get("region") or "Sin región"] = por_reg.get(v.get("region") or "Sin región", 0) + 1
        top, bot = 1.6, 6.55
        paso = (bot - top) / (len(REGION_ORDEN) - 1)
        D.rect(s, 1.18, top, 0.06, bot - top, CREMA_OSC, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
        mx = max(por_reg.values())
        for i, reg in enumerate(REGION_ORDEN):
            y = top + i * paso
            n = por_reg.get(reg, 0)
            d = 0.14 if not n else 0.3
            col = VERDE if n else CREMA_OSC
            D.rect(s, 1.21 - d / 2, y - d / 2, d, d, col, MSO_SHAPE.OVAL)
            if n:
                D.texto(s, 1.21 - d / 2, y - 0.13, d, 0.26, str(n), 10.5, BLANCO, True, PP_ALIGN.CENTER,
                        MSO_ANCHOR.MIDDLE)
            D.texto(s, 1.75, y - 0.13, 2.6, 0.26, reg, 11 if n else 9.5, TEXTO if n else GRIS, bool(n), anchor=MSO_ANCHOR.MIDDLE)
        # tarjetas por zona comercial
        por_zona = {z: 0 for z in ZONAS}
        prod_zona = {z: set() for z in ZONAS}
        for v in vis:
            z = REGION_ZONA.get(v.get("region"))
            if z:
                por_zona[z] += 1
                prod_zona[z].add(v.get("cliente"))
        mz = max(por_zona.values()) or 1
        D.etiqueta(s, 5.0, 1.6, 7, "Visitas por zona comercial", VERDE, 11)
        for i, z in enumerate(ZONAS):
            y = 2.05 + i * 0.73
            D.texto(s, 5.0, y + 0.08, 1.6, 0.35, z, 12.5, TEXTO, True)
            D.rect(s, 6.7, y + 0.1, 4.6, 0.36, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
            if por_zona[z]:
                D.rect(s, 6.7, y + 0.1, max(0.36, 4.6 * por_zona[z] / mz), 0.36, VERDE if i % 2 == 0 else ROJO,
                       MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
            D.texto(s, 11.45, y + 0.04, 1.6, 0.45, [[(str(por_zona[z]), VERDE, True),
                                                     (f" visitas · {len(prod_zona[z])} prod.", GRIS, False)]], 11, anchor=MSO_ANCHOR.MIDDLE)
        D.icono(s, "frutilla", 11.85, 0.42, 0.85, (251, 227, 228))

        # 4. calendario del mes (mapa de calor) + semana más activa
        s = D.slide(CREMA)
        D.marco(s, "Ritmo del mes", "Visitas por día: más verde, más visitas")
        por_dia = {}
        for v in vis:
            por_dia[str(v.get("fecha"))[:10]] = por_dia.get(str(v.get("fecha"))[:10], 0) + 1
        escala = [CREMA_OSC, RGBColor(0xC6, 0xE0, 0xCD), RGBColor(0x6E, 0xB0, 0x86), VERDE]
        x0, y0, c = 0.6, 1.95, 0.6
        for j, dname in enumerate(["L", "M", "M", "J", "V", "S", "D"]):
            D.texto(s, x0 + j * (c + 0.08), y0 - 0.38, c, 0.3, dname, 11, GRIS, True, PP_ALIGN.CENTER)
        dia = ini
        fila = 0
        while dia <= fin:
            col_i = dia.weekday()
            n = por_dia.get(dia.isoformat(), 0)
            futuro = dia > corte
            color = BLANCO if futuro else escala[min(n, 3)]
            x, y = x0 + col_i * (c + 0.08), y0 + fila * (c + 0.08)
            D.rect(s, x, y, c, c, color, MSO_SHAPE.ROUNDED_RECTANGLE, 0.18)
            D.texto(s, x + 0.06, y + 0.04, c - 0.1, 0.22, str(dia.day), 8.5, BLANCO if n >= 2 else GRIS, n >= 2)
            if n:
                D.texto(s, x, y + 0.22, c, 0.36, str(n), 14, BLANCO if n >= 2 else VERDE, True, PP_ALIGN.CENTER)
            if col_i == 6:
                fila += 1
            dia += timedelta(days=1)
        for k, (txt, colr) in enumerate([("sin visitas", CREMA_OSC), ("1", escala[1]), ("2", escala[2]), ("3 o más", VERDE)]):
            D.rect(s, x0 + k * 1.3, 6.42, 0.22, 0.22, colr, MSO_SHAPE.ROUNDED_RECTANGLE, 0.2)
            D.texto(s, x0 + k * 1.3 + 0.3, 6.39, 1.0, 0.3, txt, 10, GRIS)
        # columnas por semana (gráfico nativo)
        semanas = OrderedDict()
        d = ini - timedelta(days=ini.weekday())
        while d <= corte:
            semanas[d] = sum(n for k, n in por_dia.items() if d.isoformat() <= k <= (d + timedelta(days=6)).isoformat())
            d += timedelta(days=7)
        cd = CategoryChartData()
        cd.categories = [f"S{k.isocalendar()[1]}" for k in semanas]
        cd.add_series("Visitas", list(semanas.values()))
        gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(7.1), Inches(2.0), Inches(5.6), Inches(3.6), cd)
        ch = gf.chart
        D._ejes_limpios(ch, 11)
        ch.has_legend = False
        pl = ch.plots[0]
        pl.gap_width = 45
        for i, pt in enumerate(pl.series[0].points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = ROJO if list(semanas.values())[i] == max(semanas.values()) else VERDE
        pl.has_data_labels = True
        pl.data_labels.font.size, pl.data_labels.font.bold = Pt(13), True
        pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        ch.value_axis.visible = False
        ch.value_axis.has_major_gridlines = False
        ch.category_axis.format.line.color.rgb = CREMA_OSC
        D.etiqueta(s, 7.2, 1.6, 5.4, "Visitas por semana", VERDE, 11)
        smax = max(semanas, key=lambda k: semanas[k])
        D.texto(s, 7.2, 5.75, 5.4, 0.6, [[("Semana más activa: ", GRIS, False),
                                          (f"{fcorta(smax.isoformat())} – {fcorta((smax + timedelta(days=6)).isoformat())} · {semanas[smax]} visitas", ROJO, True)]], 12)

        # 5. estado de las plantaciones: dona + variedades apiladas + fitosanidad
        s = D.slide()
        D.marco(s, "Estado de las plantaciones", "Evaluación general de cada visita y de cada variedad")
        ests = [e for e in ESTADOS]
        cnt = [sum(1 for v in vis if estado_visita(v) == e) for e in ests]
        con = [(e, n) for e, n in zip(ests, cnt) if n]
        if con:
            D.dona(s, 0.5, 1.55, 3.6, [e for e, _ in con], [n for _, n in con], [EST_COL[e] for e, _ in con],
                   sum(n for _, n in con), "visitas evaluadas")
            D.leyenda(s, 0.95, 5.35, [(f"{e} · {n}", EST_COL[e]) for e, n in con], 11, 0.33)
        var_est = {}
        for v in vis:
            for n_, e in v["evalVar"].items():
                if e.get("estado") in ESTADOS:
                    var_est.setdefault(n_, {k: 0 for k in ESTADOS})[e["estado"]] += 1
        if var_est:
            top_var = sorted(var_est, key=lambda k: -sum(var_est[k].values()))[:8]
            usados = [e for e in ESTADOS if any(var_est[k][e] for k in top_var)]
            D.etiqueta(s, 4.6, 1.6, 5, "Estado por variedad (evaluaciones)", VERDE, 11)
            D.barras_apiladas(s, 4.5, 1.9, 5.0, 4.7, top_var, [(e, [var_est[k][e] or None for k in top_var]) for e in usados],
                              [EST_COL[e] for e in usados])
        fito = {"Sano": 0, "Observación": 0, "Problema": 0}
        for v in vis:
            if v.get("fito") in fito:
                fito[v["fito"]] += 1
        tot = sum(fito.values())
        D.etiqueta(s, 9.95, 1.6, 3, "Fitosanidad", VERDE, 11)
        if tot:
            colf = {"Sano": VERDE, "Observación": RGBColor(0xE0, 0x9A, 0x2B), "Problema": ROJO}
            for i, (k, n) in enumerate(fito.items()):
                y = 2.05 + i * 1.45
                pct = round(100 * n / tot)
                D.texto(s, 9.95, y, 2.8, 0.6, f"{pct}%", 34, colf[k], True)
                D.texto(s, 9.95, y + 0.62, 2.8, 0.3, f"{k} · {n} visitas", 11, GRIS)
                D.rect(s, 9.95, y + 0.98, 2.75, 0.12, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
                if n:
                    D.rect(s, 9.95, y + 0.98, max(0.12, 2.75 * n / tot), 0.12, colf[k], MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
        else:
            D.texto(s, 9.95, 2.1, 2.8, 0.6, "Sin registros de fitosanidad.", 11, GRIS)

        # 6. reclamos y oportunidades en números
        s = D.slide(CREMA)
        D.marco(s, "Reclamos y oportunidades", "Lo que pide atención y lo que abre negocio")
        D.icono(s, "frutilla", 0.6, 1.65, 0.6, (251, 227, 228))
        D.texto(s, 1.35, 1.7, 4.5, 0.5, [[(str(len(recs)), ROJO, True), ("  reclamos", TEXTO, True)]], 22)
        if recs:
            cats = {}
            for _, d in recs:
                cats[d.get("categoria") or "Otro"] = cats.get(d.get("categoria") or "Otro", 0) + 1
            est_r = {k: sum(1 for _, d in recs if (d.get("estado") or "Abierto") == k) for k in ("Abierto", "En gestión", "Resuelto")}
            D.texto(s, 1.35, 2.25, 5, 0.35, " · ".join(f"{k}: {n}" for k, n in est_r.items()), 11.5, GRIS)
            orden = sorted(cats, key=lambda k: -cats[k])[:6]
            D.barras(s, 0.5, 2.7, 5.7, min(3.9, 0.55 * len(orden) + 0.5), orden, [cats[k] for k in orden], ROJO)
        else:
            D.texto(s, 1.35, 2.4, 5, 0.4, "Sin reclamos registrados este mes.", 13, GRIS)
        D.rect(s, 6.62, 1.65, 0.02, 4.9, CREMA_OSC)
        D.icono(s, "flor", 6.95, 1.65, 0.6, (9, 98, 66))
        D.texto(s, 7.7, 1.7, 5, 0.5, [[(str(len(ops)), VERDE, True), ("  oportunidades", TEXTO, True)]], 22)
        if ops:
            tipos = {}
            for _, d in ops:
                tipos[d.get("sub") or "Otra"] = tipos.get(d.get("sub") or "Otra", 0) + 1
            vars_op = {}
            for v, d in ops:
                for x in [z.strip() for z in variedad_oport(v, d).split(",")]:
                    if x and x != "Sin especificar":
                        vars_op[x] = vars_op.get(x, 0) + 1
            if vars_op:
                D.texto(s, 7.7, 2.25, 5.2, 0.35, "Variedades: " + ", ".join(f"{k} ({n})" for k, n in sorted(vars_op.items(), key=lambda t: -t[1])[:5]), 11.5, GRIS)
            orden = sorted(tipos, key=lambda k: -tipos[k])[:6]
            D.barras(s, 6.85, 2.7, 5.9, min(3.9, 0.55 * len(orden) + 0.5), orden, [tipos[k] for k in orden], VERDE)
        else:
            D.texto(s, 7.7, 2.4, 5, 0.4, "Sin oportunidades registradas este mes.", 13, GRIS)

    # 3. apariencia general de los predios
    if predios:
        s = D.slide()
        D.marco(s, "Apariencia general de los predios", "Una fotografía por productor visitado")
        D.grilla(s, predios, 0.6, 1.6, 4, 2.88, 1.95, 0.2, 0.08, 0.55, sub=True)
    # 4. planta y fruto
    if plantas or frutas:
        s = D.slide()
        D.marco(s, "Apariencia de planta y de fruto", "Cómo se ven las plantaciones este mes")
        y = 1.55
        for nombre, items, tipo in (("Planta", plantas, "hoja"), ("Fruto", frutas, "frutilla")):
            if not items:
                continue
            D.icono(s, tipo, 0.6, y, 0.42, (230, 240, 232))
            D.texto(s, 1.12, y + 0.06, 4, 0.35, nombre, 14, VERDE, True)
            D.grilla(s, items, 0.6, y + 0.52, 4, 2.88, 1.5, 0.2, 0, 0.5, sub=True)
            y += 2.6

    # 5. visitas
    def tarjeta(s, v, x0, y0, w):
        D.rect(s, x0, y0, w, 5.15, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE, 0.04)
        D.texto(s, x0 + 0.22, y0 + 0.2, w - 1.6, 0.4, corta(v.get("cliente"), 40), 15, TEXTO, True)
        D.pastilla(s, x0 + w - 1.27, y0 + 0.22, estado_visita(v))
        D.texto(s, x0 + 0.22, y0 + 0.58, w - 0.44, 0.3, ubic(v), 10, GRIS)
        fs = sorted([f for f in v["fotos"] if f.get("id")], key=lambda f: (not f.get("desc"), not f.get("cat")))[:3]
        fw = (w - 0.44 - 0.24) / 3
        if fs:
            for i, f in enumerate(fs):
                x = x0 + 0.22 + i * (fw + 0.12)
                D.foto_redonda(s, f["id"], x, y0 + 0.98, fw, 1.25)
                D.texto(s, x, y0 + 2.27, fw, 0.3, corta(caption(f), 34), 8.5, GRIS)
        else:
            D.rect(s, x0 + 0.22, y0 + 0.98, w - 0.44, 1.25, BLANCO, MSO_SHAPE.ROUNDED_RECTANGLE, 0.08)
            D.texto(s, x0 + 0.22, y0 + 1.5, w - 0.44, 0.3, "Sin fotos registradas en la app", 10, GRIS, align=PP_ALIGN.CENTER)
        D.tabla(s, x0 + 0.22, y0 + 2.62, w - 0.44, filas_visita(v), 1.0, 8.5)

    def par_slide(par):
        s = D.slide()
        fechas = [fd(v.get("fecha")) for v in par]
        D.marco(s, "Resumen de visitas", "Visitas del " + (fechas[0] if fechas[0] == fechas[-1] else f"{fechas[0]} al {fechas[-1]}"))
        for v, x0 in zip(par, (0.6, 6.79)):
            tarjeta(s, v, x0, 1.55, 5.94)
        for v in par:
            if any(es_nueva(v, n) for n in v["variedades"]):
                nuevas_slide(v)

    def nuevas_slide(v):
        nuevas = [n for n in v["variedades"] if es_nueva(v, n)]
        fs = [f for f in v["fotos"] if f.get("id") and f.get("variedad") in nuevas]
        for k in range(0, max(len(fs), 1), 6):
            s = D.slide()
            D.poligono(s, [(0, 0), (4.6, 0), (3.7, H), (0, H)], VERDE)
            D.icono(s, "flor", 0.6, 0.55, 0.85, (214, 40, 57))
            D.texto(s, 0.6, 1.6, 3.4, 1.0, "Variedades nuevas", 28, BLANCO, True)
            D.texto(s, 0.6, 2.65, 3.2, 0.8, corta(v.get("cliente"), 40), 16, RGBColor(0xFF, 0xC9, 0xD0), True)
            lineas = [fd(v.get("fecha")) + " · " + ", ".join(x for x in [v.get("comuna"), v.get("region")] if x), ""]
            for n in nuevas:
                fp = fplant(v, n)
                e = v["evalVar"].get(n) or {}
                lineas.append(f"{n}" + (f" · {e['estado']}" if e.get("estado") else "")
                              + (f" · plant. {', '.join(fd(x) for x in fp)}" if fp else ""))
            D.texto(s, 0.6, 3.5, 3.1, 3.0, lineas, 12, BLANCO, espacio=4)
            D.texto(s, 11.9, 6.95, 0.83, 0.3, str(D.n), 10, GRIS, align=PP_ALIGN.RIGHT)
            items = [{"id": f["id"], "cap": corta(" · ".join(x for x in [caption(f), f.get("desc")] if x), 70)} for f in fs[k:k + 6]]
            if items:
                D.grilla(s, items, 4.75, 0.6, 3, 2.5, 2.35, 0.2, 0.35, 0.55)
            else:
                D.texto(s, 4.75, 3.4, 8, 0.4, "Sin fotos de estas variedades en la app", 14, GRIS)

    def problema_slide(v):
        s = D.slide()
        D.marco(s, corta(v.get("cliente"), 40), " · ".join(x for x in [fd(v.get("fecha")), v.get("comuna"), v.get("region")] if x),
                titulo_extra=" · problemas detectados")
        D.rect(s, 0.6, 1.55, 5.3, 5.15, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE, 0.04)
        D.icono(s, "hoja", 0.82, 1.75, 0.5, (251, 227, 230))
        D.texto(s, 1.45, 1.83, 3.0, 0.35, "Estado de la visita", 12, GRIS)
        D.pastilla(s, 4.6, 1.85, estado_visita(v))
        D.tabla(s, 0.82, 2.45, 4.86, filas_visita(v, 1.45), 1.0, 8.5)
        fs = [{"id": f["id"], "cap": corta(" · ".join(x for x in [caption(f), f.get("desc")] if x), 55)} for f in v["fotos"] if f.get("id")][:6]
        if fs:
            D.grilla(s, fs, 6.2, 1.55, 3, 2.04, 1.95, 0.2, 0.08, 0.5)
        else:
            D.texto(s, 6.2, 3.8, 6.5, 0.4, "Sin fotos registradas en la app para esta visita", 13, GRIS)

    # primero las visitas sin problemas, de a dos y en orden de fecha; luego las con problemas
    normales = [v for v in vis if not problema(v)]
    for k in range(0, len(normales), 2):
        par_slide(normales[k:k + 2])
    for v in vis:
        if problema(v):
            problema_slide(v)
            if any(es_nueva(v, n) for n in v["variedades"]):
                nuevas_slide(v)

    # casos registrados sin visita presencial (no cuentan como visitas)
    for k in range(0, len(casos), 4):
        s = D.slide()
        D.marco(s, "Casos sin visita presencial", "Reclamos, riesgos y oportunidades registrados por llamado u otro medio")
        for i, v in enumerate(casos[k:k + 4]):
            f, c = divmod(i, 2)
            x, y = 0.6 + c * 6.17, 1.6 + f * 2.75
            D.rect(s, x, y, 5.97, 2.55, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE, 0.05)
            tipos = sorted({d.get("tipo") for d in v["destacados"] if d.get("tipo")})
            es_rec = any(t in ("Reclamo", "Riesgo") for t in tipos)
            D.icono(s, "frutilla" if es_rec else "hoja", x + 0.22, y + 0.22, 0.6, (251, 227, 230) if es_rec else (230, 240, 232))
            D.texto(s, x + 0.95, y + 0.2, 4.8, 0.4, corta(v.get("cliente"), 48), 12.5, TEXTO, True)
            D.texto(s, x + 0.95, y + 0.55, 4.8, 0.3, f"{fd(v.get('fecha'))} · {v.get('medio') or 'Sin medio indicado'}", 10, GRIS)
            D.chip(s, x + 0.22, y + 0.97, " · ".join(tipos) or "Registro")
            txt = " · ".join(f"{d.get('tipo')}: {d.get('texto') or ''}" for d in v["destacados"])
            D.texto(s, x + 0.22, y + 1.35, 5.5, 1.1, corta(txt, 260), 10, TEXTO)

    # 6. oportunidades (tarjetas con ícono, como "Complete toolkit")
    for k in range(0, max(len(ops), 1), 6):
        s = D.slide()
        D.marco(s, "Oportunidades de venta", "Detectadas en terreno para derivar a Ventas")
        if not ops:
            D.icono(s, "frutilla", 5.97, 2.4, 1.4, (251, 227, 230))
            D.texto(s, 1.0, 4.1, 11.3, 0.5, "Sin oportunidades registradas este mes.", 18, GRIS, align=PP_ALIGN.CENTER)
            break
        for i, (v, d) in enumerate(ops[k:k + 6]):
            f, c = divmod(i, 3)
            x, y = 0.6 + c * 4.13, 1.6 + f * 2.6
            D.rect(s, x, y, 3.93, 2.4, GRIS_CLARO, MSO_SHAPE.ROUNDED_RECTANGLE, 0.06)
            D.icono(s, ["frutilla", "flor", "hoja"][i % 3], x + 0.22, y + 0.22, 0.6, (251, 227, 230) if i % 3 == 0 else (230, 240, 232))
            D.texto(s, x + 0.95, y + 0.24, 2.8, 0.6, corta(v.get("cliente"), 42), 12.5, TEXTO, True)
            D.chip(s, x + 0.22, y + 0.97, corta(variedad_oport(v, d), 30))
            D.texto(s, x + 0.22, y + 1.35, 3.5, 0.3, d.get("sub") or "", 10, VERDE, True)
            D.texto(s, x + 0.22, y + 1.62, 3.5, 0.75, corta(d.get("texto"), 115), 10, TEXTO)

    # 7. agenda próxima (columnas por semana, como "The new payments stack")
    s = D.slide()
    D.marco(s, "Agenda próxima", "Visitas programadas en la app para las próximas 4 semanas")
    if agenda:
        semanas = OrderedDict()
        for a in agenda:
            d = date.fromisoformat(a["fecha"][:10])
            lunes = d - timedelta(days=d.weekday())
            semanas.setdefault(lunes, []).append(a)
        cols = list(semanas.items())[:4]
        cw = (W - 1.2 - 0.25 * (len(cols) - 1)) / max(len(cols), 1)
        cw = min(cw, 3.0)
        for j, (lunes, items) in enumerate(cols):
            x = 0.6 + j * (cw + 0.25)
            D.rect(s, x, 1.6, cw, 0.62, VERDE, MSO_SHAPE.ROUNDED_RECTANGLE, 0.15)
            D.texto(s, x + 0.15, 1.66, cw - 0.3, 0.5, [[(f"Semana {lunes.isocalendar()[1]}", BLANCO, True)],
                                                     [(f"{fcorta(lunes.isoformat())} – {fcorta((lunes + timedelta(days=6)).isoformat())}", VERDE_MED, False)]], 11)
            for i, a in enumerate(items[:6]):
                y = 2.35 + i * 0.73
                D.rect(s, x, y, cw, 0.63, VERDE_CLARO if i % 2 == 0 else ROSA, MSO_SHAPE.ROUNDED_RECTANGLE, 0.15)
                cuando = fcorta(a["fecha"]) + (f" {a['hora']}" if a.get("hora") else "")
                D.texto(s, x + 0.14, y + 0.07, cw - 0.28, 0.5,
                        [[(cuando + " · ", VERDE if i % 2 == 0 else ROJO, True), (corta(a.get("cliente"), 26), TEXTO, True)],
                         [(corta(" · ".join(z for z in [a.get("motivo"), a.get("ciudad") or a.get("zona")] if z), 40), GRIS, False)]], 9.5)
            if len(items) > 6:
                D.texto(s, x, 2.35 + 6 * 0.73, cw, 0.3, f"+ {len(items) - 6} más", 9.5, GRIS)
    else:
        D.icono(s, "flor", 5.97, 2.4, 1.4, (214, 40, 57))
        D.texto(s, 1.0, 4.1, 11.3, 0.5, "Sin visitas programadas en la agenda de la app.", 18, GRIS, align=PP_ALIGN.CENTER)

    # cierre: foto del vivero al atardecer (llahuen.com) con velo verde
    s = D.slide(VERDE)
    D.foto_fondo(s, img_atardecer or img_campo, 0.72)
    D.racimo(s, 8.6, 1.0, 1.0, fondo_claro=False)
    D.texto(s, 0.6, 2.2, 8.0, 1.3, "¡Muchas gracias!", 60, BLANCO, True)
    D.texto(s, 0.6, 3.55, 7.6, 1.0, [[("Seguimos acompañando a nuestros productores ", VERDE_MED, False),
                                      ("en cada temporada.", RGBColor(0xFF, 0xC9, 0xCC), True)]], 20)
    if logo:
        D.rect(s, 0.6, 5.3, 2.3, 1.6, CREMA, MSO_SHAPE.ROUNDED_RECTANGLE, 0.08)
        D.imagen(s, logo, 0.78, 5.42, h=1.36)
    D.etiqueta(s, 3.2, 5.95, 6, "Postventa y atención a clientes estratégicos", VERDE_MED, 11)
    D.texto(s, 3.2, 6.3, 6, 0.3, "www.llahuen.com", 12, BLANCO)

    nombre = f"Reporte_mensual_visitas_Llahuen_{ini.year}-{ini.month:02d}.pptx"
    D.p.save(nombre)
    achicar_pptx(nombre)
    resumen = (f"{mes_txt.capitalize()} {ini.year}: {len(vis)} visitas a {len({v.get('cliente') for v in vis})} productores; "
               f"{len(recs)} reclamos y {len(ops)} oportunidades.")
    return nombre, resumen


def achicar_pptx(ruta, limite=15_000_000):
    """Gmail rechaza correos sobre 25 MB (el adjunto crece ~35 % al codificarse).
    Si la presentación pasa el límite, recomprime las fotos JPG hasta que quepa."""
    import zipfile
    for lado, calidad in [(1100, 72), (900, 66), (720, 60), (560, 55), (440, 50)]:
        if os.path.getsize(ruta) <= limite:
            return
        tmp = ruta + ".tmp"
        with zipfile.ZipFile(ruta) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                data = zin.read(it.filename)
                if it.filename.startswith("ppt/media/") and it.filename.lower().endswith((".jpg", ".jpeg")):
                    try:
                        im = PILImage.open(io.BytesIO(data)).convert("RGB")
                        im.thumbnail((lado, lado))
                        b = io.BytesIO()
                        im.save(b, "JPEG", quality=calidad, optimize=True)
                        if len(b.getvalue()) < len(data):
                            data = b.getvalue()
                    except Exception:
                        pass
                zout.writestr(it, data)
        os.replace(tmp, ruta)
    print(f"Tamaño final de la presentación: {os.path.getsize(ruta) / 1e6:.1f} MB")


def enviar(archivo, resumen, mes):
    to = [x.strip() for x in (os.environ.get("MAIL_TO_MENSUAL") or "postventa@llahuen.com").split(",") if x.strip()]
    msg = EmailMessage()
    msg["Subject"] = f"Reporte mensual de visitas - Agrícola Llahuen - {MESES[mes.month - 1]} {mes.year}"
    msg["From"] = os.environ["MAIL_USER"]
    msg["To"] = ", ".join(to)
    msg.set_content(f"Hola,\n\nAdjunto la presentación mensual de visitas a productores.\n\n{resumen}\n\n"
                    "Generado automáticamente desde la app de visitas.")
    with open(archivo, "rb") as f:
        msg.add_attachment(f.read(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.presentationml.presentation", filename=archivo)
    with smtplib.SMTP_SSL(os.environ.get("SMTP_HOST") or "smtp.gmail.com", int(os.environ.get("SMTP_PORT") or 465), timeout=120) as s:
        s.login(os.environ["MAIL_USER"], os.environ["MAIL_PASS"])
        s.send_message(msg)


if __name__ == "__main__":
    hoy = date.fromisoformat(os.environ["HOY"]) if os.environ.get("HOY") else datetime.now(TZ).date()
    if os.environ.get("MES"):
        # Acepta AAAA-MM, AA-MM, MM-AAAA o MM/AAAA (un año de 2 dígitos se entiende como 20AA)
        import re
        a, b = map(int, re.findall(r"\d+", os.environ["MES"])[:2])
        y, m = (a, b) if b <= 12 and (a > 12 or a >= 100) else (b, a)
        y = y + 2000 if y < 100 else y
        mes = date(y, m, 1)
        if (y, m) != (hoy.year, hoy.month):
            hoy = (mes.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    else:
        mes = hoy.replace(day=1)
    archivo, resumen = construir(mes, hoy, os.environ.get("AUTOR") or "Gonzalo Ibarra Castro")
    print("Presentación generada.")
    if os.environ.get("SOLO_PPT") != "1":
        enviar(archivo, resumen, mes)
        print("Correo enviado.")
