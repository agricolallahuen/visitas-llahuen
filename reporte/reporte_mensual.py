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
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from resumen_semanal import (  # noqa: E402
    MESES, TZ, cargar_docs, cargar_fotos, estado_visita, migrar, variedad_oport,
)

VERDE = RGBColor(0x2A, 0x68, 0x46)
ROJO = RGBColor(0xB1, 0x24, 0x3B)
TEXTO = RGBColor(0x17, 0x25, 0x1D)
GRIS = RGBColor(0x5B, 0x6A, 0x60)
FONDO = RGBColor(0xF6, 0xF4, 0xEC)
MARCO = RGBColor(0xEC, 0xEF, 0xEA)
SUBT = RGBColor(0xD7, 0xE8, 0xDC)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)
LINEA = RGBColor(0xD2, 0xDB, 0xD1)
EST_COL = {"Excelente": RGBColor(0x2E, 0x7D, 0x4F), "Bueno": RGBColor(0x6B, 0x9A, 0x3A),
           "Regular": RGBColor(0xC2, 0x7A, 0x2C), "Crítico": ROJO}
FILA = {"Oportunidad": (RGBColor(0xF3, 0xD9, 0x8A), VERDE), "Reclamos": (RGBColor(0xF2, 0xC2, 0xCA), ROJO)}
FILA_DEF = (RGBColor(0xE8, 0xF1, 0xEA), VERDE)
PIE = "Agrícola Llahuen · Postventa y clientes estratégicos"
FONT = "Calibri"
FRUTILLA_B64 = "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"


# ----------------------------------------------------------------- utilidades
def fd(iso):
    try:
        d = date.fromisoformat(str(iso)[:10])
        return f"{d.day:02d}/{d.month:02d}/{d.year}"
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


def foto_bytes(data_url, lado=1000):
    """Foto comprimida (JPEG) y su relación ancho/alto."""
    try:
        im = PILImage.open(io.BytesIO(base64.b64decode(data_url.split(",", 1)[1])))
        try:
            from PIL import ImageOps
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        im = im.convert("RGB")
        im.thumbnail((lado, lado))
        b = io.BytesIO()
        im.save(b, "JPEG", quality=76)
        return b.getvalue(), im.width / im.height
    except Exception:
        return None, None


# ----------------------------------------------------------------- dibujo
class Deck:
    def __init__(self):
        self.p = Presentation()
        self.p.slide_width, self.p.slide_height = Inches(13.333), Inches(7.5)
        self.blank = self.p.slide_layouts[6]
        self.n = 0
        self.fotos = {}

    def slide(self, fondo=FONDO):
        s = self.p.slides.add_slide(self.blank)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = fondo
        self.n += 1
        return s

    def texto(self, s, x, y, w, h, txt, size, color=TEXTO, bold=False, align=PP_ALIGN.LEFT,
              anchor=MSO_ANCHOR.TOP, wrap=True):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        tf.vertical_anchor = anchor
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, Inches(0.03))
        lineas = txt if isinstance(txt, list) else [txt]
        for i, ln in enumerate(lineas):
            par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            par.alignment = align
            r = par.add_run()
            r.text = ln
            r.font.size, r.font.bold, r.font.name = Pt(size), bold, FONT
            r.font.color.rgb = color
        return tb

    def rect(self, s, x, y, w, h, color, forma=MSO_SHAPE.RECTANGLE):
        r = s.shapes.add_shape(forma, Inches(x), Inches(y), Inches(w), Inches(h))
        r.fill.solid()
        r.fill.fore_color.rgb = color
        r.line.fill.background()
        r.shadow.inherit = False
        return r

    def cabecera(self, s, titulo, sub=""):
        self.rect(s, 0, 0, 13.333, 0.78, VERDE)
        self.texto(s, 0.45, 0.1, 9.0, 0.6, titulo, 24, BLANCO, True, anchor=MSO_ANCHOR.MIDDLE)
        if sub:
            self.texto(s, 8.7, 0.1, 4.2, 0.6, sub, 12, SUBT, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        self.texto(s, 0.45, 7.12, 8.0, 0.3, PIE, 9, GRIS)
        self.texto(s, 11.9, 7.12, 1.0, 0.3, str(self.n), 9, GRIS, align=PP_ALIGN.RIGHT)

    def pastilla(self, s, x, y, estado):
        if estado not in EST_COL:
            return
        r = self.rect(s, x, y, 1.2, 0.3, EST_COL[estado], MSO_SHAPE.ROUNDED_RECTANGLE)
        tf = r.text_frame
        tf.margin_top = tf.margin_bottom = 0
        par = tf.paragraphs[0]
        par.alignment = PP_ALIGN.CENTER
        run = par.add_run()
        run.text = estado
        run.font.size, run.font.bold, run.font.name = Pt(11), True, FONT
        run.font.color.rgb = BLANCO

    def foto(self, s, x, y, w, h, fid, marco=MARCO):
        self.rect(s, x, y, w, h, marco)
        datos = self.fotos.get(fid)
        if not datos:
            return False
        raw, ar = foto_bytes(datos)
        if not raw:
            return False
        if w / h > ar:
            ph, pw = h, h * ar
        else:
            pw, ph = w, w / ar
        s.shapes.add_picture(io.BytesIO(raw), Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2), Inches(pw), Inches(ph))
        return True

    def tabla(self, s, x, y, w, filas, ancho_et=1.0, size=9.5):
        t = s.shapes.add_table(len(filas), 2, Inches(x), Inches(y), Inches(w), Inches(0.265 * len(filas))).table
        t.columns[0].width = Inches(ancho_et)
        t.columns[1].width = Inches(w - ancho_et)
        t.first_row = False
        t.horz_banding = False
        for i, (et, val) in enumerate(filas):
            fondo, col = FILA.get(et, FILA_DEF)
            for j, (txt, f, c, b, sz) in enumerate([(et, fondo, col, True, size), (val, BLANCO, TEXTO, False, size + 0.5)]):
                cel = t.cell(i, j)
                cel.fill.solid()
                cel.fill.fore_color.rgb = f
                cel.margin_left = cel.margin_right = Inches(0.06)
                cel.margin_top = cel.margin_bottom = Inches(0.03)
                tf = cel.text_frame
                tf.word_wrap = True
                r = tf.paragraphs[0].add_run()
                r.text = txt
                r.font.size, r.font.bold, r.font.name = Pt(sz), b, FONT
                r.font.color.rgb = c
        return t

    def grilla_fotos(self, s, items, x0, y0, cols, w, h, gap_x, gap_y, cap_h, sub=False):
        for i, it in enumerate(items):
            fila, col = divmod(i, cols)
            x = x0 + col * (w + gap_x)
            y = y0 + fila * (h + cap_h + gap_y)
            self.foto(s, x, y, w, h, it["id"])
            lineas = [it["cap"]] + ([it["sub"]] if sub and it.get("sub") else [])
            tb = self.texto(s, x, y + h + 0.02, w, cap_h, lineas[0], 9, TEXTO)
            if len(lineas) > 1:
                par = tb.text_frame.add_paragraph()
                r = par.add_run()
                r.text = lineas[1]
                r.font.size, r.font.name = Pt(8.5), FONT
                r.font.color.rgb = GRIS


# ----------------------------------------------------------------- contenido
def caption(f):
    return " · ".join(x for x in [f.get("cat") or "Foto", f.get("variedad")] if x)


def es_nueva(v, n):
    e = v["evalVar"].get(n) or {}
    return bool(e.get("nueva"))


def ubic(v):
    partes = [fd(v.get("fecha")), ", ".join(x for x in [v.get("comuna"), v.get("region")] if x), v.get("tipo")]
    if v.get("superficie"):
        partes.append(f"{str(v['superficie']).replace('.', ',')} ha")
    return " · ".join(p for p in partes if p)


def filas_visita(v, largo=1.0):
    n = lambda k: int(k * largo)  # noqa: E731
    dat = " · ".join([v.get("predio") or "Predio no indicado", v.get("etapa") or "etapa no indicada",
                      "Var.: " + (", ".join(v["variedades"]) or "no indicada")])
    if v.get("fito"):
        fito = v["fito"] + (": " + v["fitoDetalle"] if v.get("fitoDetalle") else "")
    else:
        fito = "Sin información sanitaria registrada."
    notas = " ".join(x for x in [v.get("notas"), ("Manejo: " + v["manejo"]) if v.get("manejo") else ""] if x)
    hall = []
    for nom in v["variedades"]:
        e = v["evalVar"].get(nom) or {}
        fp = sorted({l.get("fplant") for l in (e.get("lotes") or {}).values() if l.get("fplant")})
        meta = ", ".join(x for x in [e.get("estado"), ("plant. " + ", ".join(fd(x) for x in fp)) if fp else ""] if x)
        txt = nom + (f" ({meta})" if meta else "")
        if e.get("resultado"):
            txt += ": " + e["resultado"]
        hall.append(txt)
    op = [(d.get("texto") or "") + f" [Var.: {variedad_oport(v, d)}]" for d in v["destacados"] if d.get("tipo") == "Oportunidad"]
    rec = [(d.get("texto") or "") + f" ({d.get('estado', 'Abierto')})" for d in v["destacados"] if d.get("tipo") == "Reclamo"]
    return [("Datos", corta(dat, n(150))), ("Fitosanidad", corta(fito, n(150))),
            ("Notas de la visita", corta(notas, n(300)) or "Sin notas registradas."),
            ("Hallazgos", corta(" · ".join(hall), n(280)) or "Sin hallazgos registrados."),
            ("Oportunidad", corta(" · ".join(op), n(170)) or "Sin oportunidad registrada."),
            ("Reclamos", corta(" · ".join(rec), n(170)) or "Sin reclamos.")]


def problema(v):
    return (estado_visita(v) in ("Regular", "Crítico") or v.get("fito") == "Problema"
            or any(d.get("tipo") == "Reclamo" and d.get("estado") != "Resuelto" for d in v["destacados"]))


def construir(mes, hoy, autor):
    ini = mes
    fin = (mes.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    corte = min(fin, hoy)
    docs = [migrar(d) for d in cargar_docs() if not d.get("ejemplo")]
    vis = sorted([v for v in docs if not v.get("prog") and ini.isoformat() <= str(v.get("fecha", ""))[:10] <= corte.isoformat()],
                 key=lambda v: (v.get("fecha", ""), v.get("creado") or 0))
    agenda = sorted([a for a in docs if a.get("prog") and a.get("estado") == "pendiente"
                     and hoy.isoformat() < str(a.get("fecha", ""))[:10] <= (hoy + timedelta(days=28)).isoformat()],
                    key=lambda a: (a.get("fecha", ""), a.get("hora", "")))

    # selección de fotos
    def fotos_cat(palabras, n, una_por_cliente=True):
        out, vistos = [], set()
        for v in vis:
            for f in v["fotos"]:
                if f.get("id") and any(p in (f.get("cat") or "").lower() for p in palabras):
                    if una_por_cliente and v.get("cliente") in vistos:
                        continue
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

    D = Deck()
    D.fotos = cargar_fotos(ids)
    logo = logo_bytes()

    # 1. portada
    s = D.slide(BLANCO)
    if logo:
        s.shapes.add_picture(io.BytesIO(logo), Inches(3.87), Inches(0.55), Inches(5.6))
    D.texto(s, 0.5, 4.55, 12.33, 0.9, "Postventa y atención a clientes estratégicos", 40, VERDE, True, PP_ALIGN.CENTER)
    D.texto(s, 0.5, 5.45, 12.33, 0.7, "Agrícola Llahuen", 32, ROJO, True, PP_ALIGN.CENTER)
    rango = f"{ini.day} al {corte.day} de {MESES[ini.month - 1]} de {ini.year}"
    D.texto(s, 0.5, 6.3, 12.33, 0.9, [f"Resumen de las visitas registradas en la app Visitas Llahuen · {rango}",
                                      autor, fd(hoy.isoformat())], 14, GRIS, False, PP_ALIGN.CENTER)

    if not vis:
        s = D.slide()
        D.cabecera(s, "Resumen de visitas", rango)
        D.texto(s, 0.5, 3.2, 12.33, 1.0, "No hay visitas registradas en la app para este mes.", 24, GRIS, False, PP_ALIGN.CENTER)
    # 2. apariencia general de los predios
    if predios:
        s = D.slide()
        D.cabecera(s, "Apariencia general de los predios", "Fotografías de las visitas")
        D.grilla_fotos(s, predios, 0.45, 0.95, 4, 3.01, 2.5, 0.12, 0.05, 0.5, sub=True)
    # 3. planta y fruto
    if plantas or frutas:
        s = D.slide()
        D.cabecera(s, "Apariencia de planta y de fruto", "Fotografías de las visitas")
        if plantas:
            D.texto(s, 0.45, 0.85, 6, 0.3, "Apariencia de planta", 12, VERDE, True)
            D.grilla_fotos(s, plantas, 0.45, 1.15, 4, 3.01, 2.0, 0.12, 0, 0.5, sub=True)
        if frutas:
            y = 3.95 if plantas else 1.15
            D.texto(s, 0.45, y - 0.3, 6, 0.3, "Apariencia de fruto", 12, VERDE, True)
            D.grilla_fotos(s, frutas, 0.45, y, 4, 3.01, 2.0, 0.12, 0, 0.5, sub=True)

    # 4. visitas: normales de a 2; con problemas, en su propia diapositiva
    def nuevas_slide(v):
        nuevas = [n for n in v["variedades"] if es_nueva(v, n)]
        fs = [f for f in v["fotos"] if f.get("id") and f.get("variedad") in nuevas]
        for k in range(0, len(fs), 8):
            s = D.slide()
            D.cabecera(s, f"{corta(v.get('cliente'), 40)} · variedades nuevas",
                       f"Registro fotográfico: {', '.join(nuevas)} · {fd(v.get('fecha'))}")
            items = [{"id": f["id"], "cap": " · ".join(x for x in [caption(f), f.get("desc")] if x)} for f in fs[k:k + 8]]
            D.grilla_fotos(s, items, 0.45, 1.0, 4, 3.01, 2.45, 0.12, 0.15, 0.45)

    def problema_slide(v):
        s = D.slide()
        D.cabecera(s, f"{corta(v.get('cliente'), 38)} · problemas detectados",
                   " · ".join(x for x in [fd(v.get("fecha")), v.get("comuna"), v.get("region")] if x))
        D.texto(s, 0.45, 0.9, 3.7, 0.4, corta(v.get("cliente"), 34), 16, TEXTO, True)
        D.pastilla(s, 4.2, 0.95, estado_visita(v))
        D.tabla(s, 0.45, 1.4, 5.0, filas_visita(v, 1.6), 1.0, 9)
        fs = [{"id": f["id"], "cap": " · ".join(x for x in [caption(f), f.get("desc")] if x)} for f in v["fotos"] if f.get("id")][:8]
        if fs:
            D.grilla_fotos(s, fs, 5.75, 1.0, 4, 1.7, 2.3, 0.12, 0.3, 0.5)
        else:
            D.rect(s, 5.75, 1.0, 7.15, 2.3, MARCO)
            D.texto(s, 5.75, 2.0, 7.15, 0.3, "Sin fotos registradas en la app para esta visita", 9.5, GRIS, align=PP_ALIGN.CENTER)

    def par_slide(par):
        s = D.slide()
        fechas = sorted({fd(v.get("fecha")) for v in par}, key=lambda x: x[6:] + x[3:5] + x[:2])
        D.cabecera(s, "Resumen de visitas", "Visitas del " + (fechas[0] if len(fechas) == 1 else f"{fechas[0]} al {fechas[-1]}"))
        if len(par) == 2:
            D.rect(s, 6.62, 0.95, 0.01, 6.0, LINEA)
        for v, x0 in zip(par, (0.5, 6.78)):
            D.texto(s, x0, 0.86, 4.75, 0.4, corta(v.get("cliente"), 42), 16, TEXTO, True)
            D.pastilla(s, x0 + 4.85, 0.92, estado_visita(v))
            D.texto(s, x0, 1.25, 6.05, 0.3, ubic(v), 10.5, GRIS)
            fs = [f for f in v["fotos"] if f.get("id")]
            fs.sort(key=lambda f: (not f.get("desc"), not f.get("cat")))
            fs = fs[:3]
            if fs:
                for i, f in enumerate(fs):
                    x = x0 + i * 2.05
                    D.foto(s, x, 1.61, 1.95, 2.0, f["id"])
                    D.texto(s, x, 3.64, 1.95, 0.42, corta(caption(f), 48), 9, GRIS)
            else:
                D.rect(s, x0, 1.61, 6.05, 2.0, MARCO)
                D.texto(s, x0, 2.5, 6.05, 0.3, "Sin fotos registradas en la app para esta visita", 9.5, GRIS, align=PP_ALIGN.CENTER)
            D.tabla(s, x0, 4.1, 6.05, filas_visita(v), 1.0, 9)
        for v in par:
            if any(es_nueva(v, n) for n in v["variedades"]):
                nuevas_slide(v)

    pend = []
    for v in vis:
        if problema(v):
            if pend:
                par_slide(pend)
                pend = []
            problema_slide(v)
            if any(es_nueva(v, n) for n in v["variedades"]):
                nuevas_slide(v)
        else:
            pend.append(v)
            if len(pend) == 2:
                par_slide(pend)
                pend = []
    if pend:
        par_slide(pend)

    # 5. oportunidades
    ops = [(v, d) for v in vis for d in v["destacados"] if d.get("tipo") == "Oportunidad"]
    for k in range(0, max(len(ops), 1), 9):
        s = D.slide()
        D.cabecera(s, "Oportunidades de venta detectadas", "Interés registrado en la app durante las visitas")
        if not ops:
            D.texto(s, 0.5, 3.2, 12.33, 0.8, "Sin oportunidades registradas este mes.", 20, GRIS, align=PP_ALIGN.CENTER)
            break
        filas = [("Productor", "Variedad", "Tipo", "Detalle", "Estado")]
        filas += [(corta(v.get("cliente"), 40), variedad_oport(v, d), d.get("sub") or "", corta(d.get("texto"), 190),
                   d.get("estado") or "") for v, d in ops[k:k + 9]]
        anchos = [2.6, 2.0, 1.9, 4.6, 1.3]
        t = s.shapes.add_table(len(filas), 5, Inches(0.45), Inches(1.05), Inches(sum(anchos)), Inches(0.4 * len(filas))).table
        for j, a in enumerate(anchos):
            t.columns[j].width = Inches(a)
        for i, fila in enumerate(filas):
            for j, txt in enumerate(fila):
                cel = t.cell(i, j)
                cel.fill.solid()
                cel.fill.fore_color.rgb = VERDE if i == 0 else BLANCO
                cel.margin_left = cel.margin_right = Inches(0.08)
                r = cel.text_frame.paragraphs[0].add_run()
                r.text = str(txt)
                r.font.size, r.font.name, r.font.bold = Pt(11 if i == 0 else 10.5), FONT, i == 0
                r.font.color.rgb = BLANCO if i == 0 else TEXTO

    # 6. agenda próxima
    s = D.slide()
    D.cabecera(s, "Agenda próxima", "Visitas programadas en la app")
    D.rect(s, 2.9, 1.2, 7.5, 5.6, BLANCO, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.04
    lineas = []
    if agenda:
        semanas = OrderedDict()
        for a in agenda:
            d = date.fromisoformat(a["fecha"][:10])
            semanas.setdefault(f"W{d.isocalendar()[1]}", []).append(a)
        for w, items in semanas.items():
            lineas.append((w, True))
            for a in items[:7]:
                d = date.fromisoformat(a["fecha"][:10])
                txt = f"{d.day} {MESES[d.month - 1][:3]}{(' ' + a['hora']) if a.get('hora') else ''} · {a.get('cliente', '')}"
                if a.get("motivo"):
                    txt += f" · {a['motivo']}"
                if a.get("ciudad") or a.get("zona"):
                    txt += f" ({a.get('ciudad') or a.get('zona')})"
                lineas.append((corta(txt, 85), False))
    else:
        lineas = [("Sin visitas programadas en la agenda de la app para las próximas 4 semanas.", False)]
    tb = s.shapes.add_textbox(Inches(3.2), Inches(1.45), Inches(6.9), Inches(5.2))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, (txt, cab) in enumerate(lineas[:22]):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        par.space_before = Pt(8 if cab and i else 2)
        r = par.add_run()
        r.text = txt if cab else "✓  " + txt
        r.font.size, r.font.bold, r.font.name = Pt(16 if cab else 13), cab, FONT
        r.font.color.rgb = VERDE if cab else TEXTO

    # 7. cierre
    s = D.slide()
    fr = base64.b64decode(FRUTILLA_B64)
    s.shapes.add_picture(io.BytesIO(fr), Inches(5.17), Inches(0.55), Inches(3.0), Inches(3.0))
    D.texto(s, 0.5, 3.45, 12.33, 1.6, "¡Muchas gracias!", 88, ROJO, True, PP_ALIGN.CENTER)
    D.texto(s, 0.5, 5.05, 12.33, 0.5, "Agrícola Llahuen · Postventa y atención a clientes estratégicos", 20, TEXTO, False, PP_ALIGN.CENTER)
    if logo:
        s.shapes.add_picture(io.BytesIO(logo), Inches(5.97), Inches(5.75), Inches(1.4))

    nombre = f"Reporte_mensual_visitas_Llahuen_{ini.year}-{ini.month:02d}.pptx"
    D.p.save(nombre)
    resumen = (f"{MESES[ini.month - 1].capitalize()} {ini.year}: {len(vis)} visitas a "
               f"{len({v.get('cliente') for v in vis})} productores; "
               f"{sum(1 for v in vis for d in v['destacados'] if d.get('tipo') == 'Reclamo')} reclamos y {len(ops)} oportunidades.")
    return nombre, resumen


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
        y, m = map(int, os.environ["MES"].split("-")[:2])
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
