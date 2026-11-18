"""Resumen semanal de visitas - Agrícola Llahuen.

Lee las visitas desde Supabase (mismas tablas que la app), arma un PDF
con los puntos principales de la semana y lo envía por correo.

Variables de entorno (se configuran como "secrets" en GitHub):
  SUPABASE_URL, SUPABASE_KEY   -> los de config.js de la app
  TEAM_CODE                    -> código del equipo
  MAIL_USER, MAIL_PASS         -> cuenta Gmail y "contraseña de aplicación"
  MAIL_TO                      -> destinatarios separados por coma
Opcionales: SMTP_HOST (smtp.gmail.com), SMTP_PORT (465), SEMANA (AAAA-MM-DD,
cualquier día de la semana a reportar), SOLO_PDF=1 (no envía correo),
DEMO_JSON (ruta a datos de prueba, no consulta Supabase).
"""
import base64
import io
import json
import os
import smtplib
import sys
from collections import Counter, OrderedDict
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from zoneinfo import ZoneInfo

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import requests  # noqa: E402
from PIL import Image as PILImage  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_CENTER  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa: E402
from reportlab.lib.units import cm  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

TZ = ZoneInfo("America/Santiago")
ESTADOS = ["Excelente", "Bueno", "Regular", "Crítico"]
EST_COL = {"Excelente": "#2E7D4F", "Bueno": "#7BAE5B", "Regular": "#E0A33A", "Crítico": "#C0392B"}
ROJO = colors.HexColor("#B3262E")
VERDE = colors.HexColor("#2E5E3E")
GRIS = colors.HexColor("#6B6B6B")
CLARO = colors.HexColor("#F6F1EC")
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


# ----------------------------------------------------------------- datos
def _headers():
    key = os.environ["SUPABASE_KEY"]
    h = {"apikey": key, "x-llahuen-key": os.environ["TEAM_CODE"], "Accept": "application/json"}
    if key.startswith("eyJ"):  # llave antigua tipo JWT
        h["Authorization"] = "Bearer " + key
    return h


def _base():
    return os.environ["SUPABASE_URL"].rstrip("/") + "/rest/v1/"


def cargar_docs():
    if os.environ.get("DEMO_JSON"):
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


# ----------------------------------------------------------------- utilidades
def fdate(iso):
    try:
        d = date.fromisoformat(str(iso)[:10])
        return f"{d.day:02d}-{d.month:02d}-{d.year}"
    except ValueError:
        return str(iso or "—")


def largo(d):
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def esc(t):
    return (str(t or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace("\n", "<br/>"))


def estado_visita(v):
    """Estado general de la visita; si no hay, el peor estado por variedad."""
    if v.get("estado") in ESTADOS:
        return v["estado"]
    ests = [e.get("estado") for e in v["evalVar"].values() if e.get("estado") in ESTADOS]
    return min(ests, key=lambda s: -ESTADOS.index(s)) if ests else ""


def img_flow(data_url, ancho_cm, alto_max_cm):
    try:
        raw = base64.b64decode(data_url.split(",", 1)[1])
        im = PILImage.open(io.BytesIO(raw)).convert("RGB")
        im.thumbnail((700, 700))
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=78)
        buf.seek(0)
        w, h = im.size
        ancho = ancho_cm * cm
        alto = ancho * h / w
        if alto > alto_max_cm * cm:
            alto = alto_max_cm * cm
            ancho = alto * w / h
        return Image(buf, width=ancho, height=alto)
    except Exception:  # foto dañada o formato no soportado
        return None


VARIEDADES_BASE = ["Camarosa", "Camino Real", "Fronteras", "Adelanto", "Merced", "Belvedere", "Castaic",
                   "San Andreas", "Albion", "Monterey", "Cabrillo", "Alturas", "Portola", "Artesia",
                   "Carpinteria", "Sweet Carolina"]


def variedad_oport(v, d):
    """Variedad de la oportunidad: la marcada en la app, la(s) mencionada(s) en el detalle,
    o las variedades de la visita."""
    if d.get("variedad"):
        return d["variedad"]
    import unicodedata

    def n(t):
        return "".join(c for c in unicodedata.normalize("NFD", str(t or "").lower()) if unicodedata.category(c) != "Mn")
    texto = n(d.get("texto"))
    nombres = list(dict.fromkeys(VARIEDADES_BASE + list(v["variedades"])))
    hall = [x for x in nombres if n(x) and n(x) in texto]
    if hall:
        return ", ".join(hall)
    return ", ".join(v["variedades"]) or "No especificada"


def cargar_logo():
    """Logo de Agrícola Llahuen tomado de la propia app (index.html del repositorio)."""
    import re
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "index.html")
    try:
        with open(ruta, encoding="utf-8") as f:
            html = f.read()
    except OSError:
        return None
    m = re.search(r'<img[^>]*\bid="logo"[^>]*>', html)
    src = re.search(r'src="(data:image/[^"]+)"', m.group(0)) if m else None
    if not src:
        return None
    try:
        raw = base64.b64decode(src.group(1).split(",", 1)[1])
        im = PILImage.open(io.BytesIO(raw))
        if im.mode in ("RGBA", "LA", "P"):
            im = im.convert("RGBA")
            fondo = PILImage.new("RGB", im.size, "white")
            fondo.paste(im, mask=im.split()[-1])
            im = fondo
        buf = io.BytesIO()
        im.save(buf, "PNG")
        buf.seek(0)
        alto = 2.4 * cm
        return Image(buf, width=alto * im.width / im.height, height=alto)
    except Exception:
        return None


def grafico_barras(etiquetas, valores, titulo, col=None, horizontal=False):
    alto = max(1.4, 0.32 * len(valores) + 0.7) if horizontal else 2.2
    fig, ax = plt.subplots(figsize=(5.6, alto), dpi=170)
    cols = col or ["#2E5E3E"] * len(valores)
    if horizontal:
        ax.barh(etiquetas, valores, color=cols, height=0.55)
        ax.invert_yaxis()
        for i, val in enumerate(valores):
            ax.text(val + max(valores) * 0.02, i, str(val), va="center", fontsize=8)
        ax.set_xticks([])
    else:
        ax.bar(etiquetas, valores, color=cols, width=0.45 if len(valores) > 4 else 0.3)
        for i, val in enumerate(valores):
            ax.text(i, val + max(valores) * 0.03, str(val), ha="center", fontsize=8)
        ax.set_yticks([])
        ax.tick_params(axis="x", labelsize=8)
    ax.set_title(titulo, fontsize=10, loc="left", color="#333333")
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis="y", labelsize=8)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    buf.seek(0)
    return Image(buf, width=15.5 * cm, height=15.5 * cm * alto / 5.6)


# ----------------------------------------------------------------- informe
def construir(hoy):
    lunes = hoy - timedelta(days=hoy.weekday())
    domingo = lunes + timedelta(days=6)
    docs = [migrar(d) for d in cargar_docs() if not d.get("ejemplo")]
    visitas = [v for v in docs if not v.get("prog")]
    agenda = [v for v in docs if v.get("prog")]
    sem = sorted([v for v in visitas if lunes.isoformat() <= str(v.get("fecha", ""))[:10] <= domingo.isoformat()],
                 key=lambda v: (v.get("fecha", ""), v.get("cliente", "")))
    ag_sem = [a for a in agenda if lunes.isoformat() <= str(a.get("fecha", ""))[:10] <= domingo.isoformat()]

    ss = getSampleStyleSheet()
    H1 = ParagraphStyle("H1", parent=ss["Heading1"], fontName="Helvetica-Bold", fontSize=17, textColor=VERDE,
                        spaceAfter=6)
    H2 = ParagraphStyle("H2", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=13.5, textColor=VERDE,
                        spaceBefore=12, spaceAfter=6)
    H3 = ParagraphStyle("H3", parent=ss["Heading3"], fontName="Helvetica-Bold", fontSize=12, textColor=ROJO,
                        spaceBefore=8, spaceAfter=3)
    P = ParagraphStyle("P", parent=ss["BodyText"], fontName="Helvetica", fontSize=11, leading=14.5)
    PS = ParagraphStyle("PS", parent=P, fontSize=9.5, leading=12, textColor=GRIS)
    CEL = ParagraphStyle("CEL", parent=P, fontSize=9.5, leading=12)
    HDR = ParagraphStyle("HDR", parent=CEL, textColor=colors.white, fontName="Helvetica-Bold")
    KPI = ParagraphStyle("KPI", parent=P, fontName="Helvetica-Bold", fontSize=20, alignment=TA_CENTER,
                         textColor=ROJO, leading=24)
    KPL = ParagraphStyle("KPL", parent=PS, alignment=TA_CENTER)

    def tabla(filas, anchos, cab=True):
        data = [[Paragraph(esc(c), HDR if (cab and i == 0) else CEL) for c in f] for i, f in enumerate(filas)]
        t = Table(data, colWidths=[a * cm for a in anchos], repeatRows=1 if cab else 0)
        st = [("VALIGN", (0, 0), (-1, -1), "TOP"),
              ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#DDD3C8")),
              ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
        if cab:
            st += [("BACKGROUND", (0, 0), (-1, 0), VERDE)]
        t.setStyle(TableStyle(st))
        return t

    el = []
    productores = OrderedDict()
    for v in sem:
        productores.setdefault((v.get("cliente") or "Sin nombre").strip(), []).append(v)
    regiones = Counter(v.get("region") or "Sin región" for v in sem)
    reclamos = [(v, d) for v in sem for d in v["destacados"] if d.get("tipo") == "Reclamo"]
    oport = [(v, d) for v in sem for d in v["destacados"] if d.get("tipo") == "Oportunidad"]
    riesgos = [(v, d) for v in sem for d in v["destacados"] if d.get("tipo") == "Riesgo"]

    # --- portada / resumen
    titulo = [Paragraph("Resumen semanal de visitas a productores", H1),
              Paragraph(f"Agrícola Llahuen · Postventa y clientes estratégicos<br/>"
                        f"Semana del {largo(lunes)} al {largo(domingo)} · generado el {fdate(hoy.isoformat())}", PS)]
    logo = cargar_logo()
    if logo is not None:
        cab = Table([[logo, titulo]], colWidths=[logo.drawWidth + 0.5 * cm, None])
        cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (0, 0), 0)]))
        el.append(cab)
    else:
        el.extend(titulo)
    el.append(Spacer(1, 10))
    kpis = [(len(sem), "visitas"), (len(productores), "productores"), (len(regiones), "regiones"),
            (len(reclamos), "reclamos"), (len(oport), "oportunidades")]
    kt = Table([[Paragraph(str(n), KPI) for n, _ in kpis], [Paragraph(t, KPL) for _, t in kpis]],
               colWidths=[3.3 * cm] * 5)
    kt.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), CLARO), ("TOPPADDING", (0, 0), (-1, -1), 6),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    el.append(kt)

    if not sem:
        el.append(Spacer(1, 14))
        el.append(Paragraph("No hay visitas registradas en la app para esta semana.", P))
    else:
        # puntos clave
        el.append(Paragraph("Puntos clave", H2))
        claves = []
        abiertos = [x for x in reclamos if x[1].get("estado") != "Resuelto"]
        if reclamos:
            claves.append(f"{len(reclamos)} reclamo(s) registrados; {len(abiertos)} siguen abiertos o en gestión.")
        criticos = [v for v in sem if estado_visita(v) in ("Crítico", "Regular")]
        if criticos:
            claves.append("Plantaciones en estado Regular o Crítico: " +
                          ", ".join(sorted({esc(v.get('cliente')) for v in criticos})) + ".")
        fito = [v for v in sem if v.get("fito") == "Problema"]
        if fito:
            claves.append(f"{len(fito)} visita(s) con problema fitosanitario detectado.")
        if oport:
            claves.append(f"{len(oport)} oportunidad(es) comercial(es) para derivar a Ventas.")
        top = regiones.most_common(1)[0]
        claves.append(f"Región con más visitas: {esc(top[0])} ({top[1]}).")
        for c in claves:
            el.append(Paragraph("• " + c, P))

        # zonas y productores
        el.append(Paragraph("Zonas y productores visitados", H2))
        filas = [["Fecha", "Productor", "Región / comuna", "Tipo", "Variedades", "Estado"]]
        for v in sem:
            filas.append([fdate(v.get("fecha")), v.get("cliente", ""),
                          " / ".join(x for x in [v.get("region"), v.get("comuna")] if x),
                          v.get("tipo", ""), ", ".join(v["variedades"]) or "—", estado_visita(v) or "—"])
        el.append(tabla(filas, [2.3, 3.6, 3.3, 2.4, 3.5, 2.0]))

        # gráficos
        el.append(Paragraph("Gráficos", H2))
        reg = regiones.most_common()
        el.append(grafico_barras([r for r, _ in reg], [n for _, n in reg], "Visitas por región", horizontal=True))

        # reclamos
        el.append(Paragraph("Reclamos", H2))
        if reclamos:
            filas = [["Productor", "Categoría", "Detalle", "Reposición", "Estado"]]
            for v, d in reclamos:
                repo = f"{d.get('repoCant') or 0} pl. {d.get('repoVar') or ''}".strip() if d.get("repo") else "—"
                filas.append([v.get("cliente", ""), d.get("categoria", ""), d.get("texto", ""), repo,
                              d.get("estado", "")])
            el.append(tabla(filas, [3.2, 2.5, 6.1, 2.8, 2.5]))
        else:
            el.append(Paragraph("Sin reclamos registrados esta semana.", P))

        # problemas detectados
        el.append(Paragraph("Problemas detectados en plantaciones", H2))
        probs = []
        for v in sem:
            items = []
            if v.get("fito") in ("Problema", "Observación"):
                items.append(f"Fitosanitario ({v['fito']}): {v.get('fitoDetalle') or 'sin detalle'}")
            for n, e in v["evalVar"].items():
                for k, l in (e.get("lotes") or {}).items():
                    if l.get("fito") in ("Problema", "Observación") or l.get("estado") in ("Regular", "Crítico"):
                        txt = ", ".join(x for x in [l.get("estado"), l.get("fito"), l.get("fitoD")] if x)
                        items.append(f"{n}{'' if k == '0' else ' (' + k + ')'}: {txt}")
            items += [f"Riesgo: {d.get('texto')}" for vv, d in riesgos if vv is v]
            if items:
                probs.append((v, items))
        if probs:
            filas = [["Productor", "Problema / observación"]]
            for v, items in probs:
                filas.append([f"{v.get('cliente', '')}<br/><font size=8 color='#6B6B6B'>{esc(v.get('region'))}</font>",
                              "<br/>".join("• " + esc(i) for i in items)])
            data = [[Paragraph(c, HDR if i == 0 else CEL) for c in f] for i, f in enumerate(filas)]
            t = Table(data, colWidths=[4.2 * cm, 12.9 * cm], repeatRows=1)
            t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BACKGROUND", (0, 0), (-1, 0), VERDE),
                                   ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#DDD3C8"))]))
            el.append(t)
        else:
            el.append(Paragraph("No se registraron problemas sanitarios ni plantaciones en estado Regular o Crítico.", P))

        # oportunidades
        el.append(Paragraph("Oportunidades para Ventas", H2))
        if oport:
            filas = [["Productor", "Variedad", "Tipo", "Detalle", "Estado"]]
            for v, d in oport:
                var = variedad_oport(v, d)
                filas.append([v.get("cliente", ""), var, d.get("sub", ""), d.get("texto", ""), d.get("estado", "")])
            el.append(tabla(filas, [3.3, 2.8, 2.8, 5.9, 2.3]))
        else:
            el.append(Paragraph("Sin oportunidades registradas esta semana.", P))

        # agenda
        no_real = [a for a in ag_sem if a.get("estado") == "no_realizada"]
        pend = [a for a in ag_sem if a.get("estado") == "pendiente" and str(a.get("fecha", ""))[:10] <= hoy.isoformat()]
        if no_real or pend:
            el.append(Paragraph("Agenda: visitas no realizadas o pendientes", H2))
            filas = [["Fecha", "Cliente", "Motivo de la visita", "Situación"]]
            for a in sorted(no_real + pend, key=lambda a: a.get("fecha", "")):
                if a.get("estado") == "no_realizada":
                    sit = f"No realizada: {a.get('motivoNo') or ''} {a.get('detalleNo') or ''}".strip()
                    if a.get("reprogFecha"):
                        sit += f". Reprogramada al {fdate(a['reprogFecha'])}"
                else:
                    sit = "Pendiente (sin marcar)"
                filas.append([fdate(a.get("fecha")), a.get("cliente", ""), a.get("motivo", ""), sit])
            el.append(tabla(filas, [2.2, 3.8, 4.4, 6.7]))

        # detalle por productor
        el.append(PageBreak())
        el.append(Paragraph("Detalle por productor", H2))
        fotos_sel = {}
        for nombre, vs in productores.items():
            fs = [f for v in vs for f in v["fotos"] if f.get("id")]
            fs.sort(key=lambda f: (not f.get("desc"), not f.get("cat")))
            fotos_sel[nombre] = fs[:3]
        datos_foto = cargar_fotos({f["id"] for fs in fotos_sel.values() for f in fs})
        for nombre, vs in productores.items():
            bloque = [Paragraph(esc(nombre), H3)]
            for v in vs:
                lin = [f"<b>{fdate(v.get('fecha'))}</b> · {esc(v.get('tipo'))} · "
                       f"{esc(' / '.join(x for x in [v.get('region'), v.get('comuna')] if x))}"]
                if v.get("predio"):
                    lin.append(f"Predio: {esc(v['predio'])}")
                if v["variedades"]:
                    det = []
                    for n in v["variedades"]:
                        e = v["evalVar"].get(n) or {}
                        fp = sorted({l.get("fplant") for l in (e.get("lotes") or {}).values() if l.get("fplant")})
                        det.append(esc(n) + (f" ({e['estado']})" if e.get("estado") else "") +
                                   (f", plantación {', '.join(fdate(x) for x in fp)}" if fp else ""))
                    lin.append("Variedades: " + "; ".join(det))
                if estado_visita(v):
                    lin.append(f"Estado general: <b>{estado_visita(v)}</b>" +
                               (f" · Etapa: {esc(v.get('etapa'))}" if v.get("etapa") else ""))
                for d in v["destacados"]:
                    lin.append(f"<font color='#B3262E'><b>{esc(d.get('tipo'))}:</b></font> {esc(d.get('texto'))}")
                if v.get("notas"):
                    lin.append("Notas: " + esc(v["notas"]))
                if v.get("recomendaciones"):
                    lin.append("Recomendaciones: " + esc(v["recomendaciones"]))
                if v.get("proxima"):
                    lin.append(f"Próxima visita: {fdate(v['proxima'])}")
                bloque.append(Paragraph("<br/>".join(lin), P))
                bloque.append(Spacer(1, 4))
            celdas = []
            for f in fotos_sel[nombre]:
                im = img_flow(datos_foto[f["id"]], 5.2, 4.2) if f["id"] in datos_foto else None
                if im is None:
                    continue
                pie = " · ".join(esc(x) for x in [f.get("cat"), f.get("variedad")] if x)
                txt = (f"<b>{pie}</b><br/>" if pie else "") + esc(f.get("desc") or "Sin descripción")
                celdas.append([im, Paragraph(txt, PS)])
            if celdas:
                while len(celdas) < 3:
                    celdas.append(["", ""])
                ft = Table([[c[0] for c in celdas], [c[1] for c in celdas]], colWidths=[5.7 * cm] * 3)
                ft.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (0, 0), (-1, 0), "CENTER")]))
                bloque.append(ft)
            else:
                bloque.append(Paragraph("Sin fotos registradas.", PS))
            bloque.append(Spacer(1, 8))
            el.append(KeepTogether(bloque[:2]))
            el.extend(bloque[2:])

    def pie(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(GRIS)
        canvas.drawString(1.8 * cm, 1.1 * cm, "Agrícola Llahuen · Resumen semanal de visitas")
        canvas.drawRightString(A4[0] - 1.8 * cm, 1.1 * cm, f"Página {doc.page}")
        canvas.restoreState()

    nombre_pdf = f"Resumen_visitas_Llahuen_{hoy.isoformat()}.pdf"
    doc = SimpleDocTemplate(nombre_pdf, pagesize=A4, leftMargin=1.8 * cm, rightMargin=1.8 * cm,
                            topMargin=1.6 * cm, bottomMargin=1.8 * cm,
                            title="Resumen semanal de visitas - Agrícola Llahuen")
    doc.build(el, onFirstPage=pie, onLaterPages=pie)

    resumen = (f"Semana del {largo(lunes)} al {largo(domingo)}: {len(sem)} visitas a {len(productores)} "
               f"productores en {len(regiones)} regiones; {len(reclamos)} reclamos y {len(oport)} oportunidades.")
    return nombre_pdf, resumen


def enviar(pdf, resumen):
    to = [x.strip() for x in os.environ["MAIL_TO"].split(",") if x.strip()]
    msg = EmailMessage()
    msg["Subject"] = "Resumen semanal de visitas - Agrícola Llahuen - " + pdf[-14:-4]
    msg["From"] = os.environ["MAIL_USER"]
    msg["To"] = ", ".join(to)
    msg.set_content(f"Hola,\n\nAdjunto el resumen semanal de visitas a productores.\n\n{resumen}\n\n"
                    "Generado automáticamente desde la app de visitas.")
    with open(pdf, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=pdf)
    host = os.environ.get("SMTP_HOST") or "smtp.gmail.com"
    port = int(os.environ.get("SMTP_PORT") or 465)
    with smtplib.SMTP_SSL(host, port, timeout=60) as s:
        s.login(os.environ["MAIL_USER"], os.environ["MAIL_PASS"])
        s.send_message(msg)


if __name__ == "__main__":
    hoy = date.fromisoformat(os.environ["SEMANA"]) if os.environ.get("SEMANA") else datetime.now(TZ).date()
    pdf, resumen = construir(hoy)
    print("PDF generado.")  # sin datos de clientes: el registro del proceso es público
    if os.environ.get("SOLO_PDF") != "1":
        enviar(pdf, resumen)
        print("Correo enviado.")
